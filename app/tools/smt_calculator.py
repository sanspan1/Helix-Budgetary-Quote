"""
smt_calculator.py
------------------
Deterministic implementation of the BMC Helix SMT's formula chains, so
the agent calls this code for arithmetic instead of reasoning through
multi-table lookups and discount math in prose.

Entry points:
  - compute_new_business_line(...)   -> Part 1, New Input Form (one line)
  - compute_renewal_line(...)        -> Part 2, Renewal Input Form (one line)
  - compute_renewal_quote(...)       -> Part 2, every line + totals + summary
  - render_renewal_table_markdown()  -> chat table built from computed values
  - write_renewal_output_xlsx(...)   -> Renewal Input Form .xlsx

All totals, list-price benchmarks and overall discount figures are
computed here. The LLM must never sum or derive any of them itself.

Raises ValueError with a clear message when a required input is missing
or a product can't be matched — surface that to the user, never
substitute a guess.
"""

SKILL_VERSION = 8

import csv
import os
from datetime import date, datetime

REF_DIR = os.path.join(os.path.dirname(__file__), "..", "references")

BASE_LIST_TIER_SENTINELS = ("Helix Base List Price", "BMC Base List Price")
OPS_LICENSE_TYPES = ("OPS",)
MONTHS_PER_YEAR = 12
MONEY_DP = 2

CODE_FLAT = ""
CODE_ADD = "a"
CODE_DECREASE = "d"
CODE_EXCLUDE = "x"
CODE_LABELS = {
    CODE_FLAT: "Flat renewal",
    CODE_ADD: "a (growth)",
    CODE_DECREASE: "d (decrease)",
    CODE_EXCLUDE: "x (exclude)",
}


def _load_csv(filename):
    path = os.path.join(REF_DIR, filename)
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    # Strip header/value whitespace (e.g. "BMC Base List Price " has a
    # trailing space in channel_discount_table.csv).
    return [{(k or "").strip(): (v.strip() if isinstance(v, str) else v)
             for k, v in row.items()} for row in rows]


def _normalize(name):
    return "".join(str(name).split()).lower()


def _to_date(value):
    """Accepts date, datetime, 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM:SS'."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.fromisoformat(str(value).strip()).date()


def _money(x):
    return round(float(x), MONEY_DP)


def _is_ops(license_type):
    return license_type in OPS_LICENSE_TYPES


class KnowledgeBase:
    """Loads all reference CSVs once; reused across line-item calcs."""

    def __init__(self):
        self.price_list = _load_csv("bmc_price_list.csv")
        self.channel_table = _load_csv("channel_discount_table.csv")
        self.sdn_table = _load_csv("sdn_disnet_gss_discount_table.csv")
        self.ervfactor_table = _load_csv("ervfactor_table.csv")
        self.trvfactor_table = _load_csv("trvfactor_table.csv")
        self.obsoleting = _load_csv("obsoleting_table.csv")
        self.cap_discount = _load_csv("cap_discount_table.csv")
        self.support_rates = _load_csv("support_rate_table.csv")

        self._price_index = {}
        for row in self.price_list:
            key = (_normalize(row["Product Name"]), row["License Type"])
            self._price_index[key] = row

        # Blank Excel dates export as "00:00:00"; treat as no obsolete date
        # (Excel treats a blank AO as later than any date).
        self._obsoleting_index = {}
        self.obsoleting_blank_dates = []
        for r in self.obsoleting:
            raw = r.get("Projected Obsolete Date")
            try:
                parsed = _to_date(raw)
            except ValueError:
                parsed = None
            if parsed is None:
                self.obsoleting_blank_dates.append(r["Product Name"])
                continue
            self._obsoleting_index[_normalize(r["Product Name"])] = parsed
        self._cap_index = {
            _normalize(r["Product Name"]): float(r["Cap Discount %"])
            for r in self.cap_discount if r.get("Cap Discount %") not in (None, "")
        }
        self._support_index = {
            _normalize(r["Support Tier"]): float(r["Support Rate"])
            for r in self.support_rates if r.get("Support Rate") not in (None, "")
        }

    def find_product(self, product_name, license_type):
        """Exact match required (after whitespace normalization).
        Returns None if not found — caller must flag, not guess."""
        return self._price_index.get((_normalize(product_name), license_type))

    def channel_discount(self, product_family, customer_tier):
        """Customer Tier discount for a Product Family. Returns None when
        the cell is blank or the tier column doesn't exist."""
        tier = (customer_tier or "").strip()
        for row in self.channel_table:
            if row["PF"] == product_family:
                val = row.get(tier)
                return float(val) if val not in (None, "") else None
        return None

    def sdn_discount(self, product_line, deal_type, license_type):
        """Part 1 only. deal_type: 'SDN' / 'Disnet' / 'GSS'."""
        col = f"{deal_type} % ({license_type})"
        for row in self.sdn_table:
            if row["Product Line (PL)"] == product_line:
                val = row.get(col)
                return float(val) if val not in (None, "") else 0.0
        return 0.0

    def obsolete_date(self, product_name):
        return self._obsoleting_index.get(_normalize(product_name))

    def cap_discount_pct(self, product_name):
        return self._cap_index.get(_normalize(product_name))

    def support_rate(self, support_tier):
        return self._support_index.get(_normalize(support_tier or ""))

    def growth_factor(self, table, old_term_years, new_term_years):
        old_term_years = max(1, min(5, round(old_term_years)))
        new_term_years = max(1, min(5, round(new_term_years)))
        for row in table:
            if int(row["OldTermYrs"]) == old_term_years:
                return float(row[f"NewTermYrs_{new_term_years}"])
        raise ValueError(f"No growth factor row for OldTermYrs={old_term_years}")


# ---------------------------------------------------------------------
# PART 1 — New Input Form (genuinely new business)
# ---------------------------------------------------------------------

def compute_new_business_line(kb: KnowledgeBase, product_name, qty, license_type,
                                customer_tier, deal_type, term_months=12):
    """
    license_type: 'OPS' or 'Perpetual'
    deal_type: 'SDN', 'Disnet', 'GSS', or None. None raises.
    """
    row = kb.find_product(product_name, license_type)
    if row is None:
        raise ValueError(f"No exact price-list match for '{product_name}' ({license_type}). "
                          f"Flag this — do not guess a substitute.")

    list_price = float(row["List Price (USD)"])
    product_line = row["Product Line (PL)"]
    product_family = row["Product Family"]

    channel_disc = kb.channel_discount(product_family, customer_tier)
    if channel_disc is None:
        raise ValueError(f"No channel discount found for family='{product_family}', "
                          f"tier='{customer_tier}'. Flag — do not assume 0.")

    partner_price = list_price * (1 - channel_disc)

    if deal_type is None:
        raise ValueError("Deal Type (SDN/Disnet/GSS) is required and was not provided. "
                          "This is a per-deal input — ask, do not default.")

    bw = kb.sdn_discount(product_line, deal_type, license_type)
    if bw == 0:
        raise ValueError(
            f"BW=0 for Product Line '{product_line}' under Deal Type '{deal_type}' — "
            "flag this line for manual review rather than computing AS/AX automatically."
        )

    AS = 1 - (1 - bw) / (1 - channel_disc)
    AT = 1 - (1 - (0.9999 if deal_type == "GSS" else 0.9998)) / (1 - bw)

    AX = partner_price * (1 - AS) * (1 - AT)
    AZ = AX * term_months * qty
    BB = AZ
    AK = list_price * MONTHS_PER_YEAR * qty

    return {
        "product": product_name, "license_type": license_type,
        "product_line": product_line, "product_family": product_family,
        "list_price": list_price, "partner_price": partner_price,
        "channel_discount": channel_disc, "BW": bw, "AS": AS, "AT": AT,
        "AX": AX, "AZ": AZ, "BA": 0.0, "BB": BB, "AK": AK,
        "obsolete_date": kb.obsolete_date(product_name),
    }


# ---------------------------------------------------------------------
# PART 2 — Renewal Input Form (existing subscription)
# ---------------------------------------------------------------------

def _resolve_channel_discount(kb, product_family, customer_tier):
    if (customer_tier or "").strip() in BASE_LIST_TIER_SENTINELS:
        return 0.0
    channel_disc = kb.channel_discount(product_family, customer_tier)
    if channel_disc is None:
        raise ValueError(
            f"No Customer Tier discount in channel_discount_table.csv for "
            f"family='{product_family}', tier='{customer_tier}'. Flag — do not assume 0."
        )
    return channel_disc


def compute_renewal_line(kb: KnowledgeBase, product_name, license_type,
                           existing_qty, proposed_qty, code,
                           orig_brv_acv, orig_trv_acv, orig_erv_acv, orig_term_months,
                           new_term_months, customer_tier,
                           support_rate=None,
                           obsolete_date=None, max_term_end_date_across_quote=None,
                           new_term_years_rounded_across_quote=None,
                           u_override=None, v_override=None, att_a_unit_cost=None):
    """
    Renewal (ACV growth model) for one line.

    Discount structure (all against the monthly Base List Price, LP):
      - AS (License Trans, SMT column AS) = 1 − Proposed ACV / (LP × 12 × qty)
        for OPS. This is the SMT's own number and is the EFFECTIVE
        discount vs base list.
      - AX (License Net Price) = LP × (1 − AS). The Customer Tier
        discount is NOT multiplied in again: the renewal price is
        anchored to the prior ACV, so AX × 12 × qty == Proposed ACV.
      - Customer Tier discount (channel table) is reported as the first
        layer, and the remainder as the additional license discount on
        the Partner Discounted Price:
            additional = 1 − (1 − AS) / (1 − tier)
        so (1 − tier) × (1 − additional) == (1 − AS) exactly.

    code: 'a' (add), 'd' (decrease), 'x' (exclude), or '' (flat renewal)
    customer_tier: exact channel_discount_table.csv column name, or
        "BMC Base List Price" / "Helix Base List Price" for 0%.
    support_rate: required for Perpetual lines, None for OPS.
    obsolete_date / max_term_end_date_across_quote /
    new_term_years_rounded_across_quote: growth-factor cap inputs,
        computed once per quote by compute_renewal_quote().
    """
    if None in (orig_brv_acv, orig_trv_acv, orig_erv_acv, orig_term_months):
        raise ValueError(
            "Prior-period BRV/TRV/ERV ACV and term data are required for a "
            "renewal calculation and cannot come from the current quote PDF."
        )
    is_ops = _is_ops(license_type)
    if not is_ops and code not in (CODE_DECREASE, CODE_EXCLUDE) and support_rate is None:
        raise ValueError(
            f"support_rate is required for a Perpetual line ('{product_name}'). "
            "Resolve it from support_rate_table.csv via the UFR Support Tier."
        )

    row = kb.find_product(product_name, license_type)
    if row is None:
        raise ValueError(f"No exact price-list match for '{product_name}' ({license_type}).")

    list_price = float(row["List Price (USD)"])
    product_family = row["Product Family"]
    channel_disc = _resolve_channel_discount(kb, product_family, customer_tier)
    partner_price = list_price * (1 - channel_disc)

    if obsolete_date is None:
        obsolete_date = kb.obsolete_date(product_name)
    obsolete_date = _to_date(obsolete_date)
    max_end = _to_date(max_term_end_date_across_quote)
    cap_pct = kb.cap_discount_pct(product_name)

    # Real sheet stores prior-period ACVs as whole-dollar figures.
    orig_brv_acv = round(orig_brv_acv)
    orig_trv_acv = round(orig_trv_acv)
    orig_erv_acv = round(orig_erv_acv)

    # "Remediate On Prem" family convention: TRV set to the rounded ERV.
    if product_family == "Remediate On Prem":
        orig_trv_acv = orig_erv_acv

    obsolete_during_period = bool(obsolete_date and max_end and obsolete_date < max_end)

    common_meta = {
        "product": product_name, "license_type": license_type,
        "prp_number": row.get("PRP Number", ""),
        "PL": row["Product Line (PL)"], "product_family": product_family,
        "uom": row.get("Unit of Measure", ""), "status": row.get("Status", ""),
        "list_price": list_price, "partner_price": partner_price,
        "customer_tier": customer_tier, "channel_discount": channel_disc,
        "new_term_months": new_term_months,
        "orig_brv_acv": orig_brv_acv, "orig_trv_acv": orig_trv_acv,
        "orig_erv_acv": orig_erv_acv, "orig_term_months": orig_term_months,
        "support_rate": support_rate if not is_ops else None,
        "obsolete_date": obsolete_date, "obsolete_during_period": obsolete_during_period,
        "cap_discount": cap_pct,
    }

    if code in (CODE_DECREASE, CODE_EXCLUDE):
        return {
            **common_meta,
            "code": code, "existing_qty": existing_qty, "proposed_qty": 0,
            "AS": 0.0, "AT": 0.0, "additional_license_discount": 0.0,
            "effective_discount": 0.0, "AX": list_price, "AY": 0.0, "AQ": None,
            "AZ": 0.0, "BA": 0.0, "BB": 0.0, "AK": 0.0, "term_list_value": 0.0,
            "ERVFactor": None, "TRVFactor": None, "Y": None, "AD": None, "AE": None,
            "BN": None, "warnings": [],
            "note": "Code is decrease/exclude — real sheet zeroes AZ/BA/BB. If this "
                    "quantity reappears as a new line elsewhere in the quote, flag "
                    "as a possible license conversion.",
        }

    # Growth factors (old-term-years -> new-term-years lookup)
    old_yrs = orig_term_months / MONTHS_PER_YEAR
    new_yrs = new_term_months / MONTHS_PER_YEAR
    erv_factor = kb.growth_factor(kb.ervfactor_table, old_yrs, new_yrs)
    trv_factor_raw = kb.growth_factor(kb.trvfactor_table, old_yrs, new_yrs)

    cap_condition = (
        obsolete_during_period
        and new_term_years_rounded_across_quote is not None
        and round(new_yrs) < new_term_years_rounded_across_quote
    )
    trv_factor = min(trv_factor_raw, 1) if cap_condition else trv_factor_raw

    # BN: chosen annual unit list (ceiling reference)
    bn = list_price * MONTHS_PER_YEAR if is_ops else list_price * support_rate
    overrides = [v for v in (u_override, v_override) if v]
    bo = max(overrides) if overrides else bn

    brv_unit = orig_brv_acv / existing_qty if existing_qty else 0.0
    trv_unit = orig_trv_acv / existing_qty if existing_qty else 0.0

    # Y (Anchor Price TRV, per unit) = MIN(brv + (trv − brv) × factor, BN, BO)
    y = min(brv_unit + (trv_unit - brv_unit) * trv_factor, bn, bo)
    ad = y * proposed_qty          # Anchor TRV ACV
    ae = ad                        # Proposed ACV

    at = 0.0
    denom = bn * proposed_qty
    AS = 1 - (ae / denom) / (1 - at) if denom else 0.0

    AX = list_price * (1 - AS) * (1 - at)
    effective_discount = 1 - (AX / list_price) if list_price else 0.0
    additional_license_discount = (
        1 - (1 - effective_discount) / (1 - channel_disc) if channel_disc < 1 else 0.0
    )

    AK = list_price * MONTHS_PER_YEAR * proposed_qty    # Total Annual List (BP use)

    if is_ops:
        AQ = None
        AY = 0.0
        AZ = AX * new_term_months * proposed_qty
        BA = 0.0
        term_list_value = list_price * new_term_months * proposed_qty
    else:
        net_new_units = max(proposed_qty - existing_qty, 0)
        AQ = list_price * support_rate / MONTHS_PER_YEAR
        AY = AQ * (1 - AS)
        AZ = AX * net_new_units
        BA = AY * proposed_qty * new_term_months
        term_list_value = list_price * net_new_units + AQ * proposed_qty * new_term_months

    AZ, BA = _money(AZ), _money(BA)
    BB = _money(AZ + BA)

    warnings = []
    if cap_pct is not None and effective_discount > cap_pct:
        warnings.append(
            f"Effective discount {effective_discount:.2%} exceeds the cap_discount_table "
            f"cap of {cap_pct:.2%} for this product — needs discount approval / review."
        )
    if obsolete_during_period:
        warnings.append(
            f"Projected obsolete date {obsolete_date.isoformat()} falls before the quote "
            f"end date {max_end.isoformat()}."
        )
    if is_ops:
        expected = ae * new_term_months / MONTHS_PER_YEAR
        if abs(AZ - expected) > 0.01 * max(proposed_qty, 1):
            warnings.append(
                f"Internal check failed: Total License {AZ:.2f} != Proposed ACV × term/12 "
                f"({expected:.2f})."
            )

    return {
        **common_meta,
        "code": code, "existing_qty": existing_qty, "proposed_qty": proposed_qty,
        "ERVFactor": erv_factor, "TRVFactor": trv_factor, "TRVFactorRaw": trv_factor_raw,
        "BN": bn, "Y": y, "AD": ad, "AE": ae, "AQ": AQ,
        "AS": AS, "AT": at, "AU": None, "AX": AX, "AY": AY,
        "additional_license_discount": additional_license_discount,
        "effective_discount": effective_discount,
        "AZ": AZ, "BA": BA, "BB": BB, "AK": AK, "term_list_value": term_list_value,
        "warnings": warnings,
    }


def compute_renewal_quote(kb: KnowledgeBase, line_inputs, customer_tier):
    """
    Computes every line of one renewal quote plus quote-level totals.

    line_inputs: list of dicts with keys product_name, license_type,
        existing_qty, proposed_qty, code, orig_brv_acv, orig_trv_acv,
        orig_erv_acv, orig_term_months, new_term_months, term_end,
        support_rate (None for OPS).

    Returns {"lines": [...], "totals": {...}, "summary": {...}}.
    """
    if not line_inputs:
        raise ValueError("No renewal lines to compute.")

    term_ends = [_to_date(l.get("term_end")) for l in line_inputs if l.get("term_end")]
    max_term_end = max(term_ends) if term_ends else None
    new_term_years_rounded = max(
        round(l["new_term_months"] / MONTHS_PER_YEAR) for l in line_inputs
    )

    results = []
    for l in line_inputs:
        results.append(compute_renewal_line(
            kb, l["product_name"], license_type=l["license_type"],
            existing_qty=l["existing_qty"], proposed_qty=l["proposed_qty"], code=l["code"],
            orig_brv_acv=l["orig_brv_acv"], orig_trv_acv=l["orig_trv_acv"],
            orig_erv_acv=l["orig_erv_acv"], orig_term_months=l["orig_term_months"],
            new_term_months=l["new_term_months"], customer_tier=customer_tier,
            support_rate=l.get("support_rate"),
            max_term_end_date_across_quote=max_term_end,
            new_term_years_rounded_across_quote=new_term_years_rounded,
        ))

    total_az = _money(sum(r["AZ"] for r in results))
    total_ba = _money(sum(r["BA"] for r in results))
    total_bb = _money(sum(r["BB"] for r in results))
    total_list_term = _money(sum(r["term_list_value"] for r in results))
    total_annual_list = _money(sum(r["AK"] for r in results))
    total_prior_brv = _money(sum(r["orig_brv_acv"] for r in results))
    total_proposed_acv = _money(sum(r["AE"] or 0 for r in results))
    overall_discount = 1 - total_bb / total_list_term if total_list_term else 0.0

    totals = {
        "total_license_AZ": total_az, "total_support_BA": total_ba, "total_BB": total_bb,
        "total_list_value_for_term": total_list_term,
        "total_annual_list_AK": total_annual_list,
        "total_prior_brv_acv": total_prior_brv,
        "total_proposed_acv": total_proposed_acv,
        "overall_discount": overall_discount,
    }

    by_code = {}
    for r in results:
        by_code.setdefault(r["code"], []).append(r["product"])
    tiers = sorted({(r["customer_tier"], r["channel_discount"]) for r in results},
                   key=lambda t: str(t))
    support_lines = [(r["product"], r["support_rate"]) for r in results
                     if not _is_ops(r["license_type"]) and r["code"] not in (CODE_DECREASE, CODE_EXCLUDE)]

    summary = {
        "line_count": len(results),
        "customer_tiers": [{"tier": t, "discount": d} for t, d in tiers],
        "lines_by_code": {CODE_LABELS.get(k, k): v for k, v in by_code.items()},
        "perpetual_support_lines": support_lines,
        "max_term_end": max_term_end.isoformat() if max_term_end else None,
        "warnings": [f"{r['product']}: {w}" for r in results for w in r["warnings"]],
    }
    summary["text"] = _summary_text(results, totals, summary)
    return {"lines": results, "totals": totals, "summary": summary}


def _summary_text(results, totals, summary):
    n = summary["line_count"]
    tier_bits = ", ".join(f"{t['tier']} ({t['discount']:.2%})" for t in summary["customer_tiers"])
    parts = [f"All {n} lines were priced with Customer Tier {tier_bits}."]

    ops_count = sum(1 for r in results if _is_ops(r["license_type"]))
    if summary["perpetual_support_lines"]:
        sup = "; ".join(f"{p} at {rate:.0%}" for p, rate in summary["perpetual_support_lines"])
        parts.append(f"Support rates were applied to Perpetual lines only: {sup}.")
    if ops_count:
        parts.append(f"{ops_count} line(s) are OPS (on-prem subscription), where support "
                     "is included in the monthly unit price, so no separate support rate "
                     "or support charge applies.")

    code_bits = []
    for label, products in summary["lines_by_code"].items():
        if label == CODE_LABELS[CODE_FLAT]:
            code_bits.append(f"{len(products)} flat renewal(s)")
        else:
            code_bits.append(f"{label}: {', '.join(products)}")
    parts.append("Line changes: " + "; ".join(code_bits) + ".")

    parts.append(
        f"Undiscounted list value for the term (monthly list price × term months × "
        f"proposed quantity) is ${totals['total_list_value_for_term']:,.2f}; the grand "
        f"total of ${totals['total_BB']:,.2f} is an overall discount of "
        f"{totals['overall_discount']:.2%} (1 − grand total ÷ list value)."
    )
    parts.append(
        f"Total proposed ACV is ${totals['total_proposed_acv']:,.2f} against prior booked "
        f"ACV of ${totals['total_prior_brv_acv']:,.2f}."
    )
    if summary["warnings"]:
        parts.append("Review flags: " + " | ".join(summary["warnings"]))
    return " ".join(parts)


def _pct(x):
    return "" if x is None else f"{x:.2%}"


def _usd(x, dp=2):
    return "" if x is None else f"{x:,.{dp}f}"


def render_renewal_table_markdown(quote):
    """Builds the chat table from computed values. TOTAL row sums only
    the money columns, from the same rounded line values."""
    header = ("| Product | License Type | Code | Existing Qty | Proposed Qty | Support Rate "
              "| List Price (monthly) | Customer Tier Discount % | Additional License Discount % "
              "| Effective Discount % (SMT AS) | Net Unit Price (AX, monthly) "
              "| Total License (AZ) | Total Support (BA) | Total (BB) |")
    sep = "|" + "---|" * 14
    rows = [header, sep]
    for r in quote["lines"]:
        support = ("Included (OPS)" if _is_ops(r["license_type"])
                   else _pct(r["support_rate"]))
        rows.append(
            f"| {r['product']} | {r['license_type']} | {CODE_LABELS.get(r['code'], r['code'])} "
            f"| {r['existing_qty']:,} | {r['proposed_qty']:,} | {support} "
            f"| {_usd(r['list_price'])} | {_pct(r['channel_discount'])} "
            f"| {_pct(r['additional_license_discount'])} | {_pct(r['effective_discount'])} "
            f"| {_usd(r['AX'], 4)} | {_usd(r['AZ'])} | {_usd(r['BA'])} | {_usd(r['BB'])} |"
        )
    t = quote["totals"]
    rows.append(
        f"| **TOTAL** | | | | | | | | | | | **{_usd(t['total_license_AZ'])}** "
        f"| **{_usd(t['total_support_BA'])}** | **{_usd(t['total_BB'])}** |"
    )
    return "\n".join(rows)


# ---------------------------------------------------------------------
# xlsx OUTPUT — exact Renewal Input Form column layout
# ---------------------------------------------------------------------

# Exact header row, order, and labels from the real 'Renewal Input
# Form' tab. Do not reorder, rename, or drop columns.
RENEWAL_HEADERS = {
    "A": "Product (Marketing Schedule)", "B": "Product Number", "C": "Charge Number",
    "D": "Subscription Number", "E": "Status", "F": "PL", "G": "Product Family",
    "H": "UoM", "I": "BMC Base List Price ", "J": "Partner Discounted Price",
    "K": "BMC Base List Price ", "L": "Partner Discounted Price",
    "M": "Perp or OPS   (if applicable)",
    "N": "Approximate Rate Plan Name \n(ensure correct Revenue/Service Type tags if copying)",
    "O": "PRP Number (ensure correct Serv Type <--)", "P": "Code",
    "Q": "Existing\nQty", "R": "Proposed\nQty",
    "S": "Installmt/ Ramp? (edit if needed)",
    "U": '"Att A Unit Cost" (if applicable)', "V": "TRV Override (e.g., Price Capped)",
    "W": "BRV", "X": "'Hold the Line'\nERV", "Y": "'Anchor Price'\nTRV",
    "Z": "Proposed (incl all discounts)", "AA": "Calculated",
    "AB": "BRV (ACV)", "AC": "'Hold the Line'\nERV (ACV)",
    "AD": "'Anchor Price'\nTRV (ACV)", "AE": "Proposed (incl all discounts)",
    "AF": "Calculated", "AG": "On Anchor", "AH": "Calculated",
    "AI": "Estimated ERV Deviation", "AJ": "Estimated TRV Deviation",
    "AK": '(BP Use)      "Total Annual List"', "AL": "Start Date", "AM": "End Date",
    "AN": "Approx. Term in Months", "AO": "Projected Obsolete Date",
    "AP": "Perp Support Rate", "AQ": "Est. Perp Supp Mthly List Price",
    "AR": "Hold Disc", "AS": "License\nTrans", "AT": "Lic One-Time",
    "AU": "Total License", "AV": "Supp Trans", "AW": "Supp One-Time",
    "AX": "License Net Price", "AY": "Support Net Price", "AZ": "Total License",
    "BA": "Total Support", "BB": "Total", "BC": "orig TRV ACV",
    "BD": "orig ERV ACV", "BE": "orig term length", "BF": "chosen unit list",
    "BG": "code", "BH": "sum orig TRV ACV", "BI": "sum exist qty",
    "BJ": "sum decr qty", "BK": "sum rem qty", "BL": "sum new qty",
    "BM": "orig TRV ACV unit", "BN": "chosen annual unit list",
    "BO": "override unit if any", "BP": "new TRV ACV unit", "BQ": "decrease count",
    "BR": "disc form w vlookup", "BS": "max disc", "BT": "total trans disc",
    "BU": "disc cap", "BV": "total disc excl ARG", "BW": "orig end date",
    "BX": "dedup ramp/install qty count", "BY": "dedup ramp/install qty",
    "BZ": "Obsolete During Period", "CA": "SaaS margin", "CB": "SaaS location",
    "CC": "SaaS cost code", "CD": "SaaS margin cost", "CE": "deal score type",
}
RENEWAL_LAST_COL = "CE"

# Reference-only audit columns after CE (grey header). Strip before
# pasting into the real sheet.
REFERENCE_HEADERS = {
    "CF": "Customer Tier Discount %\n(reference only)",
    "CG": "Additional License Discount %\n(on Partner Price, reference only)",
    "CH": "Effective Discount %\n(Tier + License combined = AS)",
    "CI": "Review Flags\n(reference only)",
}
REFERENCE_LAST_COL = "CI"

FMT_MONEY = '$#,##0.00'
FMT_MONEY_UNIT = '$#,##0.0000'
FMT_MONEY_WHOLE = '$#,##0'
FMT_PCT = '0.00%'
FMT_DATE = 'DD-MMM-YYYY'
FMT_QTY = '#,##0'
FMT_MONTHS = '0.00'

HEADER_FILL_HEX = "0064FF"
REF_HEADER_FILL_HEX = "7D7D7D"
TOTAL_FILL_HEX = "F2F2F2"
WHITE_HEX = "FFFFFF"
BORDER_HEX = "D9D9D9"
NOTE_HEX = "808080"
WARN_HEX = "C00000"


def write_renewal_output_xlsx(quote, output_path, start_date=None, end_date=None,
                                subscription_number=None):
    """
    quote: return value of compute_renewal_quote() (a plain list of
        compute_renewal_line() results is also accepted).
    Writes the 'Renewal Input Form' sheet (real template columns + grey
    reference columns) and a 'Summary' sheet with code-computed totals.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import column_index_from_string, get_column_letter

    if isinstance(quote, list):
        results = quote
        totals = None
        summary = None
    else:
        results, totals, summary = quote["lines"], quote["totals"], quote["summary"]

    start_date = _to_date(start_date)
    end_date = _to_date(end_date)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Renewal Input Form"

    FONT = "Arial"
    HEADER_FILL = PatternFill("solid", fgColor=HEADER_FILL_HEX)
    REF_HEADER_FILL = PatternFill("solid", fgColor=REF_HEADER_FILL_HEX)
    TOTAL_FILL = PatternFill("solid", fgColor=TOTAL_FILL_HEX)
    HEADER_FONT = Font(name=FONT, bold=True, color=WHITE_HEX, size=8)
    DATA_FONT = Font(name=FONT, size=9)
    BOLD_FONT = Font(name=FONT, bold=True, size=9)
    NA_FONT = Font(name=FONT, italic=True, size=9, color=NOTE_HEX)
    WARN_FONT = Font(name=FONT, size=9, color=WARN_HEX)
    thin = Side(style="thin", color=BORDER_HEX)
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    last_idx = column_index_from_string(RENEWAL_LAST_COL)
    ref_last_idx = column_index_from_string(REFERENCE_LAST_COL)
    hdr_row = 1
    for c in range(1, ref_last_idx + 1):
        col = get_column_letter(c)
        cell = ws.cell(row=hdr_row, column=c,
                       value=RENEWAL_HEADERS.get(col) or REFERENCE_HEADERS.get(col, ""))
        cell.fill = REF_HEADER_FILL if c > last_idx else HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[hdr_row].height = 42

    def setcell(row, col, value, font=DATA_FONT, fmt=None):
        c = ws.cell(row=row, column=column_index_from_string(col), value=value)
        c.font = font
        if fmt:
            c.number_format = fmt
        c.border = BORDER
        return c

    start_row = hdr_row + 1
    for idx, r in enumerate(results):
        row = start_row + idx
        license_type = r.get("license_type", "")
        is_ops = _is_ops(license_type)

        setcell(row, "A", r.get("product", ""))
        if subscription_number:
            setcell(row, "D", subscription_number)
        setcell(row, "E", r.get("status", ""))
        setcell(row, "F", r.get("PL", ""))
        setcell(row, "G", r.get("product_family", ""))
        setcell(row, "H", r.get("uom", ""))

        list_col, partner_col = ("K", "L") if is_ops else ("I", "J")
        setcell(row, list_col, r.get("list_price"), fmt=FMT_MONEY_UNIT)
        setcell(row, partner_col, r.get("partner_price"), fmt=FMT_MONEY_UNIT)

        setcell(row, "M", license_type)
        setcell(row, "O", r.get("prp_number", ""))
        setcell(row, "P", r.get("code", ""))
        setcell(row, "Q", r.get("existing_qty"), fmt=FMT_QTY)
        setcell(row, "R", r.get("proposed_qty"), fmt=FMT_QTY)

        if r.get("Y") is not None:
            setcell(row, "Y", r["Y"], fmt=FMT_MONEY_UNIT)
        setcell(row, "AB", r.get("orig_brv_acv"), fmt=FMT_MONEY_WHOLE)
        if r.get("AD") is not None:
            setcell(row, "AD", r["AD"], fmt=FMT_MONEY)
        if r.get("AE") is not None:
            setcell(row, "AE", r["AE"], fmt=FMT_MONEY)
        setcell(row, "AK", r.get("AK"), fmt=FMT_MONEY)

        if start_date:
            setcell(row, "AL", start_date, fmt=FMT_DATE)
        if end_date:
            setcell(row, "AM", end_date, fmt=FMT_DATE)
        setcell(row, "AN", r.get("new_term_months"), fmt=FMT_MONTHS)
        if r.get("obsolete_date"):
            setcell(row, "AO", r["obsolete_date"], fmt=FMT_DATE)

        if not is_ops:
            setcell(row, "AP", r.get("support_rate") or 0.0, fmt=FMT_PCT)
            if r.get("AQ") is not None:
                setcell(row, "AQ", r["AQ"], fmt=FMT_MONEY_UNIT)

        setcell(row, "AS", r.get("AS"), fmt='0.0000%')
        setcell(row, "AT", r.get("AT"), fmt='0.0000%')
        setcell(row, "AX", r.get("AX"), fmt='$#,##0.000000')
        setcell(row, "AY", r.get("AY"), fmt='$#,##0.000000')
        setcell(row, "AZ", r.get("AZ"), BOLD_FONT, fmt=FMT_MONEY)
        setcell(row, "BA", r.get("BA"), BOLD_FONT, fmt=FMT_MONEY)
        setcell(row, "BB", r.get("BB"), BOLD_FONT, fmt=FMT_MONEY)

        setcell(row, "BC", r.get("orig_trv_acv"), fmt=FMT_MONEY_WHOLE)
        setcell(row, "BD", r.get("orig_erv_acv"), fmt=FMT_MONEY_WHOLE)
        setcell(row, "BE", r.get("orig_term_months"), fmt=FMT_MONTHS)
        if r.get("note"):
            setcell(row, "BF", r["note"], NA_FONT)
        if r.get("BN") is not None:
            setcell(row, "BN", r["BN"], fmt=FMT_MONEY_UNIT)
        if r.get("cap_discount") is not None:
            setcell(row, "BU", r["cap_discount"], fmt=FMT_PCT)
        setcell(row, "BZ", "Yes" if r.get("obsolete_during_period") else "No")

        setcell(row, "CF", r.get("channel_discount"), fmt=FMT_PCT)
        setcell(row, "CG", r.get("additional_license_discount"), fmt=FMT_PCT)
        setcell(row, "CH", r.get("effective_discount"), BOLD_FONT, fmt=FMT_PCT)
        if r.get("warnings"):
            setcell(row, "CI", " | ".join(r["warnings"]), WARN_FONT)

        for c in range(1, ref_last_idx + 1):
            ws.cell(row=row, column=c).border = BORDER

    # TOTAL row: live SUM formulas over every money column that adds up.
    total_row = start_row + len(results)
    end_row = total_row - 1
    label = ws.cell(row=total_row, column=1, value="TOTAL")
    label.font = BOLD_FONT
    for col in ("AB", "AD", "AE", "AK", "AZ", "BA", "BB"):
        setcell(total_row, col, f"=SUM({col}{start_row}:{col}{end_row})", BOLD_FONT, fmt=FMT_MONEY)
    for c in range(1, ref_last_idx + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.border = BORDER

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["BF"].width = 40
    ws.column_dimensions["CI"].width = 50
    for c in range(1, ref_last_idx + 1):
        col = get_column_letter(c)
        if col not in ("A", "G", "H", "BF", "CI"):
            ws.column_dimensions[col].width = 13
    ws.freeze_panes = f"B{start_row}"

    if totals is not None:
        ss = wb.create_sheet("Summary")
        rows = [
            ("Subscription Number", subscription_number, None),
            ("Start Date", start_date, FMT_DATE),
            ("End Date", end_date, FMT_DATE),
            ("Lines", summary["line_count"], None),
            ("Total License (AZ)", totals["total_license_AZ"], FMT_MONEY),
            ("Total Support (BA)", totals["total_support_BA"], FMT_MONEY),
            ("Grand Total (BB)", totals["total_BB"], FMT_MONEY),
            ("List Value for Term (list × term months × qty)",
             totals["total_list_value_for_term"], FMT_MONEY),
            ("Total Annual List (AK)", totals["total_annual_list_AK"], FMT_MONEY),
            ("Overall Discount (1 − BB ÷ list value)", totals["overall_discount"], FMT_PCT),
            ("Prior Booked ACV (BRV)", totals["total_prior_brv_acv"], FMT_MONEY),
            ("Proposed ACV", totals["total_proposed_acv"], FMT_MONEY),
        ]
        for i, h in enumerate(("Metric", "Value"), start=1):
            c = ss.cell(row=1, column=i, value=h)
            c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        for i, (k, v, fmt) in enumerate(rows, start=2):
            a = ss.cell(row=i, column=1, value=k)
            b = ss.cell(row=i, column=2, value=v)
            a.font, b.font = DATA_FONT, BOLD_FONT
            a.border = b.border = BORDER
            if fmt:
                b.number_format = fmt
        note_row = len(rows) + 3
        ss.cell(row=note_row, column=1, value=summary["text"]).font = DATA_FONT
        ss.cell(row=note_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ss.merge_cells(start_row=note_row, start_column=1, end_row=note_row + 6, end_column=2)
        ss.column_dimensions["A"].width = 48
        ss.column_dimensions["B"].width = 60

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    print(f"smt_calculator.py SKILL_VERSION={SKILL_VERSION}")
    kb = KnowledgeBase()
    r2 = compute_renewal_line(kb, "BMC Helix Service Management OnPrem - Standard - Named User",
                                license_type="OPS", existing_qty=50, proposed_qty=50, code="",
                                orig_brv_acv=807, orig_trv_acv=920, orig_erv_acv=892,
                                orig_term_months=11.9667, new_term_months=11,
                                customer_tier="BMC Base List Price")
    print(f"AS={r2['AS']:.10f}  AX={r2['AX']:.4f}  BB={r2['BB']:.2f}")
