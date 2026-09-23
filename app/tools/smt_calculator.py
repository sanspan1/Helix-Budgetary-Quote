"""
smt_calculator.py
------------------
Deterministic implementation of the BMC Helix SMT's traced formula
chains. This exists so the agent CALLS this code for arithmetic
instead of reasoning through multi-table lookups and discount math
in prose — that approach was tested and failed twice (a dropped
discount factor, then a near-empty output).

Two entry points:
  - compute_new_business_line(...)  -> Part 1, New Input Form
  - compute_renewal_line(...)       -> Part 2, Renewal Input Form

Both raise ValueError with a clear message when a required input is
missing or a product can't be matched — the caller (the agent) should
surface that as a flag to the user, never silently substitute a guess.
"""

SKILL_VERSION = 5  # bump this every time this file changes; check it
                    # with `python3 -c "from smt_calculator import SKILL_VERSION; print(SKILL_VERSION)"`
                    # to confirm which version is actually on disk, since
                    # downloaded zips have repeatedly landed with stale
                    # filenames like "(1)"/"(2)" that don't reflect content.

import csv
import os
from datetime import date, datetime

REF_DIR = os.path.join(os.path.dirname(__file__), "..", "references")


def _load_csv(filename):
    path = os.path.join(REF_DIR, filename)
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _normalize(name):
    return "".join(name.split()).lower()


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

        self._price_index = {}
        for row in self.price_list:
            key = (_normalize(row["Product Name"]), row["License Type"])
            self._price_index[key] = row

        self._obsoleting_index = {
            _normalize(r["Product Name"]): r["Projected Obsolete Date"]
            for r in self.obsoleting
        }
        self._cap_index = {
            _normalize(r["Product Name"]): float(r["Cap Discount %"])
            for r in self.cap_discount
        }

    def find_product(self, product_name, license_type):
        """Exact match required (after whitespace normalization).
        Returns None if not found — caller must flag, not guess."""
        key = (_normalize(product_name), license_type)
        return self._price_index.get(key)

    def channel_discount(self, product_family, customer_tier):
        for row in self.channel_table:
            if row["PF"] == product_family:
                val = row.get(customer_tier)
                if val not in (None, ""):
                    return float(val)
        return None

    def sdn_discount(self, product_line, deal_type, license_type):
        """deal_type: 'SDN' / 'Disnet' / 'GSS'. license_type: 'OPS' or 'Perpetual'."""
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
    deal_type: 'SDN', 'Disnet', 'GSS', or None (if None, AS/AT cannot be
               computed via the SDN branch — this function raises rather
               than silently falling into the renewal-else-branch bug
               we found, since that branch requires Part 2 data this
               function doesn't have).
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
        # AND(...) gate fails -> real sheet falls to a renewal-Target-Value
        # else-branch that evaluates to AS=1 for lines with no renewal
        # baseline. That's a structural gap for new-business lines with
        # no Existing Qty. Surface it, don't silently compute a wrong number.
        raise ValueError(
            f"BW=0 for Product Line '{product_line}' under Deal Type '{deal_type}' — "
            "the real sheet's formula falls into a renewal-baseline branch here "
            "that doesn't apply to new business. Flag this line for manual review "
            "rather than computing AS/AX automatically."
        )

    AS = 1 - (1 - bw) / (1 - channel_disc)
    AT = 1 - (1 - (0.9999 if deal_type == "GSS" else 0.9998)) / (1 - bw)

    AX = partner_price * (1 - AS) * (1 - AT)
    AZ = AX * term_months * qty
    BB = AZ  # BA (Total Support) = 0 for OPS in this model
    AK = list_price * 12 * qty  # pre-discount annual benchmark

    obs_date = kb.obsolete_date(product_name)

    return {
        "product": product_name, "license_type": license_type,
        "product_line": product_line, "product_family": product_family,
        "list_price": list_price, "partner_price": partner_price,
        "channel_discount": channel_disc, "BW": bw, "AS": AS, "AT": AT,
        "AX": AX, "AZ": AZ, "BA": 0.0, "BB": BB, "AK": AK,
        "obsolete_date": obs_date,
    }


# ---------------------------------------------------------------------
# PART 2 — Renewal Input Form (existing subscription)
# ---------------------------------------------------------------------

def compute_renewal_line(kb: KnowledgeBase, product_name, license_type,
                           existing_qty, proposed_qty, code,
                           orig_brv_acv, orig_trv_acv, orig_erv_acv, orig_term_months,
                           new_term_months, customer_tier,
                           support_rate=None,
                           obsolete_date=None, max_term_end_date_across_quote=None,
                           new_term_years_rounded_across_quote=None,
                           u_override=None, v_override=None, att_a_unit_cost=None):
    """
    code: 'a' (add), 'd' (decrease), 'x' (exclude), or '' (straight renewal)
    customer_tier: the exact channel_discount_table.csv column name
        (e.g. "GOSI Strategic", "Reseller (Customer Tier 1-3)"), OR the
        literal sentinel "Helix Base List Price" which disables channel
        discount entirely (0%, hardcoded, not a lookup). Confirmed from
        the real B4=Lists!I3 formula: Part 2 uses the SAME ChannelTable
        lookup as Part 1, keyed by this value and the product's Product
        Family — this was previously NOT wired in (basis was hardcoded
        to raw list_price regardless of tier), which only produced
        correct output for lines whose real tier happened to be the
        sentinel. A real, nonzero discount (e.g. GOSI Strategic = 45%
        for ITOM/ITSM/Remediate On Prem) would have been silently
        ignored before this fix.
    orig_brv_acv: prior period's actual BOOKED ACV ($) for this line —
        column AB in the real sheet. A raw historical input, distinct
        from orig_trv_acv/orig_erv_acv below.
    orig_trv_acv / orig_erv_acv / orig_term_months: from the customer's
        OWN prior-period UFR/subscription record. Cannot be derived from
        the current quote PDF. Raise if not supplied.
    support_rate: REQUIRED for Perpetual lines (meaningless for OPS/SaaS,
        pass None). Confirmed by Helix: this is NOT a fixed 20% — it
        varies by Support Tier: BMC Continuous Support=20%, Partner
        Continuous Support L0=19%, Partner Lite=17%, Partner Continuous
        Support L1=16%, Partner Emerging Market L1=14%. Look this up via
        get_support_rate(support_tier) before calling — previously this
        was hardcoded to 0.20 internally regardless of the real tier,
        which silently overcharged/undercharged every non-"BMC
        Continuous Support" Perpetual line. Raises if a Perpetual line
        is called without this.
    obsolete_date / max_term_end_date_across_quote / new_term_years_rounded_across_quote:
        feed the real sheet's growth-factor CAPPING condition
        (AND(AO<CG$11, ROUND(AN/12,0)<NewTL) -> MIN(factor,1) instead of
        the raw factor). CG$11 and NewTL are WHOLE-QUOTE aggregates (max
        term-end date and rounded new-term-years across every line in
        the renewal sheet), not per-line values — the caller must compute
        these once across all lines and pass them into every call. If
        omitted, the raw (uncapped) factor is used and this function
        will NOT silently apply the cap — pass these explicitly for any
        quote where correctness matters, since this branch was already
        found to be silently skipped once.
    """
    if orig_brv_acv is None or orig_trv_acv is None or orig_erv_acv is None or orig_term_months is None:
        raise ValueError(
            "Prior-period BRV/TRV/ERV ACV and term data are required for a "
            "renewal calculation and cannot come from the current quote PDF. "
            "Ask for the customer's UFR/prior-subscription export before proceeding."
        )
    if license_type != "OPS" and code not in ("d", "x") and support_rate is None:
        raise ValueError(
            f"support_rate is required for a Perpetual line ('{product_name}') "
            "— it is NOT a fixed 20%. Call get_support_rate(support_tier) "
            "first and pass its result here."
        )

    row = kb.find_product(product_name, license_type)
    if row is None:
        raise ValueError(f"No exact price-list match for '{product_name}' ({license_type}).")

    list_price = float(row["List Price (USD)"])
    product_family = row["Product Family"]

    # Whole-dollar rounding — now enforced here in code, not left to the
    # caller/LLM to remember. The real sheet stores these as whole-
    # dollar figures; feeding raw UFR decimals produced a small but
    # real drift (confirmed: AS off by ~0.00003 for a real example).
    # round() is idempotent, so already-rounded input is unaffected.
    orig_brv_acv = round(orig_brv_acv)
    orig_trv_acv = round(orig_trv_acv)
    orig_erv_acv = round(orig_erv_acv)

    # "Remediate On Prem" family convention — confirmed from the real
    # Amdocs quote across all 6 TrueSight/Remediate products: BOTH
    # orig_trv_acv AND orig_erv_acv are set to the rounded ERV value
    # (TRV's own raw column is not used). Previously only an agent
    # instruction the LLM had to remember to apply per-product; now
    # enforced here unconditionally for any product in this family.
    if product_family == "Remediate On Prem":
        orig_trv_acv = orig_erv_acv

    common_meta = {
        "product": product_name, "license_type": license_type,
        "PL": row["Product Line (PL)"], "product_family": product_family,
        "uom": row.get("Unit of Measure", ""), "status": row.get("Status", ""),
        "list_price": list_price, "new_term_months": new_term_months,
        "orig_brv_acv": orig_brv_acv, "orig_trv_acv": orig_trv_acv,
        "orig_erv_acv": orig_erv_acv, "orig_term_months": orig_term_months,
        "support_rate": support_rate,
    }

    if code in ("d", "x"):
        # Real sheet: Y/AC/AD blank, AZ/BA forced to 0, BB=0.
        return {
            **common_meta,
            "code": code, "existing_qty": existing_qty, "proposed_qty": 0,
            "AS": 0.0, "AT": 0.0, "AU": 0.0, "AX": list_price, "AY": 0.0,
            "AZ": 0.0, "BA": 0.0, "BB": 0.0, "ERVFactor": None, "TRVFactor": None,
            "AQ": None,
            "note": "Code is decrease/exclude — real sheet zeroes AZ/BA/BB. "
                    "If this quantity reappears as a new OPS/Perpetual line "
                    "elsewhere in the same quote, flag as a possible license "
                    "conversion, not a genuine reduction.",
        }

    # Growth factors (old-term-years -> new-term-years lookup)
    old_yrs = orig_term_months / 12
    new_yrs = new_term_months / 12
    erv_factor = kb.growth_factor(kb.ervfactor_table, old_yrs, new_yrs)
    trv_factor_raw = kb.growth_factor(kb.trvfactor_table, old_yrs, new_yrs)

    # Real formula: IF(AND(AO < CG$11, ROUND(AN/12,0) < NewTL), MIN(factor,1), factor)
    # AO blank (no obsolescence entry) behaves as "greater than any date" in
    # Excel (text > any number/date), so the AND is FALSE and the raw
    # (uncapped) factor applies whenever obsolete_date is None.
    cap_condition = (
        obsolete_date is not None
        and max_term_end_date_across_quote is not None
        and new_term_years_rounded_across_quote is not None
        and obsolete_date < max_term_end_date_across_quote
        and round(new_yrs) < new_term_years_rounded_across_quote
    )
    trv_factor = min(trv_factor_raw, 1) if cap_condition else trv_factor_raw

    # BN (ceiling reference) needs list-basis; for OPS: list_price*12, Perpetual: list_price*support_rate
    bn = list_price * 12 if license_type == "OPS" else list_price * support_rate
    bo = max(filter(None, [u_override, v_override])) if (u_override or v_override) else bn

    brv_unit = orig_brv_acv / existing_qty if existing_qty else 0.0   # AB/Q, real formula's "W" column
    trv_unit = orig_trv_acv / existing_qty if existing_qty else 0.0   # BC/Q, real formula's "BM" column

    # Y (Anchor Price TRV, per-unit) = MIN( brv_unit + (trv_unit - brv_unit) * factor(capped or raw), BN, BO )
    y = min(brv_unit + (trv_unit - brv_unit) * trv_factor, bn, bo)
    ad = y * proposed_qty  # Anchor TRV ACV
    ae = ad  # Proposed ACV = straight copy

    at = 0.0  # optional manual override field in this form; default 0 unless supplied
    denom = (list_price * 12 * proposed_qty) if license_type == "OPS" else (list_price * support_rate * proposed_qty)
    AS = 1 - (ae / denom) / (1 - at) if denom else 0.0

    basis = list_price  # default: sentinel active, or no channel row for this tier/family
    if customer_tier != "Helix Base List Price":
        row_for_family = row  # already the matched price-list row (has Product Family)
        channel_disc = kb.channel_discount(row_for_family["Product Family"], customer_tier)
        if channel_disc is not None:
            basis = list_price * (1 - channel_disc)
        # else: no channel-table row for this tier/family combo — falls
        # back to raw list_price. This should be rare; flag if it
        # happens for a tier that's supposed to have full coverage.
    AX = basis * (1 - AS) * (1 - at)

    if license_type == "OPS":
        AZ = AX * new_term_months * proposed_qty
        BA = 0.0
        AQ = None
    else:
        AZ = AX * max(proposed_qty - existing_qty, 0)
        # Perpetual support fee (AQ -> AV -> AY -> BA). This was previously
        # hardcoded to 0.0, which silently understated every Perpetual
        # renewal's total by its support component — caught via a real
        # example (BMC Discovery: BA should be $3.67, not $0.00). Then
        # later found to be hardcoded to a fixed 20% regardless of the
        # real Support Tier — now takes the confirmed per-tier rate.
        AQ = list_price * support_rate / 12   # AP (support rate) — now tier-based, not fixed
        AV = AS                       # AV = AS for Perpetual lines
        AW = 0.0                      # no manual support-discount override modeled
        AY = AQ * (1 - AV) * (1 - AW)
        BA = AY * proposed_qty * new_term_months

    BB = AZ + BA

    return {
        **common_meta,
        "code": code, "existing_qty": existing_qty, "proposed_qty": proposed_qty,
        "ERVFactor": erv_factor, "TRVFactor": trv_factor, "AQ": AQ,
        "AS": AS, "AT": at, "AU": None, "AX": AX,
        "AY": (AY if license_type != "OPS" else 0.0),
        "AZ": AZ, "BA": BA, "BB": BB,
    }



# ---------------------------------------------------------------------
# xlsx OUTPUT — exact Renewal Input Form column layout
# ---------------------------------------------------------------------

# Exact header row, exact order, exact labels — copied verbatim from the
# real 'Renewal Input Form' tab's own header row. Do not reorder, rename,
# or drop columns: this must be directly paste-compatible into the real
# sheet. Columns not listed here (blank spacer columns) are intentional.
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


def write_renewal_output_xlsx(results, output_path, start_date=None, end_date=None,
                                subscription_number=None):
    """
    results: list of dicts, each the return value of compute_renewal_line(),
        in the order they should appear as rows.
    output_path: where to save the .xlsx.
    start_date / end_date: the quote's contract dates (applied to every
        row's Start Date / End Date columns — these are per-quote, not
        per-line, unless individual lines genuinely have different terms).
    subscription_number: applied to every row's Subscription Number
        column if provided (all lines in one renewal quote normally
        share one subscription number).

    Populates only the columns this calculator actually computes or was
    given as input. Every other column is left genuinely blank — this
    is meaningful (matches what the real sheet would show for columns
    this tool doesn't model, like BH-BY cross-row rollups or the
    formula-tampering/duplicate-detector checks) and must not be filled
    with a guessed or zeroed value.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import column_index_from_string, get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Renewal Input Form"

    FONT = "Arial"
    HEADER_FILL = PatternFill("solid", fgColor="0064FF")
    HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=8)
    DATA_FONT = Font(name=FONT, size=9)
    NA_FONT = Font(name=FONT, italic=True, size=9, color="808080")
    thin = Side(style="thin", color="D9D9D9")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    last_idx = column_index_from_string(RENEWAL_LAST_COL)
    hdr_row = 1
    for c in range(1, last_idx + 1):
        col = get_column_letter(c)
        cell = ws.cell(row=hdr_row, column=c, value=RENEWAL_HEADERS.get(col, ""))
        cell.fill = HEADER_FILL
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
        is_ops = license_type == "OPS"

        setcell(row, "A", r.get("product", ""))
        if subscription_number:
            setcell(row, "D", subscription_number)
        setcell(row, "E", r.get("status", ""))
        setcell(row, "F", r.get("PL", ""))
        setcell(row, "G", r.get("product_family", ""))
        setcell(row, "H", r.get("uom", ""))

        list_price = r.get("list_price")
        if license_type == "Perpetual":
            setcell(row, "I", list_price, fmt='$#,##0.00000')
        elif is_ops:
            setcell(row, "K", list_price, fmt='$#,##0.00000')

        setcell(row, "M", license_type)
        setcell(row, "P", r.get("code", ""))
        setcell(row, "Q", r.get("existing_qty"))
        setcell(row, "R", r.get("proposed_qty"))

        # AB (BRV ACV) is the prior-period actual booked value, supplied
        # directly by the caller as orig_brv_acv.
        if r.get("orig_brv_acv") is not None:
            setcell(row, "AB", r["orig_brv_acv"], fmt='$#,##0')

        setcell(row, "BC", r.get("orig_trv_acv"), fmt='$#,##0')
        setcell(row, "BD", r.get("orig_erv_acv"), fmt='$#,##0')
        setcell(row, "BE", r.get("orig_term_months"), fmt='0.0000')

        if start_date:
            setcell(row, "AL", start_date, fmt='DD-MMM-YYYY')
        if end_date:
            setcell(row, "AM", end_date, fmt='DD-MMM-YYYY')
        setcell(row, "AN", r.get("new_term_months"), fmt='0.0000')

        if license_type == "Perpetual":
            setcell(row, "AP", r.get("support_rate") or 0.0, fmt='0%')
            if r.get("AQ") is not None:
                setcell(row, "AQ", r["AQ"], fmt='$#,##0.00000')

        if r.get("AS") is not None:
            setcell(row, "AS", r["AS"], fmt='0.0000%')
        if r.get("AT") is not None:
            setcell(row, "AT", r["AT"], fmt='0.0000%')
        if r.get("AX") is not None:
            setcell(row, "AX", r["AX"], fmt='$#,##0.000000')
        if r.get("AY") is not None:
            setcell(row, "AY", r["AY"], fmt='$#,##0.000000')
        setcell(row, "AZ", r.get("AZ"), Font(name=FONT, bold=True, size=9), fmt='$#,##0.00')
        setcell(row, "BA", r.get("BA"), Font(name=FONT, bold=True, size=9), fmt='$#,##0.00')
        setcell(row, "BB", r.get("BB"), Font(name=FONT, bold=True, size=9), fmt='$#,##0.00')

        if r.get("note"):
            setcell(row, "BF", r["note"], NA_FONT)

        for c in range(1, last_idx + 1):
            ws.cell(row=row, column=c).border = BORDER

    total_row = start_row + len(results)
    ws.cell(row=total_row, column=2, value="TOTAL").font = Font(bold=True)
    bb_col = column_index_from_string("BB")
    tot = ws.cell(row=total_row, column=bb_col,
                   value=f"=SUM({get_column_letter(bb_col)}{start_row}:{get_column_letter(bb_col)}{total_row-1})")
    tot.font = Font(bold=True)
    tot.number_format = '$#,##0.00'

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["BF"].width = 50
    for c in range(1, last_idx + 1):
        col = get_column_letter(c)
        if col not in ("A", "G", "H", "BF"):
            ws.column_dimensions[col].width = 11
    ws.freeze_panes = f"A{start_row}"

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    # Basic smoke test — confirms the script runs and both code paths
    # execute without errors. This does NOT verify correctness; it only
    # confirms nothing crashes. Do not add expected/known-correct values
    # here — this file ships inside the agent's own skill folder, and an
    # agent that can read its own source code would see them.
    print(f"smt_calculator.py SKILL_VERSION={SKILL_VERSION}")
    print()
    kb = KnowledgeBase()

    print("=== Part 1 smoke test: BMC Discovery, SDN, Reseller Tier 1-3 ===")
    r = compute_new_business_line(kb, "BMC Discovery - Resource Unit", qty=1,
                                    license_type="OPS", customer_tier="Reseller (Customer Tier 1-3)",
                                    deal_type="SDN")
    print(f"AX={r['AX']:.6f}  BB={r['BB']:.4f}")

    print()
    print("=== Part 2 smoke test: Named User, straight renewal ===")
    r2 = compute_renewal_line(kb, "BMC Helix Service Management OnPrem - Standard - Named User",
                                license_type="OPS", existing_qty=50, proposed_qty=50, code="",
                                orig_brv_acv=807, orig_trv_acv=920, orig_erv_acv=892, orig_term_months=11.9667,
                                new_term_months=11, customer_tier="Helix Base List Price")
    print(f"AS={r2['AS']:.10f}")
    print(f"AX={r2['AX']:.4f}  BB={r2['BB']:.2f}")

