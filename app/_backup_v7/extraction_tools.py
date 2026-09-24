"""
extraction_tools.py
--------------------
Parses the two independent input sources (quote PDF, UFR export) from
local file paths, matches line items between them, and resolves
Customer Tier / Support Rate from reference tables.
"""

import csv
import os
import re

import openpyxl
import pypdf

REFERENCES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "references")

# Only mappings confirmed against real data. Anything else must come
# back as "needs_confirmation" rather than being guessed.
CONFIRMED_TIER_MAPPING = {
    "GOSI": "GOSI Strategic",
}


def _normalize(name: str) -> str:
    return "".join(str(name).split()).lower()


def _to_number(value):
    """Coerces a cell value or tool-call argument into a number.
    Returns 0 for blank/None."""
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)):
        return value
    try:
        cleaned = str(value).replace(",", "").strip()
        return float(cleaned) if "." in cleaned else int(cleaned)
    except (ValueError, TypeError):
        return 0


def _clean_pdf_text(value: str) -> str:
    """Repairs line-wrap artifacts from pypdf's column extraction before
    a captured product/UoM string is used for matching.

    Rejoins a word broken mid-hyphen with no space before the hyphen
    (e.g. "Add-\\non" -> "Add-on"), then collapses all remaining
    whitespace/newlines to single spaces. A genuine " - " separator
    that falls at a line break is left intact by the whitespace
    collapse (e.g. "Solutions - \\nMarketZone" -> "Solutions - MarketZone").
    """
    value = re.sub(r"([A-Za-z0-9])-\s*\n\s*", r"\1-", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


_MONTH_ABBREV = {m: i + 1 for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
)}


def _parse_quote_date(cleaned_date_str: str):
    """Parses a 'DD-MON-YYYY' date (the quote PDF's format) into a
    datetime.date. Expects _clean_pdf_text has already been applied."""
    from datetime import date
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", cleaned_date_str.strip())
    if not m:
        raise ValueError(f"Could not parse quote date {cleaned_date_str!r} as DD-MON-YYYY.")
    day, mon, year = m.groups()
    mon_num = _MONTH_ABBREV.get(mon.upper())
    if mon_num is None:
        raise ValueError(f"Unrecognized month abbreviation {mon!r} in date {cleaned_date_str!r}.")
    return date(int(year), mon_num, int(day))


def _whole_month_term(start_date, end_date) -> int:
    """Whole-month term length, end-date inclusive (30-Jun-2026 to
    29-Jun-2027 is 12 months). Equivalent to Excel's
    DATEDIF(start, end+1day, "m")."""
    from datetime import timedelta
    end_exclusive = end_date + timedelta(days=1)
    months = (end_exclusive.year - start_date.year) * 12 + (end_exclusive.month - start_date.month)
    if end_exclusive.day < start_date.day:
        months -= 1
    return months


# ---------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------

def extract_quote_pdf(pdf_path: str) -> list[dict]:
    """Extracts line items from a budgetary quote PDF at a local file
    path. Tolerates the Support Plan column ("BMC Continuous Support")
    being split across up to three lines by pypdf, comma-formatted or
    line-wrapped unit costs, and multi-line Unit of Measure text.
    Captures the Term column and returns term_start/term_end/
    new_term_months per line.

    Currently anchors on the literal text "BMC Continuous Support" for
    the Support Plan column; a different Support Plan text will not match.
    """
    reader = pypdf.PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    lines = []
    row_pattern = re.compile(
        r"([A-Za-z][A-Za-z0-9\-\s/]+?)\s*"
        r"BMC\s*\n?\s*Continuous\s*\n?\s*Support\s*"
        r"(\d{1,2}-[A-Za-z]{3}-\d{4})\s*to\s*(\d{1,2}-[A-Za-z]{3}-\s*\n?\s*\d{4})\s*"
        r"per\s+([a-zA-Z0-9\-\s]+?)\s*"
        r"([\d,]+)\s*"
        r"USD\s*([\d,.]+?)\s*"
        r"USD\s*([\d,]+\.\d{2})",
        re.DOTALL,
    )
    for match in row_pattern.finditer(text):
        product, term_start_raw, term_end_raw, uom, qty, unit_cost, fee = match.groups()
        term_start = _parse_quote_date(_clean_pdf_text(term_start_raw))
        term_end = _parse_quote_date(_clean_pdf_text(term_end_raw))
        lines.append({
            "product": _clean_pdf_text(product),
            "uom": _clean_pdf_text(uom),
            "term_start": term_start.isoformat(),
            "term_end": term_end.isoformat(),
            "new_term_months": _whole_month_term(term_start, term_end),
            "quantity": _to_number(qty),
            "unit_cost": _to_number(unit_cost),
            "fee": _to_number(fee),
        })
    if not lines:
        raise ValueError(
            "No line items extracted from PDF — table format may differ "
            "from the expected BMC budgetary quote layout. Flag for review."
        )
    return lines


# ---------------------------------------------------------------------
# UFR extraction
# ---------------------------------------------------------------------

def _parse_ufr_date(value):
    """Parses a UFR 'Term Start/End Date' cell into a datetime.date.
    Handles both a real datetime.date/datetime cell and a text string
    like '30-Jun-2026'."""
    from datetime import date, datetime
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", str(value).strip())
    if not m:
        raise ValueError(f"Could not parse UFR date {value!r} as DD-Mon-YYYY or a real date cell.")
    day, mon, year = m.groups()
    mon_num = _MONTH_ABBREV.get(mon.upper())
    if mon_num is None:
        raise ValueError(f"Unrecognized month abbreviation {mon!r} in UFR date {value!r}.")
    return date(int(year), mon_num, int(day))


def parse_ufr_export(xlsx_path: str, subscription_number: str) -> list[dict]:
    """Parses a UFR export xlsx at a local file path, scoped to one
    subscription_number (a real export spans many subscriptions and
    customers).

    'Subscription Number' and 'Subscription Owner Account CSN' are
    carried forward across blank rows, since the real export only
    populates them on the first row of each subscription's block.

    orig_term_months is derived from 'Term Start Date' / 'Term End
    Date' using the same whole-month convention as the quote PDF's
    new_term_months. 'Support Tier' is also captured, for
    get_support_rate().
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    required = [
        "Subscription Number", "Marketing Schedule Name", "Service Type",
        "RPC Quantity", "Renewal Qty Number Of Licenses",
        "BRV ACV Local", "ERV ACV Local", "TRV ACV Local",
        "Subscription Owner Account CSN",
        "Term Start Date", "Term End Date", "Support Tier",
    ]
    missing = [r for r in required if r not in col]
    if missing:
        raise ValueError(f"UFR export missing expected column(s): {missing}")

    records = []
    last_subscription_number = None
    last_account_csn = None

    for r in range(2, ws.max_row + 1):
        raw_sub = ws.cell(row=r, column=col["Subscription Number"]).value
        raw_csn = ws.cell(row=r, column=col["Subscription Owner Account CSN"]).value

        if raw_sub not in (None, ""):
            last_subscription_number = raw_sub
        if raw_csn not in (None, ""):
            last_account_csn = raw_csn

        product = ws.cell(row=r, column=col["Marketing Schedule Name"]).value
        if not product:
            continue
        if str(last_subscription_number).strip() != str(subscription_number).strip():
            continue

        orig_term_start = _parse_ufr_date(ws.cell(row=r, column=col["Term Start Date"]).value)
        orig_term_end = _parse_ufr_date(ws.cell(row=r, column=col["Term End Date"]).value)
        orig_term_months = (
            _whole_month_term(orig_term_start, orig_term_end)
            if orig_term_start and orig_term_end else None
        )

        records.append({
            "product": product,
            "service_type": ws.cell(row=r, column=col["Service Type"]).value,
            "existing_qty": _to_number(ws.cell(row=r, column=col["RPC Quantity"]).value),
            "renewal_qty": _to_number(ws.cell(row=r, column=col["Renewal Qty Number Of Licenses"]).value),
            "brv_acv": _to_number(ws.cell(row=r, column=col["BRV ACV Local"]).value),
            "erv_acv": _to_number(ws.cell(row=r, column=col["ERV ACV Local"]).value),
            "trv_acv": _to_number(ws.cell(row=r, column=col["TRV ACV Local"]).value),
            "account_csn": last_account_csn,
            "subscription_number": last_subscription_number,
            "orig_term_start": orig_term_start.isoformat() if orig_term_start else None,
            "orig_term_end": orig_term_end.isoformat() if orig_term_end else None,
            "orig_term_months": orig_term_months,
            "support_tier": ws.cell(row=r, column=col["Support Tier"]).value,
        })

    if not records:
        raise ValueError(
            f"No UFR rows found for subscription_number={subscription_number!r}. "
            "Check the value matches exactly what's in the export's "
            "'Subscription Number' column."
        )
    return records


# ---------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------

def match_quote_to_ufr(quote_lines: list[dict], ufr_records: list[dict]) -> dict:
    """Groups quote_lines by product, matches each against its UFR
    record (already scoped to one subscription), and returns
    resolved/ambiguous match results.

    A quantity increase (PDF qty > existing qty) is modeled as ONE row:
    existing_qty is carried forward as-is, proposed_qty is
    existing_qty + pdf_qty, Code="a".

    Returns {"lines": [...], "suggested_batch_resolution": {...} | None}.
    suggested_batch_resolution, when present, covers every line whose
    PDF quantity equals its existing UFR quantity — the caller should
    confirm this pattern once for the whole batch rather than asking
    about each line individually.
    """
    ufr_index = {_normalize(r["product"]): r for r in ufr_records}

    totals = {}
    for line in quote_lines:
        key = _normalize(line["product"])
        totals.setdefault(key, {
            "product": line["product"], "quantity": 0, "lines": [],
            "term_start": line.get("term_start"), "term_end": line.get("term_end"),
            "new_term_months": line.get("new_term_months"),
            "term_mismatch": False,
        })
        agg = totals[key]
        agg["quantity"] += _to_number(line["quantity"])
        agg["lines"].append(line)
        if (line.get("new_term_months") is not None
                and agg["new_term_months"] is not None
                and line["new_term_months"] != agg["new_term_months"]):
            agg["term_mismatch"] = True

    results = []
    for key, agg in totals.items():
        ufr = ufr_index.get(key)
        if ufr is None:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "scope_flag": "possible_new_business",
                "reason": f"'{agg['product']}' has no matching UFR record for this "
                          "subscription. Either this is a genuinely NEW product for "
                          "this customer, or the name doesn't match the UFR export "
                          "(typo/formatting difference).",
                "guardrail": "If this is a new product, it belongs on the SMT NEW "
                             "INPUT FORM (Part 1 — new business, SDN/Disnet/GSS "
                             "discount model), which is OUT OF SCOPE for this agent. "
                             "Do NOT attempt to compute a renewal line for it. Tell "
                             "the user explicitly that this product needs to go "
                             "through the New Input Form process instead, and ask "
                             "them to confirm whether it's a genuine new product or "
                             "a name-matching issue with an existing UFR line.",
            })
            continue

        existing_qty = _to_number(ufr["existing_qty"])
        pdf_qty = _to_number(agg["quantity"])

        term_fields = {
            "term_start": agg.get("term_start"),
            "term_end": agg.get("term_end"),
            "new_term_months": agg.get("new_term_months"),
        }
        if agg.get("term_mismatch"):
            term_fields["term_warning"] = (
                f"'{agg['product']}' has multiple PDF lines with DIFFERENT terms — "
                "do not assume they match. Confirm the correct term with the user "
                "before calling compute_renewal_line."
            )

        if pdf_qty > existing_qty:
            results.append({
                "product": agg["product"], "status": "resolved",
                **term_fields,
                "rows": [
                    {
                        "code": "a",
                        "existing_qty": existing_qty,
                        "proposed_qty": existing_qty + pdf_qty,
                        "note": "renewal with growth: existing quantity carried "
                                "forward plus the PDF's added quantity",
                    },
                ],
                "ufr": ufr,
            })
        elif pdf_qty == existing_qty:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "ambiguity_type": "qty_equal",
                **term_fields,
                "reason": f"PDF quantity ({pdf_qty}) equals UFR existing quantity "
                          f"({existing_qty}) — could be a flat renewal or the growth "
                          "amount being added. Confirm which applies.",
                "ufr": ufr,
                "existing_qty": existing_qty,
                "pdf_qty": pdf_qty,
            })
        else:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "ambiguity_type": "qty_decrease",
                **term_fields,
                "reason": f"PDF quantity ({pdf_qty}) is less than UFR existing "
                          f"quantity ({existing_qty}) — could be Code 'd' (decrease) "
                          "or 'x' (exclude). Confirm which applies.",
                "ufr": ufr,
                "existing_qty": existing_qty,
                "pdf_qty": pdf_qty,
            })

    flat_candidates = [r for r in results if r.get("ambiguity_type") == "qty_equal"]
    suggested_batch_resolution = None
    if flat_candidates:
        suggested_batch_resolution = {
            "pattern": "flat_renewal_likely",
            "count": len(flat_candidates),
            "products": [r["product"] for r in flat_candidates],
            "message": (
                f"{len(flat_candidates)} product(s) have quote quantity equal to "
                "their existing UFR quantity, which usually means a flat renewal "
                "(no quantity change). Ask the user ONCE: 'It looks like "
                f"{len(flat_candidates)} product(s) — {', '.join(r['product'] for r in flat_candidates)} — "
                "are flat renewals (quantity unchanged). Should I treat all of "
                "these as flat renewals?' Do not ask about each product "
                "individually when they all share this same pattern."
            ),
        }

    return {"lines": results, "suggested_batch_resolution": suggested_batch_resolution}


def extract_and_match_quote(pdf_path: str, xlsx_path: str, subscription_number: str) -> dict:
    """Single combined tool call: extracts the PDF, parses the UFR
    export scoped to one subscription, and matches them.

    Returns {"lines": [...], "suggested_batch_resolution": {...} | None} —
    see match_quote_to_ufr for the shape of each entry in "lines" and
    how to use suggested_batch_resolution."""
    quote_lines = extract_quote_pdf(pdf_path)
    ufr_records = parse_ufr_export(xlsx_path, subscription_number)
    return match_quote_to_ufr(quote_lines, ufr_records)


# ---------------------------------------------------------------------
# Customer Tier resolution
# ---------------------------------------------------------------------

def _read_csv_dicts(filename: str) -> list[dict]:
    path = os.path.join(REFERENCES_DIR, filename)
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def get_valid_customer_tier_columns() -> list[str]:
    """Returns the real Customer Tier column names from
    channel_discount_table.csv. Call this before offering the user a
    list of tier choices, so only real column names are presented."""
    rows = _read_csv_dicts("channel_discount_table.csv")
    if not rows:
        raise ValueError("channel_discount_table.csv is empty or missing headers.")
    fieldnames = list(rows[0].keys())
    return [c for c in fieldnames[1:] if c]


def get_customer_tier(account_csn) -> dict:
    """Looks up account_csn in account_csn_tier_mapping.csv, then maps
    the raw 'Tier Global Helix' value to a real channel_discount_table.csv
    column via CONFIRMED_TIER_MAPPING. Only confirmed mappings are
    auto-resolved; everything else returns needs_confirmation."""
    if account_csn in (None, ""):
        return {
            "status": "needs_confirmation",
            "reason": "No account_csn available for this row (UFR export "
                      "didn't carry an Account CSN forward for it). Ask the "
                      "user to pick from get_valid_customer_tier_columns().",
        }

    rows = _read_csv_dicts("account_csn_tier_mapping.csv")
    match = next(
        (r for r in rows if str(r.get("Account CSN", "")).strip() == str(account_csn).strip()),
        None,
    )
    if match is None:
        return {
            "status": "needs_confirmation",
            "reason": f"account_csn={account_csn!r} not found in "
                      "account_csn_tier_mapping.csv. Ask the user to pick "
                      "from get_valid_customer_tier_columns().",
        }

    raw_tier = (match.get("Tier Global Helix") or "").strip()
    resolved = CONFIRMED_TIER_MAPPING.get(raw_tier)
    if resolved is None:
        return {
            "status": "needs_confirmation",
            "raw_tier": raw_tier,
            "reason": f"Tier Global Helix={raw_tier!r} has no confirmed mapping "
                      "to a channel_discount_table.csv column yet. Ask the user "
                      "to pick from get_valid_customer_tier_columns().",
        }

    return {"status": "resolved", "raw_tier": raw_tier, "customer_tier": resolved}


# ---------------------------------------------------------------------
# Support Rate resolution
# ---------------------------------------------------------------------

def get_support_rate(support_tier: str) -> dict:
    """Looks up the real support rate for a Support Tier from
    support_rate_table.csv. Rate varies by tier and only applies to
    Perpetual lines."""
    rows = _read_csv_dicts("support_rate_table.csv")
    match = next(
        (r for r in rows if _normalize(r.get("Support Tier", "")) == _normalize(support_tier)),
        None,
    )
    if match is None:
        return {
            "status": "needs_confirmation",
            "reason": f"support_tier={support_tier!r} not found in "
                      "support_rate_table.csv. Confirm the correct rate with "
                      "the user before computing this line.",
        }
    return {"status": "resolved", "support_rate": _to_number(match.get("Rate"))}
