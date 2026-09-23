"""
extraction_tools.py
--------------------
Parses the two independent input sources (quote PDF, UFR export) from
local file paths. Matches line items between them.

NOTE: chat-uploaded files are NOT used here — `agents-cli playground`
(and `adk web`) do not persist chat attachments as loadable artifacts
by default (known ADK limitation, tracked upstream). Files must be
placed in a local folder (e.g. data/) and referenced by path.

CRITICAL: a real UFR export contains rows for MANY subscriptions
across MANY customers — a real one had 100+ rows. Matching must
ALWAYS be scoped to one subscription_number first; matching against
the whole file by product name alone caused a real bug where a
different customer's "BMC Discovery - Resource Unit" row (existing
qty 9000) silently overwrote the correct Amdocs row (existing qty 2)
because both share the same product name.

Per confirmed Helix process (Sep 10 transcript): a quantity increase
is always modeled as TWO separate rows — a flat-renewal row for the
existing quantity, and a growth row (Code="a") for only the net-new
units. This function auto-resolves that case since it's a fixed rule.

Two cases remain genuinely ambiguous and are NOT auto-resolved:
  - PDF quantity == UFR existing quantity (could mean "flat renewal"
    OR "this quantity is the growth being added" — both patterns were
    seen in real data with different correct answers).
  - PDF quantity < UFR existing quantity (could mean decrease "d" or
    exclude "x" — no rule distinguishes these from data alone).
Both are returned with status="ambiguous" and must be resolved by a
human before compute_renewal_line() is called.
"""

import csv
import os
import re
from datetime import datetime

import pypdf
import openpyxl

REGION = "USA"
REF_DIR = os.path.join(os.path.dirname(__file__), "..", "references")

# Confirmed mapping from Anand's raw "Tier Global Helix" values to the
# exact column names in channel_discount_table.csv. ONLY "GOSI" has a
# confirmed 1:1 mapping (exact name match to "GOSI Strategic"). The
# other 10 real values seen (Tier 1, Tier 2a/2b/2c, Tier 3a/3b/3c/3d,
# Tier 4, Tier 5) do NOT have a confirmed mapping to the 14 real
# channel-table columns — do not guess one. Flag and ask instead.
CONFIRMED_TIER_MAPPING = {
    "GOSI": "GOSI Strategic",
}


def _normalize(name):
    return "".join(name.split()).lower()


def _to_number(value):
    """Coerces an openpyxl cell value to a number. Cells can come back
    as int, float, or str depending on how the source Excel formatted
    them — this failed once already ('>' not supported between int
    and str) when a quantity cell was text-typed."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _datedif_whole_months(start_date, end_date) -> int:
    """Excel DATEDIF(start, end, "M") — whole calendar months between
    two dates, matching the real SMT sheet's own date arithmetic.
    Verified: 30-Jun-2026 to 29-Jun-2027 = 11 months (not 12), because
    day 29 is earlier than day 30 in the final month."""
    months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    if end_date.day < start_date.day:
        months -= 1
    return months


def _parse_date_token(text):
    """Parses a date like '30-JUN-2026' (case-insensitive) into a date object."""
    return datetime.strptime(text.strip().upper(), "%d-%b-%Y").date()


def _to_date(value):
    """Normalizes a UFR cell's date value, which can come back as either
    a real datetime/date object (typical Excel date-formatted cell) or a
    string like '30-Jun-2026' (if stored as text) — handle both."""
    if value is None:
        return None
    if hasattr(value, "date"):
        return value.date()
    if hasattr(value, "year"):  # already a date object
        return value
    try:
        return _parse_date_token(str(value))
    except ValueError:
        return None


def extract_quote_pdf(pdf_path: str) -> list[dict]:
    """Extracts line items from a budgetary quote PDF at a local path."""
    reader = pypdf.PdfReader(pdf_path)
    if reader.is_encrypted:
        reader.decrypt("")

    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    header_marker = re.search(r"Per Product Fee\s*\n\s*\(in USD\)", text)
    start_idx = header_marker.end() if header_marker else 0

    # NOTE: this anchors on "BMC Continuous Support" literally. Anand
    # confirmed Support Tier can also be "L1" or "L1 Emerging Markets"
    # (different support rates: 16%, 14%). If a quote's PDF ever shows
    # a different Support Plan value, this regex will NOT match that
    # row and it will be silently dropped — this is a known, real gap,
    # not yet fixed, since no real example with a non-"BMC Continuous
    # Support" quote has been seen. Flag this if a future PDF returns
    # fewer lines than expected.
    row_re = re.compile(
        r"BMC\s*\n\s*Continuous\s*\n\s*Support.*?"
        r"per\s+([a-z\-\s]+?)\s+([\d,]+)\s+USD\s*([\d,]+\.\d+)\s+USD\s*([\d,]+\.\d{2})",
        re.DOTALL,
    )

    lines = []
    prev_end = start_idx
    term_start, term_end, new_term_months = None, None, None
    date_re = re.compile(
        r"(\d{1,2}-[A-Za-z]{3}-\s*\d{4})\s*\n?\s*to\s*(\d{1,2}-[A-Za-z]{3}-\s*\d{4})",
        re.DOTALL,
    )
    clean_date = lambda s: " ".join(s.split()).replace("- ", "-")
    for m in row_re.finditer(text, start_idx):
        product_name = " ".join(text[prev_end:m.start()].split())
        uom, qty, unit_cost, fee = m.groups()
        lines.append({
            "product": product_name,
            "uom": " ".join(uom.split()),
            "quantity": int(qty.replace(",", "")),
            "unit_cost": float(unit_cost.replace(",", "")),
            "fee": float(fee.replace(",", "")),
        })
        if term_start is None:
            preceding = text[max(0, m.start() - 200):m.start()]
            dm = date_re.search(preceding) or date_re.search(text[prev_end:m.start()])
            if dm:
                term_start, term_end = clean_date(dm.group(1)), clean_date(dm.group(2))
        prev_end = m.end()

    if not lines:
        raise ValueError(
            "No line items extracted from PDF — table format may differ "
            "from the expected BMC budgetary quote layout, OR the "
            "Support Plan isn't 'BMC Continuous Support' (see code note "
            "on the extraction regex). Flag for review."
        )

    if term_start and term_end:
        try:
            new_term_months = _datedif_whole_months(
                _parse_date_token(term_start), _parse_date_token(term_end)
            )
        except ValueError:
            new_term_months = None

    return {
        "lines": lines,
        "term_start": term_start,
        "term_end": term_end,
        "new_term_months": new_term_months,
    }


def parse_ufr_export(xlsx_path: str, subscription_number: str) -> list[dict]:
    """Parses a UFR export xlsx at a local path, filtered to ONE
    subscription_number. A real UFR export spans many subscriptions
    across many customers — never match without this filter."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    required = ["Subscription Number", "Subscription Owner Account CSN",
                "Marketing Schedule Name", "Service Type", "Support Tier",
                "RPC Quantity", "Renewal Qty Number Of Licenses",
                "BRV ACV Local", "ERV ACV Local", "TRV ACV Local",
                "Term Start Date", "Term End Date"]
    missing = [r for r in required if r not in col]
    if missing:
        raise ValueError(f"UFR export missing expected column(s): {missing}")

    records = []
    last_sub_number = None
    last_account_csn = None
    for r in range(2, ws.max_row + 1):
        # Subscription Number AND Subscription Owner Account CSN are
        # only populated on the FIRST row of each subscription's block
        # in this export format (confirmed from the real file) — carry
        # both forward for subsequent rows. Missing this for account_csn
        # caused a real bug: only the first product of a subscription
        # got a usable CSN, every other product came back with no CSN
        # at all, blocking the automatic Customer Tier lookup for them.
        sub_val = ws.cell(row=r, column=col["Subscription Number"]).value
        if sub_val:
            last_sub_number = sub_val
        csn_val = ws.cell(row=r, column=col["Subscription Owner Account CSN"]).value
        if csn_val:
            last_account_csn = csn_val
        product = ws.cell(row=r, column=col["Marketing Schedule Name"]).value
        if not product:
            continue
        if last_sub_number != subscription_number:
            continue

        term_start = _to_date(ws.cell(row=r, column=col["Term Start Date"]).value)
        term_end = _to_date(ws.cell(row=r, column=col["Term End Date"]).value)
        orig_term_months = None
        if term_start and term_end:
            orig_term_months = _datedif_whole_months(term_start, term_end)

        records.append({
            "product": product,
            "subscription_number": last_sub_number,
            "account_csn": last_account_csn,
            "service_type": ws.cell(row=r, column=col["Service Type"]).value,
            "support_tier": ws.cell(row=r, column=col["Support Tier"]).value,
            "existing_qty": _to_number(ws.cell(row=r, column=col["RPC Quantity"]).value),
            "renewal_qty": _to_number(ws.cell(row=r, column=col["Renewal Qty Number Of Licenses"]).value),
            "brv_acv": _to_number(ws.cell(row=r, column=col["BRV ACV Local"]).value),
            "erv_acv": _to_number(ws.cell(row=r, column=col["ERV ACV Local"]).value),
            "trv_acv": _to_number(ws.cell(row=r, column=col["TRV ACV Local"]).value),
            # NOTE: as of this build, these UFR date columns have been
            # observed to hold the SAME dates as the new quote's PDF
            # term, not an obviously distinct prior-period term. This is
            # unconfirmed with Helix — treat orig_term_months derived
            # from this as a proposed value the user must confirm, not
            # a verified fact.
            "ufr_term_start": term_start.strftime("%d-%b-%Y") if term_start else None,
            "ufr_term_end": term_end.strftime("%d-%b-%Y") if term_end else None,
            "orig_term_months_from_ufr": orig_term_months,
        })

    if not records:
        raise ValueError(
            f"No UFR rows found for subscription_number={subscription_number!r}. "
            "Check the exact value (case/format) against the UFR export."
        )
    return records


def match_quote_to_ufr(quote_lines: list[dict], ufr_records: list[dict]) -> list[dict]:
    """Groups quote_lines by product, matches each against its UFR
    record, and returns resolved/ambiguous match results.

    ufr_records MUST already be filtered to a single subscription —
    this function does not filter by subscription itself."""
    ufr_index = {_normalize(r["product"]): r for r in ufr_records}

    totals = {}
    for line in quote_lines:
        key = _normalize(line["product"])
        totals.setdefault(key, {"product": line["product"], "quantity": 0, "lines": []})
        line_qty = _to_number(line["quantity"])
        if line_qty is None:
            raise ValueError(
                f"PDF quantity for '{line['product']}' could not be read as "
                f"a number: {line['quantity']!r}"
            )
        totals[key]["quantity"] += line_qty
        totals[key]["lines"].append(line)

    results = []
    for key, agg in totals.items():
        ufr = ufr_index.get(key)
        if ufr is None:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "reason": "No matching UFR record for this subscription — "
                          "genuinely new product or name mismatch. Confirm "
                          "before proceeding.",
            })
            continue

        existing_qty = _to_number(ufr["existing_qty"])
        pdf_qty = _to_number(agg["quantity"])
        if existing_qty is None or pdf_qty is None:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "reason": "Existing Qty or PDF Qty could not be read as a "
                          "number for this product — check the source data.",
                "ufr": ufr,
            })
            continue

        if pdf_qty > existing_qty:
            # CORRECTED: previously split into two synthetic rows (flat
            # + growth-only with existing_qty=0), which broke the
            # per-unit baseline math (orig_brv_acv/existing_qty divides
            # by zero/undefined) and produced a wrong ~$1.99 instead of
            # the real, verified $13.67 for a live Amdocs example. The
            # real master workbook uses ONE row for an "add": existing
            # qty carried forward, proposed = existing + the PDF's
            # stated (delta) quantity, Code="a".
            proposed_qty = existing_qty + pdf_qty
            results.append({
                "product": agg["product"], "status": "resolved",
                "code": "a", "existing_qty": existing_qty, "proposed_qty": proposed_qty,
                "note": "quantity increase — one row, existing carried "
                        "forward plus the PDF's stated delta",
                "ufr": ufr,
            })
        elif pdf_qty == existing_qty:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "reason": f"PDF quantity ({pdf_qty}) equals UFR existing quantity "
                          f"({existing_qty}) — could be a flat renewal or the growth "
                          "amount being added. Confirm which applies.",
                "ufr": ufr,
            })
        else:
            results.append({
                "product": agg["product"], "status": "ambiguous",
                "reason": f"PDF quantity ({pdf_qty}) is less than UFR existing "
                          f"quantity ({existing_qty}) — could be Code 'd' (decrease) "
                          "or 'x' (exclude). Confirm which applies.",
                "ufr": ufr,
            })
    return results


def extract_and_match_quote(pdf_path: str, xlsx_path: str, subscription_number: str) -> dict:
    """Extracts the quote PDF and UFR export (filtered to ONE
    subscription), and matches them — entirely in Python, in a single
    tool call. This exists so the (potentially very large — a real UFR
    export had 100+ rows across many subscriptions) UFR data never has
    to be reproduced inside a function-call payload passed between
    separate LLM-orchestrated tool calls. Splitting extraction/parsing/
    matching into three separate tool calls caused a real
    MALFORMED_FUNCTION_CALL failure once already.

    subscription_number is REQUIRED — matching without it caused a
    real bug (a different customer's identically-named product row
    silently overwrote the correct one).

    Returns a dict:
      {"new_term_months": ..., "term_start": ..., "term_end": ...,
       "matches": [...]}
    new_term_months/term_start/term_end are derived automatically from
    the PDF's own stated dates using the confirmed DATEDIF rule — do
    NOT ask the user for these, report them and let the user confirm
    if anything looks off.

    Note on orig_term_months_proposed: derived from the UFR's own Term
    Start/End Date columns using the same DATEDIF logic. This is
    reported as a PROPOSAL, not a confirmed fact — these UFR date
    columns have been observed (Amdocs case) to hold the same dates as
    the NEW quote's PDF term, which is not obviously the prior period.
    This is unconfirmed with Helix. The agent must present it as
    "I propose X, confirm or override" and never silently treat it as
    verified the way new_term_months is.
    """
    pdf_result = extract_quote_pdf(pdf_path)
    ufr_records = parse_ufr_export(xlsx_path, subscription_number)
    matches = match_quote_to_ufr(pdf_result["lines"], ufr_records)

    # All lines in one subscription share the same UFR term dates — take
    # it from the first record. Explicitly labeled "proposed", not
    # confirmed: these UFR date columns have been observed to match the
    # NEW quote's own dates rather than an obviously distinct prior
    # term, which is unconfirmed with Helix. The agent must report this
    # as a proposal for the user to confirm or override, never as a
    # derived fact the way new_term_months is.
    orig_term_months_proposed = ufr_records[0].get("orig_term_months_from_ufr") if ufr_records else None
    ufr_term_start = ufr_records[0].get("ufr_term_start") if ufr_records else None
    ufr_term_end = ufr_records[0].get("ufr_term_end") if ufr_records else None

    return {
        "new_term_months": pdf_result["new_term_months"],
        "term_start": pdf_result["term_start"],
        "term_end": pdf_result["term_end"],
        "orig_term_months_proposed": orig_term_months_proposed,
        "orig_term_months_proposed_source": (
            f"UFR Term Start/End Date: {ufr_term_start} to {ufr_term_end}"
            if ufr_term_start else None
        ),
        "matches": matches,
    }


def get_valid_customer_tier_columns() -> list[str]:
    """Returns the exact, real column names from channel_discount_table.csv
    that represent Customer Tier options — nothing else. Call this
    before asking the user to pick a Customer Tier; do NOT reconstruct
    this list from memory or from raw Tier Global Helix values (e.g.
    "Tier 3a", "Distributor", "Partner") — those are NOT the same
    strings as the real columns and have caused the agent to offer
    fabricated options twice already."""
    path = os.path.join(REF_DIR, "channel_discount_table.csv")
    with open(path, newline="", encoding="utf-8") as f:
        header = next(csv.reader(f))
    # First two columns are "PF" (Product Family, the row key) and
    # "BMC Base List Price " — not tier options. Also add the sentinel.
    return [h for h in header[2:]] + ["Helix Base List Price"]


def get_customer_tier(account_csn) -> dict:
    """Looks up an account's Customer Tier from the confirmed CSN
    mapping file, and translates it to the exact channel_discount_table.csv
    column name where a confirmed translation exists.

    Returns:
      {"raw_tier": ..., "channel_table_column": ... or None,
       "status": "resolved" or "needs_confirmation", "reason": ...}

    Only "GOSI" currently has a confirmed 1:1 mapping to a real
    channel-table column ("GOSI Strategic"). The other 10 real tier
    values seen in the mapping file (Tier 1, Tier 2a/2b/2c, Tier 3a/3b/
    3c/3d, Tier 4, Tier 5) do NOT have a confirmed mapping — do not
    guess one. Ask the user which channel-table column applies.
    """
    path = os.path.join(REF_DIR, "account_csn_tier_mapping.csv")
    target = str(account_csn).strip()
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if str(row["Account CSN"]).strip() == target:
                raw_tier = row["Tier Global Helix"]
                mapped = CONFIRMED_TIER_MAPPING.get(raw_tier)
                if mapped:
                    return {
                        "raw_tier": raw_tier, "channel_table_column": mapped,
                        "status": "resolved",
                        "reason": f"'{raw_tier}' has a confirmed mapping to "
                                  f"channel table column '{mapped}'.",
                    }
                return {
                    "raw_tier": raw_tier, "channel_table_column": None,
                    "status": "needs_confirmation",
                    "reason": f"Account CSN {account_csn} has Tier Global Helix "
                              f"= '{raw_tier}', which has NO confirmed mapping to "
                              "a channel_discount_table.csv column yet. Ask the "
                              "user which of the 14 real columns applies — do "
                              "not guess based on the tier number alone.",
                }
    return {
        "raw_tier": None, "channel_table_column": None,
        "status": "needs_confirmation",
        "reason": f"Account CSN {account_csn} was not found in the tier "
                  "mapping file at all. Ask the user for the Customer Tier.",
    }


def get_support_rate(support_tier) -> dict:
    """Looks up the confirmed Support Rate for a given Support Tier
    label (from the UFR export's own 'Support Tier' column), per
    Anand's confirmed rule: BMC Continuous / SaaS Continuous = 20%,
    L1 = 16%, L1 Emerging Markets = 14%."""
    path = os.path.join(REF_DIR, "support_rate_table.csv")
    target = _normalize(str(support_tier))
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if _normalize(row["Support Tier"]) == target:
                return {
                    "support_rate": float(row["Support Rate"]),
                    "status": "resolved",
                    "reason": f"Support Tier '{support_tier}' -> {row['Support Rate']}",
                }
    return {
        "support_rate": None, "status": "needs_confirmation",
        "reason": f"Support Tier '{support_tier}' not found in "
                  "support_rate_table.csv — ask the user for the correct "
                  "rate rather than defaulting to 20%.",
    }
