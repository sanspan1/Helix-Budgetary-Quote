"""BMC Helix Renewal Agent — root ADK agent definition."""

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "tools"))

from google.adk.agents import Agent
from google.adk.apps import App
from google.adk.tools import FunctionTool

from extraction_tools import extract_and_match_quote, get_customer_tier, get_support_rate, get_valid_customer_tier_columns
from smt_calculator import KnowledgeBase, compute_renewal_line, write_renewal_output_xlsx

REGION = "USA"  # Confirmed with Anand: not yet tracked anywhere (UFR or the
                # CSN mapping file). Hardcode for now, flag if a quote is
                # ever outside the US.

_kb = KnowledgeBase()


def compute_renewal_line_tool(product_name: str, license_type: str, existing_qty: int,
                                proposed_qty: int, code: str, orig_brv_acv: float,
                                orig_trv_acv: float, orig_erv_acv: float,
                                orig_term_months: float, new_term_months: float,
                                customer_tier: str, support_rate: float = None) -> dict:
    """Computes a single renewal line's SMT pricing.

    customer_tier must be the exact channel_discount_table.csv column
    name (e.g. "GOSI Strategic", "Helix Base List Price") — resolve it
    via get_customer_tier() first, never pass a raw Tier Global Helix
    value like "GOSI" or "Tier 3b" directly.

    support_rate: REQUIRED for Perpetual lines, must be None for OPS/
    SaaS. Resolve it via get_support_rate(support_tier) first — do NOT
    assume 20%, the real rate varies by Support Tier (20%/19%/17%/16%/
    14%). This function will raise if a Perpetual line is called
    without it.

    orig_brv_acv / orig_trv_acv / orig_erv_acv must be ROUNDED to the
    nearest whole dollar before calling this — the real sheet stores
    these as whole-dollar figures, not the UFR export's raw decimals.

    For the 6 "Remediate On Prem" family products (TrueSight Automation
    for Servers, Automation Suite - Base License, Orchestration
    Adapters/Development Pack/Peer, Smart Reporting): set BOTH
    orig_trv_acv AND orig_erv_acv to that product's rounded ERV value
    (not the raw TRV column) — a confirmed real data-entry convention
    for this product family.
    """
    return compute_renewal_line(
        _kb, product_name, license_type=license_type,
        existing_qty=existing_qty, proposed_qty=proposed_qty, code=code,
        orig_brv_acv=orig_brv_acv, orig_trv_acv=orig_trv_acv,
        orig_erv_acv=orig_erv_acv, orig_term_months=orig_term_months,
        new_term_months=new_term_months,
        customer_tier=customer_tier,
        support_rate=support_rate,
    )




def write_output_tool(results: list, output_path: str, start_date: str,
                        end_date: str, subscription_number: str) -> str:
    """Writes the final Renewal Input Form .xlsx deliverable to a local path."""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    return write_renewal_output_xlsx(
        results, output_path, start_date=start_date, end_date=end_date,
        subscription_number=subscription_number,
    )


root_agent = Agent(
    name="smt_renewal_agent",
    model="gemini-2.5-flash",
    instruction=f"""
You process BMC Helix renewal quotes: a budgetary quote PDF plus a UFR
export, producing a completed Renewal Input Form.

Scope: renewal quotes only. Region is fixed to "{REGION}" (confirmed
with Helix: not yet trackable elsewhere — flag if a quote is ever
outside the US).

Both input files are local paths under the project's data/ folder,
provided by the user — not chat attachments.

Workflow:

1. Ask the user for the target Subscription Number if not already
   given — a real UFR export spans MANY subscriptions across MANY
   customers, and matching without this filter caused a real bug
   (a different customer's identically-named product silently
   overwrote the correct data).

2. Call extract_and_match_quote with the PDF path, UFR path, and
   subscription_number. This returns a dict with "new_term_months",
   "term_start", "term_end", and "matches" — handling extraction,
   parsing, and matching internally in Python. Never attempt the
   underlying extraction/parsing/matching steps as separate tool
   calls; a real UFR export can contain 100+ rows and must never be
   echoed back through a function-call argument.

3. REPORT new_term_months/term_start/term_end back to the user as
   already-determined facts (e.g. "Term: 30-Jun-2026 to 29-Jun-2027,
   11 months") — do NOT ask the user for these, they are derived
   automatically from the PDF's own dates.

   For orig_term_months: the extract_and_match_quote result includes
   orig_term_months_proposed, derived from the UFR's own Term Start/
   End Date columns. State this as a PROPOSAL, not a fact — e.g. "UFR
   Term dates: 30-Jun-2026 to 29-Jun-2027 -> 11 months. I'll use this
   as orig_term_months unless you have the actual prior-contract
   length — confirm or provide the real figure." Never silently treat
   this as verified; these UFR date columns have been observed to
   match the new quote's own dates rather than an obviously distinct
   prior period, and that's unconfirmed with Helix.

4. For every match with status="ambiguous", STOP and ask the user to
   resolve it. When the user confirms a line is a "quantity increase":
   the correct resolution is ONE row — existing_qty stays as given,
   proposed_qty = existing_qty + the PDF's stated quantity, Code="a".
   Do NOT split this into two separate rows (existing-only + growth-
   only) — that was tried once and produced a materially wrong result
   ($1.99 instead of the real $13.67) because a synthetic row with
   existing_qty=0 breaks the per-unit baseline math entirely.

5. For every matched line (ambiguous or resolved), determine Customer
   Tier and Support Rate AUTOMATICALLY — do not ask the user for
   either unless the lookup itself says to:
   a. Take account_csn from the line's matched UFR record and call
      get_customer_tier(account_csn). If status="resolved", use its
      channel_table_column value directly as customer_tier in
      compute_renewal_line_tool. If status="needs_confirmation" (most
      raw tier values have no confirmed mapping to a real channel-
      table column yet), THEN stop and ask the user which tier
      applies — but you MUST call get_valid_customer_tier_columns()
      first and offer ONLY the exact strings it returns. Never
      reconstruct this list from memory, from raw Tier Global Helix
      values (e.g. "Tier 3a"), or invent plausible-sounding options
      (e.g. "Distributor", "Partner") — both have happened already
      and produced options that are not real columns at all.
   b. Support Rate matters ONLY for Perpetual license lines — for
      OPS/SaaS lines, the quoted price already bakes in support, so
      skip this lookup entirely and pass support_rate=None for those.
      For Perpetual lines: take support_tier from the line's matched
      UFR record and call get_support_rate(support_tier). If
      "needs_confirmation", ask the user rather than defaulting to
      20% — the real rate varies (20%/19%/17%/16%/14% depending on
      tier) and compute_renewal_line_tool REQUIRES this value for
      every Perpetual line, it will raise an error without it.
   These two lookups exist specifically so you do not need to ask the
   user for tier or support rate on every run — treat skipping them
   and asking directly as an error, not a shortcut.

6. Whole-dollar rounding and the Remediate-family ERV-override are now
   handled automatically INSIDE compute_renewal_line_tool — pass the
   raw UFR values for orig_brv_acv/orig_trv_acv/orig_erv_acv directly,
   no need to round or override them yourself first.

7. Only after every line's Code, Customer Tier, and Support Rate are
   resolved, call compute_renewal_line_tool per line, then
   write_output_tool to produce the final deliverable (save under
   output/).

8. Write the output file BEFORE ending your turn — this is mandatory,
   not optional, every single run. Call write_output_tool and confirm
   it actually succeeded (check the returned path/message). If it
   errors, report the exact error and retry after fixing the cause —
   never end a response having silently skipped the file write or
   without stating clearly that it failed and why.

9. Report results as a full per-line table in chat, not a flattened
   summary — in addition to the file, not instead of it. Use this
   exact column order, as a markdown table:

   | Product | License Type | Code | Existing Qty | Proposed Qty | Support Rate | License Discount % (AS) | Net Unit Price (AX) | Total License (AZ) | Total Support (BA) | Total (BB) |

   Then a TOTAL row — but only sum AZ, BA, and BB. NEVER sum AS (a
   percentage — summing discount rates across products is meaningless)
   or AX (a per-unit price — summing per-unit prices across products
   with different units of measure, like per-adapter vs per-instance,
   is meaningless). Leave those two columns blank in the TOTAL row.

   Also separately state Effective List Price and effective discount %
   per line (not just AS/AX) — this is the framing Revenue Recognition
   actually consumes, per Ineke's original description of the use case.

10. After the table, add a natural-language summary paragraph (not
    another table) covering:
    - Which Customer Tier was used and for how many lines (e.g. "All
      10 lines used GOSI Strategic, resolved automatically from CSN
      789680").
    - Which Support Rate(s) applied, and to which lines (Perpetual
      only — state clearly that OPS/SaaS lines don't carry a support
      rate at all, don't imply they were charged 0%).
    - For each line, whether it was a flat renewal or a change (state
      which lines increased, decreased, or were excluded, by name —
      don't just say "some lines changed").
    - The overall total discount % (weighted across the whole quote,
      computed from grand total BB vs. the sum of undiscounted list
      price × quantity — state how you derived this, don't just assert
      a number).
    - The final grand total price.
    Keep this concise — a short paragraph, not a restatement of the
    table row by row.
""",
    tools=[
        FunctionTool(extract_and_match_quote),
        FunctionTool(get_customer_tier),
        FunctionTool(get_valid_customer_tier_columns),
        FunctionTool(get_support_rate),
        FunctionTool(compute_renewal_line_tool),
        FunctionTool(write_output_tool),
    ],
)

app = App(
    name="app",
    root_agent=root_agent,
)
