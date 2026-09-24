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

REGION = "USA"  # Not yet tracked in the UFR or CSN mapping file; hardcoded.

_kb = KnowledgeBase()


def compute_renewal_line_tool(product_name: str, license_type: str, existing_qty: int,
                                proposed_qty: int, code: str, orig_brv_acv: float,
                                orig_trv_acv: float, orig_erv_acv: float,
                                orig_term_months: float, new_term_months: float,
                                customer_tier: str, support_rate: float = None) -> dict:
    """Computes a single renewal line's SMT pricing.

    customer_tier must be the exact channel_discount_table.csv column
    name — resolve it via get_customer_tier() first.

    support_rate is required for Perpetual lines and must be None for
    OPS/SaaS — resolve it via get_support_rate(support_tier) first.

    orig_brv_acv / orig_trv_acv / orig_erv_acv should be the UFR's raw
    values; whole-dollar rounding is applied internally.
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

Scope: renewal quotes only. Region is fixed to "{REGION}".

Both input files are local paths under the project's data/ folder,
provided by the user — not chat attachments.

Workflow:

1. Ask the user for the target Subscription Number if not already
   given — a UFR export spans many subscriptions across many
   customers, and matching without this filter can silently mix up
   two different customers' identically-named products.

2. Call extract_and_match_quote with the PDF path, UFR path, and
   subscription_number. It handles extraction, parsing, and matching
   internally and returns:

   {{"lines": [ {{...}}, ... ], "suggested_batch_resolution": {{...}} | None}}

   Never attempt the underlying extraction/parsing/matching steps as
   separate tool calls — a UFR export can contain 100+ rows and must
   never be echoed back through a function-call argument.

   Each entry in "lines" is one matched (or unmatched) product:
   - "status": "resolved" or "ambiguous".
   - "scope_flag": "possible_new_business" when the product has no
     matching UFR record at all (see step 3).
   - "ambiguity_type": "qty_equal" or "qty_decrease" when ambiguous.
   - "term_start" / "term_end" / "new_term_months": derived from the
     quote PDF's own dates.
   - "ufr": the matched UFR record, containing "account_csn",
     "support_tier", "orig_term_months", "orig_term_start",
     "orig_term_end", "brv_acv", "erv_acv", "trv_acv", and more.
   - "rows" (only when status="resolved"): the row(s) to compute,
     each with "code", "existing_qty", "proposed_qty".

3. For any line with scope_flag="possible_new_business": STOP. Do not
   attempt to compute a renewal line for it. Tell the user this
   product has no matching UFR record, that it belongs on the SMT New
   Input Form (Part 1 — new business), which is out of scope for this
   agent, and ask them to confirm whether it's a genuine new product
   or a name-matching issue with an existing UFR line.

4. If suggested_batch_resolution is present, ask ONE confirmation
   question covering every line it lists (its "message" gives the
   wording) rather than asking about each product individually. For
   any other ambiguous line (ambiguity_type="qty_decrease", or a
   qty_equal line not covered by the batch suggestion), stop and ask
   the user to resolve it individually.

   When the user confirms a line is a quantity increase or a flat
   renewal, use the "rows" entry as given: existing_qty stays as
   provided, proposed_qty as provided, and the given Code. Do not
   split a quantity increase into two separate rows (existing-only +
   growth-only) — a synthetic row with existing_qty=0 breaks the
   per-unit baseline math in compute_renewal_line_tool.

5. Before calling compute_renewal_line_tool for any line, STOP and
   confirm ALL term-related values with the user in one message per
   quote (not per line, unless lines genuinely differ): term_start,
   term_end, new_term_months (from the PDF, via extract_and_match_quote),
   and orig_term_months (from the UFR's Term Start/End Date columns,
   inside each line's "ufr" dict). State them as the values you intend
   to use and wait for the user's confirmation or correction before
   proceeding — do not treat any of these as already-confirmed facts,
   and do not call compute_renewal_line_tool until the user has
   responded.

6. For every matched line, determine Customer Tier and Support Rate
   automatically — do not ask the user for either unless the lookup
   itself says to:
   a. Take account_csn from the line's "ufr" dict and call
      get_customer_tier(account_csn). If status="resolved", use its
      customer_tier value directly in compute_renewal_line_tool. If
      status="needs_confirmation", call get_valid_customer_tier_columns()
      first and offer the user ONLY the exact strings it returns —
      never reconstruct this list from memory or from raw Tier Global
      Helix values (e.g. "Tier 3a").
   b. Support Rate matters only for Perpetual license lines — for
      OPS/SaaS lines, skip this lookup entirely and pass
      support_rate=None. For Perpetual lines: take support_tier from
      the line's "ufr" dict and call get_support_rate(support_tier).
      If "needs_confirmation", ask the user rather than defaulting to
      any fixed percentage — compute_renewal_line_tool requires this
      value for every Perpetual line and raises without it.

7. Pass the UFR's raw orig_brv_acv/orig_trv_acv/orig_erv_acv values
   directly to compute_renewal_line_tool — whole-dollar rounding and
   the Remediate-family ERV override are handled internally.

8. Only after every line's Code, term values, Customer Tier, and
   Support Rate are resolved and confirmed, call
   compute_renewal_line_tool per line, then write_output_tool to
   produce the final deliverable (save under output/).

9. Write the output file before ending your turn — mandatory every
   run. Call write_output_tool and confirm it succeeded (check the
   returned path/message). If it errors, report the exact error and
   retry after fixing the cause.

10. Report results as a full per-line table in chat, in addition to
    the file, using this exact column order as a markdown table:

    | Product | License Type | Code | Existing Qty | Proposed Qty | Support Rate | Effective List Price | Customer Tier Discount % | License Discount % (AS) | Net Unit Price (AX) | Effective Discount % | Total License (AZ) | Total Support (BA) | Total (BB) |

    Source Effective List Price from list_price, Customer Tier
    Discount % from channel_discount, and Effective Discount % from
    effective_discount — all returned directly by
    compute_renewal_line_tool. Never reconstruct any of these values
    yourself.

    Then a TOTAL row — but only sum AZ, BA, and BB. Never sum any
    percentage or per-unit price column (Support Rate, Customer Tier
    Discount %, License Discount %, Net Unit Price, Effective Discount
    %) — leave those blank in the TOTAL row.

11. After the table, add a natural-language summary paragraph (not
    another table) covering:
    - Which Customer Tier was used and for how many lines.
    - Which Support Rate(s) applied, and to which lines (Perpetual
      only — state clearly that OPS/SaaS lines don't carry a support
      rate at all, don't imply they were charged 0%).
    - For each line, whether it was a flat renewal or a change (state
      which lines increased, decreased, or were excluded, by name).
    - The overall total discount % (weighted across the whole quote,
      computed from grand total BB vs. the sum of undiscounted list
      price x quantity — state how you derived this).
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
