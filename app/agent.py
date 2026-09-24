"""BMC Helix Renewal Agent — root ADK agent definition."""

import os
import sys
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent / "tools"))

from google.adk.agents import Agent
from google.adk.apps import App
from google.adk.tools import FunctionTool, ToolContext

from extraction_tools import (
    RESOLUTION_DECREASE, RESOLUTION_EXCLUDE, RESOLUTION_FLAT, RESOLUTION_GROWTH,
    _normalize, build_renewal_line_inputs, extract_and_match_quote,
    get_customer_tier, get_valid_customer_tier_columns,
)
from smt_calculator import (
    KnowledgeBase, compute_renewal_quote, render_renewal_table_markdown,
    write_renewal_output_xlsx,
)

REGION = "USA"  # Not yet tracked in the UFR or CSN mapping file; hardcoded.
MATCH_STATE_KEY = "smt_match_result"
SUBSCRIPTION_STATE_KEY = "smt_subscription_number"
DEFAULT_OUTPUT_DIR = "output"

_kb = KnowledgeBase()


def extract_and_match_quote_tool(pdf_path: str, xlsx_path: str, subscription_number: str,
                                   tool_context: ToolContext) -> dict:
    """Extracts the quote PDF, parses the UFR export for one subscription,
    and matches them. The full result (all UFR ACV values) is stored in
    session state for build_renewal_form_tool — you never need to pass
    any ACV, quantity or price back in yourself.

    Returns a compact per-line view for the confirmation questions.
    """
    result = extract_and_match_quote(pdf_path, xlsx_path, subscription_number)
    tool_context.state[MATCH_STATE_KEY] = result
    tool_context.state[SUBSCRIPTION_STATE_KEY] = subscription_number

    view = []
    for line in result["lines"]:
        ufr = line.get("ufr") or {}
        view.append({
            "product": line["product"],
            "status": line["status"],
            "scope_flag": line.get("scope_flag"),
            "ambiguity_type": line.get("ambiguity_type"),
            "reason": line.get("reason"),
            "existing_qty": ufr.get("existing_qty"),
            "pdf_qty": line.get("pdf_qty"),
            "license_type": ufr.get("service_type"),
            "support_tier": ufr.get("support_tier"),
            "account_csn": ufr.get("account_csn"),
            "term_start": line.get("term_start"),
            "term_end": line.get("term_end"),
            "new_term_months": line.get("new_term_months"),
            "orig_term_start": ufr.get("orig_term_start"),
            "orig_term_end": ufr.get("orig_term_end"),
            "orig_term_months": ufr.get("orig_term_months"),
            "term_warning": line.get("term_warning"),
        })
    return {"lines": view, "suggested_batch_resolution": result["suggested_batch_resolution"]}


def build_renewal_form_tool(tool_context: ToolContext,
                              customer_tier: str = "",
                              confirm_all_flat: bool = False,
                              flat_products: Optional[list[str]] = None,
                              growth_products: Optional[list[str]] = None,
                              decrease_products: Optional[list[str]] = None,
                              exclude_products: Optional[list[str]] = None,
                              new_term_months_override: float = 0.0,
                              orig_term_months_override: float = 0.0,
                              perpetual_support_rate_override: float = -1.0,
                              output_path: str = "") -> dict:
    """Computes every renewal line, all totals and the summary in code,
    and writes the Renewal Input Form .xlsx. Call once, after the user
    has confirmed line resolutions and terms.

    customer_tier: leave empty to auto-resolve from the account CSN.
        Pass an exact get_valid_customer_tier_columns() value only when
        auto-resolution needed confirmation.
    confirm_all_flat: True when the user confirmed the batch flat-renewal
        suggestion (applies to every qty_equal line not listed elsewhere).
    flat_products / growth_products / decrease_products / exclude_products:
        product names (as returned by extract_and_match_quote_tool) for
        ambiguous lines the user resolved individually.
    new_term_months_override / orig_term_months_override: only when the
        user corrected a term; 0 means use the extracted value.
    perpetual_support_rate_override: only when a Perpetual line's
        support tier could not be resolved and the user gave a rate; -1
        means none.
    output_path: defaults to output/Renewal_Input_Form_<subscription>.xlsx

    Returns table_markdown and summary_text to show the user VERBATIM,
    plus output_path, totals, and any pending/blocked lines.
    """
    match_result = tool_context.state.get(MATCH_STATE_KEY)
    if not match_result:
        return {"status": "error",
                "message": "No extracted quote in session. Call extract_and_match_quote_tool first."}
    subscription_number = tool_context.state.get(SUBSCRIPTION_STATE_KEY, "")

    resolutions = {}
    if confirm_all_flat:
        for line in match_result["lines"]:
            if line.get("ambiguity_type") == "qty_equal":
                resolutions[_normalize(line["product"])] = RESOLUTION_FLAT
    for names, choice in ((flat_products, RESOLUTION_FLAT),
                          (growth_products, RESOLUTION_GROWTH),
                          (decrease_products, RESOLUTION_DECREASE),
                          (exclude_products, RESOLUTION_EXCLUDE)):
        for name in names or []:
            resolutions[_normalize(name)] = choice

    built = build_renewal_line_inputs(
        match_result, resolutions,
        new_term_months_override=new_term_months_override or None,
        orig_term_months_override=orig_term_months_override or None,
        perpetual_support_rate_override=(
            perpetual_support_rate_override if perpetual_support_rate_override >= 0 else None
        ),
    )
    if built["pending"] or built["blocked"]:
        return {"status": "needs_input", "pending": built["pending"], "blocked": built["blocked"],
                "message": "Resolve these lines with the user, then call again."}

    tier = customer_tier.strip()
    if not tier:
        csns = {l["account_csn"] for l in built["line_inputs"]}
        if len(csns) != 1:
            return {"status": "needs_input",
                    "message": f"Lines span multiple account CSNs {sorted(map(str, csns))}; "
                               "ask the user which Customer Tier to use.",
                    "valid_tiers": get_valid_customer_tier_columns()}
        lookup = get_customer_tier(csns.pop())
        if lookup["status"] != "resolved":
            return {"status": "needs_input", "message": lookup["reason"],
                    "valid_tiers": get_valid_customer_tier_columns()}
        tier = lookup["customer_tier"]

    quote = compute_renewal_quote(_kb, built["line_inputs"], tier)

    if not output_path:
        output_path = os.path.join(DEFAULT_OUTPUT_DIR,
                                   f"Renewal_Input_Form_{subscription_number}.xlsx")
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    write_renewal_output_xlsx(quote, output_path, start_date=built["start_date"],
                              end_date=built["end_date"],
                              subscription_number=subscription_number)

    return {
        "status": "ok",
        "output_path": output_path,
        "customer_tier": tier,
        "table_markdown": render_renewal_table_markdown(quote),
        "summary_text": quote["summary"]["text"],
        "totals": quote["totals"],
        "warnings": quote["summary"]["warnings"],
    }


root_agent = Agent(
    name="smt_renewal_agent",
    model="gemini-2.5-flash",
    instruction=f"""
You process BMC Helix renewal quotes: a budgetary quote PDF plus a UFR
export, producing a completed Renewal Input Form.

Scope: renewal quotes only. Region is fixed to "{REGION}". Both input
files are local paths under the project's data/ folder.

You never do arithmetic. Every price, discount, total, list value and
overall discount comes from build_renewal_form_tool.

Workflow:

1. Ask for the Subscription Number if not given.

2. Call extract_and_match_quote_tool(pdf_path, xlsx_path,
   subscription_number). It stores all UFR values in session state.

3. Lines with scope_flag="possible_new_business": STOP for that line.
   Tell the user it has no matching UFR record, belongs on the SMT New
   Input Form (out of scope), and ask whether it is genuinely new or a
   name-matching issue.

4. If suggested_batch_resolution is present, ask ONE question covering
   every listed product (use its message). Ask about any other ambiguous
   line (qty_decrease: decrease or exclude?) individually.

5. In the same message, state the terms you will use (term_start,
   term_end, new_term_months, orig_term_months, from the tool output)
   and ask the user to confirm or correct them.

6. After the user answers, call build_renewal_form_tool once:
   - confirm_all_flat=True if they accepted the batch flat suggestion.
   - Put individually resolved products in flat_products /
     growth_products / decrease_products / exclude_products.
   - Term overrides only if the user corrected a term.
   - Leave customer_tier empty; it is resolved from the account CSN.
   If it returns status="needs_input", ask the user exactly what it
   lists (offer only the valid_tiers strings for a tier question), then
   call it again. Never guess a tier or support rate.

7. When status="ok", reply with exactly:
   - table_markdown, copied verbatim (do not edit, reorder, round or
     recompute any cell, and do not add rows).
   - summary_text, copied verbatim as one paragraph.
   - One line: the output file path.
   - If warnings is non-empty, list them.
""",
    tools=[
        FunctionTool(extract_and_match_quote_tool),
        FunctionTool(get_valid_customer_tier_columns),
        FunctionTool(build_renewal_form_tool),
    ],
)

app = App(
    name="app",
    root_agent=root_agent,
)
