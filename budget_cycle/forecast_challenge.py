"""Driver/cash challenge on the existing budget; no workbook or inputs mutated."""

from __future__ import annotations

import hashlib
import json
import math
from copy import deepcopy
from pathlib import Path

import pandas as pd

from budget_cycle import budget_model as model

HERE = Path(__file__).resolve().parent
SCENARIOS = {
    "approved_budget": {},
    "combined_downside": {"cost_inflation": .02, "dso_days": 7, "dio_days": 5},
    "pricing_response": {"cost_inflation": .02, "dso_days": 7, "dio_days": 5,
                         "price_change": .01, "elasticity": -1.5},
    "working_capital_response": {"cost_inflation": .02, "dso_days": 4, "dio_days": 3},
}


def evaluate(name: str, drivers: dict) -> tuple[dict, pd.DataFrame]:
    allowed = {"cost_inflation", "dso_days", "dio_days", "price_change", "elasticity"}
    if set(drivers) - allowed or any(not math.isfinite(float(v)) for v in drivers.values()):
        raise ValueError("Unknown or non-finite driver")
    cost = float(drivers.get("cost_inflation", 0))
    price = float(drivers.get("price_change", 0))
    elasticity = float(drivers.get("elasticity", 0))
    if not -.5 <= cost <= 1 or not -.2 <= price <= .2 or not -5 <= elasticity <= 0:
        raise ValueError("Driver outside the bounded planning sensitivity")
    inputs = model.inputs()
    build = model.build(inputs).copy(deep=True)
    bs = deepcopy(model.A["balance_sheet"])
    for measure, driver in (("dso", "dso_days"), ("dio", "dio_days")):
        bs[measure] += float(drivers.get(driver, 0))
        if not 0 <= bs[measure] <= 120:
            raise ValueError("Working-capital days outside 0–120")
    # Action starts in April. Q1 is the approved budget here, NOT the Q1 actuals.
    remaining = build.month >= 4
    volume = (1 + price) ** elasticity
    build.loc[remaining, "cases"] *= volume
    build.loc[remaining, "revenue_v2"] *= (1 + price) * volume
    build.loc[remaining, "cogs"] *= (1 + cost) * volume
    build.loc[remaining, "freight"] *= volume
    statement = model.statements(inputs, build, balance_sheet=bs)
    if statement.balance_check.abs().max() > 1e-6:
        raise AssertionError("Accounting identity failed")
    if statement.cash.min() < bs["minimum_cash"] - 1e-6 or statement.revolver.min() < -1e-6:
        raise AssertionError("Cash/revolver mechanism failed")
    result = {"scenario": name, "drivers": drivers,
              "revenue": float(statement.revenue.sum()), "ebitda": float(statement.ebitda.sum()),
              "net_income": float(statement.net_income.sum()), "interest": float(statement.interest.sum()),
              "closing_revolver": float(statement.revolver.iloc[-1]),
              "peak_revolver": float(statement.revolver.max()), "closing_cash": float(statement.cash.iloc[-1]),
              "closing_receivables": float(statement.ar.iloc[-1]),
              "closing_inventory": float(statement.inventory.iloc[-1]),
              "maximum_balance_residual": float(statement.balance_check.abs().max()),
              "remaining_volume_multiplier": volume}
    return result, statement


def packet() -> tuple[dict, pd.DataFrame]:
    cases, frames = [], []
    for name, drivers in SCENARIOS.items():
        case, frame = evaluate(name, drivers)
        cases.append(case)
        frames.append(frame.reset_index().assign(scenario=name))
    base = cases[0]
    for case in cases:
        case["ebitda_vs_budget"] = case["ebitda"] - base["ebitda"]
        case["closing_revolver_vs_budget"] = case["closing_revolver"] - base["closing_revolver"]
    inputs = [HERE / "assumptions.json", *sorted((HERE / "data").glob("*.csv"))]
    return {
        "schema_version": 1, "classification": "synthetic_budget_sensitivity",
        "basis": "Approved FY2027 budget; shocks affect April–December, working-capital day changes affect all modeled months. This is separate from the existing actuals-based 3+9 reforecast.",
        "input_sha256": {path.relative_to(HERE).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs},
        "scenarios": cases,
        "decision": "CHALLENGE REQUIRED; no budget or management approval recorded",
        "owner_questions": {
            "Sales": "Does the 1% price response retain the modeled demand under the assumed -1.5 elasticity?",
            "Operations": "Can inventory days fall without increasing stockouts or expiry costs?",
            "Finance": "Are collection timing, tax-loss treatment and financing capacity acceptable?",
        },
        "limits": ["Elasticity is an assumption, not estimated causal demand response",
                   "No facility limit or lender covenant is supplied; peak debt is exposure, not confirmed funding",
                   "Existing tax-loss credit and opening-balance interest conventions are unchanged",
                   "No real business-partner challenge session has occurred"],
    }, pd.concat(frames, ignore_index=True)


def main():
    result, monthly = packet()
    output = HERE / "challenge_output"
    output.mkdir(exist_ok=True)
    (output / "forecast_challenge.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    monthly.to_csv(output / "monthly_statements.csv", index=False, float_format="%.6f", lineterminator="\n")
    lines = ["# Budget challenge: margin is only part of the decision", "",
             result["basis"], "", "All data and scenarios are synthetic. No management sign-off is claimed.", "",
             "| Scenario | EBITDA | Interest | Closing revolver | Peak revolver |", "|---|---:|---:|---:|---:|"]
    for case in result["scenarios"]:
        lines.append(f"| {case['scenario']} | ${case['ebitda']:,.0f} | ${case['interest']:,.0f} | ${case['closing_revolver']:,.0f} | ${case['peak_revolver']:,.0f} |")
    lines += ["", "Challenge the demand assumption before approving the price response. Compare the working-capital response on debt and cash as well as EBITDA. Collection and inventory actions require customer-term and service validation; they are not free improvements.", "",
              "The monthly export preserves all three statements and checks the balance identity under every case. A financing limit is unknown, so no scenario is called funded or covenant-compliant.", "",
              "## Questions for the decision meeting", ""]
    lines += [f"- **{owner}:** {question}" for owner, question in result["owner_questions"].items()]
    lines += ["", "## Limits", ""] + [f"- {limit}" for limit in result["limits"]]
    (output / "DECISION_MEMO.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(json.dumps({"scenarios": len(result["scenarios"]), "maximum_balance_residual": max(x["maximum_balance_residual"] for x in result["scenarios"])}))


if __name__ == "__main__":
    main()
