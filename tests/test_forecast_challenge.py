from copy import deepcopy

import pandas as pd
import pytest

from budget_cycle import budget_model as model
from budget_cycle import forecast_challenge as challenge


def test_zero_shock_is_exactly_the_existing_statement_model():
    before = deepcopy(model.A)
    inputs = model.inputs()
    _, result = challenge.evaluate("unchanged", {})
    pd.testing.assert_frame_equal(result, model.statements(inputs, model.build(inputs)))
    assert model.A == before


def test_all_cases_preserve_accounting_and_q1_budget_earnings():
    packet, monthly = challenge.packet()
    assert monthly.groupby("scenario").size().eq(12).all()
    assert monthly.balance_check.abs().max() < 1e-6
    q1 = monthly[monthly.month <= 3].groupby("scenario").ebitda.sum()
    assert q1.max() == pytest.approx(q1.min())
    assert len(packet["scenarios"]) == 4


def test_working_capital_releases_debt_without_invented_operating_profit():
    cases = {case["scenario"]: case for case in challenge.packet()[0]["scenarios"]}
    down, action = cases["combined_downside"], cases["working_capital_response"]
    assert action["ebitda"] == pytest.approx(down["ebitda"])
    assert action["closing_revolver"] < down["closing_revolver"]
    assert action["interest"] < down["interest"]
    assert cases["pricing_response"]["remaining_volume_multiplier"] < 1


@pytest.mark.parametrize("drivers", [{"dso_days": -100}, {"price_change": float("nan")},
                                      {"elasticity": 2}, {"unknown": 3}])
def test_unsupported_assumptions_are_rejected(drivers):
    with pytest.raises(ValueError):
        challenge.evaluate("invalid", drivers)
