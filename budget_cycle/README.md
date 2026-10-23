# Kestrel Bay FY2027: the budget cycle, end to end

The repository's [Kestrel Bay workbook](../examples/kestrel-bay/) ends FY2026 with revenue ahead of
budget and EBITDA behind it. This picks the same synthetic distributor up for the next year's
budget: from leadership's target, through the departments' submissions and the challenge round,
to an approved budget with three linked statements and, after the first quarter, a 3+9 reforecast.

**[Download the workbook](Kestrel_Bay_FY2027_Budget.xlsx)** (Excel or LibreOffice). Every figure is
a formula over the inputs and the data sheets; nothing is pasted in.

## What the cycle found

| FY2027 | Net revenue | EBITDA | Margin |
|---|---:|---:|---:|
| Last twelve months (the base) | $45.88M | $2.22M | 4.84% |
| V0: leadership's target | $48.63M | $2.92M | 6.00% |
| V1: the departments' submissions | $49.56M | **$1.99M** | 4.02% |
| V2: after the challenge round (approved) | $49.74M | **$2.33M** | 4.68% |
| 3+9 reforecast after Q1 | $49.76M | **$1.66M** | 3.33% |

**1. Bottom-up, the year gets worse, not better.** The submissions grow revenue 8.0%, above the 6%
target, yet EBITDA falls below last year's. Procurement's cost inflation (beef +6%) outruns sales'
price increases (beef +4%), so COGS grows 9.3% and gross margin slips from 23.8% to 22.9%, and
department opex grows 7.3%. V1 misses the EBITDA target by $926K.

**2. The challenge round closes $335K of it.** An extra 1.5% on beef and seafood in the foodservice
channels (+$179K), a warehouse non-pay cut (+$111K), the sales hire moved from April to July
(+$24K) and technology growth capped at 5% (+$21K). The bridge closes to the cent.

**3. The rest is a decision, not a model.** Approved EBITDA is still $591K short. Closing it takes
a further **1.28%** on every price from February, which the workbook computes rather than assumes.
The alternative is resetting the target; the workbook leaves that to the people who own it.

**4. The balance sheet funds the year.** Opening working capital is the August month-end, at
summer volumes, so January releases cash and the revolver drops from $1.2M to $0.5M. It peaks
again at $1.0M in June as stock and receivables build for the summer, and is repaid by December,
closing with $405K of cash. Net income is $1.29M. The balance sheet balances and the cash flow
ties to it in every month.

**5. Q1 says the approved budget is already $671K too high.** Grocery Retail ran 6% above budget
and Restaurants 4% below, beef cost 2.75% more per case and the warehouse spent 5.7% over. Revenue
looks fine, $22K ahead for the year, because the cases moved to the cheaper channel. The EBITDA
bridge from budget to reforecast:

| Q1 actual vs budget | Volume and mix, Apr-Dec | Beef and other unit costs, Apr-Dec | Opex run-rate, Apr-Dec | Total |
|---:|---:|---:|---:|---:|
| −$155K | −$75K | −$281K | −$159K | **−$671K** |

## What's in the workbook

| Sheet | What it does |
|---|---|
| **Summary** | The cycle in six steps and the headline numbers, all formulas |
| **Assumptions** | Every input: targets, category and channel submissions, department opex and hires, challenge levers, balance sheet and financing, capex |
| **Consolidation** | Last twelve months, V0, V1 and V2 side by side; gaps to target; the V1-to-V2 bridge; the price that would close the rest |
| **Statements** | Monthly P&L, balance sheet and cash flow for the approved budget, with a revolver that holds minimum cash |
| **Reforecast** | Q1 run-rates by channel, category and department; the 3+9 against budget; the EBITDA bridge |
| **Build** | 24 category × channel segments × 12 months: cases, price, cost and freight for V1, V2 and the reforecast |
| **Opex** | Department opex by month for V1 and V2, hires from their start month |
| **Data_Base, Data_Q1** | The extract the cycle starts from and the synthetic Q1 actuals, loaded unchanged |
| **Checks** | 15 controls, each with its tolerance: seasonality, cases, the statements against the consolidation, the balance sheet and cash every month, the revolver, both bridges, Q1 against actuals |

How the months are built: each segment's last-twelve-month cases grow by its channel's submitted
rate and are spread by its category's 2025 seasonality. Price increases start in the month set on
the Assumptions sheet; cost and freight inflation apply all year. Interest accrues on the opening
revolver, so the three statements have no circular reference.

## How it's verified

- **A twin in pandas.** [`budget_model.py`](budget_model.py) computes the whole cycle from the same
  inputs without the workbook. [`tests/test_budget_cycle.py`](../tests/test_budget_cycle.py) reads the
  saved workbook and holds it to the twin to the cent: the consolidation, both bridges, every month
  of 22 statement lines, every reforecast ratio, total and bridge step. 73 tests, including one that holds this page's numbers to the model.
- **The workbook checks itself.** 15 in-workbook controls, all passing; a test fails if one does,
  or if any cell holds an error.
- **No typed numbers.** A test walks the six calculation sheets and fails on any number that is not
  a formula. Inputs live on Assumptions; data on the Data sheets.
- **The starting point is traceable.** [`extract_base.py`](extract_base.py) takes the last twelve
  months, the seasonality, the opex and the opening working capital from the Kestrel Bay workbook;
  a test reruns it and requires the committed CSVs back byte for byte. The synthetic Q1 actuals
  regenerate identically from a fixed seed.

## How it's built

[`build_workbook.py`](build_workbook.py) writes the formulas with openpyxl, then LibreOffice opens and
saves the file headless so it carries calculated values. Unlike the flagship utility workbook, which
was built in Microsoft 365 Excel with Power Query and Power Pivot, this one keeps to functions Excel
and LibreOffice share (`SUMIFS`, `SUMPRODUCT`, `INDEX`/`MATCH`, `EOMONTH`), so anyone can rebuild it:

```bash
python budget_cycle/extract_base.py      # the base, from the Kestrel Bay workbook
python budget_cycle/budget_model.py --q1  # the synthetic Q1 actuals
python budget_cycle/build_workbook.py    # the workbook, recalculated by LibreOffice
python -m pytest tests/test_budget_cycle.py
```

Kestrel Bay is synthetic: no real company, customer or price is represented.
