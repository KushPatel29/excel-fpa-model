# lululemon athletica: a DCF from its own filings

A discounted cash flow valuation of lululemon (Vancouver; Nasdaq: LULU) as of 17 March 2026, the
day it filed its FY2025 10-K. Every historical figure comes from SEC EDGAR's XBRL data with the
accession number of the 10-K it was taken from; every market input names its source; every
forecast driver is a stated judgement in three cases; every output is an Excel formula.

**[Download the workbook](Lululemon_DCF_FY2025.xlsx)** (Excel or LibreOffice). Blue on yellow is
an input; everything else is calculated.

> Not investment advice. This is a worked valuation: what the business is worth under explicit
> assumptions, and how much each assumption moves it. It does not compare the result with the share
> price.

## What the filings say

| | FY2021 | FY2022 | FY2023 | FY2024 | FY2025 |
|---|---:|---:|---:|---:|---:|
| Revenue (USD M) | 6,257 | 8,111 | 9,619 | 10,588 | 11,103 |
| Growth | | 29.6% | 18.6% | 10.1% | 4.9% |
| Operating margin | 21.3% | 16.4% | 22.2% | 23.7% | 19.9% |
| Operating working capital, % of revenue | 7.7% | 13.2% | 9.6% | 7.6% | 9.6% |

Growth has slowed from 30% to 5% in three years, and FY2025's operating margin fell almost four
points from FY2024's peak while inventory rose 18% on 5% more sales. FY2022's margin is the Mirror
write-down ($408M of impairment and restructuring), not the business. The valuation question is
whether FY2025 is a reset or the start of a decline.

## What it is worth

| | Bear | Base | Bull |
|---|---:|---:|---:|
| **Value per share** | **$171** | **$278** | **$369** |
| Revenue growth, FY2026–30 | −1% to 3% | 3% to 6% | 6% to 9% |
| Operating margin by FY2030 | 15.5% | 20.0% | 22.0% |
| Terminal growth | 2.0% | 2.5% | 3.0% |
| Return on new capital, terminal | 12% | 20% | 25% |

**Base case: enterprise value $30.3B, equity $32.1B, $278 a share** on 115.6M shares (common plus
the exchangeable shares of Lulu Canadian Holding, which convert one for one). The discount rate is
the cost of equity, 7.72%: the 10-year Treasury at 4.20% on 17 March 2026, plus Damodaran's
January 2026 apparel beta (0.79, unlevered and corrected for cash) times his January 2026 US
equity risk premium (4.46%). lululemon has no borrowings, so this is also its WACC.

Three things a reader should know before trusting the number:

1. **79% of the value is the terminal value.** It is computed with the value-driver formula,
   NOPAT × (1 − g/RONIC) / (WACC − g), so growth has to be paid for with reinvestment rather than
   arriving free, and the implied exit multiple is 9.9x FY2030 EBITDA, a check on whether the
   perpetuity is plausible.
2. **The discount rate moves it most.** Across WACC 6.5% to 10% and terminal growth 1.5% to 3.5%,
   the base case runs from $189 to $435 a share. Using the specialty-retail beta (1.00) instead of
   apparel's raises the cost of equity to 8.66% and takes the base case to $238.
3. **Two conventions are deliberate.** Operating lease cost stays inside operating income, so the
   lease liabilities are not also subtracted as debt; and share-based compensation is treated as
   the real cost it is, not added back.

## How it is built and checked

- [`edgar_facts.json`](edgar_facts.json): 24 line items for five fiscal years, from
  `data.sec.gov`'s company facts, each with its tag, value and accession number. For each period
  the latest 10-K presentation wins, so a recast comparative replaces the original. The file was
  read in the browser and written here; the two sides hashed to the same SHA-256.
- [`market_inputs.json`](market_inputs.json) and [`assumptions.json`](assumptions.json): the
  market inputs with their sources, and the three cases with the story behind each.
- [`dcf_model.py`](dcf_model.py) computes the valuation in plain Python;
  [`build_workbook.py`](build_workbook.py) writes the same calculation as formulas, including the
  sensitivity grid, which re-discounts the base case's cash flows with `SUMPRODUCT` at each WACC and
  growth, and LibreOffice recalculates it.
- [`tests/test_dcf.py`](../tests/test_dcf.py) holds every valuation figure, forecast year, history
  ratio and sensitivity cell in the workbook to the Python twin; checks the filed figures against
  their own arithmetic (gross profit is revenue less cost of goods sold, net income is pre-tax income
  less tax, in every year); requires all 18 workbook checks to pass and no calculation sheet to hold
  a typed number; and holds this page to the numbers.
