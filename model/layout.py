"""Where things live in the workbook, shared by the builder and the tests.

Every figure the tests check is reached through a defined name, so a layout
change in the builder cannot silently point a test at the wrong cell.
"""
WORKBOOK = "Kestrel_Bay_FPA_Model.xlsx"
COMPANY = "Kestrel Bay Provisions"

CATEGORIES = ["Beef", "Pork", "Poultry", "Seafood", "Charcuterie", "Cheese & Dairy"]
CHANNELS = ["Restaurants", "Grocery Retail", "Hotels & Institutions", "Online Direct"]
DEPARTMENTS = ["Sales & Marketing", "Warehouse & Logistics", "General & Admin", "Technology"]
SCENARIO_NAMES = ["Base", "Upside", "Downside"]
SENS_PRICE = [-0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03]
SENS_VOLUME = [-0.06, -0.04, -0.02, 0.0, 0.02, 0.04, 0.06]

SHEETS = [
    "Cover", "Dashboard", "PnL", "PVM", "Forecast", "Scenarios", "WorkingCapital",
    "Channels", "Checks", "Assumptions", "PQ_Budget", "Data_Sales", "Data_BudgetUnits",
    "Data_BudgetRates", "Data_Opex", "Data_Balances", "Data_Dims",
]

# P&L lines: key, label, is_cost, kind ("$", "%" or "cases")
PNL_LINES = [
    ("cases", "Cases sold", False, "cases"),
    ("revenue", "Revenue", False, "$"),
    ("cogs", "Cost of goods sold", True, "$"),
    ("gm", "Gross margin", False, "$"),
    ("gm_pct", "Gross margin %", False, "%"),
    ("freight", "Freight", True, "$"),
    ("contribution", "Contribution", False, "$"),
    ("opex_sm", "Sales & Marketing", True, "$"),
    ("opex_wl", "Warehouse & Logistics", True, "$"),
    ("opex_ga", "General & Admin", True, "$"),
    ("opex_tech", "Technology", True, "$"),
    ("opex_total", "Total operating expenses", True, "$"),
    ("ebitda", "EBITDA", False, "$"),
    ("ebitda_pct", "EBITDA margin %", False, "%"),
]
PNL_BLOCKS = {"outlook": 7, "budget": 24, "py": 41, "variance": 58}  # first line row
MONTH_COLS = "DEFGHIJKLMNO"                                         # Jan..Dec
