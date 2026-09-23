"""Names shared by the builder and the tests.

Every figure the tests check is reached through a defined name, so a layout
change in the builder cannot silently point a test at the wrong cell.
"""
WORKBOOK = "PGE_Utility_FPA_Model.xlsx"
COMPANY = "Portland General Electric"
COMPANY_ID = 250                    # FERC Form 1 respondent id

CLASSES = ["residential", "commercial", "industrial", "transportation"]
CLASS_LABELS = {"residential": "Residential", "commercial": "Commercial",
                "industrial": "Industrial", "transportation": "Transportation"}
WEATHER_CLASSES = ["residential", "commercial"]

FERC_CLASSES = [  # label, FERC revenue_type
    ("Residential", "residential_sales"),
    ("Commercial", "small_or_commercial"),
    ("Industrial", "large_or_industrial"),
    ("Street lighting", "public_street_and_highway_lighting"),
    ("Public authorities", "other_sales_to_public_authorities"),
]
FERC_YEARS = list(range(2014, 2026))
PEERS = [  # FERC id, short name
    (250, "PGE"), (303, "PacifiCorp"), (162, "Puget Sound Energy"), (182, "Avista"), (216, "Idaho Power"),
]

WEATHER_SCENARIOS = [("Normal", 1.00, 1.00), ("Mild", 0.88, 1.00), ("Cold", 1.12, 1.00)]
SENS_RATE = [-0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04]
SENS_MW = [0, 50, 100, 150, 200, 250, 300]
BACKTEST_YEARS = [2023, 2024, 2025, 2026]

SHEETS = [
    "Cover", "Dashboard", "PnL", "PVM", "Monthly", "Weather", "Plan", "Forecast", "Scenarios",
    "Peers", "Explore", "Checks", "Assumptions", "PQ_Monthly", "PQ_Weather", "Data_EIA", "Data_NOAA",
    "Data_FERC", "Data_Utilities",
]
