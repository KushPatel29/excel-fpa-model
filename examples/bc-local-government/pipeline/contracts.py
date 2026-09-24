"""Versioned source contracts for B.C. municipal statistics, 2019-2024."""

from __future__ import annotations

KEY_COLUMNS = ("Municipalities", "Type", "RD")

SCHEDULE_COLUMNS = {
    "201": {
        "population": (
            "BC Stats Population Estimates",
        ),
        "land_area_ha": ("Land Area (ha)",),
        "road_km": ("Distance of Roads (km)",),
    },
    "301": {
        "cash_and_investments": ("Cash and Investments",),
        "taxes_receivable": ("Taxes Receivable",),
        "total_financial_assets": ("Total Financial Assets",),
    },
    "302": {
        "accounts_payable": ("Accounts Payable and Accrued Liabilities",),
        "long_term_debt": ("Long-Term Debt",),
        "total_liabilities": ("Total Liabilities",),
    },
    "304": {
        "net_financial_assets": ("Net Financial Assets (Debt)",),
        "total_non_financial_assets": ("Total Non-Financial Assets",),
        "water_reserves": ("Water Statutory Reserves",),
        "sewer_reserves": ("Sewer Statutory Reserves",),
        "other_reserves": ("Other Statutory Reserves",),
        "equity_in_tca": ("Equity in Tangible Capital Assets",),
        "accumulated_surplus": ("Accumulated Surplus",),
    },
    "401": {
        "taxation_revenue": (
            "Total Own Purpose Taxation and Grants in Lieu",
        ),
        "sale_of_services": ("Sale of Services",),
        "federal_transfers": ("Federal Government Transfers",),
        "provincial_transfers": ("Provincial Government Transfers",),
        "regional_transfers": ("Regional and Other Governments Transfers",),
        "investment_income": ("Investment Income",),
        "government_business_income": (
            "Income from Government Business Enterprise",
        ),
        "developer_contributions": (
            "Developer and Other Contributions/ Donations",
        ),
        "gain_on_sale_assets": ("Gain on Sale of Assets",),
        "other_revenue": ("Other Revenue",),
        "total_revenue": ("Total Revenue",),
    },
    "402": {
        "general_government_expense": ("General Government",),
        "protective_services_expense": ("Protective Services",),
        "solid_waste_expense": ("Solid Waste Management and Recycling",),
        "health_social_housing_expense": ("Health, Social Services and Housing",),
        "development_services_expense": ("Development Services",),
        "transportation_expense": ("Transportation and Transit",),
        "parks_recreation_expense": ("Parks, Recreation and Culture",),
        "water_expense": ("Water Services",),
        "sewer_expense": ("Sewer Services",),
        "other_services_expense": ("Other Services",),
        "amortization": ("Amortization",),
        "asset_retirement_accretion": (
            "Asset Retirement Obligation Accretion",
        ),
        "loss_on_disposition": ("Loss on Disposition of Assets",),
        "other_adjustments_expense": ("Other Adjustments",),
        "total_expenses": ("Total Expenses",),
    },
    "502": {
        "total_tca_nbv": ("Total Tangible Capital Assets",),
        "total_non_financial_assets_502": ("Total Non-Financial Assets",),
    },
    "601_1": {
        "opening_debt": ("Total Debt at Year Start",),
        "new_debt": ("Proceeds from New Debt",),
        "debt_repayment": ("Debt  Repayment", "Debt Repayment"),
        "closing_debt": ("Total Debt at Year End",),
        "interest_expense": ("Interest Expense",),
        "debt_financing_cost": ("Debt Financing Cost in Year",),
    },
    "602_1": {
        "revenue_for_limit": ("Total Revenue for Purposes of Limit",),
        "liability_servicing_limit": ("Liability Servicing Limit",),
        "actual_debt_servicing": ("Actual Debt Servicing Cost",),
        "debt_servicing_capacity": (
            "Liability Servicing Capacity Available",
        ),
    },
    "706": {
        "tax_levy": ("Current Year Tax Levy",),
        "taxes_imposed": ("Total Taxes Imposed and Outstanding",),
        "current_taxes_collected": ("Current Year Taxes Collected",),
        "arrears_collected": ("Arrears and Delinquent Taxes Collected",),
        "total_taxes_collected": ("Total Taxes Collected",),
        "ending_taxes_receivable": ("Total Taxes Receivable at End of Year",),
    },
}

ASSET_COLUMNS = {
    "historical_cost": ("Historical Cost",),
    "accumulated_amortization": ("Accumulated Amortization",),
    "net_book_value": ("Net Book Value",),
    "asset_additions": ("Total Asset Additions in Year",),
    "replacement_value": ("Estimated Current Replacement Value",),
    "replacement_value_amortization": (
        "Estimated Annual Replacement Value Amortization",
    ),
    "asset_management_plan": ("Asset Management Plan in Place",),
}

ASSET_CATEGORIES = {
    "Buildings",
    "Drainage System",
    "Fleet",
    "Land",
    "Other Engineering Systems",
    "Other Tangible Capital Assets",
    "Parks and Recreation",
    "Sewer System",
    "Total TCA",
    "Transportation System",
    "Water System",
    "Work In Progress",
}

BLANK_MARKERS = {None, "", "No Data Submitted", "N/A", "n/a", "-"}

OPTIONAL_FIELDS = {("402", "asset_retirement_accretion")}
