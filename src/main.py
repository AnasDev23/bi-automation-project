"""
BI Automation & Reporting Pipeline
-----------------------------------
Extracts retail data from SQL Server, performs BI analysis,
and generates ONE timestamped Excel report containing:
- Executive Summary
- Interactive-style KPI Dashboard
- Sales and performance analysis
- Excel charts
- Conditional formatting
- Data quality information

Generated output:
    automated_reports/BI_Report_<timestamp>.xlsx
"""

from datetime import datetime
from pathlib import Path

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.page import PageMargins

from database import create_database_engine


# ============================================================
# PROJECT CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

AUTOMATED_REPORTS_DIR = BASE_DIR / "automated_reports"

SQL_QUERY = """
SELECT
    o.OrderID,
    o.CustomerID,
    o.EmployeeID,
    o.OrderDate,
    o.Sales,
    c.CustomerName,
    c.CityID,
    ci.CityName,
    e.EmployeeName,
    e.Department
FROM Orders AS o
INNER JOIN Customers AS c
    ON o.CustomerID = c.CustomerID
INNER JOIN Cities AS ci
    ON c.CityID = ci.CityID
INNER JOIN Employees AS e
    ON o.EmployeeID = e.EmployeeID;
"""


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def add_sales_contribution(summary: pd.DataFrame) -> pd.DataFrame:
    """Add each group's percentage contribution to total sales."""

    summary = summary.copy()

    total_sales = summary["Total_Sales"].sum()

    if total_sales == 0:
        summary["Sales_Contribution_%"] = 0.0
    else:
        summary["Sales_Contribution_%"] = (
            summary["Total_Sales"] / total_sales
        ) * 100

    return summary


def calculate_growth(series: pd.Series) -> pd.Series:
    """
    Calculate percentage-point growth.

    Values are multiplied by 100, so Excel formatting must use
    0.00"%" rather than 0.00%.
    """

    growth = series.pct_change(fill_method=None) * 100

    return growth.replace(
        [float("inf"), float("-inf")],
        float("nan")
    )


def get_top_row(
    dataframe: pd.DataFrame,
    column: str = "Total_Sales"
) -> pd.Series:
    """Return the row with the highest value in a column."""

    if dataframe.empty:
        raise ValueError(
            f"Cannot calculate top result because {column} data is empty."
        )

    return dataframe.loc[dataframe[column].idxmax()]


# ============================================================
# 1. LOAD DATA FROM SQL SERVER
# ============================================================

def load_source_data() -> pd.DataFrame:

    print("\n" + "=" * 65)
    print("BI AUTOMATION PROJECT")
    print("=" * 65)

    print("\nConnecting to SQL Server...")

    engine = create_database_engine()

    try:
        with engine.connect() as connection:
            df = pd.read_sql_query(SQL_QUERY, connection)

        print("SQL Server connection successful.")

    finally:
        engine.dispose()

    if df.empty:
        raise ValueError(
            "The SQL query returned no records. "
            "Please check your database tables and data."
        )

    print(f"Records retrieved: {len(df):,}")

    return df


# ============================================================
# 2. DATA VALIDATION AND PREPARATION
# ============================================================

def prepare_data(df: pd.DataFrame):

    print("\n--- DATA VALIDATION ---")

    original_rows = len(df)

    required_columns = [
        "OrderID",
        "CustomerID",
        "EmployeeID",
        "OrderDate",
        "Sales",
        "CustomerName",
        "CityName",
        "EmployeeName",
        "Department"
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Required database columns are missing: "
            + ", ".join(missing_columns)
        )

    # Convert sales to numeric values.
    df["Sales"] = pd.to_numeric(
        df["Sales"],
        errors="coerce"
    )

    # Convert dates to proper datetime values.
    df["OrderDate"] = pd.to_datetime(
        df["OrderDate"],
        errors="coerce"
    )

    # Capture missing values before excluding invalid records.
    missing_values_df = pd.DataFrame({
        "Column": df.columns,
        "Missing_Values": [
            int(df[column].isna().sum())
            for column in df.columns
        ]
    })

    invalid_sales_rows = int(df["Sales"].isna().sum())
    invalid_date_rows = int(df["OrderDate"].isna().sum())

    # Keep records with valid sales and order dates.
    df = df.dropna(
        subset=["Sales", "OrderDate"]
    ).copy()

    if df.empty:
        raise ValueError(
            "No valid records remain after date and sales validation."
        )

    negative_sales = int(
        (df["Sales"] < 0).sum()
    )

    # Add date-related analysis columns.
    df["Year"] = df["OrderDate"].dt.year

    df["Month"] = df["OrderDate"].dt.month

    df["YearMonth"] = (
        df["OrderDate"].dt.to_period("M").astype(str)
    )

    df["Quarter"] = df["OrderDate"].dt.quarter

    data_quality_summary = pd.DataFrame({
        "Metric": [
            "Rows Retrieved from Database",
            "Valid Rows Included in Analysis",
            "Rows Excluded",
            "Rows with Invalid Sales",
            "Rows with Invalid Order Dates",
            "Rows with Negative Sales",
            "Missing Cells in Retrieved Data"
        ],
        "Value": [
            original_rows,
            len(df),
            original_rows - len(df),
            invalid_sales_rows,
            invalid_date_rows,
            negative_sales,
            int(missing_values_df["Missing_Values"].sum())
        ]
    })

    print(f"Valid records: {len(df):,}")
    print(f"Invalid sales records: {invalid_sales_rows:,}")
    print(f"Invalid date records: {invalid_date_rows:,}")
    print(f"Negative sales records: {negative_sales:,}")

    print("\nData validation completed.")

    return df, data_quality_summary, missing_values_df


# ============================================================
# 3. BUSINESS INTELLIGENCE ANALYSIS
# ============================================================

def perform_bi_analysis(df: pd.DataFrame):

    print("\n--- CALCULATING BUSINESS INTELLIGENCE METRICS ---")

    total_orders = int(df["OrderID"].nunique())

    total_sales = float(df["Sales"].sum())

    average_order_value = (
        total_sales / total_orders
        if total_orders
        else 0
    )

    minimum_order_value = float(df["Sales"].min())

    maximum_order_value = float(df["Sales"].max())

    # --------------------------------------------------------
    # MONTHLY SALES
    # --------------------------------------------------------

    monthly_summary = (
        df.groupby("YearMonth")
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum")
        )
        .reset_index()
        .sort_values("YearMonth")
    )

    monthly_summary["Previous_Month_Sales"] = (
        monthly_summary["Total_Sales"].shift(1)
    )

    monthly_summary["MoM_Growth_%"] = calculate_growth(
        monthly_summary["Total_Sales"]
    )

    # --------------------------------------------------------
    # YEARLY SALES
    # --------------------------------------------------------

    yearly_performance = (
        df.groupby("Year")
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values("Year")
    )

    yearly_performance["Previous_Year_Sales"] = (
        yearly_performance["Total_Sales"].shift(1)
    )

    yearly_performance["YoY_Growth_%"] = calculate_growth(
        yearly_performance["Total_Sales"]
    )

    # --------------------------------------------------------
    # QUARTERLY SALES
    # --------------------------------------------------------

    quarterly_summary = (
        df.groupby(["Year", "Quarter"])
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values(["Year", "Quarter"])
    )

    # --------------------------------------------------------
    # EMPLOYEE PERFORMANCE
    # --------------------------------------------------------

    employee_performance = (
        df.groupby(["EmployeeID", "EmployeeName"])
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values("Total_Sales", ascending=False)
    )

    employee_performance = add_sales_contribution(
        employee_performance
    )

    # --------------------------------------------------------
    # CUSTOMER PERFORMANCE
    # --------------------------------------------------------

    customer_performance = (
        df.groupby(["CustomerID", "CustomerName"])
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values("Total_Sales", ascending=False)
    )

    customer_performance = add_sales_contribution(
        customer_performance
    )

    # --------------------------------------------------------
    # CITY PERFORMANCE
    # --------------------------------------------------------

    city_performance = (
        df.groupby("CityName")
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values("Total_Sales", ascending=False)
    )

    city_performance = add_sales_contribution(
        city_performance
    )

    # --------------------------------------------------------
    # DEPARTMENT PERFORMANCE
    # --------------------------------------------------------

    department_performance = (
        df.groupby("Department")
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum"),
            Average_Order_Value=("Sales", "mean")
        )
        .reset_index()
        .sort_values("Total_Sales", ascending=False)
    )

    department_performance = add_sales_contribution(
        department_performance
    )

    # --------------------------------------------------------
    # CITY + DEPARTMENT ANALYSIS
    # --------------------------------------------------------

    city_department_summary = (
        df.groupby(["CityName", "Department"])
        .agg(
            Total_Orders=("OrderID", "nunique"),
            Total_Sales=("Sales", "sum")
        )
        .reset_index()
        .sort_values("Total_Sales", ascending=False)
    )

    city_department_pivot = pd.pivot_table(
        df,
        values="Sales",
        index="CityName",
        columns="Department",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # --------------------------------------------------------
    # TOP CUSTOMERS AND CITIES
    # --------------------------------------------------------

    top_customers = customer_performance.head(10).copy()

    top_cities = city_performance.head(10).copy()

    # --------------------------------------------------------
    # SALES DISTRIBUTION STATISTICS
    # --------------------------------------------------------

    sales_statistics = pd.DataFrame({
        "Statistic": [
            "Mean",
            "Median",
            "Standard Deviation",
            "Minimum",
            "Maximum",
            "25th Percentile",
            "75th Percentile"
        ],
        "Value": [
            df["Sales"].mean(),
            df["Sales"].median(),
            df["Sales"].std(),
            df["Sales"].min(),
            df["Sales"].max(),
            df["Sales"].quantile(0.25),
            df["Sales"].quantile(0.75)
        ]
    })

    # --------------------------------------------------------
    # KPI SUMMARY
    # --------------------------------------------------------

    kpi_summary = pd.DataFrame({
        "KPI": [
            "Total Orders",
            "Total Sales",
            "Average Order Value",
            "Minimum Order Value",
            "Maximum Order Value",
            "Unique Customers",
            "Unique Employees",
            "Unique Cities",
            "Unique Departments"
        ],
        "Value": [
            total_orders,
            total_sales,
            average_order_value,
            minimum_order_value,
            maximum_order_value,
            int(df["CustomerID"].nunique()),
            int(df["EmployeeID"].nunique()),
            int(df["CityName"].nunique()),
            int(df["Department"].nunique())
        ]
    })

    # --------------------------------------------------------
    # KEY BUSINESS HIGHLIGHTS
    # --------------------------------------------------------

    best_year = get_top_row(
        yearly_performance
    )

    worst_year = yearly_performance.loc[
        yearly_performance["Total_Sales"].idxmin()
    ]

    best_month = get_top_row(
        monthly_summary
    )

    worst_month = monthly_summary.loc[
        monthly_summary["Total_Sales"].idxmin()
    ]

    best_quarter = get_top_row(
        quarterly_summary
    )

    worst_quarter = quarterly_summary.loc[
        quarterly_summary["Total_Sales"].idxmin()
    ]

    top_customer = get_top_row(
        customer_performance
    )

    top_city = get_top_row(
        city_performance
    )

    top_department = get_top_row(
        department_performance
    )

    print("Total Orders:", f"{total_orders:,}")
    print("Total Sales:", f"{total_sales:,.2f}")
    print("Average Order Value:", f"{average_order_value:,.2f}")

    print("Highest Sales Year:", int(best_year["Year"]))
    print("Top Customer:", top_customer["CustomerName"])
    print("Top City:", top_city["CityName"])
    print("Top Department:", top_department["Department"])

    print("BI analysis completed successfully.")

    # --------------------------------------------------------
    # RETURN ANALYSIS RESULTS
    # --------------------------------------------------------

    dataframes = {
        "KPI Summary": kpi_summary,
        "Monthly Sales": monthly_summary,
        "Yearly Sales": yearly_performance,
        "Quarterly Sales": quarterly_summary,
        "Employee Performance": employee_performance,
        "Customer Performance": customer_performance,
        "City Performance": city_performance,
        "Department Performance": department_performance,
        "Top Customers": top_customers,
        "Top Cities": top_cities,
        "City-Department": city_department_summary,
        "City-Dept Pivot": city_department_pivot,
        "Sales Statistics": sales_statistics
    }

    highlights = {
        "total_orders": total_orders,
        "total_sales": total_sales,
        "average_order_value": average_order_value,
        "unique_customers": int(df["CustomerID"].nunique()),
        "unique_employees": int(df["EmployeeID"].nunique()),
        "unique_cities": int(df["CityName"].nunique()),
        "best_year": best_year,
        "worst_year": worst_year,
        "best_month": best_month,
        "worst_month": worst_month,
        "best_quarter": best_quarter,
        "worst_quarter": worst_quarter,
        "top_customer": top_customer,
        "top_city": top_city,
        "top_department": top_department
    }

    return dataframes, highlights


# ============================================================
# 4. EXCEL STYLE CONFIGURATION
# ============================================================

NAVY = "17365D"
BLUE = "1F4E78"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EAF3F8"
WHITE = "FFFFFF"
DARK_TEXT = "243447"
GREEN_FILL = "E2F0D9"
RED_FILL = "FCE4D6"
BORDER_COLOR = "B4C7E7"

THIN_SIDE = Side(
    style="thin",
    color=BORDER_COLOR
)

CARD_BORDER = Border(
    left=THIN_SIDE,
    right=THIN_SIDE,
    top=THIN_SIDE,
    bottom=THIN_SIDE
)

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor=NAVY
)

HEADER_FONT = Font(
    bold=True,
    color=WHITE,
    size=10
)

TITLE_FILL = PatternFill(
    fill_type="solid",
    fgColor=BLUE
)

TITLE_FONT = Font(
    bold=True,
    size=22,
    color=WHITE
)

LABEL_FILL = PatternFill(
    fill_type="solid",
    fgColor=LIGHT_BLUE
)

VALUE_FILL = PatternFill(
    fill_type="solid",
    fgColor=PALE_BLUE
)

LABEL_FONT = Font(
    bold=True,
    size=10,
    color=DARK_TEXT
)

VALUE_FONT = Font(
    bold=True,
    size=15,
    color=BLUE
)


# ============================================================
# 5. DASHBOARD KPI CARD FUNCTION
# ============================================================

def create_kpi_card(
    worksheet,
    label_range: str,
    value_range: str,
    label: str,
    value: str
):
    """Create a dashboard card with a readable text value."""

    worksheet.merge_cells(label_range)
    worksheet.merge_cells(value_range)

    label_start = label_range.split(":")[0]
    value_start = value_range.split(":")[0]

    label_cell = worksheet[label_start]
    value_cell = worksheet[value_start]

    label_cell.value = label
    value_cell.value = str(value)

    label_cell.font = LABEL_FONT
    value_cell.font = VALUE_FONT

    label_cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    value_cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    for row in worksheet[label_range]:
        for cell in row:
            cell.fill = LABEL_FILL
            cell.border = CARD_BORDER

    for row in worksheet[value_range]:
        for cell in row:
            cell.fill = VALUE_FILL
            cell.border = CARD_BORDER


# ============================================================
# 6. CREATE DASHBOARD
# ============================================================

def create_dashboard(workbook, highlights):

    dashboard = workbook.create_sheet(
        "Dashboard",
        0
    )

    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    dashboard.column_dimensions["A"].width = 3

    for column in range(2, 18):
        letter = dashboard.cell(
            row=1,
            column=column
        ).column_letter

        dashboard.column_dimensions[letter].width = 14

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    dashboard.merge_cells("B2:Q3")

    dashboard["B2"] = "BI AUTOMATION DASHBOARD"

    dashboard["B2"].font = TITLE_FONT

    dashboard["B2"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    for row in dashboard["B2:Q3"]:
        for cell in row:
            cell.fill = TITLE_FILL

    dashboard.row_dimensions[2].height = 28
    dashboard.row_dimensions[3].height = 28

    # --------------------------------------------------------
    # REPORT TIMESTAMP
    # --------------------------------------------------------

    dashboard.merge_cells("B4:Q4")

    dashboard["B4"] = (
        "Generated on: "
        + datetime.now().strftime("%d %B %Y, %I:%M:%S %p")
    )

    dashboard["B4"].font = Font(
        italic=True,
        size=9,
        color="666666"
    )

    dashboard["B4"].alignment = Alignment(
        horizontal="center"
    )

    # --------------------------------------------------------
    # FIRST ROW KPI CARDS
    # --------------------------------------------------------

    create_kpi_card(
        dashboard,
        "B6:D6",
        "B7:D8",
        "TOTAL ORDERS",
        f"{highlights['total_orders']:,}"
    )

    create_kpi_card(
        dashboard,
        "F6:H6",
        "F7:H8",
        "TOTAL SALES",
        f"{highlights['total_sales']:,.2f}"
    )

    create_kpi_card(
        dashboard,
        "J6:L6",
        "J7:L8",
        "AVERAGE ORDER VALUE",
        f"{highlights['average_order_value']:,.2f}"
    )

    # --------------------------------------------------------
    # SECOND ROW KPI CARDS
    # --------------------------------------------------------

    create_kpi_card(
        dashboard,
        "B10:D10",
        "B11:D12",
        "HIGHEST SALES YEAR",
        str(int(highlights["best_year"]["Year"]))
    )

    create_kpi_card(
        dashboard,
        "F10:H10",
        "F11:H12",
        "TOP CITY",
        highlights["top_city"]["CityName"]
    )

    create_kpi_card(
        dashboard,
        "J10:L10",
        "J11:L12",
        "TOP DEPARTMENT",
        highlights["top_department"]["Department"]
    )

    for row in [6, 7, 8, 10, 11, 12]:
        dashboard.row_dimensions[row].height = 25

    # --------------------------------------------------------
    # CHART SOURCES
    # --------------------------------------------------------

    monthly_sheet = workbook["Monthly Sales"]
    yearly_sheet = workbook["Yearly Sales"]
    city_sheet = workbook["Top Cities"]
    department_sheet = workbook["Department Performance"]

    # --------------------------------------------------------
    # MONTHLY SALES LINE CHART
    # --------------------------------------------------------

    monthly_chart = LineChart()

    monthly_chart.title = "Monthly Sales Trend"
    monthly_chart.y_axis.title = "Sales"
    monthly_chart.x_axis.title = "Year / Month"

    monthly_chart.style = 13

    monthly_data = Reference(
        monthly_sheet,
        min_col=3,
        min_row=1,
        max_row=monthly_sheet.max_row
    )

    monthly_categories = Reference(
        monthly_sheet,
        min_col=1,
        min_row=2,
        max_row=monthly_sheet.max_row
    )

    monthly_chart.add_data(
        monthly_data,
        titles_from_data=True
    )

    monthly_chart.set_categories(
        monthly_categories
    )

    monthly_chart.height = 8
    monthly_chart.width = 15

    dashboard.add_chart(
        monthly_chart,
        "B15"
    )

    # --------------------------------------------------------
    # YEARLY SALES COLUMN CHART
    # --------------------------------------------------------

    yearly_chart = BarChart()

    yearly_chart.type = "col"

    yearly_chart.title = "Yearly Sales Comparison"
    yearly_chart.y_axis.title = "Sales"
    yearly_chart.x_axis.title = "Year"

    yearly_chart.style = 10

    yearly_data = Reference(
        yearly_sheet,
        min_col=3,
        min_row=1,
        max_row=yearly_sheet.max_row
    )

    yearly_categories = Reference(
        yearly_sheet,
        min_col=1,
        min_row=2,
        max_row=yearly_sheet.max_row
    )

    yearly_chart.add_data(
        yearly_data,
        titles_from_data=True
    )

    yearly_chart.set_categories(
        yearly_categories
    )

    yearly_chart.height = 8
    yearly_chart.width = 15

    dashboard.add_chart(
        yearly_chart,
        "J15"
    )

    # --------------------------------------------------------
    # TOP CITIES BAR CHART
    # --------------------------------------------------------

    city_chart = BarChart()

    city_chart.type = "bar"

    city_chart.title = "Top 10 Cities by Sales"
    city_chart.x_axis.title = "Sales"
    city_chart.y_axis.title = "City"

    city_chart.style = 12

    city_max_row = min(
        city_sheet.max_row,
        11
    )

    city_data = Reference(
        city_sheet,
        min_col=3,
        min_row=1,
        max_row=city_max_row
    )

    city_categories = Reference(
        city_sheet,
        min_col=1,
        min_row=2,
        max_row=city_max_row
    )

    city_chart.add_data(
        city_data,
        titles_from_data=True
    )

    city_chart.set_categories(
        city_categories
    )

    city_chart.height = 9
    city_chart.width = 15

    dashboard.add_chart(
        city_chart,
        "B32"
    )

    # --------------------------------------------------------
    # DEPARTMENT PERFORMANCE BAR CHART
    # --------------------------------------------------------

    department_chart = BarChart()

    department_chart.type = "bar"

    department_chart.title = "Department Sales Performance"
    department_chart.x_axis.title = "Sales"
    department_chart.y_axis.title = "Department"

    department_chart.style = 12

    department_data = Reference(
        department_sheet,
        min_col=3,
        min_row=1,
        max_row=department_sheet.max_row
    )

    department_categories = Reference(
        department_sheet,
        min_col=1,
        min_row=2,
        max_row=department_sheet.max_row
    )

    department_chart.add_data(
        department_data,
        titles_from_data=True
    )

    department_chart.set_categories(
        department_categories
    )

    department_chart.height = 9
    department_chart.width = 15

    dashboard.add_chart(
        department_chart,
        "J32"
    )

    # --------------------------------------------------------
    # DASHBOARD VIEW AND PRINT SETTINGS
    # --------------------------------------------------------

    dashboard.freeze_panes = "B5"

    dashboard.sheet_view.showGridLines = False

    dashboard.sheet_view.zoomScale = 85

    dashboard.page_setup.orientation = "landscape"

    dashboard.page_setup.fitToWidth = 1
    dashboard.page_setup.fitToHeight = 0

    dashboard.sheet_properties.pageSetUpPr.fitToPage = True

    dashboard.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.5,
        bottom=0.5,
        header=0.2,
        footer=0.2
    )

    dashboard.print_area = "A1:Q50"

    return dashboard


# ============================================================
# 7. CREATE EXECUTIVE SUMMARY
# ============================================================

def create_executive_summary(workbook, highlights):

    worksheet = workbook.create_sheet(
        "Executive Summary",
        1
    )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    worksheet.merge_cells("A1:F2")

    worksheet["A1"] = "BI AUTOMATION — EXECUTIVE SUMMARY"

    worksheet["A1"].font = Font(
        size=18,
        bold=True,
        color=WHITE
    )

    worksheet["A1"].fill = TITLE_FILL

    worksheet["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    for row in worksheet["A1:F2"]:
        for cell in row:
            cell.fill = TITLE_FILL

    worksheet.row_dimensions[1].height = 27
    worksheet.row_dimensions[2].height = 27

    # --------------------------------------------------------
    # SECTION HEADINGS
    # --------------------------------------------------------

    worksheet.merge_cells("A4:B4")
    worksheet.merge_cells("D4:F4")

    worksheet["A4"] = "KEY PERFORMANCE INDICATORS"
    worksheet["D4"] = "BUSINESS HIGHLIGHTS"

    for cell in ["A4", "D4"]:
        worksheet[cell].font = Font(
            bold=True,
            color=WHITE,
            size=11
        )

        worksheet[cell].fill = PatternFill(
            fill_type="solid",
            fgColor=BLUE
        )

        worksheet[cell].alignment = Alignment(
            horizontal="center"
        )

    # --------------------------------------------------------
    # KEY PERFORMANCE INDICATORS
    # --------------------------------------------------------

    kpis = [
        ("Total Orders", highlights["total_orders"], "#,##0"),
        ("Total Sales", highlights["total_sales"], "#,##0.00"),
        (
            "Average Order Value",
            highlights["average_order_value"],
            "#,##0.00"
        ),
        ("Unique Customers", highlights["unique_customers"], "#,##0"),
        ("Unique Employees", highlights["unique_employees"], "#,##0"),
        ("Unique Cities", highlights["unique_cities"], "#,##0")
    ]

    for row_number, (label, value, number_format) in enumerate(
        kpis,
        start=5
    ):
        worksheet.cell(
            row=row_number,
            column=1,
            value=label
        )

        value_cell = worksheet.cell(
            row=row_number,
            column=2,
            value=value
        )

        value_cell.number_format = number_format

    # --------------------------------------------------------
    # BUSINESS HIGHLIGHTS
    # --------------------------------------------------------

    best_year = highlights["best_year"]
    worst_year = highlights["worst_year"]
    best_month = highlights["best_month"]
    worst_month = highlights["worst_month"]
    best_quarter = highlights["best_quarter"]
    worst_quarter = highlights["worst_quarter"]

    business_highlights = [
        (
            "Highest Sales Year",
            str(int(best_year["Year"])),
            best_year["Total_Sales"]
        ),
        (
            "Lowest Sales Year",
            str(int(worst_year["Year"])),
            worst_year["Total_Sales"]
        ),
        (
            "Top City",
            str(highlights["top_city"]["CityName"]),
            highlights["top_city"]["Total_Sales"]
        ),
        (
            "Top Department",
            str(highlights["top_department"]["Department"]),
            highlights["top_department"]["Total_Sales"]
        ),
        (
            "Best Sales Month",
            str(best_month["YearMonth"]),
            best_month["Total_Sales"]
        ),
        (
            "Lowest Sales Month",
            str(worst_month["YearMonth"]),
            worst_month["Total_Sales"]
        ),
        (
            "Best Sales Quarter",
            f"{int(best_quarter['Year'])} Q{int(best_quarter['Quarter'])}",
            best_quarter["Total_Sales"]
        ),
        (
            "Lowest Sales Quarter",
            f"{int(worst_quarter['Year'])} Q{int(worst_quarter['Quarter'])}",
            worst_quarter["Total_Sales"]
        )
    ]

    for row_number, (label, period, sales) in enumerate(
        business_highlights,
        start=5
    ):
        worksheet.cell(
            row=row_number,
            column=4,
            value=label
        )

        worksheet.cell(
            row=row_number,
            column=5,
            value=period
        )

        sales_cell = worksheet.cell(
            row=row_number,
            column=6,
            value=float(sales)
        )

        sales_cell.number_format = "#,##0.00"

    # --------------------------------------------------------
    # AUTOMATIC BUSINESS OBSERVATIONS
    # --------------------------------------------------------

    worksheet.merge_cells("A15:F15")

    worksheet["A15"] = "AUTOMATIC BUSINESS OBSERVATIONS"

    worksheet["A15"].font = Font(
        bold=True,
        color=WHITE,
        size=11
    )

    worksheet["A15"].fill = PatternFill(
        fill_type="solid",
        fgColor=BLUE
    )

    observations = [
        (
            f"The database analysis contains "
            f"{highlights['total_orders']:,} unique orders."
        ),
        (
            f"Total recorded sales amount to "
            f"{highlights['total_sales']:,.2f}."
        ),
        (
            f"The average order value is "
            f"{highlights['average_order_value']:,.2f}."
        ),
        (
            f"{int(best_year['Year'])} recorded the highest "
            f"annual sales of {best_year['Total_Sales']:,.2f}."
        ),
        (
            f"{highlights['top_city']['CityName']} is the top city "
            f"by sales, with "
            f"{highlights['top_city']['Total_Sales']:,.2f}."
        ),
        (
            f"{highlights['top_department']['Department']} is the "
            f"highest-performing department by sales."
        ),
        (
            f"The top customer is "
            f"{highlights['top_customer']['CustomerName']}."
        ),
        (
            f"{int(best_quarter['Year'])} Q"
            f"{int(best_quarter['Quarter'])} recorded the highest "
            f"quarterly sales."
        )
    ]

    for row_number, observation in enumerate(
        observations,
        start=16
    ):
        worksheet.merge_cells(
            start_row=row_number,
            start_column=1,
            end_row=row_number,
            end_column=6
        )

        cell = worksheet.cell(
            row=row_number,
            column=1,
            value=f"• {observation}"
        )

        cell.alignment = Alignment(
            wrap_text=True,
            vertical="center"
        )

        cell.font = Font(
            size=10,
            color=DARK_TEXT
        )

        worksheet.row_dimensions[row_number].height = 25

    # --------------------------------------------------------
    # FORMATTING
    # --------------------------------------------------------

    for row in worksheet.iter_rows(
        min_row=5,
        max_row=12,
        min_col=1,
        max_col=6
    ):
        for cell in row:
            cell.border = CARD_BORDER
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

    worksheet.column_dimensions["A"].width = 27
    worksheet.column_dimensions["B"].width = 19
    worksheet.column_dimensions["C"].width = 4
    worksheet.column_dimensions["D"].width = 25
    worksheet.column_dimensions["E"].width = 19
    worksheet.column_dimensions["F"].width = 21

    worksheet.freeze_panes = "A5"

    worksheet.sheet_view.showGridLines = False

    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0

    worksheet.sheet_properties.pageSetUpPr.fitToPage = True

    return worksheet


# ============================================================
# 8. FORMAT ANALYSIS WORKSHEETS
# ============================================================

def format_analysis_sheet(worksheet):

    if worksheet.max_row < 1:
        return

    worksheet.sheet_view.showGridLines = False

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = worksheet.dimensions

    worksheet.row_dimensions[1].height = 30

    # Header styling
    for cell in worksheet[1]:

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Column sizing and data formatting
    for column_index in range(
        1,
        worksheet.max_column + 1
    ):

        column_letter = worksheet.cell(
            row=1,
            column=column_index
        ).column_letter

        header = str(
            worksheet.cell(
                row=1,
                column=column_index
            ).value or ""
        )

        header_lower = header.lower()

        # Automatic width with a practical maximum.
        max_length = max(
            (
                len(str(cell.value))
                for cell in worksheet[column_letter]
                if cell.value is not None
            ),
            default=10
        )

        width = max(
            13,
            min(max_length + 3, 36)
        )

        worksheet.column_dimensions[
            column_letter
        ].width = width

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):

            cell = worksheet.cell(
                row=row_number,
                column=column_index
            )

            if cell.value is None:
                continue

            cell.alignment = Alignment(
                vertical="center"
            )

            # All percentages in this project are stored
            # as percentage points, for example 12.5 = 12.5%.
            if "%" in header:
                cell.number_format = '0.00"%"'

            # Sales and monetary values.
            elif any(
                term in header_lower
                for term in [
                    "total_sales",
                    "previous_year_sales",
                    "previous_month_sales",
                    "average_order_value"
                ]
            ):
                cell.number_format = "#,##0.00"

            # Order and ID counts.
            elif (
                "total_orders" in header_lower
                or "missing_values" in header_lower
                or header_lower in [
                    "customerid",
                    "employeeid",
                    "cityid"
                ]
            ):
                cell.number_format = "#,##0"

    # Special handling for KPI Summary mixed numeric values.
    if worksheet.title == "KPI Summary":

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):

            label = str(
                worksheet.cell(row=row_number, column=1).value
            )

            value_cell = worksheet.cell(
                row=row_number,
                column=2
            )

            if any(
                term in label
                for term in [
                    "Orders",
                    "Customers",
                    "Employees",
                    "Cities",
                    "Departments"
                ]
            ):
                value_cell.number_format = "#,##0"

            else:
                value_cell.number_format = "#,##0.00"

    # Sales statistics values.
    if worksheet.title == "Sales Statistics":

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):
            worksheet.cell(
                row=row_number,
                column=2
            ).number_format = "#,##0.00"

    # Data quality counts.
    if worksheet.title == "Data Quality":

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):
            worksheet.cell(
                row=row_number,
                column=2
            ).number_format = "#,##0"

    # --------------------------------------------------------
    # CONDITIONAL FORMATTING FOR GROWTH
    # --------------------------------------------------------

    for column_index in range(
        1,
        worksheet.max_column + 1
    ):

        header = str(
            worksheet.cell(
                row=1,
                column=column_index
            ).value or ""
        )

        if header in ["MoM_Growth_%", "YoY_Growth_%"]:

            column_letter = worksheet.cell(
                row=1,
                column=column_index
            ).column_letter

            cell_range = (
                f"{column_letter}2:"
                f"{column_letter}{worksheet.max_row}"
            )

            worksheet.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="greaterThan",
                    formula=["0"],
                    fill=PatternFill(
                        fill_type="solid",
                        fgColor=GREEN_FILL
                    )
                )
            )

            worksheet.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="lessThan",
                    formula=["0"],
                    fill=PatternFill(
                        fill_type="solid",
                        fgColor=RED_FILL
                    )
                )
            )

    # --------------------------------------------------------
    # DATA BARS FOR SALES
    # --------------------------------------------------------

    for column_index in range(
        1,
        worksheet.max_column + 1
    ):

        header = str(
            worksheet.cell(
                row=1,
                column=column_index
            ).value or ""
        )

        if header == "Total_Sales" and worksheet.max_row >= 2:

            column_letter = worksheet.cell(
                row=1,
                column=column_index
            ).column_letter

            worksheet.conditional_formatting.add(
                f"{column_letter}2:{column_letter}{worksheet.max_row}",
                DataBarRule(
                    start_type="min",
                    end_type="max",
                    color="5B9BD5",
                    showValue=True
                )
            )


# ============================================================
# 9. GENERATE ONE AUTOMATED EXCEL REPORT
# ============================================================

def generate_automated_report(
    dataframes: dict,
    highlights: dict,
    data_quality_summary: pd.DataFrame,
    missing_values_df: pd.DataFrame
) -> Path:

    print("\n--- GENERATING AUTOMATED EXCEL REPORT ---")

    AUTOMATED_REPORTS_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    # Use a timestamp with microseconds to avoid filename collisions.
    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S_%f"
    )

    output_file = (
        AUTOMATED_REPORTS_DIR
        / f"BI_Report_{timestamp}.xlsx"
    )

    # --------------------------------------------------------
    # WRITE ALL DATA SHEETS DIRECTLY TO FINAL REPORT
    # --------------------------------------------------------

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        for sheet_name, dataframe in dataframes.items():

            dataframe.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

        data_quality_summary.to_excel(
            writer,
            sheet_name="Data Quality",
            index=False
        )

        missing_values_df.to_excel(
            writer,
            sheet_name="Missing Values",
            index=False
        )

    print("Analysis sheets written.")

    # --------------------------------------------------------
    # OPEN THE FINAL WORKBOOK
    # --------------------------------------------------------

    workbook = load_workbook(output_file)

    # Create dashboard and executive summary inside
    # the same final automated workbook.
    create_dashboard(
        workbook,
        highlights
    )

    create_executive_summary(
        workbook,
        highlights
    )

    # --------------------------------------------------------
    # FORMAT ALL ANALYSIS SHEETS
    # --------------------------------------------------------

    excluded_sheets = {
        "Dashboard",
        "Executive Summary"
    }

    for worksheet in workbook.worksheets:

        if worksheet.title not in excluded_sheets:
            format_analysis_sheet(worksheet)

    # --------------------------------------------------------
    # SAVE FINAL WORKBOOK
    # --------------------------------------------------------

    workbook.active = 0

    workbook.save(output_file)

    return output_file


# ============================================================
# 10. MAIN EXECUTION
# ============================================================

def main():

    start_time = datetime.now()

    try:

        # Load database records.
        df = load_source_data()

        # Validate and prepare data.
        (
            df,
            data_quality_summary,
            missing_values_df
        ) = prepare_data(df)

        # Perform all BI calculations.
        dataframes, highlights = perform_bi_analysis(df)

        # Generate one final Excel workbook.
        report_file = generate_automated_report(
            dataframes=dataframes,
            highlights=highlights,
            data_quality_summary=data_quality_summary,
            missing_values_df=missing_values_df
        )

        execution_seconds = (
            datetime.now() - start_time
        ).total_seconds()

        report_size_kb = report_file.stat().st_size / 1024

        print("\n" + "=" * 65)
        print("BI AUTOMATION COMPLETED SUCCESSFULLY")
        print("=" * 65)

        print("\nFINAL AUTOMATED REPORT")
        print("----------------------")

        print("File:", report_file.name)
        print("Location:", report_file)
        print(f"Size: {report_size_kb:,.2f} KB")
        print(f"Execution Time: {execution_seconds:.2f} seconds")

        print("\nREPORT CONTENTS")
        print("---------------")

        print("Dashboard")
        print("Executive Summary")
        print("KPI Summary")
        print("Monthly Sales")
        print("Yearly Sales")
        print("Quarterly Sales")
        print("Employee Performance")
        print("Customer Performance")
        print("City Performance")
        print("Department Performance")
        print("Top Customers")
        print("Top Cities")
        print("City-Department Analysis")
        print("City-Department Pivot")
        print("Sales Statistics")
        print("Data Quality")
        print("Missing Values")

        print("\nOnly ONE automated Excel report was generated.")

        print("=" * 65)

    except Exception as error:

        print("\n" + "=" * 65)
        print("BI AUTOMATION FAILED")
        print("=" * 65)

        print("Error:", str(error))

        # Preserve the traceback for debugging in PyCharm.
        raise


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()