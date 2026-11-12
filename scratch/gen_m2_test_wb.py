import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()

# Sheet 1: Formatting & Styles
ws1 = wb.active
ws1.title = "Formatting_Styles"

ws1["A1"] = "ExcelClone Formatting & Number Formats Test"
ws1["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")

labels = [
    ("Currency", 12500.75, "$#,##0.00"),
    ("Percentage", 0.175, "0.00%"),
    ("Date Serial", 45600, "yyyy-mm-dd"),
    ("Scientific", 1234567, "0.00E+00"),
    ("Thousands Separator", 987654321, "#,##0"),
    ("Negative Red", -450.50, "$#,##0.00;[Red]($#,##0.00)"),
]

for idx, (label, val, fmt) in enumerate(labels, 3):
    ws1.cell(row=idx, column=1, value=label).font = Font(bold=True)
    c = ws1.cell(row=idx, column=2, value=val)
    c.number_format = fmt

# Merged Cells & Alignment
ws1.merge_cells("A11:C11")
m_cell = ws1["A11"]
m_cell.value = "Merged Banner (Centered, Gray Fill)"
m_cell.alignment = Alignment(horizontal="center", vertical="center")
m_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
m_cell.font = Font(bold=True)

# Sheet 2: Tables & Charts
ws2 = wb.create_sheet(title="Tables_And_Charts")
ws2.append(["Quarter", "Revenue ($)", "Expenses ($)", "Net Profit ($)"])

rows_data = [
    ["Q1", 120000, 75000, "=B2-C2"],
    ["Q2", 145000, 82000, "=B3-C3"],
    ["Q3", 160000, 91000, "=B4-C4"],
    ["Q4", 210000, 110000, "=B5-C5"],
]

for r in rows_data:
    ws2.append(r)

# Create Excel Table
tab = Table(displayName="QuarterlyFinancials", ref="A1:D5")
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style
ws2.add_table(tab)

# Bar Chart
chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "Quarterly Revenue & Expenses"
chart.y_axis.title = "USD"
chart.x_axis.title = "Quarter"

data_ref = Reference(ws2, min_col=2, min_row=1, max_col=3, max_row=5)
cats_ref = Reference(ws2, min_col=1, min_row=2, max_row=5)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
chart.width = 16
chart.height = 10
ws2.add_chart(chart, "F2")

# Sheet 3: Data Validation & Filters
ws3 = wb.create_sheet(title="DataValidation_Filters")
ws3.append(["Item ID", "Status (Dropdown)", "Quantity (1-100)", "Approval"])

items = [
    ["ITM-001", "Required", 10, "Approved"],
    ["ITM-002", "Optional", 5, "Pending"],
    ["ITM-003", "Required", 25, "Approved"],
    ["ITM-004", "In Review", 2, "Pending"],
]
for row in items:
    ws3.append(row)

# Add Dropdown Data Validation for Status Column
dv = DataValidation(type="list", formula1='"Required,Optional,In Review,Discontinued"', allow_blank=True)
ws3.add_data_validation(dv)
dv.add("B2:B10")

# AutoFilter
ws3.auto_filter.ref = "A1:D5"

test_m2_file = r"C:\Users\it\Documents\4D PO\clone exc\test_milestone2_features.xlsx"
wb.save(test_m2_file)
print(f"Generated Milestone 2 test workbook: {test_m2_file}")

