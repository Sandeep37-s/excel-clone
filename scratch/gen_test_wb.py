import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()

# Sheet 1: Main Calculations
ws1 = wb.active
ws1.title = "Formulas_Summary"

# Header styling
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF')
)

headers = ["Item Code", "Product", "Category", "Quantity", "Unit Price", "Total Price", "Discount Rate", "Final Amount", "Priority Check"]
for col_idx, h in enumerate(headers, 1):
    c = ws1.cell(row=1, column=col_idx, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

# Sample Data
data = [
    ("A101", "STM32 MCU", "Electronics", 10, 15.50),
    ("A102", "ADS1256 ADC", "Electronics", 5, 24.00),
    ("A103", "BLDC Motor", "Mechanical", 2, 85.00),
    ("A104", "Enclosure", "Mechanical", 4, 45.00),
    ("A105", "RJ45 MagJack", "Connectors", 20, 3.20),
    ("A106", "Optical Sensor", "Sensors", 8, 6.75),
]

for row_idx, (code, name, cat, qty, price) in enumerate(data, 2):
    ws1.cell(row=row_idx, column=1, value=code)
    ws1.cell(row=row_idx, column=2, value=name)
    ws1.cell(row=row_idx, column=3, value=cat)
    ws1.cell(row=row_idx, column=4, value=qty)
    ws1.cell(row=row_idx, column=5, value=price)
    # Total Price formula: Qty * Price
    ws1.cell(row=row_idx, column=6, value=f"=D{row_idx}*E{row_idx}")
    # Discount Rate with IF formula: IF(Quantity >= 10, 0.10, 0.05)
    ws1.cell(row=row_idx, column=7, value=f"=IF(D{row_idx}>=10, 0.10, 0.05)")
    # Final Amount formula: Total * (1 - Discount)
    ws1.cell(row=row_idx, column=8, value=f"=F{row_idx}*(1-G{row_idx})")
    # Priority Check: IF formula
    ws1.cell(row=row_idx, column=9, value=f"=IF(H{row_idx}>100, \"HIGH\", \"STANDARD\")")

# Summary section using SUM, AVERAGE, COUNTIF
r_sum = 9
ws1.cell(row=r_sum, column=3, value="Total Units:")
ws1.cell(row=r_sum, column=4, value=f"=SUM(D2:D7)")

r_total = 10
ws1.cell(row=r_total, column=5, value="Grand Total Amount:")
ws1.cell(row=r_total, column=6, value=f"=SUM(F2:F7)")
ws1.cell(row=r_total, column=8, value=f"=SUM(H2:H7)")

r_count = 11
ws1.cell(row=r_count, column=3, value="Electronics Count:")
ws1.cell(row=r_count, column=4, value=f"=COUNTIF(C2:C7, \"Electronics\")")

# Sheet 2: Cross-sheet Lookup Reference table
ws2 = wb.create_sheet(title="Lookup_Reference")
ws2.cell(row=1, column=1, value="Lookup Code").font = Font(bold=True)
ws2.cell(row=1, column=2, value="Tax Zone").font = Font(bold=True)
ws2.cell(row=1, column=3, value="Tariff Rate").font = Font(bold=True)

lookup_data = [
    ("A101", "Zone-1", 0.05),
    ("A102", "Zone-1", 0.05),
    ("A103", "Zone-2", 0.12),
    ("A104", "Zone-2", 0.12),
    ("A105", "Zone-3", 0.02),
    ("A106", "Zone-1", 0.05),
]
for idx, (c, z, r) in enumerate(lookup_data, 2):
    ws2.cell(row=idx, column=1, value=c)
    ws2.cell(row=idx, column=2, value=z)
    ws2.cell(row=idx, column=3, value=r)

# Sheet 3: VLOOKUP & Cross-Sheet test
ws3 = wb.create_sheet(title="CrossSheet_Tests")
ws3.cell(row=1, column=1, value="Query Code").font = Font(bold=True)
ws3.cell(row=1, column=2, value="Retrieved Category (VLOOKUP from Sheet1)").font = Font(bold=True)
ws3.cell(row=1, column=3, value="Retrieved Tariff (VLOOKUP from Sheet2)").font = Font(bold=True)
ws3.cell(row=1, column=4, value="Cross Sheet Sum Total").font = Font(bold=True)

test_codes = ["A101", "A103", "A105"]
for idx, q_code in enumerate(test_codes, 2):
    ws3.cell(row=idx, column=1, value=q_code)
    # VLOOKUP from Formulas_Summary sheet
    ws3.cell(row=idx, column=2, value=f"=VLOOKUP(A{idx}, Formulas_Summary!A2:E7, 3, FALSE)")
    # VLOOKUP from Lookup_Reference sheet
    ws3.cell(row=idx, column=3, value=f"=VLOOKUP(A{idx}, Lookup_Reference!A2:C7, 3, FALSE)")

# Cross sheet reference: sum of grand total from sheet 1
ws3.cell(row=2, column=4, value="=Formulas_Summary!H10")

# Adjust column widths
for ws in [ws1, ws2, ws3]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

test_file = r"C:\Users\it\Documents\4D PO\clone exc\test_comprehensive_formulas.xlsx"
wb.save(test_file)
print(f"Generated test workbook successfully: {test_file}")

