import openpyxl
from openpyxl.utils import get_column_letter

f = r'C:\Users\it\Downloads\EFM_PRODUCTION_MASTER_BOM.xlsx'
wb = openpyxl.load_workbook(f, data_only=False)

print('=== ALL DATA VALUES ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    print(f'\n--- SHEET: {shname} ---')
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                print(f'  {cell.coordinate}: {repr(cell.value)}')

print('\n=== DATA VALIDATION ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    for dv in ws.data_validations.dataValidation:
        print(f'Sheet={shname} Range={dv.sqref} Type={dv.type} Formula1={dv.formula1} ShowDropDown={dv.showDropDown}')

print('\n=== CONDITIONAL FORMATTING ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    for cf_range, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            print(f'Sheet={shname} Range={cf_range} Type={rule.type} Priority={rule.priority}')
            if hasattr(rule, 'dxf') and rule.dxf:
                dxf = rule.dxf
                if dxf.fill and dxf.fill.fgColor:
                    fc = dxf.fill.fgColor
                    print(f'  Fill fgColor type={fc.type} value={fc.rgb if fc.type=="rgb" else fc.theme}')
                if dxf.font and dxf.font.color:
                    fc = dxf.font.color
                    print(f'  Font color type={fc.type} value={fc.rgb if fc.type=="rgb" else fc.theme}')
            if hasattr(rule, 'formula') and rule.formula:
                print(f'  Formula: {rule.formula}')
            if hasattr(rule, 'operator') and rule.operator:
                print(f'  Operator: {rule.operator}')
            if hasattr(rule, 'text') and rule.text:
                print(f'  Text: {rule.text}')

print('\n=== NAMED RANGES ===')
for name, defn in wb.defined_names.items():
    print(f'  {name}: {defn.attr_text}')

print('\n=== COL WIDTHS / ROW HEIGHTS ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    print(f'Sheet: {shname}')
    for col, dim in sorted(ws.column_dimensions.items()):
        if dim.width:
            print(f'  Col {col}: width={dim.width}')
    for row, dim in sorted(ws.row_dimensions.items()):
        if dim.height:
            print(f'  Row {row}: height={dim.height}')

print('\n=== MERGED CELLS ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    for m in ws.merged_cells.ranges:
        print(f'Sheet={shname} Merge={m}')

print('\n=== FORMULAS ===')
for shname in wb.sheetnames:
    ws = wb[shname]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).startswith('='):
                print(f'Sheet={shname} Cell={cell.coordinate} Formula={cell.value}')

