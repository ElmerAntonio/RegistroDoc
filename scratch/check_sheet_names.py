import openpyxl

wb = openpyxl.load_workbook("Registro_2026_FINAL_v2.xlsx")
for name in wb.sheetnames:
    if "Planilla" in name or "PROM" in name:
        print(repr(name))
wb.close()
