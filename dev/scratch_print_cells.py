import openpyxl

wb = openpyxl.load_workbook("Registro_2026.xlsx")
ws = wb["Asistencia (9\u00b0)"]

print("Current Sheet contents (Rows 42-50):")
for r in range(42, 51):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 7)]
    print(f"Row {r}: {vals}")

wb.close()
