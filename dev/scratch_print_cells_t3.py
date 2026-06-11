import openpyxl

wb = openpyxl.load_workbook("Registro_2026.xlsx")
ws = wb["Asistencia (9\u00b0)"]

print("Current Sheet contents (Rows 80-131):")
for r in range(80, 132):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 4)]
    print(f"Row {r:03d}: {vals}")

wb.close()
