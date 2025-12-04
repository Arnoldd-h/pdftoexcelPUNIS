#!/usr/bin/env python3
"""
Abre el Excel generado para verificar si los valores son visibles
"""

import openpyxl

xlsx_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v(153148).xlsx"

print(f"Abriendo: {xlsx_path}\n")

wb = openpyxl.load_workbook(xlsx_path)
ws = wb.active

print("Verificando valores en columna J:")
print("=" * 60)

# Verificar primeras 20 celdas de columna J que deberían tener datos
test_rows = [12, 13, 16, 17, 18, 21, 24, 50, 51, 54, 55, 56, 57, 60, 61, 62, 63, 66, 92, 93]

for row in test_rows:
    cell = ws[f'J{row}']
    value = cell.value
    data_type = cell.data_type
    number_format = cell.number_format
    
    if value:
        print(f"J{row}: value='{value}' | data_type='{data_type}' | format='{number_format}'")
    else:
        print(f"J{row}: ❌ VACÍA")

wb.close()

print("\n\nSi todas las celdas muestran un valor, el problema es de PUNIS, no del Excel.")
