import xlsxwriter
import zipfile

wb = xlsxwriter.Workbook('test_xlsxwriter.xlsx')
ws = wb.add_worksheet('ANALISIS DE PUNIS')
ws.write('A1', 'Test')
ws.write('B1', 514704408)
ws.write('I1', '4299217233')
wb.close()

print('=== CONTENIDO XLSXWRITER ===')
with zipfile.ZipFile('test_xlsxwriter.xlsx', 'r') as z:
    for f in z.namelist():
        print(f'  {f}')
    
    # Ver el XML de la celda I1
    with z.open('xl/worksheets/sheet1.xml') as f:
        content = f.read().decode('utf-8')
        start = content.find('r="I1"')
        if start != -1:
            tag_start = content.rfind('<c ', 0, start)
            end = content.find('</c>', start) + 4
            print()
            print('Celda I1:')
            print(content[tag_start:end])
