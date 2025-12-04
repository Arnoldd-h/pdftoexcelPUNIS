import zipfile

def check_cell(xlsx_path, cell_ref):
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        with z.open('xl/worksheets/sheet1.xml') as f:
            content = f.read().decode('utf-8')
            start = content.find(f'r="{cell_ref}"')
            if start == -1:
                return f'{cell_ref}: Celda no encontrada'
            tag_start = content.rfind('<c ', 0, start)
            end = content.find('</c>', start) + 4
            return f'{cell_ref}: {content[tag_start:end]}'

print(check_cell('APU_CON_VAE_CONVERTIDO_v(150856).xlsx', 'J13'))
print(check_cell('APU_CON_VAE_CONVERTIDO_v(150856).xlsx', 'J16'))
print(check_cell('APU_CON_VAE_CONVERTIDO_v(150856).xlsx', 'K13'))
print(check_cell('APU_CON_VAE_CONVERTIDO_v(150856).xlsx', 'L13'))
