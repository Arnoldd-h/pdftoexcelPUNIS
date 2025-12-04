import zipfile

def get_cell_xml(xlsx_path, cell_ref):
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        with z.open('xl/worksheets/sheet1.xml') as f:
            content = f.read().decode('utf-8')
            # Buscar la celda específica
            start = content.find(f'r="{cell_ref}"')
            if start == -1:
                return 'Celda no encontrada'
            # Retroceder para encontrar el inicio del tag <c
            tag_start = content.rfind('<c ', 0, start)
            # Buscar el cierre
            end = content.find('</c>', start) + 4
            return content[tag_start:end]

# Comparar celda B12 (código 514704408)
print('=== CELDA B12 (código 514704408) ===')
print('PUNIS Original:')
print(get_cell_xml('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'B12'))
print()
print('Generado:')
print(get_cell_xml('APU_CON_VAE_CONVERTIDO.xlsx', 'B12'))

# Comparar celda I13 (CPC del equipo)
print()
print('=== CELDA I13 (CPC elemento) ===')
print('PUNIS Original:')
print(get_cell_xml('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'I13'))
print()
print('Generado:')
print(get_cell_xml('APU_CON_VAE_CONVERTIDO.xlsx', 'I13'))

# Comparar celda L13 (VAE elemento)
print()
print('=== CELDA L13 (VAE elemento) ===')
print('PUNIS Original:')
print(get_cell_xml('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'L13'))
print()
print('Generado:')
print(get_cell_xml('APU_CON_VAE_CONVERTIDO.xlsx', 'L13'))
