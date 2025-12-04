import zipfile

def get_cell_value_and_type(xlsx_path, cell_ref):
    """Obtiene el valor y tipo de una celda específica."""
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
            cell_xml = content[tag_start:end]
            
            return cell_xml

print("=== Columna J (NP/EP/ND) en archivo de referencia ===")
print("\nJ13 (Equipo):", get_cell_value_and_type('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'J13'))
print("\nJ15 (Mano de Obra línea 1):", get_cell_value_and_type('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'J15'))
print("\nJ16 (Mano de Obra línea 2):", get_cell_value_and_type('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 'J16'))

print("\n\n=== Columna J en nuestro archivo generado ===")
print("\nJ13 (Equipo):", get_cell_value_and_type('APU_CON_VAE_CONVERTIDO_v(134007).xlsx', 'J13'))
print("\nJ15 (Mano de Obra línea 1):", get_cell_value_and_type('APU_CON_VAE_CONVERTIDO_v(134007).xlsx', 'J15'))
print("\nJ16 (Mano de Obra línea 2):", get_cell_value_and_type('APU_CON_VAE_CONVERTIDO_v(134007).xlsx', 'J16'))
