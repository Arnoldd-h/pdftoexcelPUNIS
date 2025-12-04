import zipfile

def get_shared_string(xlsx_path, index):
    """Obtiene el string compartido por índice."""
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        with z.open('xl/sharedStrings.xml') as f:
            content = f.read().decode('utf-8')
            
            # Buscar todos los <si> tags
            import re
            si_tags = re.findall(r'<si><t>(.*?)</t></si>', content)
            
            if index < len(si_tags):
                return si_tags[index]
            else:
                return f"Índice {index} fuera de rango (total: {len(si_tags)})"

print("=== Archivo de referencia PUNIS ===")
print(f"Índice 21 (J13): '{get_shared_string('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 21)}'")
print(f"Índice 16 (J16): '{get_shared_string('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx', 16)}'")

print("\n=== Nuestro archivo generado ===")
print(f"Índice 19 (J13): '{get_shared_string('APU_CON_VAE_CONVERTIDO_v(134007).xlsx', 19)}'")
print(f"Índice 14 (J16): '{get_shared_string('APU_CON_VAE_CONVERTIDO_v(134007).xlsx', 14)}'")
