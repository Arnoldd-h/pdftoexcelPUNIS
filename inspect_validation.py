import zipfile
import xml.etree.ElementTree as ET

def inspect_data_validation(xlsx_path):
    """Inspecciona las validaciones de datos en el archivo."""
    print(f"\n=== Analizando: {xlsx_path} ===")
    
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        # Leer el worksheet
        with z.open('xl/worksheets/sheet1.xml') as f:
            content = f.read().decode('utf-8')
            
            # Buscar dataValidations
            if '<dataValidations' in content:
                print("\n✓ Encontradas validaciones de datos:")
                start = content.find('<dataValidations')
                end = content.find('</dataValidations>') + len('</dataValidations>')
                validations = content[start:end]
                print(validations[:2000])  # Primeros 2000 caracteres
            else:
                print("\n✗ No se encontraron validaciones de datos")
            
            # Buscar si hay definedNames (rangos nombrados)
            print("\n--- Buscando rangos nombrados en workbook.xml ---")
        
        try:
            with z.open('xl/workbook.xml') as f:
                wb_content = f.read().decode('utf-8')
                if '<definedNames>' in wb_content:
                    start = wb_content.find('<definedNames>')
                    end = wb_content.find('</definedNames>') + len('</definedNames>')
                    names = wb_content[start:end]
                    print(names[:1000])
                else:
                    print("No hay rangos nombrados")
        except:
            print("No se pudo leer workbook.xml")

# Analizar el archivo de referencia
inspect_data_validation('ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx')

# Analizar nuestro archivo generado
inspect_data_validation('APU_CON_VAE_CONVERTIDO_v(134007).xlsx')
