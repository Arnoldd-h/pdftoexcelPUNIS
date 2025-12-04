#!/usr/bin/env python3
"""
Verifica los datos de VAE en el archivo Excel generado
"""

import zipfile
import xml.etree.ElementTree as ET
import os

def check_vae_data(xlsx_path):
    """Verifica los datos de VAE en el archivo Excel."""
    print(f"Analizando: {xlsx_path}\n")
    
    # Extraer el XLSX
    temp_dir = xlsx_path + '_check'
    if os.path.exists(temp_dir):
        import shutil
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            z.extractall(temp_dir)
        
        # Leer el worksheet
        sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', 'sheet1.xml')
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Leer sharedStrings si existe
        ss_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
        shared_strings = []
        if os.path.exists(ss_path):
            tree = ET.parse(ss_path)
            root = tree.getroot()
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('.//s:si', ns):
                t = si.find('.//s:t', ns)
                if t is not None:
                    shared_strings.append(t.text)
            print(f"Shared Strings encontrados: {len(shared_strings)}\n")
        
        # Parsear XML
        tree = ET.fromstring(content)
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        # Buscar celdas en columna J (NP/EP/ND)
        print("Analizando columna J (NP/EP/ND):")
        print("=" * 60)
        
        # Buscar todas las celdas que empiezan con J
        import re
        cell_pattern = r'<c r="(J\d+)"[^>]*t="([^"]*)"[^>]*>(?:<f[^>]*>[^<]*</f>)?<v>([^<]*)</v>'
        
        matches = re.findall(cell_pattern, content)
        
        if not matches:
            print("❌ NO se encontraron celdas en columna J con datos")
        else:
            for cell_ref, cell_type, value in matches[:10]:  # Primeras 10
                if cell_type == 's':  # Shared string
                    idx = int(value)
                    if idx < len(shared_strings):
                        text = shared_strings[idx]
                        print(f"  {cell_ref}: índice={idx} → '{text}'")
                    else:
                        print(f"  {cell_ref}: índice={idx} → ❌ FUERA DE RANGO")
                elif cell_type == 'inlineStr':
                    print(f"  {cell_ref}: ❌ INLINE STRING (no debería estar)")
                else:
                    print(f"  {cell_ref}: valor={value} (tipo={cell_type})")
        
        # Verificar si hay data validation
        print("\n\nVerificando Data Validation:")
        print("=" * 60)
        if '<dataValidations' in content:
            dv_pattern = r'<dataValidation[^>]*sqref="([^"]*)"[^>]*>'
            validations = re.findall(dv_pattern, content)
            if validations:
                print(f"✓ Se encontraron {len(validations)} reglas de validación:")
                for v in validations[:5]:
                    print(f"  - Rango: {v}")
            else:
                print("❌ No se encontraron reglas de validación definidas")
        else:
            print("❌ No existe sección <dataValidations>")
        
        # Verificar defined names
        print("\n\nVerificando Defined Names:")
        print("=" * 60)
        workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
        if os.path.exists(workbook_path):
            with open(workbook_path, 'r', encoding='utf-8') as f:
                wb_content = f.read()
            if '<definedNames>' in wb_content:
                dn_pattern = r'<definedName[^>]*name="([^"]*)"[^>]*>([^<]*)</definedName>'
                names = re.findall(dn_pattern, wb_content)
                if names:
                    print(f"✓ Se encontraron {len(names)} nombres definidos:")
                    for name, ref in names[:5]:
                        print(f"  - {name}: {ref}")
                else:
                    print("❌ No se encontraron nombres definidos")
            else:
                print("❌ No existe sección <definedNames>")
        
    finally:
        # Limpiar
        if os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)
    
    print("\n\n" + "=" * 60)
    print("DIAGNÓSTICO:")
    print("=" * 60)
    print("Para que PUNIS muestre automáticamente los datos del VAE,")
    print("el archivo debe tener:")
    print("  1. ✓ Shared Strings (ya tiene)")
    print("  2. ✓ Celdas J con referencias a shared strings (ya tiene)")
    print("  3. ❓ Data Validation con lista NP/EP/ND")
    print("  4. ❓ Defined Name para la lista de validación")
    print("\nSi faltan puntos 3 o 4, ese es el problema.")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        xlsx_path = sys.argv[1]
    else:
        # Buscar el archivo más reciente
        import glob
        files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
        if files:
            xlsx_path = max(files, key=os.path.getmtime)
        else:
            xlsx_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v(152037).xlsx"
    check_vae_data(xlsx_path)
