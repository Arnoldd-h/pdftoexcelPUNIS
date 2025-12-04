#!/usr/bin/env python3
"""
Compara la validación de datos entre el archivo de referencia y el generado
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re

def analyze_validation(xlsx_path, label):
    """Analiza la validación de datos en un archivo Excel."""
    print(f"\n{'='*60}")
    print(f"Analizando: {label}")
    print(f"Archivo: {xlsx_path}")
    print('='*60)
    
    # Extraer el XLSX
    temp_dir = xlsx_path + '_validate_check'
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
        
        # Buscar la sección de validación
        print("\n1. ESTRUCTURA DE DATA VALIDATION:")
        print("-" * 60)
        
        # Buscar toda la sección dataValidations
        dv_match = re.search(r'<dataValidations[^>]*>(.*?)</dataValidations>', content, re.DOTALL)
        if dv_match:
            dv_section = dv_match.group(0)
            print("✓ Sección <dataValidations> encontrada")
            print(f"Contenido completo:\n{dv_section[:500]}...")
            
            # Extraer atributos del dataValidations
            attrs_match = re.search(r'<dataValidations\s+([^>]*)>', dv_section)
            if attrs_match:
                attrs = attrs_match.group(1)
                print(f"\nAtributos de <dataValidations>: {attrs}")
            
            # Extraer cada dataValidation individual
            individual_dvs = re.findall(r'<dataValidation[^>]*>', content)
            print(f"\n✓ {len(individual_dvs)} reglas de validación encontradas")
            
            for i, dv in enumerate(individual_dvs[:3], 1):  # Primeras 3
                print(f"\nRegla {i}:")
                print(f"  {dv}")
                
                # Extraer atributos específicos
                type_match = re.search(r'type="([^"]*)"', dv)
                formula_match = re.search(r'formula1="([^"]*)"', dv)
                sqref_match = re.search(r'sqref="([^"]*)"', dv)
                
                if type_match:
                    print(f"  - type: {type_match.group(1)}")
                if formula_match:
                    print(f"  - formula1: {formula_match.group(1)}")
                if sqref_match:
                    sqref = sqref_match.group(1)
                    cells = sqref.split()
                    print(f"  - sqref: {len(cells)} celdas")
                    print(f"    Primeras celdas: {' '.join(cells[:5])}")
        else:
            print("❌ NO se encontró sección <dataValidations>")
        
        # Verificar celdas específicas en columna J
        print("\n\n2. ANÁLISIS DE CELDAS INDIVIDUALES (Columna J):")
        print("-" * 60)
        
        # Buscar celdas J12, J13, J16, J17
        test_cells = ['J12', 'J13', 'J16', 'J17']
        for cell_ref in test_cells:
            cell_pattern = f'<c r="{cell_ref}"[^>]*>([^<]*)<'
            match = re.search(cell_pattern, content)
            if match:
                cell_content = match.group(0)
                print(f"\n{cell_ref}:")
                print(f"  {cell_content[:200]}")
            else:
                print(f"\n{cell_ref}: NO ENCONTRADA")
        
        # Leer sharedStrings
        print("\n\n3. SHARED STRINGS (primeros 30):")
        print("-" * 60)
        ss_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
        if os.path.exists(ss_path):
            tree = ET.parse(ss_path)
            root = tree.getroot()
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            for idx, si in enumerate(root.findall('.//s:si', ns)[:30]):
                t = si.find('.//s:t', ns)
                if t is not None:
                    text = t.text
                    if idx in [14, 19, 26]:  # Índices clave para NP/EP/ND
                        print(f"  [{idx}] → '{text}' ⭐")
                    else:
                        print(f"  [{idx}] → '{text}'")
        
    finally:
        # Limpiar
        if os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # Analizar archivo de referencia (PUNIS original)
    ref_file = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE.xlsx"
    
    if os.path.exists(ref_file):
        analyze_validation(ref_file, "ARCHIVO DE REFERENCIA (PUNIS ORIGINAL)")
    else:
        print(f"❌ No se encuentra el archivo de referencia: {ref_file}")
    
    # Analizar archivo generado
    import glob
    files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
    if files:
        gen_file = max(files, key=os.path.getmtime)
        analyze_validation(gen_file, "ARCHIVO GENERADO")
    else:
        print("\n❌ No se encuentra archivo generado")
