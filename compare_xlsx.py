#!/usr/bin/env python3
"""
Compara el archivo de referencia (que funciona en PUNIS) con el generado
para encontrar las diferencias en estructura XML
"""

import zipfile
import os
import re

def extract_and_compare(ref_path, gen_path):
    """Compara dos archivos XLSX a nivel XML."""
    
    ref_dir = ref_path + '_compare'
    gen_dir = gen_path + '_compare'
    
    # Limpiar directorios
    import shutil
    for d in [ref_dir, gen_dir]:
        if os.path.exists(d):
            shutil.rmtree(d)
        os.makedirs(d)
    
    try:
        # Extraer ambos archivos
        with zipfile.ZipFile(ref_path, 'r') as z:
            z.extractall(ref_dir)
        with zipfile.ZipFile(gen_path, 'r') as z:
            z.extractall(gen_dir)
        
        print("=" * 80)
        print("COMPARACIÓN DE ARCHIVOS XLSX")
        print("=" * 80)
        print(f"\nREFERENCIA: {os.path.basename(ref_path)}")
        print(f"GENERADO:   {os.path.basename(gen_path)}")
        
        # Comparar sheet1.xml
        print("\n\n" + "=" * 80)
        print("1. COMPARACIÓN DE CELDAS EN SHEET1.XML")
        print("=" * 80)
        
        ref_sheet = os.path.join(ref_dir, 'xl', 'worksheets', 'sheet1.xml')
        gen_sheet = os.path.join(gen_dir, 'xl', 'worksheets', 'sheet1.xml')
        
        with open(ref_sheet, 'r', encoding='utf-8') as f:
            ref_content = f.read()
        with open(gen_sheet, 'r', encoding='utf-8') as f:
            gen_content = f.read()
        
        # Buscar celdas en columna I y J para filas específicas
        test_cells = ['I13', 'J13', 'I17', 'J17', 'I18', 'J18']
        
        for cell_ref in test_cells:
            print(f"\n--- Celda {cell_ref} ---")
            
            # Buscar en referencia
            ref_pattern = f'<c r="{cell_ref}"[^>]*>.*?</c>'
            ref_match = re.search(ref_pattern, ref_content, re.DOTALL)
            
            # Buscar en generado
            gen_pattern = f'<c r="{cell_ref}"[^>]*>.*?</c>'
            gen_match = re.search(gen_pattern, gen_content, re.DOTALL)
            
            print(f"REF: {ref_match.group(0)[:150] if ref_match else 'NO ENCONTRADA'}...")
            print(f"GEN: {gen_match.group(0)[:150] if gen_match else 'NO ENCONTRADA'}...")
        
        # Comparar sharedStrings.xml
        print("\n\n" + "=" * 80)
        print("2. COMPARACIÓN DE SHARED STRINGS")
        print("=" * 80)
        
        ref_ss = os.path.join(ref_dir, 'xl', 'sharedStrings.xml')
        gen_ss = os.path.join(gen_dir, 'xl', 'sharedStrings.xml')
        
        if os.path.exists(ref_ss):
            with open(ref_ss, 'r', encoding='utf-8') as f:
                ref_ss_content = f.read()
            ref_count = len(re.findall(r'<si>', ref_ss_content))
            print(f"REF: {ref_count} shared strings")
            
            # Mostrar primeros 20
            ref_strings = re.findall(r'<si><t>([^<]*)</t></si>', ref_ss_content)[:20]
            print("Primeros 20:")
            for i, s in enumerate(ref_strings):
                print(f"  [{i}] '{s[:40]}...'")
        else:
            print("REF: NO TIENE sharedStrings.xml")
        
        if os.path.exists(gen_ss):
            with open(gen_ss, 'r', encoding='utf-8') as f:
                gen_ss_content = f.read()
            gen_count = len(re.findall(r'<si>', gen_ss_content))
            print(f"\nGEN: {gen_count} shared strings")
            
            # Mostrar primeros 20
            gen_strings = re.findall(r'<si><t>([^<]*)</t></si>', gen_ss_content)[:20]
            print("Primeros 20:")
            for i, s in enumerate(gen_strings):
                print(f"  [{i}] '{s[:40]}...'")
        else:
            print("GEN: NO TIENE sharedStrings.xml")
        
        # Buscar diferencias en estructura de celdas
        print("\n\n" + "=" * 80)
        print("3. ANÁLISIS DE ESTRUCTURA DE CELDAS VAE")
        print("=" * 80)
        
        # En el archivo de referencia, ¿cómo están las celdas I y J?
        # Buscar patrón de celdas con t="s" vs t="inlineStr"
        ref_inline_i = len(re.findall(r'<c r="I\d+"[^>]*t="inlineStr"', ref_content))
        ref_shared_i = len(re.findall(r'<c r="I\d+"[^>]*t="s"', ref_content))
        ref_inline_j = len(re.findall(r'<c r="J\d+"[^>]*t="inlineStr"', ref_content))
        ref_shared_j = len(re.findall(r'<c r="J\d+"[^>]*t="s"', ref_content))
        
        gen_inline_i = len(re.findall(r'<c r="I\d+"[^>]*t="inlineStr"', gen_content))
        gen_shared_i = len(re.findall(r'<c r="I\d+"[^>]*t="s"', gen_content))
        gen_inline_j = len(re.findall(r'<c r="J\d+"[^>]*t="inlineStr"', gen_content))
        gen_shared_j = len(re.findall(r'<c r="J\d+"[^>]*t="s"', gen_content))
        
        print("\nColumna I (CPC):")
        print(f"  REF: {ref_inline_i} inline, {ref_shared_i} shared")
        print(f"  GEN: {gen_inline_i} inline, {gen_shared_i} shared")
        
        print("\nColumna J (NP/EP/ND):")
        print(f"  REF: {ref_inline_j} inline, {ref_shared_j} shared")
        print(f"  GEN: {gen_inline_j} inline, {gen_shared_j} shared")
        
    finally:
        # Limpiar
        for d in [ref_dir, gen_dir]:
            if os.path.exists(d):
                shutil.rmtree(d)

if __name__ == "__main__":
    import glob
    
    ref_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx"
    
    # Buscar archivo generado más reciente
    files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
    if files:
        gen_path = max(files, key=os.path.getmtime)
        extract_and_compare(ref_path, gen_path)
    else:
        print("No se encontró archivo generado")
