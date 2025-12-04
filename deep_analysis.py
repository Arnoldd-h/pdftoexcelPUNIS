#!/usr/bin/env python3
"""
Análisis profundo del archivo de referencia que SÍ funciona en PUNIS
"""

import zipfile
import os
import re
import shutil

def deep_analysis(ref_path):
    """Análisis profundo del archivo de referencia."""
    
    ref_dir = ref_path + '_deep'
    
    if os.path.exists(ref_dir):
        shutil.rmtree(ref_dir)
    os.makedirs(ref_dir)
    
    try:
        with zipfile.ZipFile(ref_path, 'r') as z:
            z.extractall(ref_dir)
        
        print("=" * 80)
        print("ANÁLISIS PROFUNDO DEL ARCHIVO DE REFERENCIA")
        print("=" * 80)
        
        # Listar todos los archivos
        print("\n1. ESTRUCTURA DE ARCHIVOS:")
        print("-" * 60)
        for root, dirs, files in os.walk(ref_dir):
            level = root.replace(ref_dir, '').count(os.sep)
            indent = ' ' * 2 * level
            print(f'{indent}{os.path.basename(root)}/')
            subindent = ' ' * 2 * (level + 1)
            for file in files:
                filepath = os.path.join(root, file)
                size = os.path.getsize(filepath)
                print(f'{subindent}{file} ({size} bytes)')
        
        # Leer sheet1.xml
        sheet_path = os.path.join(ref_dir, 'xl', 'worksheets', 'sheet1.xml')
        with open(sheet_path, 'r', encoding='utf-8') as f:
            sheet_content = f.read()
        
        # Buscar definedNames en workbook.xml
        print("\n\n2. WORKBOOK.XML - DEFINED NAMES:")
        print("-" * 60)
        wb_path = os.path.join(ref_dir, 'xl', 'workbook.xml')
        with open(wb_path, 'r', encoding='utf-8') as f:
            wb_content = f.read()
        
        if '<definedNames>' in wb_content:
            dn_match = re.search(r'<definedNames>(.*?)</definedNames>', wb_content, re.DOTALL)
            if dn_match:
                print(dn_match.group(0)[:1000])
        else:
            print("NO hay definedNames")
        
        # Buscar estilos
        print("\n\n3. STYLES.XML - ESTILOS USADOS:")
        print("-" * 60)
        styles_path = os.path.join(ref_dir, 'xl', 'styles.xml')
        with open(styles_path, 'r', encoding='utf-8') as f:
            styles_content = f.read()
        
        # Buscar cellXfs (formatos de celda)
        xfs_match = re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles_content, re.DOTALL)
        if xfs_match:
            xfs = xfs_match.group(0)
            xf_count = len(re.findall(r'<xf ', xfs))
            print(f"Total de formatos de celda (xf): {xf_count}")
            
            # Mostrar estilos 29 y 30 (usados en columnas I y J)
            xf_items = re.findall(r'<xf [^>]*/>', xfs)
            for i, xf in enumerate(xf_items[:35]):
                if i in [29, 30]:
                    print(f"  [{i}] {xf} ⭐ (usado en VAE)")
                else:
                    print(f"  [{i}] {xf}")
        
        # Analizar filas de datos VAE específicas
        print("\n\n4. FILAS DE DATOS VAE EN DETALLE:")
        print("-" * 60)
        
        # Buscar fila 13 completa (primera fila de datos de EQUIPO)
        row13_match = re.search(r'<row r="13"[^>]*>(.*?)</row>', sheet_content, re.DOTALL)
        if row13_match:
            row13 = row13_match.group(0)
            print("FILA 13 (Primera fila de datos EQUIPO):")
            # Buscar celdas de H a L
            for col in ['H', 'I', 'J', 'K', 'L']:
                cell_match = re.search(f'<c r="{col}13"[^>]*>.*?</c>', row13, re.DOTALL)
                if cell_match:
                    print(f"  {col}13: {cell_match.group(0)}")
                else:
                    # Buscar celda vacía o sin contenido
                    cell_match2 = re.search(f'<c r="{col}13"[^/]*/>', row13)
                    if cell_match2:
                        print(f"  {col}13: {cell_match2.group(0)} (vacía)")
                    else:
                        print(f"  {col}13: NO EXISTE")
        
        # Buscar fila 17 completa (primera fila de MANO DE OBRA datos)
        row17_match = re.search(r'<row r="17"[^>]*>(.*?)</row>', sheet_content, re.DOTALL)
        if row17_match:
            row17 = row17_match.group(0)
            print("\nFILA 17 (Primera fila de datos MANO DE OBRA):")
            for col in ['H', 'I', 'J', 'K', 'L']:
                cell_match = re.search(f'<c r="{col}17"[^>]*>.*?</c>', row17, re.DOTALL)
                if cell_match:
                    print(f"  {col}17: {cell_match.group(0)}")
                else:
                    cell_match2 = re.search(f'<c r="{col}17"[^/]*/>', row17)
                    if cell_match2:
                        print(f"  {col}17: {cell_match2.group(0)} (vacía)")
                    else:
                        print(f"  {col}17: NO EXISTE")
        
        # Leer shared strings y mostrar índices relevantes
        print("\n\n5. SHARED STRINGS - ÍNDICES RELEVANTES:")
        print("-" * 60)
        ss_path = os.path.join(ref_dir, 'xl', 'sharedStrings.xml')
        with open(ss_path, 'r', encoding='utf-8') as f:
            ss_content = f.read()
        
        strings = re.findall(r'<si><t>([^<]*)</t></si>', ss_content)
        indices_importantes = [17, 18, 19, 20, 21, 22, 27, 28, 29]
        for i in indices_importantes:
            if i < len(strings):
                print(f"  [{i}] = '{strings[i]}'")
        
    finally:
        if os.path.exists(ref_dir):
            shutil.rmtree(ref_dir)

if __name__ == "__main__":
    ref_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx"
    deep_analysis(ref_path)
