#!/usr/bin/env python3
"""
Verificar estructura de filas en el archivo de referencia
"""

import zipfile
import os
import re
import shutil

def check_row_structure(ref_path):
    """Verifica la estructura de filas."""
    
    ref_dir = ref_path + '_rows'
    
    if os.path.exists(ref_dir):
        shutil.rmtree(ref_dir)
    os.makedirs(ref_dir)
    
    try:
        with zipfile.ZipFile(ref_path, 'r') as z:
            z.extractall(ref_dir)
        
        sheet_path = os.path.join(ref_dir, 'xl', 'worksheets', 'sheet1.xml')
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Leer shared strings
        ss_path = os.path.join(ref_dir, 'xl', 'sharedStrings.xml')
        strings = []
        with open(ss_path, 'r', encoding='utf-8') as f:
            ss_content = f.read()
        for si_match in re.finditer(r'<si>(.*?)</si>', ss_content, re.DOTALL):
            si_content = si_match.group(1)
            t_match = re.search(r'<t[^>]*>([^<]*)</t>', si_content)
            if t_match:
                strings.append(t_match.group(1))
            else:
                strings.append('')
        
        print("ESTRUCTURA DE FILAS 1-30 EN ARCHIVO DE REFERENCIA:")
        print("=" * 100)
        
        # Buscar todas las filas
        for row_num in range(1, 40):
            row_match = re.search(f'<row r="{row_num}"[^>]*>(.*?)</row>', content, re.DOTALL)
            if row_match:
                row_content = row_match.group(1)
                
                # Extraer celdas de columna A
                a_match = re.search(r'<c r="A' + str(row_num) + r'"[^>]*>.*?</c>', row_content, re.DOTALL)
                a_text = ""
                if a_match:
                    v_match = re.search(r't="s"[^>]*><v>(\d+)</v>', a_match.group(0))
                    if v_match:
                        idx = int(v_match.group(1))
                        a_text = strings[idx][:40] if idx < len(strings) else "?"
                
                # Verificar si tiene celdas VAE (H-L)
                has_vae = any(re.search(f'<c r="{col}{row_num}"', row_content) for col in ['H', 'I', 'J', 'K', 'L'])
                
                vae_info = ""
                if has_vae:
                    for col in ['H', 'I', 'J']:
                        cell_match = re.search(f'<c r="{col}{row_num}"[^>]*>.*?</c>', row_content, re.DOTALL)
                        if cell_match:
                            cell = cell_match.group(0)
                            v_match = re.search(r'<v>([^<]*)</v>', cell)
                            t_match = re.search(r't="s"', cell)
                            if v_match:
                                val = v_match.group(1)
                                if t_match and val.isdigit():
                                    idx = int(val)
                                    text = strings[idx][:15] if idx < len(strings) else "?"
                                    vae_info += f" {col}='{text}'"
                                else:
                                    vae_info += f" {col}={val}"
                
                print(f"Fila {row_num:2d}: A='{a_text[:35]}'{vae_info}")
            else:
                print(f"Fila {row_num:2d}: (vacía o no existe)")
        
    finally:
        if os.path.exists(ref_dir):
            shutil.rmtree(ref_dir)

if __name__ == "__main__":
    ref_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx"
    check_row_structure(ref_path)
