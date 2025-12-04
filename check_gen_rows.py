#!/usr/bin/env python3
"""
Verificar estructura de filas en el archivo generado
"""

import zipfile
import os
import re
import shutil
import glob

def check_row_structure(file_path, label):
    """Verifica la estructura de filas."""
    
    temp_dir = file_path + '_rows_check'
    
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(temp_dir)
        
        sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', 'sheet1.xml')
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Leer shared strings
        ss_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
        strings = []
        if os.path.exists(ss_path):
            with open(ss_path, 'r', encoding='utf-8') as f:
                ss_content = f.read()
            for si_match in re.finditer(r'<si>(.*?)</si>', ss_content, re.DOTALL):
                si_content = si_match.group(1)
                t_match = re.search(r'<t[^>]*>([^<]*)</t>', si_content)
                if t_match:
                    strings.append(t_match.group(1))
                else:
                    strings.append('')
        
        print(f"\n{'='*100}")
        print(f"ESTRUCTURA DE FILAS EN: {label}")
        print(f"{'='*100}")
        
        for row_num in range(1, 40):
            row_match = re.search(f'<row r="{row_num}"[^>]*>(.*?)</row>', content, re.DOTALL)
            if row_match:
                row_content = row_match.group(1)
                
                # Extraer celda A
                a_match = re.search(r'<c r="A' + str(row_num) + r'"[^>]*>.*?</c>', row_content, re.DOTALL)
                a_text = ""
                if a_match:
                    # Buscar shared string
                    v_match = re.search(r't="s"[^>]*><v>(\d+)</v>', a_match.group(0))
                    if v_match:
                        idx = int(v_match.group(1))
                        a_text = strings[idx][:40] if idx < len(strings) else "?"
                    else:
                        # Buscar inline string
                        inline_match = re.search(r'<t>([^<]*)</t>', a_match.group(0))
                        if inline_match:
                            a_text = inline_match.group(1)[:40]
                
                # Verificar celdas VAE
                vae_info = ""
                for col in ['H', 'I', 'J']:
                    cell_match = re.search(f'<c r="{col}{row_num}"[^>]*>.*?</c>', row_content, re.DOTALL)
                    if cell_match:
                        cell = cell_match.group(0)
                        v_match = re.search(r'<v>([^<]*)</v>', cell)
                        t_shared = re.search(r't="s"', cell)
                        if v_match:
                            val = v_match.group(1)
                            if t_shared and val.isdigit():
                                idx = int(val)
                                text = strings[idx][:15] if idx < len(strings) else "?"
                                vae_info += f" {col}='{text}'"
                            else:
                                vae_info += f" {col}={val}"
                        else:
                            # Inline string
                            inline_match = re.search(r'<t>([^<]*)</t>', cell)
                            if inline_match:
                                vae_info += f" {col}='{inline_match.group(1)[:15]}'"
                
                print(f"Fila {row_num:2d}: A='{a_text[:35]}'{vae_info}")
            else:
                print(f"Fila {row_num:2d}: (vacía o no existe)")
        
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # Archivo generado
    files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
    if files:
        gen_path = max(files, key=os.path.getmtime)
        check_row_structure(gen_path, "ARCHIVO GENERADO")
