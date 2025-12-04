#!/usr/bin/env python3
"""
Comparación exacta celda por celda entre referencia y generado
"""

import zipfile
import os
import re
import shutil
import xml.etree.ElementTree as ET

def compare_cells(ref_path, gen_path):
    """Compara celdas exactas."""
    
    ref_dir = ref_path + '_cmp'
    gen_dir = gen_path + '_cmp'
    
    for d in [ref_dir, gen_dir]:
        if os.path.exists(d):
            shutil.rmtree(d)
        os.makedirs(d)
    
    try:
        with zipfile.ZipFile(ref_path, 'r') as z:
            z.extractall(ref_dir)
        with zipfile.ZipFile(gen_path, 'r') as z:
            z.extractall(gen_dir)
        
        # Leer shared strings de ambos
        def load_shared_strings(base_dir):
            ss_path = os.path.join(base_dir, 'xl', 'sharedStrings.xml')
            strings = []
            if os.path.exists(ss_path):
                with open(ss_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # Usar regex para extraer todos los <t> dentro de <si>
                for si_match in re.finditer(r'<si>(.*?)</si>', content, re.DOTALL):
                    si_content = si_match.group(1)
                    t_match = re.search(r'<t[^>]*>([^<]*)</t>', si_content)
                    if t_match:
                        strings.append(t_match.group(1))
                    else:
                        strings.append('')
            return strings
        
        ref_strings = load_shared_strings(ref_dir)
        gen_strings = load_shared_strings(gen_dir)
        
        print("=" * 80)
        print("COMPARACIÓN CELDA POR CELDA")
        print("=" * 80)
        
        # Leer worksheets
        ref_sheet_path = os.path.join(ref_dir, 'xl', 'worksheets', 'sheet1.xml')
        gen_sheet_path = os.path.join(gen_dir, 'xl', 'worksheets', 'sheet1.xml')
        
        with open(ref_sheet_path, 'r', encoding='utf-8') as f:
            ref_content = f.read()
        with open(gen_sheet_path, 'r', encoding='utf-8') as f:
            gen_content = f.read()
        
        # Comparar filas 13, 17, 18 (datos de EQUIPO y MANO DE OBRA)
        test_rows = [13, 17, 18, 21, 22, 23, 24, 25]  # Filas de datos
        
        for row in test_rows:
            print(f"\n{'='*60}")
            print(f"FILA {row}")
            print(f"{'='*60}")
            
            for col in ['H', 'I', 'J', 'K', 'L']:
                cell_ref = f'{col}{row}'
                
                # Buscar en referencia
                ref_pattern = f'<c r="{cell_ref}"[^>]*>.*?</c>'
                ref_match = re.search(ref_pattern, ref_content, re.DOTALL)
                
                # Buscar en generado
                gen_pattern = f'<c r="{cell_ref}"[^>]*>.*?</c>'
                gen_match = re.search(gen_pattern, gen_content, re.DOTALL)
                
                print(f"\n{cell_ref}:")
                
                if ref_match:
                    ref_cell = ref_match.group(0)
                    # Extraer valor
                    v_match = re.search(r'<v>([^<]*)</v>', ref_cell)
                    t_match = re.search(r't="([^"]*)"', ref_cell)
                    s_match = re.search(r's="([^"]*)"', ref_cell)
                    
                    cell_type = t_match.group(1) if t_match else 'n'
                    style = s_match.group(1) if s_match else '?'
                    
                    if v_match:
                        val = v_match.group(1)
                        if cell_type == 's' and val.isdigit():
                            idx = int(val)
                            text = ref_strings[idx] if idx < len(ref_strings) else '???'
                            print(f"  REF: s={style} t={cell_type} v={val} → '{text}'")
                        else:
                            print(f"  REF: s={style} t={cell_type} v={val}")
                    else:
                        print(f"  REF: s={style} (sin valor)")
                else:
                    print(f"  REF: NO EXISTE")
                
                if gen_match:
                    gen_cell = gen_match.group(0)
                    v_match = re.search(r'<v>([^<]*)</v>', gen_cell)
                    t_match = re.search(r't="([^"]*)"', gen_cell)
                    s_match = re.search(r's="([^"]*)"', gen_cell)
                    
                    cell_type = t_match.group(1) if t_match else 'n'
                    style = s_match.group(1) if s_match else '?'
                    
                    if v_match:
                        val = v_match.group(1)
                        if cell_type == 's' and val.isdigit():
                            idx = int(val)
                            text = gen_strings[idx] if idx < len(gen_strings) else '???'
                            print(f"  GEN: s={style} t={cell_type} v={val} → '{text}'")
                        else:
                            print(f"  GEN: s={style} t={cell_type} v={val}")
                    else:
                        print(f"  GEN: s={style} (sin valor)")
                else:
                    print(f"  GEN: NO EXISTE")
        
    finally:
        for d in [ref_dir, gen_dir]:
            if os.path.exists(d):
                shutil.rmtree(d)

if __name__ == "__main__":
    import glob
    
    ref_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx"
    
    files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
    if files:
        gen_path = max(files, key=os.path.getmtime)
        compare_cells(ref_path, gen_path)
