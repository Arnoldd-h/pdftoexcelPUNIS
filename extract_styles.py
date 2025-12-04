#!/usr/bin/env python3
"""
Extrae los estilos del archivo de referencia para usarlos en el generado
"""

import zipfile
import os
import re
import shutil

def extract_styles(ref_path):
    """Extrae los estilos del archivo de referencia."""
    
    ref_dir = ref_path + '_styles'
    
    if os.path.exists(ref_dir):
        shutil.rmtree(ref_dir)
    os.makedirs(ref_dir)
    
    try:
        with zipfile.ZipFile(ref_path, 'r') as z:
            z.extractall(ref_dir)
        
        styles_path = os.path.join(ref_dir, 'xl', 'styles.xml')
        with open(styles_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print("ESTILOS RELEVANTES PARA VAE (28-31):\n")
        
        # Buscar cellXfs
        xfs_match = re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', content, re.DOTALL)
        if xfs_match:
            xfs_content = xfs_match.group(1)
            # Encontrar cada xf
            xf_items = re.findall(r'<xf [^>]*/>', xfs_content)
            
            for i in [28, 29, 30, 31]:
                if i < len(xf_items):
                    print(f"Estilo [{i}]: {xf_items[i]}")
        
        print("\n\nFORMATOS DE NÚMERO (numFmts):")
        numfmts_match = re.search(r'<numFmts[^>]*>(.*?)</numFmts>', content, re.DOTALL)
        if numfmts_match:
            print(numfmts_match.group(0)[:1000])
        
        print("\n\nBORDES:")
        borders_match = re.search(r'<borders[^>]*>(.*?)</borders>', content, re.DOTALL)
        if borders_match:
            borders_content = borders_match.group(1)
            border_items = re.findall(r'<border[^>]*>.*?</border>', borders_content, re.DOTALL)
            for i, border in enumerate(border_items[:10]):
                print(f"Borde [{i}]: {border}")
        
    finally:
        if os.path.exists(ref_dir):
            shutil.rmtree(ref_dir)

if __name__ == "__main__":
    ref_path = r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\ANALISIS_PU_VAE_PUNIS_SS_HH__LAGO_SAN_PEDRO.xlsx"
    extract_styles(ref_path)
