#!/usr/bin/env python3
"""
Verifica que las celdas NP/EP/ND sean inline strings
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re

def check_inline_vae(xlsx_path):
    """Verifica que NP/EP/ND sean inline strings."""
    print(f"Analizando: {xlsx_path}\n")
    
    # Extraer el XLSX
    temp_dir = xlsx_path + '_inline_check'
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
        
        print("Verificando celdas de columna J (NP/EP/ND):")
        print("=" * 60)
        
        # Buscar celdas J con t="inlineStr"
        inline_pattern = r'<c r="(J\d+)"[^>]*t="inlineStr"[^>]*><is><t>([^<]*)</t></is></c>'
        inline_matches = re.findall(inline_pattern, content)
        
        print(f"\n✓ {len(inline_matches)} celdas J con inline strings")
        print("\nPrimeras 20 celdas:")
        for cell_ref, text in inline_matches[:20]:
            if text in ['NP', 'EP', 'ND']:
                print(f"  {cell_ref}: '{text}' ⭐")
            else:
                print(f"  {cell_ref}: '{text}'")
        
        # Buscar celdas J con t="s" (shared string) - NO deberían tener NP/EP/ND
        shared_pattern = r'<c r="(J\d+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>'
        shared_matches = re.findall(shared_pattern, content)
        
        print(f"\n\n✓ {len(shared_matches)} celdas J con shared strings")
        
        # Verificar si alguna shared string es NP/EP/ND (no debería)
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
        
        print("\nPrimeras 10 celdas J con shared string (verificando contenido):")
        for cell_ref, idx in shared_matches[:10]:
            idx_int = int(idx)
            if idx_int < len(shared_strings):
                text = shared_strings[idx_int]
                if text in ['NP', 'EP', 'ND']:
                    print(f"  {cell_ref} → índice {idx} → '{text}' ❌ ERROR: debería ser inline!")
                else:
                    print(f"  {cell_ref} → índice {idx} → '{text[:30]}...'")
        
        # Verificar data validation
        print("\n\nVerificando Data Validation:")
        print("=" * 60)
        if '<dataValidations' in content:
            dv_match = re.search(r'<dataValidation[^>]*sqref="([^"]*)"[^>]*>', content)
            if dv_match:
                sqref = dv_match.group(1)
                cells = sqref.split()
                print(f"✓ Validación aplicada a {len(cells)} celdas")
                print(f"  Primeras celdas: {' '.join(cells[:10])}")
                
                # Verificar atributos
                attr_match = re.search(r'<dataValidation\s+([^>]*)', content)
                if attr_match:
                    attrs = attr_match.group(1)
                    if 'showDropDown="0"' in attrs:
                        print("  ⚠ showDropDown=0 (dropdown oculto)")
                    if 'allowBlank="0"' in attrs:
                        print("  ✓ allowBlank=0 (no permite vacío)")
                    if 'type="list"' in attrs:
                        print("  ✓ type=list")
        
        print("\n\n" + "=" * 60)
        print("RESUMEN:")
        print("=" * 60)
        vae_inline = [text for _, text in inline_matches if text in ['NP', 'EP', 'ND']]
        print(f"✓ {len(vae_inline)} celdas con NP/EP/ND como inline strings")
        print(f"  - NP: {vae_inline.count('NP')}")
        print(f"  - EP: {vae_inline.count('EP')}")
        print(f"  - ND: {vae_inline.count('ND')}")
        print("\n✅ Si todos los NP/EP/ND son inline, PUNIS debería reconocerlos automáticamente")
        
    finally:
        # Limpiar
        if os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    import glob
    files = glob.glob(r"c:\Users\User\Downloads\Nueva carpeta (2)\Nueva carpeta (2)\pdftoexcelPUNIS-main\APU_CON_VAE_CONVERTIDO_v*.xlsx")
    if files:
        latest = max(files, key=os.path.getmtime)
        check_inline_vae(latest)
    else:
        print("No se encontró archivo generado")
