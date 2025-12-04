"""
Script para convertir inline strings a shared strings en un archivo XLSX.
Esto es necesario para compatibilidad con PUNIS.
"""
import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
import os
import shutil
import re

def convert_to_shared_strings(input_path, output_path=None):
    """Convierte un archivo XLSX de inline strings a shared strings."""
    
    if output_path is None:
        output_path = input_path
    
    # Crear directorio temporal
    temp_dir = input_path + '_temp'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    # Extraer el XLSX
    with zipfile.ZipFile(input_path, 'r') as z:
        z.extractall(temp_dir)
    
    # Parsear el worksheet
    sheet_path = os.path.join(temp_dir, 'xl', 'worksheets', 'sheet1.xml')
    
    # Registrar namespace
    ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    
    # Leer y parsear
    with open(sheet_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Recopilar todos los inline strings y crear tabla de shared strings
    shared_strings = []
    string_map = {}  # mapa de string -> índice
    
    # Buscar todos los inline strings
    pattern = r't="inlineStr"[^>]*><is><t>([^<]*)</t></is>'
    
    for match in re.finditer(pattern, content):
        text = match.group(1)
        if text not in string_map:
            string_map[text] = len(shared_strings)
            shared_strings.append(text)
    
    print(f"Encontrados {len(shared_strings)} strings únicos")
    
    # Reemplazar inline strings con referencias a shared strings
    def replace_inline(match):
        text = match.group(1)
        idx = string_map[text]
        return f't="s"><v>{idx}</v>'
    
    new_content = re.sub(pattern, replace_inline, content)
    
    # Escribir el worksheet modificado
    with open(sheet_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    # Crear el archivo sharedStrings.xml
    shared_strings_path = os.path.join(temp_dir, 'xl', 'sharedStrings.xml')
    
    ss_content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    ss_content += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    ss_content += f'count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
    
    for s in shared_strings:
        # Escapar caracteres especiales
        s_escaped = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        ss_content += f'<si><t>{s_escaped}</t></si>'
    
    ss_content += '</sst>'
    
    with open(shared_strings_path, 'w', encoding='utf-8') as f:
        f.write(ss_content)
    
    # Actualizar [Content_Types].xml para incluir sharedStrings
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    with open(content_types_path, 'r', encoding='utf-8') as f:
        ct_content = f.read()
    
    if 'sharedStrings' not in ct_content:
        # Agregar el override para sharedStrings
        insert_pos = ct_content.find('</Types>')
        override = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        ct_content = ct_content[:insert_pos] + override + ct_content[insert_pos:]
        
        with open(content_types_path, 'w', encoding='utf-8') as f:
            f.write(ct_content)
    
    # Actualizar xl/_rels/workbook.xml.rels para incluir relación con sharedStrings
    rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels_content = f.read()
    
    if 'sharedStrings' not in rels_content:
        # Encontrar el próximo rId
        rids = re.findall(r'rId(\d+)', rels_content)
        next_rid = max(int(r) for r in rids) + 1 if rids else 1
        
        insert_pos = rels_content.find('</Relationships>')
        rel = f'<Relationship Id="rId{next_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        rels_content = rels_content[:insert_pos] + rel + rels_content[insert_pos:]
        
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)
    
    # Crear el nuevo XLSX
    if os.path.exists(output_path):
        os.remove(output_path)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                z.write(file_path, arcname)
    
    # Limpiar
    shutil.rmtree(temp_dir)
    
    print(f"Archivo convertido guardado en: {output_path}")
    return output_path

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        convert_to_shared_strings(input_file, output_file)
    else:
        # Convertir el archivo generado
        convert_to_shared_strings('APU_CON_VAE_CONVERTIDO.xlsx')
