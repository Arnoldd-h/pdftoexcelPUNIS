import zipfile
import re

z = zipfile.ZipFile('APU_CON_VAE_CONVERTIDO_v(151250).xlsx')
sheet = z.open('xl/worksheets/sheet1.xml').read().decode('utf-8')
ss = z.open('xl/sharedStrings.xml').read().decode('utf-8')
strings = re.findall(r'<si><t>([^<]*)</t></si>', ss)

j13_match = re.search(r'r="J13"[^>]*><v>(\d+)</v>', sheet)
if j13_match:
    idx = int(j13_match.group(1))
    print(f'J13 apunta a índice: {idx}')
    print(f'Ese índice contiene: "{strings[idx]}"')
else:
    print('J13 no encontrado')

# Ver también J16
j16_match = re.search(r'r="J16"[^>]*><v>(\d+)</v>', sheet)
if j16_match:
    idx = int(j16_match.group(1))
    print(f'\nJ16 apunta a índice: {idx}')
    print(f'Ese índice contiene: "{strings[idx]}"')
