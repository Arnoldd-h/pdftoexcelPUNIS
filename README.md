# Convertidor de APU - PDF a Excel

Este script automatiza la conversión de archivos PDF de **Análisis de Precios Unitarios (APU) con VAE** al formato Excel estandarizado.

## Estructura del Formato

El script reconoce y extrae:

### Información de encabezado:
- Nombre del profesional responsable
- Proyecto y ubicación
- Número de hoja

### Por cada rubro:
- Número de rubro, unidad y detalle
- **Equipos**: Descripción, costo, peso relativo, CPC, NP/EP/ND, VAE
- **Mano de Obra**: Descripción, categoría (EO C1, EO D2, etc.), cantidad, jornal/hr, costo hora, rendimiento, costo, peso relativo, CPC, NP/EP/ND, VAE
- **Materiales**: Descripción, unidad, cantidad, precio unitario, costo, peso relativo, CPC, NP/EP/ND, VAE
- **Transporte**: Descripción, unidad, cantidad, tarifa, costo, peso relativo, CPC, NP/EP/ND, VAE
- Subtotales (M, N, O, P)
- Total costo directo
- Indirectos y utilidad
- Costo total del rubro
- Valor unitario
- Texto del valor en letras
- Fecha

## Uso

### Opción 1: Línea de comandos
```bash
python pdf_to_excel_apu.py archivo.pdf [archivo_salida.xlsx]
```

### Opción 2: Arrastrar y soltar
1. Arrastra tu archivo PDF sobre `convertir_apu.bat`
2. El archivo Excel se generará en la misma carpeta

### Opción 3: Interfaz gráfica
1. Ejecuta `Convertidor_APU.bat`
2. Selecciona el archivo PDF
3. Haz clic en "Convertir PDF a Excel"

## Requisitos

- Python 3.8 o superior
- Librerías: pdfplumber, openpyxl, pandas

Las librerías se instalan automáticamente en el entorno virtual `.venv`.

## Instalación manual de dependencias

```bash
pip install pdfplumber openpyxl pandas
```

## Formatos soportados

El script está diseñado para PDFs con el formato estándar de APU ecuatoriano que incluye:
- Secciones de EQUIPO, MANO DE OBRA, MATERIALES, TRANSPORTE
- Cálculos de VAE (Valor Agregado Ecuatoriano)
- Códigos CPC
- Clasificación NP/EP/ND

## Notas

- El archivo Excel de salida tendrá el sufijo `_CONVERTIDO.xlsx`
- Si el PDF tiene múltiples rubros, todos se incluirán en el mismo archivo Excel
- El formato de salida replica exactamente la estructura del Excel de ejemplo

## Autor

Herramienta de automatización para conversión de APUs.

## Versión

1.0 - Noviembre 2025
