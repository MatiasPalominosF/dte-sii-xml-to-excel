# DTE-SII-XML to Excel Converter

Este proyecto proporciona una herramienta para convertir archivos XML de Documentos Tributarios Electrónicos (DTE) del Servicio de Impuestos Internos (SII) de Chile a un archivo Excel consolidado.

## Funcionalidades

- Parsea múltiples archivos XML de DTE en un directorio especificado.
- Extrae información relevante de cada DTE, incluyendo datos del emisor, receptor, detalles del documento y totales.
- Combina los datos de todos los XML en un único DataFrame de pandas.
- Genera un archivo Excel con todos los datos extraídos.
- Ajusta automáticamente el ancho de las columnas y la altura de las filas en el archivo Excel para mejorar la legibilidad.

## Requisitos

- Python 3.x
- pandas
- openpyxl
- xml.etree.ElementTree (incluido en la biblioteca estándar de Python)

## Uso

1. Coloca tus archivos XML de DTE en el directorio `archivos/`.
2. Ejecuta el script principal:

```
python prueba.py
```

3. El script generará un archivo Excel llamado `DTE_Recibidos_Combined.xlsx` con todos los datos extraídos.

## Estructura del Proyecto

- `programa.py`: Script principal que contiene toda la lógica de conversión.
- `archivos/`: Directorio donde se deben colocar los archivos XML a procesar.
- `DTE_Recibidos_Combined.xlsx`: Archivo Excel de salida con los datos combinados.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue para discutir cambios mayores antes de crear un pull request.
