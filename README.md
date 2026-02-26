# 📧 Outlook Email Search Tool

Herramienta de búsqueda y gestión de correos electrónicos de Outlook para uso interno de Banco Tanner. Permite buscar, filtrar, exportar adjuntos y generar reportes desde la bandeja de correo corporativa.

## Requisitos

- **Python 3.9+**
- **Microsoft Outlook** instalado y configurado con la cuenta corporativa
- **Windows** (usa COM automation via `pywin32`)

## Instalación

```bash
# Clonar el repositorio
git clone https://github.com/TU_USUARIO/Correo_Python.git
cd Correo_Python

# Instalar dependencias
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

Se abrirá un menú interactivo en consola:

| Opción | Función |
|--------|---------|
| 1 | 🔍 Buscar correos con filtros (asunto, remitente, fechas, etc.) |
| 2 | 📎 Exportar adjuntos de los correos encontrados |
| 3 | 📊 Exportar resultados a Excel o CSV |
| 4 | 📄 Ver detalle completo de un correo |
| 5 | 📈 Ver resumen estadístico de la búsqueda |
| 6 | 📁 Listar carpetas del buzón |
| 7 | ⚡ Búsqueda rápida |
| 0 | 🚪 Salir |

## Filtros de Búsqueda

- **Asunto**: búsqueda parcial en el encabezado del correo
- **Remitente**: por nombre o dirección de email
- **Rango de fechas**: formato `DD-MM-YYYY`
- **Adjuntos**: filtrar solo correos con/sin adjuntos
- **Cuerpo**: buscar texto dentro del cuerpo del correo
- **Carpeta**: buscar en Inbox, Sent, Drafts, etc.

## Exportación de Adjuntos

Los adjuntos se pueden organizar en subcarpetas por:
- `flat` — todos en la misma carpeta
- `sender` — agrupados por remitente
- `date` — agrupados por fecha
- `subject` — agrupados por asunto

También se puede filtrar por tipo de archivo (ej: `.pdf`, `.xlsx`).

## Reportes

Los reportes incluyen las columnas:
- Remitente, Email, Asunto, Fecha, Hora
- Cantidad de adjuntos, Importancia, Categorías, Tamaño

Formatos disponibles: **Excel (.xlsx)** y **CSV (.csv)**.

## Estructura del Proyecto

```
Correo_Python/
├── main.py            # Interfaz CLI principal
├── outlook_client.py  # Conexión COM con Outlook
├── search.py          # Motor de búsqueda con filtros DASL
├── attachments.py     # Exportación de archivos adjuntos
├── reports.py         # Generación de tablas y reportes
├── requirements.txt   # Dependencias Python
├── .gitignore         # Archivos ignorados por Git
└── README.md          # Este archivo
```

## Notas

- Outlook debe estar **abierto** al ejecutar la herramienta
- Las búsquedas usan filtros DASL de Outlook para rendimiento óptimo
- Los reportes se guardan en la carpeta `reportes/` por defecto
- Las fechas se manejan en formato `DD-MM-YYYY`
