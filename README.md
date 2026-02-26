# 📧 Outlook Email Search Tool

Herramienta gráfica para buscar, filtrar y exportar correos desde Microsoft Outlook.

**Banco Tanner — Herramienta Interna**

## Características

- **Búsqueda Avanzada**: Filtra por asunto, remitente, fechas, carpeta, adjuntos y contenido del cuerpo
- **Búsqueda Rápida**: Busca por un solo término en asunto, remitente o ambos
- **Tabla de Resultados**: Visualiza resultados ordenables con información clave
- **Exportar a Excel**: Exporta los resultados directamente a un archivo `.xlsx` con un botón
- **Exportar a CSV**: Exporta los resultados a formato CSV
- **Exportar Adjuntos**: Descarga archivos adjuntos organizados por remitente, fecha o asunto
- **Ver Detalle**: Visualiza información completa de cada correo
- **Resumen Estadístico**: Top remitentes, rango de fechas, conteo de adjuntos
- **Explorar Carpetas**: Navega la estructura de carpetas del buzón

## Requisitos

- Windows con Microsoft Outlook instalado y configurado
- Python 3.9+
- Outlook debe estar abierto al ejecutar la aplicación

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

La aplicación abrirá una interfaz gráfica con pestañas para:

1. **Búsqueda**: Configura filtros y busca correos. Los resultados se muestran en una tabla interactiva.
2. **Carpetas**: Explora la estructura de carpetas de tu buzón.

### Exportar resultados

Después de realizar una búsqueda, usa los botones en la parte inferior:
- **📊 Exportar Excel** — Genera un archivo .xlsx con los resultados
- **📋 Exportar CSV** — Genera un archivo .csv
- **📎 Exportar Adjuntos** — Descarga los archivos adjuntos a un directorio
- **📄 Ver Detalle** — Abre la información completa del correo seleccionado
- **📈 Resumen** — Muestra estadísticas de los resultados

## Dependencias

| Paquete | Uso |
|---------|-----|
| `pywin32` | Conexión COM con Outlook |
| `pandas` | Manipulación de datos para exportación |
| `openpyxl` | Escritura de archivos Excel |
| `ttkbootstrap` | Interfaz gráfica moderna |

## Estructura del Proyecto

```
Correo_Python/
├── main.py              # Punto de entrada
├── gui_app.py           # Ventana principal y navegación
├── gui_search.py        # Pestañas de búsqueda y tabla de resultados
├── gui_detail.py        # Ventana de detalle de correo
├── gui_attachments.py   # Diálogo de exportación de adjuntos
├── gui_folders.py       # Pestaña de carpetas del buzón
├── outlook_client.py    # Conexión COM con Outlook
├── search.py            # Motor de búsqueda con filtros DASL
├── attachments.py       # Lógica de exportación de adjuntos
├── reports.py           # Exportación a Excel/CSV y estadísticas
├── requirements.txt     # Dependencias
└── README.md            # Este archivo
```
