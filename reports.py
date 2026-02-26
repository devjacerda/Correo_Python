"""
Módulo de generación de reportes y exportación de resultados.
Soporta exportación a Excel y CSV, y generación de resúmenes estadísticos.
"""

import os
from datetime import datetime
import pandas as pd


def export_to_excel(results: list, filepath: str = None) -> str:
    """
    Exporta los resultados de búsqueda a un archivo Excel.
    
    Args:
        results: Lista de diccionarios con datos de correos
        filepath: Ruta del archivo. Si None, genera nombre automático.
        
    Returns:
        Ruta del archivo generado
    """
    if not results:
        return ""

    if filepath is None:
        os.makedirs("reportes", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join("reportes", f"busqueda_{timestamp}.xlsx")

    # Preparar datos (sin objetos COM)
    clean_results = _clean_for_export(results)
    df = pd.DataFrame(clean_results)

    # Renombrar columnas a español
    column_names = {
        "subject": "Asunto",
        "sender_name": "Remitente",
        "sender_email": "Email Remitente",
        "to": "Destinatario",
        "cc": "CC",
        "date": "Fecha",
        "time": "Hora",
        "has_attachments": "Tiene Adjuntos",
        "attachment_count": "Nº Adjuntos",
        "attachment_names": "Nombres Adjuntos",
        "importance": "Importancia",
        "categories": "Categorías",
        "size_kb": "Tamaño (KB)",
        "body_preview": "Vista Previa",
    }
    df = df.rename(columns=column_names)

    # Exportar con formato
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")

        # Ajustar ancho de columnas
        worksheet = writer.sheets["Resultados"]
        for i, col in enumerate(df.columns, 1):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(col),
            )
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=i).column_letter
            ].width = min(max_length + 2, 50)

    return os.path.abspath(filepath)


def export_to_csv(results: list, filepath: str = None) -> str:
    """
    Exporta los resultados de búsqueda a un archivo CSV.
    
    Args:
        results: Lista de diccionarios con datos de correos
        filepath: Ruta del archivo. Si None, genera nombre automático.
        
    Returns:
        Ruta del archivo generado
    """
    if not results:
        return ""

    if filepath is None:
        os.makedirs("reportes", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = os.path.join("reportes", f"busqueda_{timestamp}.csv")

    clean_results = _clean_for_export(results)
    df = pd.DataFrame(clean_results)

    column_names = {
        "subject": "Asunto",
        "sender_name": "Remitente",
        "sender_email": "Email Remitente",
        "to": "Destinatario",
        "cc": "CC",
        "date": "Fecha",
        "time": "Hora",
        "has_attachments": "Tiene Adjuntos",
        "attachment_count": "Nº Adjuntos",
        "attachment_names": "Nombres Adjuntos",
        "importance": "Importancia",
        "categories": "Categorías",
        "size_kb": "Tamaño (KB)",
        "body_preview": "Vista Previa",
    }
    df = df.rename(columns=column_names)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")

    return os.path.abspath(filepath)


def generate_summary(results: list) -> dict:
    """
    Genera un resumen estadístico de los resultados de búsqueda.
    
    Args:
        results: Lista de diccionarios con datos de correos
        
    Returns:
        Diccionario con el resumen estadístico
    """
    if not results:
        return {}

    total = len(results)
    with_att = sum(1 for r in results if r.get("has_attachments"))
    total_att = sum(r.get("attachment_count", 0) for r in results)

    # Top remitentes
    senders = {}
    for r in results:
        name = r.get("sender_name", "Desconocido")
        senders[name] = senders.get(name, 0) + 1
    top_senders = sorted(senders.items(), key=lambda x: x[1], reverse=True)[:5]

    # Rango de fechas
    dates = [r.get("date", "") for r in results if r.get("date") != "N/A"]

    summary = {
        "total": total,
        "with_attachments": with_att,
        "pct_attachments": round(with_att / total * 100) if total > 0 else 0,
        "total_attachments": total_att,
        "date_min": min(dates) if dates else "N/A",
        "date_max": max(dates) if dates else "N/A",
        "top_senders": top_senders,
    }

    return summary


def _clean_for_export(results: list) -> list:
    """Limpia los resultados para exportación (elimina objetos COM)."""
    clean = []
    for r in results:
        row = {k: v for k, v in r.items() if k != "_outlook_item"}
        # Convertir listas a string
        if "attachment_names" in row:
            row["attachment_names"] = ", ".join(row["attachment_names"])
        clean.append(row)
    return clean


def _truncate(text: str, max_len: int) -> str:
    """Trunca texto a longitud máxima."""
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."
