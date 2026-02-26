"""
Módulo de generación de reportes y tablas con resultados de búsqueda.
Soporta visualización en consola, exportación a Excel y CSV.
"""

import os
from datetime import datetime
import pandas as pd
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

console = Console()


def display_table(results: list, max_rows: int = 50):
    """
    Muestra los resultados de búsqueda en una tabla formateada en consola.
    
    Args:
        results: Lista de diccionarios con datos de correos
        max_rows: Máximo de filas a mostrar
    """
    if not results:
        console.print("[yellow]No hay resultados para mostrar.[/yellow]")
        return

    table = Table(
        title=f"📧 Resultados de Búsqueda ({len(results)} correos)",
        show_lines=True,
        header_style="bold cyan",
        border_style="blue",
    )

    table.add_column("#", style="dim", width=4, justify="right")
    table.add_column("Fecha", width=12)
    table.add_column("Hora", width=10)
    table.add_column("Remitente", width=25, no_wrap=True)
    table.add_column("Asunto", width=45)
    table.add_column("📎", width=3, justify="center")
    table.add_column("Importancia", width=10, justify="center")

    for i, email in enumerate(results[:max_rows], 1):
        att_icon = "✓" if email.get("has_attachments") else ""
        imp = email.get("importance", "Normal")
        imp_style = {"Alta": "[red]Alta[/red]", "Baja": "[dim]Baja[/dim]"}.get(
            imp, imp
        )

        table.add_row(
            str(i),
            email.get("date", ""),
            email.get("time", ""),
            _truncate(email.get("sender_name", ""), 25),
            _truncate(email.get("subject", ""), 45),
            att_icon,
            imp_style,
        )

    console.print(table)

    if len(results) > max_rows:
        console.print(
            f"[dim]... mostrando {max_rows} de {len(results)} resultados[/dim]"
        )


def display_detail(email_data: dict):
    """
    Muestra el detalle completo de un correo.
    
    Args:
        email_data: Diccionario con datos del correo
    """
    content = (
        f"[bold]De:[/bold] {email_data.get('sender_name', 'N/A')} "
        f"<{email_data.get('sender_email', 'N/A')}>\n"
        f"[bold]Para:[/bold] {email_data.get('to', 'N/A')}\n"
        f"[bold]CC:[/bold] {email_data.get('cc', 'N/A')}\n"
        f"[bold]Fecha:[/bold] {email_data.get('date', 'N/A')} "
        f"{email_data.get('time', '')}\n"
        f"[bold]Importancia:[/bold] {email_data.get('importance', 'Normal')}\n"
        f"[bold]Categorías:[/bold] {email_data.get('categories', 'N/A')}\n"
        f"[bold]Tamaño:[/bold] {email_data.get('size_kb', 0)} KB\n"
    )

    if email_data.get("has_attachments"):
        att_names = ", ".join(email_data.get("attachment_names", []))
        content += (
            f"[bold]Adjuntos ({email_data.get('attachment_count', 0)}):[/bold] "
            f"{att_names}\n"
        )

    content += f"\n[bold]Vista previa:[/bold]\n{email_data.get('body_preview', 'N/A')}"

    console.print(
        Panel(
            content,
            title=f"📧 {email_data.get('subject', 'Sin asunto')}",
            border_style="cyan",
        )
    )


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
        console.print("[yellow]No hay resultados para exportar.[/yellow]")
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

    console.print(f"[green]✓ Reporte exportado a:[/green] {os.path.abspath(filepath)}")
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
        console.print("[yellow]No hay resultados para exportar.[/yellow]")
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

    console.print(f"[green]✓ Reporte exportado a:[/green] {os.path.abspath(filepath)}")
    return os.path.abspath(filepath)


def generate_summary(results: list):
    """
    Genera un resumen estadístico de los resultados de búsqueda.
    
    Args:
        results: Lista de diccionarios con datos de correos
    """
    if not results:
        console.print("[yellow]No hay resultados para resumir.[/yellow]")
        return

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

    summary = Table(title="📊 Resumen de Búsqueda", border_style="green")
    summary.add_column("Métrica", style="bold")
    summary.add_column("Valor", justify="right")

    summary.add_row("Total correos", str(total))
    summary.add_row("Con adjuntos", f"{with_att} ({round(with_att/total*100)}%)")
    summary.add_row("Total adjuntos", str(total_att))
    if dates:
        summary.add_row("Fecha más antigua", min(dates))
        summary.add_row("Fecha más reciente", max(dates))

    console.print(summary)

    # Top remitentes
    if top_senders:
        sender_table = Table(title="👤 Top Remitentes", border_style="cyan")
        sender_table.add_column("Remitente", width=35)
        sender_table.add_column("Correos", justify="right", width=10)
        for name, count in top_senders:
            sender_table.add_row(_truncate(name, 35), str(count))
        console.print(sender_table)


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
