"""
Módulo de exportación de archivos adjuntos.
Permite descargar adjuntos de los correos encontrados a un directorio.
"""

import os
from datetime import datetime
from typing import Optional
from rich.console import Console
from rich.progress import Progress, BarColumn, TextColumn, MofNCompleteColumn

console = Console()


def export_attachments(
    results: list,
    output_dir: str,
    organize_by: str = "flat",
    file_types: Optional[list] = None,
    skip_inline: bool = True,
) -> dict:
    """
    Exporta archivos adjuntos de los correos encontrados.
    
    Args:
        results: Lista de resultados de búsqueda (con _outlook_item)
        output_dir: Directorio destino para guardar los archivos
        organize_by: Modo de organización:
            - 'flat': Todos en la misma carpeta
            - 'sender': Subcarpeta por remitente
            - 'date': Subcarpeta por fecha (YYYY-MM-DD)
            - 'subject': Subcarpeta por asunto
        file_types: Lista de extensiones a filtrar (ej: ['.pdf', '.xlsx'])
                    None = todos los tipos
        skip_inline: Si True, omite imágenes embebidas (inline)
        
    Returns:
        Diccionario con resumen de la exportación
    """
    # Crear directorio base
    os.makedirs(output_dir, exist_ok=True)

    stats = {
        "total_emails": 0,
        "emails_with_attachments": 0,
        "total_attachments": 0,
        "exported": 0,
        "skipped": 0,
        "errors": 0,
        "files": [],
    }

    # Filtrar solo correos con adjuntos
    emails_with_att = [r for r in results if r.get("has_attachments")]
    stats["total_emails"] = len(results)
    stats["emails_with_attachments"] = len(emails_with_att)

    if not emails_with_att:
        console.print("[yellow]No se encontraron correos con adjuntos.[/yellow]")
        return stats

    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        console=console,
    ) as progress:
        task = progress.add_task(
            "Exportando adjuntos...", total=len(emails_with_att)
        )

        for email_data in emails_with_att:
            item = email_data.get("_outlook_item")
            if not item:
                progress.advance(task)
                continue

            try:
                att_count = item.Attachments.Count
                stats["total_attachments"] += att_count

                for i in range(att_count):
                    att = item.Attachments.Item(i + 1)

                    try:
                        filename = att.FileName

                        # Omitir imágenes inline
                        if skip_inline:
                            try:
                                # Los adjuntos inline tienen PropertyAccessor
                                content_id = att.PropertyAccessor.GetProperty(
                                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                                )
                                if content_id:
                                    stats["skipped"] += 1
                                    continue
                            except Exception:
                                pass

                        # Filtrar por tipo de archivo
                        if file_types:
                            ext = os.path.splitext(filename)[1].lower()
                            if ext not in [ft.lower() for ft in file_types]:
                                stats["skipped"] += 1
                                continue

                        # Determinar subdirectorio
                        sub_dir = _get_subfolder(
                            organize_by, email_data
                        )
                        target_dir = os.path.join(output_dir, sub_dir) if sub_dir else output_dir
                        os.makedirs(target_dir, exist_ok=True)

                        # Manejar nombres duplicados
                        filepath = _get_unique_path(target_dir, filename)

                        # Guardar archivo
                        att.SaveAsFile(filepath)
                        stats["exported"] += 1
                        stats["files"].append(
                            {
                                "filename": os.path.basename(filepath),
                                "path": filepath,
                                "from_subject": email_data.get("subject", ""),
                                "from_sender": email_data.get("sender_name", ""),
                                "date": email_data.get("date", ""),
                            }
                        )

                    except Exception as e:
                        stats["errors"] += 1
                        console.print(
                            f"  [red]Error al exportar '{att.FileName}': {e}[/red]"
                        )

            except Exception as e:
                stats["errors"] += 1
                console.print(f"  [red]Error procesando correo: {e}[/red]")

            progress.advance(task)

    # Resumen
    console.print(f"\n[green]✓ Exportación completada:[/green]")
    console.print(f"  📧 Correos procesados: {stats['emails_with_attachments']}")
    console.print(f"  📎 Adjuntos exportados: {stats['exported']}")
    if stats["skipped"]:
        console.print(f"  ⏭️  Omitidos: {stats['skipped']}")
    if stats["errors"]:
        console.print(f"  ❌ Errores: {stats['errors']}")
    console.print(f"  📁 Directorio: {output_dir}")

    return stats


def _get_subfolder(organize_by: str, email_data: dict) -> str:
    """Genera nombre de subcarpeta según el modo de organización."""
    if organize_by == "sender":
        name = email_data.get("sender_name", "Desconocido")
        return _sanitize_foldername(name)
    elif organize_by == "date":
        date_str = email_data.get("date", "")
        try:
            dt = datetime.strptime(date_str, "%d-%m-%Y")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            return "sin_fecha"
    elif organize_by == "subject":
        subject = email_data.get("subject", "Sin asunto")
        return _sanitize_foldername(subject[:50])
    return ""


def _sanitize_foldername(name: str) -> str:
    """Sanitiza un nombre de carpeta eliminando caracteres no válidos."""
    invalid_chars = '<>:"/\\|?*'
    for c in invalid_chars:
        name = name.replace(c, "_")
    return name.strip(". ") or "sin_nombre"


def _get_unique_path(directory: str, filename: str) -> str:
    """Genera una ruta única si el archivo ya existe (agrega sufijo numérico)."""
    filepath = os.path.join(directory, filename)
    if not os.path.exists(filepath):
        return filepath

    name, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(filepath):
        filepath = os.path.join(directory, f"{name}_{counter}{ext}")
        counter += 1
    return filepath
