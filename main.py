"""
Outlook Email Search Tool - Interfaz Principal
Aplicación interactiva para buscar, filtrar y exportar correos desde Outlook.
"""

import os
import sys
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt, IntPrompt, Confirm
from rich.table import Table
from rich.text import Text

from outlook_client import OutlookClient
from search import EmailSearch
from reports import display_table, display_detail, export_to_excel, export_to_csv, generate_summary
from attachments import export_attachments

console = Console()

# Banner de la aplicación
BANNER = """
[bold cyan]╔══════════════════════════════════════════════════════╗
║         📧  Outlook Email Search Tool  📧            ║
║         Banco Tanner - Herramienta Interna           ║
╚══════════════════════════════════════════════════════╝[/bold cyan]
"""


def main():
    """Punto de entrada principal de la aplicación."""
    console.print(BANNER)

    # Conectar a Outlook
    try:
        client = OutlookClient()
        email = client.get_account_email()
        console.print(f"  📬 Cuenta: [bold]{email}[/bold]\n")
    except Exception:
        console.print("[red]No se pudo conectar a Outlook. Cerrando.[/red]")
        sys.exit(1)

    searcher = EmailSearch(client)
    last_results = []  # Almacena la última búsqueda

    while True:
        console.print()
        _show_menu()
        choice = Prompt.ask(
            "\n[bold]Selecciona una opción[/bold]",
            choices=["1", "2", "3", "4", "5", "6", "7", "0"],
            default="1",
        )

        if choice == "1":
            last_results = _action_search(searcher)
        elif choice == "2":
            _action_export_attachments(searcher, last_results)
        elif choice == "3":
            _action_export_report(last_results)
        elif choice == "4":
            _action_view_detail(last_results)
        elif choice == "5":
            _action_summary(last_results)
        elif choice == "6":
            _action_list_folders(client)
        elif choice == "7":
            _action_quick_search(searcher)
        elif choice == "0":
            console.print("[cyan]¡Hasta luego! 👋[/cyan]")
            break


def _show_menu():
    """Muestra el menú principal."""
    menu = Table(show_header=False, border_style="cyan", box=None, padding=(0, 2))
    menu.add_column("Opción", style="bold yellow", width=4)
    menu.add_column("Descripción")

    menu.add_row("1", "🔍 Buscar correos (con filtros)")
    menu.add_row("2", "📎 Exportar adjuntos de búsqueda")
    menu.add_row("3", "📊 Exportar reporte (Excel/CSV)")
    menu.add_row("4", "📄 Ver detalle de un correo")
    menu.add_row("5", "📈 Resumen estadístico")
    menu.add_row("6", "📁 Ver carpetas del buzón")
    menu.add_row("7", "⚡ Búsqueda rápida")
    menu.add_row("0", "🚪 Salir")

    console.print(Panel(menu, title="[bold]Menú Principal[/bold]", border_style="cyan"))


def _action_search(searcher: EmailSearch) -> list:
    """Acción: Búsqueda con filtros."""
    console.print("\n[bold cyan]═══ Búsqueda de Correos ═══[/bold cyan]")
    console.print("[dim]Presiona Enter para omitir un filtro[/dim]\n")

    subject = Prompt.ask("  Asunto (texto parcial)", default="").strip() or None
    sender = Prompt.ask("  Remitente (nombre o email)", default="").strip() or None
    date_from = Prompt.ask("  Fecha desde (DD-MM-YYYY)", default="").strip() or None
    date_to = Prompt.ask("  Fecha hasta (DD-MM-YYYY)", default="").strip() or None

    # Opciones avanzadas
    advanced = Confirm.ask("  ¿Filtros avanzados?", default=False)

    folder = "inbox"
    has_attachments = None
    body_contains = None
    max_results = 100

    if advanced:
        folder_input = Prompt.ask(
            "  Carpeta (inbox/sent/drafts)", default="inbox"
        ).strip()
        folder = folder_input if folder_input else "inbox"

        att_filter = Prompt.ask(
            "  ¿Solo con adjuntos? (si/no/todos)", default="todos"
        ).strip().lower()
        if att_filter == "si":
            has_attachments = True
        elif att_filter == "no":
            has_attachments = False

        body_contains = (
            Prompt.ask("  Buscar en cuerpo", default="").strip() or None
        )
        max_results = IntPrompt.ask("  Máx. resultados", default=100)

    # Ejecutar búsqueda
    console.print()
    results = searcher.search(
        subject=subject,
        sender=sender,
        date_from=date_from,
        date_to=date_to,
        folder=folder,
        has_attachments=has_attachments,
        body_contains=body_contains,
        max_results=max_results,
    )

    if results:
        display_table(results)
    else:
        console.print("[yellow]No se encontraron correos con esos filtros.[/yellow]")

    return results


def _action_export_attachments(searcher: EmailSearch, last_results: list):
    """Acción: Exportar adjuntos."""
    console.print("\n[bold cyan]═══ Exportar Adjuntos ═══[/bold cyan]\n")

    if last_results:
        use_last = Confirm.ask(
            f"  ¿Usar última búsqueda ({len(last_results)} correos)?", default=True
        )
        if not use_last:
            last_results = _action_search(searcher)
    else:
        console.print("  [yellow]No hay búsqueda previa. Realizando nueva búsqueda...[/yellow]\n")
        last_results = _action_search(searcher)

    if not last_results:
        return

    # Directorio destino
    default_dir = os.path.join(os.path.expanduser("~"), "Desktop", "adjuntos_outlook")
    output_dir = Prompt.ask("  Directorio destino", default=default_dir).strip()

    # Organización
    organize = Prompt.ask(
        "  Organizar por (flat/sender/date/subject)",
        default="flat",
        choices=["flat", "sender", "date", "subject"],
    )

    # Filtro de tipos
    types_input = Prompt.ask(
        "  Tipos de archivo (ej: .pdf,.xlsx o Enter=todos)", default=""
    ).strip()
    file_types = None
    if types_input:
        file_types = [t.strip() for t in types_input.split(",")]

    # Ejecutar
    console.print()
    export_attachments(
        results=last_results,
        output_dir=output_dir,
        organize_by=organize,
        file_types=file_types,
    )


def _action_export_report(last_results: list):
    """Acción: Exportar reporte."""
    console.print("\n[bold cyan]═══ Exportar Reporte ═══[/bold cyan]\n")

    if not last_results:
        console.print("[yellow]No hay resultados. Realiza una búsqueda primero (opción 1).[/yellow]")
        return

    format_choice = Prompt.ask(
        "  Formato", default="excel", choices=["excel", "csv"]
    )

    filepath = Prompt.ask(
        "  Ruta del archivo (Enter=automático)", default=""
    ).strip() or None

    if format_choice == "excel":
        export_to_excel(last_results, filepath)
    else:
        export_to_csv(last_results, filepath)


def _action_view_detail(last_results: list):
    """Acción: Ver detalle de un correo."""
    if not last_results:
        console.print("\n[yellow]No hay resultados. Realiza una búsqueda primero (opción 1).[/yellow]")
        return

    console.print(f"\n[dim]Hay {len(last_results)} correos en la última búsqueda.[/dim]")
    idx = IntPrompt.ask("  Número de correo a ver", default=1)

    if 1 <= idx <= len(last_results):
        console.print()
        display_detail(last_results[idx - 1])
    else:
        console.print(f"[red]Número inválido. Rango: 1-{len(last_results)}[/red]")


def _action_summary(last_results: list):
    """Acción: Resumen estadístico."""
    if not last_results:
        console.print("\n[yellow]No hay resultados. Realiza una búsqueda primero (opción 1).[/yellow]")
        return

    console.print()
    generate_summary(last_results)


def _action_list_folders(client: OutlookClient):
    """Acción: Listar carpetas del buzón."""
    console.print("\n[bold cyan]═══ Carpetas del Buzón ═══[/bold cyan]\n")

    folders = client.list_folders(max_depth=2)

    table = Table(border_style="blue", header_style="bold")
    table.add_column("Carpeta", width=50)
    table.add_column("Items", justify="right", width=10)

    for name, path, count, indent in folders:
        prefix = "  " * indent + ("📁 " if indent == 0 else "📂 ")
        style = "bold" if indent == 0 else ""
        table.add_row(f"{prefix}{name}", str(count) if count else "", style=style)

    console.print(table)


def _action_quick_search(searcher: EmailSearch):
    """Acción: Búsqueda rápida (un solo término)."""
    console.print("\n[bold cyan]═══ Búsqueda Rápida ═══[/bold cyan]\n")

    term = Prompt.ask("  Término de búsqueda").strip()
    if not term:
        return

    search_in = Prompt.ask(
        "  Buscar en", default="subject", choices=["subject", "sender", "all"]
    )

    console.print()
    if search_in == "subject":
        results = searcher.search(subject=term, max_results=50)
    elif search_in == "sender":
        results = searcher.search(sender=term, max_results=50)
    else:
        # Buscar en todo: primero por asunto, luego por remitente
        results = searcher.search(subject=term, max_results=50)
        results_sender = searcher.search(sender=term, max_results=50)
        # Combinar sin duplicados (por asunto+fecha)
        seen = {(r["subject"], r["date"], r["time"]) for r in results}
        for r in results_sender:
            key = (r["subject"], r["date"], r["time"])
            if key not in seen:
                results.append(r)
                seen.add(key)

    if results:
        display_table(results)
        generate_summary(results)
    else:
        console.print("[yellow]No se encontraron resultados.[/yellow]")

    return results


if __name__ == "__main__":
    main()
