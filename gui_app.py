"""
Outlook Email Search Tool - Aplicación GUI Principal
Ventana principal con tabs para búsqueda, búsqueda rápida y carpetas.
"""

import sys
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

from outlook_client import OutlookClient
from search import EmailSearch
from gui_search import SearchFrame
from gui_folders import FoldersFrame


class OutlookSearchApp(ttk.Window):
    """Ventana principal de la aplicación."""

    def __init__(self):
        super().__init__(
            title="📧 Outlook Email Search Tool — Banco Tanner",
            themename="cosmo",
            size=(1100, 720),
            minsize=(900, 600),
        )

        # Variables
        self.client = None
        self.searcher = None

        # Mostrar splash y conectar
        self._show_splash()

    def _show_splash(self):
        """Muestra pantalla de carga mientras conecta a Outlook."""
        self.splash_frame = ttk.Frame(self)
        self.splash_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

        ttk.Label(
            self.splash_frame,
            text="📧",
            font=("Segoe UI Emoji", 48),
        ).pack(pady=(0, 10))

        ttk.Label(
            self.splash_frame,
            text="Outlook Email Search Tool",
            font=("Segoe UI", 22, "bold"),
        ).pack()

        ttk.Label(
            self.splash_frame,
            text="Banco Tanner — Herramienta Interna",
            font=("Segoe UI", 11),
            foreground="gray",
        ).pack(pady=(5, 20))

        self.splash_status = ttk.StringVar(value="Conectando a Outlook...")
        ttk.Label(
            self.splash_frame,
            textvariable=self.splash_status,
            font=("Segoe UI", 10),
        ).pack()

        self.splash_progress = ttk.Progressbar(
            self.splash_frame, mode="indeterminate", length=300, bootstyle=INFO
        )
        self.splash_progress.pack(pady=(10, 0))
        self.splash_progress.start(15)

        # Conectar en thread
        thread = threading.Thread(target=self._connect_outlook, daemon=True)
        thread.start()

    def _connect_outlook(self):
        """Conecta a Outlook en background."""
        try:
            self.client = OutlookClient()
            self.searcher = EmailSearch(self.client)
            email = self.client.get_account_email()
            self.after(0, self._on_connected, email)
        except Exception as e:
            self.after(0, self._on_connection_error, str(e))

    def _on_connected(self, email: str):
        """Outlook conectado exitosamente."""
        self.splash_progress.stop()
        self.splash_frame.destroy()
        self._build_main_ui(email)

    def _on_connection_error(self, error_msg: str):
        """Error de conexión a Outlook."""
        self.splash_progress.stop()
        self.splash_status.set(f"❌ Error: {error_msg}")
        messagebox.showerror(
            "Error de Conexión",
            f"No se pudo conectar a Outlook.\n\n{error_msg}\n\n"
            "Asegúrate de que Outlook esté abierto y configurado.",
        )
        self.destroy()

    def _build_main_ui(self, email: str):
        """Construye la interfaz principal tras conexión exitosa."""
        # --- Barra de menú ---
        menubar = ttk.Menu(self)
        self.configure(menu=menubar)

        file_menu = ttk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Salir", command=self.destroy, accelerator="Alt+F4")
        menubar.add_cascade(label="Archivo", menu=file_menu)

        help_menu = ttk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Acerca de...", command=self._show_about)
        menubar.add_cascade(label="Ayuda", menu=help_menu)

        # --- Header ---
        header = ttk.Frame(self, padding=(15, 10))
        header.pack(fill=X)

        ttk.Label(
            header,
            text="📧 Outlook Email Search Tool",
            font=("Segoe UI", 16, "bold"),
        ).pack(side=LEFT)

        ttk.Label(
            header,
            text=f"📬 {email}",
            font=("Segoe UI", 10),
            foreground="gray",
        ).pack(side=RIGHT)

        ttk.Separator(self).pack(fill=X)

        # --- Notebook principal ---
        self.notebook = ttk.Notebook(self, padding=5)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=5)

        # Tab 1: Búsqueda
        self.search_frame = SearchFrame(self.notebook, self.searcher)
        self.notebook.add(self.search_frame, text="  🔍 Búsqueda  ")

        # Tab 2: Carpetas
        self.folders_frame = FoldersFrame(self.notebook, self.client)
        self.notebook.add(self.folders_frame, text="  📁 Carpetas  ")

        # --- Barra de estado ---
        status_bar = ttk.Frame(self, padding=(15, 5))
        status_bar.pack(fill=X, side=BOTTOM)

        ttk.Separator(self).pack(fill=X, side=BOTTOM)

        ttk.Label(
            status_bar,
            text=f"✓ Conectado a Outlook  |  Cuenta: {email}",
            font=("Segoe UI", 9),
            foreground="gray",
        ).pack(side=LEFT)

        ttk.Label(
            status_bar,
            text="Banco Tanner — Herramienta Interna",
            font=("Segoe UI", 9),
            foreground="gray",
        ).pack(side=RIGHT)

    def _show_about(self):
        """Muestra diálogo 'Acerca de'."""
        messagebox.showinfo(
            "Acerca de",
            "📧 Outlook Email Search Tool\n\n"
            "Herramienta para buscar, filtrar y exportar\n"
            "correos desde Microsoft Outlook.\n\n"
            "Banco Tanner — Herramienta Interna\n"
            "Versión 2.0 — Interfaz Gráfica",
        )


def main():
    """Punto de entrada de la aplicación."""
    app = OutlookSearchApp()
    app.mainloop()


if __name__ == "__main__":
    main()
