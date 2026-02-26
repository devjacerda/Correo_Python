"""
Outlook Email Search Tool - Aplicación GUI Principal
Ventana principal con tabs para búsqueda y carpetas.
Usa un worker thread dedicado para operaciones COM con Outlook.
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

from outlook_worker import OutlookWorker
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

        self.worker = None
        self.search_frame = None
        self.folders_frame = None

        self._show_splash()
        self._start_worker()

    def _show_splash(self):
        """Muestra pantalla de carga."""
        self.splash = ttk.Frame(self)
        self.splash.place(relx=0.5, rely=0.5, anchor=CENTER)

        ttk.Label(self.splash, text="📧", font=("Segoe UI Emoji", 48)).pack(pady=(0, 10))
        ttk.Label(
            self.splash, text="Outlook Email Search Tool",
            font=("Segoe UI", 22, "bold"),
        ).pack()
        ttk.Label(
            self.splash, text="Banco Tanner — Herramienta Interna",
            font=("Segoe UI", 11), foreground="gray",
        ).pack(pady=(5, 20))

        self.splash_status = ttk.StringVar(value="Conectando a Outlook...")
        ttk.Label(self.splash, textvariable=self.splash_status, font=("Segoe UI", 10)).pack()

        self.splash_bar = ttk.Progressbar(self.splash, mode="indeterminate", length=300, bootstyle=INFO)
        self.splash_bar.pack(pady=(10, 0))
        self.splash_bar.start(15)

    def _start_worker(self):
        """Inicia el worker thread de Outlook."""
        self.worker = OutlookWorker(self)
        self.worker.start()

    # === Callbacks del worker ===

    def _on_worker_ready(self, email: str):
        """Worker conectado exitosamente."""
        self.splash_bar.stop()
        self.splash.destroy()
        self._build_main_ui(email)

    def _on_worker_error(self, error_msg: str):
        """Error de conexión."""
        self.splash_bar.stop()
        self.splash_status.set(f"❌ Error")
        messagebox.showerror(
            "Error de Conexión",
            f"No se pudo conectar a Outlook.\n\n{error_msg}\n\n"
            "Asegúrate de que Outlook esté abierto y configurado.",
        )
        self.destroy()

    def _on_search_progress(self, current, msg):
        """Progreso de búsqueda — actualiza la UI."""
        if self.search_frame:
            self.search_frame.status_var.set(f"🔍 {msg}")

    def _on_attachment_progress(self, current, total, msg):
        """Progreso de exportación de adjuntos."""
        pass  # Manejado desde gui_attachments

    # === UI Principal ===

    def _build_main_ui(self, email: str):
        """Construye la interfaz principal."""
        # --- Menú ---
        menubar = ttk.Menu(self)
        self.configure(menu=menubar)

        file_menu = ttk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Salir", command=self.destroy)
        menubar.add_cascade(label="Archivo", menu=file_menu)

        help_menu = ttk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Acerca de...", command=self._show_about)
        menubar.add_cascade(label="Ayuda", menu=help_menu)

        # --- Header ---
        header = ttk.Frame(self, padding=(15, 10))
        header.pack(fill=X)

        ttk.Label(
            header, text="📧 Outlook Email Search Tool",
            font=("Segoe UI", 16, "bold"),
        ).pack(side=LEFT)

        ttk.Label(
            header, text=f"📬 {email}",
            font=("Segoe UI", 10), foreground="gray",
        ).pack(side=RIGHT)

        ttk.Separator(self).pack(fill=X)

        # --- Tabs ---
        notebook = ttk.Notebook(self, padding=5)
        notebook.pack(fill=BOTH, expand=True, padx=10, pady=5)

        self.search_frame = SearchFrame(notebook, self.worker)
        notebook.add(self.search_frame, text="  🔍 Búsqueda  ")

        self.folders_frame = FoldersFrame(notebook, self.worker)
        notebook.add(self.folders_frame, text="  📁 Carpetas  ")

        # --- Status bar ---
        ttk.Separator(self).pack(fill=X, side=BOTTOM)
        status_bar = ttk.Frame(self, padding=(15, 5))
        status_bar.pack(fill=X, side=BOTTOM)

        ttk.Label(
            status_bar, text=f"✓ Conectado  |  {email}",
            font=("Segoe UI", 9), foreground="gray",
        ).pack(side=LEFT)

        ttk.Label(
            status_bar, text="Banco Tanner",
            font=("Segoe UI", 9), foreground="gray",
        ).pack(side=RIGHT)

    def _show_about(self):
        messagebox.showinfo(
            "Acerca de",
            "📧 Outlook Email Search Tool\n\n"
            "Buscar, filtrar y exportar correos de Outlook.\n\n"
            "Banco Tanner — v2.0 GUI",
        )


def main():
    app = OutlookSearchApp()
    app.mainloop()


if __name__ == "__main__":
    main()
