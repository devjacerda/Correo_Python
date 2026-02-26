"""
Diálogo de exportación de archivos adjuntos.
Permite configurar directorio, organización y tipos de archivo antes de exportar.
"""

import os
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

from attachments import export_attachments


class AttachmentsDialog(ttk.Toplevel):
    """Diálogo modal para exportar adjuntos."""

    def __init__(self, parent, results: list):
        super().__init__(parent)
        self.title("📎 Exportar Adjuntos")
        self.geometry("500x420")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.results = results
        self._build_ui()

        # Centrar
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 500) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 420) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _build_ui(self):
        """Construye la interfaz del diálogo."""
        main = ttk.Frame(self, padding=20)
        main.pack(fill=BOTH, expand=True)

        # Info
        emails_with_att = sum(1 for r in self.results if r.get("has_attachments"))
        ttk.Label(
            main,
            text=f"Se encontraron {emails_with_att} correos con adjuntos de {len(self.results)} totales.",
            font=("Segoe UI", 10),
        ).pack(anchor=W, pady=(0, 15))

        # --- Directorio destino ---
        ttk.Label(main, text="Directorio destino:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        dir_frame = ttk.Frame(main)
        dir_frame.pack(fill=X, pady=(2, 10))

        default_dir = os.path.join(os.path.expanduser("~"), "Desktop", "adjuntos_outlook")
        self.dir_var = ttk.StringVar(value=default_dir)
        ttk.Entry(dir_frame, textvariable=self.dir_var, font=("Segoe UI", 9)).pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        ttk.Button(dir_frame, text="📂 Examinar", bootstyle=OUTLINE, command=self._browse_dir).pack(side=RIGHT)

        # --- Organización ---
        ttk.Label(main, text="Organizar por:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        self.organize_var = ttk.StringVar(value="flat")
        org_frame = ttk.Frame(main)
        org_frame.pack(fill=X, pady=(2, 10))

        options = [
            ("Sin organizar (flat)", "flat"),
            ("Por remitente", "sender"),
            ("Por fecha", "date"),
            ("Por asunto", "subject"),
        ]
        for text, value in options:
            ttk.Radiobutton(
                org_frame, text=text, variable=self.organize_var, value=value
            ).pack(side=LEFT, padx=(0, 15))

        # --- Filtro de tipos ---
        ttk.Label(main, text="Tipos de archivo (vacío = todos):", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        self.types_var = ttk.StringVar(value="")
        ttk.Entry(main, textvariable=self.types_var, font=("Segoe UI", 9)).pack(fill=X, pady=(2, 5))
        ttk.Label(
            main, text="Ejemplo: .pdf, .xlsx, .docx", font=("Segoe UI", 8), foreground="gray"
        ).pack(anchor=W, pady=(0, 15))

        # --- Barra de progreso ---
        self.progress_var = ttk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(main, variable=self.progress_var, bootstyle=SUCCESS)
        self.progress_bar.pack(fill=X, pady=(0, 5))

        self.progress_label = ttk.StringVar(value="")
        ttk.Label(main, textvariable=self.progress_label, font=("Segoe UI", 9)).pack(anchor=W, pady=(0, 10))

        # --- Botones ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=X)

        self.btn_export = ttk.Button(
            btn_frame, text="📎 Exportar Adjuntos", bootstyle=SUCCESS,
            command=self._start_export
        )
        self.btn_export.pack(side=LEFT, padx=(0, 10))

        ttk.Button(
            btn_frame, text="Cerrar", bootstyle=SECONDARY,
            command=self.destroy
        ).pack(side=RIGHT)

    def _browse_dir(self):
        """Abre diálogo para seleccionar directorio."""
        directory = filedialog.askdirectory(
            title="Seleccionar directorio destino",
            initialdir=self.dir_var.get()
        )
        if directory:
            self.dir_var.set(directory)

    def _start_export(self):
        """Inicia la exportación en un thread."""
        output_dir = self.dir_var.get().strip()
        if not output_dir:
            messagebox.showwarning("Atención", "Selecciona un directorio destino.", parent=self)
            return

        organize = self.organize_var.get()
        types_str = self.types_var.get().strip()
        file_types = None
        if types_str:
            file_types = [t.strip() for t in types_str.split(",") if t.strip()]

        self.btn_export.configure(state=DISABLED)
        self.progress_var.set(0)
        self.progress_label.set("Exportando...")

        thread = threading.Thread(
            target=self._export_thread,
            args=(output_dir, organize, file_types),
            daemon=True,
        )
        thread.start()

    def _export_thread(self, output_dir, organize, file_types):
        """Thread de exportación."""
        try:
            stats = export_attachments(
                results=self.results,
                output_dir=output_dir,
                organize_by=organize,
                file_types=file_types,
                progress_callback=self._on_progress,
            )
            self.after(0, self._on_complete, stats)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_progress(self, current, total, message):
        """Callback de progreso."""
        pct = (current / total * 100) if total > 0 else 0
        self.after(0, lambda: self.progress_var.set(pct))
        self.after(0, lambda: self.progress_label.set(message))

    def _on_complete(self, stats):
        """Exportación completada."""
        self.progress_var.set(100)
        self.progress_label.set("✓ Exportación completada")
        self.btn_export.configure(state=NORMAL)

        msg = (
            f"📧 Correos procesados: {stats['emails_with_attachments']}\n"
            f"📎 Adjuntos exportados: {stats['exported']}\n"
        )
        if stats["skipped"]:
            msg += f"⏭️ Omitidos: {stats['skipped']}\n"
        if stats["errors"]:
            msg += f"❌ Errores: {stats['errors']}\n"
        msg += f"\n📁 Directorio: {self.dir_var.get()}"

        messagebox.showinfo("Exportación Completada", msg, parent=self)

    def _on_error(self, error_msg):
        """Error en la exportación."""
        self.progress_label.set(f"❌ Error: {error_msg}")
        self.btn_export.configure(state=NORMAL)
        messagebox.showerror("Error", f"Error durante la exportación:\n{error_msg}", parent=self)
