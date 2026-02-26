"""
Diálogo de exportación de archivos adjuntos.
Usa el OutlookWorker para la exportación (operaciones COM en su thread).
"""

import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox


class AttachmentsDialog(ttk.Toplevel):
    """Diálogo para exportar adjuntos desde los resultados de búsqueda."""

    def __init__(self, parent, worker):
        super().__init__(parent)
        self.title("📎 Exportar Adjuntos")
        self.geometry("500x360")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.worker = worker
        n_att = len([r for r in worker.last_results if r.get("has_attachments")])
        n_total = len(worker.last_results)
        self._build_ui(n_att, n_total)

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 500) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 360) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _build_ui(self, n_att, n_total):
        m = ttk.Frame(self, padding=20)
        m.pack(fill=BOTH, expand=True)

        ttk.Label(m, text=f"{n_att} correos con adjuntos de {n_total} totales.",
                  font=("Segoe UI", 10)).pack(anchor=W, pady=(0, 12))

        # Directorio
        ttk.Label(m, text="Directorio destino:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        df = ttk.Frame(m)
        df.pack(fill=X, pady=(2, 8))
        self.v_dir = ttk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "adjuntos_outlook"))
        ttk.Entry(df, textvariable=self.v_dir, font=("Segoe UI", 9)).pack(side=LEFT, fill=X, expand=True, padx=(0, 4))
        ttk.Button(df, text="📂", bootstyle=OUTLINE, command=self._browse, width=3).pack(side=RIGHT)

        # Organización
        ttk.Label(m, text="Organizar por:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        of = ttk.Frame(m)
        of.pack(fill=X, pady=(2, 8))
        self.v_org = ttk.StringVar(value="flat")
        for txt, val in [("Flat", "flat"), ("Remitente", "sender"), ("Fecha", "date"), ("Asunto", "subject")]:
            ttk.Radiobutton(of, text=txt, variable=self.v_org, value=val).pack(side=LEFT, padx=(0, 12))

        # Tipos
        ttk.Label(m, text="Tipos (vacío=todos):", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        self.v_types = ttk.StringVar()
        ttk.Entry(m, textvariable=self.v_types, font=("Segoe UI", 9)).pack(fill=X, pady=(2, 3))
        ttk.Label(m, text="Ej: .pdf, .xlsx, .docx", font=("Segoe UI", 8), foreground="gray").pack(anchor=W, pady=(0, 8))

        # Progreso
        self.v_prog = ttk.DoubleVar()
        self.prog_bar = ttk.Progressbar(m, variable=self.v_prog, bootstyle=SUCCESS)
        self.prog_bar.pack(fill=X, pady=(0, 3))
        self.v_status = ttk.StringVar()
        ttk.Label(m, textvariable=self.v_status, font=("Segoe UI", 9)).pack(anchor=W, pady=(0, 8))

        # Botones
        bf = ttk.Frame(m)
        bf.pack(fill=X)
        self.btn_export = ttk.Button(bf, text="📎 Exportar", bootstyle=SUCCESS, command=self._do_export)
        self.btn_export.pack(side=LEFT, padx=(0, 8))
        ttk.Button(bf, text="Cerrar", bootstyle=SECONDARY, command=self.destroy).pack(side=RIGHT)

    def _browse(self):
        d = filedialog.askdirectory(title="Directorio destino", initialdir=self.v_dir.get())
        if d: self.v_dir.set(d)

    def _do_export(self):
        out = self.v_dir.get().strip()
        if not out:
            messagebox.showwarning("Atención", "Selecciona un directorio.", parent=self)
            return

        types_str = self.v_types.get().strip()
        file_types = [t.strip() for t in types_str.split(",") if t.strip()] if types_str else None

        self.btn_export.configure(state=DISABLED)
        self.v_status.set("Exportando...")
        self.v_prog.set(0)

        # Registrar callback de progreso en el app
        app = self.winfo_toplevel()
        original_progress = app._on_attachment_progress

        def progress_cb(cur, total, msg):
            pct = (cur / total * 100) if total > 0 else 0
            self.v_prog.set(pct)
            self.v_status.set(msg)

        app._on_attachment_progress = progress_cb

        self.worker.submit(
            "export_attachments",
            {"output_dir": out, "organize_by": self.v_org.get(), "file_types": file_types},
            lambda stats: self._on_done(stats, app, original_progress),
            lambda err: self._on_err(err, app, original_progress),
        )

    def _on_done(self, stats, app, orig_cb):
        app._on_attachment_progress = orig_cb
        self.v_prog.set(100)
        self.v_status.set("✓ Completado")
        self.btn_export.configure(state=NORMAL)

        msg = (f"📧 Correos: {stats['emails_with_attachments']}\n"
               f"📎 Exportados: {stats['exported']}\n")
        if stats["skipped"]: msg += f"⏭️ Omitidos: {stats['skipped']}\n"
        if stats["errors"]: msg += f"❌ Errores: {stats['errors']}\n"
        msg += f"\n📁 {self.v_dir.get()}"
        messagebox.showinfo("Completado", msg, parent=self)

    def _on_err(self, err, app, orig_cb):
        app._on_attachment_progress = orig_cb
        self.v_status.set("❌ Error")
        self.btn_export.configure(state=NORMAL)
        messagebox.showerror("Error", err, parent=self)
