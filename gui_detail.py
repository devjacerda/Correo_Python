"""
Diálogo de detalle de un correo electrónico.
Muestra toda la información del correo en una ventana separada.
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolledtext import ScrolledText


class EmailDetailDialog(ttk.Toplevel):
    """Ventana de detalle de un correo."""

    def __init__(self, parent, email_data: dict):
        super().__init__(parent)
        self.title(f"📧 {email_data.get('subject', 'Sin asunto')}")
        self.geometry("700x550")
        self.resizable(True, True)
        self.transient(parent)

        self._build_ui(email_data)

        # Centrar respecto al padre
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 700) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 550) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _build_ui(self, data: dict):
        """Construye la interfaz del detalle."""
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        # --- Encabezado ---
        header_frame = ttk.LabelFrame(main_frame, text="Información del Correo", padding=10)
        header_frame.pack(fill=X, pady=(0, 10))

        fields = [
            ("De:", f"{data.get('sender_name', 'N/A')} <{data.get('sender_email', 'N/A')}>"),
            ("Para:", data.get("to", "N/A")),
            ("CC:", data.get("cc", "N/A") or "—"),
            ("Fecha:", f"{data.get('date', 'N/A')}  {data.get('time', '')}"),
            ("Importancia:", data.get("importance", "Normal")),
            ("Categorías:", data.get("categories", "") or "—"),
            ("Tamaño:", f"{data.get('size_kb', 0)} KB"),
        ]

        for row, (label, value) in enumerate(fields):
            lbl = ttk.Label(header_frame, text=label, font=("Segoe UI", 10, "bold"), width=12, anchor=E)
            lbl.grid(row=row, column=0, sticky=E, padx=(0, 8), pady=2)

            val = ttk.Label(header_frame, text=value, font=("Segoe UI", 10), wraplength=500, anchor=W)
            val.grid(row=row, column=1, sticky=W, pady=2)

        # --- Adjuntos ---
        if data.get("has_attachments"):
            att_frame = ttk.LabelFrame(main_frame, text=f"📎 Adjuntos ({data.get('attachment_count', 0)})", padding=10)
            att_frame.pack(fill=X, pady=(0, 10))

            att_names = data.get("attachment_names", [])
            for i, name in enumerate(att_names):
                icon = "📄"
                ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                if ext in ("pdf",):
                    icon = "📕"
                elif ext in ("xlsx", "xls", "csv"):
                    icon = "📊"
                elif ext in ("png", "jpg", "jpeg", "gif", "bmp"):
                    icon = "🖼️"
                elif ext in ("zip", "rar", "7z"):
                    icon = "📦"
                elif ext in ("doc", "docx"):
                    icon = "📝"

                ttk.Label(
                    att_frame, text=f"  {icon} {name}", font=("Segoe UI", 9)
                ).pack(anchor=W)

        # --- Cuerpo ---
        body_frame = ttk.LabelFrame(main_frame, text="Vista Previa", padding=10)
        body_frame.pack(fill=BOTH, expand=True)

        body_text = ScrolledText(body_frame, font=("Consolas", 9), wrap="word")
        body_text.pack(fill=BOTH, expand=True)
        body_text.insert("1.0", data.get("body_preview", "Sin contenido"))
        body_text.configure(state="disabled")

        # --- Botón cerrar ---
        ttk.Button(
            main_frame, text="Cerrar", bootstyle=SECONDARY, command=self.destroy
        ).pack(pady=(10, 0))
