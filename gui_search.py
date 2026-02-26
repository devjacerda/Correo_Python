"""
Frame de búsqueda de correos con filtros, tabla de resultados y exportación.
Incluye búsqueda avanzada y búsqueda rápida en sub-tabs internos.
"""

import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tableview import Tableview
from tkinter import filedialog, messagebox

from search import EmailSearch
from reports import export_to_excel, export_to_csv, generate_summary
from gui_detail import EmailDetailDialog
from gui_attachments import AttachmentsDialog


class SearchFrame(ttk.Frame):
    """Frame principal de búsqueda con sub-tabs para búsqueda avanzada y rápida."""

    def __init__(self, parent, searcher: EmailSearch):
        super().__init__(parent, padding=10)
        self.searcher = searcher
        self.last_results = []

        self._build_ui()

    def _build_ui(self):
        """Construye la interfaz completa."""
        # === Sub-notebook para búsqueda avanzada / rápida ===
        self.sub_notebook = ttk.Notebook(self)
        self.sub_notebook.pack(fill=X, pady=(0, 10))

        # --- Tab Búsqueda Avanzada ---
        adv_frame = ttk.Frame(self.sub_notebook, padding=10)
        self.sub_notebook.add(adv_frame, text="🔍 Búsqueda con Filtros")
        self._build_advanced_filters(adv_frame)

        # --- Tab Búsqueda Rápida ---
        quick_frame = ttk.Frame(self.sub_notebook, padding=10)
        self.sub_notebook.add(quick_frame, text="⚡ Búsqueda Rápida")
        self._build_quick_search(quick_frame)

        # === Tabla de resultados ===
        self._build_results_table()

        # === Barra de acciones ===
        self._build_action_bar()

        # === Barra de estado ===
        self.status_var = ttk.StringVar(value="Listo. Ingresa los filtros y presiona Buscar.")
        ttk.Label(
            self, textvariable=self.status_var,
            font=("Segoe UI", 9), foreground="gray"
        ).pack(fill=X, pady=(5, 0))

    def _build_advanced_filters(self, parent):
        """Construye el panel de filtros avanzados."""
        # Row 1: Asunto y Remitente
        row1 = ttk.Frame(parent)
        row1.pack(fill=X, pady=(0, 5))

        ttk.Label(row1, text="Asunto:", width=10, anchor=E).pack(side=LEFT)
        self.subject_var = ttk.StringVar()
        ttk.Entry(row1, textvariable=self.subject_var, width=30).pack(side=LEFT, padx=(5, 15))

        ttk.Label(row1, text="Remitente:", width=10, anchor=E).pack(side=LEFT)
        self.sender_var = ttk.StringVar()
        ttk.Entry(row1, textvariable=self.sender_var, width=30).pack(side=LEFT, padx=(5, 0))

        # Row 2: Fechas
        row2 = ttk.Frame(parent)
        row2.pack(fill=X, pady=(0, 5))

        ttk.Label(row2, text="Desde:", width=10, anchor=E).pack(side=LEFT)
        self.date_from_var = ttk.StringVar()
        entry_from = ttk.Entry(row2, textvariable=self.date_from_var, width=14)
        entry_from.pack(side=LEFT, padx=(5, 5))
        ttk.Label(row2, text="DD-MM-YYYY", font=("Segoe UI", 8), foreground="gray").pack(side=LEFT, padx=(0, 15))

        ttk.Label(row2, text="Hasta:", width=10, anchor=E).pack(side=LEFT)
        self.date_to_var = ttk.StringVar()
        entry_to = ttk.Entry(row2, textvariable=self.date_to_var, width=14)
        entry_to.pack(side=LEFT, padx=(5, 5))
        ttk.Label(row2, text="DD-MM-YYYY", font=("Segoe UI", 8), foreground="gray").pack(side=LEFT)

        # Row 3: Opciones avanzadas
        row3 = ttk.Frame(parent)
        row3.pack(fill=X, pady=(0, 5))

        ttk.Label(row3, text="Carpeta:", width=10, anchor=E).pack(side=LEFT)
        self.folder_var = ttk.StringVar(value="inbox")
        folder_combo = ttk.Combobox(
            row3, textvariable=self.folder_var, width=12,
            values=["inbox", "sent", "drafts", "deleted", "junk", "outbox"],
            state="readonly"
        )
        folder_combo.pack(side=LEFT, padx=(5, 15))

        self.has_att_var = ttk.StringVar(value="todos")
        ttk.Label(row3, text="Adjuntos:", width=10, anchor=E).pack(side=LEFT)
        att_combo = ttk.Combobox(
            row3, textvariable=self.has_att_var, width=10,
            values=["todos", "sí", "no"], state="readonly"
        )
        att_combo.pack(side=LEFT, padx=(5, 15))

        ttk.Label(row3, text="Máx:", width=5, anchor=E).pack(side=LEFT)
        self.max_results_var = ttk.IntVar(value=100)
        ttk.Spinbox(
            row3, from_=10, to=1000, increment=50,
            textvariable=self.max_results_var, width=6
        ).pack(side=LEFT, padx=(5, 15))

        # Row 4: Cuerpo + Botón
        row4 = ttk.Frame(parent)
        row4.pack(fill=X)

        ttk.Label(row4, text="En cuerpo:", width=10, anchor=E).pack(side=LEFT)
        self.body_var = ttk.StringVar()
        ttk.Entry(row4, textvariable=self.body_var, width=30).pack(side=LEFT, padx=(5, 15))

        self.btn_search = ttk.Button(
            row4, text="🔍 Buscar", bootstyle=PRIMARY,
            command=self._do_advanced_search, width=15
        )
        self.btn_search.pack(side=RIGHT)

    def _build_quick_search(self, parent):
        """Construye el panel de búsqueda rápida."""
        frame = ttk.Frame(parent)
        frame.pack(fill=X)

        ttk.Label(frame, text="Término:", width=10, anchor=E).pack(side=LEFT)
        self.quick_term_var = ttk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.quick_term_var, width=40)
        entry.pack(side=LEFT, padx=(5, 15))
        entry.bind("<Return>", lambda e: self._do_quick_search())

        ttk.Label(frame, text="Buscar en:", anchor=E).pack(side=LEFT, padx=(0, 5))
        self.quick_scope_var = ttk.StringVar(value="subject")
        scope_combo = ttk.Combobox(
            frame, textvariable=self.quick_scope_var, width=10,
            values=["subject", "sender", "all"], state="readonly"
        )
        scope_combo.pack(side=LEFT, padx=(0, 15))

        self.btn_quick = ttk.Button(
            frame, text="⚡ Buscar", bootstyle=SUCCESS,
            command=self._do_quick_search, width=12
        )
        self.btn_quick.pack(side=RIGHT)

    def _build_results_table(self):
        """Construye la tabla de resultados con Treeview."""
        table_frame = ttk.Frame(self)
        table_frame.pack(fill=BOTH, expand=True)

        columns = (
            "num", "date", "time", "sender", "subject", "att", "importance"
        )
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings",
            selectmode="browse", height=15
        )

        # Configurar columnas
        col_config = {
            "num": ("#", 40, CENTER),
            "date": ("Fecha", 90, CENTER),
            "time": ("Hora", 75, CENTER),
            "sender": ("Remitente", 200, W),
            "subject": ("Asunto", 320, W),
            "att": ("📎", 35, CENTER),
            "importance": ("Importancia", 80, CENTER),
        }

        for col_id, (heading, width, anchor) in col_config.items():
            self.tree.heading(col_id, text=heading, command=lambda c=col_id: self._sort_column(c))
            self.tree.column(col_id, width=width, anchor=anchor, minwidth=30)

        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        v_scroll.pack(side=RIGHT, fill=Y)
        h_scroll.pack(side=BOTTOM, fill=X)

        # Doble clic para ver detalle
        self.tree.bind("<Double-1>", self._on_double_click)

        # Tag para importancia alta
        self.tree.tag_configure("alta", foreground="#e74c3c")

    def _build_action_bar(self):
        """Construye la barra de botones de acción."""
        bar = ttk.Frame(self, padding=(0, 8, 0, 0))
        bar.pack(fill=X)

        self.btn_excel = ttk.Button(
            bar, text="📊 Exportar Excel", bootstyle=(SUCCESS, OUTLINE),
            command=self._export_excel, state=DISABLED
        )
        self.btn_excel.pack(side=LEFT, padx=(0, 5))

        self.btn_csv = ttk.Button(
            bar, text="📋 Exportar CSV", bootstyle=(INFO, OUTLINE),
            command=self._export_csv, state=DISABLED
        )
        self.btn_csv.pack(side=LEFT, padx=(0, 5))

        self.btn_attachments = ttk.Button(
            bar, text="📎 Exportar Adjuntos", bootstyle=(WARNING, OUTLINE),
            command=self._export_attachments, state=DISABLED
        )
        self.btn_attachments.pack(side=LEFT, padx=(0, 5))

        self.btn_detail = ttk.Button(
            bar, text="📄 Ver Detalle", bootstyle=(SECONDARY, OUTLINE),
            command=self._view_detail, state=DISABLED
        )
        self.btn_detail.pack(side=LEFT, padx=(0, 5))

        self.btn_summary = ttk.Button(
            bar, text="📈 Resumen", bootstyle=(DARK, OUTLINE),
            command=self._show_summary, state=DISABLED
        )
        self.btn_summary.pack(side=LEFT)

        # Contador de resultados
        self.count_var = ttk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count_var, font=("Segoe UI", 10, "bold")).pack(side=RIGHT)

    # =========================
    # Acciones de búsqueda
    # =========================

    def _do_advanced_search(self):
        """Ejecuta la búsqueda avanzada."""
        subject = self.subject_var.get().strip() or None
        sender = self.sender_var.get().strip() or None
        date_from = self.date_from_var.get().strip() or None
        date_to = self.date_to_var.get().strip() or None
        folder = self.folder_var.get()
        body_contains = self.body_var.get().strip() or None
        max_results = self.max_results_var.get()

        has_att = self.has_att_var.get()
        has_attachments = None
        if has_att == "sí":
            has_attachments = True
        elif has_att == "no":
            has_attachments = False

        self._run_search(
            subject=subject, sender=sender, date_from=date_from,
            date_to=date_to, folder=folder, has_attachments=has_attachments,
            body_contains=body_contains, max_results=max_results,
        )

    def _do_quick_search(self):
        """Ejecuta la búsqueda rápida."""
        term = self.quick_term_var.get().strip()
        if not term:
            messagebox.showwarning("Atención", "Ingresa un término de búsqueda.", parent=self)
            return

        scope = self.quick_scope_var.get()
        if scope == "subject":
            self._run_search(subject=term, max_results=50)
        elif scope == "sender":
            self._run_search(sender=term, max_results=50)
        else:
            # Buscar en todo
            self._run_search(subject=term, max_results=50, _merge_sender=term)

    def _run_search(self, _merge_sender=None, **kwargs):
        """Ejecuta la búsqueda en un thread separado."""
        self._set_searching(True)
        self.status_var.set("🔍 Buscando correos...")
        self.tree.delete(*self.tree.get_children())

        thread = threading.Thread(
            target=self._search_thread, args=(_merge_sender,),
            kwargs=kwargs, daemon=True
        )
        thread.start()

    def _search_thread(self, merge_sender=None, **kwargs):
        """Thread de búsqueda."""
        try:
            results = self.searcher.search(
                progress_callback=self._on_search_progress,
                **kwargs
            )

            # Si merge_sender, combinar con búsqueda por remitente
            if merge_sender:
                results_sender = self.searcher.search(sender=merge_sender, max_results=50)
                seen = {(r["subject"], r["date"], r["time"]) for r in results}
                for r in results_sender:
                    key = (r["subject"], r["date"], r["time"])
                    if key not in seen:
                        results.append(r)
                        seen.add(key)

            self.after(0, self._on_search_complete, results)

        except Exception as e:
            self.after(0, self._on_search_error, str(e))

    def _on_search_progress(self, current, message):
        """Callback de progreso de búsqueda."""
        self.after(0, lambda: self.status_var.set(f"🔍 {message}"))

    def _on_search_complete(self, results):
        """Búsqueda completada."""
        self.last_results = results
        self._populate_table(results)
        self._set_searching(False)

        count = len(results)
        self.count_var.set(f"{count} correo{'s' if count != 1 else ''}")
        if count > 0:
            self.status_var.set(f"✓ Búsqueda completada: {count} correos encontrados.")
            self._set_buttons_state(NORMAL)
        else:
            self.status_var.set("No se encontraron correos con esos filtros.")
            self._set_buttons_state(DISABLED)

    def _on_search_error(self, error_msg):
        """Error en la búsqueda."""
        self._set_searching(False)
        self.status_var.set(f"❌ Error: {error_msg}")
        messagebox.showerror("Error de Búsqueda", error_msg, parent=self)

    # =========================
    # Tabla de resultados
    # =========================

    def _populate_table(self, results):
        """Puebla la tabla con los resultados."""
        self.tree.delete(*self.tree.get_children())

        for i, email in enumerate(results, 1):
            att_icon = "✓" if email.get("has_attachments") else ""
            importance = email.get("importance", "Normal")
            tags = ("alta",) if importance == "Alta" else ()

            self.tree.insert("", END, values=(
                i,
                email.get("date", ""),
                email.get("time", ""),
                self._truncate(email.get("sender_name", ""), 30),
                self._truncate(email.get("subject", ""), 50),
                att_icon,
                importance,
            ), tags=tags)

    def _sort_column(self, col):
        """Ordena la tabla por columna."""
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        try:
            items.sort(key=lambda t: int(t[0]))
        except ValueError:
            items.sort(key=lambda t: t[0].lower())

        for index, (val, k) in enumerate(items):
            self.tree.move(k, "", index)

    def _on_double_click(self, event):
        """Doble clic en un resultado."""
        self._view_detail()

    # =========================
    # Acciones de exportación
    # =========================

    def _export_excel(self):
        """Exporta resultados a Excel."""
        if not self.last_results:
            return

        filepath = filedialog.asksaveasfilename(
            title="Guardar reporte Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"busqueda_outlook.xlsx",
            parent=self,
        )
        if not filepath:
            return

        try:
            result_path = export_to_excel(self.last_results, filepath)
            messagebox.showinfo(
                "Exportación Exitosa",
                f"Reporte guardado en:\n{result_path}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{e}", parent=self)

    def _export_csv(self):
        """Exporta resultados a CSV."""
        if not self.last_results:
            return

        filepath = filedialog.asksaveasfilename(
            title="Guardar reporte CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"busqueda_outlook.csv",
            parent=self,
        )
        if not filepath:
            return

        try:
            result_path = export_to_csv(self.last_results, filepath)
            messagebox.showinfo(
                "Exportación Exitosa",
                f"Reporte guardado en:\n{result_path}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{e}", parent=self)

    def _export_attachments(self):
        """Abre el diálogo de exportación de adjuntos."""
        if not self.last_results:
            return
        AttachmentsDialog(self.winfo_toplevel(), self.last_results)

    def _view_detail(self):
        """Muestra el detalle del correo seleccionado."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Info", "Selecciona un correo de la tabla.", parent=self)
            return

        item = selection[0]
        idx = int(self.tree.set(item, "num")) - 1
        if 0 <= idx < len(self.last_results):
            EmailDetailDialog(self.winfo_toplevel(), self.last_results[idx])

    def _show_summary(self):
        """Muestra el resumen estadístico."""
        if not self.last_results:
            return

        summary = generate_summary(self.last_results)
        if not summary:
            return

        # Construir texto del resumen
        lines = [
            f"📊 Resumen de Búsqueda\n",
            f"{'─' * 35}",
            f"Total correos:        {summary['total']}",
            f"Con adjuntos:         {summary['with_attachments']} ({summary['pct_attachments']}%)",
            f"Total adjuntos:       {summary['total_attachments']}",
            f"Fecha más antigua:    {summary['date_min']}",
            f"Fecha más reciente:   {summary['date_max']}",
            f"\n👤 Top Remitentes:",
            f"{'─' * 35}",
        ]
        for name, count in summary.get("top_senders", []):
            lines.append(f"  {self._truncate(name, 25):28s} {count}")

        messagebox.showinfo(
            "Resumen Estadístico",
            "\n".join(lines),
            parent=self,
        )

    # =========================
    # Helpers
    # =========================

    def _set_searching(self, searching: bool):
        """Habilita/deshabilita controles durante búsqueda."""
        state = DISABLED if searching else NORMAL
        self.btn_search.configure(state=state)
        self.btn_quick.configure(state=state)

    def _set_buttons_state(self, state):
        """Establece el estado de los botones de acción."""
        for btn in (self.btn_excel, self.btn_csv, self.btn_attachments, self.btn_detail, self.btn_summary):
            btn.configure(state=state)

    @staticmethod
    def _truncate(text: str, max_len: int) -> str:
        """Trunca texto."""
        if not text or len(text) <= max_len:
            return text or ""
        return text[:max_len - 3] + "..."
