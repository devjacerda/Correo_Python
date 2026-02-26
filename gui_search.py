"""
Frame de búsqueda de correos con filtros, tabla de resultados y acciones de exportación.
Toda la comunicación con Outlook se hace a través del OutlookWorker.
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

from reports import export_to_excel, export_to_csv, generate_summary
from gui_detail import EmailDetailDialog
from gui_attachments import AttachmentsDialog


class SearchFrame(ttk.Frame):
    """Frame de búsqueda con filtros avanzados, búsqueda rápida y resultados."""

    def __init__(self, parent, worker):
        super().__init__(parent, padding=10)
        self.worker = worker
        self.last_results = []  # resultados limpios (sin COM refs)

        self._build_ui()

    def _build_ui(self):
        # === Sub-tabs ===
        notebook = ttk.Notebook(self)
        notebook.pack(fill=X, pady=(0, 8))

        adv = ttk.Frame(notebook, padding=8)
        notebook.add(adv, text="🔍 Búsqueda con Filtros")
        self._build_filters(adv)

        quick = ttk.Frame(notebook, padding=8)
        notebook.add(quick, text="⚡ Búsqueda Rápida")
        self._build_quick(quick)

        # === Tabla ===
        self._build_table()

        # === Botones ===
        self._build_buttons()

        # === Estado ===
        self.status_var = ttk.StringVar(value="Listo. Ingresa filtros y presiona Buscar.")
        ttk.Label(self, textvariable=self.status_var, font=("Segoe UI", 9), foreground="gray").pack(fill=X, pady=(4, 0))

    # ──────────── Filtros avanzados ────────────

    def _build_filters(self, parent):
        # Row 1
        r1 = ttk.Frame(parent)
        r1.pack(fill=X, pady=2)
        ttk.Label(r1, text="Asunto:", width=10, anchor=E).pack(side=LEFT)
        self.v_subject = ttk.StringVar()
        ttk.Entry(r1, textvariable=self.v_subject, width=30).pack(side=LEFT, padx=(4, 12))
        ttk.Label(r1, text="Remitente:", width=10, anchor=E).pack(side=LEFT)
        self.v_sender = ttk.StringVar()
        ttk.Entry(r1, textvariable=self.v_sender, width=30).pack(side=LEFT, padx=4)

        # Row 2
        r2 = ttk.Frame(parent)
        r2.pack(fill=X, pady=2)
        ttk.Label(r2, text="Desde:", width=10, anchor=E).pack(side=LEFT)
        self.v_from = ttk.StringVar()
        ttk.Entry(r2, textvariable=self.v_from, width=12).pack(side=LEFT, padx=4)
        ttk.Label(r2, text="DD-MM-YYYY", font=("Segoe UI", 8), foreground="gray").pack(side=LEFT, padx=(0, 12))
        ttk.Label(r2, text="Hasta:", width=10, anchor=E).pack(side=LEFT)
        self.v_to = ttk.StringVar()
        ttk.Entry(r2, textvariable=self.v_to, width=12).pack(side=LEFT, padx=4)
        ttk.Label(r2, text="DD-MM-YYYY", font=("Segoe UI", 8), foreground="gray").pack(side=LEFT)

        # Row 3
        r3 = ttk.Frame(parent)
        r3.pack(fill=X, pady=2)
        ttk.Label(r3, text="Carpeta:", width=10, anchor=E).pack(side=LEFT)
        self.v_folder = ttk.StringVar(value="inbox")
        ttk.Combobox(r3, textvariable=self.v_folder, width=10, state="readonly",
                     values=["inbox", "sent", "drafts", "deleted", "junk", "outbox"]).pack(side=LEFT, padx=(4, 12))
        ttk.Label(r3, text="Adjuntos:", anchor=E).pack(side=LEFT)
        self.v_att = ttk.StringVar(value="todos")
        ttk.Combobox(r3, textvariable=self.v_att, width=8, state="readonly",
                     values=["todos", "sí", "no"]).pack(side=LEFT, padx=(4, 12))
        ttk.Label(r3, text="Máx:", anchor=E).pack(side=LEFT)
        self.v_max = ttk.IntVar(value=100)
        ttk.Spinbox(r3, from_=10, to=1000, increment=50, textvariable=self.v_max, width=6).pack(side=LEFT, padx=4)

        # Row 4
        r4 = ttk.Frame(parent)
        r4.pack(fill=X, pady=2)
        ttk.Label(r4, text="En cuerpo:", width=10, anchor=E).pack(side=LEFT)
        self.v_body = ttk.StringVar()
        ttk.Entry(r4, textvariable=self.v_body, width=30).pack(side=LEFT, padx=4)

        self.btn_search = ttk.Button(r4, text="🔍 Buscar", bootstyle=PRIMARY,
                                     command=self._search_advanced, width=14)
        self.btn_search.pack(side=RIGHT)

    # ──────────── Búsqueda rápida ────────────

    def _build_quick(self, parent):
        f = ttk.Frame(parent)
        f.pack(fill=X)
        ttk.Label(f, text="Término:", width=10, anchor=E).pack(side=LEFT)
        self.v_quick = ttk.StringVar()
        e = ttk.Entry(f, textvariable=self.v_quick, width=40)
        e.pack(side=LEFT, padx=4)
        e.bind("<Return>", lambda _: self._search_quick())

        ttk.Label(f, text="En:", anchor=E).pack(side=LEFT, padx=(8, 4))
        self.v_scope = ttk.StringVar(value="subject")
        ttk.Combobox(f, textvariable=self.v_scope, width=8, state="readonly",
                     values=["subject", "sender", "all"]).pack(side=LEFT, padx=(0, 8))

        self.btn_quick = ttk.Button(f, text="⚡ Buscar", bootstyle=SUCCESS,
                                    command=self._search_quick, width=12)
        self.btn_quick.pack(side=RIGHT)

    # ──────────── Tabla de resultados ────────────

    def _build_table(self):
        tf = ttk.Frame(self)
        tf.pack(fill=BOTH, expand=True)

        cols = ("num", "date", "time", "sender", "subject", "att", "importance")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse", height=16)

        cfg = {
            "num": ("#", 40, CENTER), "date": ("Fecha", 90, CENTER),
            "time": ("Hora", 70, CENTER), "sender": ("Remitente", 200, W),
            "subject": ("Asunto", 320, W), "att": ("📎", 35, CENTER),
            "importance": ("Imp.", 60, CENTER),
        }
        for cid, (hd, w, anch) in cfg.items():
            self.tree.heading(cid, text=hd, command=lambda c=cid: self._sort(c))
            self.tree.column(cid, width=w, anchor=anch, minwidth=30)

        vsb = ttk.Scrollbar(tf, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky=NSEW)
        vsb.grid(row=0, column=1, sticky=NS)
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda _: self._view_detail())
        self.tree.tag_configure("alta", foreground="#e74c3c")

    # ──────────── Botones de acción ────────────

    def _build_buttons(self):
        bar = ttk.Frame(self, padding=(0, 6, 0, 0))
        bar.pack(fill=X)

        self.btn_excel = ttk.Button(bar, text="📊 Exportar Excel", bootstyle=(SUCCESS, OUTLINE),
                                    command=self._export_excel, state=DISABLED)
        self.btn_excel.pack(side=LEFT, padx=(0, 4))

        self.btn_csv = ttk.Button(bar, text="📋 Exportar CSV", bootstyle=(INFO, OUTLINE),
                                  command=self._export_csv, state=DISABLED)
        self.btn_csv.pack(side=LEFT, padx=(0, 4))

        self.btn_att = ttk.Button(bar, text="📎 Exportar Adjuntos", bootstyle=(WARNING, OUTLINE),
                                  command=self._export_attachments, state=DISABLED)
        self.btn_att.pack(side=LEFT, padx=(0, 4))

        self.btn_det = ttk.Button(bar, text="📄 Detalle", bootstyle=(SECONDARY, OUTLINE),
                                  command=self._view_detail, state=DISABLED)
        self.btn_det.pack(side=LEFT, padx=(0, 4))

        self.btn_sum = ttk.Button(bar, text="📈 Resumen", bootstyle=(DARK, OUTLINE),
                                  command=self._show_summary, state=DISABLED)
        self.btn_sum.pack(side=LEFT)

        self.v_count = ttk.StringVar()
        ttk.Label(bar, textvariable=self.v_count, font=("Segoe UI", 10, "bold")).pack(side=RIGHT)

    # ══════════════ Búsquedas ══════════════

    def _search_advanced(self):
        kwargs = {}
        s = self.v_subject.get().strip()
        if s: kwargs["subject"] = s
        s = self.v_sender.get().strip()
        if s: kwargs["sender"] = s
        s = self.v_from.get().strip()
        if s: kwargs["date_from"] = s
        s = self.v_to.get().strip()
        if s: kwargs["date_to"] = s
        kwargs["folder"] = self.v_folder.get()
        att = self.v_att.get()
        if att == "sí": kwargs["has_attachments"] = True
        elif att == "no": kwargs["has_attachments"] = False
        s = self.v_body.get().strip()
        if s: kwargs["body_contains"] = s
        kwargs["max_results"] = self.v_max.get()

        self._submit_search("search", kwargs)

    def _search_quick(self):
        term = self.v_quick.get().strip()
        if not term:
            messagebox.showwarning("Atención", "Ingresa un término.", parent=self)
            return

        scope = self.v_scope.get()
        if scope == "subject":
            self._submit_search("search", {"subject": term, "max_results": 50})
        elif scope == "sender":
            self._submit_search("search", {"sender": term, "max_results": 50})
        else:
            self._submit_search("quick_search_all", {"term": term, "max_results": 50})

    def _submit_search(self, task_name, kwargs):
        self._set_searching(True)
        self.status_var.set("🔍 Buscando correos...")
        self.tree.delete(*self.tree.get_children())
        self.worker.submit(task_name, kwargs, self._on_results, self._on_error)

    def _on_results(self, clean_results):
        """Callback con resultados limpios del worker."""
        self.last_results = clean_results
        self._fill_table(clean_results)
        self._set_searching(False)

        n = len(clean_results)
        self.v_count.set(f"{n} correo{'s' if n != 1 else ''}")
        if n > 0:
            self.status_var.set(f"✓ {n} correos encontrados.")
            self._set_action_buttons(NORMAL)
        else:
            self.status_var.set("No se encontraron correos con esos filtros.")
            self._set_action_buttons(DISABLED)

    def _on_error(self, msg):
        self._set_searching(False)
        self.status_var.set(f"❌ Error")
        messagebox.showerror("Error de Búsqueda", msg, parent=self)

    # ══════════════ Tabla ══════════════

    def _fill_table(self, results):
        self.tree.delete(*self.tree.get_children())
        for i, e in enumerate(results, 1):
            att = "✓" if e.get("has_attachments") else ""
            imp = e.get("importance", "Normal")
            tags = ("alta",) if imp == "Alta" else ()
            self.tree.insert("", END, values=(
                i, e.get("date", ""), e.get("time", ""),
                _trunc(e.get("sender_name", ""), 30),
                _trunc(e.get("subject", ""), 50),
                att, imp,
            ), tags=tags)

    def _sort(self, col):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        try:
            items.sort(key=lambda t: int(t[0]))
        except ValueError:
            items.sort(key=lambda t: t[0].lower())
        for idx, (_, k) in enumerate(items):
            self.tree.move(k, "", idx)

    # ══════════════ Acciones ══════════════

    def _export_excel(self):
        if not self.last_results: return
        fp = filedialog.asksaveasfilename(
            title="Guardar Excel", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")], initialfile="busqueda_outlook.xlsx", parent=self)
        if not fp: return
        try:
            path = export_to_excel(self.last_results, fp)
            messagebox.showinfo("Exportado", f"Guardado en:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _export_csv(self):
        if not self.last_results: return
        fp = filedialog.asksaveasfilename(
            title="Guardar CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv")], initialfile="busqueda_outlook.csv", parent=self)
        if not fp: return
        try:
            path = export_to_csv(self.last_results, fp)
            messagebox.showinfo("Exportado", f"Guardado en:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _export_attachments(self):
        if not self.last_results: return
        AttachmentsDialog(self.winfo_toplevel(), self.worker)

    def _view_detail(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Selecciona un correo.", parent=self)
            return
        idx = int(self.tree.set(sel[0], "num")) - 1
        if 0 <= idx < len(self.last_results):
            EmailDetailDialog(self.winfo_toplevel(), self.last_results[idx])

    def _show_summary(self):
        if not self.last_results: return
        s = generate_summary(self.last_results)
        if not s: return
        lines = [
            "📊 Resumen de Búsqueda\n" + "─" * 35,
            f"Total correos:        {s['total']}",
            f"Con adjuntos:         {s['with_attachments']} ({s['pct_attachments']}%)",
            f"Total adjuntos:       {s['total_attachments']}",
            f"Fecha más antigua:    {s['date_min']}",
            f"Fecha más reciente:   {s['date_max']}",
            "\n👤 Top Remitentes:\n" + "─" * 35,
        ]
        for name, cnt in s.get("top_senders", []):
            lines.append(f"  {_trunc(name, 25):28s} {cnt}")
        messagebox.showinfo("Resumen", "\n".join(lines), parent=self)

    # ══════════════ Helpers ══════════════

    def _set_searching(self, busy):
        st = DISABLED if busy else NORMAL
        self.btn_search.configure(state=st)
        self.btn_quick.configure(state=st)

    def _set_action_buttons(self, state):
        for b in (self.btn_excel, self.btn_csv, self.btn_att, self.btn_det, self.btn_sum):
            b.configure(state=state)


def _trunc(t, n):
    if not t or len(t) <= n: return t or ""
    return t[:n - 3] + "..."
