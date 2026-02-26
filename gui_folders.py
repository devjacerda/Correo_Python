"""
Frame de visualización de carpetas del buzón.
Usa el OutlookWorker para cargar carpetas (COM en su thread).
"""

import ttkbootstrap as ttk
from ttkbootstrap.constants import *


class FoldersFrame(ttk.Frame):
    """Frame con el árbol de carpetas del buzón."""

    def __init__(self, parent, worker):
        super().__init__(parent, padding=10)
        self.worker = worker
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=X, pady=(0, 8))

        ttk.Label(top, text="📁 Carpetas del Buzón", font=("Segoe UI", 14, "bold")).pack(side=LEFT)
        self.btn_ref = ttk.Button(top, text="🔄 Actualizar", bootstyle=INFO, command=self._load)
        self.btn_ref.pack(side=RIGHT)

        tf = ttk.Frame(self)
        tf.pack(fill=BOTH, expand=True)

        self.tree = ttk.Treeview(tf, columns=("items",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Carpeta", anchor=W)
        self.tree.heading("items", text="Items", anchor=E)
        self.tree.column("#0", width=400, stretch=True)
        self.tree.column("items", width=80, anchor=E, stretch=False)

        vsb = ttk.Scrollbar(tf, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        vsb.pack(side=RIGHT, fill=Y)

        self.v_status = ttk.StringVar(value="Presiona 'Actualizar' para cargar.")
        ttk.Label(self, textvariable=self.v_status, font=("Segoe UI", 9)).pack(fill=X, pady=(4, 0))

    def _load(self):
        self.btn_ref.configure(state=DISABLED)
        self.v_status.set("Cargando carpetas...")
        self.tree.delete(*self.tree.get_children())
        self.worker.submit("list_folders", {"max_depth": 2}, self._on_data, self._on_err)

    def _on_data(self, folders):
        parent_map = {-1: ""}
        for name, path, count, indent in folders:
            pid = parent_map.get(indent - 1, "")
            icon = "📁" if indent == 0 else "📂"
            nid = self.tree.insert(pid, END, text=f" {icon} {name}",
                                   values=(str(count) if count else "",), open=(indent == 0))
            parent_map[indent] = nid

        self.v_status.set(f"✓ {len(folders)} carpetas cargadas")
        self.btn_ref.configure(state=NORMAL)

    def _on_err(self, msg):
        self.v_status.set(f"❌ Error: {msg}")
        self.btn_ref.configure(state=NORMAL)
