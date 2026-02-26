"""
Frame de visualización de carpetas del buzón de Outlook.
Muestra la estructura jerárquica de carpetas en un Treeview.
"""

import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


class FoldersFrame(ttk.Frame):
    """Frame con el árbol de carpetas del buzón."""

    def __init__(self, parent, outlook_client):
        super().__init__(parent, padding=10)
        self.client = outlook_client
        self._build_ui()

    def _build_ui(self):
        """Construye la interfaz."""
        # --- Barra superior ---
        top_bar = ttk.Frame(self)
        top_bar.pack(fill=X, pady=(0, 10))

        ttk.Label(
            top_bar, text="📁 Carpetas del Buzón",
            font=("Segoe UI", 14, "bold")
        ).pack(side=LEFT)

        self.btn_refresh = ttk.Button(
            top_bar, text="🔄 Actualizar", bootstyle=INFO,
            command=self._load_folders
        )
        self.btn_refresh.pack(side=RIGHT)

        # --- Treeview ---
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=BOTH, expand=True)

        columns = ("items",)
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="tree headings",
            selectmode="browse"
        )
        self.tree.heading("#0", text="Carpeta", anchor=W)
        self.tree.heading("items", text="Items", anchor=E)
        self.tree.column("#0", width=400, stretch=True)
        self.tree.column("items", width=80, anchor=E, stretch=False)

        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # --- Status ---
        self.status_var = ttk.StringVar(value="Presiona 'Actualizar' para cargar las carpetas")
        ttk.Label(self, textvariable=self.status_var, font=("Segoe UI", 9)).pack(fill=X, pady=(5, 0))

    def _load_folders(self):
        """Carga las carpetas en un thread separado."""
        self.btn_refresh.configure(state=DISABLED)
        self.status_var.set("Cargando carpetas...")
        self.tree.delete(*self.tree.get_children())

        thread = threading.Thread(target=self._load_folders_thread, daemon=True)
        thread.start()

    def _load_folders_thread(self):
        """Thread que carga las carpetas."""
        try:
            folders = self.client.list_folders(max_depth=2)
            self.after(0, self._populate_tree, folders)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _populate_tree(self, folders):
        """Puebla el Treeview con los datos de carpetas."""
        self.tree.delete(*self.tree.get_children())

        # Mapeo de indent -> último nodo padre insertado
        parent_map = {-1: ""}  # root

        for name, path, count, indent in folders:
            parent_indent = indent - 1
            parent_id = parent_map.get(parent_indent, "")

            icon = "📁" if indent == 0 else "📂"
            node_id = self.tree.insert(
                parent_id, END,
                text=f" {icon} {name}",
                values=(str(count) if count else "",),
                open=(indent == 0)
            )
            parent_map[indent] = node_id

        total = len(folders)
        self.status_var.set(f"✓ {total} carpetas cargadas")
        self.btn_refresh.configure(state=NORMAL)

    def _on_error(self, error_msg):
        """Maneja errores de carga."""
        self.status_var.set(f"❌ Error: {error_msg}")
        self.btn_refresh.configure(state=NORMAL)
