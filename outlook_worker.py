"""
Worker thread dedicado para operaciones COM con Outlook.
Todas las interacciones con Outlook ocurren en este thread,
evitando problemas de threading COM y manteniendo la GUI responsive.
"""

import threading
import queue
import pythoncom

from outlook_client import OutlookClient
from search import EmailSearch
from attachments import export_attachments as _export_attachments


class OutlookWorker(threading.Thread):
    """
    Thread dedicado que posee todos los objetos COM de Outlook.
    Recibe tareas via queue y devuelve resultados via root.after().
    """

    def __init__(self, app):
        super().__init__(daemon=True)
        self.app = app
        self.tasks = queue.Queue()
        self.client = None
        self.searcher = None
        self.last_results = []  # resultados CON _outlook_item (viven en este thread)
        self.cancel_event = threading.Event()  # señal para detener búsqueda

    def run(self):
        """Loop principal del worker thread."""
        pythoncom.CoInitialize()
        try:
            self.client = OutlookClient()
            self.searcher = EmailSearch(self.client)
            email = self.client.get_account_email()
            self.app.after(0, self.app._on_worker_ready, email)
        except Exception as e:
            self.app.after(0, self.app._on_worker_error, str(e))
            return

        # Procesar tareas indefinidamente
        while True:
            task = self.tasks.get()
            if task is None:
                break

            task_name, kwargs, on_success, on_error = task
            try:
                if task_name == "search":
                    self._do_search(kwargs, on_success)
                elif task_name == "quick_search_all":
                    self._do_quick_search_all(kwargs, on_success)
                elif task_name == "export_attachments":
                    self._do_export_attachments(kwargs, on_success)
                elif task_name == "list_folders":
                    self._do_list_folders(kwargs, on_success)
            except Exception as e:
                self.app.after(0, on_error, str(e))

    def submit(self, task_name, kwargs, on_success, on_error):
        """Envía una tarea al worker thread."""
        self.cancel_event.clear()  # resetear señal de cancelación
        self.tasks.put((task_name, kwargs, on_success, on_error))

    def cancel_search(self):
        """Detiene la búsqueda en curso."""
        self.cancel_event.set()

    # === Tareas ===

    def _do_search(self, kwargs, on_success):
        """Ejecuta búsqueda y almacena resultados con refs COM."""
        def progress_cb(current, msg):
            self.app.after(0, self.app._on_search_progress, current, msg)

        results = self.searcher.search(
            progress_callback=progress_cb,
            cancel_event=self.cancel_event,
            **kwargs,
        )
        self.last_results = results

        # Enviar resultados limpios (sin COM refs) a la GUI
        clean = self.searcher.get_results_without_item(results)
        cancelled = self.cancel_event.is_set()
        self.app.after(0, on_success, clean, cancelled)

    def _do_quick_search_all(self, kwargs, on_success):
        """Búsqueda rápida en subject + sender, combinada."""
        term = kwargs["term"]
        max_results = kwargs.get("max_results", 50)

        def progress_cb(current, msg):
            self.app.after(0, self.app._on_search_progress, current, msg)

        results = self.searcher.search(
            subject=term, max_results=max_results,
            progress_callback=progress_cb, cancel_event=self.cancel_event,
        )
        if self.cancel_event.is_set():
            clean = self.searcher.get_results_without_item(results)
            self.app.after(0, on_success, clean, True)
            return
        results_sender = self.searcher.search(
            sender=term, max_results=max_results,
            cancel_event=self.cancel_event,
        )

        seen = {(r["subject"], r["date"], r["time"]) for r in results}
        for r in results_sender:
            key = (r["subject"], r["date"], r["time"])
            if key not in seen:
                results.append(r)
                seen.add(key)

        self.last_results = results
        clean = self.searcher.get_results_without_item(results)
        cancelled = self.cancel_event.is_set()
        self.app.after(0, on_success, clean, cancelled)

    def _do_export_attachments(self, kwargs, on_success):
        """Exporta adjuntos usando las refs COM almacenadas."""
        def progress_cb(current, total, msg):
            self.app.after(0, self.app._on_attachment_progress, current, total, msg)

        stats = _export_attachments(
            results=self.last_results,
            progress_callback=progress_cb,
            **kwargs,
        )
        self.app.after(0, on_success, stats)

    def _do_list_folders(self, kwargs, on_success):
        """Lista carpetas del buzón."""
        folders = self.client.list_folders(**kwargs)
        self.app.after(0, on_success, folders)
