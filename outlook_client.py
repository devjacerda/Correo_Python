"""
Módulo de conexión a Outlook via COM (win32com).
Gestiona la conexión al cliente de Outlook y acceso a carpetas del buzón.
"""

import win32com.client
import pythoncom


class OutlookClient:
    """Cliente para interactuar con Microsoft Outlook via COM."""

    # Constantes de tipos de carpeta de Outlook
    FOLDER_TYPES = {
        "inbox": 6,          # olFolderInbox
        "sent": 5,           # olFolderSentMail
        "drafts": 16,        # olFolderDrafts
        "deleted": 3,        # olFolderDeletedItems
        "outbox": 4,         # olFolderOutbox
        "junk": 23,          # olFolderJunk
        "calendar": 9,       # olFolderCalendar
        "contacts": 10,      # olFolderContacts
    }

    def __init__(self):
        """Inicializa la conexión con Outlook."""
        self.outlook = None
        self.namespace = None
        self._connect()

    def _connect(self):
        """Establece la conexión COM con Outlook."""
        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            raise ConnectionError(
                f"No se pudo conectar con Outlook. "
                f"Asegúrate de que Outlook esté abierto y configurado. Error: {e}"
            )

    def get_default_folder(self, folder_type: str = "inbox"):
        """
        Obtiene una carpeta predeterminada de Outlook.
        
        Args:
            folder_type: Tipo de carpeta ('inbox', 'sent', 'drafts', etc.)
        
        Returns:
            Objeto carpeta de Outlook
        """
        folder_id = self.FOLDER_TYPES.get(folder_type.lower())
        if folder_id is None:
            raise ValueError(
                f"Tipo de carpeta '{folder_type}' no reconocido. "
                f"Opciones: {', '.join(self.FOLDER_TYPES.keys())}"
            )
        return self.namespace.GetDefaultFolder(folder_id)

    def get_folder_by_path(self, path: str):
        """
        Obtiene una carpeta por su ruta completa.
        Ejemplo: 'Bandeja de entrada/Proyectos/2024'
        
        Args:
            path: Ruta de la carpeta separada por '/'
        
        Returns:
            Objeto carpeta de Outlook
        """
        parts = path.strip("/").split("/")
        folder = self.get_default_folder("inbox")

        # Si la ruta tiene más de un nivel, navegar desde la raíz
        if len(parts) > 1 or parts[0].lower() != "inbox":
            root = self.namespace.Folders
            folder = None
            for part in parts:
                if folder is None:
                    # Buscar en las cuentas raíz
                    for i in range(root.Count):
                        account_folder = root.Item(i + 1)
                        if account_folder.Name.lower() == part.lower():
                            folder = account_folder
                            break
                    if folder is None:
                        # Intentar desde la bandeja de entrada
                        folder = self.get_default_folder("inbox")
                        try:
                            folder = folder.Folders[part]
                        except Exception:
                            raise ValueError(f"No se encontró la carpeta: {part}")
                else:
                    try:
                        folder = folder.Folders[part]
                    except Exception:
                        raise ValueError(
                            f"No se encontró la subcarpeta: {part} en {folder.Name}"
                        )
        return folder

    def list_folders(self, parent=None, indent=0, max_depth=3):
        """
        Lista las carpetas disponibles en el buzón.
        
        Args:
            parent: Carpeta padre (None = raíz)
            indent: Nivel de indentación actual
            max_depth: Profundidad máxima de recursión
            
        Returns:
            Lista de tuplas (nombre, ruta, cantidad_items, indent)
        """
        folders_info = []

        if parent is None:
            # Listar todas las cuentas
            for i in range(self.namespace.Folders.Count):
                account = self.namespace.Folders.Item(i + 1)
                folders_info.append((account.Name, account.Name, 0, indent))
                if indent < max_depth:
                    folders_info.extend(
                        self.list_folders(account, indent + 1, max_depth)
                    )
        else:
            try:
                for i in range(parent.Folders.Count):
                    folder = parent.Folders.Item(i + 1)
                    try:
                        item_count = folder.Items.Count
                    except Exception:
                        item_count = 0
                    full_path = f"{parent.Name}/{folder.Name}"
                    folders_info.append(
                        (folder.Name, full_path, item_count, indent)
                    )
                    if indent < max_depth:
                        folders_info.extend(
                            self.list_folders(folder, indent + 1, max_depth)
                        )
            except Exception:
                pass

        return folders_info

    def get_account_email(self):
        """Obtiene la dirección de email de la cuenta principal."""
        try:
            accounts = self.namespace.Accounts
            if accounts.Count > 0:
                return accounts.Item(1).SmtpAddress
        except Exception:
            pass
        return "No disponible"
