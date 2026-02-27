"""
Módulo de búsqueda de correos en Outlook.
Permite búsquedas flexibles por múltiples criterios con soporte de filtros DASL.
"""

import threading
from datetime import datetime, timedelta
from typing import Optional, Callable


class EmailSearch:
    """Motor de búsqueda de correos en Outlook."""

    def __init__(self, outlook_client):
        """
        Args:
            outlook_client: Instancia de OutlookClient
        """
        self.client = outlook_client

    def search(
        self,
        subject: Optional[str] = None,
        sender: Optional[str] = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        folder: str = "inbox",
        has_attachments: Optional[bool] = None,
        body_contains: Optional[str] = None,
        recipient: Optional[str] = None,
        max_results: int = 500,
        subfolder: Optional[str] = None,
        progress_callback: Optional[Callable] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> list:
        """
        Busca correos con múltiples filtros.
        
        Args:
            subject: Texto a buscar en el asunto (parcial)
            sender: Nombre o email del remitente
            date_from: Fecha inicio 'DD-MM-YYYY'
            date_to: Fecha fin 'DD-MM-YYYY'
            folder: Tipo de carpeta ('inbox', 'sent', etc.)
            has_attachments: Filtrar por adjuntos (True/False/None)
            body_contains: Texto a buscar en el cuerpo
            recipient: Destinatario (para carpeta sent)
            max_results: Máximo de resultados a retornar
            subfolder: Subcarpeta dentro de la carpeta principal
            progress_callback: Función opcional (current, message) para reportar progreso
            
        Returns:
            Lista de diccionarios con datos de cada correo
        """
        # Obtener la carpeta
        try:
            target_folder = self.client.get_default_folder(folder)
            if subfolder:
                target_folder = target_folder.Folders[subfolder]
        except Exception as e:
            raise ValueError(f"Error al acceder a la carpeta: {e}")

        # Construir filtro DASL para mejor rendimiento
        dasl_filter = self._build_dasl_filter(
            subject, sender, date_from, date_to, has_attachments
        )

        # Ejecutar búsqueda
        results = []
        try:
            items = target_folder.Items
            items.Sort("[ReceivedTime]", True)  # Más recientes primero

            if dasl_filter:
                items = items.Restrict(dasl_filter)

            count = 0
            for item in items:
                if count >= max_results:
                    break
                if cancel_event and cancel_event.is_set():
                    break

                try:
                    # Filtros adicionales que no se pueden hacer con DASL
                    if body_contains and body_contains.lower() not in (
                        item.Body or ""
                    ).lower():
                        continue

                    if recipient:
                        recipients_str = ""
                        try:
                            for r in range(item.Recipients.Count):
                                recip = item.Recipients.Item(r + 1)
                                recipients_str += (
                                    f"{recip.Name} {recip.Address} "
                                )
                        except Exception:
                            pass
                        if recipient.lower() not in recipients_str.lower():
                            continue

                    email_data = self._extract_email_data(item)
                    results.append(email_data)
                    count += 1

                    if progress_callback:
                        progress_callback(count, f"Encontrados: {count} correos...")

                except Exception:
                    continue

        except Exception as e:
            raise RuntimeError(f"Error durante la búsqueda: {e}")

        return results

    def _build_dasl_filter(
        self,
        subject=None,
        sender=None,
        date_from=None,
        date_to=None,
        has_attachments=None,
    ) -> str:
        """
        Construye un filtro DASL para Outlook.
        
        Returns:
            String con el filtro DASL o vacío si no hay filtros
        """
        conditions = []

        if subject:
            # urn:schemas:httpmail:subject para búsqueda parcial
            conditions.append(
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject}%'"
            )

        if sender:
            # Buscar por nombre o email del remitente
            sender_filter = (
                f"@SQL=(\"urn:schemas:httpmail:fromemail\" LIKE '%{sender}%' "
                f"OR \"urn:schemas:httpmail:fromname\" LIKE '%{sender}%')"
            )
            conditions.append(sender_filter)

        if date_from:
            try:
                dt_from = datetime.strptime(date_from, "%d-%m-%Y")
                date_str = dt_from.strftime("%m/%d/%Y")
                conditions.append(
                    f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{date_str}'"
                )
            except ValueError:
                raise ValueError(
                    f"Formato de fecha_desde inválido: {date_from}. Use DD-MM-YYYY"
                )

        if date_to:
            try:
                dt_to = datetime.strptime(date_to, "%d-%m-%Y")
                # Agregar un día para incluir todo el día final
                dt_to += timedelta(days=1)
                date_str = dt_to.strftime("%m/%d/%Y")
                conditions.append(
                    f"@SQL=\"urn:schemas:httpmail:datereceived\" < '{date_str}'"
                )
            except ValueError:
                raise ValueError(
                    f"Formato de fecha_hasta inválido: {date_to}. Use DD-MM-YYYY"
                )

        if has_attachments is not None:
            val = "1" if has_attachments else "0"
            conditions.append(
                f"@SQL=\"urn:schemas:httpmail:hasattachment\" = {val}"
            )

        if not conditions:
            return ""

        if len(conditions) == 1:
            return conditions[0]

        # Construir filtro combinado con AND
        inner_parts = []
        for c in conditions:
            # Extraer la parte interna del @SQL=
            if c.startswith("@SQL="):
                inner = c[5:]
            elif c.startswith("@SQL("):
                inner = c[4:]
            else:
                inner = c
            inner_parts.append(inner)

        combined = " AND ".join(inner_parts)
        return f"@SQL={combined}"

    def _extract_email_data(self, item) -> dict:
        """
        Extrae datos relevantes de un objeto de correo de Outlook.
        
        Args:
            item: Objeto MailItem de Outlook
            
        Returns:
            Diccionario con los campos del correo
        """
        try:
            received_time = item.ReceivedTime
            date_str = received_time.strftime("%d-%m-%Y")
            time_str = received_time.strftime("%H:%M:%S")
        except Exception:
            date_str = "N/A"
            time_str = "N/A"

        try:
            sender_email = item.SenderEmailAddress or "N/A"
            # Si es una dirección Exchange, intentar obtener SMTP
            if sender_email and "/" in sender_email:
                try:
                    sender_obj = item.Sender
                    if sender_obj:
                        exch_user = sender_obj.GetExchangeUser()
                        if exch_user:
                            sender_email = exch_user.PrimarySmtpAddress
                except Exception:
                    pass
        except Exception:
            sender_email = "N/A"

        try:
            attachment_count = item.Attachments.Count
            attachment_names = []
            for i in range(attachment_count):
                att = item.Attachments.Item(i + 1)
                attachment_names.append(att.FileName)
        except Exception:
            attachment_count = 0
            attachment_names = []

        try:
            importance_map = {0: "Baja", 1: "Normal", 2: "Alta"}
            importance = importance_map.get(item.Importance, "Normal")
        except Exception:
            importance = "Normal"

        try:
            body_preview = (item.Body or "")[:200].replace("\r\n", " ").strip()
        except Exception:
            body_preview = ""

        try:
            categories = item.Categories or ""
        except Exception:
            categories = ""

        try:
            cc = item.CC or ""
        except Exception:
            cc = ""

        try:
            to = item.To or ""
        except Exception:
            to = ""

        return {
            "subject": getattr(item, "Subject", "Sin asunto") or "Sin asunto",
            "sender_name": getattr(item, "SenderName", "N/A") or "N/A",
            "sender_email": sender_email,
            "to": to,
            "cc": cc,
            "date": date_str,
            "time": time_str,
            "body_preview": body_preview,
            "has_attachments": attachment_count > 0,
            "attachment_count": attachment_count,
            "attachment_names": attachment_names,
            "importance": importance,
            "categories": categories,
            "size_kb": round(getattr(item, "Size", 0) / 1024, 1),
            "_outlook_item": item,  # Referencia interna para adjuntos
        }

    def get_results_without_item(self, results: list) -> list:
        """
        Retorna los resultados sin la referencia al objeto COM (para serializar).
        
        Args:
            results: Lista de resultados de búsqueda
            
        Returns:
            Lista limpia sin objetos COM
        """
        clean = []
        for r in results:
            row = {k: v for k, v in r.items() if k != "_outlook_item"}
            clean.append(row)
        return clean
