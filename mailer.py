"""
Mailer Module.
Sends the generated Copilot Usage HTML report via Microsoft Graph API.
Uses POST /users/{sender}/sendMail with application permission Mail.Send.
Supports file attachments (e.g. .xlsx, .csv) encoded in Base64.
"""

import base64
import logging
from pathlib import Path
from typing import List, Optional
import requests
import config

logger = logging.getLogger(__name__)


class MailerError(Exception):
    """Exception raised for mail sending failures."""
    pass


def send_report_email(
    token: str,
    html_content: str,
    subject: str = None,
    sender: str = None,
    recipients: List[str] = None,
    attachment_paths: Optional[List[Path]] = None,
) -> bool:
    """
    Sends the HTML report via Microsoft Graph sendMail endpoint.
    Attaches optional Excel or CSV audit files.
    """
    subject = subject or config.MAIL_SUBJECT
    sender = sender or config.MAIL_SENDER
    recipients = recipients or config.MAIL_RECIPIENTS

    if not sender:
        raise MailerError(
            "MAIL_SENDER is not configured. Specify MAIL_SENDER in your environment or .env file."
        )

    if not recipients:
        raise MailerError(
            "MAIL_RECIPIENTS is empty. Provide at least one recipient email address."
        )

    url = f"{config.GRAPH_API_V1}/users/{sender}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    message_payload = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": html_content,
        },
        "toRecipients": [
            {"emailAddress": {"address": r}} for r in recipients
        ],
    }

    # Add attachments if provided
    attachments_list = []
    if attachment_paths:
        for file_path in attachment_paths:
            p = Path(file_path)
            if p.exists() and p.is_file():
                with open(p, "rb") as f:
                    file_b64 = base64.b64encode(f.read()).decode("utf-8")

                content_type = (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    if p.suffix.lower() == ".xlsx"
                    else "text/csv"
                )
                attachments_list.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": p.name,
                    "contentType": content_type,
                    "contentBytes": file_b64,
                })
                logger.info(f"Attached file: {p.name} ({p.stat().st_size} bytes)")

    if attachments_list:
        message_payload["attachments"] = attachments_list

    payload = {
        "message": message_payload,
        "saveToSentItems": "true",
    }

    logger.info(f"Dispatching email: {sender} -> {', '.join(recipients)}")
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=45)
        if response.status_code in (200, 202):
            logger.info("Report email with attachments dispatched successfully.")
            return True

        response.raise_for_status()
        return True
    except requests.exceptions.HTTPError as e:
        status_code = response.status_code if "response" in locals() else None
        err_msg = response.text if "response" in locals() else str(e)
        if status_code == 403:
            raise MailerError(
                "Mail.Send permission required (HTTP 403 Forbidden). Ensure 'Mail.Send' "
                "application permission is granted and consented in Microsoft Entra ID."
            ) from e
        raise MailerError(f"Failed to dispatch email (HTTP {status_code}): {err_msg}") from e
    except requests.exceptions.RequestException as e:
        raise MailerError(f"Network error while dispatching email: {e}") from e
