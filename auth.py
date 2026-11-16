"""
Authentication module for Microsoft Graph API.
Seamlessly supports:
  1. Service Principal (Client ID & Client Secret) - Ideal for local development & testing.
  2. Azure Managed Identity (System-Assigned or User-Assigned) - Ideal for Azure Automation Runbooks.
"""

import os
import logging
import requests
import config

logger = logging.getLogger(__name__)


class GraphAuthError(Exception):
    """Custom exception for authentication failures."""
    pass


def get_graph_access_token() -> str:
    """
    Acquires an OAuth2 Bearer token for Microsoft Graph API.
    Auto-detects whether running in Azure (Managed Identity) or Local (Service Principal).
    """
    # 1. Try Azure Automation Managed Identity if environment variables are set
    if config.IDENTITY_ENDPOINT and config.IDENTITY_HEADER:
        logger.info("Azure Automation Managed Identity ortamı tespit edildi. Token alınıyor...")
        try:
            token = _get_token_via_azure_automation_identity()
            if token:
                logger.info("Managed Identity ile Graph token başarıyla alındı.")
                return token
        except Exception as e:
            logger.warning(f"Managed Identity ile token alınamadı: {e}. Diğer yöntemler deneniyor...")

    # 2. Try IMDS Managed Identity (e.g. Azure VM / Container) if available
    if os.getenv("AZURE_USE_IMDS", "false").lower() == "true":
        try:
            token = _get_token_via_imds()
            if token:
                return token
        except Exception as e:
            logger.debug(f"IMDS token denemesi: {e}")

    # 3. Fallback to Service Principal (Tenant ID, Client ID, Client Secret)
    if config.AZURE_TENANT_ID and config.AZURE_CLIENT_ID and config.AZURE_CLIENT_SECRET:
        logger.info("Service Principal (Client Secret) ile Graph token alınıyor...")
        return _get_token_via_client_credentials()

    # If neither is configured
    raise GraphAuthError(
        "Kimlik doğrulama bilgileri bulunamadı!\n"
        "Local çalıştırma için lütfen .env dosyasında AZURE_TENANT_ID, AZURE_CLIENT_ID ve AZURE_CLIENT_SECRET tanımlayın.\n"
        "Azure Automation için ise Runbook'ta Managed Identity'yi aktif edin ve Graph API izinlerini atayın.\n"
        "Örnek veriyle test etmek için scripti '--mock' parametresiyle çalıştırabilirsiniz."
    )


def _get_token_via_client_credentials() -> str:
    """
    Acquires a token using Entra ID OAuth 2.0 Client Credentials Grant.
    """
    token_url = f"{config.LOGIN_MICROSOFT_ONLINE}/{config.AZURE_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": config.AZURE_CLIENT_ID,
        "client_secret": config.AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    try:
        response = requests.post(token_url, data=payload, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        token = data.get("access_token")
        if not token:
            raise GraphAuthError(f"Token yanıtında 'access_token' bulunamadı: {data}")
        return token
    except requests.exceptions.HTTPError as e:
        err_detail = response.text if "response" in locals() else str(e)
        raise GraphAuthError(
            f"Service Principal ile token alma başarısız oldu (HTTP {response.status_code}): {err_detail}"
        ) from e
    except requests.exceptions.RequestException as e:
        raise GraphAuthError(f"Token isteği sırasında ağ hatası oluştu: {e}") from e


def _get_token_via_azure_automation_identity() -> str:
    """
    Acquires a token using Azure Automation System/User Assigned Managed Identity.
    Azure Automation exposes IDENTITY_ENDPOINT and IDENTITY_HEADER.
    """
    url = f"{config.IDENTITY_ENDPOINT}?resource=https://graph.microsoft.com/&api-version=2019-08-01"
    headers = {"X-IDENTITY-HEADER": config.IDENTITY_HEADER}

    response = requests.get(url, headers=headers, timeout=20)
    response.raise_for_status()
    data = response.json()
    token = data.get("access_token")
    if not token:
        raise GraphAuthError(f"Managed Identity yanıtında 'access_token' bulunamadı: {data}")
    return token


def _get_token_via_imds() -> str:
    """
    Fallback for standard Azure VM IMDS endpoint.
    """
    url = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https://graph.microsoft.com/"
    headers = {"Metadata": "true"}
    response = requests.get(url, headers=headers, timeout=5)
    response.raise_for_status()
    return response.json().get("access_token")
