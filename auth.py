import os
import msal

POWERBI_RESOURCE = "https://analysis.windows.net/powerbi/api"

TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]


def acquire_token() -> str:
    """
    Service Principal authentication (client credentials flow).
    Works in Streamlit Cloud (no interaction required).
    """
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )

    result = app.acquire_token_for_client(
        scopes=[f"{POWERBI_RESOURCE}/.default"]
    )

    if "access_token" not in result:
        raise RuntimeError(
            f"Service Principal auth failed:\n{result}"
        )

    return result["access_token"]
