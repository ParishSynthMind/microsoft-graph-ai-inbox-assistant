"""Signing in to Microsoft Graph with MSAL, two ways.

Delegated, device code (default). A person signs in once in any browser by
typing a short code at microsoft.com/devicelogin. The assistant then acts as
that person, can only touch what they can touch, and only within the scopes
below. MSAL keeps a refresh token in a local cache file (owner-only
permissions), so later runs are silent. Works with personal Microsoft
accounts (Outlook.com) and with work or school accounts.

App-only, client credentials. For a business tenant where the assistant runs
as a background service with no one signed in. An admin grants application
permissions once; the secret (or certificate) is read from the environment,
never from a file in the repo. App-only Mail permissions reach every mailbox
in the tenant unless an admin scopes them down, so the README says how to
limit it to one mailbox (Exchange RBAC for Applications) and one SharePoint
site (Sites.Selected).
"""
from __future__ import annotations

import os
import stat
from pathlib import Path
from typing import Callable

DELEGATED_SCOPES = ["User.Read", "Mail.ReadWrite", "Calendars.ReadWrite", "Files.ReadWrite"]
APP_SCOPES = ["https://graph.microsoft.com/.default"]
DEFAULT_CACHE = Path.home() / ".graph-assistant" / "token_cache.json"


class AuthError(RuntimeError):
    pass


def _load_cache(path: Path):
    import msal

    cache = msal.SerializableTokenCache()
    if path.exists():
        cache.deserialize(path.read_text())
    return cache


def _save_cache(cache, path: Path) -> None:
    if not cache.has_state_changed:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(cache.serialize())
    os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)


def device_code_token(client_id: str, tenant: str = "consumers", cache_path: Path = DEFAULT_CACHE,
                      scopes: list[str] = DELEGATED_SCOPES, interactive: bool = False,
                      show: Callable[[str], None] = print) -> Callable[[], str]:
    """Returns a function that gives a valid access token. With interactive=True and no cached sign-in,
    it starts the device-code flow and shows the message (URL and code) through `show`.

    tenant: "consumers" for personal Microsoft accounts, "organizations" for any work account,
    or a tenant id / domain to lock it to one organisation."""
    import msal

    cache = _load_cache(cache_path)
    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant}",
                                       token_cache=cache)

    def token() -> str:
        accounts = app.get_accounts()
        result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
        if not result:
            if not interactive:
                raise AuthError("not signed in: run `graph-assistant login` first")
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                raise AuthError(f"could not start device sign-in: {flow.get('error_description', flow)}")
            show(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
        _save_cache(cache, cache_path)
        if "access_token" not in result:
            raise AuthError(f"sign-in failed: {result.get('error')}: {result.get('error_description')}")
        return result["access_token"]

    return token


def client_credentials_token(client_id: str, tenant: str, secret: str | None = None) -> Callable[[], str]:
    """App-only. `secret` defaults to the GRAPH_CLIENT_SECRET environment variable."""
    import msal

    secret = secret or os.environ.get("GRAPH_CLIENT_SECRET")
    if not secret:
        raise AuthError("GRAPH_CLIENT_SECRET is not set")
    if tenant in ("common", "consumers", "organizations"):
        raise AuthError("app-only sign-in needs your tenant id or domain, not 'common'")
    app = msal.ConfidentialClientApplication(client_id, client_credential=secret,
                                             authority=f"https://login.microsoftonline.com/{tenant}")

    def token() -> str:
        result = app.acquire_token_for_client(scopes=APP_SCOPES)  # MSAL caches it in memory until expiry
        if "access_token" not in result:
            raise AuthError(f"app sign-in failed: {result.get('error')}: {result.get('error_description')}")
        return result["access_token"]

    return token


def signed_out(cache_path: Path = DEFAULT_CACHE) -> bool:
    if cache_path.exists():
        cache_path.unlink()
        return True
    return False
