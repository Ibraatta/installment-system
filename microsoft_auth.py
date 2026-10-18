from __future__ import annotations
from typing import Any
from pathlib import Path
import os, json, hashlib, tempfile, time
import msal
import streamlit as st
AUTH_FLOW_DIR=Path(tempfile.gettempdir())/"installment_oauth"
AUTH_FLOW_TTL_SECONDS=900
GRAPH_SCOPES=["User.Read","Files.ReadWrite"]
class CloudAuthenticationError(RuntimeError): pass
def microsoft_config() -> dict[str, str]:
    """Read Microsoft/OneDrive config from Azure env vars or Streamlit secrets."""
    try:
        section = st.secrets["microsoft"]
    except Exception:
        section = {}

    def setting(env_name: str, secret_name: str, default: str = "") -> str:
        env_value = os.getenv(env_name, "").strip()
        if env_value:
            return env_value
        try:
            return str(section.get(secret_name, default)).strip()
        except Exception:
            return default

    return {
        "tenant_id": setting("INSTALLMENT_TENANT_ID", "tenant_id"),
        "client_id": setting("INSTALLMENT_CLIENT_ID", "client_id"),
        "client_secret": setting("INSTALLMENT_CLIENT_SECRET", "client_secret"),
        "onedrive_folder": setting(
            "INSTALLMENT_ONEDRIVE_FOLDER",
            "onedrive_folder",
            "Installment System V2",
        ),
        "workbook_name": setting("INSTALLMENT_WORKBOOK_NAME", "workbook_name", "Installment_System_V2.xlsx"),
        "redirect_uri": setting(
            "INSTALLMENT_REDIRECT_URI",
            "redirect_uri",
            "http://localhost:8501",
        ),
        "authority": setting(
            "INSTALLMENT_AUTHORITY",
            "authority",
            "consumers",
        ) or "consumers",
    }


def cloud_mode_enabled() -> bool:
    """Return True when all required OneDrive credentials are configured."""
    config = microsoft_config()
    return all(config.get(key) for key in ("client_id", "client_secret"))


def msal_cache() -> msal.SerializableTokenCache:
    """Load the per-session Microsoft token cache."""
    cache = msal.SerializableTokenCache()
    serialized = st.session_state.get("msal_cache")
    if serialized:
        cache.deserialize(serialized)
    return cache


def save_msal_cache(cache: msal.SerializableTokenCache) -> None:
    """Persist the token cache in the current Streamlit session only."""
    if cache.has_state_changed:
        st.session_state["msal_cache"] = cache.serialize()


def build_msal_app(cache: msal.SerializableTokenCache | None = None) -> msal.ConfidentialClientApplication:
    """Create the confidential Microsoft client used by the web application."""
    config = microsoft_config()
    authority = config.get("authority", "consumers")
    return msal.ConfidentialClientApplication(
        client_id=config["client_id"],
        authority=f"https://login.microsoftonline.com/{authority}",
        client_credential=config["client_secret"],
        token_cache=cache,
    )


def microsoft_auth_response() -> dict[str, str]:
    """Convert Streamlit query parameters into an MSAL-compatible dictionary."""
    response: dict[str, str] = {}
    for key in ("code", "state", "error", "error_description"):
        value = st.query_params.get(key)
        if isinstance(value, list):
            value = value[0] if value else None
        if value is not None:
            response[key] = str(value)
    return response



def auth_flow_file(state: str) -> Path:
    """Return a safe temporary path for one Microsoft OAuth flow."""
    digest = hashlib.sha256(state.encode("utf-8")).hexdigest()
    return AUTH_FLOW_DIR / f"{digest}.json"


def save_auth_flow(flow: dict[str, Any]) -> None:
    """Persist the short-lived OAuth flow across the Microsoft redirect."""
    state = str(flow.get("state", "")).strip()
    if not state:
        raise CloudAuthenticationError("Microsoft login flow did not include state.")

    AUTH_FLOW_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "created_at": time.time(),
        "flow": flow,
    }
    path = auth_flow_file(state)
    path.write_text(json.dumps(payload), encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def load_auth_flow(state: str) -> dict[str, Any] | None:
    """Load a valid short-lived OAuth flow after returning from Microsoft."""
    if not state:
        return None

    path = auth_flow_file(state)
    if not path.exists():
        return None

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        created_at = float(payload.get("created_at", 0))
        if time.time() - created_at > AUTH_FLOW_TTL_SECONDS:
            path.unlink(missing_ok=True)
            return None

        flow = payload.get("flow")
        return flow if isinstance(flow, dict) else None
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        path.unlink(missing_ok=True)
        return None


def delete_auth_flow(state: str) -> None:
    """Remove a completed or invalid temporary OAuth flow."""
    if not state:
        return
    try:
        auth_flow_file(state).unlink(missing_ok=True)
    except OSError:
        pass


def acquire_microsoft_token() -> str | None:
    """Authenticate the signed-in user and return a delegated Microsoft Graph token."""
    if not cloud_mode_enabled():
        return None

    config = microsoft_config()
    cache = msal_cache()
    app = build_msal_app(cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        save_msal_cache(cache)
        if result and result.get("access_token"):
            st.session_state["graph_access_token"] = result["access_token"]
            st.session_state["graph_token_expires_at"] = time.time() + float(
                result.get("expires_in", 3000)
            )
            return result["access_token"]

    auth_response = microsoft_auth_response()
    if auth_response.get("error"):
        description = auth_response.get("error_description") or auth_response["error"]
        st.query_params.clear()
        st.error(f"تعذر تسجيل الدخول إلى Microsoft: {description}")
        return None

    if auth_response.get("code"):
        state = str(auth_response.get("state", "")).strip()
        flow = st.session_state.get("microsoft_auth_flow")
        if not flow or str(flow.get("state", "")).strip() != state:
            flow = load_auth_flow(state)

        if not flow:
            st.query_params.clear()
            st.error(
                "انتهت جلسة تسجيل الدخول المؤقتة. اضغط تسجيل الدخول مرة أخرى."
            )
            return None

        try:
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError as exc:
            delete_auth_flow(state)
            st.session_state.pop("microsoft_auth_flow", None)
            st.query_params.clear()
            st.error(
                "تعذر التحقق من تسجيل الدخول إلى Microsoft. "
                "اضغط تسجيل الدخول مرة أخرى."
            )
            return None

        delete_auth_flow(state)
        save_msal_cache(cache)
        st.session_state.pop("microsoft_auth_flow", None)
        st.query_params.clear()

        if result.get("access_token"):
            st.session_state["graph_access_token"] = result["access_token"]
            st.session_state["graph_token_expires_at"] = time.time() + float(
                result.get("expires_in", 3000)
            )
            st.rerun()

        st.error(
            "تعذر تسجيل الدخول إلى Microsoft: "
            + str(result.get("error_description") or result.get("error") or "Unknown error")
        )
        return None

    if st.session_state.get("graph_access_token"):
        expires_at = float(st.session_state.get("graph_token_expires_at", 0))
        if time.time() < expires_at - 120:
            return str(st.session_state["graph_access_token"])

    flow = app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPES,
        redirect_uri=config["redirect_uri"],
    )
    save_auth_flow(flow)
    save_msal_cache(cache)
    st.session_state["microsoft_auth_flow"] = flow

    st.info("سجّل الدخول بحساب Microsoft الشخصي للوصول إلى ملف OneDrive.")
    st.link_button(
        "تسجيل الدخول إلى Microsoft",
        flow["auth_uri"],
        type="primary",
        width='stretch',
    )
    return None


