"""
fetch_server_info.py — Fetch Therefore server configuration via REST API.
"""
import base64
import json
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

# GetGlobalSettings numeric keys → (field_name, display_label)
# Keys with label=None are parsed from XML rather than surfaced directly.
_SETTINGS = {
    1:    ("db_connection",      "Server / Database"),
    3:    ("db_username",        "Username"),
    15:   ("db_options",         "Connection Options"),
    100:  ("smtp_server",        "SMTP Server"),
    101:  ("smtp_sender",        "Sender Address"),
    104:  ("buffer_path",        "Buffer Storage"),
    105:  ("cache_path",         "Cache Path"),
    123:  ("preview_blob",       "Preview Storage"),
    168:  ("report_server",      "Report Server"),
    174:  ("tenant_name",        "Tenant Name"),
    175:  ("locale",             "Locale"),
    194:  ("admin_email",        "Admin Email"),
    320:  ("region",             "Region"),
    501:  ("fulltext_path",      "Full-Text Index Storage"),
    524:  ("fulltext_blob",      "Full-Text Blob Storage"),
    604:  ("smtp_config_xml",    None),
    816:  ("storage_blob",       "Document Storage Blob"),
    1021: ("terms_url",          "Terms of Service"),
    1300: ("server_name",        "Server Name"),
    1301: ("storage_status_xml", None),
    1612: ("webviewer_blob",     "Web Viewer Storage"),
}


def _restun_base(api_url: str) -> str:
    """Normalise any Therefore URL form to the /restun base."""
    url = api_url.rstrip("/")
    if "/restun" in url:
        return url.split("/restun")[0] + "/restun"
    if "/theservice" in url:
        return url.split("/theservice")[0] + "/theservice/v0001/restun"
    return url + "/theservice/v0001/restun"


def derive_tenant(api_url: str, tenant: str) -> str:
    """Return tenant from explicit value or auto-detect from thereforeonline.com URL."""
    if tenant:
        return tenant
    parsed = urlparse(api_url)
    host = parsed.hostname or ""
    if "thereforeonline.com" in host:
        return host.split(".")[0]
    return ""


def _post(base: str, operation: str, body: dict, headers: dict) -> dict:
    url = f"{base}/{operation}"
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode("utf-8"))


def _get(base: str, operation: str, headers: dict) -> dict:
    url = f"{base}/{operation}"
    get_headers = {k: v for k, v in headers.items() if k != "Content-Type"}
    req = urllib.request.Request(url, headers=get_headers)
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode("utf-8"))


def _global_setting(base: str, key: int, headers: dict) -> str | None:
    try:
        result = _post(base, "GetGlobalSettings", {"Settings": [key]}, headers)
        vals = result.get("SettingValues", [])
        if vals:
            return vals[0].get("Value") or None
    except Exception:
        pass
    return None


def fetch_server_info(api_url: str, tenant: str, username: str, password: str) -> dict:
    """
    Fetch Therefore server configuration details.

    Returns a dict of discovered values; individual keys absent on failure.
    Raises urllib.error.HTTPError / URLError on connection or auth failure.
    """
    base   = _restun_base(api_url)
    tenant = derive_tenant(api_url, tenant)

    auth_header = base64.b64encode(f"{username}:{password}".encode()).decode()
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Basic {auth_header}",
    }
    if tenant:
        headers["TenantName"] = tenant

    anon_headers = {"Content-Type": "application/json; charset=utf-8"}
    if tenant:
        anon_headers["TenantName"] = tenant

    info: dict = {
        "api_url":    api_url.rstrip("/"),
        "api_tenant": tenant,
    }

    # Service version — no auth needed
    try:
        v = _post(base, "GetWebAPIServerVersion", {}, anon_headers)
        info["service_version"]     = v.get("ServiceVersion", "")
        info["version_description"] = v.get("VersionDescription", "")
    except Exception:
        pass

    # Customer ID — GET, no auth needed
    try:
        r = _get(base, "GetSystemCustomerId", anon_headers)
        info["customer_id"] = r.get("CustomerId", "")
    except Exception:
        pass

    # Domain info
    try:
        d = _post(base, "GetDomainInfo", {}, headers)
        info["domain_names"]   = d.get("DomainNames", [])
        info["default_domain"] = d.get("DefaultDomain", "")
    except Exception:
        pass

    # Global settings
    for key, (field, _label) in _SETTINGS.items():
        val = _global_setting(base, key, headers)
        if val is not None:
            info[field] = val

    # Parse SMTP config XML (key 604)
    smtp_xml = info.pop("smtp_config_xml", None)
    if smtp_xml:
        try:
            el = ET.fromstring(smtp_xml)
            info["smtp_server"]    = el.findtext("Server")  or info.get("smtp_server", "")
            info["smtp_sender"]    = el.findtext("Sender")  or info.get("smtp_sender", "")
            info["smtp_ssl"]       = el.findtext("UseSsl") == "1"
            auth_el = el.find("Auth")
            if auth_el is not None:
                info["smtp_auth_user"] = auth_el.findtext("User") or ""
        except Exception:
            pass

    # Parse storage status XML (key 1301)
    status_xml = info.pop("storage_status_xml", None)
    if status_xml:
        try:
            el = ET.fromstring(status_xml)
            info["storage_licensed_gb"] = el.findtext("LicensedStorage") or ""
            info["storage_used_gb"]     = el.findtext("Exceeded") or ""
        except Exception:
            pass

    return info
