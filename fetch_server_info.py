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
SETTINGS = {
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


def parse_settings(raw: dict) -> dict:
    """
    Map raw setting values {key: str} to named fields, expanding the SMTP (604)
    and storage status (1301) XML values. Shared by the REST and TheXMLServer paths.
    """
    info: dict = {}
    for key, val in raw.items():
        if key in SETTINGS and val:
            info[SETTINGS[key][0]] = val

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


def _make_headers(username: str, password: str, tenant: str) -> dict:
    auth_header = base64.b64encode(f"{username}:{password}".encode()).decode()
    h = {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": f"Basic {auth_header}",
    }
    if tenant:
        h["TenantName"] = tenant
    return h


def fetch_users_groups(api_url: str, tenant: str, username: str, password: str) -> dict:
    """
    Fetch all users and group memberships from the Therefore API.

    Returns a dict compatible with parse_security() output:
        {"users": [...], "groups": [...], "roles": [], "role_assignments": [], "source": "api"}

    Each user has: user_no, user_type, user_name, display_name, email, disabled, member_of.
    Each group has: user_no, user_type, user_name, display_name, members.
    """
    base    = _restun_base(api_url)
    tenant  = derive_tenant(api_url, tenant)
    headers = _make_headers(username, password, tenant)

    # 1. All named users (Flags=4; 0–3 return empty list)
    raw_users = _post(base, "ExecuteUsersQuery", {"Flags": 4}, headers).get("Users", [])

    users = []
    user_by_name: dict = {}
    for u in raw_users:
        uname = u.get("UserName", "")
        obj = {
            "user_no":      u.get("UserId") or 0,
            "user_type":    1,
            "user_name":    uname,
            "display_name": u.get("DisplayName") or uname,
            "email":        u.get("SMTP", ""),
            "disabled":     bool(u.get("Disabled")),
            "members":      [],
            "member_of":    [],
        }
        users.append(obj)
        user_by_name[uname] = obj

    # 2. All groups via GetObjects Type=11; Data==2 → group
    items = _post(base, "GetObjects", {"Flags": 0, "Type": 11}, headers).get("ItemList", [])
    raw_groups = [i for i in items if i.get("Data") == 2]

    # 3. Members per group
    member_of: dict = {u["user_name"]: [] for u in users}
    groups = []
    for g in raw_groups:
        gname = g.get("Name", "")
        try:
            members_raw = _post(base, "GetUsersFromGroup", {"GroupName": gname}, headers).get("Users", [])
        except Exception:
            members_raw = []

        members = []
        for m in members_raw:
            mname = m.get("UserName", "")
            members.append({
                "user_no":      m.get("UserId") or 0,
                "user_type":    1,
                "user_name":    mname,
                "display_name": m.get("DisplayName") or mname,
            })
            if mname in member_of:
                member_of[mname].append(gname)

        groups.append({
            "user_no":      g.get("ID") or 0,
            "user_type":    2,
            "user_name":    gname,
            "display_name": gname,
            "email":        "",
            "disabled":     False,
            "members":      members,
        })

    # Attach member_of lists to users
    for u in users:
        u["member_of"] = [
            {"user_name": gn, "display_name": gn, "user_no": 0, "user_type": 2}
            for gn in member_of.get(u["user_name"], [])
        ]

    return {
        "users":            users,
        "groups":           groups,
        "roles":            [],
        "role_assignments": [],
        "source":           "api",
    }


def fetch_server_info(api_url: str, tenant: str, username: str, password: str) -> dict:
    """
    Fetch Therefore server configuration details.

    Returns a dict of discovered values; individual keys absent on failure.
    Raises urllib.error.HTTPError / URLError on connection or auth failure.
    """
    base    = _restun_base(api_url)
    tenant  = derive_tenant(api_url, tenant)
    headers = _make_headers(username, password, tenant)

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
    raw = {}
    for key in SETTINGS:
        val = _global_setting(base, key, headers)
        if val is not None:
            raw[key] = val
    info.update(parse_settings(raw))

    return info
