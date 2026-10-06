"""
therefore_xmlserver.py — Export configuration and server details directly from a
Therefore server via the /TheXMLServer SOAP endpoint used by Solution Designer.

This is an internal Therefore protocol (not the public /theservice REST API).
Behaviour here was verified against Therefore Online build 35.0.3; see
https://github.com/Fybre/therefore-console-reference for the protocol findings.

Configuration export uses the same two-step route as Solution Designer's
"Export configuration": DoImportExport FactoryType 1 (the server builds the
export tree) followed by FactoryType 2 (the server serialises the selected
objects into the TheConfiguration XML).

Limitations:
  * Password login only — ASCII passwords; SSO / MFA accounts are not supported.
  * Therefore Online closes requests that run longer than about five minutes
    (an IIS ARR proxy limit). Very large tenants can exceed this while the
    server builds the export; this raises ExportTimeoutError.
"""
from __future__ import annotations

import http.client
import secrets
import socket
import ssl
import time
import urllib.error
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
OP_NS = "http://tempuri.org/"
ZERO_GUID = "00000000-0000-0000-0000-000000000000"

# Solution Designer client identity (CT 11) and the client/minimum server versions
# observed for build 35.0.3.
CLIENT_TYPE = 11
CLIENT_VERSION = 587202563
MIN_SERVER_VERSION = 587202560

# Session codes meaning "log in again" (session terminated / client not connected).
INVALID_SESSION_CODES = {-1073741614, -1073741672}

# Therefore Online's proxy cuts requests at ~300 s; wait slightly longer so its
# 502 (rather than our own timeout) is what we normally see.
DEFAULT_TIMEOUT = 330

# Login failures worth explaining (texts from TheMessages.dll, build 35.0.3).
LOGIN_ERRORS = {
    0xC000001B: "Invalid user name or password.",
    0xC000001D: "The tenant does not exist.",
    0xC000001E: "A tenant name is required for this server.",
    0xC0000018: "Access denied.",
    0xC0000019: "All license points are currently in use.",
    0xC00000B4: "All licenses are currently in use.",
    0xC0000024: "This account is locked after too many failed login attempts.",
    0xC0000025: "This user has another active session that must be disconnected first.",
    0x4000001F: "The user must change their password before logging on.",
}

_RADIX64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_-"


class XMLServerError(Exception):
    """A TheXMLServer call failed."""


class ResultError(XMLServerError):
    def __init__(self, method: str, code: int):
        self.code = code
        super().__init__(f"{method} returned {code} (0x{code & 0xFFFFFFFF:08X})")


class ExportTimeoutError(XMLServerError):
    """The server did not finish the export before the connection was closed."""


# ---------------------------------------------------------------------------
# Login proof encoding
# ---------------------------------------------------------------------------
def _encode_radix64(data: bytes) -> str:
    value = int.from_bytes(data, "little")
    digits = []
    while value:
        value, digit = divmod(value, 64)
        digits.append(_RADIX64[digit])
    width = (len(data) * 8 + 5) // 6
    return ("".join(reversed(digits)) or "0").rjust(width, "0")


def _decode_radix64(encoded: str, length: int) -> bytes:
    value = 0
    for digit in encoded:
        value = value * 64 + _RADIX64.index(digit)
    return value.to_bytes(length, "little")


def _login_proof(challenge: str, password: str, public_key: str) -> str:
    """Encrypt challenge + password with the server's CryptoAPI RSA public key."""
    from cryptography.hazmat.primitives.asymmetric import padding, rsa

    try:
        plaintext = (challenge + password).encode("ascii") + b"\x00"
    except UnicodeEncodeError:
        raise XMLServerError("Direct server login supports ASCII passwords only.") from None
    blob = _decode_radix64(public_key, 276)
    if blob[:8] != bytes.fromhex("0602000000a40000") or blob[8:12] != b"RSA1":
        raise XMLServerError("Server returned an unexpected login key format.")
    bits = int.from_bytes(blob[12:16], "little")
    exponent = int.from_bytes(blob[16:20], "little")
    modulus = int.from_bytes(blob[20:20 + bits // 8], "little")
    key = rsa.RSAPublicNumbers(exponent, modulus).public_key()
    # cryptography returns big-endian ciphertext; CryptoAPI expects little-endian.
    return "str:" + _encode_radix64(key.encrypt(plaintext, padding.PKCS1v15())[::-1])


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------
def xmlserver_url(server_url: str) -> str:
    """Normalise any Therefore URL form (tenant URL, /theservice, /TheXMLServer) to /TheXMLServer."""
    url = server_url.strip()
    if "://" not in url:
        url = "https://" + url
    parsed = urlparse(url)
    return f"{parsed.scheme}://{parsed.netloc}/TheXMLServer"


def derive_tenant(server_url: str, tenant: str = "") -> str:
    if tenant:
        return tenant
    host = urlparse(server_url if "://" in server_url else "https://" + server_url).hostname or ""
    return host.split(".")[0] if host.endswith(".thereforeonline.com") else ""


def _typed(value: str | None) -> str:
    """Strip the int:/str: type tag from a serialised value."""
    value = value or ""
    tag, sep, rest = value.partition(":")
    return rest if sep and tag in ("int", "str") else value


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------
class XMLServerClient:
    """Minimal authenticated TheXMLServer session. Use as a context manager."""

    def __init__(self, server_url: str, username: str, password: str, tenant: str = "",
                 timeout: float = DEFAULT_TIMEOUT, log_fn=None):
        self.url = xmlserver_url(server_url)
        self.tenant = derive_tenant(server_url, tenant)
        self.username = username
        self.password = password
        self.timeout = timeout
        self.session_id: str | None = None
        self.login_result: ET.Element | None = None
        self._log = log_fn or (lambda _msg: None)
        self._tls = ssl.create_default_context()

    def __enter__(self):
        self.login()
        return self

    def __exit__(self, *_exc):
        self.disconnect()

    # -- transport ----------------------------------------------------------
    def call(self, method: str, fields: list[tuple[str, object]], sec: bool = False,
             timeout: float | None = None) -> dict[str, str]:
        """POST one SOAP operation; return its response fields as {name: text}."""
        envelope = ET.Element(f"{{{SOAP_NS}}}Envelope")
        body = ET.SubElement(envelope, f"{{{SOAP_NS}}}Body")
        op = ET.SubElement(body, f"{{{OP_NS}}}{method}")
        for name, value in fields:
            ET.SubElement(op, f"{{{OP_NS}}}{name}").text = str(value)
        request = urllib.request.Request(
            self.url + ("/Sec" if sec else ""),
            data=ET.tostring(envelope, encoding="utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8",
                     "SOAPAction": f'"{OP_NS}ITheXMLService/{method}"'},
        )
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=self._tls))
        try:
            response = opener.open(request, timeout=timeout or self.timeout)
        except urllib.error.HTTPError as exc:
            response = exc  # SOAP faults arrive as HTTP 500 with an XML body
        with response:
            content = response.read()
            status = response.code
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            raise XMLServerError(f"{method}: server returned HTTP {status} without a SOAP response") from None
        if any(el.tag.endswith("}Fault") for el in root.iter()):
            raise XMLServerError(f"{method}: SOAP fault (HTTP {status})")
        result_el = next((el for el in root.iter() if el.tag == f"{{{OP_NS}}}{method}Response"), None)
        if result_el is None:
            raise XMLServerError(f"{method}: missing {method}Response")
        return {child.tag.split("}")[-1]: child.text or "" for child in result_el}

    def invoke(self, method: str, fields: list[tuple[str, object]], timeout: float | None = None) -> dict[str, str]:
        """Call a session method and raise ResultError on a nonzero METHODResult."""
        out = self.call(method, fields, timeout=timeout)
        code = int(out.get(f"{method}Result") or 0)
        if code:
            raise ResultError(method, code)
        return out

    # -- session ------------------------------------------------------------
    def _params(self, flags: int, proof: str = "", with_user: bool = True) -> str:
        root = ET.Element("ConParams")
        values = [("CT", CLIENT_TYPE), ("CV", CLIENT_VERSION), ("MSV", MIN_SERVER_VERSION),
                  ("Tenant", self.tenant)]
        if with_user:
            values.append(("U", self.username))
        values += [("P", proof), ("MM", 0)]
        for key, value in values:
            ET.SubElement(root, key).text = str(value)
        node = ET.SubElement(root, "NodeName")
        ET.SubElement(node, "I").text = socket.gethostname()
        ET.SubElement(node, "IP").text = "0.0.0.0"
        ET.SubElement(root, "Flags").text = str(flags)
        ET.SubElement(root, "LCID").text = "1033"
        return ET.tostring(root, encoding="unicode", short_empty_elements=False)

    def _login_fields(self, session: str, params: str):
        return [("strTenant", self.tenant), ("type", CLIENT_TYPE), ("sessionID", session),
                ("connectResult", ""), ("isStreamingAllowed", "false"), ("retADOS", 0),
                ("retDCOM", 0), ("connectParams", params)]

    def _check_login(self, out: dict[str, str]):
        for name in ("Connect11LoginExResult", "retADOS", "retDCOM"):
            code = int(out.get(name) or 0)
            if code:
                known = LOGIN_ERRORS.get(code & 0xFFFFFFFF)
                if known:
                    raise XMLServerError(f"Login failed: {known}")
                raise ResultError(f"Login ({name})", code)

    def login(self):
        """Password challenge login: pre-auth, challenge, then RSA-encrypted proof."""
        self.session_id = None
        self._log(f"Connecting to {self.url} as {self.username} ...")
        self.call("Connect11Ex", [
            ("SessionId", ZERO_GUID), ("connectResult", ""), ("isStreamingAllowed", "false"),
            ("retADOS", 0), ("retDCOM", 0), ("strTenant", self.tenant), ("type", CLIENT_TYPE),
            ("connectParams", self._params(2, with_user=False))], sec=True)
        challenge = self.call("Connect11LoginEx", self._login_fields(ZERO_GUID, self._params(4)))
        self._check_login(challenge)
        provisional = str(uuid.UUID(challenge["sessionID"]))
        conres = ET.fromstring(challenge["connectResult"])
        proof = _login_proof(conres.findtext("Chlg") or "", self.password, conres.findtext("Key") or "")
        final = self.call("Connect11LoginEx", self._login_fields(provisional, self._params(34, proof)))
        self._check_login(final)
        session = str(uuid.UUID(final["sessionID"]))
        result = ET.fromstring(final["connectResult"])
        user_info = result.find("UserInfo")
        if session == ZERO_GUID or user_info is None or not len(user_info):
            raise XMLServerError("Login did not return an authenticated session.")
        self.session_id = session
        self.login_result = result
        self._log("Logged in.")

    def disconnect(self):
        if not self.session_id:
            return
        try:
            self.call("Disconnect", [("sessionID", self.session_id)], timeout=30)
        except (XMLServerError, OSError):
            pass
        self.session_id = None


# ---------------------------------------------------------------------------
# Configuration export
# ---------------------------------------------------------------------------
_TIMEOUT_MESSAGE = (
    "The server did not finish building the configuration export before the connection "
    "was closed (Therefore Online closes requests after about five minutes). This tenant "
    "is probably too large for a direct export. Try again later, or upload an XML export instead."
)


def _do_import_export(client: XMLServerClient, request_xml: str) -> ET.Element:
    try:
        out = client.invoke("DoImportExport", [("response", ""), ("sessionID", client.session_id),
                                               ("request", request_xml)])
    except (TimeoutError, socket.timeout, ConnectionError, http.client.RemoteDisconnected) as exc:
        raise ExportTimeoutError(_TIMEOUT_MESSAGE) from exc
    except XMLServerError as exc:
        if "HTTP 502" in str(exc) or "HTTP 504" in str(exc):
            raise ExportTimeoutError(_TIMEOUT_MESSAGE) from exc
        raise
    return ET.fromstring(out["response"])


def _export_tree_request(include_role_assignments: bool) -> str:
    return (
        '<Request FactoryType="1"><ExportRoot><ObjType>17</ObjType><ObjNo>0</ObjNo>'
        "<SubObjNo>1000</SubObjNo><UIObjType>0</UIObjType><ObjName></ObjName>"
        f"<Guid>{ZERO_GUID}</Guid><Target></Target><Flags>0</Flags><Sel>0</Sel>"
        "<Childs/><Actions/><SecretsObjects/></ExportRoot>"
        f"<ExportRoleAssignments>{int(include_role_assignments)}</ExportRoleAssignments>"
        "<EnableLogging>0</EnableLogging>"
        "<Debug><Logs/><ExceptOccur>0</ExceptOccur><OldImportVersion>0</OldImportVersion></Debug>"
        "</Request>\r\n"
    )


def export_configuration(client: XMLServerClient, include_role_assignments: bool = True,
                         log_fn=None) -> str:
    """
    Export the full configuration as TheConfiguration XML text, selecting every
    object the server offers (as Solution Designer does when nothing is unticked).
    Raises ExportTimeoutError when the server takes too long.
    """
    log = log_fn or (lambda _msg: None)

    log("Server is building the export tree (this can take a few minutes) ...")
    started = time.time()
    tree = _do_import_export(client, _export_tree_request(include_role_assignments))
    exp_objects = tree.find("ExpObjects")
    if exp_objects is None:
        raise XMLServerError("Export tree response did not contain ExpObjects.")
    log(f"Export tree built: {len(exp_objects.findall('.//C'))} items in {time.time() - started:.0f}s.")

    # Return the tree with everything selected, as Solution Designer does on Save.
    exp_objects.find("ExportRoleAssignments").text = str(int(include_role_assignments))
    request = ET.Element("Request", FactoryType="2")
    request.append(exp_objects)
    request.append(tree.find("OriginalExpObjects"))
    request.append(tree.find("Debug"))
    ET.SubElement(request, "EnableLogging").text = "0"
    # The server rejects an empty password (0xC0000001). It only protects secrets
    # inside the export, which documentation never needs, so a throwaway value is used.
    ET.SubElement(request, "EncryptionPassword").text = secrets.token_urlsafe(16)

    log("Server is exporting the configuration ...")
    started = time.time()
    result = _do_import_export(client, ET.tostring(request, encoding="unicode") + "\r\n")
    xml_text = "".join(line.text or "" for line in result.findall("ConfigLines/Line"))
    if not xml_text.strip():
        raise XMLServerError("Export response did not contain configuration XML.")
    log(f"Configuration exported ({len(xml_text) / 1e6:.1f} MB) in {time.time() - started:.0f}s.")
    return xml_text


# ---------------------------------------------------------------------------
# Server details (equivalent of the REST-based fetch_server_info)
# ---------------------------------------------------------------------------
def _get_settings(client: XMLServerClient, keys: list[int]) -> dict[int, str]:
    def request(batch):
        keys_xml = "<SettingKeys><Values>" + "".join(f"<V>int:{k}</V>" for k in batch) + "</Values></SettingKeys>"
        out = client.invoke("GetSettings", [("values", ""), ("sessionID", client.session_id), ("keys", keys_xml)])
        values = [v.text for v in ET.fromstring(out["values"]).findall("Values/V")]
        return {k: _typed(v) for k, v in zip(batch, values)}

    try:
        return request(keys)
    except XMLServerError:
        # One unreadable key can fail the whole batch; fall back to one key per call.
        found = {}
        for key in keys:
            try:
                found.update(request([key]))
            except XMLServerError:
                pass
        return found


def _server_version(client: XMLServerClient) -> str:
    """GetServerInfo: first call starts the build; poll until BuildStat=1."""
    stamp = "<TimeStamp><ObjectNo>0</ObjectNo></TimeStamp>\r\n"
    for attempt, mode in enumerate([1, 2, 2, 2, 2, 2]):
        out = client.invoke("GetServerInfo", [("sessionID", client.session_id), ("nType", mode),
                                              ("strTimeStamp", stamp), ("strServerInfo", "")])
        info = ET.fromstring(out["strServerInfo"])
        if info.findtext("BuildStat") == "1":
            packed = int(info.findtext("LocSvr/Version") or 0)
            if packed and packed != 0x7F00FF00:
                return f"{packed >> 24}.{(packed >> 16) & 0xFF}.{packed & 0xFF}"
            return ""
        if info.findtext("BuildStat") == "3":
            return ""
        if attempt:
            time.sleep(1)
    return ""


def fetch_server_info(client: XMLServerClient, server_url: str) -> dict:
    """
    Collect the Server Configuration details over TheXMLServer.
    Returns the same keys as fetch_server_info.fetch_server_info (REST) where available.
    """
    from fetch_server_info import SETTINGS, parse_settings

    info: dict = {"api_url": server_url.rstrip("/"), "api_tenant": client.tenant}

    def attempt(label, fn):
        try:
            fn()
        except (XMLServerError, OSError, ET.ParseError, ValueError) as exc:
            client._log(f"Warning: could not read {label} — {exc}")

    def license_info():
        lic = ET.fromstring(client.invoke("GetLicenseInfoEx", [("licenseInfo", ""),
                                                               ("sessionID", client.session_id)])["licenseInfo"])
        info["product_name"] = lic.findtext("ProductName") or ""
        info["licensee"] = lic.findtext("Licensee") or ""
        maint = lic.findtext("Maintenance") or ""
        info["maintenance_until"] = f"{maint[:4]}-{maint[4:6]}-{maint[6:8]}" if len(maint) == 8 else maint
        license_no = lic.findtext("LicenseNo") or ""
        # Solution Designer shows the customer ID as the first ten characters of the licence number.
        if len(license_no) == 20:
            info["customer_id"] = license_no[:10]

    def version():
        info["server_version"] = _server_version(client)

    def domains():
        out = client.invoke("GetDomainNames", [("xml", ""), ("sessionID", client.session_id), ("nFlags", 2)])
        info["domain_names"] = [_typed(v.text) for v in ET.fromstring(out["xml"]).iter("V") if v.text]

    def settings():
        info.update(parse_settings(_get_settings(client, list(SETTINGS))))

    attempt("licence information", license_info)
    attempt("server version", version)
    attempt("domain names", domains)
    attempt("server settings", settings)
    return {k: v for k, v in info.items() if v not in ("", None, [])}
