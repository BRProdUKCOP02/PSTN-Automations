"""
oauth_setup.py — One-time setup script to obtain an Adobe Sign OAuth refresh token.

Run this ONCE from a machine with a browser. It will:
  1. Print the authorisation URL (and attempt to open it automatically)
  2. Start a temporary local HTTPS server to catch Adobe Sign's redirect
  3. Exchange the authorisation code for access + refresh tokens
  4. Print the refresh token and write it to .env automatically

After this script completes successfully, you never need to run it again unless:
  - You revoke the application in Adobe Sign
  - The refresh token is unused for more than 60 days

Usage:
    python oauth_setup.py
"""
import http.server
import os
import ssl
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Load client credentials from .env (before config.py, which requires all
# fields to be set — REFRESH_TOKEN won't be set yet on first run)
# ---------------------------------------------------------------------------
from dotenv import load_dotenv, set_key

_ENV_PATH = Path(__file__).parent / ".env"
load_dotenv(_ENV_PATH)

CLIENT_ID     = os.getenv("ADOBE_SIGN_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("ADOBE_SIGN_CLIENT_SECRET", "")
REDIRECT_URI  = os.getenv("ADOBE_SIGN_REDIRECT_URI", "https://localhost")
BASE_URL      = os.getenv("ADOBE_SIGN_BASE_URL", "https://api.na1.adobesign.com")

# Scopes must EXACTLY match what is ticked in the Adobe Sign app Configure OAuth page.
# All 9 ticked scopes with :group modifier (to access all group data).
SCOPES = "+".join([
    "user_login:group",
    "user_read:group",
    "user_write:group",
    "agreement_read:group",
    "agreement_write:group",
    "agreement_send:group",
    "widget_read:group",
    "library_read:group",
    "workflow_read:group",
])

# ---------------------------------------------------------------------------
# Validate prerequisites
# ---------------------------------------------------------------------------

def _check_config() -> None:
    missing = []
    if not CLIENT_ID or CLIENT_ID.startswith("YOUR_"):
        missing.append("ADOBE_SIGN_CLIENT_ID")
    if not CLIENT_SECRET or CLIENT_SECRET.startswith("YOUR_"):
        missing.append("ADOBE_SIGN_CLIENT_SECRET")
    if missing:
        print(f"\n❌  Missing in .env: {', '.join(missing)}")
        print(f"   Edit: {_ENV_PATH}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Local redirect-catcher server
# Adobe Sign redirects to http://localhost:8080?code=...
# Plain HTTP on a high port — no SSL needed, no admin rights required.
# ---------------------------------------------------------------------------

_auth_code: str = ""
_server_error: str = ""

_REDIRECT_PORT = 8080


class _RedirectHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        global _auth_code, _server_error
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)

        if "code" in params:
            _auth_code = params["code"][0]
            body = b"<html><body><h2>Authorisation successful!</h2><p>You can close this tab and return to the terminal.</p></body></html>"
        elif "error" in params:
            _server_error = params.get("error_description", params.get("error", ["unknown"]))[0]
            body = b"<html><body><h2>Authorisation failed.</h2><p>Check the terminal for details.</p></body></html>"
        else:
            body = b"<html><body><p>Waiting...</p></body></html>"

        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass  # suppress server access logs


def _make_self_signed_cert() -> tuple[str, str]:
    """Generate a temporary self-signed cert for https://localhost:8080."""
    try:
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509.oid import NameOID
        import datetime

        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        subject = issuer = x509.Name([
            x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        ])
        cert = (
            x509.CertificateBuilder()
            .subject_name(subject)
            .issuer_name(issuer)
            .public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(datetime.datetime.utcnow())
            .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(hours=1))
            .add_extension(
                x509.SubjectAlternativeName([x509.DNSName("localhost")]),
                critical=False,
            )
            .sign(key, hashes.SHA256())
        )
        tmp_dir = tempfile.mkdtemp()
        cert_path = os.path.join(tmp_dir, "cert.pem")
        key_path = os.path.join(tmp_dir, "key.pem")
        with open(cert_path, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))
        with open(key_path, "wb") as f:
            f.write(key.private_bytes(
                serialization.Encoding.PEM,
                serialization.PrivateFormat.TraditionalOpenSSL,
                serialization.NoEncryption(),
            ))
        return cert_path, key_path
    except ImportError:
        return "", ""


def _start_redirect_server() -> http.server.HTTPServer:
    """Start an HTTPS server on port 8080 to catch the OAuth redirect."""
    cert_path, key_path = _make_self_signed_cert()
    server = http.server.HTTPServer(("localhost", _REDIRECT_PORT), _RedirectHandler)
    if cert_path:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(cert_path, key_path)
        server.socket = context.wrap_socket(server.socket, server_side=True)
    return server


# ---------------------------------------------------------------------------
# Token exchange
# ---------------------------------------------------------------------------

def _exchange_code(code: str) -> dict:
    # Use the shard-specific token endpoint.
    token_url = f"{BASE_URL}/oauth/v2/token"
    payload = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "code": code,
    }
    response = requests.post(token_url, data=payload, timeout=30)
    if not response.ok:
        print(f"\n❌  Token exchange failed: HTTP {response.status_code}")
        print(response.text)
        sys.exit(1)
    return response.json()


# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------

def _auth_host_from_base_url(base_url: str) -> str:
    """
    Derive the browser-facing OAuth host from the API base URL.
    e.g. https://api.na1.adobesign.com  ->  secure.na1.adobesign.com
         https://api.eu1.adobesign.com  ->  secure.eu1.adobesign.com
    """
    # Strip scheme and 'api.' prefix to get the shard hostname
    host = base_url.replace("https://", "").replace("http://", "")
    host = host.replace("api.", "secure.", 1)
    return host.rstrip("/")


def main():
    _check_config()

    # EU2 shard (user's actual shard: secure.eu2.adobesign.com)
    encoded_redirect = urllib.parse.quote(REDIRECT_URI, safe="")
    # SCOPES already uses + as separator; keep : unencoded as Adobe Sign expects
    encoded_scopes = "+".join(urllib.parse.quote(s, safe=":") for s in SCOPES.split("+"))
    auth_host = _auth_host_from_base_url(BASE_URL)
    auth_url = (
        f"https://{auth_host}/public/oauth/v2"
        f"?response_type=code"
        f"&client_id={CLIENT_ID}"
        f"&redirect_uri={encoded_redirect}"
        f"&scope={encoded_scopes}"
    )

    print("=" * 65)
    print("  Adobe Sign OAuth Setup — One-Time Token Acquisition")
    print("=" * 65)
    print()
    print("  PRE-FLIGHT CHECK — before continuing, confirm:")
    print(f"  1. In the Adobe Sign developer console, your app's")
    print(f"     Redirect URI is set to exactly:  {REDIRECT_URI}")
    print( "     (Account > Adobe Sign API > Applications > Configure OAuth)")
    print( "  2. The following scopes are enabled (account modifier):")
    print( "     agreement_read, agreement_write, agreement_send,")
    print( "     library_read, user_read")
    print()
    print("  Press Enter to continue, or Ctrl+C to cancel and fix the app first.")
    try:
        input("  > ")
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(0)

    print("\nStep 1: Authorise the application in your browser.")
    print("        (The browser will open automatically in 3 seconds.)\n")
    print(f"  Auth host : secure.eu2.adobesign.com")
    print(f"  URL       : {auth_url}\n")

    # Try to start the redirect-catching server
    server = None
    use_server = False
    try:
        server = _start_redirect_server()
        use_server = True
        t = threading.Thread(target=server.handle_request, daemon=True)
        t.start()
        print(f"  Local redirect server started on http://localhost:{_REDIRECT_PORT}")
        print("  After approving in your browser, it will redirect here automatically.\n")
    except PermissionError:
        print(f"  ⚠️  Could not start local server on port {_REDIRECT_PORT}.")
        print("     After authorising in the browser, Adobe Sign will redirect to")
        print(f"     http://localhost:{_REDIRECT_PORT}?code=... — copy the full URL from your browser")
        print("     address bar and paste it below.\n")

    import time
    time.sleep(3)
    webbrowser.open(auth_url)

    if use_server:
        print("Waiting for Adobe Sign to redirect to localhost...")
        t.join(timeout=120)
        if server:
            server.server_close()
        if _server_error:
            print(f"\n❌  Adobe Sign returned an error: {_server_error}")
            sys.exit(1)
        if not _auth_code:
            print("\n⚠️  Auto-capture timed out. Switching to manual method.")
            print("   In your browser, after approving the app, the address bar will show")
            print("   a URL starting with https://localhost:8080/?code=...")
            print("   Copy that full URL and paste it below (even if the page shows an error).\n")
            redirect_response = input("Paste the full redirect URL here:\n> ").strip()
            parsed = urllib.parse.urlparse(redirect_response)
            params = urllib.parse.parse_qs(parsed.query)
            if "code" not in params:
                print("❌  Could not find 'code' in the URL. Ensure you copied the full URL.")
                sys.exit(1)
            code = params["code"][0]
        else:
            code = _auth_code
    else:
        # Manual fallback — user pastes the redirect URL
        redirect_response = input("Paste the full redirect URL here:\n> ").strip()
        parsed = urllib.parse.urlparse(redirect_response)
        params = urllib.parse.parse_qs(parsed.query)
        if "code" not in params:
            print("❌  Could not find 'code' in the URL. Ensure you copied the full URL.")
            sys.exit(1)
        code = params["code"][0]

    print("\nStep 2: Exchanging authorisation code for tokens...")
    tokens = _exchange_code(code)

    access_token = tokens.get("access_token", "")
    refresh_token = tokens.get("refresh_token", "")
    api_access_point = tokens.get("api_access_point", BASE_URL).rstrip("/")

    if not refresh_token:
        print(f"\n❌  No refresh token in response: {tokens}")
        sys.exit(1)

    print("\n" + "=" * 65)
    print("  ✅  SUCCESS")
    print("=" * 65)
    print(f"\n  Refresh Token : {refresh_token}")
    print(f"  API Base URL  : {api_access_point}")

    # Write to .env automatically
    set_key(str(_ENV_PATH), "ADOBE_SIGN_REFRESH_TOKEN", refresh_token)
    set_key(str(_ENV_PATH), "ADOBE_SIGN_BASE_URL", api_access_point)

    print(f"\n  ✅  Written to: {_ENV_PATH}")
    print("      ADOBE_SIGN_REFRESH_TOKEN and ADOBE_SIGN_BASE_URL updated.\n")
    print("  You can now run the bot. This script does not need to be run again.")
    print("=" * 65)


if __name__ == "__main__":
    main()
