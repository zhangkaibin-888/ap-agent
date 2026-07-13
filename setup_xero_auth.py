#!/usr/bin/env python3
"""
Xero OAuth Setup Script
========================
Run this ONCE to authenticate with Xero. It will:
1. Open a browser for you to log into Xero
2. Save the tokens to xero_tokens.json
3. The pipeline and server will refresh tokens automatically after that

Prerequisites:
  pip install requests
  You have a Xero App created at https://developer.xero.com
  Your app has redirect URI: http://localhost:8080/callback
"""

import json
import webbrowser
from pathlib import Path
from urllib.parse import urlencode, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests

CONFIG_PATH = Path(__file__).parent / "config.json"
TOKEN_PATH = Path(__file__).parent / "xero_tokens.json"

if not CONFIG_PATH.exists():
    print("FATAL: config.json not found. Create it first.")
    exit(1)

with open(CONFIG_PATH) as f:
    cfg = json.load(f)

APP_NAME = cfg.get("app_name", "InvoiceXero")
CLIENT_ID = cfg["xero"]["client_id"]
CLIENT_SECRET = cfg["xero"]["client_secret"]
REDIRECT_URI = "http://localhost:8080/callback"
SCOPES = "openid profile email accounting.transactions accounting.contacts accounting.attachments offline_access"


# ── Simple local HTTP server to catch the callback ──
auth_code = None

class CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        qs = parse_qs(self.path.split("?")[1] if "?" in self.path else "")
        auth_code = qs.get("code", [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        if auth_code:
            self.wfile.write(b"<h1>Authorization successful!</h1><p>Close this tab and return to the terminal.</p>")
        else:
            self.wfile.write(b"<h1>Authorization failed.</h1>")
        # Shutdown server after response
        import threading
        threading.Thread(target=self.server.shutdown).start()

    def log_message(self, format, *args):
        pass  # Suppress HTTP log noise


def get_auth_code():
    """Open browser for user to authorize, then catch the callback."""
    global auth_code
    params = urlencode({
        "response_type": "code",
        "client_id": CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
        "state": f"{APP_NAME.lower()}_setup",
    })
    auth_url = f"https://login.xero.com/identity/connect/authorize?{params}"
    print(f"Opening browser to authorize with Xero...")
    print(f"If browser doesn't open, visit:\n{auth_url}\n")
    webbrowser.open(auth_url)

    server = HTTPServer(("localhost", 8080), CallbackHandler)
    server.timeout = 300  # 5 min timeout
    server.handle_request()

    if not auth_code:
        print("Authorization failed or timed out.")
        exit(1)
    return auth_code


def exchange_code(code):
    """Exchange auth code for tokens."""
    print("Exchanging authorization code for tokens...")
    resp = requests.post(
        "https://identity.xero.com/connect/token",
        data={
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": REDIRECT_URI,
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
        },
    )
    resp.raise_for_status()
    tokens = resp.json()
    tokens["expires_at"] = __import__("time").time() + tokens.get("expires_in", 3600) - 60
    return tokens


def get_tenants(access_token):
    """Get the Xero organisations (tenants) you have access to."""
    resp = requests.get(
        "https://api.xero.com/connections",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    resp.raise_for_status()
    return resp.json()


def main():
    print("=" * 60)
    print(f"{APP_NAME} — Xero OAuth Setup")
    print("=" * 60)
    print()
    print(f"Client ID: {CLIENT_ID[:8]}...{CLIENT_ID[-4:]}")
    print(f"Redirect URI: {REDIRECT_URI}")
    print()

    # Step 1: Get auth code
    code = get_auth_code()
    print(f"✓ Authorization code received")

    # Step 2: Exchange for tokens
    tokens = exchange_code(code)
    print(f"✓ Tokens received")

    # Step 3: Show available tenants
    tenants = get_tenants(tokens["access_token"])
    print(f"\nXero Organisations you have access to:")
    for i, t in enumerate(tenants, 1):
        print(f"  {i}. {t['tenantName']} (ID: {t['tenantId']}) — {t.get('orgType', '')}")

    if len(tenants) > 1:
        print(f"\nUsing first tenant by default: {tenants[0]['tenantName']}")
    tenant_id = tenants[0]["tenantId"] if tenants else "UNKNOWN"

    # Step 4: Save tokens
    with open(TOKEN_PATH, "w") as f:
        json.dump(tokens, f, indent=2)
    print(f"\n✓ Tokens saved to {TOKEN_PATH}")

    # Step 5: Update config.json with tenant_id
    cfg["xero"]["tenant_id"] = tenant_id
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=4)
    print(f"✓ Tenant ID written to config.json")

    print(f"\n{'='*60}")
    print(f"Setup complete! The pipeline can now connect to Xero.")
    print(f"Run the pipeline manually to test:")
    print(f"  python3 {Path(__file__).parent / 'pipeline.py'}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
