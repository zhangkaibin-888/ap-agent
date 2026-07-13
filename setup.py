#!/usr/bin/env python3
"""
Invoice-to-Xero Bridge — Setup Wizard
=======================================
Interactive configuration wizard that guides you through:

1. Application identity and basic settings
2. Microsoft Graph (email) — Azure App Registration
3. LLM provider for AI invoice parsing
4. Xero API credentials and OAuth setup
5. Invoice processing preferences
6. Testing connections
7. PM2 deployment setup

Usage:
  python3 setup.py
"""

import json
import os
import sys
import subprocess
from pathlib import Path

# ── Helpers ───────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
ECOSYSTEM_PATH = SCRIPT_DIR / "ecosystem.config.js"


def load_config():
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH) as f:
            return json.load(f)
    return None


def save_config(cfg):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=4)
    print(f"  ✓ Saved config to {CONFIG_PATH}")


def prompt(label, default="", secret=False):
    """Prompt for a value with optional default."""
    display_default = f" [{default}]" if default else ""
    prompt_text = f"{label}{display_default}: "
    value = input(prompt_text).strip()
    if not value and default:
        return default
    return value


def prompt_required(label):
    """Prompt for a required value."""
    while True:
        value = input(f"{label}: ").strip()
        if value:
            return value
        print("  This field is required.")


def prompt_int(label, default=None):
    """Prompt for an integer."""
    display = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"{label}{display}: ").strip()
        if not raw and default is not None:
            return default
        try:
            return int(raw)
        except ValueError:
            print("  Please enter a number.")


def prompt_boolean(label, default=True):
    """Prompt for yes/no."""
    hint = " [Y/n]" if default else " [y/N]"
    raw = input(f"{label}{hint}: ").strip().lower()
    if not raw:
        return default
    return raw in ("y", "yes")


def print_section(title):
    """Print a section header."""
    print()
    print("=" * 60)
    print(f"  {title}")
    print("=" * 60)
    print()


def test_llm_connection(cfg):
    """Test connectivity to the configured LLM provider."""
    llm = cfg.get("llm", {})
    provider = llm.get("provider", "openai")
    api_key = os.environ.get(llm.get("api_key_env", "LLM_API_KEY"), "")
    model = llm.get("model", "gpt-4o-mini")

    if not api_key:
        # Try the env var name from config
        env_var = llm.get("api_key_env", "LLM_API_KEY")
        print(f"\n  ⚠ {env_var} environment variable is not set.")
        set_now = prompt_boolean(f"  Set {env_var} now?", default=True)
        if set_now:
            key = input(f"  Enter your {provider} API key: ").strip()
            if key:
                os.environ[env_var] = key
                api_key = key
                print(f"  ✓ Set {env_var} for this session.")
                print(f"    Add 'export {env_var}=<your-key>' to your ~/.bashrc")
        else:
            print("  Skipping LLM connection test.")
            return False

    print(f"\n  Testing {provider} connection...")
    
    if provider == "anthropic":
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }
        payload = {
            "model": model,
            "max_tokens": 10,
            "messages": [{"role": "user", "content": "Say OK"}],
        }
    else:
        base_url = llm.get("base_url", "https://api.openai.com/v1").rstrip("/")
        url = f"{base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": "Say OK"}],
            "max_tokens": 10,
        }

    try:
        import requests
        resp = requests.post(url, headers=headers, json=payload, timeout=15)
        if resp.ok:
            print(f"  ✓ {provider} connection OK (model: {model})")
            return True
        else:
            print(f"  ✗ {provider} error: {resp.status_code} {resp.text[:200]}")
            return False
    except ImportError:
        print("  ⚠ requests not installed. Install with: pip install requests")
        return False
    except Exception as e:
        print(f"  ✗ Connection failed: {e}")
        return False


def test_graph_connection(cfg):
    """Test Microsoft Graph connectivity."""
    g = cfg.get("graph", {})
    if not g.get("tenant_id") or not g.get("client_id") or not g.get("client_secret"):
        print("  ⚠ Graph credentials not fully configured. Skipping test.")
        return False

    print(f"\n  Testing Microsoft Graph connection...")
    try:
        from msal import ConfidentialClientApplication
        import requests
        from datetime import datetime

        app = ConfidentialClientApplication(
            g["client_id"],
            authority=f"https://login.microsoftonline.com/{g['tenant_id']}",
            client_credential=g["client_secret"],
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            print(f"  ✗ Auth failed: {result.get('error_description', 'Unknown error')}")
            return False

        # List 1 email from the shared mailbox to verify
        headers = {"Authorization": f"Bearer {result['access_token']}"}
        url = f"https://graph.microsoft.com/v1.0/users/{g['shared_mailbox']}/messages?$top=1"
        resp = requests.get(url, headers=headers)
        if resp.ok:
            print(f"  ✓ Graph connection OK — can access shared mailbox")
            return True
        else:
            print(f"  ✗ Can't access mailbox: {resp.status_code} {resp.text[:200]}")
            print(f"    Ensure the app has Mail.Read Application permission and admin consent was granted.")
            return False
    except ImportError:
        print("  ⚠ msal not installed. Install with: pip install msal requests")
        return False
    except Exception as e:
        print(f"  ✗ Connection failed: {e}")
        return False


# ── Main Setup Flow ──────────────────────────────────────────────────

def setup_app_identity():
    """Step 1: Application identity and basic settings."""
    print_section("Step 1: Application Identity")

    existing = load_config() or {}

    name = prompt("Application name", default=existing.get("app_name", "InvoiceXero"))
    log_level = prompt("Log level (DEBUG/INFO/WARNING/ERROR)", default=existing.get("log_level", "INFO"))
    log_file = prompt("Log file path", default=existing.get("log_file", "/var/log/invoice-xero-bridge.log"))
    port = prompt_int("HTTP server port", default=existing.get("port", 8088))

    cfg = {
        "app_name": name,
        "log_level": log_level.upper(),
        "log_file": log_file,
        "port": port,
    }
    return cfg


def setup_graph():
    """Step 2: Microsoft Graph configuration."""
    print_section("Step 2: Microsoft Graph (Email Access)")

    print("Create an Azure App Registration at https://portal.azure.com:")
    print("  1. App Registrations → New Registration")
    print("  2. Name it (e.g., 'InvoiceXero Bridge')")
    print("  3. Under API Permissions → Add: Mail.Read, Mail.Send (Application permissions)")
    print("  4. Grant Admin Consent")
    print("  5. Under Certificates & Secrets → New Client Secret")
    print()

    existing = load_config() or {}
    g = existing.get("graph", {})

    tenant_id = prompt("Azure AD Tenant ID", default=g.get("tenant_id", ""))
    client_id = prompt("Azure App Client ID", default=g.get("client_id", ""))
    client_secret = prompt("Azure App Client Secret", default=g.get("client_secret", ""))
    mailbox = prompt("Shared mailbox email address", default=g.get("shared_mailbox", ""))

    if tenant_id and client_id and client_secret and mailbox:
        print("\n  Graph configuration complete.")

    return {
        "tenant_id": tenant_id,
        "client_id": client_id,
        "client_secret": client_secret,
        "shared_mailbox": mailbox,
    }


def setup_llm():
    """Step 3: LLM provider configuration."""
    print_section("Step 3: LLM Provider (AI Invoice Parsing)")

    print("Configure which LLM to use for parsing invoice PDFs.")
    print()
    print("  Supported providers:")
    print("    openai    → OpenAI (GPT-4o-mini, GPT-4o)")
    print("    deepseek  → DeepSeek (deepseek-chat, deepseek-reasoner)")
    print("    anthropic → Anthropic (Claude 3 Haiku, Sonnet)")
    print("    custom    → Any OpenAI-compatible API")
    print()

    existing = load_config() or {}
    llm = existing.get("llm", {})

    provider = prompt("LLM provider", default=llm.get("provider", "openai"))
    model_map = {
        "openai": "gpt-4o-mini",
        "deepseek": "deepseek-chat",
        "anthropic": "claude-3-haiku-20240307",
        "custom": "gpt-4o-mini",
    }
    default_model = llm.get("model", model_map.get(provider, "gpt-4o-mini"))
    model = prompt(f"Model name", default=default_model)
    api_key_env = prompt("Environment variable for API key", default=llm.get("api_key_env", "LLM_API_KEY"))

    base_url = llm.get("base_url", "")
    if provider == "custom":
        base_url = prompt("API base URL (e.g., https://api.openai.com/v1)", default=base_url or "https://api.openai.com/v1")
    elif provider == "deepseek":
        base_url = "https://api.deepseek.com/v1"
    elif provider == "anthropic":
        base_url = "https://api.anthropic.com/v1"
    else:
        base_url = "https://api.openai.com/v1"

    temperature = prompt("Temperature (0.0 = deterministic)", default=str(llm.get("temperature", 0.0)))
    max_tokens = prompt_int("Max tokens in response", default=llm.get("max_tokens", 500))

    return {
        "provider": provider,
        "model": model,
        "api_key_env": api_key_env,
        "base_url": base_url,
        "temperature": float(temperature),
        "max_tokens": max_tokens,
    }


def setup_xero():
    """Step 4: Xero API configuration."""
    print_section("Step 4: Xero Accounting")

    print("Set up your Xero app at https://developer.xero.com:")
    print("  1. My Apps → Create App")
    print("  2. Set redirect URI: http://localhost:8080/callback")
    print("  3. Copy Client ID and Client Secret below")
    print()

    existing = load_config() or {}
    x = existing.get("xero", {})

    client_id = prompt("Xero Client ID", default=x.get("client_id", ""))
    client_secret = prompt("Xero Client Secret", default=x.get("client_secret", ""))
    tenant_id = prompt("Xero Organisation (Tenant) ID", default=x.get("tenant_id", ""))

    print()
    account_code = prompt("Default expense account code", default=x.get("default_account_code", "200"))
    tax_type = prompt("Default tax type (NONE, GST, VAT, etc.)", default=x.get("default_tax_type", "NONE"))

    return {
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": tenant_id,
        "default_account_code": account_code,
        "default_tax_type": tax_type,
    }


def setup_pipeline():
    """Step 5: Invoice processing preferences."""
    print_section("Step 5: Invoice Processing")

    existing = load_config() or {}

    approver = prompt("Approver email address (for approval requests)", default=existing.get("approver_email", ""))
    
    recipients_raw = prompt("Notification recipients (comma-separated emails)", default=", ".join(existing.get("notification_recipients", [])))
    recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]

    logic_app = prompt("Logic App URL (for approval emails)", default=existing.get("logic_app_url", ""))

    parsing_mode = prompt("Parsing mode (regex / llm / hybrid)", default=existing.get("parsing_mode", "hybrid"))

    print("\n  PDF Regex Patterns:")
    print("  You can add custom regex patterns for specific suppliers.")
    pdf_patterns = existing.get("pdf_patterns", {})

    if prompt_boolean("  Add custom patterns?", default=False):
        if not pdf_patterns:
            pdf_patterns = {"invoice_number": [], "total_amount": [], "due_date": [], "supplier_name": []}
        for field in ["invoice_number", "total_amount", "due_date", "supplier_name"]:
            print(f"    {field}:")
            existing_patterns = " | ".join(pdf_patterns.get(field, []))
            new_pattern = prompt(f"      Regex pattern (or leave blank)", default=existing_patterns)
            if new_pattern:
                pdf_patterns[field] = [p.strip() for p in new_pattern.split("|")]

    return {
        "approver_email": approver,
        "notification_recipients": recipients,
        "logic_app_url": logic_app,
        "parsing_mode": parsing_mode,
        "pdf_patterns": pdf_patterns,
    }


def run_xero_oauth():
    """Offer to run the Xero OAuth setup."""
    print_section("Xero OAuth Setup")
    print("Now you can authenticate with Xero to get your access tokens.")
    print()

    if prompt_boolean("Run Xero OAuth setup now?", default=True):
        print()
        oauth_script = SCRIPT_DIR / "setup_xero_auth.py"
        if oauth_script.exists():
            subprocess.run([sys.executable, str(oauth_script)], cwd=str(SCRIPT_DIR))
        else:
            print(f"  ✗ setup_xero_auth.py not found at {oauth_script}")
    else:
        print("  You can run it later with: python3 setup_xero_auth.py")


def create_ecosystem(cfg):
    """Create/update PM2 ecosystem file."""
    print_section("PM2 Ecosystem File")

    app_name = cfg.get("app_name", "InvoiceXero").lower().replace(" ", "-")
    script_path = str(SCRIPT_DIR / "server.py")

    ecosystem = {
        "apps": [{
            "name": f"{app_name}-server",
            "script": script_path,
            "interpreter": "python3",
            "cwd": str(SCRIPT_DIR),
            "env": {
                "PYTHONUNBUFFERED": "1",
            },
            "log_file": f"/var/log/{app_name}-server.log",
            "error_file": f"/var/log/{app_name}-server-error.log",
            "out_file": f"/var/log/{app_name}-server-out.log",
            "max_restarts": 10,
            "restart_delay": 5000,
        }]
    }

    if prompt_boolean("Write PM2 ecosystem.config.js?", default=True):
        with open(SCRIPT_DIR / "ecosystem.config.js", "w") as f:
            # Write JavaScript, not JSON
            f.write("module.exports = ")
            f.write(json.dumps(ecosystem, indent=2))
            f.write(";\n")
        print(f"  ✓ Written to {SCRIPT_DIR}/ecosystem.config.js")
        print()
        print("  Deploy with:")
        print(f"    pm2 start {SCRIPT_DIR}/ecosystem.config.js")
        print(f"    pm2 save")
    else:
        print("  Skipping ecosystem file creation.")


def run_tests(cfg):
    """Run connection tests."""
    print_section("Connection Tests")

    test_llm = prompt_boolean("Test LLM connection?", default=True)
    if test_llm:
        test_llm_connection(cfg)

    test_graph = prompt_boolean("Test Microsoft Graph connection?", default=True)
    if test_graph:
        test_graph_connection(cfg)

    print()
    print("  To test the full pipeline:")
    print(f"    python3 {SCRIPT_DIR / 'pipeline.py'}")
    print("  To start the HTTP server:")
    print(f"    python3 {SCRIPT_DIR / 'server.py'}")
    print()


# ── Main ─────────────────────────────────────────────────────────────

def main():
    print()
    print("=" * 60)
    print("  Invoice-to-Xero Bridge — Setup Wizard")
    print("=" * 60)
    print()
    print("This wizard will help you configure the application step by step.")
    print("Press Enter to accept defaults shown in [brackets].")
    print()

    cfg = load_config()

    # Step-by-step configuration
    identity = setup_app_identity()
    graph = setup_graph()
    llm = setup_llm()
    xero = setup_xero()
    pipeline = setup_pipeline()

    # Merge with existing config if any
    full_cfg = {
        **identity,
        "graph": graph,
        "llm": llm,
        "xero": xero,
        **pipeline,
    }

    # Review
    print_section("Configuration Summary")
    print(f"  Application Name:      {full_cfg['app_name']}")
    print(f"  Log Level:             {full_cfg['log_level']}")
    print(f"  Log File:              {full_cfg['log_file']}")
    print(f"  Port:                  {full_cfg['port']}")
    print(f"  Shared Mailbox:        {full_cfg['graph']['shared_mailbox']}")
    print(f"  LLM Provider:          {full_cfg['llm']['provider']} ({full_cfg['llm']['model']})")
    print(f"  Xero Client ID:        {full_cfg['xero']['client_id'][:8]}...")
    print(f"  Approver Email:        {full_cfg.get('approver_email', '') or '(not set)'}")
    print(f"  Notifications:         {len(full_cfg.get('notification_recipients', []))} recipient(s)")
    print(f"  Parsing Mode:          {full_cfg.get('parsing_mode', 'hybrid')}")
    print()

    if prompt_boolean("Save this configuration?", default=True):
        save_config(full_cfg)
        print()
        run_xero_oauth()
        create_ecosystem(full_cfg)
        run_tests(full_cfg)

        print_section("Setup Complete!")
        print(f"  ✓ Config saved to:    {CONFIG_PATH}")
        print(f"  ✓ Next steps:")
        print(f"    1. Run the pipeline: python3 pipeline.py")
        print(f"    2. Start the server: python3 server.py")
        print(f"    3. Deploy with PM2:  pm2 start ecosystem.config.js")
        print(f"    4. Set up cron:      crontab -e")
        print(f"       Add: 0 9 * * * cd {SCRIPT_DIR} && python3 cron_trigger.py")
        print()
    else:
        print("  Configuration not saved.")


if __name__ == "__main__":
    main()
