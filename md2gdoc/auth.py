"""OAuth2 authentication flow for md2gdoc.

Handles the full setup: create credentials, run browser-based OAuth flow,
save token for future use. Users run `md2gdoc auth` once, then they're set.
"""

import json
import os
import sys
from pathlib import Path

# Scopes needed for md2gdoc operations
SCOPES = [
    "https://www.googleapis.com/auth/documents",       # Create/edit docs
    "https://www.googleapis.com/auth/drive.file",       # Upload diagram images
    "https://www.googleapis.com/auth/presentations.readonly",  # Export slides
]

TOKEN_DIR = Path.home() / ".config" / "md2gdoc"
TOKEN_PATH = TOKEN_DIR / "token.json"
CLIENT_SECRET_PATH = TOKEN_DIR / "client_secret.json"

SETUP_GUIDE = """\
To use md2gdoc, you need Google Cloud OAuth2 credentials.

One-time setup (takes ~3 minutes):

  1. Go to https://console.cloud.google.com/apis/credentials
  2. Create a project (or select an existing one)
  3. Enable these APIs (search in "Enabled APIs & services" > "Enable APIs"):
     - Google Docs API
     - Google Drive API
     - Google Slides API  (only needed for diagram embedding)
  4. Go to "Credentials" > "Create Credentials" > "OAuth client ID"
     - Application type: Desktop app
     - Name: md2gdoc (or anything)
  5. Download the JSON file (click the download icon)
  6. Run:  md2gdoc auth --client-secret /path/to/downloaded.json

This saves your credentials to ~/.config/md2gdoc/ and opens a browser
to authorize the app. You only need to do this once.
"""


def run_auth_flow(client_secret_path: str | None = None):
    """Run the full OAuth2 authentication flow.

    1. Locate or accept client_secret.json
    2. Run InstalledAppFlow (opens browser)
    3. Save token to ~/.config/md2gdoc/token.json
    """
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("Error: google-auth-oauthlib is required for authentication.", file=sys.stderr)
        print("Install it:  pip install google-auth-oauthlib", file=sys.stderr)
        sys.exit(1)

    # Find client secret
    secret_path = _resolve_client_secret(client_secret_path)
    if not secret_path:
        print(SETUP_GUIDE)
        sys.exit(1)

    print(f"Using client secret: {secret_path}")
    print("Opening browser for Google authorization...\n")

    # Run the OAuth flow
    flow = InstalledAppFlow.from_client_secrets_file(str(secret_path), SCOPES)

    try:
        creds = flow.run_local_server(
            port=0,
            prompt="consent",
            access_type="offline",
        )
    except Exception as e:
        print(f"\nBrowser auth failed: {e}", file=sys.stderr)
        print("Trying console-based auth instead...\n", file=sys.stderr)
        creds = flow.run_console()

    # Save token
    TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    token_data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }

    with open(TOKEN_PATH, "w") as f:
        json.dump(token_data, f, indent=2)
    os.chmod(TOKEN_PATH, 0o600)

    # Copy client secret to config dir for future re-auth
    if secret_path != CLIENT_SECRET_PATH:
        import shutil
        shutil.copy2(secret_path, CLIENT_SECRET_PATH)
        os.chmod(CLIENT_SECRET_PATH, 0o600)

    print(f"Token saved to {TOKEN_PATH}")
    print(f"Client secret saved to {CLIENT_SECRET_PATH}")
    print("\nYou're all set. Try:  md2gdoc input.md --style clean")


def check_auth() -> bool:
    """Check if valid credentials exist. Returns True if authenticated."""
    if not TOKEN_PATH.exists():
        return False
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request

        with open(TOKEN_PATH) as f:
            td = json.load(f)
        creds = Credentials(
            token=td.get("access_token") or td.get("token"),
            refresh_token=td.get("refresh_token"),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=td.get("client_id"),
            client_secret=td.get("client_secret"),
        )
        creds.refresh(Request())
        return True
    except Exception:
        return False


def revoke_auth():
    """Revoke stored credentials and delete token file."""
    if TOKEN_PATH.exists():
        # Try to revoke the token with Google (POST body, not URL param)
        try:
            with open(TOKEN_PATH) as f:
                td = json.load(f)
            token = td.get("token") or td.get("access_token")
            if token:
                import urllib.request
                import urllib.parse
                data = urllib.parse.urlencode({"token": token}).encode("utf-8")
                req = urllib.request.Request(
                    "https://oauth2.googleapis.com/revoke",
                    data=data,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                )
                urllib.request.urlopen(req, timeout=10)
        except Exception:
            pass  # Best effort

        TOKEN_PATH.unlink()
        print(f"Removed {TOKEN_PATH}")
    else:
        print("No stored credentials found.")

    if CLIENT_SECRET_PATH.exists():
        CLIENT_SECRET_PATH.unlink()
        print(f"Removed {CLIENT_SECRET_PATH}")

    print("Credentials revoked. Run 'md2gdoc auth' to re-authenticate.")


def _resolve_client_secret(explicit_path: str | None) -> Path | None:
    """Find client_secret.json from various sources."""
    candidates = []

    if explicit_path:
        p = Path(explicit_path)
        if p.exists():
            return p
        print(f"Error: {explicit_path} not found", file=sys.stderr)
        return None

    # Check config dir
    candidates.append(CLIENT_SECRET_PATH)

    # Check env var
    env_path = os.environ.get("GOOGLE_CLIENT_SECRET")
    if env_path:
        candidates.append(Path(env_path))

    # Check current directory
    candidates.append(Path("client_secret.json"))
    candidates.append(Path("credentials.json"))

    for p in candidates:
        if p.exists():
            return p

    return None
