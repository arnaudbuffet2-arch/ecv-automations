"""
Re-autorise Google avec tous les scopes (Calendar + Sheets + Gmail).
Lance un navigateur pour le consentement OAuth.

Usage : python "claude safe/reauth_google.py"
"""
import json
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

ECV_DIR = Path.home() / ".ecv"
CREDENTIALS_PATH = ECV_DIR / "credentials.json"
TOKEN_PATH = ECV_DIR / "tokens.json"

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/documents.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
creds = flow.run_local_server(port=8081)

tokens = {
    "access_token":  creds.token,
    "refresh_token": creds.refresh_token,
}
TOKEN_PATH.write_text(json.dumps(tokens, indent=2))
print(f"Tokens sauvegardés dans {TOKEN_PATH}")
