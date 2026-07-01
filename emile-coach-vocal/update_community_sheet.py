"""
Met à jour l'onglet 'community management' du Google Sheet ECV
avec les stats du mois précédent.

Usage : python scripts/update_community_sheet.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "claude safe"))
from google_auth import get_google_credentials
from googleapiclient.discovery import build

SHEET_ID = "1OReCrVznxOtrxTzSRqpsKEu0lvR0cwdPoCixFKhgyxs"
TAB = "community management"

# Ligne de base pour chaque section (row = base + numero_mois)
# Exemple : avril = mois 4 → TikTok row = 4+4=8, Instagram = 21+4=25, YouTube = 38+4=42
ROW_BASE = {"tiktok": 4, "instagram": 21, "youtube": 38}


def update_month(month_num: int, tiktok: list, instagram: list, youtube: list):
    """
    month_num : 1=jan, 2=fev, ... 12=dec
    tiktok/instagram/youtube : [abonnes, vues, commentaires, partages, likes]
    """
    creds = get_google_credentials()
    service = build("sheets", "v4", credentials=creds)

    row_tt = ROW_BASE["tiktok"] + month_num
    row_ig = ROW_BASE["instagram"] + month_num
    row_yt = ROW_BASE["youtube"] + month_num

    updates = [
        {"range": f"{TAB}!C{row_tt}:G{row_tt}", "values": [tiktok]},
        {"range": f"{TAB}!C{row_ig}:G{row_ig}", "values": [instagram]},
        {"range": f"{TAB}!C{row_yt}:G{row_yt}", "values": [youtube]},
    ]

    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": updates},
    ).execute()

    print(f"OK - {result.get('totalUpdatedCells')} cellules mises a jour (mois {month_num})")


if __name__ == "__main__":
    # Mai 2026 — mois complet (TikTok/Instagram 2 mai→1er juin, YouTube 1→31 mai)
    update_month(
        month_num=5,
        tiktok=    [106043, 13821878, 7679, 59047, 1145122],
        instagram= [40872,   5857905, 1604, 28629,  209504],
        youtube=   [4410,    1005269,  310,   853,   31012],
    )
