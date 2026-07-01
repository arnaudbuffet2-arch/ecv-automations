"""
Crée un brouillon Gmail récapitulatif des stats community management.
Expéditeur : emilecoachvocal@gmail.com
CC : arnaud.buffet2@gmail.com, benedictemoyat.rp@gmail.com

Usage : python scripts/draft_community_email.py
"""
import sys
import base64
from email.mime.text import MIMEText
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "claude safe"))
from google_auth import get_google_credentials
from googleapiclient.discovery import build


def pct(new, old):
    if not old:
        return "—"
    p = (new - old) / old * 100
    sign = "+" if p >= 0 else ""
    return f"{sign}{p:.1f}%"


def create_draft(subject: str, body: str):
    creds = get_google_credentials()
    service = build("gmail", "v1", credentials=creds)

    msg = MIMEText(body, "plain", "utf-8")
    msg["From"] = "emilecoachvocal@gmail.com"
    msg["To"] = "emilecoachvocal@gmail.com"
    msg["Cc"] = "arnaud.buffet2@gmail.com, benedictemoyat.rp@gmail.com"
    msg["Subject"] = subject

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}}
    ).execute()
    print(f"Brouillon créé : {draft['id']}")


def build_body(period: str, stats: dict, prev_stats: dict) -> str:
    tt  = stats["tiktok"]
    ig  = stats["instagram"]
    yt  = stats["youtube"]
    ptt = prev_stats["tiktok"]
    pig = prev_stats["instagram"]
    pyt = prev_stats["youtube"]

    cons_vues   = tt["vues"]   + ig["vues"]   + yt["vues"]
    cons_likes  = tt["likes"]  + ig["likes"]  + yt["likes"]
    cons_com    = tt["com"]    + ig["com"]    + yt["com"]
    cons_shares = tt["shares"] + ig["shares"] + yt["shares"]
    cons_abo    = tt["abo"]    + ig["abo"]    + yt["abo"]

    pcons_vues   = ptt["vues"]   + pig["vues"]   + pyt["vues"]
    pcons_likes  = ptt["likes"]  + pig["likes"]  + pyt["likes"]
    pcons_com    = ptt["com"]    + pig["com"]    + pyt["com"]
    pcons_shares = ptt["shares"] + pig["shares"] + pyt["shares"]
    pcons_abo    = ptt["abo"]    + pig["abo"]    + pyt["abo"]

    def fmt(n): return f"{n:,}".replace(",", " ")

    lines = [
        "Bonjour,",
        "",
        f"Voici le récapitulatif des statistiques community management — {period} :",
        f"(Les pourcentages d'évolution sont calculés par rapport au mois précédent.)",
        "",
        "── TikTok ──────────────────────────────────",
        f"  Abonnés     : {fmt(tt['abo'])}  ({pct(tt['abo'], ptt['abo'])})",
        f"  Vues        : {fmt(tt['vues'])}  ({pct(tt['vues'], ptt['vues'])})",
        f"  J'aime      : {fmt(tt['likes'])}  ({pct(tt['likes'], ptt['likes'])})",
        f"  Commentaires: {fmt(tt['com'])}  ({pct(tt['com'], ptt['com'])})",
        f"  Partages    : {fmt(tt['shares'])}  ({pct(tt['shares'], ptt['shares'])})",
        "",
        "── Instagram ───────────────────────────────",
        f"  Abonnés     : {fmt(ig['abo'])}  ({pct(ig['abo'], pig['abo'])})",
        f"  Vues        : {fmt(ig['vues'])}  ({pct(ig['vues'], pig['vues'])})",
        f"  J'aime      : {fmt(ig['likes'])}  ({pct(ig['likes'], pig['likes'])})",
        f"  Commentaires: {fmt(ig['com'])}  ({pct(ig['com'], pig['com'])})",
        f"  Partages    : {fmt(ig['shares'])}  ({pct(ig['shares'], pig['shares'])})",
        "",
        "── YouTube ─────────────────────────────────",
        f"  Abonnés     : {fmt(yt['abo'])}  ({pct(yt['abo'], pyt['abo'])})",
        f"  Vues        : {fmt(yt['vues'])}  ({pct(yt['vues'], pyt['vues'])})",
        f"  J'aime      : {fmt(yt['likes'])}  ({pct(yt['likes'], pyt['likes'])})",
        f"  Commentaires: {fmt(yt['com'])}  ({pct(yt['com'], pyt['com'])})",
        f"  Partages    : {fmt(yt['shares'])}  ({pct(yt['shares'], pyt['shares'])})",
        "",
        "── Consolidé ───────────────────────────────",
        f"  Abonnés     : {fmt(cons_abo)}  ({pct(cons_abo, pcons_abo)})",
        f"  Vues        : {fmt(cons_vues)}  ({pct(cons_vues, pcons_vues)})",
        f"  J'aime      : {fmt(cons_likes)}  ({pct(cons_likes, pcons_likes)})",
        f"  Commentaires: {fmt(cons_com)}  ({pct(cons_com, pcons_com)})",
        f"  Partages    : {fmt(cons_shares)}  ({pct(cons_shares, pcons_shares)})",
        "",
        "Cordialement,",
        "de la part de Arno CM Management - Arnaud Buffet",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    # Valeurs cumulatives depuis janvier (mois complet)
    mai = {
        "tiktok":    {"abo": 106043, "vues": 13821878, "likes": 1145122, "com": 7679,  "shares": 59047},
        "instagram": {"abo": 40872,  "vues": 5857905,  "likes": 209504,  "com": 1604,  "shares": 28629},
        "youtube":   {"abo": 4410,   "vues": 1005269,  "likes": 31012,   "com": 310,   "shares": 853},
    }
    avril = {
        "tiktok":    {"abo": 97899,  "vues": 12691446, "likes": 1057769, "com": 7223, "shares": 55656},
        "instagram": {"abo": 40212,  "vues": 5284161,  "likes": 181594,  "com": 1434, "shares": 24770},
        "youtube":   {"abo": 4380,   "vues": 960256,   "likes": 28344,   "com": 282,  "shares": 522},
    }

    body = build_body("Mai 2026", mai, avril)
    create_draft("Stats Community Management — Mai 2026", body)
