"""
Récupère les stats mensuelles Instagram, TikTok et YouTube.
Stocke les données brutes en JSON dans data/social_stats/.

Config requise : scripts/social_config.json (voir social_config.example.json)
"""

import json
import os
import sys
import requests
from datetime import datetime, date
from pathlib import Path

# ── Chemins ──────────────────────────────────────────────────────────────────

VAULT_ROOT = Path(__file__).parent.parent
DATA_DIR   = VAULT_ROOT / "data" / "social_stats"
CONFIG_FILE = Path(__file__).parent / "social_config.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)


# ── Config ───────────────────────────────────────────────────────────────────

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        print(f"[ERREUR] Fichier de config manquant : {CONFIG_FILE}")
        print("  → Copie social_config.example.json en social_config.json et remplis tes clés.")
        sys.exit(1)
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


# ── Instagram ─────────────────────────────────────────────────────────────────

def fetch_instagram(cfg: dict) -> dict:
    """Récupère les stats du mois via Meta Graph API."""
    token    = cfg["instagram"]["access_token"]
    ig_id    = cfg["instagram"]["ig_user_id"]

    now      = date.today()
    since    = date(now.year, now.month, 1).isoformat()
    until    = now.isoformat()

    # Liste des médias du mois
    media_url = (
        f"https://graph.facebook.com/v19.0/{ig_id}/media"
        f"?fields=id,timestamp,like_count,comments_count,saved"
        f"&access_token={token}"
    )
    media_resp = requests.get(media_url, timeout=15)
    if not media_resp.ok:
        raise RuntimeError(f"Meta API {media_resp.status_code}: {media_resp.json()}")
    media_data = media_resp.json().get("data", [])

    # Filtrer sur le mois en cours
    monthly = [
        m for m in media_data
        if m.get("timestamp", "")[:7] == now.strftime("%Y-%m")
    ]

    totals = {
        "likes":    sum(m.get("like_count", 0) for m in monthly),
        "comments": sum(m.get("comments_count", 0) for m in monthly),
        "saves":    0,
        "shares":   0,
        "posts":    len(monthly),
    }

    for m in monthly:
        ins_url = (
            f"https://graph.facebook.com/v19.0/{m['id']}/insights"
            f"?metric=shares,saved&access_token={token}"
        )
        try:
            ins = requests.get(ins_url, timeout=10).json()
            for d in ins.get("data", []):
                if d["name"] == "shares":
                    totals["shares"] += d["values"][-1]["value"]
                elif d["name"] == "saved":
                    totals["saves"] += d["values"][-1]["value"]
        except Exception:
            pass

    return {"source": "instagram", "period": f"{now.year}-{now.month:02d}", "totals": totals, "posts": monthly}


# ── TikTok ────────────────────────────────────────────────────────────────────

def fetch_tiktok(cfg: dict) -> dict:
    """Récupère les stats du mois via Windsor.ai (tiktok_organic)."""
    tk   = cfg["tiktok"]
    now  = date.today()
    start = date(now.year, now.month, 1).isoformat()
    end   = now.isoformat()

    params = {
        "api_key":    cfg["windsor_api_key"],
        "date_from":  start,
        "date_to":    end,
        "account_id": tk["windsor_account_id"],
        "fields":     "date,likes,comments,shares,unique_video_views",
    }
    resp = requests.get(
        f"https://connectors.windsor.ai/{tk['windsor_connector']}",
        params=params,
        timeout=15,
    )
    resp.raise_for_status()
    rows = resp.json().get("data", [])

    totals = {
        "likes":    sum(r.get("likes", 0) or 0 for r in rows),
        "comments": sum(r.get("comments", 0) or 0 for r in rows),
        "shares":   sum(r.get("shares", 0) or 0 for r in rows),
        "views":    sum(r.get("unique_video_views", 0) or 0 for r in rows),
    }
    return {"source": "tiktok", "period": f"{now.year}-{now.month:02d}", "totals": totals}


# ── YouTube ───────────────────────────────────────────────────────────────────

def fetch_youtube(cfg: dict) -> dict:
    """Récupère les stats du mois via Windsor.ai (youtube)."""
    yt   = cfg["youtube"]
    now  = date.today()
    start = date(now.year, now.month, 1).isoformat()
    end   = now.isoformat()

    params = {
        "api_key":    cfg["windsor_api_key"],
        "date_from":  start,
        "date_to":    end,
        "account_id": yt["windsor_account_id"],
        "fields":     "date,likes,comments,shares,views",
    }
    resp = requests.get(
        f"https://connectors.windsor.ai/{yt['windsor_connector']}",
        params=params,
        timeout=15,
    )
    resp.raise_for_status()
    rows = resp.json().get("data", [])

    totals = {
        "likes":    sum(r.get("likes", 0) or 0 for r in rows),
        "comments": sum(r.get("comments", 0) or 0 for r in rows),
        "shares":   sum(r.get("shares", 0) or 0 for r in rows),
        "views":    sum(r.get("views", 0) or 0 for r in rows),
    }
    return {"source": "youtube", "period": f"{now.year}-{now.month:02d}", "totals": totals}


# ── Sauvegarde ────────────────────────────────────────────────────────────────

def save(data: dict):
    now      = date.today()
    filename = DATA_DIR / f"{now.strftime('%Y-%m')}_{data['source']}.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"  ✓ {data['source']:12s} → {filename.name}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    cfg = load_config()
    now = date.today()
    print(f"\nRécupération des stats — {now.strftime('%B %Y')}\n")

    fetchers = {
        "instagram": fetch_instagram,
        "tiktok":    fetch_tiktok,
        "youtube":   fetch_youtube,
    }

    for name, fn in fetchers.items():
        if cfg.get(name, {}).get("enabled", False):
            try:
                data = fn(cfg)
                save(data)
                t = data["totals"]
                print(f"    likes={t.get('likes',0)}  comments={t.get('comments',0)}  saves={t.get('saves',0)}  shares={t.get('shares',0)}")
            except Exception as e:
                print(f"  [ERREUR] {name}: {e}")
        else:
            print(f"  – {name:12s} désactivé (enabled: false dans la config)")

    print("\nDone. Fichiers dans data/social_stats/\n")


if __name__ == "__main__":
    main()
