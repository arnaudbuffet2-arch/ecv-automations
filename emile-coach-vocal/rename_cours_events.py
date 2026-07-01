"""
rename_cours_events
-------------------
Renomme les événements Google Calendar dont le titre contient ce) ou ces)
en remplaçant cette fin par ce!).

Usage :
    python rename_cours_events.py          # affiche le nombre d'événements concernés
    python rename_cours_events.py --apply  # demande confirmation puis renomme
"""

import argparse
import re
from googleapiclient.discovery import build
from google_auth import get_google_credentials

PATTERN = re.compile(r"ces?\s?\)", re.IGNORECASE)


def find_events_to_rename(service):
    to_rename = []
    page_token = None

    while True:
        kwargs = dict(
            calendarId="primary",
            timeMin="2010-01-01T00:00:00Z",
            timeMax="2030-12-31T23:59:59Z",
            singleEvents=True,
            maxResults=2500,
        )
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.events().list(**kwargs).execute()

        for event in result.get("items", []):
            title = event.get("summary", "")
            if PATTERN.search(title):
                new_title = PATTERN.sub("ce!)", title)
                to_rename.append((event["id"], new_title))

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return to_rename


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="Appliquer les renommages")
    args = parser.parse_args()

    creds = get_google_credentials()
    service = build("calendar", "v3", credentials=creds)

    to_rename = find_events_to_rename(service)

    if not to_rename:
        print("Aucun événement à renommer.")
        return

    print(f"{len(to_rename)} evenement(s) a renommer (ce) ou ces) -> ce!).")

    if not args.apply:
        print("Lancez avec --apply pour appliquer.")
        return

    confirm = input("Confirmer le renommage ? (o/n) : ")
    if confirm.strip().lower() != "o":
        print("Annulé.")
        return

    for event_id, new_title in to_rename:
        service.events().patch(
            calendarId="primary",
            eventId=event_id,
            body={"summary": new_title},
        ).execute()

    print(f"{len(to_rename)} événement(s) renommé(s).")


if __name__ == "__main__":
    main()
