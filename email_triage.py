"""
Email Triage Script - Grapes & Barrels
"""

import os
import requests
import anthropic

TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
USER_EMAIL    = os.environ.get("USER_EMAIL", "floris@grapesandbarrels.nl")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

FOLDERS = {
    "Producenten": "Producenten",
    "Klanten & Bestellingen": "Klanten & Bestellingen",
    "Logistiek": "Logistiek",
    "Finance": "Finance",
    "Nieuwsbrieven": "Nieuwsbrieven",
    "Software & Tools": "Software & Tools",
    "Postvak IN": "Postvak IN",
}

TRIAGE_PROMPT = """Je bent email-assistent voor Grapes & Barrels, een Nederlandse wijnimporteur.
Classificeer de email in PRECIES EEN van deze mappen:
- Producenten -> emails van wijnproducenten, bodegas, leveranciers
- Klanten & Bestellingen -> orders, klantenvragen, webshop-orders
- Logistiek -> verzending, douane, transport
- Finance -> facturen, betalingen, Mollie, banken
- Nieuwsbrieven -> nieuwsbrieven, marketing
- Software & Tools -> Google, Shopify, software
- Postvak IN -> al het andere

Geef ALLEEN de mapnaam terug.

Van: {sender}
Onderwerp: {subject}
Preview: {preview}"""


def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def get_folder_ids(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders?$top=50"
    folder_map = {}
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        data = r.json()
        for f in data.get("value", []):
            folder_map[f["displayName"]] = f["id"]
        url = data.get("@odata.nextLink")
    return folder_map


def get_inbox_emails(token, folder_id, max_emails=30):
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/{folder_id}/messages"
        f"?$filter=isRead eq false&$select=id,subject,from,bodyPreview"
        f"&$top={max_emails}&$orderby=receivedDateTime desc"
    )
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def classify_email(client, sender, subject, preview):
    prompt = TRIAGE_PROMPT.format(sender=sender, subject=subject, preview=preview[:400])
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=20,
        messages=[{"role": "user", "content": prompt}],
    )
    folder = msg.content[0].text.strip()
    if folder not in FOLDERS:
        return "Postvak IN"
    return folder


def move_email(token, message_id, dest_folder_id):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/move"
    r = requests.post(url, headers=headers, json={"destinationId": dest_folder_id}, timeout=30)
    r.raise_for_status()


def main():
    print("Grapes & Barrels - Email Triage gestart")
    token = get_access_token()
    folder_map = get_folder_ids(token)
    print(f"{len(folder_map)} mappen gevonden")
    inbox_id = folder_map.get("Postvak IN")
    if not inbox_id:
        print("Postvak IN niet gevonden")
        return
    emails = get_inbox_emails(token, inbox_id)
    print(f"{len(emails)} ongelezen emails")
    if not emails:
        return
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    moved = skipped = 0
    for email in emails:
        subject = email.get("subject", "")
        sender  = email.get("from", {}).get("emailAddress", {}).get("address", "")
        preview = email.get("bodyPreview", "")
        target  = classify_email(client, sender, subject, preview)
        if target == "Postvak IN":
            skipped += 1
            continue
        dest_id = folder_map.get(target)
        if not dest_id:
            skipped += 1
            continue
        try:
            move_email(token, email["id"], dest_id)
            print(f"  -> {target}: {subject[:60]}")
            moved += 1
        except Exception as e:
            print(f"  Fout: {e}")
            skipped += 1
    print(f"Klaar: {moved} verplaatst, {skipped} overgeslagen")


if __name__ == "__main__":
    main()
