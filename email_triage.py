"""
Email Triage Script - Grapes & Barrels
Leest ongelezen mails, verplaatst naar juiste map,
en maakt conceptreplies aan voor mails die een reactie vereisen.
"""

import os
import re
import json
import requests
import anthropic
from html.parser import HTMLParser

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
USER_EMAIL = os.environ.get("USER_EMAIL", "floris@grapesandbarrels.nl")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

FOLDERS = {
    "Producenten": "Producenten",
    "Klanten & Bestellingen": "Klanten & Bestellingen",
    "Logistiek": "Logistiek",
    "Finance": "Finance",
    "Ruis": "Ruis",
    "Postvak IN": "Postvak IN",
}

# Snelle filter op afzenderdomein â geen Claude-aanroep nodig
RUIS_DOMAINS = [
    "facebookmail.com", "facebook.com",
    "shopify.com", "mail.shopify.com",
    "syncwith.com",
    "google.com", "merchants.google.com", "googlemerchant",
    "mailchimp.com", "mc.us", "list-manage.com",
    "intuit.com", "intuitemailservice.com",
    "linkedin.com", "twitter.com", "instagram.com",
    "klaviyo.com", "sendgrid.net", "mailgun.org",
    "constantcontact.com", "campaignmonitor.com",
    "hubspot.com", "salesforce.com", "marketo.com",
    "noreply", "no-reply", "donotreply", "do-not-reply",
    "notifications@", "newsletter@", "marketing@",
    "mailer@", "info@shopify", "postmaster",
]

RUIS_SUBJECTS = [
    "scheduled report", "your weekly", "your monthly", "your daily",
    "unsubscribe", "newsletter", "nieuwsbrief",
    "don't miss", "limited time", "special offer",
    "korting", "aanbieding", "sale", "% off",
    "new features", "product update", "release notes",
    "you have", "meldingen", "notifications",
    "marketing", "advertentie",
]


def quick_classify(sender: str, subject: str) -> str | None:
    """Classificeer direct op basis van afzender/onderwerp, zonder Claude.
    Returns folder name or None als Claude nodig is.
    """
    s = sender.lower()
    sub = subject.lower()

    for pattern in RUIS_DOMAINS:
        if pattern in s:
            return "Ruis"

    for kw in RUIS_SUBJECTS:
        if kw in sub:
            return "Ruis"

    return None


TRIAGE_PROMPT = """Je bent email-assistent voor Floris van Grapes & Barrels, een Nederlandse wijnimporteur.

Analyseer de email hieronder en geef een JSON-antwoord.

Mappen:
- Producenten: emails van wijnproducenten, bodegas, wijnmakers, leveranciers
- Klanten & Bestellingen: orders, klantenvragen, webshop-orders, leveringsverzoeken
- Logistiek: verzending, douane, transport, DHL, PostNL
- Finance: facturen, betalingen, Mollie, banken, creditnota's
- Ruis: ALLES wat niet direct actie vereist â nieuwsbrieven, marketing, aanbiedingen,
  automatische meldingen, social media, software-updates, rapporten, ontvangstbevestigingen,
  notificaties van tools, Google/Facebook/Shopify meldingen, bulk email
- Postvak IN: alleen echte persoonlijke berichten die nergens anders passen

Bij twijfel: liever Ruis dan Postvak IN.

Een conceptreply is ALLEEN nodig bij:
- Echte vragen of verzoeken van klanten of producenten die antwoord verwachten
- Bestellingen die een bevestiging nodig hebben
- Logistieke problemen die actie vereisen
- Financiele verzoeken of disputen

GEEN reply bij: alles in Ruis, automatische meldingen, marketing.

Schrijf de conceptreply in dezelfde taal als de afzender (Nederlands als afzender Nederlands schrijft, etc.).
De reply moet professioneel, vriendelijk en beknopt zijn. Begin NIET met "Beste [naam]" als je de naam niet weet - gebruik dan "Beste," of "Goedemiddag,".
Onderteken altijd met: "Met vriendelijke groet,\nFloris\nGrapes & Barrels"

Geef ALLEEN valide JSON terug (geen markdown, geen uitleg):
{{
  "folder": "<een van de mapnamen hierboven>",
  "needs_reply": true of false,
  "draft_reply": "<conceptreply als tekst, of null>"
}}

Van: {sender}
Onderwerp: {subject}
Inhoud:
{body}"""


class HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.result = []

    def handle_data(self, d):
        self.result.append(d)

    def get_text(self):
        return " ".join(self.result)


def strip_html(html):
    s = HTMLStripper()
    try:
        s.feed(html)
    except Exception:
        pass
    text = s.get_text()
    text = re.sub(r'\s+', ' ', text).strip()
    return text


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


def ensure_folder_exists(token, folder_map, folder_name):
    """Maakt een map aan als die nog niet bestaat."""
    if folder_name in folder_map:
        return folder_map[folder_name]
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders"
    r = requests.post(url, headers=headers, json={"displayName": folder_name}, timeout=30)
    if r.status_code == 409:
        # Al bestaat - haal ID op
        return folder_map.get(folder_name)
    r.raise_for_status()
    new_id = r.json()["id"]
    folder_map[folder_name] = new_id
    print(f"Map aangemaakt: {folder_name}")
    return new_id


def get_inbox_emails(token, folder_id, max_emails=30):
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/{folder_id}/messages"
        f"?$filter=isRead eq false"
        f"&$select=id,subject,from,bodyPreview,body"
        f"&$top={max_emails}&$orderby=receivedDateTime desc"
    )
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def analyze_email(client, sender, subject, body_text):
    """Classificeer email en bepaal of een conceptreply nodig is.
    Returns: (folder, needs_reply, draft_reply)
    """
    prompt = TRIAGE_PROMPT.format(
        sender=sender,
        subject=subject,
        body=body_text[:2000],
    )
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()

    # Verwijder eventuele markdown code-blocks
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        data = json.loads(raw)
        folder = data.get("folder", "Postvak IN")
        if folder not in FOLDERS:
            folder = "Postvak IN"
        needs_reply = bool(data.get("needs_reply", False))
        draft_reply = data.get("draft_reply") or None
        if draft_reply == "null":
            draft_reply = None
        return folder, needs_reply, draft_reply
    except json.JSONDecodeError:
        print(f"   JSON parse fout. Raw: {raw[:100]}")
        folder = raw if raw in FOLDERS else "Postvak IN"
        return folder, False, None


def move_email(token, message_id, dest_folder_id):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/move"
    r = requests.post(url, headers=headers, json={"destinationId": dest_folder_id}, timeout=30)
    r.raise_for_status()
    return r.json().get("id", message_id)


def mark_as_read(token, message_id):
    """Markeer email als gelezen (voor Ruis-mails)."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}"
    requests.patch(url, headers=headers, json={"isRead": True}, timeout=30)


def create_reply_draft(token, message_id, draft_body):
    """Maakt een conceptreply aan in Outlook Concepten."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Stap 1: Maak lege reply-draft (kopieert To/Subject automatisch)
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/createReply"
    r = requests.post(url, headers=headers, json={}, timeout=30)
    r.raise_for_status()
    draft_id = r.json()["id"]

    # Stap 2: Vul de body in met Claude's tekst
    body_html = draft_body.replace("\n", "<br>\n")
    patch_url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{draft_id}"
    r2 = requests.patch(
        patch_url,
        headers=headers,
        json={
            "body": {
                "contentType": "HTML",
                "content": f"<p style='font-family:Calibri,sans-serif;font-size:11pt'>{body_html}</p>",
            }
        },
        timeout=30,
    )
    r2.raise_for_status()
    return draft_id


def main():
    print("Grapes & Barrels - Email Triage gestart")
    token = get_access_token()
    folder_map = get_folder_ids(token)
    print(f"{len(folder_map)} mappen gevonden")

    # Zorg dat alle benodigde mappen bestaan
    for folder_name in FOLDERS:
        if folder_name not in ("Postvak IN",):
            ensure_folder_exists(token, folder_map, folder_name)

    inbox_id = folder_map.get("Postvak IN")
    if not inbox_id:
        print("Postvak IN niet gevonden")
        return

    emails = get_inbox_emails(token, inbox_id)
    print(f"{len(emails)} ongelezen emails")
    if not emails:
        return

    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    moved = skipped = drafted = quick = 0

    for email in emails:
        subject = email.get("subject", "(geen onderwerp)")
        sender = email.get("from", {}).get("emailAddress", {}).get("address", "")
        body_obj = email.get("body", {})
        body_content = body_obj.get("content", "")
        content_type = body_obj.get("contentType", "text")

        if content_type.lower() == "html":
            body_text = strip_html(body_content)
        else:
            body_text = body_content

        if not body_text.strip():
            body_text = email.get("bodyPreview", "")

        # Probeer snelle classificatie zonder Claude
        quick_folder = quick_classify(sender, subject)
        if quick_folder:
            dest_id = folder_map.get(quick_folder)
            if dest_id:
                try:
                    new_id = move_email(token, email["id"], dest_id)
                    print(f" [snel] -> {quick_folder}: {subject[:60]}")
                    quick += 1
                    moved += 1
                except Exception as e:
                    print(f" Fout snel filter '{subject[:40]}': {e}")
                    skipped += 1
            continue  # Geen Claude nodig

        # Claude voor onduidelijke gevallen
        target, needs_reply, draft_reply = analyze_email(client, sender, subject, body_text)

        new_message_id = email["id"]
        if target != "Postvak IN":
            dest_id = folder_map.get(target)
            if dest_id:
                try:
                    new_message_id = move_email(token, email["id"], dest_id)
                    print(f" -> {target}: {subject[:60]}")
                    moved += 1
                except Exception as e:
                    print(f" Fout verplaatsen '{subject[:40]}': {e}")
                    skipped += 1
            else:
                print(f" Map '{target}' niet gevonden, overgeslagen")
                skipped += 1
        else:
            skipped += 1

        if needs_reply and draft_reply:
            try:
                create_reply_draft(token, new_message_id, draft_reply)
                print(f"   [concept reply] {subject[:50]}")
                drafted += 1
            except Exception as e:
                print(f"   Fout concept reply '{subject[:40]}': {e}")

    print(f"\nKlaar: {moved} verplaatst ({quick} snel), {drafted} concept replies, {skipped} overgeslagen")


if __name__ == "__main__":
    main()
