"""
Email Triage Script - Grapes & Barrels
Leest mails (gelezen en ongelezen), verplaatst naar juiste map.
Maakt GEEN automatische conceptreplies meer - flagt alleen welke mails
een antwoord nodig lijken te hebben.

Wordt geleidelijk beter: elke run wordt bijgehouden in triage_history.json.
Als Floris een mail zelf naar een andere map verplaatst dan de bot koos,
wordt dat als correctie opgeslagen en gebruikt om (a) toekomstige mails van
diezelfde afzender direct goed te zetten (zonder Claude-aanroep) en (b) de
Claude-prompt te verrijken met een paar recente correctievoorbeelden.
"""

import os
import re
import json
import requests
import anthropic
from datetime import datetime, timezone
from html.parser import HTMLParser

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
USER_EMAIL = os.environ.get("USER_EMAIL", "floris@grapesandbarrels.nl")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

HISTORY_FILE = "triage_history.json"
MAX_HISTORY = 800
CORRECTION_CHECK_DELAY_HOURS = 6   # geef Floris tijd om zelf te corrigeren voor we "afronden"
MAX_CORRECTION_CHECKS_PER_RUN = 40
LEARNED_MIN_SAMPLES = 3
LEARNED_MIN_RATIO = 0.75

FOLDERS = {
    "Producenten": "Producenten",
    "Klanten & Bestellingen": "Klanten & Bestellingen",
    "Logistiek": "Logistiek",
    "Finance": "Finance",
    "Marketing & Tools": "Marketing & Tools",
    "Evenementen & Netwerk": "Evenementen & Netwerk",
    "Juridisch": "Juridisch",
    "Ruis": "Ruis",
    "Postvak IN": "Postvak IN",
}

# ---------------------------------------------------------------------------
# Snelle filters -- geen Claude-aanroep nodig
# ---------------------------------------------------------------------------

# Afzenderdomein (substring match) -> vaste map. Dit zijn bekende automatische
# afzenders die business-relevant zijn maar geen persoonlijke actie vereisen.
QUICK_DOMAIN_MAP = {
    "shopify.com": "Marketing & Tools",
    "mail.shopify.com": "Marketing & Tools",
    "syncwith.com": "Marketing & Tools",
    "merchants.google.com": "Marketing & Tools",
    "googlemerchant": "Marketing & Tools",
    "mailchimp.com": "Marketing & Tools",
    "list-manage.com": "Marketing & Tools",
    "klaviyo.com": "Marketing & Tools",
    "sendgrid.net": "Marketing & Tools",
    "mailgun.org": "Marketing & Tools",
    "constantcontact.com": "Marketing & Tools",
    "campaignmonitor.com": "Marketing & Tools",
    "hubspot.com": "Marketing & Tools",
    "salesforce.com": "Marketing & Tools",
    "marketo.com": "Marketing & Tools",
    "facebookmail.com": "Marketing & Tools",
    "facebook.com": "Marketing & Tools",
    "twitter.com": "Marketing & Tools",
    "instagram.com": "Marketing & Tools",
    "linkedin.com": "Evenementen & Netwerk",
    "intuit.com": "Finance",
    "intuitemailservice.com": "Finance",
}

# Afzender local-part patronen die vrijwel altijd pure ruis zijn (mits domein
# hierboven niet al iets specifieks toekende)
RUIS_SENDER_PATTERNS = [
    "noreply", "no-reply", "donotreply", "do-not-reply",
    "notifications@", "postmaster", "mailer@",
]

MARKETING_SUBJECT_KEYWORDS = [
    "nieuwsbrief", "newsletter", "korting", "aanbieding", "% off",
    "special offer", "limited time", "don't miss",
    "new features", "product update", "release notes",
]

RUIS_SUBJECT_KEYWORDS = [
    "scheduled report", "your weekly", "your monthly", "your daily",
    "unsubscribe", "you have", "meldingen", "notifications",
]


def quick_classify(sender: str, subject: str) -> str | None:
    s = sender.lower()
    sub = subject.lower()

    for domain, folder in QUICK_DOMAIN_MAP.items():
        if domain in s:
            return folder
    for pattern in RUIS_SENDER_PATTERNS:
        if pattern in s:
            return "Ruis"
    for kw in MARKETING_SUBJECT_KEYWORDS:
        if kw in sub:
            return "Marketing & Tools"
    for kw in RUIS_SUBJECT_KEYWORDS:
        if kw in sub:
            return "Ruis"
    return None


# ---------------------------------------------------------------------------
# Leerdata (triage_history.json) -- maakt classificatie beter over tijd
# ---------------------------------------------------------------------------

def load_history() -> dict:
    if not os.path.exists(HISTORY_FILE):
        return {"records": []}
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            data.setdefault("records", [])
            return data
    except (json.JSONDecodeError, OSError):
        return {"records": []}


def save_history(history: dict) -> None:
    history["records"] = history["records"][-MAX_HISTORY:]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def domain_of(sender: str) -> str:
    return sender.split("@")[-1].lower() if "@" in sender else sender.lower()


def record_classification(history: dict, message_id: str, sender: str, subject: str,
                           folder: str, source: str) -> None:
    history["records"].append({
        "id": message_id,
        "sender_domain": domain_of(sender),
        "subject": subject[:150],
        "folder": folder,
        "source": source,  # "quick" | "learned" | "llm" | "correction"
        "ts": datetime.now(timezone.utc).isoformat(),
        "checked": source == "correction",
    })


def build_domain_stats(records: list) -> dict:
    """sender_domain -> {folder: gewogen score}. Correcties tellen zwaarder."""
    stats: dict = {}
    for r in records:
        d, f = r.get("sender_domain"), r.get("folder")
        if not d or not f:
            continue
        weight = 3 if r.get("source") == "correction" else 1
        stats.setdefault(d, {})
        stats[d][f] = stats[d].get(f, 0) + weight
    return stats


def learned_classify(domain_stats: dict, sender: str) -> str | None:
    counts = domain_stats.get(domain_of(sender))
    if not counts:
        return None
    total = sum(counts.values())
    best_folder, best_score = max(counts.items(), key=lambda kv: kv[1])
    if total >= LEARNED_MIN_SAMPLES and (best_score / total) >= LEARNED_MIN_RATIO:
        return best_folder
    return None


def build_learned_examples_text(records: list, limit: int = 8) -> str:
    corrections = [r for r in records if r.get("source") == "correction"][-limit:]
    if not corrections:
        return ""
    lines = [
        "Geleerde voorbeelden uit eerdere correcties door Floris "
        "(hij heeft deze zelf verplaatst -- dit is leidend voor vergelijkbare mails):"
    ]
    for c in corrections:
        lines.append(f"- afzenderdomein \"{c['sender_domain']}\", onderwerp \"{c['subject']}\" -> {c['folder']}")
    return "\n".join(lines)


def get_message_folder_id(token: str, message_id: str) -> str | None:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}?$select=parentFolderId"
    r = requests.get(url, headers=headers, timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json().get("parentFolderId")


def check_for_corrections(token: str, history: dict, folder_map: dict) -> int:
    """Kijkt of eerder verplaatste mails inmiddels ergens anders staan (= Floris
    heeft de bot gecorrigeerd) en slaat dat op als leervoorbeeld."""
    id_to_name = {v: k for k, v in folder_map.items()}
    now = datetime.now(timezone.utc)
    checks_done = 0
    corrections_found = 0
    new_records = []

    for r in history["records"]:
        if r.get("checked") or checks_done >= MAX_CORRECTION_CHECKS_PER_RUN:
            continue
        try:
            record_time = datetime.fromisoformat(r["ts"])
        except Exception:
            r["checked"] = True
            continue
        if (now - record_time).total_seconds() < CORRECTION_CHECK_DELAY_HOURS * 3600:
            continue

        checks_done += 1
        r["checked"] = True
        current_folder_id = get_message_folder_id(token, r["id"])
        if current_folder_id is None:
            continue
        current_folder_name = id_to_name.get(current_folder_id)
        if current_folder_name and current_folder_name != r["folder"]:
            corrections_found += 1
            new_records.append({
                "id": f"{r['id']}-correction-{int(now.timestamp())}",
                "sender_domain": r["sender_domain"],
                "subject": r["subject"],
                "folder": current_folder_name,
                "source": "correction",
                "ts": now.isoformat(),
                "checked": True,
            })

    history["records"].extend(new_records)
    if corrections_found:
        print(f"{corrections_found} correctie(s) gevonden -> toegevoegd aan leerdata")
    return corrections_found


# ---------------------------------------------------------------------------
# Claude classificatie
# ---------------------------------------------------------------------------

TRIAGE_PROMPT = """Je bent email-assistent voor Floris van Grapes & Barrels, een Nederlandse wijnimporteur.

Analyseer de email hieronder en geef een JSON-antwoord.

Mappen:
- Producenten: emails van wijnproducenten, bodegas, wijnmakers, leveranciers
- Klanten & Bestellingen: orders, klantenvragen, webshop-orders, leveringsverzoeken
- Logistiek: verzending, douane, transport, DHL, PostNL
- Finance: facturen, betalingen, Mollie, banken, creditnota's
- Marketing & Tools: nieuwsbrieven, aanbiedingen, en automatische meldingen van
  software/tools die Floris gebruikt (Shopify, Google, social media, mailinglijsten) --
  niet urgent, maar wel relevant genoeg om af en toe te scannen
- Evenementen & Netwerk: wijnbeurzen, proeverijen, uitnodigingen, netwerkcontacten,
  LinkedIn-berichten van mensen (geen automatische meldingen)
- Juridisch: contracten, voorwaarden, verzekeringen, juridische correspondentie
- Ruis: pure spam, phishing, volstrekt irrelevante automatische mail zonder enige
  business-waarde
- Postvak IN: alleen echte persoonlijke berichten die nergens anders bij passen

Bij twijfel tussen een specifieke map en Ruis: kies de specifieke map, niet Ruis.
Bij twijfel tussen Postvak IN en Ruis: kies Postvak IN.
Gebruik Ruis alleen als je vrij zeker bent dat het geen enkele waarde heeft.

{learned_examples}

Geef ALLEEN valide JSON terug (geen markdown, geen uitleg):
{{
  "folder": "<een van de mapnamen hierboven>",
  "needs_reply": true of false
}}

"needs_reply" is true als dit een echte vraag/verzoek/bestelling/dispuut is dat een
persoonlijk antwoord van Floris verwacht. Er wordt GEEN automatische conceptreply meer
aangemaakt -- dit veld is puur om te signaleren welke mails Floris zelf moet oppakken.

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
    if folder_name in folder_map:
        return folder_map[folder_name]
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders"
    r = requests.post(url, headers=headers, json={"displayName": folder_name}, timeout=30)
    if r.status_code == 409:
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
        f"?$select=id,subject,from,bodyPreview,body,isRead"
        f"&$top={max_emails}&$orderby=receivedDateTime desc"
    )
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def analyze_email(client, sender, subject, body_text, learned_examples_text):
    prompt = TRIAGE_PROMPT.format(
        sender=sender,
        subject=subject,
        body=body_text[:2000],
        learned_examples=learned_examples_text,
    )
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=200,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```\s*$', '', raw)
    try:
        data = json.loads(raw)
        folder = data.get("folder", "Postvak IN")
        if folder not in FOLDERS:
            folder = "Postvak IN"
        needs_reply = bool(data.get("needs_reply", False))
        return folder, needs_reply
    except json.JSONDecodeError:
        print(f"   JSON parse fout. Raw: {raw[:100]}")
        folder = raw if raw in FOLDERS else "Postvak IN"
        return folder, False


def move_email(token, message_id, dest_folder_id):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/move"
    r = requests.post(url, headers=headers, json={"destinationId": dest_folder_id}, timeout=30)
    r.raise_for_status()
    return r.json().get("id", message_id)


def main():
    print("Grapes & Barrels - Email Triage gestart")
    token = get_access_token()
    folder_map = get_folder_ids(token)
    print(f"{len(folder_map)} mappen gevonden")

    for folder_name in FOLDERS:
        if folder_name != "Postvak IN":
            ensure_folder_exists(token, folder_map, folder_name)

    history = load_history()
    check_for_corrections(token, history, folder_map)
    domain_stats = build_domain_stats(history["records"])
    learned_examples_text = build_learned_examples_text(history["records"])

    inbox_id = folder_map.get("Postvak IN")
    if not inbox_id:
        print("Postvak IN niet gevonden")
        save_history(history)
        return

    emails = get_inbox_emails(token, inbox_id)
    print(f"{len(emails)} emails in inbox")
    if not emails:
        save_history(history)
        return

    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    moved = skipped = quick = learned = needs_reply_count = 0
    needs_reply_list = []

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

        target = None
        source = None

        quick_folder = quick_classify(sender, subject)
        if quick_folder:
            target, source = quick_folder, "quick"
        else:
            learned_folder = learned_classify(domain_stats, sender)
            if learned_folder:
                target, source = learned_folder, "learned"

        needs_reply = False
        if target is None:
            target, needs_reply = analyze_email(client, sender, subject, body_text, learned_examples_text)
            source = "llm"

        if source == "quick":
            quick += 1
        elif source == "learned":
            learned += 1

        if needs_reply:
            needs_reply_count += 1
            needs_reply_list.append(f"{subject[:60]} (van {sender})")

        new_message_id = email["id"]
        if target != "Postvak IN":
            dest_id = folder_map.get(target)
            if dest_id:
                try:
                    new_message_id = move_email(token, email["id"], dest_id)
                    print(f" [{source}] -> {target}: {subject[:60]}")
                    moved += 1
                except Exception as e:
                    print(f" Fout verplaatsen '{subject[:40]}': {e}")
                    skipped += 1
                    continue
            else:
                print(f" Map '{target}' niet gevonden, overgeslagen")
                skipped += 1
                continue
        else:
            skipped += 1

        record_classification(history, new_message_id, sender, subject, target, source)

    save_history(history)

    print(f"\nKlaar: {moved} verplaatst ({quick} snel, {learned} geleerd), "
          f"{skipped} in Postvak IN gelaten")
    if needs_reply_list:
        print(f"\n{needs_reply_count} mail(s) lijken een antwoord nodig te hebben (geen concept aangemaakt):")
        for item in needs_reply_list:
            print(f"  - {item}")


if __name__ == "__main__":
    main()
