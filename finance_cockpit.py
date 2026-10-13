"""
Finance Cockpit (Rechnungs-Cockpit) – läuft als GitHub Action am 1. und 15. jedes Monats
------------------------------------------------------------------------------------------
Beantwortet vier Fragen und schickt die Antwort per Mail + Google Sheet:

  1. Welche Rechnungen wurden im Monat bezahlt?              (sevDesk: payDate / Status)
  2. Welche stehen aus?                                      (sevDesk: Status offen)
  3. Welche sollten gemahnt werden, in welcher Stufe?        (Regeln im Code, nicht in der KI)
  4. Welche müssen wir noch stellen?                         (Tab "Vertragsstamm" gegen sevDesk)

Die Zahlen berechnet der Code. Die KI (Claude oder Gemini) schreibt nur die Zusammenfassung
über die fertig berechneten Tabellen. Sie versendet nichts: Mahnungen und Zahlungen gibt ein
Mensch frei.

Aufruf:
    python finance_cockpit.py                 normaler Lauf (sevDesk -> Sheet -> Mail)
    python finance_cockpit.py --demo          Beispieldaten statt sevDesk; schreibt nur in DEMO_-Tabs
    python finance_cockpit.py --check         prüft nur die sevDesk-Verbindung und zeigt die Felder
    python finance_cockpit.py --dry-run       rechnet und druckt, schreibt nichts, mailt nichts

Umgebungsvariablen (GitHub Secrets):
    GOOGLE_CREDENTIALS_B64   wie beim Ausschreibungstracker (Pflicht außer bei --dry-run/--check)
    FINANCE_SHEET_KEY        ID des Google Sheets "Finance Cockpit" (Pflicht außer bei --dry-run)
    SEVDESK_API_TOKEN        sevDesk-API-Token (Pflicht außer bei --demo)
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM   wie beim Tracker (für die Mail)
    ALERT_EMAIL              Empfänger, mehrere mit Komma; FINANCE_EMAIL überschreibt das
    ANTHROPIC_API_KEY        für die KI-Zusammenfassung mit Claude (bevorzugt)
    ANTHROPIC_MODEL          optional, Standard claude-sonnet-5-5
    GEMINI_API_KEY           Alternative, wenn kein Anthropic-Key da ist (nur bezahlter Zugang!)
    FINANCE_TODAY            optional, YYYY-MM-DD zum Testen
"""

import base64
import json
import os
import re
import smtplib
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo

# --------------------------------------------------------------------------
# Konfiguration
# --------------------------------------------------------------------------

TIMEZONE = ZoneInfo("Europe/Berlin")
SEVDESK_BASE = "https://my.sevdesk.de/api/v1"

TAB_INVOICES = "Rechnungen_Ist"
TAB_CONTRACTS = "Vertragsstamm"
TAB_EXPECTED = "Soll_vs_Ist"
TAB_SUMMARY = "Zusammenfassung"

# Zahlungsziel, falls eine Rechnung keines hat
DEFAULT_PAYMENT_DAYS = 14

# Mahnvorschlag nach Tagen NACH Fälligkeit (absteigend sortiert, die höchste passende Stufe gilt)
DUNNING_STAGES = [
    (40, "Eskalation: Geschäftsführung entscheidet"),
    (28, "Letzte Mahnung + Anruf"),
    (14, "1. Mahnung"),
    (5, "Zahlungserinnerung"),
]

# sevDesk-Status laut offizieller API-Beschreibung (Stand 2024): 100 Entwurf, 200 offen,
# 750 teilbezahlt, 1000 bezahlt, 50 deaktivierte Wiederkehr-Vorlage
STATUS_DRAFT, STATUS_OPEN, STATUS_PARTIAL, STATUS_PAID = 100, 200, 750, 1000
# Rechnungstypen, die als echte Rechnungen zählen: Rechnung, Schlussrechnung, Teil-, Abschlagsrechnung
COUNTED_TYPES = {"RE", "ER", "TR", "AR"}

INTERVAL_MONTHS = {
    "monatlich": 1, "quartalsweise": 3, "vierteljährlich": 3, "halbjährlich": 6,
    "jährlich": 12, "jaehrlich": 12,
}

CONTRACT_HEADER = [
    "Kunde (wie in sevDesk)", "Produkt", "Betrag netto", "Rhythmus",
    "Erste Rechnung", "Ende", "Rechnungstag", "Aktiv", "Hinweis",
]
INVOICE_HEADER = [
    "Rechnungsnr.", "Kunde", "Rechnungsdatum", "Fällig am", "Brutto", "Netto", "Bezahlt",
    "Offen", "Status", "Tage überfällig", "Mahnvorschlag", "Bezahlt am", "sevDesk-ID",
]
EXPECTED_HEADER = [
    "Monat", "Kunde", "Erwartet netto", "Status", "Rechnungen", "Ausgestellt netto", "Hinweis",
]

MONTHS_DE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August",
             "September", "Oktober", "November", "Dezember"]

AI_SYSTEM = (
    "Du schreibst für die Geschäftsführung der Blindleister GmbH (Berlin, B2B-Software) eine kurze, "
    "sachliche Zusammenfassung zum Stand der Kundenrechnungen. Du bekommst fertig berechnete Daten "
    "als JSON. Regeln: Nenne nur Zahlen und Namen, die in den Daten stehen; rechne nichts neu; "
    "erfinde nichts; wenn etwas fehlt, sag das. Schreibe auf Deutsch, höchstens 220 Wörter, in "
    "diesen vier Abschnitten mit je einer Überschrift in einer Zeile: 'Auf einen Blick', "
    "'Überfällig und Mahnvorschläge', 'Noch zu stellen', 'Drei Prioritäten'. Mahnungen sind nur "
    "Vorschläge; ein Mensch gibt sie frei."
)


# --------------------------------------------------------------------------
# Hilfsfunktionen
# --------------------------------------------------------------------------

def today_berlin():
    override = os.environ.get("FINANCE_TODAY")
    if override:
        return datetime.strptime(override, "%Y-%m-%d").date()
    return datetime.now(TIMEZONE).date()


def parse_date(value):
    """Akzeptiert ISO ('2026-10-01', '2026-10-01T00:00:00+02:00'), '01.10.2026', '1.10.26'."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_amount(value):
    """'1.234,50', '1234.5', 1234.5, '1.234,50 €' -> float"""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value or "").replace("€", "").replace("\u00a0", "").replace(" ", "").strip()
    if not text:
        return 0.0
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def eur(amount):
    """1234.5 -> '1.234,50 €'"""
    text = f"{abs(amount):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if amount < 0 else ''}{text} €"


def days_text(n):
    return "1 Tag" if n == 1 else f"{n} Tage"


def month_start(d):
    return d.replace(day=1)


def add_months(d, n):
    index = d.year * 12 + (d.month - 1) + n
    return date(index // 12, index % 12 + 1, 1)


def month_label(d):
    return f"{MONTHS_DE[d.month - 1]} {d.year}"


def same_month(d, month):
    return d is not None and d.year == month.year and d.month == month.month


def safe_cell(value):
    """Verhindert, dass ein Text im Sheet als Formel gelesen wird."""
    text = str(value if value is not None else "")
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


def norm_name(name):
    text = (name or "").lower()
    text = re.sub(r"\(.*?\)", " ", text)
    text = re.sub(r"[^a-z0-9äöüß ]+", " ", text)
    stop = {"gmbh", "mbh", "ltd", "limited", "ag", "se", "kg", "ug", "co", "holding",
            "holdings", "t", "a", "deutschland", "germany"}
    return " ".join(t for t in text.split() if t not in stop)


def same_customer(a, b):
    na, nb = norm_name(a), norm_name(b)
    if not na or not nb:
        return False
    return na == nb or (len(na) >= 4 and na in nb) or (len(nb) >= 4 and nb in na)


def short_error(exc):
    return (str(exc).strip() or exc.__class__.__name__).splitlines()[0][:300]


# --------------------------------------------------------------------------
# sevDesk
# --------------------------------------------------------------------------

def sevdesk_get(token, path, params=None):
    url = f"{SEVDESK_BASE}{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    last = None
    for attempt in range(3):
        request = urllib.request.Request(url, headers={
            "Authorization": token,            # sevDesk erwartet den Token direkt, ohne "Bearer"
            "Accept": "application/json",
            "User-Agent": "blindleister-finance-cockpit",
        })
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", "replace")[:200]
            if exc.code in (429, 500, 502, 503) and attempt < 2:
                time.sleep(10 * (attempt + 1))
                last = exc
                continue
            if exc.code in (401, 403):
                raise RuntimeError(
                    f"sevDesk HTTP {exc.code}: Token ungültig oder Tarif ohne API-Zugang. {body}") from exc
            raise RuntimeError(f"sevDesk HTTP {exc.code}: {body}") from exc
        except (urllib.error.URLError, TimeoutError) as exc:
            last = exc
            time.sleep(5)
    raise RuntimeError(f"sevDesk nicht erreichbar: {short_error(last)}")


def contact_name(contact):
    if not isinstance(contact, dict):
        return ""
    if contact.get("name"):
        return str(contact["name"]).strip()
    person = " ".join(str(contact.get(k) or "").strip() for k in ("surename", "familyname")).strip()
    return person


def fetch_raw_invoices(token):
    """Alle Rechnungen seitenweise holen. 'embed=contact' bringt den Kundennamen gleich mit;
    klappt das nicht, wird ohne Embed erneut geholt und die Namen über /Contact aufgelöst."""
    def pages(extra):
        offset, out = 0, []
        while True:
            params = {"limit": 1000, "offset": offset}
            params.update(extra)
            data = sevdesk_get(token, "/Invoice", params)
            objects = data.get("objects") or []
            out += objects
            if len(objects) < 1000:
                return out
            offset += 1000

    try:
        return pages({"embed": "contact"})
    except RuntimeError as exc:
        if "HTTP 4" in str(exc) and "HTTP 401" not in str(exc) and "HTTP 403" not in str(exc):
            return pages({})
        raise


def fetch_contact_names(token):
    names, offset = {}, 0
    while True:
        data = sevdesk_get(token, "/Contact", {"limit": 1000, "offset": offset, "depth": 0})
        objects = data.get("objects") or []
        for c in objects:
            names[str(c.get("id"))] = contact_name(c)
        if len(objects) < 1000:
            return names
        offset += 1000


def parse_invoice(raw, names_by_id=None):
    contact = raw.get("contact") or {}
    cid = str(contact.get("id")) if isinstance(contact, dict) and contact.get("id") else ""
    customer = contact_name(contact) or (names_by_id or {}).get(cid, "")
    if not customer:
        customer = str(raw.get("addressName") or "").splitlines()[0].strip() if raw.get("addressName") else ""
    if not customer:
        customer = f"Kontakt {cid}" if cid else "(unbekannt)"

    try:
        status = int(raw.get("status"))
    except (TypeError, ValueError):
        status = 0
    invoice_date = parse_date(raw.get("invoiceDate"))
    try:
        pay_days = int(raw.get("timeToPay")) if raw.get("timeToPay") not in (None, "") else DEFAULT_PAYMENT_DAYS
    except (TypeError, ValueError):
        pay_days = DEFAULT_PAYMENT_DAYS
    due = invoice_date + timedelta(days=pay_days) if invoice_date else None

    return {
        "id": str(raw.get("id") or ""),
        "number": str(raw.get("invoiceNumber") or ""),
        "type": str(raw.get("invoiceType") or "RE"),
        "status": status,
        "customer": customer,
        "date": invoice_date,
        "due": due,
        "gross": parse_amount(raw.get("sumGross")),
        "net": parse_amount(raw.get("sumNet")),
        "paid_amount": parse_amount(raw.get("paidAmount")),
        "pay_date": parse_date(raw.get("payDate")),
    }


def load_invoices(token):
    raw = fetch_raw_invoices(token)
    needs_names = any(
        not contact_name(r.get("contact")) and not r.get("addressName") for r in raw if r.get("contact"))
    names = fetch_contact_names(token) if needs_names else None
    return [parse_invoice(r, names) for r in raw]


def check_sevdesk(token):
    """Verbindungstest: zeigt, welche Felder und Status die API wirklich liefert."""
    data = sevdesk_get(token, "/Invoice", {"limit": 5, "embed": "contact"})
    objects = data.get("objects") or []
    print(f"sevDesk erreichbar, {len(objects)} Beispiel-Rechnung(en) geladen.")
    if objects:
        print("Felder:", ", ".join(sorted(objects[0].keys())))
        for o in objects:
            inv = parse_invoice(o)
            print(f"  {inv['number'] or '-'} | {inv['customer']} | Status {inv['status']} | "
                  f"{inv['date']} | fällig {inv['due']} | brutto {inv['gross']}")
    return objects


# --------------------------------------------------------------------------
# Regeln: Status, Mahnstufe
# --------------------------------------------------------------------------

def classify(inv, today):
    """Ergänzt eine Rechnung um Zustand, offenen Betrag, Tage überfällig und Mahnvorschlag."""
    inv = dict(inv)
    inv.update({"state": "", "open": 0.0, "days_overdue": 0, "dunning": ""})

    if inv["type"] == "SR":
        inv["state"] = "STORNO"
        return inv
    if inv["type"] not in COUNTED_TYPES or inv["status"] < STATUS_DRAFT:
        inv["state"] = "IGNORIERT"          # Wiederkehr-Vorlagen, Mahnbelege, deaktiviert
        return inv

    if inv["status"] == STATUS_DRAFT:
        inv["state"] = "ENTWURF"
    elif inv["status"] == STATUS_PAID:
        inv["state"] = "BEZAHLT"
    elif inv["status"] in (STATUS_OPEN, STATUS_PARTIAL):
        inv["open"] = max(inv["gross"] - inv["paid_amount"], 0.0)
        overdue = (today - inv["due"]).days if inv["due"] else 0
        if overdue > 0:
            inv["state"] = "ÜBERFÄLLIG"
            inv["days_overdue"] = overdue
            inv["dunning"] = next((label for days, label in DUNNING_STAGES if overdue >= days),
                                  "überfällig, noch abwarten")
        else:
            inv["state"] = "OFFEN"
        if inv["status"] == STATUS_PARTIAL:
            inv["dunning"] = (inv["dunning"] + " (teilbezahlt)").strip()
    else:
        inv["state"] = "UNBEKANNT"
    return inv


# --------------------------------------------------------------------------
# Vertragsstamm: was müssen wir noch stellen?
# --------------------------------------------------------------------------

def parse_contracts(rows):
    """rows = Tab als Liste von Zeilen (erste Zeile = Kopf). Gibt (Verträge, Hinweise) zurück."""
    if not rows:
        return [], ["Vertragsstamm ist leer."]
    header = [h.strip().lower() for h in rows[0]]
    wanted = {name.lower(): i for i, name in enumerate(CONTRACT_HEADER)}
    index = {}
    for name in CONTRACT_HEADER:
        if name.lower() in header:
            index[name] = header.index(name.lower())
    missing = [n for n in CONTRACT_HEADER[:4] if n not in index]
    if missing:
        return [], [f"Vertragsstamm: Spalte(n) fehlen: {', '.join(missing)}."]

    def cell(row, name):
        i = index.get(name)
        return row[i].strip() if i is not None and i < len(row) else ""

    contracts, notes = [], []
    for row in rows[1:]:
        customer = cell(row, "Kunde (wie in sevDesk)")
        if not customer or cell(row, "Aktiv").upper() != "JA":
            continue
        rhythm = cell(row, "Rhythmus").lower() or "monatlich"
        if rhythm not in INTERVAL_MONTHS:
            notes.append(f"{customer}: Rhythmus '{rhythm}' unbekannt, wird als monatlich gerechnet.")
            rhythm = "monatlich"
        first = parse_date(cell(row, "Erste Rechnung"))
        if INTERVAL_MONTHS[rhythm] > 1 and not first:
            notes.append(f"{customer}: Rhythmus {rhythm} braucht ein Datum bei 'Erste Rechnung', "
                         "wird übersprungen.")
            continue
        try:
            day = int(cell(row, "Rechnungstag") or 1)
        except ValueError:
            day = 1
        contracts.append({
            "customer": customer,
            "product": cell(row, "Produkt"),
            "net": parse_amount(cell(row, "Betrag netto")),
            "interval": INTERVAL_MONTHS[rhythm],
            "first": first,
            "end": parse_date(cell(row, "Ende")),
            "day": min(max(day, 1), 28),
        })
    return contracts, notes


def contract_due_in(contract, month):
    """Ist für diesen Vertrag in `month` (1. des Monats) eine Rechnung vorgesehen?"""
    if contract["end"] and month > month_start(contract["end"]):
        return False
    first = contract["first"]
    if first:
        if month < month_start(first):
            return False
        offset = (month.year - first.year) * 12 + (month.month - first.month)
        return offset % contract["interval"] == 0
    return contract["interval"] == 1


def expected_vs_issued(contracts, invoices, today):
    """Pro Monat (Vormonat und laufender Monat) und Kunde: ausgestellt, Entwurf oder fehlt."""
    rows = []
    months = [add_months(month_start(today), -1), month_start(today)]
    customers = []
    for c in contracts:
        if not any(same_customer(c["customer"], k) and c["customer"] == k for k in customers):
            customers.append(c["customer"])

    for month in months:
        for customer in customers:
            lines = [c for c in contracts if c["customer"] == customer and contract_due_in(c, month)]
            if not lines:
                continue
            expected = sum(c["net"] for c in lines)
            day = min(c["day"] for c in lines)
            found = [i for i in invoices
                     if i["state"] in ("OFFEN", "ÜBERFÄLLIG", "BEZAHLT", "ENTWURF")
                     and same_month(i["date"], month) and same_customer(customer, i["customer"])]
            sent = [i for i in found if i["state"] != "ENTWURF"]
            drafts = [i for i in found if i["state"] == "ENTWURF"]
            numbers = ", ".join(i["number"] or "(ohne Nr.)" for i in found)
            issued_net = sum(i["net"] for i in sent)

            note = ""
            if sent:
                status = "AUSGESTELLT"
                if expected and abs(issued_net - expected) > max(1.0, 0.01 * expected):
                    status = "BETRAG WEICHT AB"
                    note = f"erwartet {eur(expected)}, ausgestellt {eur(issued_net)} netto"
            elif drafts:
                status = "ENTWURF"
                note = "Rechnung liegt als Entwurf in sevDesk, noch nicht versendet"
            elif month == month_start(today) and today.day < day:
                status = "NOCH NICHT FÄLLIG"
                note = f"geplant ab dem {day}."
            else:
                status = "FEHLT"
                note = "keine Rechnung für diesen Monat in sevDesk gefunden"
            rows.append({
                "month": month, "customer": customer, "expected": expected, "status": status,
                "numbers": numbers, "issued": issued_net, "note": note,
            })
    return rows


# --------------------------------------------------------------------------
# Auswertung
# --------------------------------------------------------------------------

def build_report(invoices, contracts, contract_notes, today):
    this_month = month_start(today)
    prev_month = add_months(this_month, -1)
    classified = [classify(i, today) for i in invoices]
    real = [i for i in classified if i["state"] in ("OFFEN", "ÜBERFÄLLIG", "BEZAHLT", "ENTWURF")]

    def total(items, key="gross"):
        return round(sum(i[key] for i in items), 2)

    paid_now = [i for i in real if i["state"] == "BEZAHLT" and same_month(i["pay_date"], this_month)]
    paid_prev = [i for i in real if i["state"] == "BEZAHLT" and same_month(i["pay_date"], prev_month)]
    issued_now = [i for i in real if i["state"] != "ENTWURF" and same_month(i["date"], this_month)]
    issued_prev = [i for i in real if i["state"] != "ENTWURF" and same_month(i["date"], prev_month)]
    open_items = [i for i in real if i["state"] == "OFFEN"]
    overdue = sorted((i for i in real if i["state"] == "ÜBERFÄLLIG"),
                     key=lambda i: -i["days_overdue"])
    drafts = [i for i in real if i["state"] == "ENTWURF"]

    expected = expected_vs_issued(contracts, classified, today) if contracts else []
    notes = list(contract_notes)
    if not contracts:
        notes.append("Der Vertragsstamm ist leer: 'noch zu stellen' wurde nicht geprüft.")

    return {
        "today": today, "this_month": this_month, "prev_month": prev_month,
        "invoices": classified,
        "paid_now": paid_now, "paid_prev": paid_prev,
        "issued_now": issued_now, "issued_prev": issued_prev,
        "open": open_items, "overdue": overdue, "drafts": drafts,
        "totals": {
            "paid_now": total(paid_now), "paid_prev": total(paid_prev),
            "issued_now": total(issued_now), "issued_prev": total(issued_prev),
            "open": total(open_items, "open"), "overdue": total(overdue, "open"),
        },
        "expected": expected,
        "missing": [e for e in expected if e["status"] == "FEHLT"],
        "notes": notes,
    }


def report_text(report):
    """Der deterministische Bericht (immer vorhanden, auch ohne KI)."""
    t, today = report["totals"], report["today"]
    cur, prev = month_label(report["this_month"]), month_label(report["prev_month"])
    lines = [f"Stand: {today:%d.%m.%Y}", "", "AUF EINEN BLICK",
             f"- Bezahlt im {cur}: {len(report['paid_now'])} Rechnung(en), {eur(t['paid_now'])} "
             f"({prev}: {len(report['paid_prev'])}, {eur(t['paid_prev'])})",
             f"- Ausgestellt im {cur}: {len(report['issued_now'])}, {eur(t['issued_now'])} brutto "
             f"({prev}: {len(report['issued_prev'])}, {eur(t['issued_prev'])})",
             f"- Offen, noch nicht fällig: {len(report['open'])}, {eur(t['open'])}",
             f"- Überfällig: {len(report['overdue'])}, {eur(t['overdue'])}",
             f"- Entwürfe in sevDesk: {len(report['drafts'])}", ""]

    lines.append("ÜBERFÄLLIG UND MAHNVORSCHLÄGE (älteste zuerst)")
    if report["overdue"]:
        for i in report["overdue"]:
            lines.append(f"- {i['number'] or '(ohne Nr.)'} | {i['customer']} | {eur(i['open'])} offen | "
                         f"{days_text(i['days_overdue'])} überfällig | Vorschlag: {i['dunning']}")
    else:
        lines.append("- keine")
    lines.append("")

    lines.append("NOCH ZU STELLEN (laut Vertragsstamm)")
    problems = [e for e in report["expected"] if e["status"] in ("FEHLT", "ENTWURF", "BETRAG WEICHT AB")]
    if problems:
        for e in problems:
            lines.append(f"- {month_label(e['month'])} | {e['customer']} | erwartet {eur(e['expected'])} "
                         f"netto | {e['status']} | {e['note']}")
    elif report["expected"]:
        lines.append("- alles ausgestellt")
    else:
        lines.append("- nicht geprüft (kein Vertragsstamm)")
    lines.append("")

    if report["notes"]:
        lines.append("HINWEISE")
        lines += [f"- {n}" for n in report["notes"]]
        lines.append("")
    lines.append("Mahnungen und Zahlungen gibt ein Mensch frei; dieser Bericht verschickt nichts an Kunden.")
    return "\n".join(lines)


def ai_payload(report):
    t = report["totals"]
    return json.dumps({
        "stand": report["today"].isoformat(),
        "monat": month_label(report["this_month"]),
        "bezahlt_laufender_monat": {"anzahl": len(report["paid_now"]), "brutto": t["paid_now"]},
        "bezahlt_vormonat": {"anzahl": len(report["paid_prev"]), "brutto": t["paid_prev"]},
        "ausgestellt_laufender_monat": {"anzahl": len(report["issued_now"]), "brutto": t["issued_now"]},
        "offen_nicht_faellig": {"anzahl": len(report["open"]), "betrag": t["open"]},
        "ueberfaellig": {"anzahl": len(report["overdue"]), "betrag": t["overdue"], "liste": [
            {"nr": i["number"], "kunde": i["customer"], "offen": round(i["open"], 2),
             "tage": i["days_overdue"], "vorschlag": i["dunning"]} for i in report["overdue"][:15]]},
        "entwuerfe": len(report["drafts"]),
        "noch_zu_stellen": [
            {"monat": month_label(e["month"]), "kunde": e["customer"], "erwartet_netto": round(e["expected"], 2),
             "status": e["status"], "hinweis": e["note"]}
            for e in report["expected"] if e["status"] in ("FEHLT", "ENTWURF", "BETRAG WEICHT AB")],
        "hinweise": report["notes"],
    }, ensure_ascii=False)


# --------------------------------------------------------------------------
# KI-Zusammenfassung
# --------------------------------------------------------------------------

def _post_json(url, headers, body, timeout=120):
    for attempt in range(4):
        request = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST",
                                         headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")
            try:
                detail = json.loads(detail)["error"]["message"]
            except Exception:
                detail = " ".join(detail.split())
            if exc.code in (429, 500, 503, 529) and attempt < 3:
                time.sleep(20 * 2 ** attempt)
                continue
            raise RuntimeError(f"KI HTTP {exc.code}: {detail[:200]}") from exc
        except (TimeoutError, urllib.error.URLError) as exc:
            if attempt < 3:
                time.sleep(20)
                continue
            raise RuntimeError(f"KI nicht erreichbar: {short_error(exc)}") from exc


def ai_summary(report):
    """Claude bevorzugt, sonst Gemini. Gibt (Text, Hinweis) zurück; Text ist '' wenn keine KI läuft."""
    payload = ai_payload(report)
    claude_key = os.environ.get("ANTHROPIC_API_KEY")
    gemini_key = os.environ.get("GEMINI_API_KEY")
    try:
        if claude_key:
            data = _post_json(
                "https://api.anthropic.com/v1/messages",
                {"x-api-key": claude_key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
                {"model": os.environ.get("ANTHROPIC_MODEL") or "claude-sonnet-5-5",
                 "max_tokens": 1200, "system": AI_SYSTEM,
                 "messages": [{"role": "user", "content": payload}]})
            return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text").strip(), ""
        if gemini_key:
            model = os.environ.get("GEMINI_MODEL") or "gemini-flash-latest"
            data = _post_json(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                {"Content-Type": "application/json", "x-goog-api-key": gemini_key},
                {"systemInstruction": {"parts": [{"text": AI_SYSTEM}]},
                 "contents": [{"role": "user", "parts": [{"text": payload}]}],
                 "generationConfig": {"temperature": 0.2}})
            parts = (data.get("candidates") or [{}])[0].get("content", {}).get("parts", [])
            return "".join(p.get("text", "") for p in parts if not p.get("thought")).strip(), ""
        return "", "Keine KI konfiguriert (ANTHROPIC_API_KEY oder GEMINI_API_KEY): es gilt der berechnete Bericht."
    except Exception as exc:
        return "", f"KI-Zusammenfassung nicht verfügbar ({short_error(exc)}); es gilt der berechnete Bericht."


# --------------------------------------------------------------------------
# E-Mail
# --------------------------------------------------------------------------

def smtp_settings():
    host, user, password = (os.environ.get(k) for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD"))
    if not (host and user and password):
        return None
    return {"host": host, "user": user, "password": password,
            "port": int(os.environ.get("SMTP_PORT") or 465),
            "sender": os.environ.get("SMTP_FROM") or user}


def build_subject(report, demo=False):
    n_over, amount = len(report["overdue"]), report["totals"]["overdue"]
    n_missing = len(report["missing"])
    parts = [f"{n_over} überfällig ({eur(amount)})"]
    if report["expected"]:
        parts.append(f"{n_missing} Rechnung(en) fehlen")
    return f"{'[DEMO] ' if demo else ''}Rechnungs-Cockpit {report['today']:%d.%m.%Y}: " + ", ".join(parts)


def send_mail(subject, body):
    settings = smtp_settings()
    if not settings:
        print("SMTP nicht konfiguriert: es wird keine Mail verschickt.")
        return False
    recipients = os.environ.get("FINANCE_EMAIL") or os.environ.get("ALERT_EMAIL") or "info@blindleister.de"
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, settings["sender"], recipients
    msg.set_content(body)
    if settings["port"] == 465:
        with smtplib.SMTP_SSL(settings["host"], settings["port"], timeout=30) as smtp:
            smtp.login(settings["user"], settings["password"])
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(settings["host"], settings["port"], timeout=30) as smtp:
            smtp.starttls()
            smtp.login(settings["user"], settings["password"])
            smtp.send_message(msg)
    print(f"Mail an {recipients} verschickt: {subject}")
    return True


# --------------------------------------------------------------------------
# Google Sheet
# --------------------------------------------------------------------------

def open_sheet(sheet_key):
    import gspread
    from google.oauth2.service_account import Credentials

    raw = os.environ.get("GOOGLE_CREDENTIALS_B64")
    if not raw:
        raise RuntimeError("GOOGLE_CREDENTIALS_B64 ist nicht gesetzt.")
    info = json.loads(base64.b64decode(raw))
    key = info.get("private_key", "")
    if "\\n" in key and "\n" not in key:
        info["private_key"] = key.replace("\\n", "\n")
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds).open_by_key(sheet_key)


def get_or_create_tab(spreadsheet, name, header, rows=500):
    import gspread
    try:
        return spreadsheet.worksheet(name), False
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=rows, cols=max(len(header), 10))
        ws.update(range_name="A1", values=[header], value_input_option="RAW")
        try:
            ws.freeze(rows=1)
            ws.format(f"A1:{chr(64 + len(header))}1", {"textFormat": {"bold": True}})
        except Exception:
            pass
        return ws, True


def write_table(ws, header, rows):
    ws.clear()
    needed = len(rows) + 2
    if ws.row_count < needed:
        ws.add_rows(needed - ws.row_count + 50)
    ws.update(range_name="A1", values=[header] + rows, value_input_option="USER_ENTERED")
    try:
        ws.freeze(rows=1)
        ws.format(f"A1:{chr(64 + len(header))}1", {"textFormat": {"bold": True}})
    except Exception:
        pass


def invoice_rows(report, keep_days=400):
    today = report["today"]
    order = {"ÜBERFÄLLIG": 0, "ENTWURF": 1, "OFFEN": 2, "BEZAHLT": 3}
    items = [i for i in report["invoices"] if i["state"] in order
             and (i["state"] != "BEZAHLT" or (i["date"] and (today - i["date"]).days <= keep_days))]
    items.sort(key=lambda i: (order[i["state"]], -i["days_overdue"], -(i["date"].toordinal() if i["date"] else 0)))
    return [[
        safe_cell(i["number"]), safe_cell(i["customer"]),
        i["date"].isoformat() if i["date"] else "", i["due"].isoformat() if i["due"] else "",
        round(i["gross"], 2), round(i["net"], 2), round(i["paid_amount"], 2), round(i["open"], 2),
        i["state"], i["days_overdue"] or "", safe_cell(i["dunning"]),
        i["pay_date"].isoformat() if i["pay_date"] else "", i["id"],
    ] for i in items]


def expected_rows(report):
    return [[month_label(e["month"]), safe_cell(e["customer"]), round(e["expected"], 2), e["status"],
             safe_cell(e["numbers"]), round(e["issued"], 2), safe_cell(e["note"])]
            for e in report["expected"]]


def write_sheet(report, subject, summary_text, sheet_key, demo=False):
    """Schreibt die Auswertung ins Sheet. Der Demo-Modus nimmt eigene Tabs mit Präfix 'DEMO_'
    und fasst die echten Tabs (und den Vertragsstamm) nie an."""
    spreadsheet = open_sheet(sheet_key)
    prefix = "DEMO_" if demo else ""
    if not demo:
        ensure_contracts_tab(spreadsheet)
    write_table(get_or_create_tab(spreadsheet, prefix + TAB_INVOICES, INVOICE_HEADER)[0],
                INVOICE_HEADER, invoice_rows(report))
    write_table(get_or_create_tab(spreadsheet, prefix + TAB_EXPECTED, EXPECTED_HEADER)[0],
                EXPECTED_HEADER, expected_rows(report))
    summary_ws, _ = get_or_create_tab(spreadsheet, prefix + TAB_SUMMARY,
                                      ["Zeitpunkt", "Betreff", "Zusammenfassung"])
    stamp = datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M")
    summary_ws.append_row([stamp, subject, summary_text[:45000]], value_input_option="RAW")


def ensure_contracts_tab(spreadsheet):
    """Legt den Tab 'Vertragsstamm' mit Kopfzeile und einer Beispielzeile an, falls er fehlt."""
    ws, created = get_or_create_tab(spreadsheet, TAB_CONTRACTS, CONTRACT_HEADER, rows=100)
    if created:
        ws.update(range_name="A2", values=[[
            "BEISPIEL Kunde GmbH (Zeile löschen)", "GDI", 800, "monatlich", "2026-01-01", "", 1, "NEIN",
            "Name genau wie in sevDesk eintragen"]], value_input_option="RAW")
    return ws, created


def read_contracts_from_sheet(sheet_key):
    spreadsheet = open_sheet(sheet_key)
    ws, created = ensure_contracts_tab(spreadsheet)
    if created:
        return [], ["Der Tab 'Vertragsstamm' wurde neu angelegt und ist noch leer: bitte ausfüllen."]
    return parse_contracts(ws.get_all_values())


# --------------------------------------------------------------------------
# Demo-Daten (für die Vorführung ohne sevDesk-Zugang)
# --------------------------------------------------------------------------

def demo_data(today):
    """Beispieldaten, relativ zum heutigen Datum, sodass jede Mahnstufe einmal vorkommt."""
    this_m = month_start(today)
    prev_m = add_months(this_m, -1)
    aug_15 = add_months(prev_m, -1) + timedelta(days=14)
    sep_3, sep_21, oct_2 = prev_m + timedelta(days=2), prev_m + timedelta(days=20), this_m + timedelta(days=1)

    def paid_on(d):
        return min(today, d)

    def inv(num, customer, d, gross, status, pay=None, paid=0.0, ptype="RE"):
        return {"id": f"demo-{num}", "number": num, "type": ptype, "status": status, "customer": customer,
                "date": d, "due": d + timedelta(days=14), "gross": gross, "net": round(gross / 1.19, 2),
                "paid_amount": paid, "pay_date": pay}

    invoices = [
        # Vormonat
        inv("RE-2026-201", "Belectric GmbH", sep_3, 952.00, STATUS_PAID, pay=prev_m + timedelta(days=19)),
        inv("RE-2026-202", "Grenergy Erneuerbare Energien GmbH", sep_3, 1071.00, STATUS_PAID, pay=prev_m + timedelta(days=24)),
        inv("RE-2026-203", "Noveria Energie GmbH", sep_3, 1071.00, STATUS_OPEN),                      # lange überfällig
        inv("RE-2026-204", "Aura Power Germany GmbH", sep_21, 357.00, STATUS_OPEN),                  # wenige Tage überfällig
        inv("RE-2026-205", "Southside Energy GmbH", sep_3, 1557.71, STATUS_PAID, pay=prev_m + timedelta(days=16)),
        inv("RE-2026-206", "Lightsource Holdings 2 Limited", sep_3, 1416.10, STATUS_PARTIAL, paid=700.00),
        inv("RE-2026-207", "Recurrent Energy Deutschland GmbH", sep_3, 1547.00, STATUS_DRAFT),         # nie versendet
        inv("RE-2026-208", "Luminous Energy Deutschland GmbH", sep_3, 1416.10, STATUS_PAID, pay=prev_m + timedelta(days=15)),
        inv("RE-2026-170", "Renergo Projekte GmbH", aug_15, 849.66, STATUS_OPEN),                    # August, sehr lange offen
        # laufender Monat
        inv("RE-2026-209", "Belectric GmbH", oct_2, 952.00, STATUS_PAID, pay=paid_on(this_m + timedelta(days=8))),
        inv("RE-2026-210", "Grenergy Erneuerbare Energien GmbH", oct_2, 1071.00, STATUS_OPEN),
        inv("RE-2026-211", "Southside Energy GmbH", oct_2, 1557.71, STATUS_PAID, pay=paid_on(this_m + timedelta(days=11))),
        inv("RE-2026-212", "Lightsource Holdings 2 Limited", oct_2, 1416.10, STATUS_OPEN),
        inv("RE-2026-213", "Renergo Projekte GmbH", oct_2, 849.66, STATUS_DRAFT),
        inv("RE-2026-214", "Luminous Energy Deutschland GmbH", oct_2, 1547.00, STATUS_OPEN),           # Betrag weicht ab
        inv("WKR-1", "Belectric GmbH", aug_15, 952.00, 50, ptype="WKR"),                              # wird ignoriert
    ]
    contracts = [
        {"customer": c, "product": "GDI", "net": n, "interval": 1, "first": date(2025, 1, 1), "end": None, "day": 1}
        for c, n in [("Belectric GmbH", 800), ("Grenergy Erneuerbare Energien GmbH", 900),
                     ("Noveria Energie GmbH", 900), ("Renergo Projekte GmbH", 714),
                     ("Aura Power Germany GmbH", 300), ("Southside Energy GmbH", 1309),
                     ("Lightsource Holdings 2 Limited", 1190), ("Recurrent Energy Deutschland GmbH", 1300),
                     ("Luminous Energy Deutschland GmbH", 1190)]]
    return invoices, contracts


# --------------------------------------------------------------------------
# Hauptlauf
# --------------------------------------------------------------------------

def run(demo=False, dry_run=False):
    today = today_berlin()
    sheet_key = os.environ.get("FINANCE_SHEET_KEY")

    if demo:
        invoices, contracts = demo_data(today)
        contract_notes = ["DEMO: Beispieldaten, keine echten Rechnungen."]
    else:
        token = os.environ.get("SEVDESK_API_TOKEN")
        if not token:
            raise RuntimeError("SEVDESK_API_TOKEN ist nicht gesetzt (oder --demo benutzen).")
        invoices = load_invoices(token)
        print(f"sevDesk: {len(invoices)} Rechnungen geladen.")
        if dry_run or not sheet_key:
            contracts, contract_notes = [], ["Kein Sheet konfiguriert: Vertragsstamm nicht gelesen."]
        else:
            contracts, contract_notes = read_contracts_from_sheet(sheet_key)

    report = build_report(invoices, contracts, contract_notes, today)
    base_text = report_text(report)
    ai_text, ai_note = ai_summary(report)
    if ai_note:
        print(ai_note)

    subject = build_subject(report, demo)
    body = ((ai_text + "\n\n" + "=" * 40 + "\nDETAILS (berechnet)\n" + "=" * 40 + "\n") if ai_text else "") + base_text
    if ai_note:
        body += f"\n\n({ai_note})"

    print(body)
    if dry_run:
        print("\n--dry-run: nichts geschrieben, nichts verschickt.")
        return report

    if sheet_key:
        write_sheet(report, subject, ai_text or base_text, sheet_key, demo=demo)
        print("Google Sheet aktualisiert" + (" (Tabs mit Präfix DEMO_)." if demo else "."))
    else:
        print("FINANCE_SHEET_KEY nicht gesetzt: Sheet wird nicht aktualisiert.")

    send_mail(subject, body)
    return report


if __name__ == "__main__":
    args = set(sys.argv[1:])
    if "--check" in args:
        check_sevdesk(os.environ.get("SEVDESK_API_TOKEN") or sys.exit("SEVDESK_API_TOKEN fehlt."))
    else:
        run(demo="--demo" in args, dry_run="--dry-run" in args)
