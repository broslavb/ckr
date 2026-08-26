#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Витягує номери ТТН із листів avto.pro і кладе їх у Firestore.

Лист має вигляд:
    До вашого замовлення № 10514428 (https://avtopro.ua/orders/10514428)
    створено ТТН 20451519906235.

Пара «номер замовлення → ТТН» лягає в колекцію avtopro_ttn документом з
іменем = номер замовлення. Самі калькуляції скрипт не чіпає: ТТН до
позицій підставляє вже сам CKR (кнопка «Підтягнути ТТН» у вкладці
«Отримання»), бо тільки він знає внутрішній формат калькуляцій.

Залежностей немає — лише стандартна бібліотека Python 3.
"""

import email
import email.header
import html as html_mod
import imaplib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(HERE, "config.txt")
STATE_FILE = os.path.join(HERE, ".state.json")
LOG_FILE = os.path.join(HERE, "avtopro_ttn.log")

FIREBASE_PROJECT = "car-calculator-2dbd2"
FIREBASE_API_KEY = "AIzaSyBz8Imv904HpJdtSjJxb2UCi5nWGugvCBA"  # той самий, що в index.html
COLLECTION = "avtopro_ttn"


def log(msg):
    line = "%s  %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


# ── Налаштування ────────────────────────────────────────────────────────────

def load_env():
    """Читає tools/config.txt у вигляді KEY=VALUE. Значення в лапках теж приймає."""
    cfg = {}
    if not os.path.exists(CONFIG_FILE):
        log("НЕМАЄ файлу %s — скопіюй config.example.txt у config.txt і заповни"
            % CONFIG_FILE)
        sys.exit(1)
    with open(CONFIG_FILE, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            v = v.strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
                v = v[1:-1]
            cfg[k.strip()] = v
    missing = [k for k in ("IMAP_USER", "IMAP_PASS", "FB_EMAIL", "FB_PASS") if not cfg.get(k)]
    if missing:
        log("У config.txt не заповнено: %s" % ", ".join(missing))
        sys.exit(1)
    return cfg


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ── Розбір листа ────────────────────────────────────────────────────────────

# «до вашого замовлення № 10514428 … створено ТТН 20451519906235»
RE_PAIR = re.compile(
    r"замовлен\w*\s*[№#]?\s*(\d{5,12}).{0,400}?ТТН\s*[:№]?\s*(\d{10,18})",
    re.IGNORECASE | re.DOTALL,
)
RE_ORDER_URL = re.compile(r"avtopro\.ua/orders/(\d{5,12})", re.IGNORECASE)
RE_ORDER_TXT = re.compile(r"замовлен\w*\s*[№#]\s*(\d{5,12})", re.IGNORECASE)
RE_TTN = re.compile(r"ТТН\s*[:№]?\s*(\d{10,18})", re.IGNORECASE)
RE_TAGS = re.compile(r"<[^>]+>")
# У листах avto.pro усередині HTML лежить великий блок CSS — його треба
# викинути цілком, інакше стилі вклиняться між номером замовлення і ТТН.
RE_STYLE = re.compile(r"<(style|script|head)\b.*?</\1>", re.IGNORECASE | re.DOTALL)


def decode_header(value):
    if not value:
        return ""
    out = []
    for text, enc in email.header.decode_header(value):
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def message_text(msg):
    """Збирає докупи текстову і HTML-частини листа."""
    chunks = []
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        if part.get_content_disposition() == "attachment":
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        charset = part.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if ctype == "text/html":
            # спершу CSS і скрипти цілком, потім теги;
            # посилання в href при цьому лишаються в тексті
            text = RE_STYLE.sub(" ", text)
            text = RE_TAGS.sub(" ", text.replace("href=", " href="))
        # у листах № і пробіли часто екрановані: &#x2116;, &nbsp;
        text = html_mod.unescape(text).replace("\xa0", " ")
        chunks.append(text)
    return "\n".join(chunks)


def extract_pairs(text):
    """Повертає [(номер замовлення, ТТН)] з тексту листа."""
    pairs = RE_PAIR.findall(text)
    if pairs:
        return [(o, t) for o, t in pairs]
    # Запасний шлях: номер замовлення і ТТН знайшлися нарізно
    orders = RE_ORDER_URL.findall(text) or RE_ORDER_TXT.findall(text)
    ttns = RE_TTN.findall(text)
    if len(orders) == 1 and len(ttns) == 1:
        return [(orders[0], ttns[0])]
    if orders and ttns:
        log("  ⚠ у листі %d замовлень і %d ТТН — пропускаю, треба глянути очима"
            % (len(orders), len(ttns)))
    return []


# ── Firestore ───────────────────────────────────────────────────────────────

def http_json(url, payload=None, token=None, method=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/json")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError("HTTP %s: %s" % (e.code, body[:400]))


def firebase_login(email_addr, password):
    url = ("https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key="
           + FIREBASE_API_KEY)
    res = http_json(url, {"email": email_addr, "password": password, "returnSecureToken": True})
    return res["idToken"]


def firestore_put(token, order, ttn, subject):
    url = ("https://firestore.googleapis.com/v1/projects/%s/databases/(default)/documents/%s/%s"
           % (FIREBASE_PROJECT, COLLECTION, order))
    payload = {"fields": {
        "order":   {"stringValue": order},
        "ttn":     {"stringValue": ttn},
        "subject": {"stringValue": subject[:200]},
        "source":  {"stringValue": "email"},
        "at":      {"timestampValue": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")},
    }}
    http_json(url, payload, token=token, method="PATCH")


# ── Пошта ───────────────────────────────────────────────────────────────────

def fetch_new_messages(cfg, state):
    host = cfg.get("IMAP_HOST", "imap.gmail.com")
    port = int(cfg.get("IMAP_PORT", "993"))
    folder = cfg.get("IMAP_FOLDER", "INBOX")
    sender = cfg.get("MAIL_FROM", "").strip()

    mail = imaplib.IMAP4_SSL(host, port)
    try:
        mail.login(cfg["IMAP_USER"], cfg["IMAP_PASS"])
        mail.select(folder)

        # UIDVALIDITY змінилася — нумерація листів інша, стан скидаємо
        typ, data = mail.status(folder, "(UIDVALIDITY)")
        validity = re.search(r"UIDVALIDITY (\d+)", data[0].decode()).group(1)
        if state.get("uidvalidity") != validity:
            log("UIDVALIDITY змінилася — починаю з чистого стану")
            state["uidvalidity"] = validity
            state["last_uid"] = 0

        criteria = ["ALL"] if not sender else ["FROM", sender]
        typ, data = mail.uid("search", None, *criteria)
        uids = [int(u) for u in data[0].split()] if data and data[0] else []
        last_uid = int(state.get("last_uid", 0))
        fresh = [u for u in uids if u > last_uid]
        if not fresh:
            return [], state

        log("Нових листів: %d" % len(fresh))
        out = []
        for uid in fresh:
            typ, data = mail.uid("fetch", str(uid), "(RFC822)")
            if typ != "OK" or not data or not data[0]:
                continue
            msg = email.message_from_bytes(data[0][1])
            out.append((uid, decode_header(msg.get("Subject")), message_text(msg)))
        state["last_uid"] = max(fresh)
        return out, state
    finally:
        try:
            mail.logout()
        except Exception:
            pass


# ── Головне ─────────────────────────────────────────────────────────────────

def main():
    cfg = load_env()
    state = load_state()

    messages, state = fetch_new_messages(cfg, state)
    if not messages:
        log("Нових листів немає")
        save_state(state)
        return

    found = []
    for uid, subject, text in messages:
        pairs = extract_pairs(text)
        if not pairs:
            log("  UID %s «%s» — ТТН не знайдено" % (uid, subject[:60]))
            continue
        for order, ttn in pairs:
            log("  UID %s → замовлення %s, ТТН %s" % (uid, order, ttn))
            found.append((order, ttn, subject))

    if found:
        token = firebase_login(cfg["FB_EMAIL"], cfg["FB_PASS"])
        ok = 0
        for order, ttn, subject in found:
            try:
                firestore_put(token, order, ttn, subject)
                ok += 1
            except RuntimeError as e:
                log("  ✖ %s: %s" % (order, e))
        log("Записано в Firestore: %d з %d" % (ok, len(found)))

    # Стан рухаємо тільки після успішної обробки — інакше лист опрацюється ще раз
    save_state(state)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log("ПОМИЛКА: %s: %s" % (type(e).__name__, e))
        sys.exit(1)
