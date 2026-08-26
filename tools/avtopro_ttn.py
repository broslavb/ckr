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
Разовий запуск без фільтра за відправником (наприклад, щоб розібрати
переслані листи): python3 avtopro_ttn.py --all
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
import urllib.parse
import urllib.request
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(HERE, "config.txt")
STATE_FILE = os.path.join(HERE, ".state.json")
LOG_FILE = os.path.join(HERE, "avtopro_ttn.log")

FIREBASE_PROJECT = "car-calculator-2dbd2"
FIREBASE_API_KEY = "AIzaSyBz8Imv904HpJdtSjJxb2UCi5nWGugvCBA"  # той самий, що в index.html
# Правила бази дозволяють запис у settings, але не в довільні нові колекції,
# тому всі пари лежать одним документом: поле "o<номер замовлення>" → {ttn, at}.
DOC_PATH = "settings/avtopro_ttn"


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

# Номер замовлення avto.pro ставить у тему листа: «[10473219] До вашого …»
RE_SUBJ_ORDER = re.compile(r"\[(\d{5,12})\]")

# «створено ТТН 59001740377579». ТТН Нової Пошти — 14 цифр; коротшу межу
# беремо із запасом, але не нижче 13, щоб не хапати артикули деталей
# (у тілі листа трапляються 11-значні номери на кшталт 51118077277).
RE_TTN = re.compile(r"ТТН\s*[:№]?\s*(\d{13,18})", re.IGNORECASE)

# Запасні шляхи на випадок теми без номера
RE_ORDER_URL = re.compile(r"avtopro\.ua/orders/(\d{5,12})", re.IGNORECASE)
RE_ORDER_TXT = re.compile(r"замовлен\w*\s*[№#]\s*(\d{5,12})", re.IGNORECASE)

RE_TAGS = re.compile(r"<[^>]+>")
# У листах avto.pro усередині HTML лежить великий блок CSS — його треба
# викинути цілком, інакше стилі вклиняться між номером замовлення і ТТН.
RE_STYLE = re.compile(r"<(style|script|head)\b.*?</\1>", re.IGNORECASE | re.DOTALL)
# Навколо номера ТТН avto.pro ставить невидимі символи (zero-width space),
# через які \s у регулярці не спрацьовує. Викидаємо їх з тексту одразу.
RE_ZEROWIDTH = re.compile(r"[\u200b\u200c\u200d\u2060\ufeff\u00ad]")


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
        text = RE_ZEROWIDTH.sub("", text)
        # посилання приходять через трекер AWS із %2F замість /
        text = urllib.parse.unquote(text)
        chunks.append(text)
    return "\n".join(chunks)


def extract_pairs(text, subject=""):
    """Повертає [(номер замовлення, ТТН)] з листа.

    Номер замовлення беремо з теми — це найнадійніше джерело: у тілі
    посилання загорнуті у трекер розсилки, а сам номер може згадуватись
    у кількох місцях.
    """
    subject = RE_ZEROWIDTH.sub("", subject or "")
    ttns = RE_TTN.findall(text)
    if not ttns:
        return []

    orders = RE_SUBJ_ORDER.findall(subject)
    if not orders:
        orders = RE_ORDER_URL.findall(text) or RE_ORDER_TXT.findall(text)
    if not orders:
        log("  ⚠ ТТН є, а номера замовлення не видно — пропускаю")
        return []

    order = orders[0]
    unique_ttns = list(dict.fromkeys(ttns))
    if len(unique_ttns) > 1:
        log("  ⚠ у листі кілька різних ТТН (%s) — беру перший"
            % ", ".join(unique_ttns))
    return [(order, unique_ttns[0])]


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


def firestore_read(token):
    """Повертає поточний вміст settings/avtopro_ttn (порожньо, якщо ще нема)."""
    url = ("https://firestore.googleapis.com/v1/projects/%s/databases/(default)/documents/%s"
           % (FIREBASE_PROJECT, DOC_PATH))
    try:
        return http_json(url, token=token).get("fields", {})
    except RuntimeError as e:
        if "404" in str(e):
            return {}
        raise


def firestore_sync(token, found):
    """Дописує пари в settings/avtopro_ttn одним запитом.

    На одне замовлення avto.pro буває кілька ТТН (різні посилки), тому
    зберігаємо список `all`, а не одне значення. updateMask лишає решту
    документа недоторканою.
    """
    existing = firestore_read(token)
    updates, now = {}, datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    for order, ttn, subject in found:
        key = "o" + order
        prev = updates.get(key)
        if prev is not None:
            known = [v["stringValue"] for v in prev["mapValue"]["fields"]["all"]["arrayValue"]["values"]]
        else:
            cur = existing.get(key, {}).get("mapValue", {}).get("fields", {})
            known = [v["stringValue"] for v in cur.get("all", {}).get("arrayValue", {}).get("values", [])]
            if not known and cur.get("ttn"):
                known = [cur["ttn"]["stringValue"]]          # запис старого формату
        if ttn not in known:
            known.append(ttn)
        updates[key] = {"mapValue": {"fields": {
            "ttn":     {"stringValue": known[0]},            # для сумісності
            "all":     {"arrayValue": {"values": [{"stringValue": t} for t in known]}},
            "subject": {"stringValue": subject[:200]},
            "at":      {"timestampValue": now},
        }}}

    if not updates:
        return 0
    mask = "&".join("updateMask.fieldPaths=" + k for k in updates)
    url = ("https://firestore.googleapis.com/v1/projects/%s/databases/(default)/documents/%s?%s"
           % (FIREBASE_PROJECT, DOC_PATH, mask))
    http_json(url, {"fields": updates}, token=token, method="PATCH")
    return len(updates)


# ── Пошта ───────────────────────────────────────────────────────────────────

def fetch_new_messages(cfg, state):
    host = cfg.get("IMAP_HOST", "imap.gmail.com")
    port = int(cfg.get("IMAP_PORT", "993"))
    folder = cfg.get("IMAP_FOLDER", "INBOX")
    sender = "" if "--all" in sys.argv else cfg.get("MAIL_FROM", "").strip()

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
        pairs = extract_pairs(text, subject)
        if not pairs:
            log("  UID %s «%s» — ТТН не знайдено" % (uid, subject[:60]))
            continue
        for order, ttn in pairs:
            log("  UID %s → замовлення %s, ТТН %s" % (uid, order, ttn))
            found.append((order, ttn, subject))

    if found:
        token = firebase_login(cfg["FB_EMAIL"], cfg["FB_PASS"])
        try:
            n = firestore_sync(token, found)
            log("Записано в базу: %d замовлень (%d ТТН)" % (n, len(found)))
        except RuntimeError as e:
            log("  ✖ запис у базу не вдався: %s" % e)
            return   # стан не рухаємо — на наступному запуску спробуємо ще раз

    # Стан рухаємо тільки після успішної обробки — інакше лист опрацюється ще раз
    save_state(state)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log("ПОМИЛКА: %s: %s" % (type(e).__name__, e))
        sys.exit(1)
