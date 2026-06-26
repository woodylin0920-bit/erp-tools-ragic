"""IMAP 讀信（共用、唯讀，不更動信箱）。

電商訂單通知統一寄到 info@boptoys.com.tw（Google Workspace）。
用應用程式密碼經 IMAP 讀取，密碼存於 ~/.boptoys-info_app_password。
"""
import email
import imaplib
import os
import re
from email.header import decode_header


def decode_hdr(s: str) -> str:
    if not s:
        return ""
    return "".join(
        t.decode(enc or "utf-8", "ignore") if isinstance(t, bytes) else t
        for t, enc in decode_header(s)
    )


def connect(user: str, app_password_path: str) -> imaplib.IMAP4_SSL:
    pw = open(os.path.expanduser(app_password_path)).read().strip()
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(user, pw)
    M.select("INBOX", readonly=True)
    return M


def search(M: imaplib.IMAP4_SSL, gmail_query: str) -> list:
    """Gmail X-GM-RAW 搜尋（支援 UTF-8 中文查詢）。回傳 UID list。"""
    M.literal = gmail_query.encode("utf-8")
    typ, data = M.uid("SEARCH", "CHARSET", "UTF-8", "X-GM-RAW")
    return data[0].split() if data and data[0] else []


def _body_of(msg) -> str:
    for p in msg.walk():
        if p.get_content_type() == "text/plain":
            t = (p.get_payload(decode=True) or b"").decode(
                p.get_content_charset() or "utf-8", "ignore")
            if t.strip():
                return t
    for p in msg.walk():
        if p.get_content_type() == "text/html":
            h = (p.get_payload(decode=True) or b"").decode(
                p.get_content_charset() or "utf-8", "ignore")
            h = re.sub(r"<(style|script|head)[^>]*>.*?</\1>", "", h, flags=re.S | re.I)
            return re.sub(r"[ \t]+", " ", re.sub(r"<[^>]+>", " ", h))
    return ""


def fetch(M: imaplib.IMAP4_SSL, uid):
    """取 (主旨, 內文)。部分平台（如蝦皮）的訂單號/買家在主旨。"""
    typ, md = M.uid("FETCH", uid, "(BODY.PEEK[])")
    msg = email.message_from_bytes(md[0][1])
    return decode_hdr(msg.get("Subject", "")), _body_of(msg)


def body(M: imaplib.IMAP4_SSL, uid) -> str:
    """取信件內文（向後相容）。"""
    return fetch(M, uid)[1]
