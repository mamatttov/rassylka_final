#!/usr/bin/env python3
"""Диагностика подключения к почте: python smtp_check.py
Проверяет настройки из smtp.ini на портах 465 и 587, ничего не отправляет и не печатает пароль."""
import configparser
import os
import smtplib
import ssl
import sys
from pathlib import Path

cfg = configparser.ConfigParser(interpolation=None)
cfg.read(Path(__file__).parent / "smtp.ini", encoding="utf-8")
host, user = cfg.get("smtp", "host"), cfg.get("smtp", "user")
pw = (os.environ.get("BUH_SMTP_PASSWORD") or "").replace(" ", "")

print(f"Сервер: {host}, логин: {user}")
print(f"Пароль из BUH_SMTP_PASSWORD: длина {len(pw)}, "
      f"{'только латинские буквы' if pw.isascii() and pw.isalpha() else 'есть НЕ латинские буквы или другие символы'}")
if not pw:
    sys.exit("Переменная BUH_SMTP_PASSWORD не задана в этом окне терминала.")

for port, mode in ((465, "ssl"), (587, "starttls")):
    print(f"\n--- порт {port} ({mode}) ---")
    try:
        ctx = ssl.create_default_context()
        s = smtplib.SMTP_SSL(host, port, context=ctx, timeout=20) if mode == "ssl" else smtplib.SMTP(host, port, timeout=20)
        if mode == "starttls":
            s.ehlo()
            s.starttls(context=ctx)
        s.ehlo()
        issuer = dict(x[0] for x in s.sock.getpeercert()["issuer"])
        print("Сертификат выдан:", issuer.get("organizationName"), "/", issuer.get("commonName"))
        print("Способы входа:", s.esmtp_features.get("auth", "—"))
        s.login(user, pw)
        print("ВХОД УСПЕШЕН")
        s.quit()
    except smtplib.SMTPAuthenticationError as ex:
        print("Сервер отклонил логин/пароль:", ex.smtp_code, ex.smtp_error.decode(errors="replace")[:200])
    except Exception as ex:
        print("Ошибка:", type(ex).__name__, ex)
