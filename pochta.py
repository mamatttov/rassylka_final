"""Письма и подключение к почтовому серверу.

Настройки — в smtp.ini (образец: smtp.example.ini).
Пароль в файлах не храним: он берётся из переменной окружения BUH_SMTP_PASSWORD
или спрашивается при запуске.
"""
import configparser
import getpass
import os
import re
import smtplib
import ssl
from email.headerregistry import Address
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

from utils import money, month_title

XLSX_TYPE = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def load_settings(path, required):
    settings = configparser.ConfigParser(interpolation=None)  # без interpolation, иначе {period} и % ломаются
    if Path(path).exists():
        settings.read(path, encoding="utf-8")
    elif required:
        raise SystemExit(f"Нет файла настроек {path}. Скопируйте smtp.example.ini в smtp.ini и заполните.")
    return settings


def make_letter(p, attachment, period, org, settings, to_addr):
    values = {
        "fio": p.fio,
        "period": month_title(period),
        "org": org,
        "total": money(p.k_vyplate),
    }
    subject = settings.get("letter", "subject", fallback="Расчётный листок за {period}")
    body = settings.get("letter", "body", fallback="Во вложении ваш расчётный листок за {period}.")
    from_addr = settings.get("smtp", "from_addr", fallback="") or settings.get("smtp", "user", fallback="buh@localhost")
    from_name = settings.get("smtp", "from_name", fallback="")
    reply_to = settings.get("smtp", "reply_to", fallback="")

    letter = EmailMessage()
    letter["Subject"] = subject.format(**values)
    letter["From"] = str(Address(from_name, addr_spec=from_addr))
    letter["To"] = to_addr
    if reply_to:
        letter["Reply-To"] = reply_to
    letter["Date"] = formatdate(localtime=True)
    letter["Message-ID"] = make_msgid(domain=from_addr.split("@")[-1])
    letter.set_content(body.format(**values))
    letter.add_attachment(attachment.read_bytes(), maintype=XLSX_TYPE[0], subtype=XLSX_TYPE[1],
                          filename=attachment.name)
    return letter


def get_password(user):
    password = os.environ.get("BUH_SMTP_PASSWORD") or getpass.getpass(f"Пароль SMTP для {user}: ")
    password = password.replace(" ", "")  # Google показывает пароль приложения группами через пробел

    # В терминале Windows Ctrl+V при вводе скрытого пароля вставляет служебный символ, а не текст
    if not password or re.search(r"[\x00-\x1f]", password):
        raise SystemExit("Пароль пустой или вставился с лишними символами.\n"
                         'Задайте его так:  $env:BUH_SMTP_PASSWORD="пароль"  и запустите снова.')
    return password


def connect(settings):
    host = settings.get("smtp", "host")
    port = settings.getint("smtp", "port", fallback=465)
    security = settings.get("smtp", "security", fallback="ssl").lower()
    user = settings.get("smtp", "user", fallback="")

    context = ssl.create_default_context()
    if security == "ssl":
        server = smtplib.SMTP_SSL(host, port, context=context, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        if security == "starttls":
            server.starttls(context=context)

    if user:
        try:
            server.login(user, get_password(user))
        except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected) as error:
            raise SystemExit(f"Почтовый сервер не принял логин или пароль ({error}).\n"
                             "Для Gmail нужен пароль приложения из 16 букв, а не обычный пароль от почты.")
    return server
