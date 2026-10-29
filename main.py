"""Рассылка расчётных листков сотрудникам по email.

По умолчанию — пробный запуск: листки и письма (.eml) сохраняются в папку, но ничего не отправляется.

  python main.py "ведомость ЗП окт.xlsx"                          пробный запуск
  python main.py "ведомость ЗП окт.xlsx" --test-to я@почта.kg     все письма — себе
  python main.py "ведомость ЗП окт.xlsx" --send                   рассылка сотрудникам
  python main.py "ведомость ЗП окт.xlsx" --send --only Иванов     только одному

Для простой таблицы (CSV) месяц указывается вручную: --period 2026-10
"""
import argparse
import re
import smtplib
import sys
import time
from pathlib import Path

import pochta
import zhurnal
from listok import make_listok
from sotrudniki import add_emails
from utils import money, month_title
from vedomost import read_vedomost

HERE = Path(__file__).parent


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("vedomost", help="файл ведомости: .xlsx из 1С, .xlsx или .csv")
    parser.add_argument("--period", help="месяц в виде ГГГГ-ММ; для ведомости 1С берётся из её заголовка")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--send", action="store_true", help="отправить письма сотрудникам")
    mode.add_argument("--test-to", metavar="EMAIL", help="отправить все письма на этот адрес (проверка)")
    parser.add_argument("--only", help="только эти сотрудники: табельные номера или фамилии через запятую")
    parser.add_argument("--resend", action="store_true", help="отправить ещё раз тем, кому уже отправляли")
    parser.add_argument("--yes", action="store_true", help="не спрашивать подтверждение перед --send")
    parser.add_argument("--emails", help="справочник email (по умолчанию sotrudniki.xlsx)")
    parser.add_argument("--config", default=HERE / "smtp.ini", help="настройки почты (по умолчанию smtp.ini)")
    parser.add_argument("--out", help="куда сохранять листки (по умолчанию raschetnye_listki/ГГГГ-ММ)")
    return parser.parse_args()


def main():
    args = parse_args()
    sending = args.send or bool(args.test_to)

    people, problems, info = read_vedomost(args.vedomost)
    if info:  # ведомость из 1С — почты в ней нет, берём из справочника
        add_emails(people, problems, args.emails)

    period = args.period or info.get("period")
    if not period or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", period):
        raise SystemExit("Укажите месяц в виде ГГГГ-ММ, например --period 2026-10")

    settings = pochta.load_settings(args.config, required=sending)
    org = info.get("org") or settings.get("letter", "organization", fallback="") or "Организация"

    if args.only:
        wanted = [w.strip() for w in args.only.split(",")]
        people = [p for p in people if p.key in wanted or p.fio.split()[0] in wanted]
        if not people:
            problems.append(f"в ведомости нет: {args.only}")

    if problems:
        print("Замечания по ведомости:")
        for text in problems:
            print("  -", text)
        if sending:
            raise SystemExit("Сначала исправьте замечания, потом отправляйте.")

    folder = Path(args.out or HERE / "raschetnye_listki" / period)
    folder.mkdir(parents=True, exist_ok=True)
    journal = zhurnal.open_journal()

    # Делаем листки всем и решаем, кому отправлять
    to_send = []
    print(f"\nРасчётные листки за {month_title(period)}: {len(people)} шт. -> {folder}")
    for p in people:
        file = make_listok(p, period, org, folder)
        status = p.skip
        if not status and args.send and not args.resend:
            before = zhurnal.when_sent(journal, period, p.key)
            if before:
                status = f"уже отправлен {before['sent_at']} на {before['email']}"
        print(f"  {p.tab:<6} {p.fio[:32]:<32} {money(p.k_vyplate):>14}  {p.email or '—':<28} {status}")
        if not status:
            to_send.append((p, file))

    total = sum(p.k_vyplate for p in people)
    print(f"Итого к выплате: {money(total)}. Писем к отправке: {len(to_send)}.")

    if not sending:
        for p, file in to_send:
            letter = pochta.make_letter(p, file, period, org, settings, p.email)
            (folder / f"{file.stem}.eml").write_bytes(bytes(letter))
        print("\nПРОБНЫЙ ЗАПУСК: ничего не отправлено. Письма сохранены как .eml рядом с листками,")
        print("их можно открыть в Outlook и проверить. Дальше: --test-to свой@адрес, потом --send.")
        return

    if not to_send:
        print("Отправлять нечего.")
        return

    if args.send and not args.yes:
        answer = input(f"\nОтправить {len(to_send)} писем сотрудникам? Введите «да»: ")
        if answer.strip().lower() not in ("да", "д", "yes", "y"):
            print("Отменено.")
            return

    send_all(to_send, args, period, org, settings, journal)


def send_all(to_send, args, period, org, settings, journal):
    pause = settings.getfloat("smtp", "delay_seconds", fallback=1.0)  # почтовики не любят частые письма
    sent, errors = 0, 0
    server = pochta.connect(settings)
    try:
        for number, (p, file) in enumerate(to_send, start=1):
            address = args.test_to or p.email
            letter = pochta.make_letter(p, file, period, org, settings, address)
            try:
                server.send_message(letter)
            except smtplib.SMTPServerDisconnected:
                # сервер мог закрыть соединение по таймауту — переподключаемся один раз
                server = pochta.connect(settings)
                server.send_message(letter)
            except smtplib.SMTPException as error:
                errors += 1
                print(f"  ОШИБКА: {p.fio} -> {address}: {error}")
                continue

            sent += 1
            print(f"  отправлено: {p.fio} -> {address}")
            if args.send:  # тестовую отправку себе в журнал не пишем
                zhurnal.remember(journal, period, p, address, file)
            if number < len(to_send):
                time.sleep(pause)
    finally:
        try:
            server.quit()
        except Exception:
            pass

    print(f"\nГотово: отправлено {sent}, ошибок {errors}.")
    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
