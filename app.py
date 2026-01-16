import datetime
import io
import logging
import re
from email import policy
from email.parser import BytesParser
from html import unescape
from typing import List, Optional, Set, Tuple

import imaplib
import pandas as pd
import streamlit as st
import xlwt


EXCLUDED_USERS = {
    "Авраменко Наталия",
    "Вифлянцев А.В.",
    "Воробьева",
    "Горностаева",
    "Гринчук Ольга",
    "Гулуева Татьяна",
    "Дегтярев Алексей",
    "Дегтярева О.А.",
    "Джиоева Ирина Витальевна",
    "Заподовникова И.",
    "Зеленская Галина",
    "Земцова",
    "Золотова Наталья",
    "Кирпичева",
    "Клишина Александра",
    "КонтроллерПеремещения1",
    "Коронова О.",
    "Куприянова О.В.",
    "МагазинПриемка3",
    "Майданик Ирина",
    "Пименова Вал.Ром.",
    "Скоробогатова Вера",
    "СтройградСклад1",
}

logging.basicConfig(
    filename="app.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

logger.info("Старт приложения")


def load_email_config() -> dict:
    """Загружает настройки почты из st.secrets."""
    try:
        email_config = st.secrets["email"]
    except Exception:
        logger.exception("Ошибка чтения st.secrets")
        raise

    logger.info("Настройки st.secrets успешно прочитаны")
    return email_config


def build_search_criteria(sender: str, start_date: datetime.date, end_date: datetime.date) -> List[str]:
    """Формирует критерии поиска IMAP по отправителю и диапазону дат."""
    since_str = start_date.strftime("%d-%b-%Y")
    before_date = end_date + datetime.timedelta(days=1)
    before_str = before_date.strftime("%d-%b-%Y")
    return ["FROM", f'"{sender}"', "SINCE", since_str, "BEFORE", before_str]


def extract_text_from_message(message) -> str:
    """Извлекает текст из email-сообщения, поддерживая text/plain и text/html."""
    parts_text = []

    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type in {"text/plain", "text/html"} and "attachment" not in disposition:
                try:
                    payload = part.get_content()
                except Exception:
                    payload = part.get_payload(decode=True)
                    if payload is None:
                        continue
                    payload = payload.decode(errors="ignore")

                if content_type == "text/html":
                    payload = unescape(re.sub(r"<[^>]+>", " ", payload))
                parts_text.append(payload)
    else:
        content_type = message.get_content_type()
        try:
            payload = message.get_content()
        except Exception:
            payload = message.get_payload(decode=True)
            if payload is None:
                payload = ""
            else:
                payload = payload.decode(errors="ignore")

        if content_type == "text/html":
            payload = unescape(re.sub(r"<[^>]+>", " ", payload))
        parts_text.append(payload)

    return "\n".join(parts_text)


def extract_user_from_text(text: str) -> Optional[str]:
    """Извлекает пользователя из текста письма по шаблону "Пользователь: <имя> провел"."""
    match = re.search(r"Пользователь:\s*(.*?)\s+провел", text)
    if match:
        return match.group(1).strip()
    return None


def extract_invoices_from_text(text: str) -> List[Tuple[str, str, str]]:
    """Ищет номера накладных и даты по шаблону "Приходная накл. <номер> (дата)"."""
    pattern = r"Приходная накл\.\s+([^\s(]+)\s*\(([^)]+)\)"
    user = extract_user_from_text(text) or "Неизвестно"
    matches = re.findall(pattern, text)
    return [(invoice_number, invoice_date, user) for invoice_number, invoice_date in matches]


def fetch_invoices(sender: str, start_date: datetime.date, end_date: datetime.date) -> List[Tuple[str, str, str]]:
    """Подключается к IMAP и извлекает номера накладных из писем."""
    email_config = load_email_config()
    host = email_config["IMAP_HOST"]
    port = email_config["IMAP_PORT"]
    user = email_config["EMAIL_USER"]
    password = email_config["EMAIL_PASSWORD"]

    logger.info("Попытка подключения к IMAP %s:%s", host, port)

    try:
        with imaplib.IMAP4_SSL(host, port) as imap:
            imap.login(user, password)
            logger.info("Успешная авторизация в IMAP")

            status, _ = imap.select("INBOX")
            if status != "OK":
                raise RuntimeError("Не удалось выбрать папку INBOX")

            criteria = build_search_criteria(sender, start_date, end_date)
            status, data = imap.search(None, *criteria)
            if status != "OK":
                raise RuntimeError("Ошибка IMAP-поиска")

            message_ids = data[0].split()
            logger.info("Найдено писем: %s", len(message_ids))

            invoices: List[Tuple[str, str, str]] = []
            for msg_id in message_ids:
                status, msg_data = imap.fetch(msg_id, "(RFC822)")
                if status != "OK" or not msg_data:
                    logger.warning("Не удалось получить письмо %s", msg_id)
                    continue

                raw_email = msg_data[0][1]
                message = BytesParser(policy=policy.default).parsebytes(raw_email)
                text = extract_text_from_message(message)
                extracted = extract_invoices_from_text(text)
                logger.info("Письмо %s: найдено накладных %s", msg_id.decode(), len(extracted))
                invoices.extend(extracted)

            return invoices
    except imaplib.IMAP4.error as exc:
        logger.exception("Ошибка IMAP-аутентификации или доступа")
        raise RuntimeError("Ошибка IMAP-аутентификации. Проверьте логин и пароль.") from exc
    except Exception as exc:
        logger.exception("Ошибка работы с IMAP")
        raise RuntimeError("Ошибка подключения к IMAP. Проверьте настройки сервера и лог.") from exc


def parse_invoice_date(raw_date: str) -> Optional[datetime.date]:
    """Пытается разобрать дату накладной из строки формата dd.mm.yy или dd.mm.yyyy."""
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(raw_date.strip(), fmt).date()
        except ValueError:
            continue
    return None


def build_report(invoices: List[Tuple[str, str, str]]) -> pd.DataFrame:
    """Формирует DataFrame с уникальными накладными, датами и пользователями, сортирует по дате."""
    unique_invoices: List[Tuple[str, str, str]] = sorted(set(invoices))
    filtered_invoices = [
        (invoice_number, raw_date, user)
        for invoice_number, raw_date, user in unique_invoices
        if user not in EXCLUDED_USERS
    ]
    logger.info("Уникальных накладных: %s", len(unique_invoices))
    logger.info("Накладных после фильтрации по пользователям: %s", len(filtered_invoices))

    rows = []
    for invoice_number, raw_date, user in filtered_invoices:
        parsed_date = parse_invoice_date(raw_date)
        rows.append(
            {
                "Накладная": invoice_number,
                "Дата": raw_date,
                "Пользователь": user,
                "_sort_date": parsed_date or datetime.date.min,
            }
        )

    df = pd.DataFrame(rows)
    df = df.sort_values(by="_sort_date", ascending=True).drop(columns=["_sort_date"])
    return df


def dataframe_to_xls(df: pd.DataFrame) -> io.BytesIO:
    """Сохраняет DataFrame в XLS через xlwt и возвращает BytesIO."""
    output = io.BytesIO()
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Отчет")

    for row_index, row in enumerate(df.itertuples(index=False), start=0):
        for col_index, value in enumerate(row):
            sheet.write(row_index, col_index, value)

    workbook.save(output)
    output.seek(0)
    return output


def main() -> None:
    """Основная функция Streamlit-приложения."""
    st.title("Поиск накладных по IMAP")

    sender = "robot_volgorost@volgorost.ru"

    with st.expander("Логи (последние 200 строк)", expanded=False):
        try:
            with open("app.log", "r", encoding="utf-8") as log_file:
                log_lines = log_file.readlines()[-200:]
            if log_lines:
                st.text("".join(log_lines))
            else:
                st.info("Логи пока пусты.")
        except FileNotFoundError:
            st.info("Файл логов ещё не создан.")
        except Exception as exc:
            logger.exception("Ошибка чтения логов в UI")
            st.error(f"Не удалось прочитать лог: {exc}")

    with st.form("search_form"):
        start_date = st.date_input(
            "Дата начала периода",
            value=datetime.date.today(),
            format="DD.MM.YYYY",
        )
        end_date = st.date_input(
            "Дата окончания периода",
            value=datetime.date.today(),
            format="DD.MM.YYYY",
        )
        submitted = st.form_submit_button("Запустить поиск")

    if submitted:
        progress = st.progress(0, text="🐱 Подключаюсь к IMAP...")
        cat_placeholder = st.empty()
        cat_placeholder.markdown(
            "```\n"
            " /\\_/\\\n"
            "( o.o )\n"
            " > ^ <\n"
            "```\n"
        )

        if start_date > end_date:
            st.error("Дата начала не может быть позже даты окончания.")
            logger.error("Некорректный диапазон дат: %s - %s", start_date, end_date)
            progress.empty()
            cat_placeholder.empty()
            return

        try:
            invoices = fetch_invoices(sender, start_date, end_date)
        except KeyError:
            st.error("Не найдены настройки email в st.secrets. Проверьте secrets.toml.")
            progress.empty()
            cat_placeholder.empty()
            return
        except RuntimeError as exc:
            st.error(str(exc))
            progress.empty()
            cat_placeholder.empty()
            return

        if not invoices:
            st.warning("За выбранный период накладные не найдены")
            logger.info("Накладные за период не найдены")
            progress.empty()
            cat_placeholder.empty()
            return

        progress.progress(60, text="🐱 Готовлю отчет...")
        df = build_report(invoices)

        select_all = st.checkbox("Выделить все / снять все", value=True, key="select_all")
        df_for_editor = df.copy()
        df_for_editor.insert(0, "Выбрать", select_all)
        edited_df = st.data_editor(
            df_for_editor,
            hide_index=True,
            column_config={"Выбрать": st.column_config.CheckboxColumn(required=True)},
            key="invoice_selector",
        )

        selected_df = edited_df[edited_df["Выбрать"]].drop(columns=["Выбрать"])
        if selected_df.empty:
            st.warning("Нет выбранных накладных для выгрузки.")

        file_name = f"nakladnye_{start_date:%d.%m.%Y}-{end_date:%d.%m.%Y}.xls"
        xls_data = dataframe_to_xls(selected_df[["Дата"]])
        progress.progress(100, text="🐱 Отчет готов!")
        st.download_button(
            label="Скачать XLS",
            data=xls_data,
            file_name=file_name,
            mime="application/vnd.ms-excel",
            disabled=selected_df.empty,
        )

        progress.empty()
        cat_placeholder.empty()


if __name__ == "__main__":
    main()
