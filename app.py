import datetime
import io
import logging
import re
from email import policy
from email.parser import BytesParser
from html import unescape
from typing import List, Set

import imaplib
import pandas as pd
import streamlit as st


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


def extract_invoices_from_text(text: str) -> List[str]:
    """Ищет номера накладных по шаблону "Приходная накл. <номер> (дата)"."""
    pattern = r"Приходная накл\.\s+([^\s(]+)"
    return re.findall(pattern, text)


def fetch_invoices(sender: str, start_date: datetime.date, end_date: datetime.date) -> List[str]:
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

            invoices: List[str] = []
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
    except Exception:
        logger.exception("Ошибка работы с IMAP")
        raise


def build_report(invoices: List[str]) -> pd.DataFrame:
    """Формирует DataFrame с уникальными накладными."""
    unique_invoices: Set[str] = sorted(set(invoices))
    logger.info("Уникальных накладных: %s", len(unique_invoices))
    return pd.DataFrame({"Накладная": list(unique_invoices)})


def dataframe_to_xls(df: pd.DataFrame) -> io.BytesIO:
    """Сохраняет DataFrame в XLS (xlwt) и возвращает BytesIO."""
    output = io.BytesIO()
    df.to_excel(output, index=False, engine="xlwt")
    output.seek(0)
    return output


def main() -> None:
    """Основная функция Streamlit-приложения."""
    st.title("Поиск накладных по IMAP")

    with st.form("search_form"):
        sender = st.text_input(
            "Email отправителя",
            value="robot_volgorost@volgorost.ru",
        )
        start_date = st.date_input("Дата начала периода", value=datetime.date.today())
        end_date = st.date_input("Дата окончания периода", value=datetime.date.today())
        submitted = st.form_submit_button("Запустить поиск")

    if submitted:
        if start_date > end_date:
            st.error("Дата начала не может быть позже даты окончания.")
            logger.error("Некорректный диапазон дат: %s - %s", start_date, end_date)
            return

        try:
            invoices = fetch_invoices(sender, start_date, end_date)
        except KeyError:
            st.error("Не найдены настройки email в st.secrets. Проверьте secrets.toml.")
            return
        except Exception:
            st.error("Ошибка подключения к почте или обработки писем. Проверьте лог.")
            return

        if not invoices:
            st.warning("За выбранный период накладные не найдены")
            logger.info("Накладные за период не найдены")
            return

        df = build_report(invoices)
        st.dataframe(df)

        file_name = f"nakladnye_{start_date:%Y%m%d}-{end_date:%Y%m%d}.xls"
        xls_data = dataframe_to_xls(df)
        st.download_button(
            label="Скачать XLS",
            data=xls_data,
            file_name=file_name,
            mime="application/vnd.ms-excel",
        )


if __name__ == "__main__":
    main()
