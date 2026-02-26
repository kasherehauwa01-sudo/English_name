#!/usr/bin/env python3
"""Streamlit-интерфейс для парсера volgarost.ru."""

from __future__ import annotations

import logging
from pathlib import Path

import streamlit as st

from main import run_parsing, setup_logging


class StreamlitLogHandler(logging.Handler):
    """Лог-хендлер, который выводит сообщения в Streamlit в реальном времени."""

    def __init__(self, sink_callback):
        super().__init__()
        self.sink_callback = sink_callback

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        self.sink_callback(message)


def ensure_logging() -> None:
    """Инициализирует базовое логирование один раз для Streamlit-процесса."""
    if not logging.getLogger().handlers:
        setup_logging("parse.log")


def app() -> None:
    st.set_page_config(page_title="Парсер volgorost", layout="centered")
    st.title("Парсер товаров volgorost.ru")
    st.write(
        "Найдёт товары с латиницей в названии, выполнит транскрипцию и сформирует Excel-отчёт."
    )

    start_url = st.text_input("Стартовый URL", value="https://volgorost.ru/")
    category_url = st.text_input(
        "URL одного раздела (необязательно)",
        value="",
        help="Если заполнено, будет обработан только этот раздел.",
    )
    out_name = st.text_input("Имя файла отчёта", value="voligorost_latin_names.xlsx")
    max_pages = st.number_input(
        "Лимит страниц на раздел (0 = без лимита)",
        min_value=0,
        value=0,
        step=1,
    )

    status_placeholder = st.empty()
    logs_placeholder = st.empty()

    if st.button("Запустить парсинг", type="primary"):
        ensure_logging()

        log_lines: list[str] = []

        def push_log_line(line: str) -> None:
            log_lines.append(line)
            # Ограничиваем буфер, чтобы интерфейс не тормозил на очень длинном запуске.
            if len(log_lines) > 500:
                del log_lines[:100]
            logs_placeholder.text_area("Логи выполнения", value="\n".join(log_lines), height=280)

        def set_status(message: str) -> None:
            status_placeholder.info(f"Статус: {message}")

        root_logger = logging.getLogger()
        streamlit_handler = StreamlitLogHandler(push_log_line)
        streamlit_handler.setLevel(logging.INFO)
        streamlit_handler.setFormatter(
            logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
        )
        root_logger.addHandler(streamlit_handler)

        try:
            with st.spinner("Идёт парсинг, это может занять несколько минут..."):
                result = run_parsing(
                    start_url=start_url.strip(),
                    category_url=category_url.strip() or None,
                    out_file=out_name.strip() or "voligorost_latin_names.xlsx",
                    max_pages_per_category=int(max_pages) or None,
                    status_callback=set_status,
                )
        finally:
            root_logger.removeHandler(streamlit_handler)

        status_placeholder.success("Статус: Готово")
        st.success("Парсинг завершён")
        st.write(
            {
                "Найдено категорий": result["categories"],
                "Обработано карточек": result["processed"],
                "Попало в отчёт": result["written"],
                "Ошибок": len(result["errors"]),
            }
        )

        if result["errors"]:
            with st.expander("Показать ошибки"):
                for url, err in result["errors"][:100]:
                    st.write(f"- {url}: {err}")

        out_path = Path(result["output"])
        if out_path.exists():
            data = out_path.read_bytes()
            st.download_button(
                "Скачать Excel",
                data=data,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    app()
