#!/usr/bin/env python3
"""Streamlit-интерфейс для поиска латиницы в Excel-файлах папки."""

from __future__ import annotations

import logging
from pathlib import Path

import streamlit as st

from main import scan_folder, setup_logging


class StreamlitLogHandler(logging.Handler):
    def __init__(self, sink_callback):
        super().__init__()
        self.sink_callback = sink_callback

    def emit(self, record: logging.LogRecord) -> None:
        self.sink_callback(self.format(record))


def ensure_logging() -> None:
    if not logging.getLogger().handlers:
        setup_logging("parse.log")


def app() -> None:
    st.set_page_config(page_title="Поиск латиницы в Excel", layout="centered")
    st.title("Поиск латиницы в 'Наименование товаров'")
    st.write(
        "Укажите папку с Excel-файлами. На выходе будет .xls-отчёт с колонками: "
        "Код, Наименование товара, Транскрипция."
    )

    folder_path = st.text_input("Путь к папке с Excel", value="")
    out_name = st.text_input("Имя итогового файла (.xls)", value="latin_names_report.xls")

    status_placeholder = st.empty()
    logs_placeholder = st.empty()

    if st.button("Запустить анализ", type="primary"):
        ensure_logging()
        log_lines: list[str] = []

        def push_log_line(line: str) -> None:
            log_lines.append(line)
            if len(log_lines) > 500:
                del log_lines[:100]
            logs_placeholder.text_area("Логи выполнения", value="\n".join(log_lines), height=300)

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
            with st.spinner("Идёт обработка Excel-файлов..."):
                result = scan_folder(
                    folder_path=folder_path.strip(),
                    out_file=out_name.strip() or "latin_names_report.xls",
                    status_callback=set_status,
                )
        except Exception as exc:
            status_placeholder.error("Статус: Ошибка")
            st.error(f"Ошибка: {exc}")
            result = None
        finally:
            root_logger.removeHandler(streamlit_handler)

        if result:
            status_placeholder.success("Статус: Готово")
            st.success("Анализ завершён")
            st.write(
                {
                    "Файлов найдено": result["files"],
                    "Файлов с нужными колонками": result["files_with_columns"],
                    "Строк с латиницей": result["written"],
                    "Ошибок": len(result["errors"]),
                }
            )

            if result["errors"]:
                with st.expander("Показать ошибки"):
                    for url, err in result["errors"][:100]:
                        st.write(f"- {url}: {err}")

            out_path = Path(result["output"])
            if out_path.exists():
                st.download_button(
                    "Скачать итоговый XLS",
                    data=out_path.read_bytes(),
                    file_name=out_path.name,
                    mime="application/vnd.ms-excel",
                )


if __name__ == "__main__":
    app()
