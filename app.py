#!/usr/bin/env python3
"""Streamlit-интерфейс для поиска латиницы в Excel-файлах."""

from __future__ import annotations

import base64
import logging
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from main import scan_uploaded_files, setup_logging


def ensure_logging() -> None:
    if not logging.getLogger().handlers:
        setup_logging("parse.log")


class StreamlitLogHandler(logging.Handler):
    def __init__(self, sink_callback):
        super().__init__()
        self.sink_callback = sink_callback

    def emit(self, record: logging.LogRecord) -> None:
        self.sink_callback(self.format(record))


def trigger_auto_download(file_bytes: bytes, file_name: str, mime: str) -> None:
    b64 = base64.b64encode(file_bytes).decode("utf-8")
    href = f"data:{mime};base64,{b64}"
    components.html(
        f"""
        <a id='auto-download-link' href='{href}' download='{file_name}' style='display:none'>download</a>
        <script>document.getElementById('auto-download-link')?.click();</script>
        """,
        height=0,
    )


def app() -> None:
    st.set_page_config(page_title="Поиск латиницы в Excel", layout="centered")
    st.title("Поиск латиницы в 'Наименование товаров'")
    st.write("Загрузите Excel-файлы. Итоговый файл формируется автоматически с именем **Готовый с транскрипцией.xlsx**.")

    uploaded_files = st.file_uploader(
        "Выберите Excel-файлы с компьютера",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

    status_placeholder = st.empty()
    progress_placeholder = st.empty()
    logs_placeholder = st.empty()

    if st.button("Запустить анализ", type="primary"):
        ensure_logging()
        log_lines: list[str] = []
        progress_bar = progress_placeholder.progress(0.0, text="Прогресс обработки файлов: 0%")

        def push_log_line(line: str) -> None:
            log_lines.append(line)
            if len(log_lines) > 500:
                del log_lines[:100]
            logs_placeholder.text_area("Логи выполнения", value="\n".join(log_lines), height=300)

        def set_status(message: str) -> None:
            status_placeholder.info(f"Статус: {message}")

        def set_progress(current: int, total: int) -> None:
            total_safe = total if total > 0 else 1
            ratio = max(0.0, min(1.0, current / total_safe))
            progress_bar.progress(ratio, text=f"Прогресс обработки файлов: {int(ratio * 100)}%")

        root_logger = logging.getLogger()
        streamlit_handler = StreamlitLogHandler(push_log_line)
        streamlit_handler.setLevel(logging.INFO)
        streamlit_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s"))
        root_logger.addHandler(streamlit_handler)

        try:
            with st.spinner("Идёт обработка Excel-файлов..."):
                if not uploaded_files:
                    raise ValueError("Загрузите хотя бы один Excel-файл.")
                result = scan_uploaded_files(
                    uploaded_files=uploaded_files,
                    out_file="Готовый с транскрипцией.xlsx",
                    status_callback=set_status,
                    progress_callback=set_progress,
                )
        except Exception as exc:
            progress_bar.progress(1.0, text="Прогресс обработки файлов: 100%")
            status_placeholder.error("Статус: Ошибка")
            st.error(f"Ошибка: {exc}")
            result = None
        finally:
            root_logger.removeHandler(streamlit_handler)

        if result:
            progress_bar.progress(1.0, text="Прогресс обработки файлов: 100%")
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
                file_bytes = out_path.read_bytes()
                file_name = out_path.name
                trigger_auto_download(
                    file_bytes=file_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.info("Файл сформирован: загрузка началась автоматически. Если не сработало — нажмите кнопку ниже.")
                st.download_button(
                    "Скачать отчёт вручную",
                    data=file_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


if __name__ == "__main__":
    app()
