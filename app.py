#!/usr/bin/env python3
"""Streamlit-интерфейс для парсера volgarost.ru."""

from __future__ import annotations

import logging
from pathlib import Path

import streamlit as st

from main import run_parsing, setup_logging


def ensure_logging() -> None:
    """Инициализирует логирование один раз для Streamlit-процесса."""
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

    if st.button("Запустить парсинг", type="primary"):
        ensure_logging()
        with st.spinner("Идёт парсинг, это может занять несколько минут..."):
            result = run_parsing(
                start_url=start_url.strip(),
                category_url=category_url.strip() or None,
                out_file=out_name.strip() or "voligorost_latin_names.xlsx",
                max_pages_per_category=int(max_pages) or None,
            )

        st.success("Готово")
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
