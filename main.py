#!/usr/bin/env python3
"""
Утилита для поиска строк с латиницей в Excel-файлах папки.

Новая целевая логика:
- искать наименования строго в колонке "Наименование товаров\n" (поддержаны мягкие вариации через нормализацию),
- брать код из колонки "Код товара",
- отбирать только строки, где в наименовании есть латиница [A-Za-z],
- формировать итоговый отчёт в формате .xls с колонками:
  1) Код
  2) Наименование товара
  3) Транскрипция
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

LATIN_RE = re.compile(r"[A-Za-z]")
LATIN_TOKEN_RE = re.compile(r"[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*")
VOWELS = set("aeiouy")

EXCEPTION_MAP = {
    "usb": "усб",
    "led": "лед",
    "wifi": "вайфай",
    "wi-fi": "вай-фай",
    "bluetooth": "блютус",
    "smart": "смарт",
    "pro": "про",
    "max": "макс",
    "mini": "мини",
    "ultra": "ультра",
    "eco": "эко",
    "iphone": "айфон",
    "samsung": "самсунг",
    "tefal": "тефаль",
    "type-c": "тайп-си",
    "usb-c": "усб-си",
}


@dataclass
class MatchRow:
    code: str
    name: str
    transcription: str


def setup_logging(log_file: str = "parse.log") -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, mode="w", encoding="utf-8"),
        ],
    )


def normalize_header(value: object) -> str:
    """Нормализует заголовок колонки (регистр/пробелы/переводы строк)."""
    text = str(value or "")
    text = text.replace("\r", " ").replace("\n", " ")
    text = text.strip().lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text


def find_required_columns(columns: list[object]) -> tuple[Optional[object], Optional[object]]:
    """
    Ищет 2 обязательные колонки:
    - код товара,
    - наименование товаров.
    """
    code_col = None
    name_col = None

    for col in columns:
        n = normalize_header(col)
        if code_col is None and n in {"код товара", "код", "артикул"}:
            code_col = col
        if name_col is None and n in {"наименование товаров", "наименование товара", "наименование"}:
            name_col = col

    # Фоллбэк на частичное совпадение
    if code_col is None:
        for col in columns:
            n = normalize_header(col)
            if "код" in n and "товар" in n:
                code_col = col
                break
    if name_col is None:
        for col in columns:
            n = normalize_header(col)
            if "наимен" in n and "товар" in n:
                name_col = col
                break

    return code_col, name_col


def letter_translit(chunk: str) -> str:
    s = chunk.lower()

    for old, new in [
        ("sch", "щ"),
        ("sh", "ш"),
        ("ch", "ч"),
        ("zh", "ж"),
        ("kh", "х"),
        ("ph", "ф"),
        ("th", "т"),
        ("ck", "к"),
        ("qu", "кв"),
    ]:
        s = s.replace(old, new)

    out: list[str] = []
    for i, c in enumerate(s):
        nxt = s[i + 1] if i + 1 < len(s) else ""
        prv = s[i - 1] if i > 0 else ""

        if c.isdigit():
            out.append(c)
        elif c == "a":
            out.append("а")
        elif c == "b":
            out.append("б")
        elif c == "c":
            out.append("с" if nxt in "eiy" else "к")
        elif c == "d":
            out.append("д")
        elif c == "e":
            out.append("е")
        elif c == "f":
            out.append("ф")
        elif c == "g":
            out.append("дж" if nxt in "eiy" else "г")
        elif c == "h":
            out.append("х")
        elif c == "i":
            out.append("и")
        elif c == "j":
            out.append("дж")
        elif c == "k":
            out.append("к")
        elif c == "l":
            out.append("л")
        elif c == "m":
            out.append("м")
        elif c == "n":
            out.append("н")
        elif c == "o":
            out.append("о")
        elif c == "p":
            out.append("п")
        elif c == "q":
            out.append("кв")
        elif c == "r":
            out.append("р")
        elif c == "s":
            out.append("з" if prv in VOWELS and nxt in VOWELS else "с")
        elif c == "t":
            out.append("т")
        elif c == "u":
            out.append("у")
        elif c == "v":
            out.append("в")
        elif c == "w":
            out.append("в")
        elif c == "x":
            out.append("кс")
        elif c == "y":
            out.append("й" if i == 0 else "и")
        elif c == "z":
            out.append("з")
        elif c == "-":
            out.append("-")
        else:
            out.append(c)

    return "".join(out)


def translit_token(token: str) -> str:
    low = token.lower()
    if low in EXCEPTION_MAP:
        return EXCEPTION_MAP[low]

    parts = token.split("-")
    if len(parts) > 1:
        return "-".join(translit_token(part) for part in parts)

    mixed = re.findall(r"[A-Za-z]+|\d+", token)
    if len(mixed) > 1:
        return "".join(p if p.isdigit() else translit_token(p) for p in mixed)

    return letter_translit(token)


def translit_to_ru(text: str) -> str:
    return LATIN_TOKEN_RE.sub(lambda m: translit_token(m.group(0)), text)


def write_xls(rows: list[MatchRow], out_file: str) -> int:
    """Записывает отчёт именно в .xls (по требованию)."""
    try:
        import xlwt
    except Exception as exc:
        raise RuntimeError("Для записи .xls установите зависимость xlwt (pip install xlwt)") from exc

    book = xlwt.Workbook()
    sheet = book.add_sheet("latin")

    header_style = xlwt.easyxf("font: bold on")
    headers = ["Код", "Наименование товара", "Транскрипция"]
    for i, h in enumerate(headers):
        sheet.write(0, i, h, header_style)

    max_lens = [len(h) for h in headers]
    for row_idx, row in enumerate(rows, start=1):
        values = [row.code, row.name, row.transcription]
        for col_idx, value in enumerate(values):
            sheet.write(row_idx, col_idx, value)
            max_lens[col_idx] = max(max_lens[col_idx], len(str(value)))

    for i, ln in enumerate(max_lens):
        sheet.col(i).width = min((ln + 2) * 256, 256 * 120)

    if not out_file.lower().endswith(".xls"):
        out_file = f"{out_file}.xls"
    book.save(out_file)
    return len(rows)


def scan_folder(folder_path: str, out_file: str, status_callback=None) -> dict:
    logger = logging.getLogger("main")

    def set_status(message: str) -> None:
        logger.info(message)
        if status_callback:
            status_callback(message)

    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        msg = f"Папка не найдена: {folder_path}"
        logger.error(msg)
        raise FileNotFoundError(msg)

    files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"}])
    set_status(f"Найдено Excel-файлов: {len(files)}")

    all_matches: list[MatchRow] = []
    errors: list[tuple[str, str]] = []
    files_with_required_columns = 0

    for idx, file_path in enumerate(files, start=1):
        set_status(f"Обрабатываю файл {idx}/{len(files)}: {file_path.name}")
        try:
            xls = pd.ExcelFile(file_path)
            matched_in_file = 0
            has_required = False

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if df.empty:
                    continue

                code_col, name_col = find_required_columns(list(df.columns))
                if not code_col or not name_col:
                    continue

                has_required = True
                names = df[name_col].fillna("").astype(str)
                codes = df[code_col].fillna("").astype(str)

                for code_value, name_value in zip(codes, names):
                    name_text = name_value.strip()
                    if name_text and LATIN_RE.search(name_text):
                        all_matches.append(
                            MatchRow(
                                code=str(code_value).strip(),
                                name=name_text,
                                transcription=translit_to_ru(name_text),
                            )
                        )
                        matched_in_file += 1

            if has_required:
                files_with_required_columns += 1
            logger.info("%s: найдено строк с латиницей: %s", file_path.name, matched_in_file)
        except Exception as exc:
            errors.append((str(file_path), str(exc)))
            logger.error("Ошибка чтения %s: %s", file_path, exc)

    set_status("Формирую итоговый XLS...")
    written = write_xls(all_matches, out_file)
    final_output = out_file if out_file.lower().endswith(".xls") else f"{out_file}.xls"

    set_status("Готово")
    logger.info("=" * 50)
    logger.info("SUMMARY")
    logger.info("Файлов найдено: %s", len(files))
    logger.info("Файлов с нужными колонками: %s", files_with_required_columns)
    logger.info("Строк с латиницей: %s", written)
    logger.info("Ошибок: %s", len(errors))

    return {
        "files": len(files),
        "files_with_columns": files_with_required_columns,
        "written": written,
        "errors": errors,
        "output": final_output,
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Поиск латиницы в колонке 'Наименование товаров' и экспорт в .xls"
    )
    p.add_argument("--folder", required=True, help="Путь к папке с Excel-файлами")
    p.add_argument("--out", default="latin_names_report.xls", help="Путь к итоговому .xls")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    setup_logging()
    scan_folder(folder_path=args.folder, out_file=args.out)


if __name__ == "__main__":
    main()
