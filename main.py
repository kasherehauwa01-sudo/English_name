#!/usr/bin/env python3
"""Поиск строк с латиницей в Excel и формирование .xls-отчёта."""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from io import BytesIO
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
    text = str(value or "")
    text = text.replace("\r", " ").replace("\n", " ")
    text = text.strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def find_required_columns(columns: list[object]) -> tuple[Optional[object], Optional[object]]:
    code_col = None
    name_col = None

    for col in columns:
        n = normalize_header(col)
        if code_col is None and n in {"код товара", "код", "артикул"}:
            code_col = col
        if name_col is None and n in {"наименование товаров", "наименование товара", "наименование"}:
            name_col = col

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
        ("sch", "щ"), ("sh", "ш"), ("ch", "ч"), ("zh", "ж"), ("kh", "х"),
        ("ph", "ф"), ("th", "т"), ("ck", "к"), ("qu", "кв"),
    ]:
        s = s.replace(old, new)

    out: list[str] = []
    for i, c in enumerate(s):
        nxt = s[i + 1] if i + 1 < len(s) else ""
        prv = s[i - 1] if i > 0 else ""
        mapping = {
            "a": "а", "b": "б", "d": "д", "e": "е", "f": "ф", "h": "х", "i": "и",
            "j": "дж", "k": "к", "l": "л", "m": "м", "n": "н", "o": "о", "p": "п",
            "q": "кв", "r": "р", "t": "т", "u": "у", "v": "в", "w": "в", "x": "кс", "z": "з",
        }
        if c.isdigit():
            out.append(c)
        elif c == "c":
            out.append("с" if nxt in "eiy" else "к")
        elif c == "g":
            out.append("дж" if nxt in "eiy" else "г")
        elif c == "s":
            out.append("з" if prv in VOWELS and nxt in VOWELS else "с")
        elif c == "y":
            out.append("й" if i == 0 else "и")
        elif c == "-":
            out.append("-")
        else:
            out.append(mapping.get(c, c))
    return "".join(out)


def translit_token(token: str) -> str:
    low = token.lower()
    if low in EXCEPTION_MAP:
        base = EXCEPTION_MAP[low]
    else:
        parts = token.split("-")
        if len(parts) > 1:
            base = "-".join(translit_token(p) for p in parts)
        else:
            mixed = re.findall(r"[A-Za-z]+|\d+", token)
            if len(mixed) > 1:
                base = "".join(p if p.isdigit() else translit_token(p) for p in mixed)
            else:
                base = letter_translit(token)

    if token.isupper():
        return base.upper()
    return base


def translit_to_ru(text: str) -> str:
    return LATIN_TOKEN_RE.sub(lambda m: translit_token(m.group(0)), text)


def write_xls(rows: list[MatchRow], out_file: str) -> str:
    """Пишет отчёт в .xls. Если xlwt недоступен — использует HTML-таблицу с расширением .xls."""
    final = out_file if out_file.lower().endswith(".xls") else f"{out_file}.xls"

    try:
        import xlwt

        book = xlwt.Workbook()
        sheet = book.add_sheet("latin")
        headers = ["Код", "Наименование товара", "Транскрипция"]
        hstyle = xlwt.easyxf("font: bold on")
        for i, h in enumerate(headers):
            sheet.write(0, i, h, hstyle)

        widths = [len(h) for h in headers]
        for r, row in enumerate(rows, start=1):
            vals = [row.code, row.name, row.transcription]
            for c, v in enumerate(vals):
                sheet.write(r, c, v)
                widths[c] = max(widths[c], len(str(v)))
        for i, w in enumerate(widths):
            sheet.col(i).width = min((w + 2) * 256, 256 * 120)

        book.save(final)
        return final
    except Exception:
        # Fallback без внешней зависимости: Excel корректно открывает HTML-таблицу с расширением .xls.
        df = pd.DataFrame(
            [{"Код": r.code, "Наименование товара": r.name, "Транскрипция": r.transcription} for r in rows]
        )
        html = df.to_html(index=False)
        Path(final).write_text(html, encoding="utf-8")
        return final


def _rows_with_merged_cells_xlsx(ws) -> set[int]:
    rows: set[int] = set()
    for mr in ws.merged_cells.ranges:
        rows.update(range(mr.min_row, mr.max_row + 1))
    return rows


def _rows_with_merged_cells_xls(sheet) -> set[int]:
    rows: set[int] = set()
    for rlow, rhigh, _clow, _chigh in getattr(sheet, "merged_cells", []):
        for r in range(rlow, rhigh):
            rows.add(r + 1)  # в xlrd строки 0-based
    return rows


def _dataframe_from_xlsx_sheet(ws) -> pd.DataFrame:
    merged_rows = _rows_with_merged_cells_xlsx(ws)
    rows: list[list[object]] = []

    for r in range(1, ws.max_row + 1):
        if r <= 9:
            continue
        if r in merged_rows:
            continue
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v not in (None, "") for v in values):
            rows.append(values)

    if not rows:
        return pd.DataFrame()
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    return pd.DataFrame(rows[1:], columns=header)


def _dataframe_from_xls_sheet(sheet) -> pd.DataFrame:
    merged_rows = _rows_with_merged_cells_xls(sheet)
    rows: list[list[object]] = []

    for r in range(sheet.nrows):
        row_num = r + 1
        if row_num <= 9:
            continue
        if row_num in merged_rows:
            continue
        values = sheet.row_values(r)
        if any(v not in (None, "") for v in values):
            rows.append(values)

    if not rows:
        return pd.DataFrame()
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    return pd.DataFrame(rows[1:], columns=header)


def _collect_rows_from_xlsx(source_name: str, binary: bytes | None = None, path: str | None = None) -> tuple[list[MatchRow], bool]:
    from openpyxl import load_workbook

    wb = load_workbook(filename=BytesIO(binary) if binary is not None else path, data_only=True)
    all_rows: list[MatchRow] = []
    has_required = False

    for sheet_name in wb.sheetnames:
        df = _dataframe_from_xlsx_sheet(wb[sheet_name])
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
                all_rows.append(MatchRow(str(code_value).strip(), name_text, translit_to_ru(name_text)))

    logging.getLogger("main").info("%s: найдено строк с латиницей: %s", source_name, len(all_rows))
    return all_rows, has_required


def _collect_rows_from_xls(source_name: str, binary: bytes | None = None, path: str | None = None) -> tuple[list[MatchRow], bool]:
    import xlrd

    if binary is not None:
        book = xlrd.open_workbook(file_contents=binary, formatting_info=True)
    else:
        book = xlrd.open_workbook(path, formatting_info=True)

    all_rows: list[MatchRow] = []
    has_required = False

    for sheet in book.sheets():
        df = _dataframe_from_xls_sheet(sheet)
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
                all_rows.append(MatchRow(str(code_value).strip(), name_text, translit_to_ru(name_text)))

    logging.getLogger("main").info("%s: найдено строк с латиницей: %s", source_name, len(all_rows))
    return all_rows, has_required


def _collect_rows_from_file(source_name: str, suffix: str, binary: bytes | None = None, path: str | None = None) -> tuple[list[MatchRow], bool]:
    suffix = suffix.lower()
    if suffix == ".xlsx":
        return _collect_rows_from_xlsx(source_name, binary=binary, path=path)
    if suffix == ".xls":
        return _collect_rows_from_xls(source_name, binary=binary, path=path)
    raise ValueError(f"Неподдерживаемый формат файла: {suffix}")


def scan_folder(folder_path: str, out_file: str, status_callback=None) -> dict:
    logger = logging.getLogger("main")

    def set_status(msg: str) -> None:
        logger.info(msg)
        if status_callback:
            status_callback(msg)

    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Папка не найдена: {folder_path}")

    files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"}])
    set_status(f"Найдено Excel-файлов: {len(files)}")

    all_rows: list[MatchRow] = []
    errors: list[tuple[str, str]] = []
    files_with_columns = 0

    for idx, file_path in enumerate(files, start=1):
        set_status(f"Обрабатываю файл {idx}/{len(files)}: {file_path.name}")
        try:
            rows, has_required = _collect_rows_from_file(
                source_name=file_path.name,
                suffix=file_path.suffix,
                path=str(file_path),
            )
            all_rows.extend(rows)
            if has_required:
                files_with_columns += 1
        except Exception as exc:
            errors.append((str(file_path), str(exc)))
            logger.error("Ошибка чтения %s: %s", file_path, exc)

    set_status("Формирую итоговый XLS...")
    final_output = write_xls(all_rows, out_file)
    set_status("Готово")

    return {
        "files": len(files),
        "files_with_columns": files_with_columns,
        "written": len(all_rows),
        "errors": errors,
        "output": final_output,
    }


def scan_uploaded_files(uploaded_files: list, out_file: str, status_callback=None) -> dict:
    logger = logging.getLogger("main")

    def set_status(msg: str) -> None:
        logger.info(msg)
        if status_callback:
            status_callback(msg)

    set_status(f"Загружено файлов: {len(uploaded_files)}")
    all_rows: list[MatchRow] = []
    errors: list[tuple[str, str]] = []
    files_with_columns = 0

    for idx, uf in enumerate(uploaded_files, start=1):
        set_status(f"Обрабатываю загруженный файл {idx}/{len(uploaded_files)}: {uf.name}")
        try:
            suffix = Path(uf.name).suffix.lower()
            rows, has_required = _collect_rows_from_file(
                source_name=uf.name,
                suffix=suffix,
                binary=uf.getvalue(),
            )
            all_rows.extend(rows)
            if has_required:
                files_with_columns += 1
        except Exception as exc:
            errors.append((uf.name, str(exc)))
            logger.error("Ошибка чтения %s: %s", uf.name, exc)

    set_status("Формирую итоговый XLS...")
    final_output = write_xls(all_rows, out_file)
    set_status("Готово")

    return {
        "files": len(uploaded_files),
        "files_with_columns": files_with_columns,
        "written": len(all_rows),
        "errors": errors,
        "output": final_output,
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Поиск латиницы в 'Наименование товаров' и экспорт в .xls")
    p.add_argument("--folder", required=True, help="Путь к папке с Excel-файлами")
    p.add_argument("--out", default="latin_names_report.xls", help="Путь к итоговому .xls")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    setup_logging()
    result = scan_folder(folder_path=args.folder, out_file=args.out)
    logging.getLogger("main").info("Готово. Строк: %s | Файл: %s", result["written"], result["output"])


if __name__ == "__main__":
    main()
