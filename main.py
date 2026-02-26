#!/usr/bin/env python3
"""
Парсер vologorost.ru:
- обходит категории и карточки товаров,
- отбирает товары, где в названии есть латиница,
- строит xlsx-отчёт с артикулом, оригинальным названием и транскрипцией.

Режим получения карточки:
1) requests + BeautifulSoup
2) fallback на Playwright для конкретной карточки, если requests не дал стабильных данных.
"""

from __future__ import annotations

import argparse
import logging
import random
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from playwright.sync_api import sync_playwright
except Exception:  # playwright опционален до момента fallback
    sync_playwright = None


LATIN_RE = re.compile(r"[A-Za-z]")
LATIN_TOKEN_RE = re.compile(r"[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*")
ARTICLE_LABEL_RE = re.compile(r"(артикул|код\s*товара|sku|модель)", re.IGNORECASE)

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

VOWELS = set("aeiouy")


@dataclass
class ProductData:
    url: str
    name: str
    article: str


class VolgorostParser:
    def __init__(
        self,
        start_url: str,
        max_pages_per_category: Optional[int] = None,
        timeout: int = 20,
    ) -> None:
        self.start_url = start_url
        self.max_pages_per_category = max_pages_per_category
        self.timeout = timeout
        self.logger = logging.getLogger(self.__class__.__name__)

        self.session = requests.Session()
        retry = Retry(
            total=5,
            connect=5,
            read=5,
            status=5,
            backoff_factor=0.8,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "HEAD"],
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                ),
                "Accept-Language": "ru,en;q=0.9",
            }
        )

    def polite_delay(self) -> None:
        time.sleep(random.uniform(0.2, 0.8))

    def _is_same_domain(self, url: str) -> bool:
        return urlparse(url).netloc == urlparse(self.start_url).netloc

    def fetch(self, url: str) -> Optional[requests.Response]:
        self.polite_delay()
        try:
            resp = self.session.get(url, timeout=self.timeout)
            if resp.status_code >= 400:
                self.logger.warning("HTTP %s для %s", resp.status_code, url)
                return None
            return resp
        except Exception as exc:
            self.logger.error("Ошибка запроса %s: %s", url, exc)
            return None

    def discover_category_urls(self) -> list[str]:
        """Поиск URL категорий по навигационным и каталожным ссылкам."""
        resp = self.fetch(self.start_url)
        if not resp:
            return []

        soup = BeautifulSoup(resp.text, "html.parser")
        categories: set[str] = set()

        candidate_scopes = [
            soup.select("nav a[href]"),
            soup.select("header a[href]"),
            soup.select("a[href*='catalog']"),
            soup.select("a[href*='katalog']"),
            soup.select("a[href*='category']"),
            soup.select(".menu a[href], .catalog a[href], .catalog-menu a[href]"),
        ]

        for links in candidate_scopes:
            for a in links:
                href = a.get("href")
                if not href:
                    continue
                full = urljoin(self.start_url, href.split("#")[0])
                if not self._is_same_domain(full):
                    continue
                parsed = urlparse(full)
                path = parsed.path.lower().strip("/")
                if not path:
                    continue
                tokens = ("catalog", "katalog", "category", "shop", "product-category")
                if any(t in path for t in tokens):
                    categories.add(full.rstrip("/"))

        if not categories:
            categories.add(self.start_url.rstrip("/"))

        self.logger.info("Найдено категорий: %s", len(categories))
        return sorted(categories)

    def _looks_like_product_url(self, url: str) -> bool:
        path = urlparse(url).path.lower().strip("/")
        if not path:
            return False
        bad_parts = ["cart", "basket", "login", "register", "search", "compare", "wishlist"]
        if any(p in path for p in bad_parts):
            return False
        product_hints = ["product", "товар", "item", "/p/", "-p-"]
        if any(h in path for h in product_hints):
            return True
        # Частый формат: длинный slug товара
        return len(path.split("/")) >= 2 and re.search(r"\d", path) is not None

    def _extract_pagination_links(self, soup: BeautifulSoup, page_url: str) -> list[str]:
        urls: set[str] = set()
        for a in soup.select("a[href]"):
            text = (a.get_text(" ", strip=True) or "").lower()
            rel = (a.get("rel") or [])
            cls = " ".join(a.get("class") or []).lower()
            href = a.get("href")
            if not href:
                continue
            full = urljoin(page_url, href)
            if not self._is_same_domain(full):
                continue
            if (
                "next" in rel
                or "след" in text
                or text in {">", ">>", "→"}
                or "next" in cls
                or re.search(r"[?&](page|p|PAGEN_\d+)=\d+", full)
            ):
                urls.add(full)
        return sorted(urls)

    def discover_product_urls(self, category_url: str) -> set[str]:
        """Обходит пагинацию категории и собирает ссылки на карточки."""
        seen_pages: set[str] = set()
        to_visit: list[str] = [category_url]
        product_urls: set[str] = set()
        page_count = 0

        while to_visit:
            current = to_visit.pop(0)
            current = current.rstrip("/")
            if current in seen_pages:
                continue
            if self.max_pages_per_category and page_count >= self.max_pages_per_category:
                break

            seen_pages.add(current)
            page_count += 1
            resp = self.fetch(current)
            if not resp:
                continue

            soup = BeautifulSoup(resp.text, "html.parser")
            for a in soup.select("a[href]"):
                href = a.get("href")
                if not href:
                    continue
                full = urljoin(current, href.split("#")[0]).rstrip("/")
                if not self._is_same_domain(full):
                    continue
                if self._looks_like_product_url(full):
                    product_urls.add(full)

            for n in self._extract_pagination_links(soup, current):
                nr = n.rstrip("/")
                if nr not in seen_pages and nr not in to_visit:
                    to_visit.append(nr)

        self.logger.info("Категория %s -> карточек: %s", category_url, len(product_urls))
        return product_urls

    def _extract_name_from_soup(self, soup: BeautifulSoup) -> str:
        selectors = [
            "h1",
            "meta[property='og:title']",
            ".product-title",
            ".product-name",
            ".page-title",
        ]
        for sel in selectors:
            node = soup.select_one(sel)
            if not node:
                continue
            if node.name == "meta":
                value = (node.get("content") or "").strip()
            else:
                value = node.get_text(" ", strip=True)
            if value:
                return value
        title_tag = soup.title.get_text(" ", strip=True) if soup.title else ""
        return title_tag

    def _extract_article_from_soup(self, soup: BeautifulSoup) -> str:
        # 1) Явные блоки с data/классами
        explicit_selectors = [
            "[itemprop='sku']",
            ".sku",
            ".product-sku",
            ".article",
            ".articul",
            "[class*='sku']",
            "[class*='article']",
        ]
        for sel in explicit_selectors:
            for node in soup.select(sel):
                txt = node.get_text(" ", strip=True)
                txt = re.sub(r"^\s*(артикул|sku)\s*[:#-]?\s*", "", txt, flags=re.IGNORECASE)
                if txt and len(txt) <= 80:
                    return txt

        # 2) Таблицы/списки характеристик
        for row in soup.select("tr, li, .property, .characteristics__item, .specification__row"):
            text = row.get_text(" ", strip=True)
            if ARTICLE_LABEL_RE.search(text):
                cleaned = ARTICLE_LABEL_RE.sub("", text)
                cleaned = re.sub(r"[:#\-]\s*", " ", cleaned).strip()
                if cleaned:
                    return cleaned

        # 3) Регулярка по всему тексту
        full_text = soup.get_text("\n", strip=True)
        m = re.search(r"(?:Артикул|SKU|Код\s*товара)\s*[:#]?\s*([A-Za-zА-Яа-я0-9\-_/\.]+)", full_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

        return ""

    def parse_product_requests(self, product_url: str) -> ProductData:
        resp = self.fetch(product_url)
        if not resp:
            raise RuntimeError("requests не получил страницу")
        soup = BeautifulSoup(resp.text, "html.parser")
        name = self._extract_name_from_soup(soup)
        article = self._extract_article_from_soup(soup)
        if not name:
            raise RuntimeError("не найдено имя товара через requests")
        if not article:
            self.logger.debug("Артикул не найден через requests: %s", product_url)
        return ProductData(url=product_url, name=name, article=article)

    def parse_product_playwright(self, product_url: str) -> ProductData:
        if sync_playwright is None:
            raise RuntimeError("playwright не установлен")

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(user_agent=self.session.headers["User-Agent"])
            page.goto(product_url, wait_until="networkidle", timeout=self.timeout * 1000)
            html = page.content()
            browser.close()

        soup = BeautifulSoup(html, "html.parser")
        name = self._extract_name_from_soup(soup)
        article = self._extract_article_from_soup(soup)
        if not name:
            raise RuntimeError("не найдено имя товара через playwright")
        return ProductData(url=product_url, name=name, article=article)

    def parse_product_auto(self, product_url: str) -> ProductData:
        try:
            data = self.parse_product_requests(product_url)
            # если имя есть, но статья пуста — считаем частичным успехом
            return data
        except Exception as exc:
            self.logger.warning("Fallback на Playwright для %s (%s)", product_url, exc)
            return self.parse_product_playwright(product_url)


def letter_translit(chunk: str) -> str:
    """Приближённая транскрипция латиницы в кириллицу."""
    s = chunk.lower()

    replacements = [
        ("sch", "щ"),
        ("sh", "ш"),
        ("ch", "ч"),
        ("zh", "ж"),
        ("kh", "х"),
        ("ph", "ф"),
        ("th", "т"),
        ("ck", "к"),
        ("qu", "кв"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)

    out: list[str] = []
    i = 0
    while i < len(s):
        c = s[i]
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
        i += 1

    return "".join(out)


def translit_token(token: str) -> str:
    low = token.lower()
    if low in EXCEPTION_MAP:
        return EXCEPTION_MAP[low]

    # Если дефисный токен: транскрибируем части, сохраняя дефис.
    parts = token.split("-")
    if len(parts) > 1:
        return "-".join(translit_token(part) for part in parts)

    # Смесь букв/цифр (например X5) — транскрибируем буквенные блоки.
    mixed = re.findall(r"[A-Za-z]+|\d+", token)
    if len(mixed) > 1:
        out = []
        for p in mixed:
            if p.isdigit():
                out.append(p)
            else:
                out.append(translit_token(p))
        return "".join(out)

    return letter_translit(token)


def translit_to_ru(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        tok = match.group(0)
        return translit_token(tok)

    return LATIN_TOKEN_RE.sub(repl, text)


def write_excel(rows: Iterable[ProductData], out_file: str) -> int:
    wb = Workbook()
    ws = wb.active
    ws.title = "latin"

    headers = ["Артикул", "Наименование товара", "Транскрипция"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    count = 0
    for item in rows:
        ws.append([item.article, item.name, translit_to_ru(item.name)])
        count += 1

    # автоширина + без переноса
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for c in col:
            value = "" if c.value is None else str(c.value)
            max_len = max(max_len, len(value))
            c.alignment = c.alignment.copy(wrap_text=False)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 100)

    wb.save(out_file)
    return count


def setup_logging(log_file: str = "parse.log") -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, mode="w", encoding="utf-8"),
        ],
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Парсер vologorost.ru для поиска латиницы в товарах")
    p.add_argument("--start-url", default="https://volgorost.ru/", help="Стартовый URL")
    p.add_argument("--out", default="voligorost_latin_names.xlsx", help="Путь к xlsx")
    p.add_argument(
        "--max-pages-per-category",
        type=int,
        default=None,
        help="Ограничение страниц на категорию (для теста)",
    )
    p.add_argument(
        "--concurrency",
        type=int,
        default=1,
        help="Параметр CLI для совместимости (в текущей версии используется последовательная обработка)",
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    setup_logging()
    logger = logging.getLogger("main")

    parser = VolgorostParser(
        start_url=args.start_url,
        max_pages_per_category=args.max_pages_per_category,
    )

    categories = parser.discover_category_urls()
    all_products: set[str] = set()
    for cat in categories:
        all_products.update(parser.discover_product_urls(cat))

    logger.info("Всего уникальных карточек: %s", len(all_products))

    processed = 0
    latin_rows: list[ProductData] = []
    errors: list[tuple[str, str]] = []

    for url in sorted(all_products):
        try:
            data = parser.parse_product_auto(url)
            processed += 1
            if LATIN_RE.search(data.name):
                latin_rows.append(data)
            logger.info(
                "Обработано: %s | latin: %s | %s",
                processed,
                len(latin_rows),
                url,
            )
        except Exception as exc:
            errors.append((url, str(exc)))
            logger.error("Ошибка карточки %s: %s", url, exc)

    written = write_excel(latin_rows, args.out)

    logger.info("=" * 50)
    logger.info("SUMMARY")
    logger.info("Категорий найдено: %s", len(categories))
    logger.info("Карточек обработано: %s", processed)
    logger.info("В отчёт попало: %s", written)
    logger.info("Ошибок: %s", len(errors))
    if errors:
        for u, e in errors[:50]:
            logger.info("ERROR URL: %s | %s", u, e)


if __name__ == "__main__":
    main()
