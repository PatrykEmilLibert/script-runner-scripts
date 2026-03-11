import base64
import concurrent.futures
import csv
import json
import os
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import requests


APP_TITLE = "CDON Artykuly"
DEFAULT_ACCOUNTS_PATH = str((Path(__file__).parent / "accounts.csv").resolve())
SANDBOX_BASE_URL = "https://merchants-api.sandbox.cdon.com/api"
PROD_BASE_URL = "https://merchants-api.cdon.com/api"
ARTICLES_ENDPOINT = "/v1/articles"

ACCENT = "#ff69b4"
ACCENT_HOVER = "#ec4fa3"
BG_MAIN = "#fff7fb"
BG_PANEL = "#ffffff"
BG_INPUT = "#fff0f8"
BORDER = "#f3bdd8"
TEXT = "#3c2030"
TEXT_MUTED = "#6f4f60"
SUCCESS = "#1f8a4c"
ERROR = "#b42318"

DEFAULT_REQUESTS_PER_MINUTE = 300
MAX_REQUESTS_PER_MINUTE = 350
DEFAULT_MAX_RETRIES = 6
DEFAULT_BACKOFF_SECONDS = 1.5
DEFAULT_PAGE_LIMIT = 100
MAX_PAGE_LIMIT = 10000
PAGE_BURST_SIZE = 350
PAGE_BURST_PAUSE_SECONDS = 0
SAFE_MODE_MAX_REQUESTS_PER_MINUTE = 100

CSV_FIELDNAMES = [
    "sku",
    "name",
    "quantity",
    "price",
    "currency",
    "price_market",
    "price_SE",
    "currency_SE",
    "price_DK",
    "currency_DK",
    "price_FI",
    "currency_FI",
    "price_NO",
    "currency_NO",
    "active",
    "status",
    "real_status",
]


def safe_float(value):
    if value is None:
        return ""
    try:
        return float(str(value).replace(",", "."))
    except ValueError:
        return ""


def safe_int(value):
    if value is None or value == "":
        return ""
    try:
        return int(float(value))
    except ValueError:
        return ""


class RequestRateLimiter:
    def __init__(self, requests_per_minute):
        if requests_per_minute < 1:
            raise ValueError("requests_per_minute musi byc > 0")
        self.min_interval = 60.0 / float(requests_per_minute)
        self._lock = threading.Lock()
        self._next_allowed_ts = 0.0

    def wait_turn(self):
        while True:
            with self._lock:
                now = time.monotonic()
                if now >= self._next_allowed_ts:
                    self._next_allowed_ts = now + self.min_interval
                    return
                wait_for = self._next_allowed_ts - now
            time.sleep(min(wait_for, 0.5))


class CDONArticlesClient:
    def __init__(
        self,
        merchant_id,
        api_token,
        use_sandbox=True,
        timeout=45,
        requests_per_minute=DEFAULT_REQUESTS_PER_MINUTE,
        max_retries=DEFAULT_MAX_RETRIES,
        backoff_seconds=DEFAULT_BACKOFF_SECONDS,
        safe_mode=False,
        log_callback=None,
    ):
        self.merchant_id = merchant_id
        self.api_token = api_token
        self.base_url = SANDBOX_BASE_URL if use_sandbox else PROD_BASE_URL
        self.timeout = timeout
        self.max_retries = max_retries
        self.backoff_seconds = float(backoff_seconds)
        self.safe_mode = bool(safe_mode)
        self.rate_limiter = RequestRateLimiter(requests_per_minute)
        self.log = log_callback or (lambda _msg: None)

    def _headers(self):
        auth_raw = f"{self.merchant_id}:{self.api_token}".encode("ascii")
        auth_basic = base64.b64encode(auth_raw).decode("ascii")
        return {
            "Authorization": f"Basic {auth_basic}",
            "Accept": "application/json",
            "Content-Type": "application/json",
            "x-merchant-id": self.merchant_id,
            "User-Agent": "CDON-Articles-GUI/1.0",
        }

    def _request_with_retry(self, method, url, skip_rate_limiter=False, **kwargs):
        transient_http_codes = {429, 500, 502, 503, 504}

        for attempt in range(self.max_retries + 1):
            if not skip_rate_limiter:
                self.rate_limiter.wait_turn()
            try:
                response = requests.request(
                    method=method,
                    url=url,
                    headers=self._headers(),
                    timeout=self.timeout,
                    **kwargs,
                )
            except (requests.Timeout, requests.ConnectionError) as exc:
                if attempt >= self.max_retries:
                    raise RuntimeError(f"Blad polaczenia po {attempt + 1} probach: {exc}") from exc
                wait_time = min(90.0, self.backoff_seconds * (2**attempt))
                self.log(f"Polaczenie nieudane. Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s")
                time.sleep(wait_time)
                continue

            if response.status_code in transient_http_codes and attempt < self.max_retries:
                retry_after_header = response.headers.get("Retry-After", "").strip()
                wait_time = None
                if retry_after_header:
                    try:
                        wait_time = float(retry_after_header)
                    except ValueError:
                        wait_time = None

                if wait_time is None:
                    wait_time = min(90.0, self.backoff_seconds * (2**attempt))

                if response.status_code == 429:
                    self.log(f"HTTP 429 (limit). Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s")
                else:
                    self.log(f"HTTP {response.status_code}. Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s")

                time.sleep(wait_time)
                continue

            return response

        raise RuntimeError("Nie udalo sie wykonac zapytania po wielu probach")

    @staticmethod
    def _page_size_params(page_limit):
        # Hidden endpoints sometimes use non-standard names for page size.
        return {
            "limit": page_limit,
            "page_size": page_limit,
            "pageSize": page_limit,
            "per_page": page_limit,
            "perPage": page_limit,
        }

    @staticmethod
    def _extract_articles(payload):
        if isinstance(payload, list):
            return payload
        if not isinstance(payload, dict):
            return []

        for key in ["articles", "items", "results", "data"]:
            value = payload.get(key)
            if isinstance(value, list):
                return value
        return []

    @staticmethod
    def _extract_total_count(payload):
        if not isinstance(payload, dict):
            return None
        for key in ["total", "total_count", "totalCount", "count"]:
            value = payload.get(key)
            if isinstance(value, int) and value >= 0:
                return value
            if isinstance(value, str) and value.isdigit():
                return int(value)
        meta = payload.get("meta")
        if isinstance(meta, dict):
            for key in ["total", "total_count", "totalCount", "count"]:
                value = meta.get(key)
                if isinstance(value, int) and value >= 0:
                    return value
                if isinstance(value, str) and value.isdigit():
                    return int(value)
        return None

    @staticmethod
    def _extract_next_link(payload):
        if not isinstance(payload, dict):
            return ""

        direct = payload.get("next") or payload.get("next_url") or payload.get("nextUrl")
        if isinstance(direct, str) and direct.strip():
            return direct.strip()

        links = payload.get("links")
        if isinstance(links, dict):
            candidate = links.get("next")
            if isinstance(candidate, str) and candidate.strip():
                return candidate.strip()

        paging = payload.get("paging")
        if isinstance(paging, dict):
            candidate = paging.get("next")
            if isinstance(candidate, str) and candidate.strip():
                return candidate.strip()

        return ""

    @staticmethod
    def _extract_next_cursor(payload):
        if not isinstance(payload, dict):
            return "", ""

        cursor_map = {
            "cursor": "cursor",
            "next_cursor": "cursor",
            "nextCursor": "cursor",
            "continuation": "continuation",
            "continuation_token": "continuation",
            "continuationToken": "continuation",
            "next_token": "next_token",
            "nextToken": "next_token",
            "page_token": "page_token",
            "pageToken": "page_token",
        }

        for src_key, request_key in cursor_map.items():
            value = payload.get(src_key)
            if isinstance(value, str) and value.strip():
                return request_key, value.strip()

        paging = payload.get("paging")
        if isinstance(paging, dict):
            for src_key, request_key in cursor_map.items():
                value = paging.get(src_key)
                if isinstance(value, str) and value.strip():
                    return request_key, value.strip()

        return "", ""

    @staticmethod
    def _batch_fingerprint(items):
        if not items:
            return (0, "", "")
        first = items[0] if isinstance(items[0], dict) else {}
        last = items[-1] if isinstance(items[-1], dict) else {}
        return (
            len(items),
            str(first.get("sku") or first.get("article_sku") or ""),
            str(last.get("sku") or last.get("article_sku") or ""),
        )

    def _fetch_payload(self, request_url, params, skip_rate_limiter=False):
        response = self._request_with_retry(
            "GET",
            request_url,
            skip_rate_limiter=skip_rate_limiter,
            params=params,
        )
        if response.status_code != 200:
            text = response.text[:500]
            raise RuntimeError(f"HTTP {response.status_code} przy GET /v1/articles: {text}")
        payload = json.loads(response.content.decode("utf-8", errors="replace"))
        items = self._extract_articles(payload)
        return payload, items

    def _build_params(self, page_limit, mode, page=None, offset=None, cursor_param="", cursor_value=""):
        params = self._page_size_params(page_limit)
        if mode == "page" and page is not None:
            params["page"] = page
        elif mode == "offset" and offset is not None:
            params["offset"] = offset
        elif mode == "cursor" and cursor_param and cursor_value:
            params[cursor_param] = cursor_value
        return params

    def _discover_numeric_mode(self, base_url, page_limit, first_items):
        if not first_items:
            return None

        first_fp = self._batch_fingerprint(first_items)
        probes = [
            ("page", self._build_params(page_limit, "page", page=2)),
            ("offset", self._build_params(page_limit, "offset", offset=len(first_items))),
        ]

        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            future_map = {
                pool.submit(self._fetch_payload, base_url, params): mode
                for mode, params in probes
            }
            for future in concurrent.futures.as_completed(future_map):
                mode = future_map[future]
                try:
                    _payload, items = future.result()
                except Exception as exc:
                    self.log(f"Probe {mode} nieudany: {exc}")
                    continue

                fp = self._batch_fingerprint(items)
                if items and fp != first_fp:
                    return mode

        return None

    @staticmethod
    def _extract_title(article):
        title = article.get("title")
        if isinstance(title, str):
            return title
        if isinstance(title, list):
            preferred_langs = ["pl-PL", "en-US", "sv-SE", "da-DK", "fi-FI", "nb-NO"]
            by_lang = {}
            for item in title:
                if not isinstance(item, dict):
                    continue
                lang = str(item.get("language") or "").strip()
                value = str(item.get("value") or "").strip()
                if lang and value:
                    by_lang[lang] = value

            for lang in preferred_langs:
                if lang in by_lang:
                    return by_lang[lang]

            for item in title:
                if isinstance(item, dict) and item.get("value"):
                    return str(item.get("value"))
        if isinstance(title, dict) and title.get("value"):
            return str(title.get("value"))
        return ""

    @staticmethod
    def _extract_price_map(article):
        price_obj = article.get("price")
        out = {}

        if isinstance(price_obj, (int, float, str)):
            value = safe_float(price_obj)
            if value != "":
                out["ALL"] = {"price": value, "currency": ""}
            return out

        if isinstance(price_obj, dict):
            value = price_obj.get("value")
            if isinstance(value, dict):
                amount = value.get("amount")
                if amount is None:
                    amount = value.get("amount_including_vat")
                currency = value.get("currency") or ""
                parsed = safe_float(amount)
                if parsed != "":
                    out["ALL"] = {"price": parsed, "currency": str(currency)}
            return out

        if isinstance(price_obj, list):
            for item in price_obj:
                if not isinstance(item, dict):
                    continue
                market = str(item.get("market") or "").strip() or "ALL"
                value = item.get("value") or {}
                amount = value.get("amount")
                if amount is None:
                    amount = value.get("amount_including_vat")
                currency = str(value.get("currency") or "")
                parsed = safe_float(amount)
                if parsed != "":
                    out[market] = {"price": parsed, "currency": currency}

        return out

    @staticmethod
    def _extract_active(article):
        status = article.get("status", "")
        for_sale = article.get("for_sale")

        if isinstance(for_sale, bool):
            return for_sale
        if isinstance(for_sale, str):
            lowered = for_sale.strip().lower()
            if lowered in {"true", "1", "yes", "tak"}:
                return True
            if lowered in {"false", "0", "no", "nie"}:
                return False

        if isinstance(status, str):
            lowered = status.strip().lower()
            if lowered in {"for sale", "for_sale", "active"}:
                return True
            if lowered in {"paused", "inactive", "not for sale", "not_for_sale", "deleted"}:
                return False

        return ""

    def _fetch_articles_pages(self, page_limit, max_pages=10000, stop_event=None, on_items_callback=None):
        base_url = f"{self.base_url}{ARTICLES_ENDPOINT}"
        total_items_fetched = 0

        mode = "page"
        page = 1
        offset = 0
        cursor_param = ""
        cursor_value = ""
        next_url = ""
        previous_batch_fingerprint = None
        pages_done = 0

        while pages_done < max_pages:
            if stop_event and stop_event.is_set():
                self.log("Otrzymano sygnal stop. Koncze pobieranie kolejnych stron.")
                break

            if next_url:
                request_url = next_url
                params = self._build_params(page_limit, mode="link")
            else:
                request_url = base_url
                params = self._build_params(
                    page_limit,
                    mode=mode,
                    page=page,
                    offset=offset,
                    cursor_param=cursor_param,
                    cursor_value=cursor_value,
                )

            payload, items = self._fetch_payload(request_url, params)
            if not items:
                break

            pages_done += 1
            fingerprint = self._batch_fingerprint(items)
            if previous_batch_fingerprint == fingerprint:
                if pages_done == 2:
                    discovered = self._discover_numeric_mode(base_url, page_limit, items)
                    if discovered and discovered != mode:
                        self.log(f"Wykryto inny tryb paginacji: {discovered}. Przelaczam.")
                        mode = discovered
                        next_url = ""
                        if discovered == "page":
                            page = 2
                        else:
                            offset = total_items_fetched
                        continue

                self.log("Wykryto powtarzajaca sie strone danych. Zatrzymuje paginacje, by uniknac petli.")
                break
            previous_batch_fingerprint = fingerprint

            total_items_fetched += len(items)
            if on_items_callback:
                on_items_callback(items)
            self.log(f"Pobrano strone {pages_done}: +{len(items)} (laczenie {total_items_fetched})")

            link = self._extract_next_link(payload)
            if link:
                next_url = link if link.startswith("http") else f"{self.base_url}{link}"
                mode = "link"
                continue

            next_cursor_param, next_cursor_value = self._extract_next_cursor(payload)
            if next_cursor_param and next_cursor_value:
                cursor_param = next_cursor_param
                cursor_value = next_cursor_value
                next_url = ""
                mode = "cursor"
                continue

            total_count = self._extract_total_count(payload)
            if mode in {"offset", "cursor", "link"}:
                offset += len(items)
                next_url = ""
                if total_count is not None and offset < total_count:
                    mode = "offset"
                    continue

            if mode == "page":
                if len(items) < page_limit or page >= max_pages:
                    break

                if self.safe_mode:
                    # In safe mode keep page fetching strictly sequential.
                    page += 1
                    continue

                start_page = page + 1
                end_page = min(max_pages, page + PAGE_BURST_SIZE)
                page_jobs = []
                for next_page in range(start_page, end_page + 1):
                    job_params = self._build_params(page_limit, mode="page", page=next_page)
                    page_jobs.append((next_page, base_url, job_params))

                with concurrent.futures.ThreadPoolExecutor(max_workers=len(page_jobs)) as pool:
                    future_map = {
                        pool.submit(self._fetch_payload, url, job_params, True): pnum
                        for pnum, url, job_params in page_jobs
                    }
                    ordered = []
                    for future in concurrent.futures.as_completed(future_map):
                        pnum = future_map[future]
                        payload_p, items_p = future.result()
                        ordered.append((pnum, payload_p, items_p))

                ordered.sort(key=lambda x: x[0])
                stop_after_batch = False
                for pnum, _payload_p, items_p in ordered:
                    if stop_event and stop_event.is_set():
                        stop_after_batch = True
                        break

                    if not items_p:
                        stop_after_batch = True
                        break

                    fp = self._batch_fingerprint(items_p)
                    if fp == previous_batch_fingerprint:
                        stop_after_batch = True
                        break

                    previous_batch_fingerprint = fp
                    total_items_fetched += len(items_p)
                    if on_items_callback:
                        on_items_callback(items_p)
                    pages_done += 1
                    self.log(f"Pobrano strone {pnum}: +{len(items_p)} (laczenie {total_items_fetched})")

                    if len(items_p) < page_limit:
                        stop_after_batch = True
                        break

                page = end_page + 1
                if stop_after_batch:
                    break

                if page <= max_pages and PAGE_BURST_PAUSE_SECONDS > 0:
                    self.log(
                        f"Batch {end_page - start_page + 1} stron zakonczony. Czekam {PAGE_BURST_PAUSE_SECONDS}s przed kolejnym batchem."
                    )
                    for _ in range(PAGE_BURST_PAUSE_SECONDS):
                        if stop_event and stop_event.is_set():
                            break
                        time.sleep(1)
                    if stop_event and stop_event.is_set():
                        self.log("Zatrzymano w trakcie oczekiwania miedzy batchami.")
                        break
                continue

            # Fallback when endpoint ignores page numbers: try offset once first page returns a full batch.
            if mode != "offset" and len(items) == page_limit:
                mode = "offset"
                offset += len(items)
                next_url = ""
                continue

            break

        return total_items_fetched

    def _article_to_row(self, article):
        if not isinstance(article, dict):
            return None

        sku = str(article.get("sku") or article.get("article_sku") or "").strip()
        if not sku:
            return None

        qty = article.get("quantity")
        if qty is None:
            qty = article.get("stock", "")
        qty = safe_int(qty)

        prices = self._extract_price_map(article)
        preferred_markets = ["SE", "DK", "FI", "NO", "ALL"]
        primary_market = next((m for m in preferred_markets if m in prices), "")
        if not primary_market and prices:
            primary_market = sorted(prices.keys())[0]

        primary_price = prices.get(primary_market, {}).get("price", "")
        primary_currency = prices.get(primary_market, {}).get("currency", "")
        title = self._extract_title(article)
        active = self._extract_active(article)
        status = str(article.get("status", ""))

        is_for_sale = bool(active is True)
        has_stock = isinstance(qty, int) and qty > 0
        real_status = "aktywne" if (is_for_sale and has_stock) else "nieaktywne"

        return {
            "sku": sku,
            "name": title,
            "quantity": qty,
            "price": primary_price,
            "currency": primary_currency,
            "price_market": primary_market,
            "price_SE": prices.get("SE", {}).get("price", ""),
            "currency_SE": prices.get("SE", {}).get("currency", ""),
            "price_DK": prices.get("DK", {}).get("price", ""),
            "currency_DK": prices.get("DK", {}).get("currency", ""),
            "price_FI": prices.get("FI", {}).get("price", ""),
            "currency_FI": prices.get("FI", {}).get("currency", ""),
            "price_NO": prices.get("NO", {}).get("price", ""),
            "currency_NO": prices.get("NO", {}).get("currency", ""),
            "active": active,
            "status": status,
            "real_status": real_status,
        }

    def export_articles_to_csv(self, output_path, page_limit=DEFAULT_PAGE_LIMIT, stop_event=None):
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        written_rows = 0
        with open(output_path, "w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_FIELDNAMES, delimiter=";")
            writer.writeheader()

            def on_items(items):
                nonlocal written_rows
                rows = []
                for article in items:
                    row = self._article_to_row(article)
                    if row is not None:
                        rows.append(row)
                if rows:
                    writer.writerows(rows)
                    written_rows += len(rows)
                    handle.flush()
                    self.log(f"Dopisano do CSV: +{len(rows)} (razem zapisane {written_rows})")

            self._fetch_articles_pages(
                page_limit=page_limit,
                stop_event=stop_event,
                on_items_callback=on_items,
            )

        return written_rows

    def fetch_all_articles(self, page_limit=DEFAULT_PAGE_LIMIT, stop_event=None):
        all_items = []

        def append_items(items):
            all_items.extend(items)

        self._fetch_articles_pages(page_limit=page_limit, stop_event=stop_event, on_items_callback=append_items)

        by_sku = {}
        for article in all_items:
            if not isinstance(article, dict):
                continue
            sku = str(article.get("sku") or article.get("article_sku") or "").strip()
            if not sku:
                continue

            qty = article.get("quantity")
            if qty is None:
                qty = article.get("stock", "")
            qty = safe_int(qty)

            prices = self._extract_price_map(article)
            preferred_markets = ["SE", "DK", "FI", "NO", "ALL"]
            primary_market = next((m for m in preferred_markets if m in prices), "")
            if not primary_market and prices:
                primary_market = sorted(prices.keys())[0]

            primary_price = prices.get(primary_market, {}).get("price", "")
            primary_currency = prices.get(primary_market, {}).get("currency", "")
            title = self._extract_title(article)
            active = self._extract_active(article)
            status = str(article.get("status", ""))

            by_sku[sku] = {
                "sku": sku,
                "name": title,
                "quantity": qty,
                "price": primary_price,
                "currency": primary_currency,
                "price_market": primary_market,
                "price_SE": prices.get("SE", {}).get("price", ""),
                "currency_SE": prices.get("SE", {}).get("currency", ""),
                "price_DK": prices.get("DK", {}).get("price", ""),
                "currency_DK": prices.get("DK", {}).get("currency", ""),
                "price_FI": prices.get("FI", {}).get("price", ""),
                "currency_FI": prices.get("FI", {}).get("currency", ""),
                "price_NO": prices.get("NO", {}).get("price", ""),
                "currency_NO": prices.get("NO", {}).get("currency", ""),
                "active": active,
                "status": status,
            }

        return list(by_sku.values())


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")

        self.title(APP_TITLE)
        self.geometry("1100x760")
        self.minsize(960, 660)
        self.configure(fg_color=BG_MAIN)

        self.accounts_path_var = ctk.StringVar(value=DEFAULT_ACCOUNTS_PATH)
        self.output_csv_var = ctk.StringVar(value="")
        self.use_sandbox_var = ctk.BooleanVar(value=False)
        self.safe_mode_var = ctk.BooleanVar(value=False)
        self.rate_limit_var = ctk.StringVar(value=str(DEFAULT_REQUESTS_PER_MINUTE))
        self.page_limit_var = ctk.StringVar(value=str(DEFAULT_PAGE_LIMIT))

        self.accounts = {}
        self.account_check_vars = {}
        self.accounts_checks_frame = None
        self.is_running = False
        self.stop_event = threading.Event()

        self._build_ui()
        self._load_accounts()

    @staticmethod
    def _safe_filename_part(value):
        safe = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in value.strip())
        return safe.strip("_") or "konto"

    def _resolve_output_dir(self, output_path):
        output_path = output_path.strip()
        if not output_path:
            raise ValueError("Brak sciezki wyjsciowej.")

        if os.path.isdir(output_path):
            return output_path

        lower = output_path.lower()
        if lower.endswith(".csv"):
            return os.path.dirname(output_path) or os.getcwd()

        return output_path

    def _build_account_output_path(self, output_dir, account_name, run_stamp):
        safe_name = self._safe_filename_part(account_name)
        filename = f"{safe_name}_{run_stamp}.csv"
        return os.path.join(output_dir, filename)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        header.grid(row=0, column=0, padx=18, pady=(18, 10), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=APP_TITLE,
            text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, padx=16, pady=(14, 2), sticky="w")

        ctk.CTkLabel(
            header,
            text="Pobieranie danych artykulow: SKU, ilosc, cena, nazwa, aktywnosc.",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=13),
        ).grid(row=1, column=0, padx=16, pady=(0, 14), sticky="w")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(2, weight=1)

        config = ctk.CTkFrame(content, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        config.grid(row=0, column=0, sticky="ew")
        config.grid_columnconfigure(1, weight=1)

        row = 0
        ctk.CTkLabel(config, text="Plik kont", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkEntry(config, textvariable=self.accounts_path_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=row, column=1, padx=8, pady=10, sticky="ew"
        )
        ctk.CTkButton(
            config,
            text="Wybierz",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._pick_accounts,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        ctk.CTkLabel(config, text="Konta (ptaszki)", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="nw")
        self.accounts_checks_frame = ctk.CTkScrollableFrame(
            config,
            fg_color=BG_INPUT,
            border_color=BORDER,
            border_width=1,
            height=120,
        )
        self.accounts_checks_frame.grid(row=row, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(
            config,
            text="Odswiez",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._load_accounts,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        ctk.CTkLabel(config, text="Folder bazowy", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkEntry(config, textvariable=self.output_csv_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=row, column=1, padx=8, pady=10, sticky="ew"
        )
        ctk.CTkButton(
            config,
            text="Wybierz",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._pick_output_csv,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        options = ctk.CTkFrame(config, fg_color="transparent")
        options.grid(row=row, column=0, columnspan=3, padx=14, pady=(2, 12), sticky="ew")

        ctk.CTkCheckBox(
            options,
            text="Sandbox",
            variable=self.use_sandbox_var,
            text_color=TEXT,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            border_color=BORDER,
        ).grid(row=0, column=0, padx=(0, 14), pady=4, sticky="w")

        ctk.CTkCheckBox(
            options,
            text="Tryb bezpieczny",
            variable=self.safe_mode_var,
            text_color=TEXT,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            border_color=BORDER,
        ).grid(row=0, column=1, padx=(0, 14), pady=4, sticky="w")

        ctk.CTkLabel(options, text="Limit req/min", text_color=TEXT).grid(row=0, column=2, padx=(0, 8), sticky="w")
        ctk.CTkEntry(options, width=90, textvariable=self.rate_limit_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=0, column=3, padx=(0, 14), sticky="w"
        )

        ctk.CTkLabel(options, text="Page limit", text_color=TEXT).grid(row=0, column=4, padx=(0, 8), sticky="w")
        ctk.CTkEntry(options, width=90, textvariable=self.page_limit_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=0, column=5, padx=(0, 0), sticky="w"
        )

        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.grid(row=1, column=0, pady=(10, 10), sticky="ew")
        controls.grid_columnconfigure(2, weight=1)

        self.start_btn = ctk.CTkButton(
            controls,
            text="Pobierz artykuly",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="#ffffff",
            width=180,
            command=self._start,
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 8), sticky="w")

        self.stop_btn = ctk.CTkButton(
            controls,
            text="Zatrzymaj i zapisz",
            fg_color="#f7b267",
            hover_color="#ea9d4e",
            text_color=TEXT,
            width=170,
            state="disabled",
            command=self._request_stop,
        )
        self.stop_btn.grid(row=0, column=1, padx=(0, 8), sticky="w")

        self.clear_btn = ctk.CTkButton(
            controls,
            text="Wyczysc log",
            fg_color="#f2c7de",
            hover_color="#e9b5d1",
            text_color=TEXT,
            width=130,
            command=self._clear_log,
        )
        self.clear_btn.grid(row=0, column=2, padx=(0, 8), sticky="w")

        self.progress = ctk.CTkProgressBar(controls, progress_color=ACCENT, fg_color="#f4d5e5")
        self.progress.grid(row=0, column=3, sticky="ew")
        self.progress.set(0)

        log_panel = ctk.CTkFrame(content, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        log_panel.grid(row=2, column=0, sticky="nsew")
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_panel,
            text="Log",
            text_color=TEXT,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        self.log_box = ctk.CTkTextbox(
            log_panel,
            fg_color=BG_INPUT,
            text_color=TEXT,
            border_width=1,
            border_color=BORDER,
            wrap="word",
        )
        self.log_box.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

    def _log(self, msg, color=None):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {msg}\n"
        self.log_box.insert("end", line)
        self.log_box.see("end")

        if color:
            self.log_box.tag_add(color, "end-2l", "end-1l")
            self.log_box.tag_config(color, foreground=color)

    def _clear_log(self):
        self.log_box.delete("1.0", "end")

    def _pick_accounts(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik kont",
            filetypes=[("CSV", "*.csv"), ("Wszystkie", "*.*")],
        )
        if path:
            self.accounts_path_var.set(path)
            self._load_accounts()

    def _pick_output_csv(self):
        path = filedialog.askdirectory(title="Wybierz folder bazowy eksportu")
        if path:
            self.output_csv_var.set(path)

    def _refresh_accounts_checkboxes(self):
        if self.accounts_checks_frame is None:
            return

        for child in self.accounts_checks_frame.winfo_children():
            child.destroy()

        self.account_check_vars = {}
        names = sorted(self.accounts.keys())
        if not names:
            ctk.CTkLabel(self.accounts_checks_frame, text="Brak kont", text_color=TEXT_MUTED).grid(
                row=0, column=0, padx=8, pady=8, sticky="w"
            )
            return

        for idx, name in enumerate(names):
            var = ctk.BooleanVar(value=True)
            self.account_check_vars[name] = var
            ctk.CTkCheckBox(
                self.accounts_checks_frame,
                text=name,
                variable=var,
                text_color=TEXT,
                fg_color=ACCENT,
                hover_color=ACCENT_HOVER,
                border_color=BORDER,
            ).grid(row=idx, column=0, padx=8, pady=4, sticky="w")

    def _get_selected_accounts(self):
        selected = []
        for name, var in self.account_check_vars.items():
            if var.get() and name in self.accounts:
                selected.append((name, self.accounts[name]))
        return selected

    def _load_accounts(self):
        path = self.accounts_path_var.get().strip()
        self.accounts = {}

        if not path or not os.path.exists(path):
            self._refresh_accounts_checkboxes()
            self._log("Nie znaleziono pliku kont.", ERROR)
            return

        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                delimiter = ";" if sample.count(";") >= sample.count(",") else ","
                reader = csv.DictReader(handle, delimiter=delimiter)

                for row in reader:
                    account_name = (row.get("Nazwa Konta") or row.get("account_name") or "").strip()
                    merchant_id = (row.get("MerchantID") or row.get("merchant_id") or "").strip()
                    api_token = (row.get("APIToken") or row.get("api_token") or "").strip()

                    if account_name and merchant_id and api_token:
                        self.accounts[account_name] = {
                            "merchant_id": merchant_id,
                            "api_token": api_token,
                        }

            if not self.accounts:
                self._refresh_accounts_checkboxes()
                self._log("Plik kont wczytany, ale brak poprawnych rekordow.", ERROR)
                return

            names = sorted(self.accounts.keys())
            self._refresh_accounts_checkboxes()
            self._log(f"Wczytano konta: {len(names)}", SUCCESS)
        except Exception as exc:
            self._refresh_accounts_checkboxes()
            self._log(f"Blad czytania pliku kont: {exc}", ERROR)

    def _validate(self):
        if self.is_running:
            return False

        output_path = self.output_csv_var.get().strip()
        if not output_path:
            messagebox.showerror("Blad", "Wybierz folder bazowy eksportu.")
            return False

        if not self.accounts:
            messagebox.showerror("Blad", "Brak kont do pobrania.")
            return False

        if not self._get_selected_accounts():
            messagebox.showerror("Blad", "Zaznacz co najmniej jedno konto (ptaszkiem).")
            return False

        try:
            req_per_min = int(self.rate_limit_var.get().strip())
            if req_per_min < 1 or req_per_min > MAX_REQUESTS_PER_MINUTE:
                raise ValueError
        except ValueError:
            messagebox.showerror("Blad", f"Limit req/min musi byc liczba od 1 do {MAX_REQUESTS_PER_MINUTE}.")
            return False

        try:
            page_limit = int(self.page_limit_var.get().strip())
            if page_limit < 1 or page_limit > MAX_PAGE_LIMIT:
                raise ValueError
        except ValueError:
            messagebox.showerror("Blad", f"Page limit musi byc liczba od 1 do {MAX_PAGE_LIMIT}.")
            return False

        return True

    def _set_running(self, running):
        self.is_running = running
        state = "disabled" if running else "normal"
        self.start_btn.configure(state=state)
        self.stop_btn.configure(state="normal" if running else "disabled")

    def _request_stop(self):
        if not self.is_running:
            return
        self.stop_event.set()
        self._log("Otrzymano polecenie zatrzymania. Zapisze to, co juz pobrano.")

    def _start(self):
        if not self._validate():
            return

        self.stop_event.clear()
        self._set_running(True)
        self.progress.set(0)
        worker = threading.Thread(target=self._run_job, daemon=True)
        worker.start()

    def _run_job(self):
        try:
            output_path = self.output_csv_var.get().strip()
            req_per_min = int(self.rate_limit_var.get().strip())
            page_limit = int(self.page_limit_var.get().strip())
            use_sandbox = bool(self.use_sandbox_var.get())
            safe_mode = bool(self.safe_mode_var.get())

            if safe_mode and req_per_min > SAFE_MODE_MAX_REQUESTS_PER_MINUTE:
                self._log(
                    f"Tryb bezpieczny: limit req/min zmieniony z {req_per_min} na {SAFE_MODE_MAX_REQUESTS_PER_MINUTE}."
                )
                req_per_min = SAFE_MODE_MAX_REQUESTS_PER_MINUTE

            base_output_dir = self._resolve_output_dir(output_path)
            export_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = os.path.join(base_output_dir, f"exporty CDON {export_stamp}")
            os.makedirs(output_dir, exist_ok=True)
            run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            accounts_to_fetch = self._get_selected_accounts()

            if not accounts_to_fetch:
                raise RuntimeError("Brak kont do pobrania.")

            self._log(f"Konta do pobrania: {len(accounts_to_fetch)}")
            self._log(f"Srodowisko: {'SANDBOX' if use_sandbox else 'PRODUKCJA'}")
            self._log(f"Tryb bezpieczny: {'TAK' if safe_mode else 'NIE'}")
            self._log(f"Limit API: {req_per_min} req/min")
            self._log(f"Page limit: {page_limit}")
            self._log(f"Folder bazowy: {base_output_dir}")
            self._log(f"Folder wyjsciowy: {output_dir}")

            self._log("Pobieranie artykulow z API ze wszystkich kont...")

            def run_one_account(account_name, creds):
                account_output_path = self._build_account_output_path(output_dir, account_name, run_stamp)
                self._log(f"[{account_name}] Start -> {account_output_path}")
                client = CDONArticlesClient(
                    merchant_id=creds["merchant_id"],
                    api_token=creds["api_token"],
                    use_sandbox=use_sandbox,
                    requests_per_minute=req_per_min,
                    max_retries=DEFAULT_MAX_RETRIES,
                    backoff_seconds=DEFAULT_BACKOFF_SECONDS,
                    safe_mode=safe_mode,
                    log_callback=lambda msg: self._log(f"[{account_name}] {msg}"),
                )
                written_rows = client.export_articles_to_csv(
                    output_path=account_output_path,
                    page_limit=page_limit,
                    stop_event=self.stop_event,
                )
                self._log(f"[{account_name}] Zakonczono. Artykuly: {written_rows}")
                return account_name, written_rows, account_output_path

            results = []
            account_workers = 1 if safe_mode else len(accounts_to_fetch)
            with concurrent.futures.ThreadPoolExecutor(max_workers=account_workers) as pool:
                future_map = {
                    pool.submit(run_one_account, account_name, creds): account_name
                    for account_name, creds in accounts_to_fetch
                }
                for future in concurrent.futures.as_completed(future_map):
                    account_name = future_map[future]
                    try:
                        results.append(future.result())
                    except Exception as exc:
                        self._log(f"[{account_name}] Blad: {exc}", ERROR)

            total_rows = sum(item[1] for item in results)
            self.progress.set(0.8)
            self.progress.set(1.0)

            self._log(f"Zapisane pliki: {len(results)}", SUCCESS)
            self._log(f"Artykuly lacznie: {total_rows}", SUCCESS)

            if self.stop_event.is_set():
                title = "Zatrzymano"
                message = (
                    f"Zatrzymano pobieranie przez uzytkownika.\n"
                    f"Zapisane pliki: {len(results)}\nArtykuly lacznie: {total_rows}."
                )
            else:
                title = "Gotowe"
                message = (
                    f"Pobrano dane dla {len(results)} kont.\n"
                    f"Artykuly lacznie: {total_rows}."
                )

            self.after(
                0,
                lambda: messagebox.showinfo(title, message),
            )
        except Exception as exc:
            self._log(f"Blad: {exc}", ERROR)
            self.after(0, lambda: messagebox.showerror("Blad", str(exc)))
        finally:
            self.after(0, lambda: self._set_running(False))


if __name__ == "__main__":
    app = App()
    app.mainloop()
