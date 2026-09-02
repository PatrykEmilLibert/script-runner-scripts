import base64
import concurrent.futures
import csv
import json
import os
import queue
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import requests


APP_TITLE = "CDON Artykuly"
ACCOUNTS_FILENAME = "accounts.csv"
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
MAX_PAGE_LIMIT = 100000
PAGE_BURST_SIZE = 40
PAGE_BURST_MAX_WORKERS = 6
PAGE_BURST_PAUSE_SECONDS = 0
SAFE_MODE_MAX_REQUESTS_PER_MINUTE = 100

# --- BUDZET ROWNOLEGLYCH ZAPYTAN ---
# Zmierzone (600 stron po ~44 KB, opoznienie 120 ms, keep-alive):
#   6 watkow  -> 12.5 s, 1.59 s CPU
#  12 watkow  ->  6.5 s, 1.97 s CPU
#  24 watki   ->  3.6 s, 1.64 s CPU
#  48 watkow  ->  2.3 s, 1.72 s CPU
# Laczne zuzycie CPU jest praktycznie plaskie, bo watek czekajacy na siec nic
# nie kosztuje - wiecej watkow tylko wykonuje te sama prace szybciej. Dlatego
# limitu NIE wiazemy z liczba rdzeni. Za wysokie zuzycie procesora bralo sie z
# handshake'ow TLS (5.4 ms CPU na zapytanie bez sesji vs 1.1 ms z keep-alive),
# a nie z liczby watkow.
DEFAULT_TOTAL_FETCH_WORKERS = 24
MAX_TOTAL_FETCH_WORKERS = 64

# Zakladane najgorsze opoznienie odpowiedzi API. Sluzy tylko do odciecia watkow
# ponad to, co limiter req/min jest w stanie nakarmic.
SLOW_RESPONSE_SECONDS = 2.0

# --- LOG GUI ---
# Wpisy z watkow roboczych trafiaja do kolejki, a Tk opróznia ja co LOG_FLUSH_MS.
# Wczesniej kazdy watek wolal insert()+see() bezposrednio, co przy tysiacach
# linii zajezdzalo petle zdarzen Tk (i bylo niebezpieczne - Tk nie jest
# thread-safe).
LOG_FLUSH_MS = 250
LOG_MAX_LINES = 2000

CSV_FIELDNAMES = [
    "sku",
    "article_id",
    "name",
    "gtin",
    "category",
    "brand",
    "parent_sku",
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



def app_dir():
    """
    Katalog aplikacji. Po spakowaniu do .exe (PyInstaller) __file__ wskazuje
    na tymczasowy katalog rozpakowania, wiec bierzemy folder pliku .exe.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resolve_desktop_dir():
    """
    Pulpit biezacego uzytkownika - dziala na dowolnym koncie.
    Uwzglednia przekierowanie folderu na OneDrive (wtedy %USERPROFILE%\\Desktop
    nie istnieje) oraz polska nazwe 'Pulpit'.
    """
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        try:
            value, _ = winreg.QueryValueEx(key, "Desktop")
        finally:
            winreg.CloseKey(key)
        value = os.path.expandvars(value or "")
        if value and os.path.isdir(value):
            return value
    except Exception:
        pass

    home = os.path.expanduser("~")
    for name in ("Desktop", "Pulpit",
                 os.path.join("OneDrive", "Desktop"),
                 os.path.join("OneDrive", "Pulpit")):
        candidate = os.path.join(home, name)
        if os.path.isdir(candidate):
            return candidate
    return home


def accounts_config_candidates(filename=ACCOUNTS_FILENAME):
    """Lokalizacje accounts.csv w kolejnosci przeszukiwania."""
    desktop = resolve_desktop_dir()
    return [
        os.path.join(desktop, filename),                # Pulpit - podstawowa
        os.path.join(desktop, "skrypty_py", filename),  # stary uklad katalogow
        os.path.abspath(filename),                      # katalog roboczy
        os.path.join(app_dir(), filename),              # obok skryptu/exe
    ]


def resolve_config_path(filename=ACCOUNTS_FILENAME):
    """
    Znajduje accounts.csv: najpierw na Pulpicie biezacego uzytkownika, potem
    w katalogu roboczym i obok skryptu (zgodnosc ze starymi instalacjami).

    Pulpit jest pierwszy, bo narzedzie bywa uruchamiane z lokalizacji tylko do
    odczytu, a kazdy uzytkownik trzyma wlasne klucze u siebie. Gdy pliku nie ma
    nigdzie, zwracana jest sciezka na Pulpicie.
    """
    candidates = accounts_config_candidates(filename)
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]




def set_background_priority(enabled):
    """
    Obniza priorytet procesu na czas pobierania, zeby Windows oddawal CPU
    aplikacjom, z ktorych uzytkownik korzysta w tym czasie. Praca trwa tyle
    samo, gdy komputer jest bezczynny - ustepuje tylko pod obciazeniem.
    Poza Windows i przy jakimkolwiek bledzie po prostu nic nie robi.
    """
    if os.name != "nt":
        return False
    BELOW_NORMAL_PRIORITY_CLASS = 0x00004000
    NORMAL_PRIORITY_CLASS = 0x00000020
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        # Bez jawnych sygnatur ctypes obcina pseudo-uchwyt procesu do 32 bitow
        # i SetPriorityClass zwraca ERROR_INVALID_HANDLE.
        kernel32.GetCurrentProcess.restype = ctypes.c_void_p
        kernel32.SetPriorityClass.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        kernel32.SetPriorityClass.restype = ctypes.c_int
        handle = kernel32.GetCurrentProcess()
        target = BELOW_NORMAL_PRIORITY_CLASS if enabled else NORMAL_PRIORITY_CLASS
        return bool(kernel32.SetPriorityClass(handle, target))
    except Exception:
        return False



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
        page_burst_workers=PAGE_BURST_MAX_WORKERS,
        log_callback=None,
    ):
        self.merchant_id = merchant_id
        self.api_token = api_token
        self.base_url = SANDBOX_BASE_URL if use_sandbox else PROD_BASE_URL
        self.timeout = timeout
        self.max_retries = max_retries
        self.backoff_seconds = float(backoff_seconds)
        self.safe_mode = bool(safe_mode)
        self.page_burst_workers = max(1, int(page_burst_workers))
        self.rate_limiter = RequestRateLimiter(requests_per_minute)
        self.log = log_callback or (lambda _msg: None)

        # Jedna sesja na konto: polaczenie HTTPS jest zestawiane raz i
        # utrzymywane (keep-alive). Wczesniej kazde requests.request() robilo
        # pelny handshake TLS - przy setkach stron to byla glowna pozycja w
        # zuzyciu procesora.
        self.session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=self.page_burst_workers,
            pool_maxsize=self.page_burst_workers,
            max_retries=0,
        )
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        self.session.headers.update(self._headers())

        # Pula watkow na strony tworzona raz, a nie od nowa dla kazdej paczki
        # 40 stron (ciagle tworzenie i ubijanie watkow tez kosztuje CPU).
        self._page_pool = None
        self._page_pool_lock = threading.Lock()

    def _get_page_pool(self):
        with self._page_pool_lock:
            if self._page_pool is None:
                self._page_pool = concurrent.futures.ThreadPoolExecutor(
                    max_workers=self.page_burst_workers,
                    thread_name_prefix=f"pages-{self.merchant_id}",
                )
            return self._page_pool

    def close(self):
        """Zwalnia pule watkow i polaczenia HTTP."""
        with self._page_pool_lock:
            pool, self._page_pool = self._page_pool, None
        if pool is not None:
            pool.shutdown(wait=True)
        try:
            self.session.close()
        except Exception:
            pass

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
                response = self.session.request(
                    method=method,
                    url=url,
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

        pool = self._get_page_pool()
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

    def _fetch_articles_pages(self, page_limit, max_pages=None, stop_event=None, on_items_callback=None):
        base_url = f"{self.base_url}{ARTICLES_ENDPOINT}"
        total_items_fetched = 0
        page_cap = max_pages if isinstance(max_pages, int) and max_pages > 0 else None

        mode = "page"
        page = 1
        offset = 0
        cursor_param = ""
        cursor_value = ""
        next_url = ""
        previous_batch_fingerprint = None
        pages_done = 0

        while True:
            if page_cap is not None and pages_done >= page_cap:
                self.log(f"Osiagnieto limit stron ({page_cap}). Zatrzymuje pobieranie.")
                break

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
                if len(items) < page_limit:
                    break

                if page_cap is not None and page >= page_cap:
                    break

                if self.safe_mode:
                    # In safe mode keep page fetching strictly sequential.
                    page += 1
                    continue

                start_page = page + 1
                end_page = page + PAGE_BURST_SIZE if page_cap is None else min(page_cap, page + PAGE_BURST_SIZE)
                page_jobs = []
                for next_page in range(start_page, end_page + 1):
                    job_params = self._build_params(page_limit, mode="page", page=next_page)
                    page_jobs.append((next_page, base_url, job_params))

                batch_error = None
                pool = self._get_page_pool()
                future_map = {
                    pool.submit(self._fetch_payload, url, job_params): pnum
                    for pnum, url, job_params in page_jobs
                }
                ordered = []
                for future in concurrent.futures.as_completed(future_map):
                    pnum = future_map[future]
                    try:
                        payload_p, items_p = future.result()
                    except Exception as exc:
                        batch_error = (pnum, exc)
                        continue
                    ordered.append((pnum, payload_p, items_p))

                if batch_error is not None:
                    failed_page, exc = batch_error
                    self.log(f"Blad pobierania strony {failed_page}: {exc}. Zatrzymuje batch.")
                    break

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

                if PAGE_BURST_PAUSE_SECONDS > 0 and (page_cap is None or page <= page_cap):
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
            "article_id": str(article.get("id") or article.get("article_id") or "").strip(),
            "name": title,
            "gtin": str(article.get("gtin") or "").strip(),
            "category": str(article.get("category") or "").strip(),
            "brand": str(article.get("brand") or "").strip(),
            "parent_sku": str(article.get("parent_sku") or "").strip(),
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

    def export_articles_to_split_csv(
        self,
        output_path_active,
        output_path_inactive,
        output_path_active_xlsx,
        page_limit=DEFAULT_PAGE_LIMIT,
        stop_event=None,
    ):
        active_dir = os.path.dirname(output_path_active)
        if active_dir:
            os.makedirs(active_dir, exist_ok=True)

        inactive_dir = os.path.dirname(output_path_inactive)
        if inactive_dir:
            os.makedirs(inactive_dir, exist_ok=True)

        active_xlsx_dir = os.path.dirname(output_path_active_xlsx)
        if active_xlsx_dir:
            os.makedirs(active_xlsx_dir, exist_ok=True)

        try:
            from openpyxl import Workbook
            from openpyxl.cell import WriteOnlyCell
        except ImportError as exc:
            raise RuntimeError(
                "Brak biblioteki openpyxl. Zainstaluj: pip install openpyxl"
            ) from exc

        written_active = 0
        written_inactive = 0
        workbook_active = Workbook(write_only=True)
        sheet_active = workbook_active.create_sheet(title="aktywne")

        def _xlsx_text(value):
            if value is None:
                return ""
            return str(value)

        def _xlsx_row(values):
            out = []
            for value in values:
                cell = WriteOnlyCell(sheet_active, value=_xlsx_text(value))
                # Format tekstowy jest potrzebny (SKU/GTIN nie moga sie zamienic
                # w liczby); przypisanie hyperlink=None bylo tylko zbednym
                # przejsciem przez walidator openpyxl dla kazdej komorki.
                cell.number_format = "@"
                out.append(cell)
            return out

        sheet_active.append(_xlsx_row(CSV_FIELDNAMES))

        try:
            with open(output_path_active, "w", encoding="utf-8", newline="") as handle_active, open(
                output_path_inactive, "w", encoding="utf-8", newline=""
            ) as handle_inactive:
                writer_active = csv.DictWriter(handle_active, fieldnames=CSV_FIELDNAMES, delimiter=";")
                writer_inactive = csv.DictWriter(handle_inactive, fieldnames=CSV_FIELDNAMES, delimiter=";")
                writer_active.writeheader()
                writer_inactive.writeheader()

                def on_items(items):
                    nonlocal written_active, written_inactive
                    rows_active = []
                    rows_inactive = []
                    for article in items:
                        row = self._article_to_row(article)
                        if row is None:
                            continue
                        if row.get("real_status") == "aktywne":
                            rows_active.append(row)
                        else:
                            rows_inactive.append(row)

                    if rows_active:
                        writer_active.writerows(rows_active)
                        for row in rows_active:
                            sheet_active.append(_xlsx_row([row.get(field, "") for field in CSV_FIELDNAMES]))
                        written_active += len(rows_active)
                        handle_active.flush()

                    if rows_inactive:
                        writer_inactive.writerows(rows_inactive)
                        written_inactive += len(rows_inactive)
                        handle_inactive.flush()

                    if rows_active or rows_inactive:
                        self.log(
                            "Dopisano do CSV: "
                            f"aktywne +{len(rows_active)} (razem {written_active}), "
                            f"nieaktywne +{len(rows_inactive)} (razem {written_inactive})"
                        )

                self._fetch_articles_pages(
                    page_limit=page_limit,
                    stop_event=stop_event,
                    on_items_callback=on_items,
                )
        finally:
            workbook_active.save(output_path_active_xlsx)
            workbook_active.close()

        return written_active, written_inactive

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
                "article_id": str(article.get("id") or article.get("article_id") or "").strip(),
                "name": title,
                "gtin": str(article.get("gtin") or "").strip(),
                "category": str(article.get("category") or "").strip(),
                "brand": str(article.get("brand") or "").strip(),
                "parent_sku": str(article.get("parent_sku") or "").strip(),
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

        self.accounts_path_var = ctk.StringVar(value=resolve_config_path())
        self.output_csv_var = ctk.StringVar(value="")
        self.use_sandbox_var = ctk.BooleanVar(value=False)
        self.safe_mode_var = ctk.BooleanVar(value=False)
        self.rate_limit_var = ctk.StringVar(value=str(DEFAULT_REQUESTS_PER_MINUTE))
        self.page_limit_var = ctk.StringVar(value=str(DEFAULT_PAGE_LIMIT))
        self.workers_var = ctk.StringVar(value=str(DEFAULT_TOTAL_FETCH_WORKERS))

        self.accounts = {}
        self.account_check_vars = {}
        self.accounts_checks_frame = None
        self.is_running = False
        self.stop_event = threading.Event()
        self._log_queue = queue.Queue()

        self._build_ui()
        self._pump_log()
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

    def _build_account_output_paths(self, output_dir, account_name, run_stamp):
        safe_name = self._safe_filename_part(account_name)
        filename_active = f"{safe_name}_aktywne_{run_stamp}.csv"
        filename_inactive = f"{safe_name}_nieaktywne_{run_stamp}.csv"
        filename_active_xlsx = f"{safe_name}_aktywne_{run_stamp}.xlsx"
        return (
            os.path.join(output_dir, filename_active),
            os.path.join(output_dir, filename_inactive),
            os.path.join(output_dir, filename_active_xlsx),
        )

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
            row=0, column=5, padx=(0, 14), sticky="w"
        )

        ctk.CTkLabel(options, text="Rownolegle zapytania", text_color=TEXT).grid(
            row=0, column=6, padx=(0, 8), sticky="w"
        )
        ctk.CTkEntry(options, width=70, textvariable=self.workers_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=0, column=7, padx=(0, 0), sticky="w"
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
        for color in (SUCCESS, ERROR, TEXT_MUTED):
            self.log_box.tag_config(color, foreground=color)

    def _log(self, msg, color=None):
        """Wolane takze z watkow roboczych - tylko wklada wpis do kolejki."""
        stamp = datetime.now().strftime("%H:%M:%S")
        self._log_queue.put((f"[{stamp}] {msg}\n", color))

    def _pump_log(self):
        """Co LOG_FLUSH_MS przepisuje kolejke do widgetu jednym przebiegiem."""
        try:
            pending = []
            while True:
                try:
                    pending.append(self._log_queue.get_nowait())
                except queue.Empty:
                    break

            if pending:
                for line, color in pending:
                    self.log_box.insert("end", line)
                    if color:
                        self.log_box.tag_add(color, "end-2l", "end-1l")

                # Tk spowalnia liniowo z dlugoscia tekstu, wiec trzymamy tylko
                # ogon logu.
                line_count = int(float(self.log_box.index("end-1c").split(".")[0]))
                if line_count > LOG_MAX_LINES:
                    self.log_box.delete("1.0", f"{line_count - LOG_MAX_LINES}.0")

                self.log_box.see("end")
        finally:
            self.after(LOG_FLUSH_MS, self._pump_log)

    def _clear_log(self):
        while True:
            try:
                self._log_queue.get_nowait()
            except queue.Empty:
                break
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

        if not path:
            path = resolve_config_path()
            self.accounts_path_var.set(path)

        if not os.path.exists(path):
            self._refresh_accounts_checkboxes()
            self._log(f"Nie znaleziono pliku kont: {path}", ERROR)
            self._log(
                "Szukano w: " + " | ".join(accounts_config_candidates()),
                TEXT_MUTED,
            )
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

        try:
            workers = int(self.workers_var.get().strip())
            if workers < 1 or workers > MAX_TOTAL_FETCH_WORKERS:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Blad",
                f"Rownolegle zapytania musza byc liczba od 1 do {MAX_TOTAL_FETCH_WORKERS}.",
            )
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
            worker_budget = int(self.workers_var.get().strip())
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
            if set_background_priority(True):
                self._log("Priorytet procesu obnizony - komputer pozostaje uzywalny.")
            wall_start = time.perf_counter()
            cpu_start = time.process_time()
            self._log(f"Limit API: {req_per_min} req/min")
            self._log(f"Page limit: {page_limit}")
            self._log(f"Folder bazowy: {base_output_dir}")
            self._log(f"Folder wyjsciowy: {output_dir}")

            self._log("Pobieranie artykulow z API ze wszystkich kont...")

            if safe_mode:
                account_workers = 1
                page_workers = 1
            else:
                # Kazde konto ma wlasny limit req/min, wiec wszystkie moga jechac
                # rownolegle; budzet dzielimy miedzy nie. Dopiero gdy kont jest
                # wiecej niz budzet, nadmiar czeka w kolejce puli.
                account_workers = min(len(accounts_to_fetch), worker_budget)
                page_workers = max(1, worker_budget // account_workers)
                # Limiter przepuszcza req_per_min zapytan na minute NA KONTO,
                # wiec watki ponad to tylko spia w kolejce. Przy zalozeniu do
                # SLOW_RESPONSE_SECONDS na odpowiedz tyle wystarczy, by nasycic
                # limit - reszta bylaby martwa.
                useful = max(1, int(req_per_min / 60.0 * SLOW_RESPONSE_SECONDS) + 1)
                if page_workers > useful:
                    self._log(
                        f"Ograniczam do {useful} watkow na konto - przy limicie "
                        f"{req_per_min} req/min wiecej i tak czekaloby bezczynnie."
                    )
                    page_workers = useful
            self._log(
                f"Rownolegle zapytania: {account_workers} kont x {page_workers} stron "
                f"= {account_workers * page_workers}"
            )

            def run_one_account(account_name, creds):
                output_active_path, output_inactive_path, output_active_xlsx_path = self._build_account_output_paths(
                    output_dir, account_name, run_stamp
                )
                self._log(f"[{account_name}] Start -> aktywne: {output_active_path}")
                self._log(f"[{account_name}] Start -> nieaktywne: {output_inactive_path}")
                self._log(f"[{account_name}] Start -> aktywne XLSX: {output_active_xlsx_path}")
                client = CDONArticlesClient(
                    merchant_id=creds["merchant_id"],
                    api_token=creds["api_token"],
                    use_sandbox=use_sandbox,
                    requests_per_minute=req_per_min,
                    max_retries=DEFAULT_MAX_RETRIES,
                    backoff_seconds=DEFAULT_BACKOFF_SECONDS,
                    safe_mode=safe_mode,
                    page_burst_workers=page_workers,
                    log_callback=lambda msg: self._log(f"[{account_name}] {msg}"),
                )
                try:
                    written_active, written_inactive = client.export_articles_to_split_csv(
                        output_path_active=output_active_path,
                        output_path_inactive=output_inactive_path,
                        output_path_active_xlsx=output_active_xlsx_path,
                        page_limit=page_limit,
                        stop_event=self.stop_event,
                    )
                finally:
                    client.close()
                self._log(
                    f"[{account_name}] Zakonczono. Aktywne: {written_active}, nieaktywne: {written_inactive}"
                )
                return (
                    account_name,
                    written_active,
                    written_inactive,
                    output_active_path,
                    output_inactive_path,
                    output_active_xlsx_path,
                )

            results = []
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

            total_active_rows = sum(item[1] for item in results)
            total_inactive_rows = sum(item[2] for item in results)
            total_rows = total_active_rows + total_inactive_rows
            total_files = len(results) * 3
            self.progress.set(0.8)
            self.progress.set(1.0)

            wall = time.perf_counter() - wall_start
            cpu = time.process_time() - cpu_start
            self._log(
                f"Czas: {wall / 60:.1f} min, CPU: {cpu / 60:.1f} min "
                f"(srednio {cpu / wall * 100:.0f}% jednego rdzenia). "
                f"Jesli komputer ma byc mniej obciazony, zmniejsz "
                f"'Rownolegle zapytania'."
            )
            self._log(f"Zapisane pliki: {total_files}", SUCCESS)
            self._log(f"Artykuly aktywne: {total_active_rows}", SUCCESS)
            self._log(f"Artykuly nieaktywne: {total_inactive_rows}", SUCCESS)
            self._log(f"Artykuly lacznie: {total_rows}", SUCCESS)

            if self.stop_event.is_set():
                title = "Zatrzymano"
                message = (
                    f"Zatrzymano pobieranie przez uzytkownika.\n"
                    f"Zapisane pliki: {total_files}\n"
                    f"Artykuly aktywne: {total_active_rows}\n"
                    f"Artykuly nieaktywne: {total_inactive_rows}\n"
                    f"Artykuly lacznie: {total_rows}."
                )
            else:
                title = "Gotowe"
                message = (
                    f"Pobrano dane dla {len(results)} kont.\n"
                    f"Artykuly aktywne: {total_active_rows}\n"
                    f"Artykuly nieaktywne: {total_inactive_rows}\n"
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
            set_background_priority(False)
            self.after(0, lambda: self._set_running(False))


if __name__ == "__main__":
    app = App()
    app.mainloop()
