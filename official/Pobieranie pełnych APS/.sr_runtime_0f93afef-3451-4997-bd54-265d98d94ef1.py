#!/usr/bin/env python3
import argparse
import csv as _csv
import io
import re
import sys
from typing import Iterable, List, Set, Optional, Dict, Any

import pandas as pd
import requests

# ---------- Defaults you can override via CLI ----------
DEFAULT_API_URL = "https://sm-prods.com/api/files/static"
DEFAULT_PREFIX = "https://sm-prods.com/api/download/static/"
DEFAULT_OUTPUT_CSV = "combined_full_aps2.csv"
# NEW DEFAULT: Output for file metadata
DEFAULT_METADATA_OUT = "file_records_metadata.csv" 
DEFAULT_URLS_OUT = None  # or e.g. "feed_urls.csv" to also dump the discovered URLs
HTTP_TIMEOUT = 60

# Inline URL blacklist (full URLs) — these feeds will be skipped entirely
URL_BLACKLIST: Set[str] = set()

# ---------- URL discovery (ported from your second script) ----------
def _iter_strings(obj) -> Iterable[str]:
    """Yield all strings nested anywhere inside lists/dicts."""
    if obj is None:
        return
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, list):
        for x in obj:
            yield from _iter_strings(x)
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from _iter_strings(v)

def _build_full_url(path: str, prefix: str) -> str:
    """Turn a relative 'feeds/...' or '/feeds/...' path into a full URL using prefix."""
    p = path.lstrip("/")
    if p.startswith("feeds/"):
        p = p.split("feeds/", 1)[1]
    if prefix.endswith("/"):
        return prefix + p
    return prefix + "/" + p

def get_feed_urls(api_url: str, prefix: str, url_blacklist: Set[str]) -> List[str]:
    """Fetch the feed list JSON and return full URLs for entries ending with 'add_per_sku.csv'."""
    try:
        r = requests.get(api_url, timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"ERROR: Could not fetch feed list from {api_url}: {e}", file=sys.stderr)
        return []

    candidates = [
        s for s in _iter_strings(data)
        if isinstance(s, str) and "." in s and "://" not in s
    ]

    filtered = []
    for s in candidates:
        base = s.rsplit("/", 1)[-1].lower()
        if "add_per_sku.csv" in base:
            full_url = _build_full_url(s, prefix)
            if full_url not in url_blacklist:
                filtered.append(full_url)

    # De-duplicate while preserving order
    seen: Set[str] = set()
    result: List[str] = []
    for u in filtered:
        if u not in seen:
            seen.add(u)
            result.append(u)
    return result

def load_blacklist_file(path: str) -> Set[str]:
    """Read a newline-delimited list of full URLs to exclude."""
    bl: Set[str] = set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                u = line.strip()
                if u and not u.startswith("#"):
                    bl.add(u)
    except FileNotFoundError:
        print(f"WARNING: Blacklist file not found: {path}. Continuing without it.", file=sys.stderr)
    return bl

# ---------- Helpers ----------
def _derive_filename_from_url(url: str) -> str:
    """
    From a URL like .../static/<path>/<name>_add_per_sku.csv
    return 'name'. If pattern not matched, use the last component sans suffix.
    """
    m = re.search(r'/static/(.*?)(?:/)?([^/]+?)_add_per_sku\.csv$', url)
    if m:
        return m.group(2)
    filename = re.sub(r'\.csv$', '', url.rstrip('/').rsplit('/', 1)[-1], flags=re.IGNORECASE)
    return filename.replace('_add_per_sku', '')

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lowercase & strip column names, collapse internal spaces/underscores for robust matching."""
    def norm(s: str) -> str:
        s = s.strip().lower()
        s = re.sub(r'[\s]+', '_', s)
        return s
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]
    return df

# ---------- Processing ----------
# Updated return type to include the list of metadata dictionaries
def process_and_combine_full_files_selected_columns(urls: List[str]) -> tuple[pd.DataFrame, List[Dict[str, Any]]]:
    """
    Downloads, processes, and combines ALL rows from each CSV URL,
    keeping only two columns:
      - sku: combined as '{filename}_{original_sku}'
      - aps_weight
    
    Returns the combined DataFrame and a list of metadata dicts.
    """
    all_dfs: List[pd.DataFrame] = []
    # NEW: List to hold metadata for each successfully processed file
    file_metadata: List[Dict[str, Any]] = [] 
    
    total_urls = len(urls)
    print(f"Starting download and processing of {total_urls} files...\n")

    for i, url in enumerate(urls, 1):
        filename: Optional[str] = None # Define outside try for error logging
        try:
            print(f"({i}/{total_urls}) Downloading and processing: {url}")
            resp = requests.get(url, timeout=HTTP_TIMEOUT)
            resp.raise_for_status()

            csv_content = io.StringIO(resp.text)

            # Read entire file; keep sku as string to preserve leading zeros.
            df = pd.read_csv(
                csv_content,
                sep=';',
                on_bad_lines='skip',
                dtype={'sku': 'string', 'SKU': 'string'},  # hint; robust renaming happens below
            )

            df = _normalize_columns(df)

            # Validate expected columns exist after normalization
            missing = []
            for col in ['sku', 'aps_weight']:
                if col not in df.columns:
                    missing.append(col)
            if missing:
                print(f"  Skipping file: missing required columns {missing}")
                print("-" * 30)
                continue

            # Deriving the filename here is crucial for both output files
            filename = _derive_filename_from_url(url)

            # Build combined sku as '{filename}_{original_sku}'
            # Ensure original_sku is string
            original_sku = df['sku'].astype('string')
            combined_sku = original_sku.map(lambda x: f"{filename}_{x}" if pd.notna(x) else x)

            # Select only the two required columns in the specified order
            out = pd.DataFrame({
                'sku': combined_sku,
                'aps_weight': df['aps_weight']
            })
            
            # --- NEW METADATA COLLECTION ---
            record_count = len(out)
            file_metadata.append({
                'filename': filename,
                'record_count': record_count,
            })
            # -------------------------------

            # Optional: coerce aps_weight to numeric (keep as is if you prefer)
            # out['aps_weight'] = pd.to_numeric(out['aps_weight'], errors='coerce')

            all_dfs.append(out)
            print(f"  Successfully processed '{filename}' with {record_count:,} rows.")
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {url}: {e}")
        except pd.errors.ParserError as e:
            print(f"Error parsing CSV from {url}: {e}")
        except Exception as e:
            # Include filename in error message if it was successfully derived
            target = f"'{filename}' ({url})" if filename else url
            print(f"Unexpected error for {target}: {e}")
        print("-" * 30)

    print("\nAll files processed. Combining data...")
    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        print(f"Data combination complete. Total rows: {len(combined_df):,}")
        return combined_df, file_metadata
    else:
        print("No data was successfully processed.")
        # Return an empty DataFrame and an empty list of metadata
        return pd.DataFrame(columns=['sku', 'aps_weight']), file_metadata

# ---------- CLI glue ----------
def main():
    parser = argparse.ArgumentParser(
        description="Fetch add_per_sku.csv feeds from an API and combine all rows, keeping only sku and aps_weight. The output sku is '{filename}_{original_sku}'."
    )
    parser.add_argument("--api-url", default=DEFAULT_API_URL, help="API endpoint returning feed metadata JSON")
    parser.add_argument("--prefix", default=DEFAULT_PREFIX, help="Base prefix used to build full feed URLs")
    parser.add_argument("--blacklist-file", default=None, help="Optional path to a newline-delimited list of FULL URLs to skip")
    parser.add_argument("--out", default=DEFAULT_OUTPUT_CSV, help="Path to save the combined CSV output (sku + aps_weight)")
    # NEW ARGUMENT for metadata output
    parser.add_argument("--metadata-out", default=DEFAULT_METADATA_OUT, help="Path to save the file metadata CSV (filename + record_count)")
    parser.add_argument("--dump-urls", default=DEFAULT_URLS_OUT, help="Optional path to also save discovered URLs to a CSV (one per row)")
    parser.add_argument("--print-urls", action="store_true", help="Also print URLs to stdout")
    args = parser.parse_args()

    # Merge inline and file blacklists
    url_blacklist = set(URL_BLACKLIST)
    if args.blacklist_file:
        url_blacklist |= load_blacklist_file(args.blacklist_file)

    # Discover URLs from the API
    urls = get_feed_urls(args.api_url, args.prefix, url_blacklist)

    if args.dump_urls is not None:
        try:
            with open(args.dump_urls, "w", newline="", encoding="utf-8") as f:
                w = _csv.writer(f)
                w.writerow(["url"])
                for u in urls:
                    w.writerow([u])
            print(f"Wrote {len(urls)} URLs to {args.dump_urls}")
        except OSError as e:
            print(f"ERROR: Failed to write URL CSV '{args.dump_urls}': {e}", file=sys.stderr)

    if args.print_urls:
        for u in urls:
            print(u)

    # Process & combine ALL rows, and also get the file metadata
    combined, metadata = process_and_combine_full_files_selected_columns(urls)

    # Save combined CSV
    if not combined.empty:
        try:
            combined.to_csv(args.out, index=False)
            print(f"\nSuccessfully saved combined data to {args.out}")
        except OSError as e:
            print(f"ERROR: Could not write output '{args.out}': {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print("Final output is empty. No combined data file was saved.")

    # --- NEW: Save metadata CSV ---
    if metadata:
        metadata_df = pd.DataFrame(metadata)
        try:
            metadata_df.to_csv(args.metadata_out, index=False)
            print(f"Successfully saved file metadata to {args.metadata_out}")
        except OSError as e:
            print(f"ERROR: Could not write metadata output '{args.metadata_out}': {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print("No metadata was generated. No metadata file was saved.")
    # -----------------------------

if __name__ == "__main__":
    main()