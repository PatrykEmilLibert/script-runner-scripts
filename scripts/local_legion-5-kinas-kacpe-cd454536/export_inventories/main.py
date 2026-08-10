#!/usr/bin/env python3
import argparse
import csv
import json
import sys
from typing import List, Dict

try:
    import requests
except ImportError as e:
    sys.stderr.write("This script requires the 'requests' package. Install it with: pip install requests\n")
    raise

# --- SCRIPT CONFIGURATION ---
# IMPORTANT: Replace "YOUR_BASELINKER_TOKEN_HERE" with your actual BaseLinker API token.
BASELINKER_TOKEN = "1426-10044-8TA3V6MQU7M54479DKBFTXXZR88EJDGBHF8D1XWRP991H8Z4N99ZF9JOKTFX26RI"

# ZMIENIONA ŚCIEŻKA DO PLIKU (dodano 'r' przed stringiem dla poprawności ścieżki w Windows)
OUTPUT_FILE = r"C:\Users\kacpe\Desktop\inventories.csv"
# -----------------------------

API_URL = "https://api.baselinker.com/connector.php"

def fetch_inventories(token: str) -> List[Dict]:
    """
    Calls BaseLinker getInventories and returns the list of inventory dicts.
    """
    headers = {
        "X-BLToken": token.strip(),
    }
    # getInventories takes no parameters; send an empty JSON object as a string
    data = {
        "method": "getInventories",
        "parameters": json.dumps({}),
    }

    resp = requests.post(API_URL, headers=headers, data=data, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:500]}")
    try:
        payload = resp.json()
    except json.JSONDecodeError:
        raise RuntimeError("Response was not valid JSON.")

    if payload.get("status") != "SUCCESS":
        # BaseLinker usually includes 'error_code' and 'error_message' on failure
        err_msg = payload.get("error_message") or payload
        raise RuntimeError(f"API error: {err_msg}")

    inv = payload.get("inventories")
    if not isinstance(inv, list):
        raise RuntimeError("Unexpected API response format: 'inventories' is missing or not a list.")
    return inv

def write_csv(inventories: List[Dict], out_path: str) -> None:
    """
    Writes inventory_id and name to CSV.
    """
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["inventory_id", "name"])
        for item in inventories:
            writer.writerow([item.get("inventory_id", ""), item.get("name", "")])

def main():
    """
    Main function to fetch and write inventory data.
    """
    # Check if the token has been updated from the placeholder
    if BASELINKER_TOKEN == "YOUR_BASELINKER_TOKEN_HERE":
        sys.stderr.write("Error: Please update the 'BASELINKER_TOKEN' variable in the script with your actual BaseLinker API token.\n")
        sys.exit(1)

    try:
        inventories = fetch_inventories(BASELINKER_TOKEN)
        write_csv(inventories, OUTPUT_FILE)
        print(f"Done. Wrote {len(inventories)} inventories to {OUTPUT_FILE}")
    except Exception as e:
        sys.stderr.write(f"Error: {e}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()