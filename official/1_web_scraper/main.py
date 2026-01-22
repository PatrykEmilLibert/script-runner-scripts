#!/usr/bin/env python3
"""
Test Script 1: Web Scraper
Biblioteki: requests, beautifulsoup4, lxml
"""
import requests
from bs4 import BeautifulSoup

print("=" * 60)
print("WEB SCRAPER - Test bibliotek requests i beautifulsoup4")
print("=" * 60)

try:
    url = "https://httpbin.org/html"
    response = requests.get(url, timeout=5)
    soup = BeautifulSoup(response.text, 'lxml')
    
    print(f"✓ Połączono z: {url}")
    print(f"✓ Status code: {response.status_code}")
    print(f"✓ Tytuł strony: {soup.title.string if soup.title else 'Brak'}")
    print(f"✓ BeautifulSoup działa!")
    print(f"✓ lxml parser działa!")
    
except Exception as e:
    print(f"✗ Błąd: {e}")

print("\nTest zakończony!")
