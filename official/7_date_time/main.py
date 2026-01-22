#!/usr/bin/env python3
"""
Test Script 7: Date & Time Processing
Biblioteki: arrow, python-dateutil, pytz
"""
import arrow
from dateutil import parser, relativedelta
import pytz

print("=" * 60)
print("DATE & TIME - Test arrow, dateutil, pytz")
print("=" * 60)

# Arrow - modern datetime library
now = arrow.now()
print("✓ Arrow:")
print(f"  Teraz: {now.format('YYYY-MM-DD HH:mm:ss')}")
print(f"  Za tydzień: {now.shift(weeks=+1).humanize()}")
print(f"  Tydzień temu: {now.shift(weeks=-1).humanize()}")

# Dateutil - parsing and relative deltas
date_str = "2024-12-25 15:30:00"
parsed = parser.parse(date_str)
future = parsed + relativedelta.relativedelta(months=+3, days=+10)

print("\n✓ Python-dateutil:")
print(f"  Sparsowano: {parsed}")
print(f"  +3 miesiące, +10 dni: {future}")

# Pytz - timezone handling
utc_time = arrow.utcnow()
warsaw_tz = pytz.timezone('Europe/Warsaw')
warsaw_time = utc_time.to(warsaw_tz)

print("\n✓ Pytz:")
print(f"  UTC: {utc_time.format('HH:mm:ss')}")
print(f"  Warszawa: {warsaw_time.format('HH:mm:ss')}")

print("\n✓ Wszystkie biblioteki czasu działają!")
