#!/usr/bin/env python3
"""
Skrypt 1: Wyświetla informacje o systemie
"""
import platform
import psutil
import os

print("=" * 50)
print("INFORMACJE O SYSTEMIE")
print("=" * 50)

print(f"System operacyjny: {platform.system()} {platform.release()}")
print(f"Wersja: {platform.version()}")
print(f"Procesor: {platform.processor()}")
print(f"Architektura: {platform.architecture()[0]}")
print(f"Nazwa komputera: {platform.node()}")
print(f"Python: {platform.python_version()}")

print("\n" + "=" * 50)
print("ZASOBY SYSTEMOWE")
print("=" * 50)

cpu_percent = psutil.cpu_percent(interval=1)
memory = psutil.virtual_memory()
disk = psutil.disk_usage('/')

print(f"Użycie CPU: {cpu_percent}%")
print(f"RAM razem: {memory.total / (1024**3):.2f} GB")
print(f"RAM używany: {memory.used / (1024**3):.2f} GB ({memory.percent}%)")
print(f"Dysk razem: {disk.total / (1024**3):.2f} GB")
print(f"Dysk używany: {disk.used / (1024**3):.2f} GB ({disk.percent}%)")

input("\nNaciśnij Enter aby zamknąć...")
