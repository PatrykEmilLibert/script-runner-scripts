#!/usr/bin/env python3
"""
Skrypt 4: Organizator plików
Sortuje pliki z pulpitu do folderów wg typu
"""
import os
import shutil
from pathlib import Path

def organize_desktop():
    print("=" * 50)
    print("ORGANIZATOR PLIKÓW")
    print("=" * 50)
    
    desktop = Path.home() / "Desktop"
    
    # Zdefiniuj kategorie plików
    categories = {
        "Obrazy": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".ico"],
        "Dokumenty": [".pdf", ".doc", ".docx", ".txt", ".xlsx", ".xls"],
        "Archiwa": [".zip", ".rar", ".7z", ".tar", ".gz"],
        "Audio": [".mp3", ".wav", ".flac", ".m4a"],
        "Video": [".mp4", ".mkv", ".avi", ".mov"],
        "Programy": [".exe", ".msi"],
    }
    
    moved_files = 0
    
    for file in desktop.iterdir():
        if file.is_file():
            file_ext = file.suffix.lower()
            
            for category, extensions in categories.items():
                if file_ext in extensions:
                    dest_folder = desktop / category
                    dest_folder.mkdir(exist_ok=True)
                    
                    dest_file = dest_folder / file.name
                    
                    if not dest_file.exists():
                        shutil.move(str(file), str(dest_file))
                        print(f"✓ Przeniesiono: {file.name} → {category}/")
                        moved_files += 1
                    break
    
    print(f"\nRazem przeniesiono: {moved_files} pliki(ów)")
    input("Naciśnij Enter aby zamknąć...")

if __name__ == "__main__":
    organize_desktop()
