#!/usr/bin/env python3
"""
Skrypt 5: Notatnik z zapisywaniem
"""
import os
from datetime import datetime
from pathlib import Path

def notepad():
    print("=" * 50)
    print("NOTATNIK TEKSTOWY")
    print("=" * 50)
    
    notes_dir = Path.home() / "Desktop" / "Notatki"
    notes_dir.mkdir(exist_ok=True)
    
    while True:
        print("\n1. Nowa notatka")
        print("2. Przeczytaj notatki")
        print("3. Usuń notatkę")
        print("4. Wyjście")
        
        choice = input("\nWybierz opcję (1-4): ").strip()
        
        if choice == "1":
            title = input("Nazwa notatki: ").strip()
            if not title:
                continue
            
            print("Wpisz zawartość (wpisz 'KONIEC' w osobnej linii aby zakończyć):")
            lines = []
            while True:
                line = input()
                if line == "KONIEC":
                    break
                lines.append(line)
            
            filename = notes_dir / f"{title}.txt"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(f"Data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Tytuł: {title}\n")
                f.write("=" * 40 + "\n")
                f.write("\n".join(lines))
            
            print(f"✓ Notatka zapisana: {filename}")
        
        elif choice == "2":
            notes = list(notes_dir.glob("*.txt"))
            if not notes:
                print("Brak notatek.")
                continue
            
            for i, note in enumerate(notes, 1):
                print(f"\n{i}. {note.stem}")
            
            try:
                note_num = int(input("Wybierz numer notatki: ")) - 1
                with open(notes[note_num], "r", encoding="utf-8") as f:
                    print("\n" + f.read())
            except (ValueError, IndexError):
                print("Nieprawidłowy wybór.")
        
        elif choice == "3":
            notes = list(notes_dir.glob("*.txt"))
            if not notes:
                print("Brak notatek.")
                continue
            
            for i, note in enumerate(notes, 1):
                print(f"{i}. {note.stem}")
            
            try:
                note_num = int(input("Wybierz numer notatki do usunięcia: ")) - 1
                notes[note_num].unlink()
                print("✓ Notatka usunięta.")
            except (ValueError, IndexError):
                print("Nieprawidłowy wybór.")
        
        elif choice == "4":
            print("Do widzenia!")
            break

if __name__ == "__main__":
    notepad()
