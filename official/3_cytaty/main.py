#!/usr/bin/env python3
"""
Skrypt 3: Generator losowych cytatów
"""
import random

quotes = [
    "Jedynym sposobem robienia świetnej pracy jest kochanie tego, co robisz. - Steve Jobs",
    "Przyszłość należy do tych, którzy wierzą w piękno swoich marzeń. - Eleanor Roosevelt",
    "Nie imituj! Bądź sobą. - Steve Jobs",
    "Sukces to suma małych wysiłków powtarzanych dzień po dniu. - Robert Collier",
    "Najlepszy czas na sadzenie drzewa był 20 lat temu. Drugi najlepszy czas to teraz. - Chiński przysłów",
    "Nie czekaj na okazję. Stwórz ją. - Unknown",
    "Każdy master był kiedyś początkującym. - Unknown",
    "Rzeczy, które są warte robienia, są warte robienia źle. - Unknown",
]

def main():
    print("=" * 60)
    print("GENERATOR CYTATÓW")
    print("=" * 60)
    
    while True:
        print("\n" + random.choice(quotes))
        choice = input("\nChcesz kolejny cytat? (t/n): ").strip().lower()
        if choice != 't':
            print("Dziękujemy za użycie! Do widzenia!")
            break

if __name__ == "__main__":
    main()
