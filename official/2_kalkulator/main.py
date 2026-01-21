#!/usr/bin/env python3
"""
Skrypt 2: Kalkulator tekstowy
"""

def calculator():
    print("=" * 50)
    print("KALKULATOR TEKSTOWY")
    print("=" * 50)
    
    while True:
        try:
            print("\nOperacje dostępne: +, -, *, /, **, %")
            expression = input("Wpisz działanie (lub 'wyjście' aby zakończyć): ").strip()
            
            if expression.lower() == 'wyjście':
                print("Do widzenia!")
                break
            
            result = eval(expression)
            print(f"Wynik: {result}")
            
        except Exception as e:
            print(f"Błąd: {e}")

if __name__ == "__main__":
    calculator()
