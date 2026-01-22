#!/usr/bin/env python3
"""
Test Script 3: REST API Client
Biblioteki: httpx, pydantic
"""
import httpx
from pydantic import BaseModel
from typing import Optional

print("=" * 60)
print("REST API CLIENT - Test httpx i pydantic")
print("=" * 60)

class User(BaseModel):
    id: int
    name: str
    username: str
    email: str
    phone: Optional[str] = None

try:
    with httpx.Client(timeout=10) as client:
        response = client.get("https://jsonplaceholder.typicode.com/users/1")
        response.raise_for_status()
        
        user = User(**response.json())
        
        print(f"✓ HTTPX - Pobrano dane użytkownika")
        print(f"✓ Pydantic - Walidacja danych OK")
        print(f"\nDane użytkownika:")
        print(f"  ID: {user.id}")
        print(f"  Imię: {user.name}")
        print(f"  Username: {user.username}")
        print(f"  Email: {user.email}")
        
except Exception as e:
    print(f"✗ Błąd: {e}")

print("\n✓ Test zakończony!")
