#!/usr/bin/env python3
"""
Test Script 6: Cryptography Utils
Biblioteki: cryptography, bcrypt
"""
from cryptography.fernet import Fernet
import bcrypt

print("=" * 60)
print("CRYPTO UTILS - Test cryptography i bcrypt")
print("=" * 60)

# Symmetric encryption with Fernet
key = Fernet.generate_key()
cipher = Fernet(key)

message = b"Secret message for testing"
encrypted = cipher.encrypt(message)
decrypted = cipher.decrypt(encrypted)

print("✓ Cryptography (Fernet):")
print(f"  Wiadomość: {message.decode()}")
print(f"  Zaszyfrowane: {encrypted[:40]}...")
print(f"  Odszyfrowane: {decrypted.decode()}")

# Password hashing with bcrypt
password = b"MySecretPassword123"
hashed = bcrypt.hashpw(password, bcrypt.gensalt())
check = bcrypt.checkpw(password, hashed)

print("\n✓ Bcrypt (Password hashing):")
print(f"  Hash: {hashed.decode()[:50]}...")
print(f"  Weryfikacja: {'✓ OK' if check else '✗ Błąd'}")

print("\n✓ Wszystkie biblioteki kryptograficzne działają!")
