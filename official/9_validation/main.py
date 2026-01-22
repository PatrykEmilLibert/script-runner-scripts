#!/usr/bin/env python3
"""
Test Script 9: Data Validation
Biblioteki: marshmallow, validators, cerberus
"""
from marshmallow import Schema, fields, ValidationError
import validators
from cerberus import Validator

print("=" * 60)
print("VALIDATION - Test marshmallow, validators, cerberus")
print("=" * 60)

# Marshmallow - object serialization/validation
class UserSchema(Schema):
    name = fields.Str(required=True)
    email = fields.Email(required=True)
    age = fields.Int()

schema = UserSchema()
user_data = {'name': 'John Doe', 'email': 'john@example.com', 'age': 30}

try:
    result = schema.load(user_data)
    print("✓ Marshmallow:")
    print(f"  Walidacja OK: {result}")
except ValidationError as e:
    print(f"  Błąd walidacji: {e}")

# Validators - simple validators
email = "test@example.com"
url = "https://example.com"

print("\n✓ Validators:")
print(f"  Email '{email}': {'✓ Valid' if validators.email(email) else '✗ Invalid'}")
print(f"  URL '{url}': {'✓ Valid' if validators.url(url) else '✗ Invalid'}")

# Cerberus - schema validation
schema_def = {
    'name': {'type': 'string', 'minlength': 3, 'required': True},
    'age': {'type': 'integer', 'min': 0, 'max': 150}
}

v = Validator(schema_def)
document = {'name': 'Alice', 'age': 25}

print("\n✓ Cerberus:")
if v.validate(document):
    print(f"  Dokument valid: {document}")
else:
    print(f"  Błędy: {v.errors}")

print("\n✓ Wszystkie biblioteki walidacji działają!")
