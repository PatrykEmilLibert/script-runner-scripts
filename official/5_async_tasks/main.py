#!/usr/bin/env python3
"""
Test Script 5: Async Tasks
Biblioteki: aiohttp, asyncio (stdlib)
"""
import asyncio
import aiohttp

print("=" * 60)
print("ASYNC TASKS - Test aiohttp")
print("=" * 60)

async def fetch_url(session, url):
    async with session.get(url) as response:
        return await response.text()

async def main():
    urls = [
        'https://httpbin.org/delay/1',
        'https://httpbin.org/delay/1',
        'https://httpbin.org/delay/1',
    ]
    
    print("Rozpoczynam równoległe zapytania HTTP...")
    
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_url(session, url) for url in urls]
        results = await asyncio.gather(*tasks)
        
        print(f"✓ Wykonano {len(results)} zapytań równolegle")
        print(f"✓ Aiohttp działa!")
        print(f"✓ Asyncio działa!")

try:
    asyncio.run(main())
    print("\n✓ Test zakończony!")
except Exception as e:
    print(f"✗ Błąd: {e}")
