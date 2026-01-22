#!/usr/bin/env python3
"""
Test Script 8: CLI Tools
Biblioteki: click, rich, colorama
"""
import click
from rich.console import Console
from rich.table import Table
from colorama import init, Fore, Style

init(autoreset=True)

print("=" * 60)
print("CLI TOOLS - Test click, rich, colorama")
print("=" * 60)

# Click - command line interfaces
@click.command()
@click.option('--name', default='World', help='Name to greet')
def greet(name):
    return f"Hello, {name}!"

print("\n✓ Click:")
print(f"  {greet.callback(name='ScriptRunner')}")

# Rich - beautiful terminal output
console = Console()
table = Table(title="Test Table")
table.add_column("ID", style="cyan")
table.add_column("Name", style="magenta")
table.add_row("1", "Item One")
table.add_row("2", "Item Two")

print("\n✓ Rich:")
console.print(table)

# Colorama - colored terminal output
print("\n✓ Colorama:")
print(f"{Fore.GREEN}  Zielony tekst{Style.RESET_ALL}")
print(f"{Fore.RED}  Czerwony tekst{Style.RESET_ALL}")
print(f"{Fore.YELLOW}  Żółty tekst{Style.RESET_ALL}")

print("\n✓ Wszystkie biblioteki CLI działają!")
