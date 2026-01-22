#!/usr/bin/env python3
"""
Test Script 2: Data Analysis
Biblioteki: pandas, numpy, matplotlib
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt

print("=" * 60)
print("DATA ANALYSIS - Test pandas, numpy, matplotlib")
print("=" * 60)

# Create sample data
data = {
    'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve'],
    'Age': [25, 30, 35, 40, 28],
    'Salary': [50000, 60000, 75000, 90000, 55000]
}

df = pd.DataFrame(data)
print("\n✓ Pandas DataFrame:")
print(df)

# NumPy operations
ages_array = np.array(df['Age'])
print(f"\n✓ NumPy - Średni wiek: {np.mean(ages_array):.1f}")
print(f"✓ NumPy - Odchylenie std: {np.std(ages_array):.2f}")

# Matplotlib plot
plt.figure(figsize=(8, 5))
plt.bar(df['Name'], df['Salary'])
plt.title('Salary by Person')
plt.ylabel('Salary ($)')
plt.savefig('salary_chart.png')
print("\n✓ Matplotlib - Wykres zapisany jako salary_chart.png")

print("\n✓ Wszystkie biblioteki działają!")
