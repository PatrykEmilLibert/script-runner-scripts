#!/usr/bin/env python3
"""
Test Script 10: Scientific Computing
Biblioteki: scipy, scikit-learn, sympy
"""
import numpy as np
from scipy import stats
from sklearn.linear_model import LinearRegression
from sympy import symbols, diff, integrate, sin, cos

print("=" * 60)
print("SCIENTIFIC COMPUTING - Test scipy, sklearn, sympy")
print("=" * 60)

# SciPy - statistical analysis
data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
mean = np.mean(data)
std = np.std(data)
median = np.median(data)

print("✓ SciPy (stats):")
print(f"  Średnia: {mean}")
print(f"  Std dev: {std:.2f}")
print(f"  Mediana: {median}")

# Scikit-learn - machine learning
X = np.array([[1], [2], [3], [4], [5]])
y = np.array([2, 4, 5, 4, 5])
model = LinearRegression()
model.fit(X, y)
prediction = model.predict([[6]])

print("\n✓ Scikit-learn (Linear Regression):")
print(f"  Model wytrenowany")
print(f"  Predykcja dla X=6: {prediction[0]:.2f}")
print(f"  Współczynnik: {model.coef_[0]:.2f}")

# SymPy - symbolic mathematics
x = symbols('x')
expr = x**2 + 2*x + 1
derivative = diff(expr, x)
integral = integrate(sin(x), x)

print("\n✓ SymPy (symbolic math):")
print(f"  Wyrażenie: {expr}")
print(f"  Pochodna: {derivative}")
print(f"  Całka sin(x): {integral}")

print("\n✓ Wszystkie biblioteki naukowe działają!")
