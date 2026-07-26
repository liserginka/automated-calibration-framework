# -*- coding: utf-8 -*-
"""
Генерация синтетических данных градуировки для 4 металлов (Fe, Cu, Zn, Cr)
по методике фотометрического определения (аналог ПНДФ).

Логика:
- 7 стандартных растворов с известной концентрацией (мг/дм3)
- 5 сессий градуировки (разные дни / партии стандартов) — имитация реальной
  лабораторной практики повторных градуировок
- 3 параллельных измерения на точку (требование ПНДФ по параллельным определениям)
=> 7 * 5 * 3 = 105 измерений на металл, 420 измерений суммарно

Итоговый набор данных имитирует реальный журнал Excel лаборатории.
"""

from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

np.random.seed(42)

# Стандартные концентрации, мг/дм3
concentrations = np.array([0.0, 0.05, 0.10, 0.25, 0.50, 1.00, 2.00])

# Параметры "истинной" градуировочной зависимости A = k*C + b для каждого металла
# k - коэффициент чувствительности (тангенс угла наклона), b - фон/холостая проба
metals_params = {
    "Fe": {"k": 0.842, "b": 0.012, "noise": 0.006},
    "Cu": {"k": 1.185, "b": 0.008, "noise": 0.007},
    "Zn": {"k": 0.615, "b": 0.015, "noise": 0.005},
    "Cr": {"k": 1.430, "b": 0.010, "noise": 0.009},
}

n_sessions = 5
n_replicates = 3

rows = []
for metal, p in metals_params.items():
    for session in range(1, n_sessions + 1):
        # небольшой сдвиг фона и чувствительности между сессиями (реальный дрейф прибора)
        session_drift_k = np.random.normal(1.0, 0.012)
        session_drift_b = np.random.normal(0.0, 0.003)
        for c in concentrations:
            true_signal = p["k"] * session_drift_k * c + (p["b"] + session_drift_b)
            for rep in range(1, n_replicates + 1):
                noise = np.random.normal(0, p["noise"])
                signal = max(true_signal + noise, 0)  # оптическая плотность не может быть < 0
                rows.append({
                    "Металл": metal,
                    "Сессия": session,
                    "Концентрация_мг_дм3": round(c, 3),
                    "Повтор": rep,
                    "Оптическая_плотность_А": round(signal, 4),
                })

df = pd.DataFrame(rows)
out_path = DATA_DIR / "raw_measurements.csv"
df.to_csv(out_path, index=False, encoding="utf-8-sig")
print(f"Сохранено: {out_path}")
print(df.shape)
print(df.head(10))
