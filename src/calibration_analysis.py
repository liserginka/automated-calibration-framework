# -*- coding: utf-8 -*-
"""
Портфолио-кейс: построение градуировочных графиков и расчёт
метрологических характеристик методом наименьших квадратов (МНК).

Задача:
    По результатам фотометрических измерений стандартных растворов
    построить градуировочные зависимости "оптическая плотность —
    концентрация" для Fe, Cu, Zn, Cr, рассчитать уравнение прямой
    A = k*C + b, тангенс угла наклона (k) и коэффициент детерминации R^2.

Вход:
    raw_measurements.csv — журнал измерений (сессии х параллельные пробы)

Выход:
    calibration_summary.csv  — сводная таблица коэффициентов по металлам
    graduation_curves.png    — итоговые градуировочные графики
"""

from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
RESULTS_DIR = ROOT / "results"
RESULTS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------
# 1. Загрузка и агрегация данных
# ---------------------------------------------------------------
df = pd.read_csv(DATA_DIR / "raw_measurements.csv", encoding="utf-8-sig")

# Согласно ПНДФ, перед построением градуировочного графика параллельные
# определения усредняются для каждой концентрации в рамках сессии,
# а итоговая точка на графике - среднее по всем сессиям.
agg = (
    df.groupby(["Металл", "Концентрация_мг_дм3"])["Оптическая_плотность_А"]
    .agg(["mean", "std", "count"])
    .reset_index()
    .rename(columns={"mean": "A_среднее", "std": "A_СКО", "count": "n"})
)

metals = ["Fe", "Cu", "Zn", "Cr"]
metal_names_ru = {"Fe": "железо (Fe)", "Cu": "медь (Cu)", "Zn": "цинк (Zn)", "Cr": "хром (Cr)"}
colors = {"Fe": "#B5482A", "Cu": "#C08628", "Zn": "#3E6E8E", "Cr": "#5C8A4E"}

# ---------------------------------------------------------------
# 2. Регрессия методом наименьших квадратов для каждого металла
# ---------------------------------------------------------------
results = []
regressions = {}

for metal in metals:
    sub = agg[agg["Металл"] == metal].sort_values("Концентрация_мг_дм3")
    x = sub["Концентрация_мг_дм3"].values
    y = sub["A_среднее"].values

    # МНК: scipy.stats.linregress реализует классический метод наименьших квадратов
    lr = stats.linregress(x, y)
    k, b, r, p_value, se = lr.slope, lr.intercept, lr.rvalue, lr.pvalue, lr.stderr
    r2 = r ** 2

    regressions[metal] = (k, b, r2, x, y)
    results.append({
        "Металл": metal_names_ru[metal],
        "Тангенс угла наклона k": round(k, 4),
        "Свободный член b": round(b, 5),
        "R^2": round(r2, 5),
        "Стд. ошибка наклона": round(se, 5),
        "Число точек (усреднённых)": len(x),
    })

summary = pd.DataFrame(results)
summary.to_csv(RESULTS_DIR / "calibration_summary.csv", index=False, encoding="utf-8-sig")
print(summary.to_string(index=False))

# ---------------------------------------------------------------
# 3. Визуализация: 2x2 сетка градуировочных графиков
# ---------------------------------------------------------------
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 11,
    "axes.edgecolor": "#444444",
    "axes.labelcolor": "#222222",
})

fig, axes = plt.subplots(2, 2, figsize=(12, 9))
fig.suptitle(
    "Градуировочные графики определения концентраций металлов в воде\n"
    "(фотометрический метод, расчёт МНК)",
    fontsize=15, fontweight="bold", y=1.0,
)

for ax, metal in zip(axes.flat, metals):
    k, b, r2, x, y = regressions[metal]
    color = colors[metal]

    # экспериментальные точки (средние по параллельным/сессиям)
    sub = df[df["Металл"] == metal]
    ax.scatter(
        sub["Концентрация_мг_дм3"], sub["Оптическая_плотность_А"],
        s=18, color=color, alpha=0.25, label="Единичные измерения", zorder=2,
    )
    ax.scatter(
        x, y, s=60, color=color, edgecolor="black", linewidth=0.6,
        label="Средние значения", zorder=3,
    )

    # линия регрессии
    x_line = np.linspace(0, x.max(), 100)
    y_line = k * x_line + b
    ax.plot(x_line, y_line, color="black", linewidth=1.6, zorder=4, label="Линия МНК")

    sign = "+" if b >= 0 else "-"
    eq_text = f"A = {k:.3f}·C {sign} {abs(b):.4f}\nR² = {r2:.4f}"
    ax.text(
        0.05, 0.92, eq_text, transform=ax.transAxes,
        fontsize=11, fontweight="bold", va="top", ha="left",
        bbox=dict(boxstyle="round,pad=0.35", facecolor="white", edgecolor=color, linewidth=1.2),
    )

    ax.set_title(metal_names_ru[metal], fontsize=13, fontweight="bold", color=color)
    ax.set_xlabel("Концентрация, мг/дм³")
    ax.set_ylabel("Оптическая плотность, А")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend(loc="lower right", fontsize=8, framealpha=0.9)

plt.tight_layout(rect=[0, 0, 1, 0.96])
out_png = RESULTS_DIR / "graduation_curves.png"
plt.savefig(out_png, dpi=200, bbox_inches="tight")
print(f"\nГрафик сохранён: {out_png}")
