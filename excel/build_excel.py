# -*- coding: utf-8 -*-
"""
Формирование Excel-макета лабораторного журнала градуировки:
  - Лист "Данные"        — исходный журнал измерений (сырые данные)
  - Лист "Градуировка"   — 4 таблицы (Fe/Cu/Zn/Cr): среднее по концентрации
                           через AVERAGEIFS, k/b/R2 через SLOPE/INTERCEPT/RSQ
  - Лист "Расчёт_пробы"  — калькулятор: ввод металла + измеренной A ->
                           расчёт концентрации по формуле C=(A-b)/k через INDEX/MATCH
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
RAW_CSV = ROOT / "data" / "raw_measurements.csv"
OUT_XLSX = ROOT / "excel" / "calibration_workbook.xlsx"

df = pd.read_csv(RAW_CSV, encoding="utf-8-sig")

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=13, color="1F3864")
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="808080")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")  # ячейки для заполнения пользователем
RESULT_FILL = PatternFill("solid", fgColor="E2EFDA")
BOLD = Font(name=FONT_NAME, bold=True)
BASE_FONT = Font(name=FONT_NAME)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()

# =========================================================================
# 1. Лист "Данные" — сырой журнал измерений
# =========================================================================
ws_raw = wb.active
ws_raw.title = "Данные"

headers = ["Металл", "Сессия", "Концентрация_мг_дм3", "Повтор", "Оптическая_плотность_А"]
for j, h in enumerate(headers, start=1):
    c = ws_raw.cell(row=1, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    c.border = BORDER

for i, row in enumerate(df.itertuples(index=False), start=2):
    ws_raw.cell(row=i, column=1, value=row[0]).font = BASE_FONT
    ws_raw.cell(row=i, column=2, value=row[1]).font = BASE_FONT
    ws_raw.cell(row=i, column=3, value=row[2]).font = BASE_FONT
    ws_raw.cell(row=i, column=4, value=row[3]).font = BASE_FONT
    ws_raw.cell(row=i, column=5, value=row[4]).font = BASE_FONT

for col, width in zip("ABCDE", [10, 9, 20, 9, 24]):
    ws_raw.column_dimensions[col].width = width
ws_raw.freeze_panes = "A2"
n_raw = ws_raw.max_row

# =========================================================================
# 2. Лист "Градуировка" — 4 блока таблиц с формулами
# =========================================================================
ws_cal = wb.create_sheet("Градуировка")
ws_cal.sheet_view.showGridLines = False

standards = [0.00, 0.05, 0.10, 0.25, 0.50, 1.00, 2.00]
metals = ["Fe", "Cu", "Zn", "Cr"]
metal_titles = {
    "Fe": "Железо (Fe)", "Cu": "Медь (Cu)",
    "Zn": "Цинк (Zn)", "Cr": "Хром (Cr)",
}

# каждому металлу — блок из 3 колонок с отступом в 1 колонку
block_start_cols = {"Fe": 1, "Cu": 5, "Zn": 9, "Cr": 13}
metal_coeff_cells = {}  # metal -> (k_cell, b_cell, r2_cell) для использования на след. листе

for metal in metals:
    c0 = block_start_cols[metal]
    col_conc = get_column_letter(c0)
    col_mean = get_column_letter(c0 + 1)

    title_cell = ws_cal.cell(row=1, column=c0, value=metal_titles[metal])
    title_cell.font = TITLE_FONT
    ws_cal.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 1)

    h1 = ws_cal.cell(row=2, column=c0, value="Концентрация,\nмг/дм³")
    h2 = ws_cal.cell(row=2, column=c0 + 1, value="Среднее A\n(по журналу)")
    for h in (h1, h2):
        h.font = HEADER_FONT
        h.fill = HEADER_FILL
        h.alignment = Alignment(horizontal="center", wrap_text=True)
        h.border = BORDER

    first_data_row = 3
    for i, conc in enumerate(standards):
        r = first_data_row + i
        cc = ws_cal.cell(row=r, column=c0, value=conc)
        cc.font = BASE_FONT
        cc.border = BORDER
        cc.number_format = "0.00"
        cc.alignment = Alignment(horizontal="center")

        # AVERAGEIFS считает среднее по журналу "Данные" для данного металла и концентрации
        formula = (
            f'=AVERAGEIFS(Данные!$E$2:$E${n_raw},'
            f'Данные!$A$2:$A${n_raw},"{metal}",'
            f'Данные!$C$2:$C${n_raw},{col_conc}{r})'
        )
        mc = ws_cal.cell(row=r, column=c0 + 1, value=formula)
        mc.font = BASE_FONT
        mc.border = BORDER
        mc.number_format = "0.0000"

    last_data_row = first_data_row + len(standards) - 1

    # Коэффициенты регрессии МНК
    r_k = last_data_row + 2
    r_b = r_k + 1
    r_r2 = r_b + 1

    label_font = BOLD
    ws_cal.cell(row=r_k, column=c0, value="k (тангенс угла накл.)").font = label_font
    ws_cal.cell(row=r_b, column=c0, value="b (свободный член)").font = label_font
    ws_cal.cell(row=r_r2, column=c0, value="R²").font = label_font

    range_y = f"{col_mean}{first_data_row}:{col_mean}{last_data_row}"
    range_x = f"{col_conc}{first_data_row}:{col_conc}{last_data_row}"

    k_cell = ws_cal.cell(row=r_k, column=c0 + 1, value=f"=SLOPE({range_y},{range_x})")
    b_cell = ws_cal.cell(row=r_b, column=c0 + 1, value=f"=INTERCEPT({range_y},{range_x})")
    r2_cell = ws_cal.cell(row=r_r2, column=c0 + 1, value=f"=RSQ({range_y},{range_x})")
    for cell in (k_cell, b_cell, r2_cell):
        cell.font = BOLD
        cell.fill = RESULT_FILL
        cell.border = BORDER
        cell.number_format = "0.0000"

    metal_coeff_cells[metal] = (
        f"Градуировка!${col_mean}${r_k}",
        f"Градуировка!${col_mean}${r_b}",
    )

    for col, width in zip((col_conc, col_mean), (14, 16)):
        ws_cal.column_dimensions[col].width = width
    ws_cal.column_dimensions[get_column_letter(c0 + 2)].width = 3  # разделитель

note = ws_cal.cell(
    row=15, column=1,
    value="Примечание: A — оптическая плотность (среднее по параллельным измерениям и сессиям градуировки), рассчитано из листа «Данные». "
          "k, b, R² рассчитаны методом наименьших квадратов через встроенные функции Excel SLOPE / INTERCEPT / RSQ.",
)
note.font = NOTE_FONT
ws_cal.merge_cells(start_row=15, start_column=1, end_row=17, end_column=14)
note.alignment = Alignment(wrap_text=True, vertical="top")

# =========================================================================
# 3. Лист "Расчёт_пробы" — калькулятор концентрации неизвестной пробы
# =========================================================================
ws_calc = wb.create_sheet("Расчёт_пробы")
ws_calc.sheet_view.showGridLines = False

title = ws_calc.cell(row=1, column=1, value="Расчёт концентрации по градуировочному графику")
title.font = TITLE_FONT
ws_calc.merge_cells("A1:D1")

labels = ["Металл", "Измеренная A (проба)", "k (из листа «Градуировка»)",
          "b (из листа «Градуировка»)", "Концентрация C, мг/дм³"]
for i, lab in enumerate(labels, start=3):
    lc = ws_calc.cell(row=i, column=1, value=lab)
    lc.font = BOLD
    lc.border = BORDER

# Ячейка выбора металла (выпадающий список) — заполняет пользователь
metal_cell = ws_calc.cell(row=3, column=2, value="Fe")
metal_cell.fill = INPUT_FILL
metal_cell.font = BASE_FONT
metal_cell.border = BORDER
dv = DataValidation(type="list", formula1='"Fe,Cu,Zn,Cr"', allow_blank=False)
ws_calc.add_data_validation(dv)
dv.add(metal_cell)

# Измеренная оптическая плотность — заполняет пользователь
a_cell = ws_calc.cell(row=4, column=2, value=0.45)
a_cell.fill = INPUT_FILL
a_cell.font = BASE_FONT
a_cell.border = BORDER
a_cell.number_format = "0.0000"

# k и b подтягиваются по выбранному металлу через INDEX/MATCH
k_formula = (
    '=INDEX(Градуировка!$B$12:$B$12,1)'  # placeholder, replaced below
)

# Строим корректные ссылки через MATCH по подписанным диапазонам на листе Градуировка
# Соберём горизонтальный список меток металлов и ссылок на их k/b ячейки на скрытых служебных строках
ws_calc.cell(row=20, column=1, value="Служебная таблица (не удалять)").font = NOTE_FONT
ws_calc.cell(row=21, column=1, value="Металл").font = BOLD
ws_calc.cell(row=21, column=2, value="k").font = BOLD
ws_calc.cell(row=21, column=3, value="b").font = BOLD
for i, metal in enumerate(metals, start=22):
    ws_calc.cell(row=i, column=1, value=metal).font = BASE_FONT
    k_ref, b_ref = metal_coeff_cells[metal]
    ws_calc.cell(row=i, column=2, value=f"={k_ref}").font = BASE_FONT
    ws_calc.cell(row=i, column=3, value=f"={b_ref}").font = BASE_FONT

k_out = ws_calc.cell(row=5, column=2,
                      value='=INDEX($B$22:$B$25,MATCH($B$3,$A$22:$A$25,0))')
b_out = ws_calc.cell(row=6, column=2,
                      value='=INDEX($C$22:$C$25,MATCH($B$3,$A$22:$A$25,0))')
for cell in (k_out, b_out):
    cell.font = BASE_FONT
    cell.border = BORDER
    cell.number_format = "0.0000"

c_result = ws_calc.cell(row=7, column=2, value="=(B4-B6)/B5")
c_result.font = Font(name=FONT_NAME, bold=True, size=12)
c_result.fill = RESULT_FILL
c_result.border = BORDER
c_result.number_format = "0.0000"

for col, width in zip("ABCD", [30, 16, 16, 16]):
    ws_calc.column_dimensions[col].width = width

note2 = ws_calc.cell(
    row=9, column=1,
    value="Введите металл (B3) и измеренную оптическую плотность пробы (B4) — "
          "концентрация (B7) рассчитывается автоматически по формуле C=(A-b)/k.",
)
note2.font = NOTE_FONT
ws_calc.merge_cells("A9:D10")
note2.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT_XLSX)
print("Сохранено:", OUT_XLSX)
