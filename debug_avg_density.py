#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отладка расчёта средних данных и интенсивности
"""

import json
import re

# Загружаем JSON файл
json_path = 'reports/Молодняки_3_20260309_1402.json'
with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

page_data = data.get('page_data', {})
radius = float(data.get('project_data', {}).get('address', {}).get('radius', 1.78))
plot_area_ha = 3.14159 * (radius ** 2) / 10000

print("=" * 60)
print("ОТЛАДКА РАСЧЁТА СРЕДНИХ ДАННЫХ")
print("=" * 60)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га ({3.14159 * radius ** 2:.2f} м2)")

def parse_care_subject_density(care_text):
    if not care_text:
        return 0
    care_text = care_text.strip().upper()
    matches = re.findall(r'(\d+(?:\.\d+)?)([А-ЯA-Z]+)', care_text)
    if not matches:
        return 0
    total_density = 0
    for number_str, breed_code in matches:
        try:
            density = float(number_str)
            total_density += density
        except ValueError:
            continue
    return total_density * 1000

# Считаем количество строк (площадок)
total_rows = len([row for page in page_data.values() for row in page if any(cell for cell in row[:3] if cell)])
print(f"\nКоличество площадок (строк): {total_rows}")

# Расчёт общей густоты ПО ПЛОЩАДКАМ
print("\n=== ГУСТОТА ПО ПЛОЩАДКАМ ===")
plot_densities = []

for page_num, page_rows in page_data.items():
    for row in page_rows:
        if len(row) >= 4 and row[3]:
            plot_density = 0
            breeds_text = row[3]
            if breeds_text:
                try:
                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                    for breed_info in breeds_list:
                        if isinstance(breed_info, dict):
                            if breed_info.get('type') == 'coniferous':
                                do_05 = breed_info.get('do_05', 0)
                                _05_15 = breed_info.get('05_15', 0)
                                bolee_15 = breed_info.get('bolee_15', 0)
                                total_trees = do_05 + _05_15 + bolee_15
                                density = total_trees / plot_area_ha if plot_area_ha > 0 else 0
                                plot_density += density
                                print(f"  {breed_info.get('name', 'Н/Д')}: {total_trees} дер. = {density:.1f} шт/га")
                            else:
                                density_value = breed_info.get('density', 0)
                                density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                                plot_density += density
                                print(f"  {breed_info.get('name', 'Н/Д')}: {density_value} дер. = {density:.1f} шт/га")
                except (json.JSONDecodeError, TypeError):
                    pass
            
            if plot_density > 0:
                plot_densities.append(plot_density)
                print(f"  -> Густота площадки: {plot_density:.1f} шт/га")

# СРЕДНЯЯ густота по площадкам
if plot_densities:
    avg_density_by_plots = sum(plot_densities) / len(plot_densities)
    print(f"\nСредняя густота по площадкам: {avg_density_by_plots:.1f} шт/га")
    print(f"  (сумма {sum(plot_densities):.1f} / кол-во {len(plot_densities)})")

# Расчёт оставаемой густоты
print("\n=== ПРЕДМЕТ УХОДА ===")
care_densities = []

for page_num, page_rows in page_data.items():
    for row in page_rows:
        if len(row) >= 3 and row[2]:
            care_text = row[2].strip()
            if care_text:
                remaining_density = parse_care_subject_density(care_text)
                if remaining_density > 0:
                    care_densities.append(remaining_density)
                    print(f"  '{care_text}' = {remaining_density:.0f} шт/га")

if care_densities:
    avg_remaining_density = sum(care_densities) / len(care_densities)
    print(f"\nСредняя оставаемая густота: {avg_remaining_density:.1f} шт/га")

# Расчёт интенсивности
print("\n=== РАСЧЁТ ИНТЕНСИВНОСТИ ===")
if plot_densities and care_densities:
    # ВАЖНО: используем СРЕДНЮЮ густоту по площадкам, а не сумму!
    avg_overall_density = avg_density_by_plots
    avg_remaining = avg_remaining_density
    
    intensity = ((avg_overall_density - avg_remaining) / avg_overall_density) * 100
    
    print(f"Средняя исходная густота: {avg_overall_density:.1f} шт/га")
    print(f"Средняя оставаемая густота: {avg_remaining:.1f} шт/га")
    print(f"Интенсивность рубки: {intensity:.1f}%")
    
    if intensity > 35:
        print(f"\n⚠️ ВНИМАНИЕ: Интенсивность {intensity:.1f}% слишком высокая для осветления!")
        print(f"   Рекомендуется 25-35%")

print("\n" + "=" * 60)
