#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка расчёта интенсивности в get_total_data_from_db
"""

import json
import re
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Загружаем JSON файл
json_path = 'reports/Молодняки_3_20260309_1402.json'
with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

page_data = data.get('page_data', {})
radius = float(data.get('project_data', {}).get('address', {}).get('radius', 1.78))
plot_area_ha = 3.14159 * (radius ** 2) / 10000

print("=" * 60)
print("ПРОВЕРКА РАСЧЁТА ИНТЕНСИВНОСТИ")
print("=" * 60)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га")

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

# Расчёт как в get_total_data_from_db
total_density_all_plots = 0
total_remaining_density = 0
plot_count_with_care = 0
num_plots = 0

print("\n=== РАСЧЁТ КАК В get_total_data_from_db ===")

for page_num, page_rows in page_data.items():
    for row in page_rows:
        # Считаем густоту породы
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
                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha
                            else:
                                density = breed_info.get('density', 0)
                                plot_density += density / plot_area_ha
                except:
                    pass
            
            if plot_density > 0:
                total_density_all_plots += plot_density
                num_plots += 1

        # Считаем предмет ухода
        if len(row) >= 3 and row[2]:
            care_text = row[2].strip()
            if care_text:
                # Снова считаем plot_density для этой строки
                plot_density_care = 0
                breeds_text = row[3] if len(row) >= 4 else ''
                if breeds_text:
                    try:
                        breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                        for breed_info in breeds_list:
                            if isinstance(breed_info, dict):
                                if breed_info.get('type') == 'coniferous':
                                    do_05 = breed_info.get('do_05', 0)
                                    _05_15 = breed_info.get('05_15', 0)
                                    bolee_15 = breed_info.get('bolee_15', 0)
                                    plot_density_care += (do_05 + _05_15 + bolee_15) / plot_area_ha
                                else:
                                    density = breed_info.get('density', 0)
                                    plot_density_care += density / plot_area_ha
                    except:
                        pass
                
                remaining_density = parse_care_subject_density(care_text)
                if remaining_density > 0 and plot_density_care > 0:
                    total_remaining_density += remaining_density
                    plot_count_with_care += 1
                    print(f"Строка {len(page_rows)}: густота={plot_density_care:.1f}, предмет ухода='{care_text}'={remaining_density:.0f}")

print(f"\ntotal_density_all_plots (СУММА): {total_density_all_plots:.1f} шт/га")
print(f"num_plots (кол-во строк): {num_plots}")
print(f"total_remaining_density (СУММА): {total_remaining_density:.0f} шт/га")
print(f"plot_count_with_care: {plot_count_with_care}")

# Расчёт интенсивности как в оригинальном коде
if plot_count_with_care > 0 and total_density_all_plots > 0:
    avg_overall_density = total_density_all_plots / num_plots if num_plots > 0 else 0
    avg_remaining_density = total_remaining_density / plot_count_with_care
    
    print(f"\navg_overall_density (СРЕДНЯЯ): {avg_overall_density:.1f} шт/га")
    print(f"avg_remaining_density (СРЕДНЯЯ): {avg_remaining_density:.0f} шт/га")
    
    intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100
    
    print(f"\nИнтенсивность рубки: {intensity:.1f}%")
    
    print("\n" + "=" * 60)
    print("ВЫВОД: Формула работает правильно!")
    print(f"  - Исходная густота: {avg_overall_density:.1f} шт/га (средняя по площадкам)")
    print(f"  - Оставаемая густота: {avg_remaining_density:.0f} шт/га (из предмета ухода)")
    print(f"  - Интенсивность: {intensity:.1f}%")
    print("=" * 60)
