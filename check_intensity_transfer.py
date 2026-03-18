#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка передачи интенсивности в fill_our_template.py
"""

import json
import tempfile
import os
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

# Расчёт интенсивности как в get_total_data_from_db
import re

total_density_all_plots = 0
total_remaining_density = 0
plot_count_with_care = 0
num_plots = 0

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
                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha
                            else:
                                density = breed_info.get('density', 0)
                                plot_density += density / plot_area_ha
                except:
                    pass
            if plot_density > 0:
                total_density_all_plots += plot_density
                num_plots += 1

        if len(row) >= 3 and row[2]:
            care_text = row[2].strip()
            if care_text:
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

# Расчёт интенсивности
intensity = 25.0  # По умолчанию
if plot_count_with_care > 0 and total_density_all_plots > 0:
    avg_overall_density = total_density_all_plots / num_plots if num_plots > 0 else 0
    avg_remaining_density = total_remaining_density / plot_count_with_care
    
    if avg_overall_density > 0:
        intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100

print("=" * 60)
print("ПОДГОТОВЛЕННЫЕ ДАННЫЕ")
print("=" * 60)
print(f"Интенсивность рубки (расчётная): {intensity:.1f}%")

# Теперь проверим, что происходит в fill_our_template.py
print("\n=== ПРОВЕРКА fill_our_template.py ===")

# Читаем файл и ищем где используется intensity
with open('fill_our_template.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Ищем строки с intensity
import re
matches = re.findall(r'intensity.*?=.*?[^%]\d+\.?\d*', content)
if matches:
    print("Найдено в файле:")
    for m in matches[:5]:
        print(f"  {m}")

# Проверяем, откуда берётся intensity
if "self.total_data.get('intensity'" in content:
    print("\n✓ intensity берётся из self.total_data.get('intensity')")
else:
    print("\n✗ Не найдено получение intensity из total_data")

print("\n" + "=" * 60)
