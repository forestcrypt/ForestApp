#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест исправленного расчёта средних данных
"""

import json
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

print("=" * 70)
print("ТЕСТ ИСПРАВЛЕННОГО РАСЧЁТА")
print("=" * 70)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га")

# Симулируем исправленный расчёт в get_total_data_from_db
breeds_data = {}

for page_num, page_rows in page_data.items():
    for row_idx, row in enumerate(page_rows):
        if len(row) >= 4 and row[3]:
            try:
                breeds_list = json.loads(row[3]) if isinstance(row[3], str) else []
                
                for breed_info in breeds_list:
                    if isinstance(breed_info, dict):
                        breed_name = breed_info.get('name', '').strip()
                        breed_type = breed_info.get('type', 'deciduous')
                        
                        # Расчёт густоты
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            total_trees = do_05 + _05_15 + bolee_15
                            density = total_trees / plot_area_ha if plot_area_ha > 0 else 0
                            
                            if bolee_15 > 0:
                                height = 2.0
                            elif _05_15 > 0:
                                height = 1.0
                            elif do_05 > 0:
                                height = 0.3
                            else:
                                height = breed_info.get('height', 0) or 0
                        else:
                            density_value = breed_info.get('density', 0)
                            density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                            height = breed_info.get('height', 0) or 0
                        
                        diameter = breed_info.get('diameter', 0) or 0
                        age = breed_info.get('age', 0) or 0
                        
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {'plots': []}
                        
                        breeds_data[breed_name]['plots'].append({
                            'density': density,
                            'height': height,
                            'diameter': diameter,
                            'age': age
                        })
            except json.JSONDecodeError:
                pass

# Исправленный расчёт (как в новом коде)
plot_densities = []
plot_height_sums = []
plot_height_counts = []
plot_diameter_sums = []
plot_diameter_counts = []
plot_age_sums = []
plot_age_counts = []

for breed_name, data in breeds_data.items():
    if data['plots']:
        for i, p in enumerate(data['plots']):
            while i >= len(plot_densities):
                plot_densities.append(0)
                plot_height_sums.append(0)
                plot_height_counts.append(0)
                plot_diameter_sums.append(0)
                plot_diameter_counts.append(0)
                plot_age_sums.append(0)
                plot_age_counts.append(0)

            plot_densities[i] += p['density']
            
            if p['height'] > 0:
                plot_height_sums[i] += p['height']
                plot_height_counts[i] += 1
            
            if p.get('diameter', 0) > 0:
                plot_diameter_sums[i] += p['diameter']
                plot_diameter_counts[i] += 1
            
            if p['age'] > 0:
                plot_age_sums[i] += p['age']
                plot_age_counts[i] += 1

# Рассчитываем средние значения
avg_overall_density = sum(plot_densities) / len(plot_densities) if plot_densities else 0

plot_avg_heights = []
plot_avg_diameters = []
plot_avg_ages = []

for i in range(len(plot_densities)):
    if plot_height_counts[i] > 0:
        plot_avg_heights.append(plot_height_sums[i] / plot_height_counts[i])
    if plot_diameter_counts[i] > 0:
        plot_avg_diameters.append(plot_diameter_sums[i] / plot_diameter_counts[i])
    if plot_age_counts[i] > 0:
        plot_avg_ages.append(plot_age_sums[i] / plot_age_counts[i])

avg_overall_height = sum(plot_avg_heights) / len(plot_avg_heights) if plot_avg_heights else 0
avg_overall_diameter = sum(plot_avg_diameters) / len(plot_avg_diameters) if plot_avg_diameters else 0
avg_overall_age = sum(plot_avg_ages) / len(plot_avg_ages) if plot_avg_ages else 0

print(f"\nПлощадок: {len(plot_densities)}")
for i in range(len(plot_densities)):
    print(f"\nПлощадка {i+1}:")
    print(f"  Общая густота: {plot_densities[i]:.1f} шт/га")
    if plot_height_counts[i] > 0:
        h_avg = plot_height_sums[i] / plot_height_counts[i]
        print(f"  Средняя высота: {h_avg:.1f} м (сумма {plot_height_sums[i]:.1f} / {plot_height_counts[i]})")
    if plot_diameter_counts[i] > 0:
        d_avg = plot_diameter_sums[i] / plot_diameter_counts[i]
        print(f"  Средний диаметр: {d_avg:.1f} см (сумма {plot_diameter_sums[i]:.1f} / {plot_diameter_counts[i]})")
    if plot_age_counts[i] > 0:
        a_avg = plot_age_sums[i] / plot_age_counts[i]
        print(f"  Средний возраст: {a_avg:.1f} лет (сумма {plot_age_sums[i]:.1f} / {plot_age_counts[i]})")

print(f"\n" + "=" * 70)
print("ИТОГОВЫЕ СРЕДНИЕ ЗНАЧЕНИЯ (исправлено):")
print("-" * 70)
print(f"  Средняя густота: {avg_overall_density:.1f} шт/га")
print(f"  Средняя высота: {avg_overall_height:.1f} м")
print(f"  Средний диаметр: {avg_overall_diameter:.1f} см")
print(f"  Средний возраст: {avg_overall_age:.1f} лет")

print(f"\n" + "=" * 70)
print("ОЖИДАЕМЫЕ ЗНАЧЕНИЯ:")
print("-" * 70)
print(f"  Средняя густота: 34157.8 шт/га")
print(f"  Средняя высота: 3.3 м ((2.0 + 6.0 + 2.0) / 3)")
print(f"  Средний диаметр: 5.0 см ((4.0 + 5.0 + 6.0) / 3)")
print(f"  Средний возраст: 11.3 лет ((10 + 12 + 12) / 3)")
print("=" * 70)

if abs(avg_overall_density - 34157.8) < 1:
    print("\n✓ ГУСТОТА ВЕРНО!")
else:
    print(f"\n✗ ГУСТОТА НЕВЕРНО: {avg_overall_density:.1f} вместо 34157.8")

if abs(avg_overall_height - 3.3) < 0.1:
    print("✓ ВЫСОТА ВЕРНО!")
else:
    print(f"✗ ВЫСОТА НЕВЕРНО: {avg_overall_height:.1f} вместо 3.3")

if abs(avg_overall_diameter - 5.0) < 0.1:
    print("✓ ДИАМЕТР ВЕРНО!")
else:
    print(f"✗ ДИАМЕТР НЕВЕРНО: {avg_overall_diameter:.1f} вместо 5.0")

if abs(avg_overall_age - 11.3) < 0.1:
    print("✓ ВОЗРАСТ ВЕРНО!")
else:
    print(f"✗ ВОЗРАСТ НЕВЕРНО: {avg_overall_age:.1f} вместо 11.3")

print("=" * 70)
