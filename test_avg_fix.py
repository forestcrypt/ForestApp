#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест расчёта средних данных после исправления
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

print("=" * 60)
print("ТЕСТ РАСЧЁТА СРЕДНИХ ДАННЫХ")
print("=" * 60)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га")

# Считаем данные по площадкам как в исправленном get_total_data_from_db
plot_densities = []
plot_heights = []
plot_ages = []
plot_diameters = []

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
                            height = 2.0 if bolee_15 > 0 else (1.0 if _05_15 > 0 else (0.3 if do_05 > 0 else 0))
                        else:
                            density_value = breed_info.get('density', 0)
                            density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                            height = breed_info.get('height', 0) or 0
                        
                        diameter = breed_info.get('diameter', 0) or 0
                        age = breed_info.get('age', 0) or 0
                        
                        # Сбор данных по породе
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

# Теперь считаем средние по площадкам как в исправленном коде
for breed_name, data in breeds_data.items():
    if data['plots']:
        for i, p in enumerate(data['plots']):
            while i >= len(plot_densities):
                plot_densities.append(0)
                plot_heights.append(0)
                plot_ages.append(0)
                plot_diameters.append(0)
            
            plot_densities[i] += p['density']
            if p['height'] > 0:
                plot_heights[i] = p['height']
            if p['age'] > 0:
                plot_ages[i] = p['age']
            if p.get('diameter', 0) > 0:
                plot_diameters[i] = p['diameter']

print(f"\nПлощадки: {len(plot_densities)}")
for i in range(len(plot_densities)):
    print(f"  Площадка {i+1}: густота={plot_densities[i]:.1f}, высота={plot_heights[i]:.1f}, диаметр={plot_diameters[i]:.1f}, возраст={plot_ages[i]:.1f}")

# Рассчитываем средние значения по площадкам
avg_overall_density = sum(plot_densities) / len(plot_densities) if plot_densities else 0
avg_overall_height = sum(plot_heights) / len(plot_heights) if plot_heights else 0
avg_overall_age = sum(plot_ages) / len(plot_ages) if plot_ages else 0
avg_overall_diameter = sum(plot_diameters) / len(plot_diameters) if plot_diameters else 0

print(f"\n=== СРЕДНИЕ ЗНАЧЕНИЯ (исправлено) ===")
print(f"Средняя густота: {avg_overall_density:.1f} шт/га")
print(f"Средняя высота: {avg_overall_height:.1f} м")
print(f"Средний диаметр: {avg_overall_diameter:.1f} см")
print(f"Средний возраст: {avg_overall_age:.1f} лет")

# Для сравнения: старый способ (по породам)
all_densities_old = []
for breed_name, data in breeds_data.items():
    if data['plots']:
        all_densities_old.extend([p['density'] for p in data['plots'] if p['density'] > 0])

avg_density_old = sum(all_densities_old) / len(all_densities_old) if all_densities_old else 0

print(f"\n=== СРЕДНИЕ ЗНАЧЕНИЯ (старый способ - НЕВЕРНО) ===")
print(f"Средняя густота (по породам): {avg_density_old:.1f} шт/га")
print(f"  (сумма {sum(all_densities_old):.1f} / кол-во {len(all_densities_old)})")

print("\n" + "=" * 60)
print("ВЫВОД: Новый способ считает среднюю густоту по площадкам!")
print(f"  Старый: {avg_density_old:.1f} шт/га (неверно, т.к. это среднее по породам)")
print(f"  Новый: {avg_overall_density:.1f} шт/га (верно, т.к. это общая густота на площадку)")
print("=" * 60)
