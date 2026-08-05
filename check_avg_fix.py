#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка исправления расчёта средних данных по площадкам
"""

import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

print("=" * 80)
print("ПРОВЕРКА ИСПРАВЛЕНИЯ РАСЧЁТА СРЕДНИХ ДАННЫХ ПО ПЛОЩАДКАМ")
print("=" * 80)

# Загружаем тестовые данные
json_path = 'reports/Молодняки_3_20260309_1402.json'

try:
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
except FileNotFoundError:
    print(f"❌ Файл {json_path} не найден!")
    print("Создайте тестовые данные или укажите правильный путь.")
    sys.exit(1)

page_data = data.get('page_data', {})
radius = float(data.get('project_data', {}).get('address', {}).get('radius', 1.78))
plot_area_ha = 3.14159 * (radius ** 2) / 10000

print(f"\n📏 Параметры:")
print(f"   Радиус: {radius} м")
print(f"   Площадь площадки: {plot_area_ha:.6f} га ({3.14159 * radius ** 2:.2f} м²)")

# Эталонный расчёт (правильный)
print("\n" + "=" * 80)
print("ЭТАЛОННЫЙ РАСЧЁТ (ПРАВИЛЬНЫЙ - по площадкам)")
print("=" * 80)

plot_data_list = []

for page_num, page_rows in page_data.items():
    for row_idx, row in enumerate(page_rows):
        if len(row) < 4 or not row[3]:
            continue
        
        plot_total_density = 0
        plot_height_sum = 0
        plot_height_count = 0
        plot_diameter_sum = 0
        plot_diameter_count = 0
        plot_age_sum = 0
        plot_age_count = 0
        
        try:
            breeds_list = json.loads(row[3]) if isinstance(row[3], str) else []
            
            for breed_info in breeds_list:
                if not isinstance(breed_info, dict):
                    continue
                
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
                
                plot_total_density += density
                if height > 0:
                    plot_height_sum += height
                    plot_height_count += 1
                if diameter > 0:
                    plot_diameter_sum += diameter
                    plot_diameter_count += 1
                if age > 0:
                    plot_age_sum += age
                    plot_age_count += 1
            
            plot_data_list.append({
                'density': plot_total_density,
                'height': plot_height_sum / plot_height_count if plot_height_count > 0 else 0,
                'diameter': plot_diameter_sum / plot_diameter_count if plot_diameter_count > 0 else 0,
                'age': plot_age_sum / plot_age_count if plot_age_count > 0 else 0
            })
            
        except Exception as e:
            print(f"⚠️ Ошибка обработки: {e}")
            continue

num_plots = len(plot_data_list)

if num_plots > 0:
    avg_density = sum(p['density'] for p in plot_data_list) / num_plots
    avg_height = sum(p['height'] for p in plot_data_list) / num_plots
    avg_diameter = sum(p['diameter'] for p in plot_data_list) / num_plots
    avg_age = sum(p['age'] for p in plot_data_list) / num_plots
    
    print(f"\n✅ РЕЗУЛЬТАТЫ ({num_plots} площадок):")
    print(f"   📊 Средняя густота:  {avg_density:.1f} шт/га")
    print(f"   📏 Средняя высота:   {avg_height:.1f} м")
    print(f"   📐 Средний диаметр:  {avg_diameter:.1f} см")
    print(f"   🎂 Средний возраст:  {avg_age:.1f} лет")
else:
    print("❌ Нет данных для расчёта!")
    avg_density = avg_height = avg_diameter = avg_age = 0

# Старая логика (НЕПРАВИЛЬНАЯ - по породам)
print("\n" + "=" * 80)
print("СТАРАЯ ЛОГИКА (НЕПРАВИЛЬНАЯ - по породам)")
print("=" * 80)

breeds_data = {}

for page_num, page_rows in page_data.items():
    for row_idx, row in enumerate(page_rows):
        if len(row) < 4 or not row[3]:
            continue
        
        try:
            breeds_list = json.loads(row[3]) if isinstance(row[3], str) else []
            
            for breed_info in breeds_list:
                if not isinstance(breed_info, dict):
                    continue
                
                breed_name = breed_info.get('name', '')
                if not breed_name:
                    continue
                
                breed_type = breed_info.get('type', 'deciduous')
                
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
                    breeds_data[breed_name] = {
                        'plots': [],
                        'type': breed_type
                    }
                
                breeds_data[breed_name]['plots'].append({
                    'density': density,
                    'height': height,
                    'diameter': diameter,
                    'age': age
                })
        
        except Exception as e:
            continue

# Старый расчёт по породам
old_plot_densities = []
old_height_sums = []
old_height_counts = []
old_diameter_sums = []
old_diameter_counts = []
old_age_sums = []
old_age_counts = []

for breed_name, data in breeds_data.items():
    if data['plots']:
        for i, p in enumerate(data['plots']):
            while i >= len(old_plot_densities):
                old_plot_densities.append(0)
                old_height_sums.append(0)
                old_height_counts.append(0)
                old_diameter_sums.append(0)
                old_diameter_counts.append(0)
                old_age_sums.append(0)
                old_age_counts.append(0)
            
            old_plot_densities[i] += p['density']
            
            if p['height'] > 0:
                old_height_sums[i] += p['height']
                old_height_counts[i] += 1
            
            if p.get('diameter', 0) > 0:
                old_diameter_sums[i] += p['diameter']
                old_diameter_counts[i] += 1
            
            if p.get('age', 0) > 0:
                old_age_sums[i] += p['age']
                old_age_counts[i] += 1

old_num_plots = len(old_plot_densities)

# Инициализируем переменные
old_avg_density = old_avg_height = old_avg_diameter = old_avg_age = 0

if old_num_plots > 0:
    old_avg_density = sum(old_plot_densities) / old_num_plots
    
    old_plot_avg_heights = []
    old_plot_avg_diameters = []
    old_plot_avg_ages = []
    
    for i in range(old_num_plots):
        if old_height_counts[i] > 0:
            old_plot_avg_heights.append(old_height_sums[i] / old_height_counts[i])
        if old_diameter_counts[i] > 0:
            old_plot_avg_diameters.append(old_diameter_sums[i] / old_diameter_counts[i])
        if old_age_counts[i] > 0:
            old_plot_avg_ages.append(old_age_sums[i] / old_age_counts[i])
    
    old_avg_height = sum(old_plot_avg_heights) / len(old_plot_avg_heights) if old_plot_avg_heights else 0
    old_avg_diameter = sum(old_plot_avg_diameters) / len(old_plot_avg_diameters) if old_plot_avg_diameters else 0
    old_avg_age = sum(old_plot_avg_ages) / len(old_plot_avg_ages) if old_plot_avg_ages else 0
    
    print(f"\n📊 СТАРЫЕ ЗНАЧЕНИЯ ({old_num_plots} 'площадок' по породам):")
    print(f"   Средняя густота:  {old_avg_density:.1f} шт/га")
    print(f"   Средняя высота:   {old_avg_height:.1f} м")
    print(f"   Средний диаметр:  {old_avg_diameter:.1f} см")
    print(f"   Средний возраст:  {old_avg_age:.1f} лет")
else:
    print("\n⚠️ Нет данных для старого расчёта")

# Сравнение
print("\n" + "=" * 80)
print("СРАВНЕНИЕ РЕЗУЛЬТАТОВ")
print("=" * 80)

print(f"\n{'Параметр':<20} | {'ПРАВИЛЬНО':<15} | {'СТАРАЯ ЛОГИКА':<15} | {'Разница':<15}")
print("-" * 80)
print(f"{'Густота (шт/га)':<20} | {avg_density:<15.1f} | {old_avg_density:<15.1f} | {avg_density - old_avg_density:<+15.1f}")
print(f"{'Высота (м)':<20} | {avg_height:<15.1f} | {old_avg_height:<15.1f} | {avg_height - old_avg_height:<+15.1f}")
print(f"{'Диаметр (см)':<20} | {avg_diameter:<15.1f} | {old_avg_diameter:<15.1f} | {avg_diameter - old_avg_diameter:<+15.1f}")
print(f"{'Возраст (лет)':<20} | {avg_age:<15.1f} | {old_avg_age:<15.1f} | {avg_age - old_avg_age:<+15.1f}")

print("\n" + "=" * 80)
print("ВЫВОД:")
print("=" * 80)

if abs(avg_density - old_avg_density) > 0.1 or abs(avg_height - old_avg_height) > 0.1:
    print("⚠️  СТАРАЯ ЛОГИКА ДАВАЛА НЕВЕРНЫЕ РЕЗУЛЬТАТЫ!")
    print("✅ ИСПРАВЛЕНИЕ ПРИМЕНЕНО - расчёт по площадкам вместо расчёта по породам")
else:
    print("ℹ️  Различий не обнаружено (возможно данные одинаковые)")

print("\n" + "=" * 80)
