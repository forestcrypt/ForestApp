#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка работы get_total_data_from_db()
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Загружаем данные из JSON
import json

json_path = 'reports/Молодняки_3_20260309_1402.json'

with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

page_data = data.get('page_data', {})
radius = float(data.get('radius', 1.78))
plot_area_ha = 3.14159 * (radius ** 2) / 10000

print("=" * 80)
print("ПРОВЕРКА РАСЧЁТА ДИАМЕТРА ПО ПЛОЩАДКАМ")
print("=" * 80)
print(f"Радиус: {radius} м, Площадь: {plot_area_ha:.6f} га")

# Считаем как в get_total_data_from_db()
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
            
            print(f"\nПлощадка {len(plot_data_list) + 1}:")
            
            for breed_info in breeds_list:
                if not isinstance(breed_info, dict):
                    continue
                
                breed_name = breed_info.get('name', 'Н/Д')
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
                
                print(f"  {breed_name}: d={diameter}, h={height}, density={density:.1f}")
                
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
            print(f"  Ошибка: {e}")
            continue

# Рассчитываем средние
num_plots = len(plot_data_list)

print("\n" + "=" * 80)
print("РЕЗУЛЬТАТЫ:")
print("=" * 80)

if num_plots > 0:
    avg_diameter = sum(p['diameter'] for p in plot_data_list) / num_plots
    
    print(f"Количество площадок: {num_plots}")
    print(f"Средний диаметр по площадкам: {avg_diameter:.1f} см")
    
    # Теперь считаем по породам
    print("\n" + "=" * 80)
    print("ПО ПОРОДАМ:")
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
                    
                    diameter = breed_info.get('diameter', 0) or 0
                    
                    if breed_name not in breeds_data:
                        breeds_data[breed_name] = {
                            'diameters': [],
                            'plots': []
                        }
                    
                    breeds_data[breed_name]['diameters'].append(diameter)
                    breeds_data[breed_name]['plots'].append(diameter)
            
            except Exception as e:
                continue
    
    for breed_name, data in breeds_data.items():
        avg_d = sum(d for d in data['diameters'] if d > 0) / len([d for d in data['diameters'] if d > 0]) if any(d > 0 for d in data['diameters']) else 0
        print(f"{breed_name}: средний диаметр = {avg_d:.1f} см (из {len([d for d in data['diameters'] if d > 0])} значений)")
else:
    print("❌ Нет данных для расчёта")
