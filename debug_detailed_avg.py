#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детальная отладка расчёта средних данных по площадкам
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
print("ДЕТАЛЬНЫЙ РАСЧЁТ СРЕДНИХ ДАННЫХ ПО ПЛОЩАДКАМ")
print("=" * 70)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га ({3.14159 * radius ** 2:.2f} м2)")
print()

# Шаг 1: Показываем исходные данные по каждой строке/площадке
print("ШАГ 1: ИСХОДНЫЕ ДАННЫЕ ПО ПЛОЩАДКАМ")
print("-" * 70)

plot_index = 0
all_plot_data = []

for page_num, page_rows in page_data.items():
    for row_idx, row in enumerate(page_rows):
        if len(row) >= 4 and row[3]:  # Есть данные о породах
            print(f"\nПлощадка {plot_index + 1} (страница {page_num}, строка {row_idx}):")
            print(f"  Предмет ухода: {row[2] if len(row) >= 3 else 'Н/Д'}")
            print(f"  Живой покров: {row[4] if len(row) >= 5 else 'Н/Д'}")
            print(f"  Тип леса: {row[5] if len(row) >= 6 else 'Н/Д'}")
            print(f"  Породы:")
            
            plot_total_density = 0
            plot_avg_height = 0
            plot_avg_diameter = 0
            plot_avg_age = 0
            height_count = 0
            diameter_count = 0
            age_count = 0
            
            try:
                breeds_list = json.loads(row[3]) if isinstance(row[3], str) else []
                
                for breed_info in breeds_list:
                    if isinstance(breed_info, dict):
                        breed_name = breed_info.get('name', 'Н/Д')
                        breed_type = breed_info.get('type', 'deciduous')
                        
                        # Расчёт густоты
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            total_trees = do_05 + _05_15 + bolee_15
                            density = total_trees / plot_area_ha if plot_area_ha > 0 else 0
                            
                            # Высота для хвойных
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
                        
                        print(f"    {breed_name}: {density:.1f} шт/га, h={height:.1f}м, d={diameter:.1f}см, age={age}")
                        
                        plot_total_density += density
                        if height > 0:
                            plot_avg_height += height
                            height_count += 1
                        if diameter > 0:
                            plot_avg_diameter += diameter
                            diameter_count += 1
                        if age > 0:
                            plot_avg_age += age
                            age_count += 1
                
                # Рассчитываем средние по площадке
                avg_h = plot_avg_height / height_count if height_count > 0 else 0
                avg_d = plot_avg_diameter / diameter_count if diameter_count > 0 else 0
                avg_a = plot_avg_age / age_count if age_count > 0 else 0
                
                print(f"  → ИТОГО по площадке {plot_index + 1}:")
                print(f"     Общая густота: {plot_total_density:.1f} шт/га")
                print(f"     Средняя высота: {avg_h:.1f} м")
                print(f"     Средний диаметр: {avg_d:.1f} см")
                print(f"     Средний возраст: {avg_a:.1f} лет")
                
                all_plot_data.append({
                    'density': plot_total_density,
                    'height': avg_h,
                    'diameter': avg_d,
                    'age': avg_a
                })
                
            except json.JSONDecodeError as e:
                print(f"    Ошибка парсинга: {e}")
            
            plot_index += 1

# Шаг 2: Расчёт средних по всем площадкам
print("\n" + "=" * 70)
print("ШАГ 2: РАСЧЁТ СРЕДНИХ ЗНАЧЕНИЙ ПО ВСЕМ ПЛОЩАДКАМ")
print("-" * 70)

num_plots = len(all_plot_data)
print(f"Количество площадок: {num_plots}")

if num_plots > 0:
    total_density = sum(p['density'] for p in all_plot_data)
    avg_density = total_density / num_plots
    
    total_height = sum(p['height'] for p in all_plot_data)
    avg_height = total_height / num_plots
    
    total_diameter = sum(p['diameter'] for p in all_plot_data)
    avg_diameter = total_diameter / num_plots
    
    total_age = sum(p['age'] for p in all_plot_data)
    avg_age = total_age / num_plots
    
    print(f"\nСредняя густота: {avg_density:.1f} шт/га (сумма {total_density:.1f} / {num_plots})")
    print(f"Средняя высота: {avg_height:.1f} м (сумма {total_height:.1f} / {num_plots})")
    print(f"Средний диаметр: {avg_diameter:.1f} см (сумма {total_diameter:.1f} / {num_plots})")
    print(f"Средний возраст: {avg_age:.1f} лет (сумма {total_age:.1f} / {num_plots})")

# Шаг 3: Расчёт интенсивности
print("\n" + "=" * 70)
print("ШАГ 3: РАСЧЁТ ИНТЕНСИВНОСТИ РУБКИ")
print("-" * 70)

import re

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

# Считаем оставаемую густоту из предмета ухода
total_remaining = 0
care_count = 0

for page_num, page_rows in page_data.items():
    for row_idx, row in enumerate(page_rows):
        if len(row) >= 3 and row[2]:
            care_text = row[2].strip()
            if care_text:
                remaining = parse_care_subject_density(care_text)
                if remaining > 0:
                    total_remaining += remaining
                    care_count += 1
                    print(f"Площадка {row_idx + 1}: предмет ухода '{care_text}' = {remaining:.0f} шт/га")

if care_count > 0 and num_plots > 0:
    avg_remaining = total_remaining / care_count
    avg_overall_density = avg_density  # Из предыдущего расчёта
    
    intensity = ((avg_overall_density - avg_remaining) / avg_overall_density) * 100
    
    print(f"\nСредняя оставаемая густота: {avg_remaining:.0f} шт/га")
    print(f"Средняя исходная густота: {avg_overall_density:.1f} шт/га")
    print(f"\nИНТЕНСИВНОСТЬ РУБКИ: {intensity:.1f}%")
    print(f"  Формула: (({avg_overall_density:.1f} - {avg_remaining:.0f}) / {avg_overall_density:.1f}) * 100 = {intensity:.1f}%")

print("\n" + "=" * 70)
print("ОЖИДАЕМЫЕ ЗНАЧЕНИЯ В МЕНЮ ИТОГО:")
print("-" * 70)
if num_plots > 0:
    print(f"  Средняя густота: {avg_density:.1f} шт/га")
    print(f"  Средняя высота: {avg_height:.1f} м")
    print(f"  Средний диаметр: {avg_diameter:.1f} см")
    print(f"  Средний возраст: {avg_age:.1f} лет")
    print(f"  Интенсивность: {intensity:.1f}%")
print("=" * 70)
