#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отладка расчёта интенсивности рубки
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
print("ОТЛАДКА РАСЧЁТА ИНТЕНСИВНОСТИ")
print("=" * 60)
print(f"Площадь учётной площадки: {plot_area_ha:.6f} га ({3.14159 * radius ** 2:.2f} м2)")

def parse_care_subject_density(care_text):
    """Парсит предмет ухода и возвращает оставляемую густоту на гектар"""
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

    # Предмет ухода показывает сколько деревьев оставить на гектар
    # Например, "3С" значит оставить 3000 сосен на гектар
    return total_density * 1000

# Расчёт общей густоты
total_density_all_plots = 0
plot_count = 0

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
                                plot_density += total_trees / plot_area_ha if plot_area_ha > 0 else 0
                                print(f"  Хвойная {breed_info.get('name', 'Н/Д')}: {total_trees} дер. на площадке = {total_trees / plot_area_ha:.1f} шт/га")
                            else:
                                density_value = breed_info.get('density', 0)
                                plot_density += density_value / plot_area_ha if plot_area_ha > 0 else 0
                                print(f"  Лиственная {breed_info.get('name', 'Н/Д')}: {density_value} дер. на площадке = {density_value / plot_area_ha:.1f} шт/га")
                except (json.JSONDecodeError, TypeError):
                    pass

            if plot_density > 0:
                total_density_all_plots += plot_density
                plot_count += 1
                print(f"  Строка: общая густота = {plot_density:.1f} шт/га")

print(f"\nОбщая густота (сумма по всем строкам): {total_density_all_plots:.1f} шт/га")
print(f"Количество строк с данными: {plot_count}")

# Расчёт оставаемой густоты из предмета ухода
total_remaining_density = 0
plot_count_with_care = 0

for page_num, page_rows in page_data.items():
    for row in page_rows:
        if len(row) >= 3 and row[2]:
            care_text = row[2].strip()
            if care_text:
                remaining_density = parse_care_subject_density(care_text)
                if remaining_density > 0:
                    total_remaining_density += remaining_density
                    plot_count_with_care += 1
                    print(f"  Строка: предмет ухода '{care_text}' = {remaining_density:.0f} шт/га оставляется")

print(f"\nОставаемая густота (сумма): {total_remaining_density:.0f} шт/га")
print(f"Количество строк с предметом ухода: {plot_count_with_care}")

# Расчёт интенсивности
if plot_count > 0 and total_density_all_plots > 0:
    avg_overall_density = total_density_all_plots / plot_count
    
    if plot_count_with_care > 0:
        avg_remaining_density = total_remaining_density / plot_count_with_care
        
        intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100
        
        print(f"\n=== РЕЗУЛЬТАТЫ ===")
        print(f"Средняя исходная густота: {avg_overall_density:.1f} шт/га")
        print(f"Средняя оставаемая густота: {avg_remaining_density:.1f} шт/га")
        print(f"Интенсивность рубки: {intensity:.1f}%")
        
        # Проверка: сколько вырубаем
        cut_density = avg_overall_density - avg_remaining_density
        print(f"Вырубаемая густота: {cut_density:.1f} шт/га")
        print(f"Проверка: {cut_density / avg_overall_density * 100:.1f}% вырубаем")
    else:
        print("\nНет данных о предмете ухода!")
else:
    print("\nНет данных для расчёта!")

print("\n" + "=" * 60)
