#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отладочный скрипт для анализа данных проекта
"""

import json
import os

# Загружаем JSON файл
json_path = 'reports/Молодняки_3_20260309_1402.json'
with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=" * 60)
print("АНАЛИЗ ДАННЫХ ПРОЕКТА")
print("=" * 60)

# Извлекаем данные
page_data = data.get('page_data', {})
project_data = data.get('project_data', {})
address_data = project_data.get('address', {})
details_data = project_data.get('details', {})

print("\n1. АДРЕСНАЯ СТРОКА:")
print(f"   Квартал: {address_data.get('quarter')}")
print(f"   Выдел: {address_data.get('plot')}")
print(f"   Лесничество: {address_data.get('forestry')}")
print(f"   Район: {address_data.get('district_forestry')}")
print(f"   Радиус: {address_data.get('radius')}")
print(f"   Площадь: {address_data.get('plot_area')}")

print("\n2. ДЕТАЛИ УХОДА:")
print(f"   Очередь: {details_data.get('care_queue')}")
print(f"   Характеристики: {details_data.get('characteristics')}")
print(f"   Дата: {details_data.get('care_date')}")
print(f"   Технология: {details_data.get('technology')}")
print(f"   Назначение: {details_data.get('forest_purpose')}")

print("\n3. ДАННЫЕ СТРАНИЦЫ (page_data):")
for page_num, rows in page_data.items():
    print(f"   Страница {page_num}:")
    for i, row in enumerate(rows):
        print(f"      Строка {i}:")
        if len(row) >= 5:
            print(f"         № п/п: {row[0]}")
            print(f"         Густота: {row[1]}")
            print(f"         Предмет ухода: {row[2]}")
            print(f"         Порода: {row[3]}")
            print(f"         Живой напочвенный покров: {row[4]}")
            
            # Парсим породы
            if row[3]:
                try:
                    breeds_list = json.loads(row[3]) if isinstance(row[3], str) else []
                    print(f"         Распарсенные породы:")
                    for breed in breeds_list:
                        if isinstance(breed, dict):
                            breed_name = breed.get('name', 'Н/Д')
                            breed_type = breed.get('type', 'Н/Д')
                            density = breed.get('density', 'Н/Д')
                            height = breed.get('height', 'Н/Д')
                            diameter = breed.get('diameter', 'Н/Д')
                            age = breed.get('age', 'Н/Д')
                            
                            # Для хвойных
                            if breed_type == 'coniferous':
                                do_05 = breed.get('do_05', 'Н/Д')
                                _05_15 = breed.get('05_15', 'Н/Д')
                                bolee_15 = breed.get('bolee_15', 'Н/Д')
                                print(f"            - {breed_name} (хвойная):")
                                print(f"              до 0.5м: {do_05}, 0.5-1.5м: {_05_15}, >1.5м: {bolee_15}")
                                print(f"              Высота: {height}, Диаметр: {diameter}, Возраст: {age}")
                            else:
                                print(f"            - {breed_name} (лиственная):")
                                print(f"              Густота: {density}, Высота: {height}, Диаметр: {diameter}, Возраст: {age}")
                except json.JSONDecodeError as e:
                    print(f"         Ошибка парсинга пород: {e}")

# Рассчитываем площадь
radius = float(address_data.get('radius', 5.64))
plot_area_ha = 3.14159 * (radius ** 2) / 10000
print(f"\n4. РАСЧЁТНЫЕ ЗНАЧЕНИЯ:")
print(f"   Площадь учётной площадки (га): {plot_area_ha:.4f}")
print(f"   Площадь учётной площадки (м2): {3.14159 * radius ** 2:.2f}")

# Собираем данные по породам
print("\n5. ДАННЫЕ ПО ПОРОДАМ (с расчётом густоты):")
breeds_data = {}

for page_num, page_rows in page_data.items():
    for row in page_rows:
        if len(row) < 4:
            continue
        
        breeds_text = row[3]
        if not breeds_text:
            continue
        
        try:
            breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
        except json.JSONDecodeError:
            continue
        
        for breed_info in breeds_list:
            if not isinstance(breed_info, dict):
                continue
            
            breed_name = breed_info.get('name', '').strip()
            if not breed_name:
                continue
            
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
                    
                print(f"   {breed_name} (хвойная):")
                print(f"      Деревьев на площадке: {total_trees} (до 0.5м: {do_05}, 0.5-1.5м: {_05_15}, >1.5м: {bolee_15})")
                print(f"      Густота (шт/га): {density:.1f}")
                print(f"      Высота (м): {height}")
                
            else:
                density_value = breed_info.get('density', 0)
                density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                height = breed_info.get('height', 0) or 0
                
                print(f"   {breed_name} (лиственная):")
                print(f"      Деревьев на площадке: {density_value}")
                print(f"      Густота (шт/га): {density:.1f}")
                print(f"      Высота (м): {height}")
            
            diameter = breed_info.get('diameter', 0) or 0
            age = breed_info.get('age', 0) or 0
            
            print(f"      Диаметр (см): {diameter}")
            print(f"      Возраст (лет): {age}")
            
            if breed_name not in breeds_data:
                breeds_data[breed_name] = {
                    'density': density,
                    'height': height,
                    'diameter': diameter,
                    'age': age
                }
            else:
                # Усредняем
                old = breeds_data[breed_name]
                old['density'] = (old['density'] + density) / 2
                old['height'] = (old['height'] + height) / 2
                old['diameter'] = (old['diameter'] + diameter) / 2
                old['age'] = (old['age'] + age) / 2

print("\n6. ИТОГОВЫЕ ДАННЫЕ ПО ПОРОДАМ (усреднённые):")
for breed_name, values in breeds_data.items():
    print(f"   {breed_name}:")
    print(f"      Густота: {values['density']:.1f} шт/га")
    print(f"      Высота: {values['height']:.1f} м")
    print(f"      Диаметр: {values['diameter']:.1f} см")
    print(f"      Возраст: {values['age']:.1f} лет")

# Расчёт состава
print("\n7. КОЭФФИЦИЕНТ СОСТАВА:")
total_densities = {name: data['density'] for name, data in breeds_data.items()}
if total_densities:
    total_all_density = sum(total_densities.values())
    print(f"   Общая густота: {total_all_density:.1f} шт/га")
    
    composition_parts = []
    for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
        coeff = max(1, round(density / total_all_density * 10))
        # Получаем первую букву
        breed_letter = breed_name[0].upper()
        if breed_name.startswith('Ос'):
            breed_letter = 'Ос'
        elif breed_name.startswith('Ол'):
            breed_letter = 'Ол'
        composition_parts.append(f"{coeff}{breed_letter}")
        print(f"   {breed_name}: {density:.1f} шт/га -> коэффициент {coeff}{breed_letter}")
    
    print(f"   Итоговый состав: {''.join(composition_parts)}")

print("\n" + "=" * 60)
print("АНАЛИЗ ЗАВЕРШЁН")
print("=" * 60)
