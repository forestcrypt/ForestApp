#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка данных о диаметрах в JSON файле
"""

import json
import sys

# Загружаем JSON
json_path = 'reports/Молодняки_3_20260309_1402.json'

try:
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
except FileNotFoundError:
    print(f"❌ Файл {json_path} не найден!")
    sys.exit(1)

print("=" * 80)
print("ПРОВЕРКА ДАННЫХ О ДИАМЕТРАХ В JSON")
print("=" * 80)

# Проверяем total_data
total_data = data.get('total_data', {})

if not total_data:
    print("❌ total_data отсутствует в JSON!")
    print("   Ключи в JSON:", list(data.keys()))
    sys.exit(1)

print("\n✅ total_data найден")
print(f"   Ключи: {list(total_data.keys())}")

# Проверяем breeds
breeds = total_data.get('breeds', [])

if not breeds:
    print("❌ breeds отсутствует в total_data!")
    sys.exit(1)

print(f"\n✅ Найдено пород: {len(breeds)}")

print("\n" + "=" * 80)
print("ДАННЫЕ ПО ПОРОДАМ:")
print("=" * 80)

for i, breed in enumerate(breeds, 1):
    print(f"\n{i}. {breed.get('name', 'Н/Д')}:")
    print(f"   Тип: {breed.get('type', 'Н/Д')}")
    print(f"   Густота: {breed.get('density', 'Н/Д')} шт/га")
    print(f"   Высота: {breed.get('height', 'Н/Д')} м")
    print(f"   Возраст: {breed.get('age', 'Н/Д')} лет")
    print(f"   Диаметр: {breed.get('diameter', 'Н/Д')} см")  # ✅ ПРОВЕРЯЕМ
    
    if breed.get('type') == 'coniferous':
        print(f"   До 0.5м: {breed.get('do_05', 'Н/Д')}")
        print(f"   0.5-1.5м: {breed.get('_05_15', 'Н/Д')}")
        print(f"   >1.5м: {breed.get('bolee_15', 'Н/Д')}")

print("\n" + "=" * 80)
print("ОБЩИЕ ДАННЫЕ:")
print("=" * 80)
print(f"Средний диаметр: {total_data.get('avg_diameter', 'Н/Д')} см")
print(f"Средняя высота: {total_data.get('avg_height', 'Н/Д')} м")
print(f"Средний возраст: {total_data.get('avg_age', 'Н/Д')} лет")
print(f"Густота: {total_data.get('avg_density', 'Н/Д')} шт/га")
print(f"Состав: {total_data.get('composition', 'Н/Д')}")
print(f"Предмет ухода: {total_data.get('care_subject', 'Н/Д')}")
print(f"Интенсивность: {total_data.get('intensity', 'Н/Д')}%")

print("\n" + "=" * 80)
print("ВЫВОД:")
print("=" * 80)

# Проверяем наличие диаметров
all_have_diameter = True
for breed in breeds:
    if breed.get('diameter', 0) == 0:
        print(f"⚠️ {breed.get('name', 'Порода')}: диаметр = 0")
        all_have_diameter = False

if all_have_diameter:
    print("✅ Все породы имеют диаметр > 0")
else:
    print("❌ Некоторые породы не имеют диаметра")
