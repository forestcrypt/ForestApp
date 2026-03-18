#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый запуск fill_our_template.py с данными проекта
"""

import json
import tempfile
import os
import sys

# Загружаем JSON файл
json_path = 'reports/Молодняки_3_20260309_1402.json'
with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

# Собираем данные аналогично generate_care_project
page_data = data.get('page_data', {})
project_data = data.get('project_data', {})
address_data = project_data.get('address', {})
details_data = project_data.get('details', {})

# Импортируем логику из molodniki_extended для расчёта итоговых данных
# Для простоты создадим total_data вручную на основе анализа

# Рассчитываем площадь
radius = float(address_data.get('radius', 1.78))
plot_area_ha = 3.14159 * (radius ** 2) / 10000

# Данные по породам из JSON
breeds_raw = []
for page_num, page_rows in page_data.items():
    for row in page_rows:
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
                        
                        breeds_raw.append({
                            'name': breed_name,
                            'type': breed_type,
                            'density': density,
                            'height': height,
                            'diameter': diameter,
                            'age': age
                        })
            except json.JSONDecodeError:
                continue

# Усредняем данные по породам
breeds_dict = {}
for breed in breeds_raw:
    name = breed['name']
    if name not in breeds_dict:
        breeds_dict[name] = {
            'name': name,
            'type': breed['type'],
            'densities': [],
            'heights': [],
            'diameters': [],
            'ages': []
        }
    breeds_dict[name]['densities'].append(breed['density'])
    breeds_dict[name]['heights'].append(breed['height'])
    breeds_dict[name]['diameters'].append(breed['diameter'])
    breeds_dict[name]['ages'].append(breed['age'])

# Формируем итоговый список пород с усреднёнными значениями
breeds_list = []
for name, data in breeds_dict.items():
    avg_density = sum(data['densities']) / len(data['densities']) if data['densities'] else 0
    avg_height = sum(data['heights']) / len(data['heights']) if data['heights'] else 0
    avg_diameter = sum(data['diameters']) / len(data['diameters']) if data['diameters'] else 0
    avg_age = sum(data['ages']) / len(data['ages']) if data['ages'] else 0
    
    breeds_list.append({
        'name': name,
        'type': data['type'],
        'density': avg_density,
        'height': avg_height,
        'diameter': avg_diameter,
        'age': avg_age
    })

# Расчёт общей густоты и состава
total_densities = {b['name']: b['density'] for b in breeds_list}
total_all_density = sum(total_densities.values())

# Расчёт коэффициентов состава
composition_parts = []
for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
    coeff = max(1, round(density / total_all_density * 10))
    breed_letter = breed_name[0].upper()
    if breed_name.startswith('Ос'):
        breed_letter = 'Ос'
    elif breed_name.startswith('Ол'):
        breed_letter = 'Ол'
    composition_parts.append(f"{coeff}{breed_letter}")

composition_text = ''.join(composition_parts)

# Предмет ухода из данных
care_subject = "3С"  # Из строки данных

# Интенсивность (примерная)
intensity = 25.0

# Средние значения
avg_density = sum(b['density'] for b in breeds_list) / len(breeds_list) if breeds_list else 0
avg_height = sum(b['height'] for b in breeds_list) / len(breeds_list) if breeds_list else 0
avg_diameter = sum(b['diameter'] for b in breeds_list) / len(breeds_list) if breeds_list else 0
avg_age = sum(b['age'] for b in breeds_list) / len(breeds_list) if breeds_list else 0

# Формируем total_data
total_data = {
    'breeds': breeds_list,
    'composition': composition_text,
    'intensity': intensity,
    'care_subject': care_subject,
    'avg_density': avg_density,
    'avg_height': avg_height,
    'avg_diameter': avg_diameter,
    'avg_age': avg_age,
    'total_plots': 1,
    'activity_name': 'осветление',
    'care_queue': details_data.get('care_queue', ''),
    'characteristics': details_data.get('characteristics', ''),
    'care_date': details_data.get('care_date', ''),
    'technology': details_data.get('technology', ''),
    'forest_purpose': details_data.get('forest_purpose', '')
}

# Создаём временный файл
temp_data = {
    'address_data': address_data,
    'total_data': total_data
}

print("=" * 60)
print("ПОДГОТОВЛЕННЫЕ ДАННЫЕ ДЛЯ fill_our_template.py")
print("=" * 60)
print(f"\nAddress data:")
for k, v in address_data.items():
    print(f"  {k}: {v}")

print(f"\nTotal data:")
print(f"  composition: {total_data['composition']}")
print(f"  intensity: {total_data['intensity']}")
print(f"  care_subject: {total_data['care_subject']}")
print(f"  avg_density: {total_data['avg_density']:.1f}")
print(f"  avg_height: {total_data['avg_height']:.1f}")
print(f"  avg_diameter: {total_data['avg_diameter']:.1f}")
print(f"  avg_age: {total_data['avg_age']:.1f}")

print(f"\nBreeds:")
for breed in total_data['breeds']:
    print(f"  {breed['name']}: density={breed['density']:.1f}, height={breed['height']:.1f}, diameter={breed['diameter']:.1f}, age={breed['age']}")

# Сохраняем во временный файл
with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
    json.dump(temp_data, f, ensure_ascii=False, indent=2)
    temp_file = f.name

print(f"\nВременный файл: {temp_file}")

# Запускаем fill_our_template.py
script_path = os.path.join(os.path.dirname(__file__), 'fill_our_template.py')
os.system(f'python "{script_path}" --data-file "{temp_file}"')

# Удаляем временный файл
try:
    os.unlink(temp_file)
except:
    pass
