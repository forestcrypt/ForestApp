#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для подстановки данных из приложения "Молодняки" в Word-документ
"Проект рубок ухода в молодняках"
"""

import os
import sys
import json
import argparse
from docx import Document
import sqlite3
import datetime

# Добавляем текущую директорию в путь для импорта модулей
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

class WordDocumentFiller:
    def __init__(self, db_name='forest_data.db', data_file=None, address_data=None, total_data=None):
        self.db_name = db_name
        self.document_path = 'reports/Проект ухода для молодняков.docx'
        self.data_file = data_file

        # Данные адресной строки (можно настроить)
        self.address_data = address_data or {
            'quarter': '1',  # Квартал
            'plot': '15',    # Выдел
            'section': 'Володозерское',  # Участковое лесничество
            'forestry': 'Сегежское лесничество',  # Лесничество
            'target_purpose': 'Эксплуатационные леса',  # Целевое назначение
            'plot_area': '25.5',  # Площадь участка (га)
            'forest_type': 'Сосняк черничный'  # Тип леса
        }

        # Данные из Итого
        self.total_data = total_data or {}

    def load_data_from_json(self, file_path):
        """Загружаем данные из JSON файла"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Извлекаем данные адресной строки и итоговые данные
            self.address_data = data.get('address_data', self.address_data)
            self.total_data = data.get('total_data', self.total_data)

            print("Данные успешно загружены из JSON файла")
            return True

        except Exception as e:
            print(f"Ошибка загрузки данных из JSON файла: {e}")
            return False

    def load_data_from_db(self):
        """Загружаем данные из базы данных приложения"""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()

            # Получаем последние данные из таблицы molodniki_totals
            cursor.execute('''
                SELECT * FROM molodniki_totals
                ORDER BY created_at DESC
                LIMIT 1
            ''')

            totals_row = cursor.fetchone()
            if totals_row:
                self.total_data = {
                    'page_number': totals_row[1],
                    'section_name': totals_row[2],
                    'total_composition': totals_row[3],
                    'avg_age': totals_row[5],
                    'avg_density': totals_row[6],
                    'avg_height': totals_row[7]
                }

            # Получаем данные пород
            cursor.execute('''
                SELECT breed_name, breed_type, density, height, age, do_05, _05_15, bolee_15
                FROM molodniki_breeds
                ORDER BY id DESC
                LIMIT 20
            ''')

            breeds = cursor.fetchall()
            self.total_data['breeds'] = []
            for breed in breeds:
                self.total_data['breeds'].append({
                    'name': breed[0],
                    'type': breed[1],
                    'density': breed[2],
                    'height': breed[3],
                    'age': breed[4],
                    'do_05': breed[5],
                    '_05_15': breed[6],
                    'bolee_15': breed[7]
                })

            conn.close()
            print("Данные успешно загружены из базы данных")
            return True

        except Exception as e:
            print(f"Ошибка загрузки данных из БД: {e}")
            return False

    def calculate_totals(self):
        """Используем данные Итого напрямую без дополнительных расчетов"""
        # Данные уже загружены из таблицы molodniki_totals в load_data_from_db()
        # Никаких дополнительных расчетов не требуется

        # Убеждаемся, что все необходимые поля присутствуют
        if not self.total_data.get('composition'):
            self.total_data['composition'] = self.total_data.get('total_composition', 'Не определен')

        if not self.total_data.get('care_subject'):
            self.total_data['care_subject'] = 'Не определен'  # Предмет ухода из Итого

        if not self.total_data.get('intensity'):
            self.total_data['intensity'] = '25%'  # Интенсивность из Итого или по умолчанию

    def get_breed_letter(self, breed_name):
        """Получаем букву для породы в формуле состава"""
        letters = {
            'Сосна': 'С',
            'Ель': 'Е',
            'Пихта': 'П',
            'Кедр': 'К',
            'Лиственница': 'Л',
            'Берёза': 'Б',
            'Осина': 'Ос',
            'Ольха': 'О'
        }

        for name, letter in letters.items():
            if name.lower() in breed_name.lower():
                return letter

        return breed_name[0].upper()

    def get_test_data(self):
        """Тестовые данные для демонстрации"""
        return {
            'composition': '8С2БДр',
            'care_subject': '300шт/гаС + 50шт/гаБ',
            'avg_age': '25',
            'avg_density': '350',
            'avg_height': '12.5',
            'intensity': '25%',
            'breeds': [
                {'name': 'Сосна', 'type': 'coniferous', 'density': 300, 'height': 15, 'age': 30, 'do_05': 50, '_05_15': 150, 'bolee_15': 100},
                {'name': 'Берёза', 'type': 'deciduous', 'density': 50, 'height': 10, 'age': 20}
            ]
        }

    def fill_document(self):
        """Заполняем документ данными"""
        if not os.path.exists(self.document_path):
            print(f"Файл {self.document_path} не найден!")
            return False

        try:
            doc = Document(self.document_path)

            # Словарь замен для различных плейсхолдеров в документе
            replacements = {
                # Старые плейсхолдеры в фигурных скобках
                '{quarter}': self.address_data.get('quarter', ''),
                '{plot}': self.address_data.get('plot', ''),
                '{section}': self.address_data.get('section', ''),
                '{forestry}': self.address_data.get('forestry', ''),
                '{target_purpose}': self.address_data.get('target_purpose', ''),
                '{plot_area}': self.address_data.get('plot_area', ''),
                '{forest_type}': self.address_data.get('forest_type', ''),
                '{filename}': self.total_data.get('section_name', 'Молодняки'),
                '{plot_info}': f'Обследовано {self.total_data.get("total_plots", len(self.total_data.get("breeds", [])))} площадок',
                '{composition}': self.total_data.get('composition', 'Не определен'),
                '{care_subject}': self.total_data.get('care_subject', 'Не определен'),
                '{avg_age}': str(self.total_data.get('avg_age', 'Н/Д')),
                '{avg_diameter}': str(self.total_data.get('avg_diameter', 'Н/Д')),
                '{avg_height}': str(self.total_data.get('avg_height', 'Н/Д')),
                '{avg_density}': str(self.total_data.get('avg_density', 'Н/Д')),
                '{intensity}': self.total_data.get('intensity', '25%'),

                # Новые плейсхолдеры из текста документа
                '( подставляем данные из адресной строки)': f"{self.address_data.get('section', '')} участковом лесничестве",
                '(ПОДСТАВЛЯЕМ ДАННЫЕ ИЗ АДРЕСНОЙ СТРОКИ)': self.address_data.get('forestry', ''),
                '( подставляем данные из адресной строки)': self.address_data.get('target_purpose', ''),
                'Выдел( подставляем данные из адресной строки)': self.address_data.get('plot', ''),
                'площадь( подставляем данные из адресной строки)': self.address_data.get('plot_area', ''),
                'квартал  ( подставляем данные из адресной строки)': self.address_data.get('quarter', ''),
                '( подставляем данные по среднему показателю по типу леса по всем площадкам)': self.address_data.get('forest_type', ''),

                # Данные из Итого
                '(подставляем данные из Итого по информации о площадке)': f'Обследовано {self.total_data.get("total_plots", len(self.total_data.get("breeds", [])))} площадок',
                '(подставляем данные из Итого, коэффициент состава)': self.total_data.get('composition', 'Не определен'),
                'подставляем данные из Итого, Прдмет ухода среднее по площадкам': self.total_data.get('care_subject', 'Не определен'),
                'подставляем данные из Итого,средние показатели возраста по породам': str(self.total_data.get('avg_age', 'Н/Д')),
                'подставляем данные из Итого,средние показатели диаметра по породам': str(self.total_data.get('avg_diameter', 'Н/Д')),
                'подставляем данные из Итого,средние показатели высоты по породам': str(self.total_data.get('avg_height', 'Н/Д')),
                'подставляем данные из Итого,средние показатели густоты по породам': str(self.total_data.get('avg_density', 'Н/Д')),
                'подставляем данные из Итого,средние показатели предмета ухода по породам': self.total_data.get('care_subject', 'Не определен'),
                '(подставляем данные из Итого)': self.total_data.get('intensity', '25%'),
                '(подставляем данные из Итого,средние показатели предмета ухода по породам)': self.total_data.get('care_subject', 'Не определен')
            }

            # Проходим по всем параграфам и заменяем текст
            for paragraph in doc.paragraphs:
                for old_text, new_text in replacements.items():
                    if old_text in paragraph.text:
                        paragraph.text = paragraph.text.replace(old_text, str(new_text))

            # Также проверяем таблицы
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for old_text, new_text in replacements.items():
                                if old_text in paragraph.text:
                                    paragraph.text = paragraph.text.replace(old_text, str(new_text))

            # Сохраняем документ
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = self.document_path.replace('.docx', f'_заполненный_{timestamp}.docx')
            doc.save(output_path)

            print(f"Документ успешно заполнен и сохранен как: {output_path}")
            return True

        except Exception as e:
            print(f"Ошибка при заполнении документа: {e}")
            return False

    def run(self):
        """Основной метод выполнения"""
        print("Начинаем заполнение Word-документа...")

        # Если указан файл данных, загружаем из него
        if self.data_file:
            if not self.load_data_from_json(self.data_file):
                print(f"Ошибка: Не удалось загрузить данные из файла {self.data_file}.")
                return False
        else:
            # Загружаем данные из БД (для обратной совместимости)
            if not self.load_data_from_db():
                print("Ошибка: Не удалось загрузить данные из базы данных. Проверьте наличие данных в таблице molodniki_totals.")
                return False

        # Рассчитываем итоги (используем данные Итого напрямую)
        self.calculate_totals()

        # Заполняем документ
        success = self.fill_document()

        if success:
            print("Задача выполнена успешно!")
        else:
            print("Произошла ошибка при выполнении задачи.")

        return success

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Заполнение Word-документа данными из приложения Молодняки')
    parser.add_argument('--data-file', type=str, help='Путь к JSON файлу с данными (address_data и total_data)')

    args = parser.parse_args()

    filler = WordDocumentFiller(data_file=args.data_file)
    filler.run()
