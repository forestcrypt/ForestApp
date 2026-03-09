#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для заполнения нашего шаблона проекта ухода
Все данные из меню Итого заполняются - по каждой породе и общие
"""

import os
import sys
import json
import datetime
from docx import Document

class OurTemplateFiller:
    def __init__(self, data_file=None):
        self.document_path = 'reports/Шаблон проект_наш.docx'
        self.data_file = data_file
        self.address_data = {}
        self.total_data = {}
        self.details_data = {}

    def load_data_from_json(self, file_path):
        """Загружаем данные из JSON файла"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.address_data = data.get('address_data', {})
            self.total_data = data.get('total_data', {})
            
            # Извлекаем детали из total_data
            self.details_data = {
                'care_queue': self.total_data.get('care_queue', ''),
                'characteristics': self.total_data.get('characteristics', ''),
                'care_date': self.total_data.get('care_date', ''),
                'technology': self.total_data.get('technology', ''),
                'forest_purpose': self.total_data.get('forest_purpose', '')
            }

            print(f"[OK] Данные загружены")
            print(f"  Пород: {len(self.total_data.get('breeds', []))}")
            print(f"  Коэффициент состава: {self.total_data.get('composition', 'Н/Д')}")
            print(f"  Интенсивность: {self.total_data.get('intensity', 'Н/Д')}")
            return True

        except Exception as e:
            print(f"[ERROR] Ошибка загрузки данных: {e}")
            import traceback
            traceback.print_exc()
            return False

    def format_number(self, value, default='Н/Д'):
        """Форматирует число с одной десятичной точкой"""
        if value is None or value == 'Н/Д':
            return default
        try:
            num = float(value)
            return f"{num:.1f}"
        except (ValueError, TypeError):
            return str(value)

    def get_characteristics(self):
        """Получает характеристики деревьев"""
        characteristics = self.details_data.get('characteristics', '')
        
        if isinstance(characteristics, dict):
            return characteristics
        else:
            return {
                'best': 'здоровая, хорошо укорененная сосна, с хорошо сформированной кроной',
                'auxiliary': 'деревья всех пород обеспечивающие сохранение целостности и устойчивости насаждения',
                'undesirable': 'деревья мешающие росту и формированию крон отобранных лучших и вспомогательных деревьев; деревья неудовлетворительного состояния'
            }

    def calculate_breed_composition(self, breed_name, density, total_density):
        """Рассчитывает коэффициент состава для породы"""
        if total_density > 0:
            coeff = round((density / total_density) * 10)
            if coeff < 1:
                coeff = 1
        else:
            coeff = 1
        
        # Получаем букву породы
        letters = {
            'Сосна': 'С', 'Ель': 'Е', 'Пихта': 'П', 'Кедр': 'К', 'Лиственница': 'Л',
            'Берёза': 'Б', 'Осина': 'Ос', 'Ольха': 'О', 'Дуб': 'Д'
        }
        
        letter = 'Др'
        for name, l in letters.items():
            if name.lower() in breed_name.lower():
                letter = l
                break
        
        return f"{coeff}{letter}"

    def fill_document(self):
        """Заполняем документ данными"""
        if not os.path.exists(self.document_path):
            print(f"[ERROR] Файл {self.document_path} не найден!")
            return False

        try:
            doc = Document(self.document_path)
            
            # Получаем данные
            characteristics = self.get_characteristics()
            
            # Рассчитываем параметры площадок
            current_radius = float(self.address_data.get('radius', 1.78))
            plot_area_m2 = 3.14159 * current_radius ** 2
            total_plots = self.total_data.get('total_plots', 0)
            
            # Формируем мероприятие с очередью
            activity_name = self.total_data.get('activity_name', 'осветление')
            care_queue = self.details_data.get('care_queue', 'первая')
            care_activity_text = f"{activity_name}, {care_queue} очередь"
            
            # Словарь общих замен
            replacements = {
                # Адресные данные
                '{forestry}': self.address_data.get('forestry', ''),
                '{district_forestry}': self.address_data.get('district_forestry', ''),
                '{quarter}': self.address_data.get('quarter', ''),
                '{plot}': self.address_data.get('plot', ''),
                '{plot_area}': self.address_data.get('plot_area', ''),
                
                # Детали ухода
                '{care_activity_text}': care_activity_text,
                '{care_queue}': care_queue,
                '{care_date}': self.details_data.get('care_date', ''),
                '{technology}': self.details_data.get('technology', ''),
                '{forest_purpose}': self.details_data.get('forest_purpose', ''),
                '{best}': characteristics.get('best', ''),
                '{auxiliary}': characteristics.get('auxiliary', ''),
                '{undesirable}': characteristics.get('undesirable', ''),
                '{care_subject}': self.total_data.get('care_subject', ''),
                
                # Общие данные
                '{total_composition_isx}': self.total_data.get('composition', 'Н/Д'),
                '{total_composition_project}': self.total_data.get('composition', 'Н/Д'),
                '{total_age_isx}': self.format_number(self.total_data.get('avg_age')),
                '{total_age_project}': self.format_number(self.total_data.get('avg_age')),
                '{total_height_isx}': self.format_number(self.total_data.get('avg_height')),
                '{total_height_project}': self.format_number(self.total_data.get('avg_height')),
                '{total_diameter_isx}': self.format_number(self.total_data.get('avg_diameter', 0)),
                '{total_diameter_project}': self.format_number(self.total_data.get('avg_diameter', 0)),
                '{total_density_isx}': self.format_number(self.total_data.get('avg_density')),
                '{total_density_project}': self.format_number(self.total_data.get('avg_density')),
                '{intensity}': f"{self.total_data.get('intensity', 25):.1f}%",
                
                # Другие
                '{radius_info}': f"{total_plots} шт. {plot_area_m2:.0f}м²(R-{current_radius:.2f}м)",
                '{forest_type}': self.address_data.get('forest_type', 'Смешанный лес'),
            }
            
            # Заполняем параграфы общими данными
            for paragraph in doc.paragraphs:
                for old_text, new_text in replacements.items():
                    if old_text in paragraph.text:
                        paragraph.text = paragraph.text.replace(old_text, str(new_text))
            
            # Заполняем таблицы
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for old_text, new_text in replacements.items():
                                if old_text in paragraph.text:
                                    paragraph.text = paragraph.text.replace(old_text, str(new_text))
            
            # === ТЕПЕРЬ ЗАПОЛНЯЕМ ДАННЫМИ ПО ПОРОДАМ ===
            breeds = self.total_data.get('breeds', [])
            
            if breeds:
                # Находим таблицу с породами
                breeds_table = None
                for table in doc.tables:
                    if len(table.columns) == 11 and 'Порода' in table.rows[0].cells[0].text:
                        breeds_table = table
                        break
                
                if breeds_table:
                    # Удаляем строку-шаблон (первую после заголовка)
                    if len(breeds_table.rows) > 1:
                        breeds_table._tbl.remove(breeds_table.rows[1]._tr)
                    
                    # Добавляем строки для каждой породы
                    for breed in breeds:
                        breed_name = breed.get('name', '')
                        density = breed.get('density', 0)
                        
                        # Рассчитываем состав
                        total_density = sum(b.get('density', 0) for b in breeds)
                        composition_isx = self.calculate_breed_composition(breed_name, density, total_density)
                        composition_project = composition_isx  # Проект такой же, т.к. состав не меняется
                        
                        # Добавляем новую строку
                        row = breeds_table.add_row()
                        row.cells[0].text = breed_name
                        row.cells[1].text = composition_isx
                        row.cells[2].text = composition_project
                        row.cells[3].text = self.format_number(breed.get('age', 0))
                        row.cells[4].text = self.format_number(breed.get('age', 0))
                        row.cells[5].text = self.format_number(breed.get('diameter', 0))
                        row.cells[6].text = self.format_number(breed.get('diameter', 0))
                        row.cells[7].text = self.format_number(breed.get('height', 0))
                        row.cells[8].text = self.format_number(breed.get('height', 0))
                        row.cells[9].text = self.format_number(density)
                        row.cells[10].text = self.format_number(density)
            
            # Сохраняем документ
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = self.document_path.replace('.docx', f'_заполненный_{timestamp}.docx')
            doc.save(output_path)
            
            print(f"[OK] Документ заполнен и сохранён: {output_path}")
            return True

        except Exception as e:
            print(f"[ERROR] Ошибка при заполнении документа: {e}")
            import traceback
            traceback.print_exc()
            return False

    def run(self):
        """Основной метод выполнения"""
        print("=" * 60)
        print("ЗАПОЛНЕНИЕ НАШЕГО ШАБЛОНА ПРОЕКТА УХОДА")
        print("=" * 60)
        
        if self.data_file:
            if not self.load_data_from_json(self.data_file):
                return False
        else:
            print("[ERROR] Не указан файл данных!")
            return False
        
        success = self.fill_document()
        
        if success:
            print("=" * 60)
            print("[OK] Задача выполнена успешно!")
            print("=" * 60)
        else:
            print("=" * 60)
            print("[ERROR] Произошла ошибка при выполнении задачи")
            print("=" * 60)
        
        return success

def fill_document_from_json(json_file_path):
    """Функция для заполнения документа из JSON"""
    try:
        filler = OurTemplateFiller(data_file=json_file_path)
        return filler.run()
    except Exception as e:
        print(f"[ERROR] Ошибка при заполнении документа: {e}")
        return False

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Заполнение нашего шаблона проекта ухода')
    parser.add_argument('--data-file', type=str, help='Путь к JSON файлу с данными')
    args = parser.parse_args()
    
    if args.data_file:
        fill_document_from_json(args.data_file)
    else:
        print("[ERROR] Не указан файл данных! Используйте --data-file <путь>")
