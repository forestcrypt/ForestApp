#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Прямой тест заполнения Word документа
"""

import sys
import os
import json
import tempfile

# Добавляем текущую директорию в путь
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from fill_word_document import WordDocumentFiller

def test_direct_fill():
    """Тестируем прямое заполнение документа"""
    try:
        # Создаем тестовые данные
        address_data = {
            'quarter': '1',
            'plot': '15',
            'section': 'Володозерское',
            'district_forestry': 'Володозерское',
            'forestry': 'Сегежское лесничество',
            'target_purpose': 'Эксплуатационные леса',
            'plot_area': '25.5',
            'radius': '1.78',
            'forest_type': 'Сосняк черничный'
        }

        total_data = {
            'page_number': 0,
            'section_name': 'Молодняки',
            'total_composition': '8С2БДр',
            'avg_age': 25.0,
            'avg_density': 350.0,
            'avg_height': 12.5,
            'total_plots': 10,
            'composition': '8С2БДр',
            'care_subject': '300шт/гаС + 50шт/гаБ',
            'intensity': '25%',
            'care_queue': 'первая',
            'characteristics': {
                'best': 'здоровая, хорошо укорененная сосна, с хорошо сформированной кроной',
                'auxiliary': 'деревья всех пород обеспечивающие сохранение целостности и устойчивости насаждения',
                'undesirable': 'деревья мешающие росту и формированию крон отобранных лучших и вспомогательных деревьев; деревья неудовлетворительного состояния'
            },
            'care_date': 'сент 2025 года',
            'technology': 'Равномерное изреживание молодняка. Срубленные деревья необходимо приземлить на месте.',
            'forest_purpose': 'Эксплуатационные леса',
            'activity_name': 'осветление'
        }

        # Создаем экземпляр WordDocumentFiller с тестовыми данными
        filler = WordDocumentFiller(
            data_file=None,
            address_data=address_data,
            total_data=total_data
        )

        # Запускаем заполнение
        print("Запускаем прямое заполнение документа...")
        success = filler.run()

        if success:
            print("Прямой тест завершен успешно!")
            return True
        else:
            print("Прямой тест завершился с ошибкой!")
            return False

    except Exception as e:
        print(f"Ошибка при прямом тестировании: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    test_direct_fill()
