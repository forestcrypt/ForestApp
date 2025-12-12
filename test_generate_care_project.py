#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки генерации проекта ухода
"""

import sys
import os

# Добавляем текущую директорию в путь
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from molodniki_extended import ExtendedMolodnikiTableScreen

def test_generate_care_project():
    """Тестируем генерацию проекта ухода"""
    try:
        # Создаем экземпляр класса
        screen = ExtendedMolodnikiTableScreen()

        # Устанавливаем тестовые данные
        screen.current_section = "Тестовый участок"
        screen.current_quarter = "1"
        screen.current_plot = "15"
        screen.current_forestry = "Сегежское лесничество"
        screen.current_district_forestry = "Володозерское"
        screen.plot_area_input = "25.5"
        screen.current_radius = "1.78"

        # Вызываем метод генерации проекта ухода
        print("Запускаем генерацию проекта ухода...")
        screen.generate_care_project(None)

        print("Тест завершен успешно!")

    except Exception as e:
        print(f"Ошибка при тестировании: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_generate_care_project()
