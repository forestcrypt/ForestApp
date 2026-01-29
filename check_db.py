#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для проверки базы данных
"""

import sqlite3
import json

def check_database():
    """Проверяем содержимое базы данных"""
    try:
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()

        # Получаем список таблиц
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        print("=== TABLES ===")
        for table in tables:
            print(table[0])

        # Проверяем таблицу molodniki_data
        print("\n=== MOLODNIKI_DATA ===")
        cursor.execute("SELECT * FROM molodniki_data LIMIT 10")
        rows = cursor.fetchall()
        for row in rows:
            print(row)

        # Ищем записи с "Молодняки" в section_name
        print("\n=== RECORDS WITH MOLODNIKI ===")
        cursor.execute("SELECT * FROM molodniki_data WHERE section_name LIKE ?", ('%олодняки%',))
        molodniki_rows = cursor.fetchall()
        for row in molodniki_rows:
            print(row)

        # Проверяем таблицу molodniki_breeds
        print("\n=== MOLODNIKI_BREEDS ===")
        cursor.execute("SELECT * FROM molodniki_breeds LIMIT 5")
        breed_rows = cursor.fetchall()
        for row in breed_rows:
            print(row)

        # Ищем "Молодняки участок 1" в любых текстовых полях
        print("\n=== SEARCHING FOR 'Молодняки участок 1' ===")
        cursor.execute("SELECT * FROM molodniki_data WHERE nn LIKE ? OR gps_point LIKE ? OR predmet_uhoda LIKE ? OR section_name LIKE ?",
                      ('%Молодняки участок 1%', '%Молодняки участок 1%', '%Молодняки участок 1%', '%Молодняки участок 1%'))
        found_rows = cursor.fetchall()
        for row in found_rows:
            print("FOUND:", row)

        conn.close()

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    check_database()