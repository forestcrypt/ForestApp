#!/usr/bin/env python3
"""
Script to delete sections from database if they are absent in reports folder
"""
import sqlite3
import os
import glob

def cleanup_sections():
    db_path = 'forest_data.db'
    reports_dir = 'reports'

    if not os.path.exists(reports_dir):
        print(f"Папка {reports_dir} не существует")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Очистка обычных участков
        print("Проверка обычных участков...")
        cursor.execute('SELECT id, section_number FROM sections WHERE section_number IS NOT NULL AND section_number != ""')
        sections_to_check = cursor.fetchall()

        for section_id, section_number in sections_to_check:
            pattern = os.path.join(reports_dir, f"{section_number}_*.xlsx")
            files = glob.glob(pattern)
            if not files:
                print(f"Удаление участка (нет файла): {section_number}")
                cursor.execute('DELETE FROM sections WHERE id = ?', (section_id,))
            else:
                print(f"Файл найден для участка: {section_number}")

        # Очистка участков молодняков
        print("\nПроверка участков молодняков...")
        cursor.execute('SELECT id, section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != ""')
        molodniki_sections = cursor.fetchall()

        for section_id, section_number in molodniki_sections:
            pattern = os.path.join(reports_dir, f"Молодняки_расширенный_{section_number}_*.xlsx")
            files = glob.glob(pattern)
            if not files:
                print(f"Удаление участка молодняков (нет файла): {section_number}")
                cursor.execute('DELETE FROM molodniki_sections WHERE id = ?', (section_id,))
            else:
                print(f"Файл найден для участка молодняков: {section_number}")

        conn.commit()
        print("\nОчистка завершена")

    except Exception as e:
        print(f"Ошибка: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == "__main__":
    cleanup_sections()
