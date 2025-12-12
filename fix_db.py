import sqlite3
import os

def fix_database():
    db_path = 'forest_data.db'
    if not os.path.exists(db_path):
        print('База данных не найдена')
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print('Проверка структуры таблицы molodniki_data...')

    # Проверим структуру таблицы
    cursor.execute('PRAGMA table_info(molodniki_data)')
    columns = [row[1] for row in cursor.fetchall()]
    print('Текущие столбцы таблицы molodniki_data:')
    for col in columns:
        print(f'  - {col}')

    # Проверим, есть ли столбец tip_lesa
    if 'tip_lesa' not in columns:
        print('\nСтолбец tip_lesa отсутствует! Добавляем...')
        try:
            cursor.execute('ALTER TABLE molodniki_data ADD COLUMN tip_lesa TEXT DEFAULT ""')
            print('Столбец tip_lesa успешно добавлен')
        except sqlite3.OperationalError as e:
            print(f'Ошибка добавления столбца: {e}')
    else:
        print('\nСтолбец tip_lesa уже существует')

    conn.commit()
    conn.close()
    print('Исправление базы данных завершено')

if __name__ == '__main__':
    fix_database()
