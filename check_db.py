import sqlite3

conn = sqlite3.connect('forest_data.db')
cursor = conn.cursor()

# Получаем список всех таблиц
cursor.execute('SELECT name FROM sqlite_master WHERE type="table"')
tables = [row[0] for row in cursor.fetchall()]

print('Tables in database:')
for table in tables:
    print(f'  - {table}')

    # Получаем структуру таблицы
    cursor.execute(f'PRAGMA table_info({table})')
    columns = cursor.fetchall()
    if columns:
        print('    Columns:')
        for col in columns:
            print(f'      {col[1]} ({col[2]})')

    # Получаем количество записей
    try:
        cursor.execute(f'SELECT COUNT(*) FROM {table}')
        count = cursor.fetchone()[0]
        print(f'    Records: {count}')
    except:
        print('    Records: N/A')

    print()

conn.close()
