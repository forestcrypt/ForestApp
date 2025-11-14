# Script to check database contents
import sqlite3

def check_database():
    conn = sqlite3.connect('forest_data.db')
    cursor = conn.cursor()

    try:
        # Check molodniki_data table
        print("=== MOLODNIKI_DATA TABLE ===")
        cursor.execute('SELECT * FROM molodniki_data LIMIT 10')
        rows = cursor.fetchall()
        for row in rows:
            print(row)

        # Check molodniki_breeds table
        print("\n=== MOLODNIKI_BREEDS TABLE ===")
        cursor.execute('SELECT * FROM molodniki_breeds LIMIT 10')
        rows = cursor.fetchall()
        for row in rows:
            print(row)

        # Check suggestions
        print("\n=== MOLODNIKI_SUGGESTIONS TABLE ===")
        cursor.execute('SELECT * FROM molodniki_suggestions LIMIT 10')
        rows = cursor.fetchall()
        for row in rows:
            print(row)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    check_database()
