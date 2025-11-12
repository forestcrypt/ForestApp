from kivy.uix.screenmanager import Screen
from kivy.properties import NumericProperty, BooleanProperty, ListProperty, StringProperty
import sqlite3
import os

class BaseTableScreen(Screen):
    current_page = NumericProperty(0)
    total_pages = NumericProperty(1)
    unsaved_changes = BooleanProperty(False)
    focused_cell = ListProperty([0, 0])
    edit_mode = BooleanProperty(False)
    current_section = StringProperty("")
    MAX_PAGES = 200
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.reports_dir = "reports"
        os.makedirs(self.reports_dir, exist_ok=True)
        self.db_name = 'forest_data.db'
        self.rows_per_page = 50
        self.page_data = {}
        self.setup_database()
        
    def setup_database(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute(f'''CREATE TABLE IF NOT EXISTS {self.table_name} (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        section_id INTEGER,
                        {self.get_column_definitions()},
                        FOREIGN KEY(section_id) REFERENCES sections(id))''')
        conn.commit()
        conn.close()
        
    def get_column_definitions(self):
        return " TEXT, ".join(self.TABLE_COLUMNS) + " TEXT"
    
    # Добавьте остальные базовые методы