"""
Общие виджеты ForestApp для обратной совместимости
AutoCompleteTextInput, TreeDataInputPopup, ExitConfirmPopup
"""
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.properties import (NumericProperty, BooleanProperty,
                          ObjectProperty, ListProperty)
import sqlite3

from kivymd.uix.button import MDButton, MDButtonText
from ui_styles import Colors, Spacing, Fonts


class AutoCompleteTextInput(TextInput):
    next_widget = ObjectProperty(None)
    prev_widget = ObjectProperty(None)
    row_index = NumericProperty(0)
    col_index = NumericProperty(0)
    suggestions = ListProperty([])

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bind(text=self.show_suggestions)
        self.popup = None

    def keyboard_on_key_down(self, window, keycode, text, modifiers):
        key = keycode[1]
        if key == 'down':
            self.focus_next('down')
        elif key == 'up':
            self.focus_previous('up')
        elif key == 'right':
            self.focus_next('right')
        elif key == 'left':
            self.focus_previous('left')
        else:
            super().keyboard_on_key_down(window, keycode, text, modifiers)
        return True

    def show_suggestions(self, instance, value):
        if not value or len(value) < 3:
            return
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT value FROM suggestions
            WHERE column_index = ? AND value LIKE ?
            ORDER BY LENGTH(value) ASC, value ASC
            LIMIT 1
        ''', (self.col_index, f'{value}%'))
        results = cursor.fetchall()
        conn.close()
        if results:
            self.text = results[0][0]

    def get_table_screen(self):
        return App.get_running_app().root.get_screen('table')

    def focus_next(self, direction):
        table_screen = self.get_table_screen()
        if direction == 'right' and self.next_widget:
            self.next_widget.focus = True
        elif direction == 'down':
            next_row = self.row_index + 1
            if next_row < len(table_screen.inputs):
                table_screen.inputs[next_row][self.col_index].focus = True

    def focus_previous(self, direction):
        table_screen = self.get_table_screen()
        if direction == 'left' and self.prev_widget:
            self.prev_widget.focus = True
        elif direction == 'up':
            prev_row = self.row_index - 1
            if prev_row >= 0:
                table_screen.inputs[prev_row][self.col_index].focus = True


class TreeDataInputPopup(Popup):
    def __init__(self, table_screen, row_index, **kwargs):
        super().__init__(
            title='Ввод данных дерева',
            size_hint=(0.8, 0.9),
            separator_height=0,
            **kwargs
        )
        self.table_screen = table_screen
        self.row_index = row_index
        self.fields = [
            ('Порода', 1), ('ж/ф', 2), ('шт/либо лет', 3),
            ('D, см', 4), ('H, м', 5), ('Сост-е', 6), ('Модель', 7), ('Примечания', 8)
        ]
        self.data = {}
        self.create_ui()

    def create_ui(self):
        content = BoxLayout(orientation='vertical', spacing=15, padding=15)
        title_label = Label(
            text='Ввод данных дерева', font_name='Roboto',
            font_size='20sp', bold=True, color=(0, 0.5, 0, 1),
            size_hint=(1, None), height=50
        )
        content.add_widget(title_label)
        scroll = ScrollView(size_hint=(1, 1))
        scroll_content = GridLayout(cols=1, spacing=15, size_hint_y=None)
        scroll_content.bind(minimum_height=scroll_content.setter('height'))
        self.input_fields = []
        for field_name, col_index in self.fields:
            field_layout = BoxLayout(
                orientation='vertical', size_hint_y=None, height=70, spacing=5
            )
            field_label = Label(
                text=field_name, font_name='Roboto', font_size='16sp',
                bold=True, color=(0, 0, 0, 1), size_hint_y=None, height=25
            )
            input_field = AutoCompleteTextInput(
                multiline=False, size_hint_y=None, height=40,
                background_color=(1, 1, 1, 1), col_index=col_index, font_name='Roboto'
            )
            self.input_fields.append(input_field)
            field_layout.add_widget(field_label)
            field_layout.add_widget(input_field)
            scroll_content.add_widget(field_layout)
        scroll.add_widget(scroll_content)
        content.add_widget(scroll)
        btn_box = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=60)
        save_btn = MDButton(
            style='filled', md_bg_color=Colors.PRIMARY,
            size_hint=(0.5, None), height=60, on_release=self.save_data,
        )
        save_btn.add_widget(MDButtonText(text='Сохранить', font_size='18sp'))
        exit_btn = MDButton(
            style='outlined',
            size_hint=(0.5, None), height=60, on_release=self.dismiss,
        )
        exit_btn.add_widget(MDButtonText(
            text='Выйти', theme_text_color='Custom',
            text_color=Colors.DANGER, font_size='18sp'
        ))
        btn_box.add_widget(save_btn)
        btn_box.add_widget(exit_btn)
        content.add_widget(btn_box)
        self.content = content
        self.open()

    def save_data(self, instance):
        for i, (field_name, col_index) in enumerate(self.fields):
            value = self.input_fields[i].text.strip()
            if value:
                self.data[col_index] = value
                self.save_to_suggestions(col_index, value)
        for col_index, value in self.data.items():
            if col_index < len(self.table_screen.inputs[self.row_index]):
                self.table_screen.inputs[self.row_index][col_index].text = value
        self.table_screen.save_current_page()
        base_number = self.table_screen.current_page * self.table_screen.rows_per_page + self.row_index + 1
        for row_idx in range(self.row_index + 1, len(self.table_screen.inputs)):
            tree_number = base_number + (row_idx - self.row_index)
            self.table_screen.inputs[row_idx][0].text = str(tree_number)
        self.table_screen.show_success("Данные дерева сохранены!")
        self.dismiss()

    def save_to_suggestions(self, col_index, value):
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR IGNORE INTO suggestions (column_index, value)
                VALUES (?, ?)
            ''', (col_index, value))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Error saving suggestion: {e}")


class ExitConfirmPopup(Popup):
    def __init__(self, **kwargs):
        super().__init__(
            title='', separator_height=0,
            size_hint=(0.6, 0.45), **kwargs
        )
        content = BoxLayout(orientation='vertical', spacing=15, padding=15)
        title_label = Label(
            text='Подтверждение выхода', font_name='Roboto',
            font_size='20sp', bold=True, color=(0, 0.5, 0, 1),
            size_hint=(1, None), height=50
        )
        content.add_widget(title_label)
        label = Label(
            text='Вы уверены, что хотите выйти?', font_name='Roboto',
            font_size='18sp', color=(0.2, 0.2, 0.2, 1),
            size_hint=(1, None), height=60
        )
        content.add_widget(label)
        btn_box = BoxLayout(orientation='horizontal', spacing=15, size_hint=(1, None), height=70)
        yes_btn = MDButton(
            style='filled', md_bg_color=Colors.DANGER,
            size_hint=(0.5, None), height=70,
        )
        yes_btn.add_widget(MDButtonText(text='Выход', font_size='18sp'))
        yes_btn.bind(on_release=lambda x: App.get_running_app().stop())
        no_btn = MDButton(
            style='outlined',
            size_hint=(0.5, None), height=70,
        )
        no_btn.add_widget(MDButtonText(
            text='Отмена', theme_text_color='Custom',
            text_color=Colors.TEXT_SECONDARY, font_size='18sp'
        ))
        no_btn.bind(on_release=self.dismiss)
        btn_box.add_widget(yes_btn)
        btn_box.add_widget(no_btn)
        content.add_widget(btn_box)
        self.content = content
