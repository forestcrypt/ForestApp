from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.widget import Widget
from kivy.properties import (NumericProperty, BooleanProperty, 
                          ObjectProperty, ListProperty, StringProperty)
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle, Line, RoundedRectangle
from kivy.clock import Clock
from kivy.animation import Animation
from kivy.core.text import LabelBase
from kivy.utils import get_color_from_hex
from kivy.config import Config
from kivy.core.image import Image as CoreImage
import sqlite3
import pandas as pd
import os
import datetime
import re
import json
import shutil
import glob
from openpyxl import load_workbook, Workbook
from tkinter import Tk, filedialog
from molodniki_extended import ExtendedMolodnikiTableScreen

LabelBase.register(name='Roboto', 
                 fn_regular='fonts/Roboto-Medium.ttf',
                 fn_bold='fonts/Roboto-Bold.ttf')

class ThemeManager:
    def __init__(self):
        self.themes = []
        self.current_theme_index = 0
        self.themes_dir = 'themes'
        os.makedirs(self.themes_dir, exist_ok=True)
        self.load_themes()
        self.load_config()
        
    def load_themes(self):
        self.themes = []
        self.themes.extend([
            {
                'type': 'color',
                'name': 'light',
                'background': get_color_from_hex('#FEF7FF'),
                'text_color': '#1C1B1F'
            },
            {
                'type': 'color',
                'name': 'dark',
                'background': get_color_from_hex('#1C1B1F'),
                'text_color': '#E6E1E5'
            }
        ])
        
        for file in os.listdir(self.themes_dir):
            if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                self.themes.append({
                    'type': 'image',
                    'name': os.path.splitext(file)[0],
                    'background': os.path.join(self.themes_dir, file),
                    'text_color': '#FFFFFF'
                })
    
    def add_theme(self, image_path):
        dest = os.path.join(self.themes_dir, os.path.basename(image_path))
        shutil.copy(image_path, dest)
        self.load_themes()
    
    @property
    def current_theme(self):
        return self.themes[self.current_theme_index]
    
    def save_config(self):
        config = {
            'theme_index': self.current_theme_index,
            'themes_dir': self.themes_dir
        }
        with open('theme_config.json', 'w') as f:
            json.dump(config, f)
    
    def load_config(self):
        try:
            with open('theme_config.json', 'r') as f:
                config = json.load(f)
                self.current_theme_index = config.get('theme_index', 0)
                self.themes_dir = config.get('themes_dir', 'themes')
        except (FileNotFoundError, json.JSONDecodeError):
            self.current_theme_index = 0

class ModernButton(Button):
    bg_color = ListProperty([1, 1, 1, 1])
    no_shadow = BooleanProperty(False)

    def __init__(self, **kwargs):
        self.no_shadow = kwargs.pop('no_shadow', False)
        super().__init__(**kwargs)
        self.background_color = (0, 0, 0, 0)
        self.font_name = 'Roboto'
        self.font_size = '16sp'
        self.bold = False
        self.size_hint = (None, None)
        self.height = 50
        self.padding = (20, 5)

        with self.canvas.before:
            if not self.no_shadow:
                Color(rgba=(0, 0, 0, 0.1))
                self.shadow = RoundedRectangle(
                    pos=(self.x+3, self.y-3),
                    size=self.size,
                    radius=[10]
                )
            self.bg_color_instruction = Color(rgba=self.bg_color)
            self.background = RoundedRectangle(
                pos=self.pos,
                size=self.size,
                radius=[10]
            )

        self.bind(pos=self.update_graphics, size=self.update_graphics)
        self.bind(text=self.update_width)

    def update_width(self, instance, value):
        self.width = self.texture_size[0] + 60

    def update_graphics(self, *args):
        self.background.pos = self.pos
        if not self.no_shadow:
            self.shadow.pos = (self.x+3, self.y-3)
            self.shadow.size = self.size
        self.background.size = self.size

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            Animation(rgba=[c * 0.9 for c in self.bg_color], d=0.1).start(self.bg_color_instruction)
        return super().on_touch_down(touch)

    def on_touch_up(self, touch):
        Animation(rgba=self.bg_color, d=0.2).start(self.bg_color_instruction)
        return super().on_touch_up(touch)

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
            size_hint=(0.5, 0.5),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5),
            **kwargs
        )
        self.table_screen = table_screen
        self.row_index = row_index
        self.current_field = 0
        self.fields = [
            ('Порода', 1),
            ('ж/ф', 2),
            ('шт/либо лет', 3),
            ('D, см', 4),
            ('H, м', 5),
            ('Сост-е', 6),
            ('Модель', 7),
            ('Примечания', 8)
        ]
        self.data = {}
        self.show_next_field()

    def show_next_field(self):
        if self.current_field >= len(self.fields):
            self.save_data()
            return

        field_name, col_index = self.fields[self.current_field]
        content = FloatLayout()

        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=(content.x-10, content.y-10), size=(content.width+20, content.height+20), radius=[30])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=self.dismiss)

        label = Label(
            text=f'Введите {field_name}:',
            font_name='Roboto',
            font_size='18sp',
            color=(0.2, 0.2, 0.2, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.65},
            size_hint=(None, None),
            size=(200, 50)
        )

        self.input_field = AutoCompleteTextInput(
            multiline=False,
            size_hint=(None, None),
            size=(200, 40),
            pos_hint={'center_x': 0.5, 'center_y': 0.45},
            background_color=(1, 1, 1, 0.8),
            col_index=col_index
        )
        self.input_field.bind(on_text_validate=self.next_field)

        btn = ModernButton(
            text='Далее',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(None, None),
            size=(100, 40),
            pos_hint={'center_x': 0.5, 'center_y': 0.25},
            no_shadow=True
        )
        btn.bind(on_press=self.next_field)

        content.add_widget(close_btn)
        content.add_widget(label)
        content.add_widget(self.input_field)
        content.add_widget(btn)

        self.content = content
        self.open()
        Clock.schedule_once(lambda dt: setattr(self.input_field, 'focus', True), 0.1)

    def next_field(self, instance=None):
        value = self.input_field.text.strip()
        if value:
            field_name, col_index = self.fields[self.current_field]
            self.data[col_index] = value
            # Save to suggestions
            self.save_to_suggestions(col_index, value)
        self.current_field += 1
        self.dismiss()
        if self.current_field < len(self.fields):
            next_popup = TreeDataInputPopup(self.table_screen, self.row_index)
            next_popup.current_field = self.current_field
            next_popup.data = self.data
            next_popup.show_next_field()
        else:
            self.save_data()

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

    def save_data(self):
        # Fill the row in the table
        for col_index, value in self.data.items():
            if col_index < len(self.table_screen.inputs[self.row_index]):
                self.table_screen.inputs[self.row_index][col_index].text = value

        # Save to page_data
        self.table_screen.save_current_page()

        # Auto-fill next tree numbers
        base_number = self.table_screen.current_page * self.table_screen.rows_per_page + self.row_index + 1
        for row_idx in range(self.row_index + 1, len(self.table_screen.inputs)):
            tree_number = base_number + (row_idx - self.row_index)
            self.table_screen.inputs[row_idx][0].text = str(tree_number)

        # Show success
        self.table_screen.show_success("Данные дерева сохранены!")

class Joypad(FloatLayout):
    def __init__(self, table_screen, **kwargs):
        super().__init__(**kwargs)
        self.table_screen = table_screen
        self.size_hint = (None, None)
        self.size = (140, 140)
        self.pos_hint = {'right': 0.95, 'y': 0.02}
        
        with self.canvas.before:
            Color(0.2, 0.2, 0.2, 1)
            self.bg_rect = RoundedRectangle(
                pos=self.pos,
                size=self.size,
                radius=[70]
            )
        
        arrows = [
            ('▲', (0.5, 0.7), 'up', (60, 40)),
            ('▼', (0.5, 0.3), 'down', (60, 40)),
            ('◄', (0.3, 0.5), 'left', (40, 60)),
            ('►', (0.7, 0.5), 'right', (40, 60))
        ]
        
        for symbol, pos, direction, size in arrows:
            btn = ModernButton(
                text=symbol,
                size_hint=(None, None),
                size=size,
                pos_hint={'center_x': pos[0], 'center_y': pos[1]},
                bg_color=(0.1, 0.1, 0.1, 1),
                color=(1, 1, 1, 1),
                font_size='20sp',
                bold=True
            )
            btn.bind(on_press=lambda x, d=direction: self.move_focus(d))
            self.add_widget(btn)

        self.bind(pos=self.update_bg, size=self.update_bg)

    def update_bg(self, *args):
        self.bg_rect.pos = self.pos
        self.bg_rect.size = self.size

    def move_focus(self, direction):
        current = self.table_screen.focused_cell
        if not current: return
        row, col = current
        
        if direction == 'up': row = max(0, row-1)
        elif direction == 'down': row = min(len(self.table_screen.inputs)-1, row+1)
        elif direction == 'left': col = max(0, col-1)
        elif direction == 'right': col = min(8, col+1)
        
        self.table_screen.focused_cell = [row, col]
        inp = self.table_screen.inputs[row][col]
        inp.focus = True
        inp.cursor = (len(inp.text), 0)
        Clock.schedule_once(lambda dt: self._update_cursor(inp), 0.01)

    def _update_cursor(self, inp):
        inp.focus = True
        inp.cursor = (len(inp.text), 0)
        inp.text = inp.text

class ExitConfirmPopup(Popup):
    def __init__(self, **kwargs):
        super().__init__(
            title='',
            separator_height=0,
            size_hint=(0.4, 0.3),
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )

        content = FloatLayout()

        with content.canvas.before:
            Color(0.98, 0.98, 0.98, 1)
            RoundedRectangle(
                pos=(content.x-10, content.y-10),
                size=(content.width+20, content.height+20),
                radius=[50]
            )

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=self.dismiss)

        label = Label(
            text='Вы уверены, что хотите выйти?',
            font_name='Roboto',
            font_size='18sp',
            color=(0.2, 0.2, 0.2, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.65},
            size_hint=(None, None),
            size=(250, 50)
        )

        btn_box = BoxLayout(
            orientation='horizontal',
            spacing=15,
            size_hint=(None, None),
            size=(200, 45),
            pos_hint={'center_x': 0.5, 'center_y': 0.25}
        )
        yes_btn = ModernButton(
            text='Выход',
            bg_color=(1, 0, 0, 1),
            color=(0, 0, 0, 1),
            size_hint=(0.5, None),
            height=45,
            no_shadow=True
        )
        no_btn = ModernButton(
            text='Отмена',
            bg_color=(1, 0, 0, 1),
            color=(0, 0, 0, 1),
            size_hint=(0.5, None),
            height=45,
            no_shadow=True
        )
        yes_btn.bind(on_press=lambda x: App.get_running_app().stop())
        no_btn.bind(on_press=self.dismiss)

        btn_box.add_widget(yes_btn)
        btn_box.add_widget(no_btn)

        content.add_widget(close_btn)
        content.add_widget(label)
        content.add_widget(btn_box)

        self.content = content

class MainMenu(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bg_image = None
        self.bg_rect = None
        self.setup_ui()
        
    def setup_ui(self):
        self.clear_widgets()
        theme = App.get_running_app().theme_manager.current_theme
        
        with self.canvas.before:
            if theme['type'] == 'color':
                Color(rgba=theme['background'])
                self.bg_rect = Rectangle(pos=self.pos, size=self.size)
            else:
                self.bg_image = CoreImage(theme['background']).texture
                self.bg_rect = Rectangle(
                    texture=self.bg_image,
                    pos=self.pos,
                    size=self.size
                )
            
        self.bind(pos=self.update_bg, size=self.update_bg)
        
        main_layout = BoxLayout(orientation='vertical', padding=[50, 20, 20, 20])

        center_layout = BoxLayout(
            orientation='vertical',
            size_hint=(0.8, None),
            height=400,
            pos_hint={'center_x': 0.5, 'center_y': 0.75},
            spacing=15
        )
        
        title = Label(
            text='Фанаты Пихты',
            font_size='24sp',
            font_name='Roboto',
            size_hint_y=None,
            height=60
        )
        center_layout.add_widget(title)
        
        buttons = [
            ('Сменить тему', '#00FF00', self.change_theme),
            ('Перечётная ведомость', '#FFA500', self.show_add_section),
            ('РУМ (Молодняки)', '#00BFFF', self.show_add_molodniki_section),
            ('Выбрать тему', '#FFFF00', self.show_theme_chooser),
            ('Выход', '#FF0000', self.confirm_exit)
        ]
        
        for text, color, callback in buttons:
            btn = ModernButton(
                text=text,
                bg_color=get_color_from_hex(color),
                color=get_color_from_hex('#000000'),
                size_hint=(None, None),
                size=(250, 50),
                pos_hint={'center_x': 0.5}
            )
            btn.bind(on_press=callback)
            center_layout.add_widget(btn)
        
        main_layout.add_widget(center_layout)
        
        footer = Label(
            text='by forestcrypt®',
            font_size='12sp',
            color=(1, 1, 1, 1),
            size_hint=(1, None),
            height=30,
            pos_hint={'right': 0.98, 'y': 0.02}
        )
        main_layout.add_widget(footer)
        
        self.add_widget(main_layout)
        
    def update_bg(self, *args):
        if self.bg_rect:
            self.bg_rect.pos = self.pos
            self.bg_rect.size = self.size
            
    def change_theme(self, instance):
        themes = App.get_running_app().theme_manager.themes
        current_idx = App.get_running_app().theme_manager.current_theme_index
        new_idx = (current_idx + 1) % len(themes)
        App.get_running_app().theme_manager.current_theme_index = new_idx
        App.get_running_app().theme_manager.save_config()
        App.get_running_app().reload_theme()

    def show_theme_chooser(self, instance):
        ThemeChooser().open()

    def show_add_section(self, instance):
        content = FloatLayout()

        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=(content.x-10, content.y-10), size=(content.width+20, content.height+20), radius=[30])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=lambda x: self.section_popup.dismiss())

        self.section_number_input = TextInput(
            hint_text="Введите номер участка",
            multiline=False,
            size_hint=(None, None),
            size=(250, 40),
            pos_hint={'center_x': 0.5, 'center_y': 0.6},
            background_color=(1, 1, 1, 0.8)
        )

        btn_box = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint=(None, None),
            size=(250, 50),
            pos_hint={'center_x': 0.5, 'center_y': 0.3}
        )
        buttons = [
            ('Сохранить', '#00FF00', self.save_section),
            ('Загрузить', '#0000FF', self.show_load_popup),
            ('Отмена', '#FF0000', lambda x: self.section_popup.dismiss())
        ]

        for text, color, callback in buttons:
            btn = ModernButton(
                text=text,
                bg_color=get_color_from_hex(color),
                size_hint=(0.33, None),
                height=45,
                no_shadow=False
            )
            btn.bind(on_press=callback)
            btn_box.add_widget(btn)

        content.add_widget(close_btn)
        content.add_widget(self.section_number_input)
        content.add_widget(btn_box)

        self.section_popup = Popup(
            title="Управление участками",
            content=content,
            size_hint=(0.6, 0.5),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )
        self.section_popup.open()

    def add_section(self, instance):
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('INSERT INTO sections DEFAULT VALUES')
            conn.commit()
            self.show_success("Новый участок создан!")
        except Exception as e:
            self.show_error(f"Ошибка: {str(e)}")
        finally:
            conn.close()

    def save_section(self, instance):
        section_number = self.section_number_input.text.strip()
        if not section_number:
            self.show_error("Введите номер участка!")
            return
            
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE sections 
                SET section_number = ? 
                WHERE id = (SELECT MAX(id) FROM sections)
            ''', (section_number,))
            conn.commit()
            self.section_popup.dismiss()
            App.get_running_app().root.get_screen('table').current_section = section_number
            App.get_running_app().root.get_screen('table').update_section_label()
            App.get_running_app().root.current = 'table'
        except Exception as e:
            self.show_error(f"Ошибка сохранения: {str(e)}")
        finally:
            conn.close()

    def show_add_molodniki_section(self, instance):
        content = FloatLayout()

        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=(content.x-10, content.y-10), size=(content.width+20, content.height+20), radius=[30])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=lambda x: self.molodniki_popup.dismiss())

        self.molodniki_section_input = TextInput(
            hint_text="Введите номер участка молодняков",
            multiline=False,
            size_hint=(None, None),
            size=(300, 40),
            pos_hint={'center_x': 0.5, 'center_y': 0.6},
            background_color=(1, 1, 1, 0.8)
        )

        btn_box = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint=(None, None),
            size=(300, 50),
            pos_hint={'center_x': 0.5, 'center_y': 0.3}
        )
        buttons = [
            ('Загрузить', '#00FF00', self.show_load_molodniki_popup),
            ('Сохранить', '#0000FF', self.save_molodniki_section),
            ('Отмена', '#FF0000', lambda x: self.molodniki_popup.dismiss())
        ]

        for text, color, callback in buttons:
            btn = ModernButton(
                text=text,
                bg_color=get_color_from_hex(color),
                size_hint=(0.33, None),
                height=45,
                no_shadow=False
            )
            btn.bind(on_press=callback)
            btn_box.add_widget(btn)

        content.add_widget(close_btn)
        content.add_widget(self.molodniki_section_input)
        content.add_widget(btn_box)

        self.molodniki_popup = Popup(
            title="Управление участками молодняков",
            content=content,
            size_hint=(0.6, 0.5),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )
        self.molodniki_popup.open()

    def add_molodniki_section(self, instance):
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('INSERT INTO molodniki_sections DEFAULT VALUES')
            conn.commit()
            self.show_success("Новый участок молодняков создан!")
        except Exception as e:
            self.show_error(f"Ошибка: {str(e)}")
        finally:
            conn.close()

    def save_molodniki_section(self, instance):
        section_number = self.molodniki_section_input.text.strip()
        if not section_number:
            self.show_error("Введите номер участка!")
            return
            
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE molodniki_sections 
                SET section_number = ? 
                WHERE id = (SELECT MAX(id) FROM molodniki_sections)
            ''', (section_number,))
            conn.commit()
            self.molodniki_popup.dismiss()
            App.get_running_app().root.get_screen('molodniki').current_section = section_number
            App.get_running_app().root.get_screen('molodniki').update_section_label()
            App.get_running_app().root.current = 'molodniki'
        except Exception as e:
            self.show_error(f"Ошибка сохранения: {str(e)}")
        finally:
            conn.close()

    def confirm_exit(self, instance):
        ExitConfirmPopup().open()

    def show_success(self, message):
        Popup(
            title='Успешно',
            content=Label(
                text=message, 
                color=(0, 0.5, 0, 1), 
                font_name='Roboto'
            ),
            size_hint=(0.6, 0.3)
        ).open()

    def show_error(self, message):
        Popup(
            title='Ошибка',
            content=Label(
                text=message, 
                color=(1, 0, 0, 1), 
                font_name='Roboto'
            ),
            size_hint=(0.6, 0.3)
        ).open()

    def show_load_popup(self, instance):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('SELECT section_number FROM sections WHERE section_number IS NOT NULL AND section_number != "" ORDER BY id DESC')
        sections = cursor.fetchall()
        conn.close()
        if not sections:
            self.show_error("Нет сохраненных участков!")
            return
        content = FloatLayout()
        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=(content.x-10, content.y-10), size=(content.width+20, content.height+20), radius=[30])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=lambda x: self.load_popup.dismiss())

        scroll = ScrollView(size_hint=(1, 0.9), pos_hint={'center_x': 0.5, 'center_y': 0.45})
        layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        for section in sections:
            btn = ModernButton(
                text=section[0],
                size_hint_y=None,
                height=40,
                bg_color=(0, 1, 0, 1),
                color=(0, 0, 0, 1),
                no_shadow=True
            )
            btn.bind(on_release=lambda b, s=section[0]: self.load_saved_section(s))
            layout.add_widget(btn)
        scroll.add_widget(layout)

        content.add_widget(close_btn)
        content.add_widget(scroll)
        self.load_popup = Popup(
            title="Выберите участок",
            content=content,
            size_hint=(0.5, 0.6),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )
        self.load_popup.open()

    def load_saved_section(self, section_number):
        files = glob.glob(os.path.join(App.get_running_app().root.get_screen('table').reports_dir, f"{section_number}_*.xlsx"))
        if files:
            latest_file = max(files, key=os.path.getctime)
            table_screen = App.get_running_app().root.get_screen('table')
            try:
                df = pd.read_excel(latest_file)
                data = df.values.tolist()
                
                table_screen.current_section = section_number
                table_screen.update_section_label()
                table_screen.page_data.clear()
                
                for page_num in range(0, len(df), table_screen.rows_per_page):
                    page = page_num // table_screen.rows_per_page
                    page_data = df.iloc[page_num:page_num+table_screen.rows_per_page].values.tolist()
                    table_screen.page_data[page] = page_data
                
                table_screen.current_page = 0
                table_screen.load_page_data()
                table_screen.update_pagination()
                self.show_success("Данные участка загружены!")
                App.get_running_app().root.current = 'table'
            except Exception as e:
                self.show_error(f"Ошибка загрузки: {str(e)}")
        else:
            self.show_error("Файл участка не найден!")
        self.load_popup.dismiss()

    def show_load_molodniki_popup(self, instance):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('SELECT section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != "" ORDER BY id DESC')
        sections = cursor.fetchall()
        conn.close()
        if not sections:
            self.show_error("Нет сохраненных участков молодняков!")
            return
        content = FloatLayout()
        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=(content.x-10, content.y-10), size=(content.width+20, content.height+20), radius=[30])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=lambda x: self.load_molodniki_popup.dismiss())

        scroll = ScrollView(size_hint=(1, 0.9), pos_hint={'center_x': 0.5, 'center_y': 0.45})
        layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        for section in sections:
            btn = ModernButton(
                text=section[0],
                size_hint_y=None,
                height=40,
                bg_color=(0, 1, 0, 1),
                color=(0, 0, 0, 1),
                no_shadow=True
            )
            btn.bind(on_release=lambda b, s=section[0]: self.load_saved_molodniki_section(s))
            layout.add_widget(btn)
        scroll.add_widget(layout)

        content.add_widget(close_btn)
        content.add_widget(scroll)
        self.load_molodniki_popup = Popup(
            title="Выберите участок молодняков",
            content=content,
            size_hint=(0.5, 0.6),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )
        self.load_molodniki_popup.open()

    def load_saved_molodniki_section(self, section_number):
        molodniki_screen = App.get_running_app().root.get_screen('molodniki')
        files = glob.glob(os.path.join(molodniki_screen.reports_dir, f"Молодняки_расширенный_{section_number}_*.xlsx"))
        if files:
            latest_file = max(files, key=os.path.getctime)
            try:
                df = pd.read_excel(latest_file)

                molodniki_screen.current_section = section_number
                molodniki_screen.update_section_label()
                molodniki_screen.page_data.clear()

                for page_num in range(0, len(df), molodniki_screen.rows_per_page):
                    page = page_num // molodniki_screen.rows_per_page
                    page_data = df.iloc[page_num:page_num+molodniki_screen.rows_per_page].values.tolist()
                    # Дополняем до 29 столбцов если нужно
                    for row in page_data:
                        while len(row) < 29:
                            row.append('')
                    molodniki_screen.page_data[page] = page_data

                molodniki_screen.current_page = 0
                molodniki_screen.load_page_data()
                molodniki_screen.update_pagination()
                self.molodniki_popup.dismiss()
                self.show_success("Данные участка молодняков загружены!")
                App.get_running_app().root.current = 'molodniki'
            except Exception as e:
                self.show_error(f"Ошибка загрузки: {str(e)}")
        else:
            self.show_error("Файл участка молодняков не найден!")

class TableScreen(Screen):
    current_page = NumericProperty(0)
    total_pages = NumericProperty(1)
    unsaved_changes = BooleanProperty(False)
    focused_cell = ListProperty([0, 0])
    edit_mode = BooleanProperty(False)
    current_section = StringProperty("")
    MAX_PAGES = 200

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.theme_manager = App.get_running_app().theme_manager
        self.reports_dir = "reports"
        os.makedirs(self.reports_dir, exist_ok=True)
        self.db_name = 'forest_data.db'
        self.rows_per_page = 50
        self.page_data = {}
        self.setup_database()
        self.create_ui()
        self.load_existing_data()
        Window.bind(on_key_down=self.key_action)

    def key_action(self, window, key, scancode, codepoint, modifier):
        if key == 115 and 'ctrl' in modifier:
            self.save_current_page()

    def setup_database(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS trees (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        tree_number INTEGER,
                        species TEXT,
                        age TEXT,
                        count TEXT,
                        diameter REAL,
                        height REAL,
                        condition TEXT,
                        model TEXT,
                        notes TEXT,
                        section_id INTEGER,
                        FOREIGN KEY(section_id) REFERENCES sections(id))''')
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS sections (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        section_number TEXT UNIQUE,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS suggestions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        column_index INTEGER,
                        value TEXT,
                        UNIQUE(column_index, value))''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_suggestions ON suggestions (column_index, value)')
        conn.commit()
        conn.close()

    def create_ui(self):
        main_layout = BoxLayout(orientation='horizontal', padding=10, spacing=10)
        
        with self.canvas.before:
            self.bg_color = Color(1, 1, 1, 1)
            self.bg_rect = Rectangle(pos=self.pos, size=self.size)
            self.bind(pos=self._update_bg, size=self._update_bg)
        
        self._update_background(self.theme_manager.current_theme)

        # Табличная часть (левая панель)
        table_panel = BoxLayout(orientation='vertical', size_hint_x=0.7)
        
        # Заголовок участка
        header_layout = BoxLayout(orientation='vertical', size_hint=(1, None), height=30)
        self.section_label = Label(
            text=f"Участок: {self.current_section}", 
            font_name='Roboto',
            size_hint=(1, None),
            height=30,
            color=self._get_text_color()
        )
        header_layout.add_widget(self.section_label)
        table_panel.add_widget(header_layout)
        
        # Пагинация
        pagination = BoxLayout(size_hint_y=None, height=40, spacing=5)
        self.page_label = Label(
            text=f'Страница {self.current_page+1} из {self.total_pages}', 
            size_hint_x=0.4, 
            font_name='Roboto',
            color=self._get_text_color()
        )
        prev_btn = ModernButton(
            text='← Предыдущая', 
            size_hint_x=0.3,
            bg_color=get_color_from_hex('#00FF00'),
            color=self._get_text_color()
        )
        prev_btn.bind(on_press=lambda x: self.change_page(-1))
        next_btn = ModernButton(
            text='Следующая →', 
            size_hint_x=0.3,
            bg_color=get_color_from_hex('#00FF00'),
            color=self._get_text_color()
        )
        next_btn.bind(on_press=lambda x: self.change_page(1))
        pagination.add_widget(prev_btn)
        pagination.add_widget(self.page_label)
        pagination.add_widget(next_btn)
        table_panel.add_widget(pagination)
        
        # Основная таблица
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True, bar_width=10)
        self.table = GridLayout(cols=9, size_hint=(None, None), spacing=2)
        self.table.bind(minimum_height=self.table.setter('height'), 
                       minimum_width=self.table.setter('width'))
        
        # Заголовки столбцов
        headers = ['№ дерева*', 'Порода*', 'ж/ф', 'шт/либо лет',
                 'D, см*', 'H, м', 'Сост-е', 'Модель', 'Примечания']
        self.header_bgs = []
        for header in headers:
            lbl = Label(
                text=header,
                size_hint_y=None,
                height=30,
                font_name='Roboto',
                bold=True,
                halign='center',
                size_hint_x=None,
                width=100,
                color=get_color_from_hex('#000000')
            )
            with lbl.canvas.before:
                Color(rgba=get_color_from_hex('#00FF00'))
                bg = Rectangle(pos=lbl.pos, size=lbl.size)
                self.header_bgs.append(bg)
            lbl.bind(pos=lambda i,v, b=bg: setattr(b, 'pos', i.pos), size=lambda i,v, b=bg: setattr(b, 'size', i.size))
            self.table.add_widget(lbl)
        
        # Создаем строки таблицы
        self.inputs = []
        for row_idx in range(self.rows_per_page):
            row = []
            for col_idx in range(9):
                inp = AutoCompleteTextInput(multiline=False, size_hint_y=None, height=30)
                inp.row_index = row_idx
                inp.col_index = col_idx
                inp.bind(focus=self.update_focus)
                inp.font_name = 'Roboto'
                
                if col_idx > 0:
                    inp.prev_widget = row[col_idx-1] if row else None
                    if row:
                        row[col_idx-1].next_widget = inp
                if row_idx > 0:
                    inp.prev_widget = self.inputs[row_idx-1][col_idx] if self.inputs else None
                    if self.inputs:
                        self.inputs[row_idx-1][col_idx].next_widget = inp
                
                if col_idx in [0,4,5]:
                    inp.input_filter = 'float' if col_idx in [4,5] else 'int'
                if col_idx == 0:
                    inp.bind(focus=self.show_tree_popup)
                row.append(inp)
                self.table.add_widget(inp)
            self.inputs.append(row)
        
        scroll.add_widget(self.table)
        table_panel.add_widget(scroll)
        main_layout.add_widget(table_panel)
        
        # Правая панель управления
        control_panel = BoxLayout(
            orientation='vertical', 
            size_hint_x=0.3,
            spacing=15,
            padding=[0, 10, 0, 0]
        )
        
        # Основные кнопки управления
        controls = GridLayout(
            cols=1,
            size_hint_y=None,
            height=350,
            spacing=10,
            pos_hint={'top': 1}
        )
        
        button_handlers = {
            'Сохранить отчет': self.show_save_dialog,
            'Сохранить страницу': self.save_current_page,
            'Загрузить участок': self.load_section,
            'Открыть папку': self.open_excel_file,
            'Редакт. режим': self.toggle_edit_mode,
            'Очистить данные': self.clear_table_data,
            'В главное меню': self.go_back
        }
        
        button_colors = {
            'Сохранить отчет': '#00FF00',
            'Сохранить страницу': '#00FFFF',
            'Загрузить участок': '#006400',
            'Открыть папку': '#0000FF',
            'Редакт. режим': '#FFA500',
            'Очистить данные': '#800000',
            'В главное меню': '#FF0000'
        }
        
        for text, color in button_colors.items():
            btn = ModernButton(
                text=text,
                bg_color=get_color_from_hex(color),
                size_hint=(None, None),
                size=(220, 45),
                pos_hint={'center_x': 0.5}
            )
            btn.bind(on_press=button_handlers[text])
            controls.add_widget(btn)
        
        control_panel.add_widget(controls)
        
        # Джойстик - центрируем внизу
        joypad_container = BoxLayout(
            size_hint=(1, None), 
            height=150,
            padding=[0, 20, 0, 0]
        )
        
        joypad = Joypad(self)
        joypad.pos_hint = {'center_x': 0.5, 'center_y': 0.5}
        joypad_container.add_widget(joypad)
        
        spacer = BoxLayout(size_hint_y=1)
        control_panel.add_widget(spacer)
        control_panel.add_widget(joypad_container)
        
        main_layout.add_widget(control_panel)
        self.add_widget(main_layout)

    def _update_bg(self, instance, value):
        self.bg_rect.pos = self.pos
        self.bg_rect.size = self.size

    def _update_background(self, theme):
        if theme['type'] == 'image':
            try:
                self.bg_color.rgba = (1, 1, 1, 1)
                self.bg_rect.texture = CoreImage(theme['background']).texture
            except Exception as e:
                print(f"Error loading background image: {str(e)}")
        else:
            self.bg_color.rgba = theme['background']
            self.bg_rect.texture = None

    def _get_text_color(self):
        theme = self.theme_manager.current_theme
        if theme['type'] == 'image':
            return get_color_from_hex('#FFFFFF')
        else:
            return get_color_from_hex(theme['text_color'])

    def update_section_label(self):
        self.section_label.text = f"Участок: {self.current_section}"

    def toggle_edit_mode(self, instance):
        self.edit_mode = not self.edit_mode
        instance.bg_color = get_color_from_hex('#FFA500' if self.edit_mode else '#00FF00')

    def update_focus(self, instance, value):
        if value:
            self.focused_cell = [instance.row_index, instance.col_index]

    def show_tree_popup(self, instance, value):
        if value and instance.text.strip():
            if not self.edit_mode:
                # In normal mode, only show popup if other columns are empty
                if not any(inp.text.strip() for inp in self.inputs[instance.row_index][1:]):
                    self.save_suggestion(0, instance.text.strip())
                    TreeDataInputPopup(self, instance.row_index).open()
            else:
                # In edit mode, always show popup for editing existing data
                self.save_suggestion(0, instance.text.strip())
                TreeDataInputPopup(self, instance.row_index).open()

    def save_suggestion(self, col_index, value):
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR IGNORE INTO suggestions (column_index, value)
                VALUES (?, ?)
            ''', (col_index, value))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Error saving suggestion: {e}")



    def save_current_page(self, instance=None):
        page_data = []
        for row in self.inputs:
            page_data.append([inp.text for inp in row])
        self.page_data[self.current_page] = page_data

    def show_save_dialog(self, instance=None):
        content = FloatLayout()

        with content.canvas.before:
            Color(rgba=(0.9, 0.9, 0.9, 1))
            RoundedRectangle(pos=content.pos, size=content.size, radius=[20])

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=lambda x: self.save_popup.dismiss())

        label = Label(
            text="Введите имя файла:",
            font_name='Roboto',
            font_size='18sp',
            color=(0.2, 0.2, 0.2, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.7},
            size_hint=(None, None),
            size=(200, 50)
        )

        self.filename_input = TextInput(
            hint_text="Имя файла",
            multiline=False,
            size_hint=(None, None),
            size=(250, 40),
            pos_hint={'center_x': 0.5, 'center_y': 0.5},
            background_color=(1, 1, 1, 0.8)
        )
        default_name = f"{self.current_section}_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}"
        self.filename_input.text = default_name

        btn_box = BoxLayout(
            orientation='horizontal',
            spacing=5,
            size_hint=(None, None),
            size=(250, 50),
            pos_hint={'center_x': 0.5, 'center_y': 0.3}
        )
        ok_btn = ModernButton(
            text="Сохранить",
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, None),
            height=40,
            no_shadow=True
        )
        cancel_btn = ModernButton(
            text="Отмена",
            bg_color=get_color_from_hex('#FF0000'),
            size_hint=(0.5, None),
            height=40,
            no_shadow=True
        )
        btn_box.add_widget(ok_btn)
        btn_box.add_widget(cancel_btn)

        content.add_widget(close_btn)
        content.add_widget(label)
        content.add_widget(self.filename_input)
        content.add_widget(btn_box)

        self.save_popup = Popup(
            title="Сохранение отчета",
            content=content,
            size_hint=(0.6, 0.5),
            separator_height=0,
            background='',
            overlay_color=(0, 0, 0, 0.5)
        )
        ok_btn.bind(on_press=self.save_to_excel)
        cancel_btn.bind(on_press=self.save_popup.dismiss)
        self.save_popup.open()

    def save_to_excel(self, instance):
        filename = self.filename_input.text.strip()
        if not filename:
            self.show_error("Имя файла не может быть пустым!")
            return
        filename = re.sub(r'[\\/*?:"<>|]', "", filename)
        filename = f"{filename}.xlsx" if not filename.endswith(".xlsx") else filename
        full_path = os.path.join(self.reports_dir, filename)
        
        try:
            all_data = []
            for page in sorted(self.page_data.keys()):
                all_data.extend(self.page_data[page])
            df = pd.DataFrame(
                all_data,
                columns=['№ дерева','Порода','ж/ф','шт/либо лет','D, см','H, м','Сост-е','Модель','Примечания']
            )
            df.to_excel(full_path, index=False)
            self.save_popup.dismiss()
            self.show_success(f"Файл сохранен: {filename}")
        except Exception as e:
            self.show_error(f"Ошибка: {str(e)}")

    def load_section(self, instance):
        Tk().withdraw()
        file_path = filedialog.askopenfilename(
            initialdir=self.reports_dir,
            title="Выберите файл участка",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                data = df.values.tolist()
                
                self.current_section = os.path.splitext(os.path.basename(file_path))[0]
                self.update_section_label()
                self.page_data.clear()
                
                for page_num in range(0, len(df), self.rows_per_page):
                    page = page_num // self.rows_per_page
                    page_data = df.iloc[page_num:page_num+self.rows_per_page].values.tolist()
                    self.page_data[page] = page_data
                
                self.current_page = 0
                self.load_page_data()
                self.update_pagination()
                self.show_success("Данные успешно загружены!")
            except Exception as e:
                self.show_error(f"Ошибка загрузки: {str(e)}")

    def load_page_data(self):
        for row in self.inputs:
            for inp in row:
                inp.text = ''
        
        if self.current_page in self.page_data:
            for i, row_data in enumerate(self.page_data[self.current_page]):
                if i >= len(self.inputs):
                    break
                for j, text in enumerate(row_data):
                    if j < len(self.inputs[i]):
                        self.inputs[i][j].text = str(text) if not pd.isna(text) else ''

    def clear_table_data(self, instance=None):
        for row in self.inputs:
            for inp in row:
                inp.text = ''
        self.page_data.clear()
        self.show_success("Данные очищены!")

    def open_excel_file(self, instance):
        if os.path.exists(self.reports_dir):
            os.startfile(self.reports_dir)
        else:
            self.show_error("Папка reports не найдена!")

    def change_page(self, delta):
        new_page = self.current_page + delta
        if 0 <= new_page < self.MAX_PAGES:
            self.current_page = new_page
            self.load_page_data()
            self.update_page_label()

    def update_pagination(self):
        self.total_pages = len(self.page_data) if self.page_data else 1
        self.total_pages = min(self.total_pages, self.MAX_PAGES)
        self.update_page_label()

    def update_page_label(self):
        self.page_label.text = f'Страница {self.current_page+1} из {self.total_pages}'

    def go_back(self, instance):
        App.get_running_app().root.current = 'main'

    def show_error(self, message):
        Popup(
            title='Ошибка',
            content=Label(text=message, color=(1, 0, 0, 1)),
            size_hint=(0.6, 0.3)
        ).open()

    def show_success(self, message):
        Popup(
            title='Успешно',
            content=Label(text=message, color=(0, 0.5, 0, 1)),
            size_hint=(0.6, 0.3)
        ).open()
        
    def load_existing_data(self):
        pass

class ThemeChooser(Popup):
    def __init__(self, **kwargs):
        super().__init__(title='Выбор темы', size_hint=(0.7, 0.6))
        content = FloatLayout()

        close_btn = ModernButton(
            text='X',
            size_hint=(None, None),
            size=(40, 40),
            pos_hint={'right': 0.95, 'top': 0.95},
            bg_color=(1, 0, 0, 1),
            no_shadow=True
        )
        close_btn.bind(on_press=self.dismiss)

        self.layout = GridLayout(cols=3, spacing=10, padding=10, size_hint=(1, 0.9), pos_hint={'center_x': 0.5, 'center_y': 0.45})
        self.scroll = ScrollView()

        for theme in App.get_running_app().theme_manager.themes:
            btn = Button(size_hint=(None, None), size=(200, 200))
            if theme['type'] == 'color':
                btn.background_color = theme['background']
            else:
                btn.background_normal = theme['background']
            btn.bind(on_release=lambda x, t=theme: self.select_theme(t))
            self.layout.add_widget(btn)

        add_btn = Button(text='Добавить тему', size_hint=(None, None), size=(200, 200))
        add_btn.bind(on_release=self.add_theme)
        self.layout.add_widget(add_btn)

        self.scroll.add_widget(self.layout)
        content.add_widget(close_btn)
        content.add_widget(self.scroll)
        self.content = content

    def select_theme(self, theme):
        manager = App.get_running_app().theme_manager
        manager.current_theme_index = manager.themes.index(theme)
        manager.save_config()
        App.get_running_app().reload_theme()
        self.dismiss()

    def add_theme(self, instance):
        Tk().withdraw()
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )
        if file_path:
            App.get_running_app().theme_manager.add_theme(file_path)
            self.layout.clear_widgets()
            for theme in App.get_running_app().theme_manager.themes:
                btn = Button(size_hint=(None, None), size=(200, 200))
                if theme['type'] == 'color':
                    btn.background_color = theme['background']
                else:
                    btn.background_normal = theme['background']
                btn.bind(on_release=lambda x, t=theme: self.select_theme(t))
                self.layout.add_widget(btn)
            add_btn = Button(text='Добавить тему', size_hint=(None, None), size=(200, 200))
            add_btn.bind(on_release=self.add_theme)
            self.layout.add_widget(add_btn)


class ForestryApp(App):
    theme_manager = ThemeManager()
    
    def build(self):
        Config.set('graphics', 'multisamples', '4')
        self.apply_theme()
        sm = ScreenManager()
        sm.add_widget(MainMenu(name='main'))
        sm.add_widget(TableScreen(name='table'))
        sm.add_widget(ExtendedMolodnikiTableScreen(name='molodniki'))
        return sm
    
    def apply_theme(self):
        theme = self.theme_manager.current_theme
        if theme['type'] == 'image':
            Window.clearcolor = (1, 1, 1, 1)
        else:
            Window.clearcolor = theme['background']
    
    def reload_theme(self):
        self.apply_theme()
        for screen in self.root.screens:
            if hasattr(screen, 'setup_ui'):
                screen.setup_ui()
            if hasattr(screen, '_update_background'):
                screen._update_background(self.theme_manager.current_theme)

if __name__ == '__main__':
    ForestryApp().run()
