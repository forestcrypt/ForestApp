# Расширенная таблица молодняков по новой структуре
        # Структура: 6 основных столбцов + динамические подстолбцы для пород

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.screenmanager import Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.checkbox import CheckBox
from kivy.properties import (NumericProperty, BooleanProperty,
                          ObjectProperty, ListProperty, StringProperty)
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle, Line, RoundedRectangle
from kivy.clock import Clock
from kivy.animation import Animation
from kivy.core.text import LabelBase
from kivy.utils import get_color_from_hex
from kivy.core.image import Image as CoreImage
import sqlite3
import pandas as pd
import os
import datetime
import re
import json
import sys
import openpyxl
from openpyxl import Workbook
from tkinter import Tk, filedialog

LabelBase.register(name='Roboto',
                 fn_regular='fonts/Roboto-Medium.ttf',
                 fn_bold='fonts/Roboto-Bold.ttf')

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
            SELECT value FROM molodniki_suggestions
            WHERE column_index = ? AND value LIKE ?
            ORDER BY LENGTH(value) ASC, value ASC
            LIMIT 1
        ''', (self.col_index, f'{value}%'))
        results = cursor.fetchall()
        conn.close()

        if results:
            self.text = results[0][0]

    def get_table_screen(self):
        return App.get_running_app().root.get_screen('molodniki')

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
        elif direction == 'right': col = min(5, col+1)  # 6 столбцов (0-5)

        self.table_screen.focused_cell = [row, col]
        inp = self.table_screen.inputs[row][col]
        inp.focus = True
        inp.cursor = (len(inp.text), 0)
        Clock.schedule_once(lambda dt: self._update_cursor(inp), 0.01)

    def _update_cursor(self, inp):
        inp.focus = True
        inp.cursor = (len(inp.text), 0)
        inp.text = inp.text

class MolodnikiTreeDataInputPopup(Popup):
    def __init__(self, table_screen, row_index, **kwargs):
        super().__init__(
            title='Ввод данных площадки молодняков',
            size_hint=(0.8, 0.9),
            separator_height=0,
            background_color=(0.5, 0.5, 0.5, 1),  # Серый фон
            overlay_color=(0, 0, 0, 0.5),
            **kwargs
        )
        self.table_screen = table_screen
        self.row_index = row_index
        self.fields = [
            ('GPS точка', 1),
            ('Предмет ухода', 2),
            ('Порода', 3),
            ('Примечания', 4),
            ('Тип Леса', 5)
        ]
        self.data = {}
        self.create_ui()

    def create_ui(self):
        content = FloatLayout()

        label = Label(
            text='Введите данные площадки молодняков:',
            font_name='Roboto',
            font_size='18sp',
            color=(1, 1, 1, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.95},
            size_hint=(None, None),
            size=(350, 50)
        )

        scroll = ScrollView(size_hint=(0.9, 0.75), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))

        with layout.canvas.before:
            Color(rgba=(0, 0, 0, 0))

        self.input_fields = []
        for field_name, col_index in self.fields:
            field_layout = BoxLayout(orientation='vertical', size_hint_y=None, height=70, spacing=5)
            field_label = Label(
                text=field_name,
                font_name='Roboto',
                font_size='16sp',
                color=(1, 1, 1, 1),
                size_hint_y=None,
                height=20
            )
            input_field = AutoCompleteTextInput(
                multiline=False,
                size_hint_y=None,
                height=40,
                background_color=(1, 1, 1, 0.8),
                col_index=col_index,
                font_name='Roboto'
            )
            if col_index == 3:  # Порода - открываем popup выбора типа
                input_field.bind(focus=self.show_breed_popup)
            self.input_fields.append(input_field)
            field_layout.add_widget(field_label)
            field_layout.add_widget(input_field)
            layout.add_widget(field_layout)

        # Заполняем данными из текущей строки
        if self.table_screen.current_page in self.table_screen.page_data and self.row_index < len(self.table_screen.page_data[self.table_screen.current_page]):
            row_data = self.table_screen.page_data[self.table_screen.current_page][self.row_index]
            for i, (field_name, col_index) in enumerate(self.fields):
                if col_index < len(row_data) and row_data[col_index]:
                    self.input_fields[i].text = str(row_data[col_index])

        scroll.add_widget(layout)

        btn_box = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint=(1, None),
            height=40,
            pos_hint={'center_x': 0.5, 'center_y': 0.1}
        )
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, None),
            height=40,
            no_shadow=True
        )
        save_btn.bind(on_press=self.save_data)
        exit_btn = ModernButton(
            text='Выйти',
            bg_color=get_color_from_hex('#FF0000'),
            size_hint=(0.5, None),
            height=40,
            no_shadow=True
        )
        exit_btn.bind(on_press=self.dismiss)
        btn_box.add_widget(save_btn)
        btn_box.add_widget(exit_btn)

        content.add_widget(label)
        content.add_widget(scroll)
        content.add_widget(btn_box)

        self.content = content
        self.open()

    def show_breed_popup(self, instance, value):
        """Показать popup для выбора типа породы"""
        if not value: return

        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Кнопки выбора типа породы
        type_layout = BoxLayout(orientation='horizontal', spacing=10)
        coniferous_btn = ModernButton(
            text='Хвойные',
            bg_color=get_color_from_hex('#228B22'),
            size_hint=(0.5, None),
            height=50
        )
        deciduous_btn = ModernButton(
            text='Лиственные',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(0.5, None),
            height=50
        )
        type_layout.add_widget(coniferous_btn)
        type_layout.add_widget(deciduous_btn)
        content.add_widget(type_layout)

        # Кнопка отмены
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        content.add_widget(cancel_btn)

        popup = Popup(
            title="Выберите тип породы",
            content=content,
            size_hint=(0.8, 0.5)
        )

        def select_coniferous(btn):
            self.show_breed_selection_popup(instance, 'coniferous')
            popup.dismiss()

        def select_deciduous(btn):
            self.show_breed_selection_popup(instance, 'deciduous')
            popup.dismiss()

        coniferous_btn.bind(on_press=select_coniferous)
        deciduous_btn.bind(on_press=select_deciduous)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_breed_selection_popup(self, instance, breed_type):
        """Показать popup для выбора конкретной породы из словаря"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Заголовок
        title_label = Label(
            text=f"Выберите {'хвойную' if breed_type == 'coniferous' else 'лиственную'} породу",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Списки пород
        if breed_type == 'coniferous':
            breeds = [
                'Сосна ЛК', 'Сосна ЕВ', 'Ель ЛК', 'Ель ЕВ',
                'Пихта ЛК', 'Пихта ЕВ', 'Кедр ЛК', 'Кедр ЕВ',
                'Лиственница ЛК', 'Лиственница ЕВ'
            ]
        else:
            breeds = [
                'Берёза', 'Осина', 'Ольха чёрная', 'Ольха серая',
                'Ива', 'Ива кустарниковая'
            ]

        # ScrollView для списка пород
        scroll = ScrollView(size_hint=(1, None), height=300)
        breeds_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        breeds_layout.bind(minimum_height=breeds_layout.setter('height'))

        for breed in breeds:
            btn = ModernButton(
                text=breed,
                bg_color=get_color_from_hex('#87CEEB'),
                size_hint=(1, None),
                height=50,
                font_size='14sp'
            )
            btn.bind(on_press=lambda x, b=breed: self.select_breed(instance, breed_type, b))
            breeds_layout.add_widget(btn)

        scroll.add_widget(breeds_layout)
        content.add_widget(scroll)

        # Кнопка "Другая порода"
        other_btn = ModernButton(
            text='Другая порода',
            bg_color=get_color_from_hex('#DDA0DD'),
            size_hint=(1, None),
            height=50
        )
        other_btn.bind(on_press=lambda x: self.select_breed(instance, breed_type, 'other'))
        content.add_widget(other_btn)

        # Кнопка отмены
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(1, None),
            height=50
        )
        content.add_widget(cancel_btn)

        popup = Popup(
            title=f"Выбор {'хвойной' if breed_type == 'coniferous' else 'лиственной'} породы",
            content=content,
            size_hint=(0.85, 0.85)
        )

        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def select_breed(self, instance, breed_type, selected_breed):
        """Обработка выбора породы"""
        if selected_breed == 'other':
            # Показываем popup для ввода названия другой породы
            self.show_custom_breed_popup(instance, breed_type)
        else:
            # Показываем popup с параметрами породы, передавая название выбранной породы
            self.show_breed_details_popup(instance, breed_type, selected_breed)

    def show_breed_details_popup(self, instance, breed_type, selected_breed=None):
        """Показать popup для управления множественными породами"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Заголовок
        title_label = Label(
            text=f"Управление породами - {selected_breed}",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Список существующих пород в этой строке
        existing_breeds = self.table_screen.parse_breeds_data(instance.text)
        if not existing_breeds:
            # Пытаемся получить данные из сохраненных данных страницы
            row_idx = self.table_screen.inputs.index([inp for inp in self.table_screen.inputs if inp[3] == instance][0]) if instance in [inp for row in self.table_screen.inputs for inp in row] else -1
            if row_idx >= 0 and self.table_screen.current_page in self.table_screen.page_data and row_idx < len(self.table_screen.page_data[self.table_screen.current_page]):
                saved_text = self.table_screen.page_data[self.table_screen.current_page][row_idx][3]  # Столбец "Порода"
                existing_breeds = self.table_screen.parse_breeds_data(saved_text)

        # Отображаем список существующих пород
        if existing_breeds:
            breeds_list_label = Label(
                text="Уже добавленные породы:",
                font_name='Roboto',
                bold=True,
                size_hint=(1, None),
                height=30,
                color=(0.3, 0.3, 0.3, 1)
            )
            content.add_widget(breeds_list_label)

            # ScrollView для списка пород
            breeds_scroll = ScrollView(size_hint=(1, None), height=80)
            breeds_list_layout = BoxLayout(orientation='vertical', spacing=5, size_hint_y=None)
            breeds_list_layout.bind(minimum_height=breeds_list_layout.setter('height'))

            for i, breed_info in enumerate(existing_breeds):
                breed_name = breed_info.get('name', 'Неизвестная')
                existing_breed_type = breed_info.get('type', 'deciduous')
                params = []

                # В зависимости от типа породы показываем разные параметры
                if existing_breed_type == 'coniferous':
                    # Для хвойных показываем градации
                    if 'do_05' in breed_info and breed_info['do_05']:
                        params.append(f"До 0.5м: {breed_info['do_05']}")
                    if '05_15' in breed_info and breed_info['05_15']:
                        params.append(f"0.5-1.5м: {breed_info['05_15']}")
                    if 'bolee_15' in breed_info and breed_info['bolee_15']:
                        params.append(f">1.5м: {breed_info['bolee_15']}")
                    if 'height' in breed_info and breed_info['height']:
                        params.append(f"Высота: {breed_info['height']}м")
                    if 'age' in breed_info and breed_info['age']:
                        params.append(f"Возраст: {breed_info['age']} лет")
                else:
                    # Для лиственных показываем только основную информацию (без градаций)
                    if 'density' in breed_info and breed_info['density']:
                        params.append(f"Густота: {breed_info['density']}")
                    if 'height' in breed_info and breed_info['height']:
                        params.append(f"Высота: {breed_info['height']}м")
                    if 'age' in breed_info and breed_info['age']:
                        params.append(f"Возраст: {breed_info['age']} лет")

                breed_text = f"{i+1}. {breed_name}: {'; '.join(params)}" if params else f"{i+1}. {breed_name}"
                breed_label = Label(
                    text=breed_text,
                    font_name='Roboto',
                    size_hint=(1, None),
                    height=25,
                    color=(0.2, 0.2, 0.2, 1),
                    text_size=(None, None),
                    halign='left',
                    valign='top'
                )
                breed_label.bind(size=lambda *args: setattr(breed_label, 'text_size', (breed_label.width, None)))
                breeds_list_layout.add_widget(breed_label)

            breeds_scroll.add_widget(breeds_list_layout)
            content.add_widget(breeds_scroll)
        else:
            breeds_list_label = Label(
                text="Породы ещё не добавлены",
                font_name='Roboto',
                size_hint=(1, None),
                height=30,
                color=(0.5, 0.5, 0.5, 1)
            )
            content.add_widget(breeds_list_label)

        # Поля ввода для параметров породы с прокруткой
        scroll_fields = ScrollView(size_hint=(1, None), height=250)
        fields_layout = GridLayout(cols=2, spacing=5, size_hint_y=None)
        fields_layout.bind(minimum_height=fields_layout.setter('height'))

        if breed_type == 'coniferous':
            fields = [
                ('До 0.5м:', 'do_05'),
                ('0.5-1.5м:', '05_15'),
                ('>1.5м:', 'bolee_15'),
                ('Высота (м):', 'height'),
                ('Диаметр (см):', 'diameter'),
                ('Густота:', 'density'),
                ('Возраст (лет):', 'age')
            ]
        else:
            fields = [
                ('Густота:', 'density'),
                ('Высота (м):', 'height'),
                ('Диаметр (см):', 'diameter'),
                ('Возраст (лет):', 'age')
            ]

        self.breed_inputs = {}
        for label_text, field_key in fields:
            lbl = Label(text=label_text, font_name='Roboto', size_hint=(None, None), size=(120, 40), halign='left', valign='middle')
            lbl.bind(size=lambda *args: setattr(lbl, 'text_size', (lbl.width, None)))
            inp = TextInput(multiline=False, size_hint=(None, None), size=(120, 40))
            if field_key in ['density', 'age']:
                inp.input_filter = 'int'
            elif field_key == 'height':
                inp.input_filter = 'float'
            elif field_key in ['do_05', '05_15', 'bolee_15']:
                inp.input_filter = 'int'
                if breed_type == 'coniferous':
                    inp.bind(text=self.update_coniferous_density)
            fields_layout.add_widget(lbl)
            fields_layout.add_widget(inp)
            self.breed_inputs[field_key] = inp

        scroll_fields.add_widget(fields_layout)
        content.add_widget(scroll_fields)

        # Кнопки управления
        btn_layout = BoxLayout(orientation='horizontal', spacing=5, size_hint=(1, None), height=50)
        add_btn = ModernButton(
            text='Добавить породу',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.25, 1),
            height=50
        )
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(0.25, 1),
            height=50
        )
        view_btn = ModernButton(
            text='Просмотр',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(0.25, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.25, 1),
            height=50
        )
        btn_layout.add_widget(add_btn)
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(view_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title=f"Параметры породы: {selected_breed}",
            content=content,
            size_hint=(0.9, 0.95)
        )

        def add_breed(btn):
            breed_data = {
                'name': selected_breed,
                'type': breed_type
            }

            for key, inp in self.breed_inputs.items():
                if inp.text.strip():
                    try:
                        if key in ['density', 'age']:
                            breed_data[key] = int(inp.text)
                        elif key == 'height':
                            breed_data[key] = float(inp.text)
                        else:
                            breed_data[key] = float(inp.text)
                    except ValueError:
                        breed_data[key] = 0 if key in ['density', 'age', 'do_05', '05_15', 'bolee_15'] else 0.0

            existing_breeds = self.table_screen.parse_breeds_data(instance.text)
            existing_breeds.append(breed_data)
            instance.text = json.dumps(existing_breeds, ensure_ascii=False, indent=2)

            # Update page_data so that taxational calculations include the new breed
            if self.table_screen.current_page not in self.table_screen.page_data:
                self.table_screen.page_data[self.table_screen.current_page] = [['', '', '', '', '', ''] for _ in range(self.table_screen.rows_per_page)]
            if instance.row_index < len(self.table_screen.page_data[self.table_screen.current_page]):
                self.table_screen.page_data[self.table_screen.current_page][instance.row_index][3] = instance.text

            self.table_screen.update_plot_total(instance, instance.text)

            for inp in self.breed_inputs.values():
                inp.text = ''

            # После добавления первой породы присваиваем номер 1 и предлагаем выбор
            if len(existing_breeds) == 1:
                self.show_breed_choice_popup(instance, selected_breed)
            else:
                self.show_breed_popup(instance, True)
                self.table_screen.show_success(f"Порода '{selected_breed}' добавлена! Выберите тип следующей породы.")

        def save_breeds(btn):
            existing_breeds = self.table_screen.parse_breeds_data(instance.text)
            if not existing_breeds:
                existing_breeds = []
            instance.text = json.dumps(existing_breeds, ensure_ascii=False, indent=2)

            # Update the main table input to reflect the changes immediately
            self.table_screen.inputs[self.row_index][3].text = instance.text

            # Update page_data
            if self.table_screen.current_page not in self.table_screen.page_data:
                self.table_screen.page_data[self.table_screen.current_page] = [['', '', '', '', '', ''] for _ in range(self.table_screen.rows_per_page)]
            if self.row_index < len(self.table_screen.page_data[self.table_screen.current_page]):
                self.table_screen.page_data[self.table_screen.current_page][self.row_index][3] = instance.text

            self.table_screen.update_plot_total(instance, instance.text)
            self.table_screen.show_success("Все данные по площадке сохранены!")
            popup.dismiss()

        def view_breeds(btn):
            popup.dismiss()
            self.table_screen.show_breeds_list_popup(instance)

        add_btn.bind(on_press=add_breed)
        save_btn.bind(on_press=save_breeds)
        view_btn.bind(on_press=view_breeds)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_breed_choice_popup(self, instance, selected_breed):
        """Показать popup с выбором после добавления первой породы"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text=f"Порода '{selected_breed}' добавлена!\nВыберите действие:",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=60,
            color=(0, 0.5, 0, 1)
        )
        content.add_widget(title_label)

        # Информация о номере породы
        info_label = Label(
            text="Автоматически присвоен номер: 1 порода",
            font_name='Roboto',
            size_hint=(1, None),
            height=30,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(info_label)

        btn_layout = BoxLayout(orientation='vertical', spacing=10, size_hint=(1, None), height=120)
        add_more_btn = ModernButton(
            text='Добавить еще породу',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(1, None),
            height=50
        )
        save_exit_btn = ModernButton(
            text='Сохранить и выйти',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(1, None),
            height=50
        )
        btn_layout.add_widget(add_more_btn)
        btn_layout.add_widget(save_exit_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Выбор действия",
            content=content,
            size_hint=(0.8, 0.5)
        )

        def add_more_breed(btn):
            popup.dismiss()
            self.show_breed_popup(instance, True)

        def save_and_exit(btn):
            popup.dismiss()
            self.table_screen.show_success("Данные по площадке сохранены!")

        add_more_btn.bind(on_press=add_more_breed)
        save_exit_btn.bind(on_press=save_and_exit)

        popup.open()

    def show_custom_breed_popup(self, instance, breed_type):
        """Показать popup для ввода названия другой породы"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите название другой породы",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.custom_breed_input = TextInput(
            hint_text="Название породы",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto'
        )
        content.add_widget(self.custom_breed_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title=f"Ввод {'хвойной' if breed_type == 'coniferous' else 'лиственной'} породы",
            content=content,
            size_hint=(0.8, 0.6)
        )

        def save_custom_breed(btn):
            breed_name = self.custom_breed_input.text.strip()
            if breed_name:
                # Проверяем, не является ли порода запрещенной
                forbidden_breeds = ['семенная', 'культуры', 'подрост']
                if any(forbidden.lower() in breed_name.lower() for forbidden in forbidden_breeds):
                    self.table_screen.show_error("Эта порода не разрешена для использования!")
                    return
                instance.text = breed_name
                self.show_breed_details_popup(instance, breed_type, breed_name)
                popup.dismiss()
            else:
                self.table_screen.show_error("Название породы не может быть пустым!")

        save_btn.bind(on_press=save_custom_breed)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def update_coniferous_density(self, instance, value):
        """Автоматический расчет густоты для хвойных пород"""
        if 'density' in self.breed_inputs:
            density_input = self.breed_inputs['density']
            try:
                do_05 = int(self.breed_inputs.get('do_05', TextInput(text='0')).text or '0')
                _05_15 = int(self.breed_inputs.get('05_15', TextInput(text='0')).text or '0')
                bolee_15 = int(self.breed_inputs.get('bolee_15', TextInput(text='0')).text or '0')

                total_density = do_05 + _05_15 + bolee_15
                density_input.text = str(total_density) if total_density > 0 else ''
            except (ValueError, AttributeError):
                pass

    def save_data(self, instance):
        # Fill the row in the table
        # First, set the NN (row_index + 1)
        self.table_screen.inputs[self.row_index][0].text = str(self.row_index + 1)

        for i, (field_name, col_index) in enumerate(self.fields):
            value = self.input_fields[i].text.strip()
            if value:
                self.table_screen.inputs[self.row_index][col_index].text = value

        # Save to page_data
        self.table_screen.save_current_page()

        # Show success
        self.table_screen.show_success("Данные площадки молодняков сохранены!")
        self.dismiss()

class ExtendedMolodnikiTableScreen(Screen):
    current_page = NumericProperty(0)
    total_pages = NumericProperty(1)
    unsaved_changes = BooleanProperty(False)
    focused_cell = ListProperty([0, 0])
    edit_mode = BooleanProperty(False)
    current_section = StringProperty("")
    current_quarter = StringProperty("")
    current_plot = StringProperty("")
    current_forestry = StringProperty("")
    current_radius = StringProperty("5.64")
    current_plot_area_ha = StringProperty("")
    plot_area_input = StringProperty("")
    MAX_PAGES = 200

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        print("DEBUG: ExtendedMolodnikiTableScreen __init__ started")
        self.theme_manager = App.get_running_app().theme_manager
        self.reports_dir = "reports"
        os.makedirs(self.reports_dir, exist_ok=True)
        self.db_name = 'forest_data.db'
        self.rows_per_page = 30
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

        # Создаем таблицу для хранения данных молодняков
        cursor.execute('''CREATE TABLE IF NOT EXISTS molodniki_data (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        page_number INTEGER,
                        row_index INTEGER,
                        nn INTEGER,
                        gps_point TEXT,
                        predmet_uhoda TEXT,
                        radius REAL DEFAULT 5.64,
                        primechanie TEXT,
                        section_name TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

        # Создаем индексы для быстрого поиска
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_molodniki_data_page ON molodniki_data (page_number, row_index)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_molodniki_data_section ON molodniki_data (section_name)')

        # Создаем таблицу для хранения пород (множественные породы на одну запись)
        cursor.execute('''CREATE TABLE IF NOT EXISTS molodniki_breeds (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        molodniki_data_id INTEGER,
                        breed_name TEXT,
                        breed_type TEXT, -- 'coniferous' или 'deciduous'
                        do_05 INTEGER DEFAULT 0,
                        _05_15 INTEGER DEFAULT 0,
                        bolee_15 INTEGER DEFAULT 0,
                        density INTEGER DEFAULT 0,
                        height REAL DEFAULT 0.0,
                        diameter REAL DEFAULT 0.0,
                        age INTEGER DEFAULT 0,
                        composition_coefficient REAL DEFAULT 0.0,
                        FOREIGN KEY(molodniki_data_id) REFERENCES molodniki_data(id) ON DELETE CASCADE)''')

        # Создаем индекс для поиска данных пород
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_molodniki_breeds ON molodniki_breeds (molodniki_data_id)')

        # Добавляем недостающие столбцы, если они отсутствуют
        try:
            cursor.execute('ALTER TABLE molodniki_breeds ADD COLUMN diameter REAL DEFAULT 0.0')
        except sqlite3.OperationalError:
            pass  # Столбец уже существует

        try:
            cursor.execute('ALTER TABLE molodniki_breeds ADD COLUMN composition_coefficient REAL DEFAULT 0.0')
        except sqlite3.OperationalError:
            pass  # Столбец уже существует

        # Создаем таблицу для хранения итогов по страницам
        cursor.execute('''CREATE TABLE IF NOT EXISTS molodniki_totals (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        page_number INTEGER,
                        section_name TEXT,
                        total_composition TEXT,
                        total_area REAL DEFAULT 0.0,
                        avg_age REAL DEFAULT 0.0,
                        avg_density REAL DEFAULT 0.0,
                        avg_height REAL DEFAULT 0.0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_molodniki_totals_page ON molodniki_totals (page_number, section_name)')

        # Создаем таблицу для хранения данных пород (JSON)
        cursor.execute('''CREATE TABLE IF NOT EXISTS molodniki_suggestions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        column_index INTEGER,
                        value TEXT,
                        UNIQUE(column_index, value))''')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_molodniki_suggestions ON molodniki_suggestions (column_index, value)')

        conn.commit()
        conn.close()

    def create_ui(self):
        main_layout = BoxLayout(orientation='horizontal', padding=10, spacing=10)

        with self.canvas.before:
            self.bg_color = Color(1, 1, 1, 1)
            self.bg_rect = Rectangle(pos=self.pos, size=self.size)
            self.bind(pos=self._update_bg, size=self._update_bg)

        self._update_background(self.theme_manager.current_theme)

        # Табличная часть (левая панель) - уменьшаем для видимости кнопок
        table_panel = BoxLayout(orientation='vertical', size_hint_x=0.75)

        # Заголовок участка и адресная строка
        header_layout = BoxLayout(orientation='vertical', size_hint=(1, None), height=80, spacing=5)

        # Адресная строка с кнопками
        address_layout = BoxLayout(orientation='horizontal', size_hint=(1, None), height=35, spacing=5)

        # Кнопки Квартал, Выдел, Лесничество, Радиус, Площадь участка
        quarter_btn = ModernButton(
            text='Квартал',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(None, None),
            size=(100, 35),
            font_size='14sp'
        )
        quarter_btn.bind(on_press=self.show_quarter_popup)

        plot_btn = ModernButton(
            text='Выдел',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(None, None),
            size=(80, 35),
            font_size='14sp'
        )
        plot_btn.bind(on_press=self.show_plot_popup)

        forestry_btn = ModernButton(
            text='Лесничество',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(None, None),
            size=(180, 35),
            font_size='14sp'
        )
        forestry_btn.bind(on_press=self.show_forestry_popup)

        radius_btn = ModernButton(
            text='Радиус',
            bg_color=get_color_from_hex('#FF8C00'),
            size_hint=(None, None),
            size=(90, 35),
            font_size='14sp'
        )
        radius_btn.bind(on_press=self.show_radius_popup)

        plot_area_combined_btn = ModernButton(
            text='Площадь участка',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(None, None),
            size=(180, 35),
            font_size='14sp'
        )
        plot_area_combined_btn.bind(on_press=self.show_plot_area_input_popup)

        address_layout.add_widget(quarter_btn)
        address_layout.add_widget(plot_btn)
        address_layout.add_widget(forestry_btn)
        address_layout.add_widget(radius_btn)
        address_layout.add_widget(plot_area_combined_btn)

        # Адресная строка (текстовое поле для отображения адреса)
        self.address_label = Label(
            text=f"Адрес: {self.current_quarter} кв. {self.current_plot} выд. {self.current_forestry}",
            font_name='Roboto',
            size_hint=(1, None),
            height=35,
            color=self._get_text_color(),
            halign='left',
            valign='middle'
        )
        self.address_label.bind(size=self.address_label.setter('text_size'))
        address_layout.add_widget(self.address_label)

        header_layout.add_widget(address_layout)

        # Заголовок участка
        self.section_label = Label(
            text=f"Молодняки - Участок: {self.current_section}",
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

        # Основная таблица (6 столбцов: основные данные)
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True, bar_width=10)
        self.table = GridLayout(cols=6, size_hint=(None, None), spacing=1)
        self.table.bind(minimum_height=self.table.setter('height'),
                       minimum_width=self.table.setter('width'))

        # Заголовки столбцов (6 столбцов: основные данные)
        headers = [
            '№ППР', 'GPS точка', 'Предмет ухода', 'Порода', 'Примечания', 'Тип Леса'
        ]

        self.header_bgs = []
        for i, header in enumerate(headers):
            lbl = Label(
                text=header,
                size_hint_y=None,
                height=30,
                font_name='Roboto',
                bold=True,
                halign='center',
                size_hint_x=None,
                width=120,  # Все столбцы одинаковой ширины
                color=(0, 0, 0, 1)  # Черный цвет для заголовков
            )
            with lbl.canvas.before:
                Color(rgba=get_color_from_hex('#00FF00'))
                bg = Rectangle(pos=lbl.pos, size=lbl.size)
                self.header_bgs.append(bg)
            lbl.bind(pos=lambda i,v, b=bg: setattr(b, 'pos', i.pos), size=lambda i,v, b=bg: setattr(b, 'size', i.size))
            self.table.add_widget(lbl)

        # Создаем строки таблицы (6 столбцов: основные данные)
        self.inputs = []
        for row_idx in range(self.rows_per_page):
            row = []
            for col_idx in range(6):
                inp = AutoCompleteTextInput(multiline=False, size_hint_y=None, height=30)
                inp.row_index = row_idx
                inp.col_index = col_idx
                inp.bind(focus=self.update_focus)
                inp.bind(text=self.update_row_total)
                inp.font_name='Roboto'
                inp.size_hint_x = None
                inp.width = 120  # Все столбцы одинаковой ширины

                # Настройка фильтров ввода для числовых полей
                if col_idx == 0:  # №ППР
                    inp.input_filter = 'int'
                    inp.bind(focus=self.show_tree_popup)
                elif col_idx == 3:  # Порода - открываем popup выбора типа
                    inp.bind(focus=self.show_breed_popup)

                row.append(inp)
                self.table.add_widget(inp)

            self.inputs.append(row)



        # Добавляем кнопки "Итого" и "Проект ухода" по середине после строки итогов
        # Пустая строка для разделения
        spacer = BoxLayout(orientation='horizontal', size_hint_y=None, height=10)
        self.table.add_widget(spacer)

        # Кнопки "Итого", "Проект ухода" и "Дополнительные функции" по середине
        button_container = BoxLayout(orientation='horizontal', size_hint_y=None, height=50, size_hint_x=1, spacing=10)
        button_spacer = BoxLayout(size_hint_x=0.15)  # Спейсер слева
        button_container.add_widget(button_spacer)

        self.total_summary_button = ModernButton(
            text='Итого',
            bg_color=get_color_from_hex('#00FF00'),  # Зеленый цвет
            size_hint=(None, None),
            size=(200, 50),
            font_size='18sp',
            bold=True
        )
        self.total_summary_button.bind(on_press=self.show_total_summary_popup)
        button_container.add_widget(self.total_summary_button)

        # Кнопка "Проект ухода"
        self.care_project_button = ModernButton(
            text='Проект ухода',
            bg_color=get_color_from_hex('#FF8C00'),  # Оранжевый цвет
            size_hint=(None, None),
            size=(200, 50),
            font_size='16sp',
            bold=True
        )
        self.care_project_button.bind(on_press=self.generate_care_project)
        button_container.add_widget(self.care_project_button)

        # Кнопка "Дополнительные функции"
        self.additional_functions_button = ModernButton(
            text='Дополнительные функции',
            bg_color=get_color_from_hex('#9370DB'),  # Фиолетовый цвет
            size_hint=(None, None),
            size=(300, 50),
            font_size='16sp',
            bold=True
        )
        self.additional_functions_button.bind(on_press=self.show_additional_functions_popup)
        button_container.add_widget(self.additional_functions_button)

        button_spacer2 = BoxLayout(size_hint_x=0.15)  # Спейсер справа
        button_container.add_widget(button_spacer2)

        self.table.add_widget(button_container)

        scroll.add_widget(self.table)
        table_panel.add_widget(scroll)
        main_layout.add_widget(table_panel)

        # Правая панель управления (увеличиваем для видимости кнопок)
        control_panel = BoxLayout(
            orientation='vertical',
            size_hint_x=0.25,
            spacing=10,
            padding=[0, 10, 0, 0]
        )

        controls = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=450,
            spacing=10,
            pos_hint={'top': 1}
        )

        button_handlers = {
            'Сохранить': self.save_all_formats,
            'Сохранить страницу': self.save_current_page,
            'Загрузить': self.load_section,
            'Редактировать': self.show_edit_plots_popup,
            'Открыть папку': self.open_excel_file,
            'Очистить': self.clear_table_data,
            'В меню': self.go_back
        }

        button_colors = {
            'Сохранить': '#FFD700',
            'Сохранить страницу': '#00FFFF',
            'Загрузить': '#006400',
            'Редактировать': '#FF6347',
            'Открыть папку': '#0000FF',
            'Очистить': '#800000',
            'В меню': '#FF0000'
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
        self.section_label.text = f"Молодняки - Участок: {self.current_section}"

    def toggle_edit_mode(self, instance):
        self.edit_mode = not self.edit_mode
        instance.bg_color = get_color_from_hex('#FFA500' if self.edit_mode else '#00BFFF')

    def update_focus(self, instance, value):
        if value:
            self.focused_cell = [instance.row_index, instance.col_index]

    def move_focus(self, direction):
        current = self.focused_cell
        if not current: return
        row, col = current

        if direction == 'up': row = max(0, row-1)
        elif direction == 'down': row = min(len(self.inputs)-1, row+1)
        elif direction == 'left': col = max(0, col-1)
        elif direction == 'right': col = min(5, col+1)  # 6 столбцов (0-5)

        self.focused_cell = [row, col]
        inp = self.inputs[row][col]
        inp.focus = True
        inp.cursor = (len(inp.text), 0)

    def show_tree_popup(self, instance, value):
        """Показать popup для ввода данных площадки молодняков"""
        if value and instance.text.strip():
            if not self.edit_mode:
                # In normal mode, only show popup if other columns are empty
                if not any(inp.text.strip() for inp in self.inputs[instance.row_index][1:]):
                    MolodnikiTreeDataInputPopup(self, instance.row_index).open()
            else:
                # In edit mode, always show popup for editing existing data
                MolodnikiTreeDataInputPopup(self, instance.row_index).open()

    def auto_fill_nn(self, instance, value):
        if self.edit_mode: return
        if value and instance.focus:
            try:
                current_number = int(instance.text)
            except ValueError:
                current_number = 0
            for row_idx, row in enumerate(self.inputs):
                if row_idx > instance.row_index:
                    try:
                        prev_num = int(self.inputs[row_idx-1][0].text)
                        row[0].text = str(prev_num + 1)
                    except (ValueError, IndexError):
                        pass

    def show_breed_popup(self, instance, value):
        """Показать popup для выбора типа породы"""
        if not value: return

        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Кнопки выбора типа породы
        type_layout = BoxLayout(orientation='horizontal', spacing=10)
        coniferous_btn = ModernButton(
            text='Хвойные',
            bg_color=get_color_from_hex('#228B22'),
            size_hint=(0.5, None),
            height=50
        )
        deciduous_btn = ModernButton(
            text='Лиственные',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(0.5, None),
            height=50
        )
        type_layout.add_widget(coniferous_btn)
        type_layout.add_widget(deciduous_btn)
        content.add_widget(type_layout)

        # Кнопка отмены
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        content.add_widget(cancel_btn)

        popup = Popup(
            title="Выберите тип породы",
            content=content,
            size_hint=(0.8, 0.5)
        )

        def select_coniferous(btn):
            self.show_breed_selection_popup(instance, 'coniferous')
            popup.dismiss()

        def select_deciduous(btn):
            self.show_breed_selection_popup(instance, 'deciduous')
            popup.dismiss()

        coniferous_btn.bind(on_press=select_coniferous)
        deciduous_btn.bind(on_press=select_deciduous)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_breed_selection_popup(self, instance, breed_type):
        """Показать popup для выбора конкретной породы из словаря"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Заголовок
        title_label = Label(
            text=f"Выберите {'хвойную' if breed_type == 'coniferous' else 'лиственную'} породу",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Списки пород
        if breed_type == 'coniferous':
            breeds = [
                'Сосна', 'Ель', 'Пихта', 'Кедр', 'Лиственница'
            ]
        else:
            breeds = [
                'Берёза', 'Осина', 'Ольха чёрная', 'Ольха серая',
                'Ива', 'Ива кустарниковая'
            ]

        # ScrollView для списка пород
        scroll = ScrollView(size_hint=(1, None), height=300)
        breeds_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        breeds_layout.bind(minimum_height=breeds_layout.setter('height'))

        for breed in breeds:
            btn = ModernButton(
                text=breed,
                bg_color=get_color_from_hex('#87CEEB'),
                size_hint=(1, None),
                height=50,
                font_size='14sp'
            )
            btn.bind(on_press=lambda x, b=breed: self.select_breed(instance, breed_type, b))
            breeds_layout.add_widget(btn)

        scroll.add_widget(breeds_layout)
        content.add_widget(scroll)

        # Кнопка "Другая порода"
        other_btn = ModernButton(
            text='Другая порода',
            bg_color=get_color_from_hex('#DDA0DD'),
            size_hint=(1, None),
            height=50
        )
        other_btn.bind(on_press=lambda x: self.select_breed(instance, breed_type, 'other'))
        content.add_widget(other_btn)

        # Кнопка отмены
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(1, None),
            height=50
        )
        content.add_widget(cancel_btn)

        popup = Popup(
            title=f"Выбор {'хвойной' if breed_type == 'coniferous' else 'лиственной'} породы",
            content=content,
            size_hint=(0.85, 0.85)
        )

        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def select_breed(self, instance, breed_type, selected_breed):
        """Обработка выбора породы"""
        if selected_breed == 'other':
            # Показываем popup для ввода названия другой породы
            self.show_custom_breed_popup(instance, breed_type)
        else:
            # Показываем popup с параметрами породы, передавая название выбранной породы
            self.show_breed_details_popup(instance, breed_type, selected_breed)

    def show_breed_details_popup(self, instance, breed_type, selected_breed=None):
        """Показать popup для управления множественными породами"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Заголовок
        title_label = Label(
            text=f"Управление породами - {selected_breed}",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Список существующих пород в этой строке
        existing_breeds = self.parse_breeds_data(instance.text)
        if not existing_breeds:
            # Пытаемся получить данные из сохраненных данных страницы
            row_idx = instance.row_index
            if self.current_page in self.page_data and row_idx < len(self.page_data[self.current_page]):
                saved_text = self.page_data[self.current_page][row_idx][3]  # Столбец "Порода"
                existing_breeds = self.parse_breeds_data(saved_text)

        # Отображаем список существующих пород
        if existing_breeds:
            breeds_list_label = Label(
                text="Уже добавленные породы:",
                font_name='Roboto',
                bold=True,
                size_hint=(1, None),
                height=30,
                color=(0.3, 0.3, 0.3, 1)
            )
            content.add_widget(breeds_list_label)

            # ScrollView для списка пород
            breeds_scroll = ScrollView(size_hint=(1, None), height=80)
            breeds_list_layout = BoxLayout(orientation='vertical', spacing=5, size_hint_y=None)
            breeds_list_layout.bind(minimum_height=breeds_list_layout.setter('height'))

            for i, breed_info in enumerate(existing_breeds):
                breed_name = breed_info.get('name', 'Неизвестная')
                breed_type = breed_info.get('type', 'unknown')
                params = []
                if 'density' in breed_info and breed_info['density']:
                    params.append(f"Густота: {breed_info['density']}")
                if 'height' in breed_info and breed_info['height']:
                    params.append(f"Высота: {breed_info['height']}м")
                if 'age' in breed_info and breed_info['age']:
                    params.append(f"Возраст: {breed_info['age']} лет")
                if breed_type == 'coniferous':
                    if 'do_05' in breed_info and breed_info['do_05']:
                        params.append(f"До 0.5м: {breed_info['do_05']}")
                    if '05_15' in breed_info and breed_info['05_15']:
                        params.append(f"0.5-1.5м: {breed_info['05_15']}")
                    if 'bolee_15' in breed_info and breed_info['bolee_15']:
                        params.append(f">1.5м: {breed_info['bolee_15']}")

                breed_text = f"{i+1}. {breed_name}: {'; '.join(params)}" if params else f"{i+1}. {breed_name}"
                breed_label = Label(
                    text=breed_text,
                    font_name='Roboto',
                    size_hint=(1, None),
                    height=25,
                    color=(0.2, 0.2, 0.2, 1),
                    text_size=(None, None),
                    halign='left',
                    valign='top'
                )
                breed_label.bind(size=lambda *args: setattr(breed_label, 'text_size', (breed_label.width, None)))
                breeds_list_layout.add_widget(breed_label)

            breeds_scroll.add_widget(breeds_list_layout)
            content.add_widget(breeds_scroll)
        else:
            breeds_list_label = Label(
                text="Породы ещё не добавлены",
                font_name='Roboto',
                size_hint=(1, None),
                height=30,
                color=(0.5, 0.5, 0.5, 1)
            )
            content.add_widget(breeds_list_label)

        # Поля ввода для параметров породы
        fields_layout = GridLayout(cols=2, spacing=5, size_hint=(1, None), height=200)

        if breed_type == 'coniferous':
            fields = [
                ('До 0.5м:', 'do_05'),
                ('0.5-1.5м:', '05_15'),
                ('>1.5м:', 'bolee_15'),
                ('Высота (м):', 'height'),
                ('Густота:', 'density'),
                ('Возраст (лет):', 'age')
            ]
        else:
            fields = [
                ('Густота:', 'density'),
                ('Высота (м):', 'height'),
                ('Возраст (лет):', 'age')
            ]

        self.breed_inputs = {}
        for label_text, field_key in fields:
            lbl = Label(text=label_text, font_name='Roboto', size_hint=(None, None), size=(100, 30))
            inp = TextInput(multiline=False, size_hint=(None, None), size=(100, 30))
            if field_key in ['density', 'age']:
                inp.input_filter = 'int'
            elif field_key == 'height':
                inp.input_filter = 'float'
            elif field_key in ['do_05', '05_15', 'bolee_15']:
                inp.input_filter = 'int'
                if breed_type == 'coniferous':
                    inp.bind(text=self.update_coniferous_density)
            fields_layout.add_widget(lbl)
            fields_layout.add_widget(inp)
            self.breed_inputs[field_key] = inp

        content.add_widget(fields_layout)

        # Кнопки управления
        btn_layout = BoxLayout(orientation='horizontal', spacing=5, size_hint=(1, None), height=50)
        add_btn = ModernButton(
            text='Добавить породу',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.25, 1),
            height=50
        )
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(0.25, 1),
            height=50
        )
        view_btn = ModernButton(
            text='Просмотр',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(0.25, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.25, 1),
            height=50
        )
        btn_layout.add_widget(add_btn)
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(view_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title=f"Параметры породы: {selected_breed}",
            content=content,
            size_hint=(0.9, 0.95)
        )

        def add_breed(btn):
            breed_data = {
                'name': selected_breed,
                'type': breed_type
            }

            for key, inp in self.breed_inputs.items():
                if inp.text.strip():
                    try:
                        if key in ['density', 'age']:
                            breed_data[key] = int(inp.text)
                        elif key == 'height':
                            breed_data[key] = float(inp.text)
                        else:
                            breed_data[key] = float(inp.text)
                    except ValueError:
                        breed_data[key] = inp.text

            existing_breeds = self.parse_breeds_data(instance.text)
            existing_breeds.append(breed_data)
            instance.text = json.dumps(existing_breeds, ensure_ascii=False, indent=2)

            self.update_plot_total(instance, instance.text)

            for inp in self.breed_inputs.values():
                inp.text = ''

            # После добавления первой породы присваиваем номер 1 и предлагаем выбор
            if len(existing_breeds) == 1:
                self.show_breed_choice_popup(instance, selected_breed)
            else:
                self.show_breed_popup(instance, True)
                self.show_success(f"Порода '{selected_breed}' добавлена! Выберите тип следующей породы.")

        def save_breeds(btn):
            existing_breeds = self.parse_breeds_data(instance.text)
            if not existing_breeds:
                existing_breeds = []
            instance.text = json.dumps(existing_breeds, ensure_ascii=False, indent=2)
            self.update_plot_total(instance, instance.text)
            self.show_success("Все данные по площадке сохранены!")
            popup.dismiss()

        def view_breeds(btn):
            popup.dismiss()
            self.show_breeds_list_popup(instance)

        add_btn.bind(on_press=add_breed)
        save_btn.bind(on_press=save_breeds)
        view_btn.bind(on_press=view_breeds)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_breeds_list_popup(self, instance):
        """Показать popup со списком всех пород в этой строке"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Список пород в этой строке",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        breeds_data = self.parse_breeds_data(instance.text)

        if not breeds_data:
            no_breeds_label = Label(
                text="Породы не найдены",
                font_name='Roboto',
                size_hint=(1, None),
                height=50,
                color=(0.5, 0.5, 0.5, 1)
            )
            content.add_widget(no_breeds_label)
        else:
            scroll = ScrollView(size_hint=(1, None), height=300)
            breeds_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
            breeds_layout.bind(minimum_height=breeds_layout.setter('height'))

            for i, breed_info in enumerate(breeds_data):
                breed_card = BoxLayout(
                    orientation='vertical',
                    size_hint=(1, None),
                    height=120,
                    padding=5
                )
                breed_card.canvas.before.clear()
                with breed_card.canvas.before:
                    Color(rgba=get_color_from_hex('#E8F4FD'))
                    Rectangle(pos=breed_card.pos, size=breed_card.size)
                    Color(rgba=get_color_from_hex('#B0BEC5'))
                    Line(rectangle=(breed_card.x, breed_card.y, breed_card.width, breed_card.height), width=1)

                name_label = Label(
                    text=f"{i+1}. {breed_info.get('name', 'Неизвестная порода')}",
                    font_name='Roboto',
                    bold=True,
                    size_hint=(1, None),
                    height=25,
                    color=(0, 0, 0, 1)
                )
                breed_card.add_widget(name_label)

                params_text = []
                breed_type = breed_info.get('type', 'deciduous')

                # В зависимости от типа породы показываем разные параметры
                if breed_type == 'coniferous':
                    # Для хвойных показываем градации + общую густоту и другие параметры
                    if 'do_05' in breed_info and breed_info['do_05']:
                        params_text.append(f"До 0.5м: {breed_info['do_05']}")
                    if '05_15' in breed_info and breed_info['05_15']:
                        params_text.append(f"0.5-1.5м: {breed_info['05_15']}")
                    if 'bolee_15' in breed_info and breed_info['bolee_15']:
                        params_text.append(f">1.5м: {breed_info['bolee_15']}")
                    if 'density' in breed_info and breed_info['density']:
                        params_text.append(f"Общая густота: {breed_info['density']}")
                    if 'height' in breed_info and breed_info['height']:
                        params_text.append(f"Высота: {breed_info['height']}м")
                    if 'age' in breed_info and breed_info['age']:
                        params_text.append(f"Возраст: {breed_info['age']} лет")
                else:
                    # Для лиственных показываем только основную информацию (без градаций)
                    if 'density' in breed_info and breed_info['density']:
                        params_text.append(f"Густота: {breed_info['density']}")
                    if 'height' in breed_info and breed_info['height']:
                        params_text.append(f"Высота: {breed_info['height']}м")
                    if 'age' in breed_info and breed_info['age']:
                        params_text.append(f"Возраст: {breed_info['age']} лет")

                params_label = Label(
                    text="; ".join(params_text) if params_text else "Нет параметров",
                    font_name='Roboto',
                    size_hint=(1, None),
                    height=40,
                    color=(0.3, 0.3, 0.3, 1),
                    text_size=(None, None),
                    halign='left',
                    valign='top'
                )
                params_label.bind(size=lambda *args: setattr(params_label, 'text_size', (params_label.width, None)))
                breed_card.add_widget(params_label)

                btn_layout = BoxLayout(orientation='horizontal', spacing=5, size_hint=(1, None), height=30)
                edit_btn = ModernButton(
                    text='Изменить',
                    bg_color=get_color_from_hex('#87CEEB'),
                    size_hint=(0.5, 1),
                    font_size='12sp'
                )
                delete_btn = ModernButton(
                    text='Удалить',
                    bg_color=get_color_from_hex('#FF6347'),
                    size_hint=(0.5, 1),
                    font_size='12sp'
                )
                btn_layout.add_widget(edit_btn)
                delete_btn.bind(on_press=lambda x, idx=i: self.delete_breed_from_list(instance, idx))
                btn_layout.add_widget(delete_btn)
                breed_card.add_widget(btn_layout)

                def edit_breed(btn, idx=i):
                    self.edit_breed_in_list(instance, idx)

                edit_btn.bind(on_press=edit_breed)

                breeds_layout.add_widget(breed_card)

            scroll.add_widget(breeds_layout)
            content.add_widget(scroll)

        close_btn = ModernButton(
            text='Закрыть',
            bg_color=get_color_from_hex('#808080'),
            size_hint=(1, None),
            height=50
        )
        content.add_widget(close_btn)

        popup = Popup(
            title="Управление породами",
            content=content,
            size_hint=(0.85, 0.9)
        )

        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def edit_breed_in_list(self, instance, breed_index):
        """Редактировать породу в списке"""
        breeds_data = self.parse_breeds_data(instance.text)
        if 0 <= breed_index < len(breeds_data):
            breed_info = breeds_data[breed_index]
            self.show_edit_breed_popup(instance, breed_index, breed_info)

    def delete_breed_from_list(self, instance, breed_index):
        """Удалить породу из списка"""
        breeds_data = self.parse_breeds_data(instance.text)
        if 0 <= breed_index < len(breeds_data):
            breed_name = breeds_data[breed_index].get('name', 'Неизвестная порода')
            breeds_data.pop(breed_index)
            instance.text = json.dumps(breeds_data, ensure_ascii=False, indent=2) if breeds_data else ''
        self.update_totals()
        self.show_success("Порода удалена!")
        if hasattr(self, 'popup') and self.popup:
            self.popup.dismiss()

    def save_totals_to_excel(self, breeds_data, current_radius, plot_area_ha, plot_count, total_plot_area_ha):
        """Сохранить итоговые данные в Excel на новом листе"""
        filename = f"Итоги_молодняков_{self.current_section}_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
        full_path = os.path.join(self.reports_dir, filename)

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Итоги"

            # Заголовок
            ws['A1'] = f'ИТОГИ ПО УЧАСТКУ МОЛОДНЯКОВ - {self.current_section}'
            ws['A1'].font = openpyxl.styles.Font(bold=True, size=14)
            ws.merge_cells('A1:E1')

            # Информация о радиусе
            ws['A3'] = f'Радиус участка: {current_radius:.2f} м'
            ws['A4'] = f'1 дерево = {10000 / (3.14159 * (current_radius ** 2)):.0f} тыс.шт./га'

            # Коэффициент состава
            ws['A6'] = 'КОЭФФИЦИЕНТ СОСТАВА НАСАЖДЕНИЯ'
            ws['A6'].font = openpyxl.styles.Font(bold=True, size=12)

            # Расчет коэффициента состава
            total_densities = {}
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    if data['plots'][0].get('type') == 'coniferous':
                        total_density = 0
                        for p in data['plots']:
                            conif_density = (p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0))
                            total_density += conif_density
                    else:
                        total_density = sum(p.get('density', 0) for p in data['plots'])
                    if total_density > 0:
                        total_densities[breed_name] = total_density

            if total_densities:
                total_all_density = sum(total_densities.values())
                composition_parts = []
                for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                    if total_all_density > 0:
                        coeff = max(1, round(density / total_all_density * 10))
                    else:
                        coeff = 1
                    breed_letter = self.get_breed_letter(breed_name)
                    composition_parts.append(f"{coeff}{breed_letter}")

                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                iterations = 0
                while total_coeffs != 10 and iterations < 100:
                    if total_coeffs > 10:
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] -= 1
                    elif total_coeffs < 10:
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] += 1
                    total_coeffs = sum(coeffs_only)
                    iterations += 1

                sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
                composition_parts = []
                for i, (breed_name, _) in enumerate(sorted_breeds):
                    if i < len(coeffs_only):
                        breed_letter = self.get_breed_letter(breed_name)
                        composition_parts.append(f"{coeffs_only[i]}{breed_letter}")

                composition_text = ''.join(composition_parts) + "Др"
                ws['A7'] = f"Формула состава: {composition_text}"

            # Хвойные породы
            row = 9
            ws[f'A{row}'] = 'ХВОЙНЫЕ ПОРОДЫ - ВЫСОТА ПО ГРАДАЦИЯМ'
            ws[f'A{row}'].font = openpyxl.styles.Font(bold=True, size=12)

            has_coniferous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'coniferous' and data['plots']:
                    has_coniferous = True
                    row += 1
                    zones = data.get('coniferous_zones', {})
                    avg_do_05 = zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0
                    avg_05_15 = zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_bolee_15 = zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_height_total = sum(p['height'] for p in data['plots'] if p['height'] > 0) / len([p for p in data['plots'] if p['height'] > 0]) if any(p['height'] > 0 for p in data['plots']) else 0

                    ws[f'A{row}'] = f"{breed_name}:"
                    ws[f'B{row}'] = f"до 0.5м: {avg_do_05:.1f} шт/га"
                    row += 1
                    ws[f'B{row}'] = f"0.5-1.5м: {avg_05_15:.1f} шт/га"
                    row += 1
                    ws[f'B{row}'] = f">1.5м: {avg_bolee_15:.1f} шт/га"
                    row += 1
                    ws[f'B{row}'] = f"средняя высота породы: {avg_height_total:.1f}м"
                    row += 1

            # Лиственные породы
            if has_coniferous:
                row += 1
            ws[f'A{row}'] = 'ЛИСТВЕННЫЕ ПОРОДЫ - СРЕДНИЕ ПОКАЗАТЕЛИ'
            ws[f'A{row}'].font = openpyxl.styles.Font(bold=True, size=12)

            has_deciduous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'deciduous' and data['plots']:
                    has_deciduous = True
                    row += 1
                    avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])
                    avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                    avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0
                    avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                    avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                    ws[f'A{row}'] = f"{breed_name}:"
                    ws[f'B{row}'] = f"Средняя густота: {avg_density:.1f} шт/га"
                    row += 1
                    ws[f'B{row}'] = f"Средняя высота: {avg_height:.1f}м"
                    row += 1
                    ws[f'B{row}'] = f"Средний возраст: {avg_age:.1f} лет"
                    row += 1

            # Автоподбор ширины столбцов
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(full_path)
            self.show_success(f"Итоги сохранены в Excel: {filename}")
        except Exception as e:
            self.show_error(f"Ошибка сохранения итогов в Excel: {str(e)}")

    def save_totals_to_word(self, breeds_data, current_radius, plot_area_ha, plot_count, total_plot_area_ha):
        """Сохранить итоговые данные в Word"""
        try:
            from docx import Document
            from docx.shared import Inches

            filename = f"Итоги_молодняков_{self.current_section}_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.docx"
            full_path = os.path.join(self.reports_dir, filename)

            doc = Document()
            doc.add_heading(f'ИТОГИ ПО УЧАСТКУ МОЛОДНЯКОВ - {self.current_section}', 0)

            # Информация о радиусе
            doc.add_paragraph(f'Радиус участка: {current_radius:.2f} м')
            doc.add_paragraph(f'1 дерево = {10000 / (3.14159 * (current_radius ** 2)):.0f} тыс.шт./га')

            # Коэффициент состава
            doc.add_heading('КОЭФФИЦИЕНТ СОСТАВА НАСАЖДЕНИЯ', level=2)

            total_densities = {}
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    if data['plots'][0].get('type') == 'coniferous':
                        total_density = 0
                        for p in data['plots']:
                            conif_density = (p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0))
                            total_density += conif_density
                    else:
                        total_density = sum(p.get('density', 0) for p in data['plots'])
                    if total_density > 0:
                        total_densities[breed_name] = total_density

            if total_densities:
                total_all_density = sum(total_densities.values())
                composition_parts = []
                for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                    if total_all_density > 0:
                        coeff = max(1, round(density / total_all_density * 10))
                    else:
                        coeff = 1
                    breed_letter = self.get_breed_letter(breed_name)
                    composition_parts.append(f"{coeff}{breed_letter}")

                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                iterations = 0
                while total_coeffs != 10 and iterations < 100:
                    if total_coeffs > 10:
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] -= 1
                    elif total_coeffs < 10:
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] += 1
                    total_coeffs = sum(coeffs_only)
                    iterations += 1

                sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
                composition_parts = []
                for i, (breed_name, _) in enumerate(sorted_breeds):
                    if i < len(coeffs_only):
                        breed_letter = self.get_breed_letter(breed_name)
                        composition_parts.append(f"{coeffs_only[i]}{breed_letter}")

                composition_text = ''.join(composition_parts) + "Др"
                doc.add_paragraph(f"Формула состава: {composition_text}")

            # Хвойные породы
            doc.add_heading('ХВОЙНЫЕ ПОРОДЫ - ВЫСОТА ПО ГРАДАЦИЯМ', level=2)

            has_coniferous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'coniferous' and data['plots']:
                    has_coniferous = True
                    zones = data.get('coniferous_zones', {})
                    avg_do_05 = zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0
                    avg_05_15 = zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_bolee_15 = zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_height_total = sum(p['height'] for p in data['plots'] if p['height'] > 0) / len([p for p in data['plots'] if p['height'] > 0]) if any(p['height'] > 0 for p in data['plots']) else 0

                    p = doc.add_paragraph()
                    p.add_run(f"{breed_name}:").bold = True
                    doc.add_paragraph(f"• до 0.5м: {avg_do_05:.1f} шт/га")
                    doc.add_paragraph(f"• 0.5-1.5м: {avg_05_15:.1f} шт/га")
                    doc.add_paragraph(f"• >1.5м: {avg_bolee_15:.1f} шт/га")
                    doc.add_paragraph(f"• средняя высота породы: {avg_height_total:.1f}м")

            # Лиственные породы
            doc.add_heading('ЛИСТВЕННЫЕ ПОРОДЫ - СРЕДНИЕ ПОКАЗАТЕЛИ', level=2)

            has_deciduous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'deciduous' and data['plots']:
                    has_deciduous = True
                    avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])
                    avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                    avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0
                    avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                    avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                    p = doc.add_paragraph()
                    p.add_run(f"{breed_name}:").bold = True
                    doc.add_paragraph(f"• Средняя густота: {avg_density:.1f} шт/га")
                    doc.add_paragraph(f"• Средняя высота: {avg_height:.1f}м")
                    doc.add_paragraph(f"• Средний возраст: {avg_age:.1f} лет")

            doc.save(full_path)
            self.show_success(f"Итоги сохранены в Word: {filename}")
        except ImportError:
            self.show_error("Для сохранения в Word установите библиотеку python-docx: pip install python-docx")
        except Exception as e:
            self.show_error(f"Ошибка сохранения итогов в Word: {str(e)}")

    def show_plot_area_input_popup(self, instance):
        """Показать popup для ввода площади участка в гектарах"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите площадь обследуемого участка",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        self.plot_area_input_field = TextInput(
            hint_text="Площадь участка (га)",
            multiline=False,
            size_hint=(1, None),
            height=50,
            font_name='Roboto',
            input_filter='float',
            text=self.plot_area_input if hasattr(self, 'plot_area_input') and self.plot_area_input else ''
        )
        content.add_widget(self.plot_area_input_field)

        info_label = Label(
            text="Укажите площадь обследуемого участка в гектарах.\n"
                 "Это значение используется для расчета площади перечета\n"
                 "по всем площадкам и отображается в итоговых отчетах.",
            font_name='Roboto',
            font_size='14sp',
            color=(0.3, 0.3, 0.3, 1),
            size_hint=(1, None),
            height=80,
            halign='left',
            valign='top'
        )
        info_label.bind(size=lambda *args: setattr(info_label, 'text_size', (info_label.width, None)))
        content.add_widget(info_label)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Площадь участка",
            content=content,
            size_hint=(0.8, 0.6)
        )

        def save_plot_area(btn):
            try:
                plot_area = float(self.plot_area_input_field.text.strip())
                if plot_area <= 0:
                    self.show_error("Площадь участка должна быть положительным числом!")
                    return

                self.plot_area_input = str(plot_area)
                self.show_success(f"Площадь участка {plot_area} га сохранена")
                popup.dismiss()

            except ValueError:
                self.show_error("Введите корректное числовое значение площади!")

        save_btn.bind(on_press=save_plot_area)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_plot_area_ha_popup(self, instance):
        """Показать popup с информацией о площади участка в гектарах"""
        try:
            current_radius = float(self.current_radius) if self.current_radius else 5.64
            plot_area_m2 = 3.14159 * (current_radius ** 2)
            plot_area_ha = plot_area_m2 / 10000

            # Расчет площади перечета по всем площадкам
            total_plot_area_ha = 0.0
            plot_count = 0

            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 4 and row[3]:  # Есть данные о породах
                        try:
                            breeds_data = json.loads(row[3]) if isinstance(row[3], str) else []
                            if breeds_data:
                                plot_count += 1
                                total_plot_area_ha += plot_area_ha
                        except (json.JSONDecodeError, TypeError):
                            continue

            content = BoxLayout(orientation='vertical', spacing=10, padding=10)

            title_label = Label(
                text="Площадь участка в гектарах",
                font_name='Roboto',
                font_size='18sp',
                bold=True,
                color=(0, 0.5, 0, 1),
                size_hint=(1, None),
                height=40
            )
            content.add_widget(title_label)

            info_text = f"""
Одиночная площадка:
Радиус: {current_radius:.2f} м
Площадь: {plot_area_ha:.4f} га

Всего площадок: {plot_count}
Совокупная площадь перечета: {total_plot_area_ha:.4f} га

Расчет совокупной площади:
{plot_count} площадок × {plot_area_ha:.4f} га = {total_plot_area_ha:.4f} га

Пример расчета густоты на гектар:
Если на площадке 10 деревьев, то густота = 10 / {plot_area_ha:.4f} ≈ {10/plot_area_ha:.1f} шт/га
"""

            info_label = Label(
                text=info_text,
                font_name='Roboto',
                font_size='14sp',
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=250,
                halign='left',
                valign='top'
            )
            info_label.bind(size=lambda *args: setattr(info_label, 'text_size', (info_label.width, None)))
            content.add_widget(info_label)

            close_btn = ModernButton(
                text='Закрыть',
                bg_color=get_color_from_hex('#808080'),
                size_hint=(1, None),
                height=50
            )
            content.add_widget(close_btn)

            popup = Popup(
                title="Площадь участка (га)",
                content=content,
                size_hint=(0.8, 0.8)
            )

            close_btn.bind(on_press=popup.dismiss)
            popup.open()

        except Exception as e:
            self.show_error(f"Ошибка расчета площади: {str(e)}")

    def show_plot_area_combined_popup(self, instance):
        """Показать объединенное popup для работы с площадью участка"""
        content = BoxLayout(orientation='vertical', spacing=20, padding=20)

        try:
            current_radius = float(self.current_radius) if self.current_radius else 5.64
            plot_area_m2 = 3.14159 * (current_radius ** 2)
            plot_area_ha = plot_area_m2 / 10000

            # Расчет площади перечета по всем площадкам
            total_plot_area_ha = 0.0
            plot_count = 0

            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 4 and row[3]:  # Есть данные о породах
                        try:
                            breeds_data = json.loads(row[3]) if isinstance(row[3], str) else []
                            if breeds_data:
                                plot_count += 1
                                total_plot_area_ha += plot_area_ha
                        except (json.JSONDecodeError, TypeError):
                            continue

            title_label = Label(
                text="Площадь участка",
                font_name='Roboto',
                font_size='20sp',
                bold=True,
                color=(0, 0.5, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            content.add_widget(title_label)

            # Раздел ввода площади участка
            input_section = BoxLayout(orientation='vertical', spacing=10, size_hint=(1, None), height=120)

            input_title = Label(
                text="Ввод площади участка",
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                size_hint=(1, None),
                height=30,
                halign='center'
            )
            input_section.add_widget(input_title)

            plot_area_input_field = TextInput(
                hint_text="Площадь участка (га)",
                multiline=False,
                size_hint=(1, None),
                height=50,
                font_name='Roboto',
                input_filter='float',
                text=str(self._get_current_plot_area_input()) if hasattr(self, 'plot_area_input') and self.plot_area_input else ''
            )
            input_section.add_widget(plot_area_input_field)

            content.add_widget(input_section)

            # Раздел информации о площади
            info_label = Label(
                text="Информация о площади участка:",
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                size_hint=(1, None),
                height=30,
                halign='center'
            )
            content.add_widget(info_label)

            info_text = ScrollView(size_hint=(1, None), height=250)
            info_layout = BoxLayout(orientation='vertical', spacing=5, padding=10, size_hint_y=None)
            info_layout.bind(minimum_height=info_layout.setter('height'))

            info_data = Label(
                text=f"""Одиночная площадка:
Радиус: {current_radius:.2f} м
Площадь: {plot_area_ha:.4f} га

Всего площадок: {plot_count}
Совокупная площадь перечета: {total_plot_area_ha:.4f} га

Расчет совокупной площади:
{plot_count} площадок × {plot_area_ha:.4f} га = {total_plot_area_ha:.4f} га

Пример расчета густоты на гектар:
Если на площадке 10 деревьев, то густота = 10 / {plot_area_ha:.4f} ≈ {10/plot_area_ha:.1f} шт/га""",
                font_name='Roboto',
                font_size='14sp',
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=200,
                halign='left',
                valign='top'
            )
            info_data.bind(size=lambda *args: setattr(info_data, 'text_size', (info_data.width, None)))
            info_layout.add_widget(info_data)
            info_text.add_widget(info_layout)

            content.add_widget(info_text)

            # Кнопки управления (объединение сохранения и обновления в одну кнопку)
            btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=60)

            combined_btn = ModernButton(
                text='Сохранить и обновить',
                bg_color=get_color_from_hex('#00FF00'),
                size_hint=(0.7, 1),
                height=60
            )

            close_btn = ModernButton(
                text='Закрыть',
                bg_color=get_color_from_hex('#FF6347'),
                size_hint=(0.3, 1),
                height=60
            )

            btn_layout.add_widget(combined_btn)
            btn_layout.add_widget(close_btn)

            content.add_widget(btn_layout)

            popup = Popup(
                title="Площадь участка",
                content=content,
                size_hint=(0.8, 0.9)
            )

            def save_and_refresh(btn):
                # Сначала сохраняем площадь участка
                try:
                    plot_area = float(plot_area_input_field.text.strip())
                    if plot_area <= 0:
                        self.show_error("Площадь участка должна быть положительным числом!")
                        return

                    self.plot_area_input = str(plot_area)
                    self.show_success(f"Площадь участка {plot_area} га сохранена")
                except ValueError:
                    self.show_error("Введите корректное числовое значение площади!")
                    return

                # Затем обновляем информацию о площади
                popup.dismiss()
                self.show_plot_area_combined_popup(instance)

            combined_btn.bind(on_press=save_and_refresh)
            close_btn.bind(on_press=popup.dismiss)

            popup.open()

        except Exception as e:
            self.show_error(f"Ошибка расчета площади: {str(e)}")

    def _get_current_plot_area_input(self):
        """Получить текущее значение площади участка"""
        # If stored in instance variable
        if hasattr(self, 'plot_area_input') and self.plot_area_input:
            return self.plot_area_input
        return ''

    def update_plot_total(self, instance, value):
        """Обновляем итог по площадке при изменении данных"""
        row_idx = instance.row_index

        breeds_text = self.inputs[row_idx][3].text
        breeds_data = self.parse_breeds_data(breeds_text)

        if not breeds_data:
            return

        total_density = 0
        total_height = 0.0
        total_age = 0
        breed_count = 0
        breed_names = []

        for breed_info in breeds_data:
            breed_count += 1
            breed_name = breed_info.get('name', 'Неизвестная')
            breed_names.append(breed_name)

            if breed_info.get('type') == 'coniferous':
                coniferous_density = (breed_info.get('do_05', 0) +
                                    breed_info.get('05_15', 0) +
                                    breed_info.get('bolee_15', 0))
                if coniferous_density > 0:
                    total_density += coniferous_density
            elif 'density' in breed_info and breed_info['density']:
                total_density += breed_info['density']

            if 'height' in breed_info and breed_info['height']:
                total_height += breed_info['height']
            if 'age' in breed_info and breed_info['age']:
                total_age += breed_info['age']

        # Обновляем общие итоги
        self.update_totals()

    def show_care_queue_popup(self, instance):
        """Показать popup для выбора мероприятий рубки"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Выберите мероприятие:",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        # Поле для ввода мероприятия
        self.activity_input = TextInput(
            hint_text="Введите мероприятие",
            multiline=False,
            size_hint=(1, None),
            height=50,
            font_name='Roboto'
        )
        content.add_widget(self.activity_input)

        # Выбор очереди
        queue_label = Label(
            text="Очередь:",
            font_name='Roboto',
            font_size='16sp',
            bold=True,
            size_hint=(1, None),
            height=30,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(queue_label)

        # Радио-кнопки для выбора очереди
        from kivy.uix.spinner import Spinner
        self.queue_spinner = Spinner(
            text='Выберите очередь',
            values=('первая', 'вторая', 'третья'),
            size_hint=(1, None),
            height=50,
            font_name='Roboto'
        )
        content.add_widget(self.queue_spinner)

        # Чекбоксы для выбора типов мероприятий
        self.activity_checkboxes = {}
        activities = ['осветление', 'прочистка']

        activities_label = Label(
            text="Типы мероприятий:",
            font_name='Roboto',
            font_size='16sp',
            bold=True,
            size_hint=(1, None),
            height=30,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(activities_label)

        for activity in activities:
            checkbox_layout = BoxLayout(orientation='horizontal', size_hint=(1, None), height=40)
            checkbox = CheckBox(size_hint=(None, 1), width=40)
            label = Label(
                text=activity,
                font_name='Roboto',
                font_size='16sp',
                size_hint=(1, 1),
                halign='left',
                valign='middle'
            )
            label.bind(size=lambda *args: setattr(label, 'text_size', (label.width, None)))
            checkbox_layout.add_widget(checkbox)
            checkbox_layout.add_widget(label)
            content.add_widget(checkbox_layout)
            self.activity_checkboxes[activity] = checkbox

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Мероприятие рубки",
            content=content,
            size_hint=(0.8, 0.8)
        )

        def save_activity(btn):
            activity_text = self.activity_input.text.strip()
            selected_queue = self.queue_spinner.text
            selected_activities = [activity for activity, checkbox in self.activity_checkboxes.items() if checkbox.active]

            if not activity_text and not selected_activities:
                self.show_error("Введите мероприятие или выберите тип мероприятия!")
                return

            if selected_queue == 'Выберите очередь':
                self.show_error("Выберите очередь!")
                return

            result_parts = []
            if activity_text:
                result_parts.append(f"Мероприятие: {activity_text}")
            if selected_queue != 'Выберите очередь':
                result_parts.append(f"Очередь: {selected_queue}")
            if selected_activities:
                result_parts.append(f"Типы: {', '.join(selected_activities)}")

            self.show_success(f"Мероприятие сохранено: {'; '.join(result_parts)}")
            popup.dismiss()

        save_btn.bind(on_press=save_activity)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_characteristics_popup(self, instance):
        """Показать popup для характеристики молодняков"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Характеристика молодняков:",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        # Поля для ввода характеристик
        self.characteristics_inputs = {}
        characteristics = ['Лучшие', 'Вспомогательные', 'Нежелательные']

        for char in characteristics:
            char_layout = BoxLayout(orientation='vertical', size_hint=(1, None), height=80, spacing=5)
            char_label = Label(
                text=f"{char}:",
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                size_hint=(1, None),
                height=25
            )
            char_input = TextInput(
                hint_text="Введите название породы",
                multiline=True,
                size_hint=(1, None),
                height=50,
                font_name='Roboto'
            )
            char_layout.add_widget(char_label)
            char_layout.add_widget(char_input)
            content.add_widget(char_layout)
            self.characteristics_inputs[char] = char_input

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Характеристика молодняков",
            content=content,
            size_hint=(0.8, 0.8)
        )

        def save_characteristics(btn):
            filled_characteristics = {}
            for char, input_field in self.characteristics_inputs.items():
                value = input_field.text.strip()
                if value:
                    filled_characteristics[char] = value

            if filled_characteristics:
                characteristics_text = "\n".join([f"{k}: {v}" for k, v in filled_characteristics.items()])
                self.show_success(f"Характеристики сохранены:\n{characteristics_text}")
            else:
                self.show_error("Заполните хотя бы одну характеристику!")
                return
            popup.dismiss()

        save_btn.bind(on_press=save_characteristics)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_date_popup(self, instance):
        """Показать popup для ввода даты рубки"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите дату рубки:",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        # Поле для ввода даты
        self.date_input = TextInput(
            hint_text="ДД.ММ.ГГГГ",
            multiline=False,
            size_hint=(1, None),
            height=50,
            font_name='Roboto',
            input_filter='0123456789.'
        )
        content.add_widget(self.date_input)

        info_label = Label(
            text="Формат: ДД.ММ.ГГГГ\nНапример: 15.06.2025",
            font_name='Roboto',
            font_size='14sp',
            color=(0.3, 0.3, 0.3, 1),
            size_hint=(1, None),
            height=50,
            halign='left',
            valign='top'
        )
        info_label.bind(size=lambda *args: setattr(info_label, 'text_size', (info_label.width, None)))
        content.add_widget(info_label)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Дата рубки",
            content=content,
            size_hint=(0.8, 0.6)
        )

        def save_date(btn):
            date_text = self.date_input.text.strip()
            if date_text:
                # Простая валидация формата даты
                import re
                if re.match(r'^\d{2}\.\d{2}\.\d{4}$', date_text):
                    self.show_success(f"Дата рубки сохранена: {date_text}")
                else:
                    self.show_error("Неверный формат даты! Используйте ДД.ММ.ГГГГ")
                    return
            else:
                self.show_error("Введите дату рубки!")
                return
            popup.dismiss()

        save_btn.bind(on_press=save_date)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_technology_popup(self, instance):
        """Показать popup для ввода технологии ухода"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите технологию ухода:",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        # Поле для ввода технологии
        self.technology_input = TextInput(
            hint_text="Опишите технологию ухода",
            multiline=True,
            size_hint=(1, None),
            height=100,
            font_name='Roboto'
        )
        content.add_widget(self.technology_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Технология ухода",
            content=content,
            size_hint=(0.8, 0.7)
        )

        def save_technology(btn):
            technology_text = self.technology_input.text.strip()
            if technology_text:
                self.show_success(f"Технология ухода сохранена: {technology_text[:50]}...")
            else:
                self.show_error("Введите технологию ухода!")
                return
            popup.dismiss()

        save_btn.bind(on_press=save_technology)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_forest_purpose_popup(self, instance):
        """Показать popup для выбора назначения лесов"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Выберите назначение лесов:",
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        # Чекбоксы выбора назначения лесов
        forest_purposes = [
            ('Эксплуатационные', 'Эксплуатационные леса'),
            ('Защитные', 'Защитные леса'),
            ('Резервные', 'Резервные леса')
        ]

        self.forest_purpose_checkboxes = {}
        self.selected_forest_purpose = None

        for short_name, full_name in forest_purposes:
            checkbox_layout = BoxLayout(orientation='horizontal', size_hint=(1, None), height=50, spacing=10)
            checkbox = CheckBox(size_hint=(None, 1), width=40)
            label = Label(
                text=f"{short_name} ({full_name})",
                font_name='Roboto',
                font_size='16sp',
                size_hint=(1, 1),
                halign='left',
                valign='middle',
                color=(0, 0, 0, 1)
            )
            label.bind(size=lambda *args: setattr(label, 'text_size', (label.width, None)))
            checkbox_layout.add_widget(checkbox)
            checkbox_layout.add_widget(label)
            content.add_widget(checkbox_layout)
            self.forest_purpose_checkboxes[full_name] = checkbox

        # Кнопки управления
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Назначение лесов",
            content=content,
            size_hint=(0.8, 0.7)
        )

        def save_forest_purpose(btn):
            selected_purpose = None
            for purpose, checkbox in self.forest_purpose_checkboxes.items():
                if checkbox.active:
                    selected_purpose = purpose
                    break

            if selected_purpose:
                self.selected_forest_purpose = selected_purpose
                self.show_success(f"Назначение лесов установлено: {selected_purpose}")
                popup.dismiss()
            else:
                self.show_error("Выберите назначение лесов!")

        save_btn.bind(on_press=save_forest_purpose)
        cancel_btn.bind(on_press=popup.dismiss)
        self.forest_purpose_popup = popup
        popup.open()

    def select_forest_purpose(self, purpose):
        """Обработка выбора назначения лесов"""
        self.selected_forest_purpose = purpose
        self.show_success(f"Назначение лесов установлено: {purpose}")
        if hasattr(self, 'forest_purpose_popup'):
            self.forest_purpose_popup.dismiss()

    def show_additional_functions_popup(self, instance):
        """Показать popup с дополнительными функциями"""
        content = BoxLayout(orientation='vertical', spacing=15, padding=15)

        title_label = Label(
            text="Дополнительные функции:",
            font_name='Roboto',
            font_size='20sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=50,
            halign='center'
        )
        content.add_widget(title_label)

        # Кнопки дополнительных функций
        buttons_layout = GridLayout(cols=2, spacing=10, size_hint=(1, None), height=400)

        # Кнопка Очередь рубки
        care_queue_btn = ModernButton(
            text='Очередь рубки',
            bg_color=get_color_from_hex('#FF8C00'),
            size_hint=(1, None),
            height=60,
            font_size='16sp'
        )
        care_queue_btn.bind(on_press=self.show_care_queue_popup)
        buttons_layout.add_widget(care_queue_btn)

        # Кнопка Характеристика молодняков
        characteristics_btn = ModernButton(
            text='Характеристика\nмолодняков',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(1, None),
            height=60,
            font_size='14sp'
        )
        characteristics_btn.bind(on_press=self.show_characteristics_popup)
        buttons_layout.add_widget(characteristics_btn)

        # Кнопка Дата рубки
        date_btn = ModernButton(
            text='Дата рубки',
            bg_color=get_color_from_hex('#87CEEB'),
            size_hint=(1, None),
            height=60,
            font_size='16sp'
        )
        date_btn.bind(on_press=self.show_date_popup)
        buttons_layout.add_widget(date_btn)

        # Кнопка Технология ухода
        technology_btn = ModernButton(
            text='Технология\nухода',
            bg_color=get_color_from_hex('#DDA0DD'),
            size_hint=(1, None),
            height=60,
            font_size='14sp'
        )
        technology_btn.bind(on_press=self.show_technology_popup)
        buttons_layout.add_widget(technology_btn)

        # Кнопка Назначение лесов
        forest_purpose_btn = ModernButton(
            text='Назначение\nлесов',
            bg_color=get_color_from_hex('#8B4513'),
            size_hint=(1, None),
            height=60,
            font_size='14sp'
        )
        forest_purpose_btn.bind(on_press=self.show_forest_purpose_popup)
        buttons_layout.add_widget(forest_purpose_btn)

        content.add_widget(buttons_layout)

        # Кнопка закрытия
        close_btn = ModernButton(
            text='Закрыть',
            bg_color=get_color_from_hex('#808080'),
            size_hint=(1, None),
            height=50
        )
        content.add_widget(close_btn)

        popup = Popup(
            title="Дополнительные функции",
            content=content,
            size_hint=(0.9, 0.9)
        )

        close_btn.bind(on_press=popup.dismiss)
        popup.open()

    def get_breed_letter(self, breed_name):
        """Получение первой буквы для коэффициента состава породы"""
        breed_letters = {
            'Сосна': 'С',
            'Ель': 'Е',
            'Пихта': 'П',
            'Кедр': 'К',
            'Лиственница': 'Л',
            'Берёза': 'Б',
            'Осина': 'Ос',
            'Ольха чёрная': 'ОЧ',
            'Ольха серая': 'ОС',
            'Ива': 'И',
            'Ива кустарниковая': 'ИК'
        }

        for full_name, letter in breed_letters.items():
            if full_name.lower() in breed_name.lower():
                return letter

        # Возвращаем первую букву имени породы, если не найдено
        return breed_name[0].upper() if breed_name else 'Н'

    def get_breed_letter(self, breed_name):
        """Получение первой буквы для коэффициента состава породы"""
        breed_letters = {
            'Сосна': 'С',
            'Ель': 'Е',
            'Пихта': 'П',
            'Кедр': 'К',
            'Лиственница': 'Л',
            'Берёза': 'Б',
            'Осина': 'Ос',
            'Ольха чёрная': 'ОЧ',
            'Ольха серая': 'ОС',
            'Ива': 'И',
            'Ива кустарниковая': 'ИК'
        }

        for full_name, letter in breed_letters.items():
            if full_name.lower() in breed_name.lower():
                return letter

        # Возвращаем первую букву имени породы, если не найдено
        return breed_name[0].upper() if breed_name else 'Н'

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

    def show_quarter_popup(self, instance):
        """Показать popup для ввода квартала"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите номер квартала",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.quarter_input = TextInput(
            hint_text="Номер квартала",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            input_filter='int',
            text=self.current_quarter
        )
        content.add_widget(self.quarter_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=40)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка квартала",
            content=content,
            size_hint=(0.6, 0.5)
        )

        def save_quarter(btn):
            quarter = self.quarter_input.text.strip()
            if quarter:
                self.current_quarter = quarter
                self.update_address_label()
                self.show_success(f"Квартал установлен: {quarter}")
                popup.dismiss()
            else:
                self.show_error("Номер квартала не может быть пустым!")

        save_btn.bind(on_press=save_quarter)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_plot_popup(self, instance):
        """Показать popup для ввода выдела"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите номер выдела",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.plot_input = TextInput(
            hint_text="Номер выдела",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            input_filter='int',
            text=self.current_plot
        )
        content.add_widget(self.plot_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка выдела",
            content=content,
            size_hint=(0.6, 0.5)
        )

        def save_plot(btn):
            plot = self.plot_input.text.strip()
            if plot:
                self.current_plot = plot
                self.update_address_label()
                self.show_success(f"Выдел установлен: {plot}")
                popup.dismiss()
            else:
                self.show_error("Номер выдела не может быть пустым!")

        save_btn.bind(on_press=save_plot)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_forestry_popup(self, instance):
        """Показать popup для ввода лесничества"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите название лесничества",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Поле для лесничества
        forestry_label = Label(
            text="Лесничество:",
            font_name='Roboto',
            size_hint=(1, None),
            height=25,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(forestry_label)

        self.forestry_input = TextInput(
            hint_text="Название лесничества",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            text=self.current_forestry
        )
        content.add_widget(self.forestry_input)

        # Поле для участкового лесничества
        district_forestry_label = Label(
            text="Участковое лесничество:",
            font_name='Roboto',
            size_hint=(1, None),
            height=25,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(district_forestry_label)

        self.district_forestry_input = TextInput(
            hint_text="Название участкового лесничества",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            text=getattr(self, 'current_district_forestry', '')
        )
        content.add_widget(self.district_forestry_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка лесничества",
            content=content,
            size_hint=(0.6, 0.7)
        )

        def save_forestry(btn):
            forestry = self.forestry_input.text.strip()
            district_forestry = self.district_forestry_input.text.strip()
            if forestry:
                self.current_forestry = forestry
                self.current_district_forestry = district_forestry
                self.update_address_label()
                self.show_success(f"Лесничество установлено: {forestry}" + (f", участковое: {district_forestry}" if district_forestry else ""))
                popup.dismiss()
            else:
                self.show_error("Название лесничества не может быть пустым!")

        save_btn.bind(on_press=save_forestry)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def calculate_section_totals(self):
        """Расчет итогов по всему разделу (все страницы)"""
        breed_composition = {}
        total_stats = {'density': [], 'height': [], 'age': []}
        coniferous_stats = {'do_05': [], '05_15': [], 'bolee_15': [], 'height': [], 'age': []}

        for page_num, page_data in self.page_data.items():
            for row_data in page_data:
                # row_data имеет 6 элементов: nn, gps_point, predmet_uhoda, poroda, primechanie, tip_lesa
                if row_data[2]:  # predmet_uhoda
                    composition = self.parse_composition(row_data[2])
                    for breed, count in composition.items():
                        if breed not in breed_composition:
                            breed_composition[breed] = []
                        breed_composition[breed].append(count)

                breeds_text = row_data[3]  # poroda
                if breeds_text:
                    breeds_data = self.parse_breeds_data(breeds_text)
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'coniferous':
                            # Густота хвойных = сумма градаций
                            coniferous_density = (breed_info.get('do_05', 0) +
                                                breed_info.get('05_15', 0) +
                                                breed_info.get('bolee_15', 0))
                            if coniferous_density > 0:
                                total_stats['density'].append(coniferous_density)
                            else:
                                if 'density' in breed_info and breed_info['density']:
                                    total_stats['density'].append(breed_info['density'])

                        elif 'density' in breed_info and breed_info['density']:
                            total_stats['density'].append(breed_info['density'])

                        if 'height' in breed_info and breed_info['height']:
                            total_stats['height'].append(breed_info['height'])
                        if 'age' in breed_info and breed_info['age']:
                            total_stats['age'].append(breed_info['age'])

        # Рассчитываем остальные итоги

        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)  # Площадь пробной площади в м²

        # Расчет средних по градациям для хвойных по формулам лесного хозяйства на гектар
        coniferous_stats_ha = []
        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                if breeds_data and any(b.get('type') == 'coniferous' for b in breeds_data):
                    coniferous_density_ha = 0
                    height_sum = 0
                    age_sum = 0
                    count = 0
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            coniferous_density_ha += (do_05 * 10000 / plot_area_m2) + (_05_15 * 10000 / plot_area_m2) + (bolee_15 * 10000 / plot_area_m2)
                            if breed_info.get('height'):
                                height_sum += breed_info['height']
                                count += 1
                            if breed_info.get('age'):
                                age_sum += breed_info['age']

                    coniferous_stats_ha.append({
                        'density_ha': coniferous_density_ha if coniferous_density_ha > 0 else 0,
                        'height': height_sum / count if count > 0 else 0,
                        'age': age_sum / count if count > 0 else 0
                    })

        # Итоги по лиственным
        deciduous_stats = []
        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                if breeds_data:
                    deciduous_density_total = 0
                    deciduous_height = []
                    deciduous_age = []
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'deciduous':
                            deciduous_density_total += breed_info.get('density', 0)
                            if breed_info.get('height', 0) > 0:
                                deciduous_height.append(breed_info['height'])
                            if breed_info.get('age', 0) > 0:
                                deciduous_age.append(breed_info['age'])
                    if deciduous_height or deciduous_age:
                        deciduous_density_ha = deciduous_density_total
                        avg_height = sum(deciduous_height) / len(deciduous_height) if deciduous_height else 0
                        avg_age = sum(deciduous_age) / len(deciduous_age) if deciduous_age else 0
                        deciduous_stats.append({'density': deciduous_density_ha, 'height': avg_height, 'age': avg_age})

        # Сводные итоги
        avg_composition = {}
        for breed, counts in breed_composition.items():
            if counts:
                avg_composition[breed] = sum(counts) / len(counts)

        composition_text = ""
        for breed in sorted(avg_composition.keys()):
            count = avg_composition[breed]
            if count > 0:
                composition_text += f"{int(count)}{breed}"

        # Обновляем итоговую строку или возвращаем данные
        # В зависимости от логики приложения
        return {
            'composition_text': composition_text,
            'forestry_formulas_text': forestry_formulas_text if 'forestry_formulas_text' in locals() else "",
            'total_plots': total_plots if 'total_plots' in locals() else 0
        }

        forestry_formulas_text = ""

        # Расчет градаций для хвойных пород
        coniferous_gradiations_stats = {'do_05_ha': [], '05_15_ha': [], 'bolee_15_ha': [], 'height': [], 'age': []}

        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)

        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                for breed_info in breeds_data:
                    if breed_info.get('type') == 'coniferous':
                        do_05_ha = breed_info.get('do_05', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        _05_15_ha = breed_info.get('05_15', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        bolee_15_ha = breed_info.get('bolee_15', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        height = breed_info.get('height', 0)
                        age = breed_info.get('age', 0)

                        coniferous_gradiations_stats['do_05_ha'].append(do_05_ha)
                        coniferous_gradiations_stats['05_15_ha'].append(_05_15_ha)
                        coniferous_gradiations_stats['bolee_15_ha'].append(bolee_15_ha)
                        if height > 0:
                            coniferous_gradiations_stats['height'].append(height)
                        if age > 0:
                            coniferous_gradiations_stats['age'].append(age)

        if coniferous_gradiations_stats['do_05_ha'] or coniferous_gradiations_stats['05_15_ha'] or coniferous_gradiations_stats['bolee_15_ha']:
            forestry_formulas_text += "Хвойные: "
            gradiations = []
            avg_do_05 = sum(coniferous_gradiations_stats['do_05_ha']) / len(coniferous_gradiations_stats['do_05_ha']) if coniferous_gradiations_stats['do_05_ha'] else 0
            gradiations.append(f"до 0.5м: {avg_do_05:.1f} шт/га")

            avg_05_15 = sum(coniferous_gradiations_stats['05_15_ha']) / len(coniferous_gradiations_stats['05_15_ha']) if coniferous_gradiations_stats['05_15_ha'] else 0
            gradiations.append(f"0.5-1.5м: {avg_05_15:.1f} шт/га")

            avg_bolee_15 = sum(coniferous_gradiations_stats['bolee_15_ha']) / len(coniferous_gradiations_stats['bolee_15_ha']) if coniferous_gradiations_stats['bolee_15_ha'] else 0
            gradiations.append(f">1.5м: {avg_bolee_15:.1f} шт/га")

            forestry_formulas_text += ", ".join(gradiations)

            if coniferous_gradiations_stats['height']:
                avg_height = sum(coniferous_gradiations_stats['height']) / len(coniferous_gradiations_stats['height'])
                forestry_formulas_text += f", высота: {avg_height:.1f}м"
            if coniferous_gradiations_stats['age']:
                avg_age = sum(coniferous_gradiations_stats['age']) / len(coniferous_gradiations_stats['age'])
                forestry_formulas_text += f", возраст: {avg_age:.1f} лет"

        # Лиственные итоги
        if deciduous_stats:
            if forestry_formulas_text:
                forestry_formulas_text += "; "
            forestry_formulas_text += "Лиственные: "
            avg_deciduous_density = sum(d['density'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0
            avg_deciduous_height = sum(d['height'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0
            avg_deciduous_age = sum(d['age'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0

            if avg_deciduous_density > 0:
                forestry_formulas_text += f"густота: {avg_deciduous_density:.1f} шт/га "
            if avg_deciduous_height > 0:
                forestry_formulas_text += f"высота: {avg_deciduous_height:.1f}м "
            if avg_deciduous_age > 0:
                forestry_formulas_text += f"возраст: {avg_deciduous_age:.1f} лет"

        return {
            'composition_text': composition_text,
            'forestry_formulas_text': forestry_formulas_text,
            'total_plots': sum(len(page) for page in self.page_data.values() if page)
        }

    def show_total_summary_popup(self, *args, **kwargs):
        """Показать popup со сводными итогами и таксационными расчетами (как в меню таксационные показатели)"""
        try:
            default_radius = float(self.current_radius) if self.current_radius else 1.78
            plot_area_m2_default = 3.14159 * (default_radius ** 2)
            trees_per_ha = 10000 / plot_area_m2_default if plot_area_m2_default > 0 else 0

            # Словарь для сбора данных по породам
            breeds_data = {}

            # Обрабатываем все страницы
            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) < 4:
                        continue

                    # Get radius for this specific plot row
                    plot_radius = default_radius  # always use default radius for consistent calculations

                    plot_area_m2 = 3.14159 * (plot_radius ** 2)
                    plot_area_ha = plot_area_m2 / 10000  # Гектары

                    # Столбец "Порода" в row[3]
                    breeds_text = row[3]
                    if not breeds_text:
                        continue

                    try:
                        breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                    except json.JSONDecodeError:
                        continue

                    for breed_info in breeds_list:
                        if not isinstance(breed_info, dict):
                            continue

                        breed_name = breed_info.get('name', '').strip()
                        if not breed_name:
                            continue

                        breed_type = breed_info.get('type', 'deciduous')
                        density = 0
                        height = None
                        age = None

                        # Расчет густоты и высоты в зависимости от типа породы
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            density = (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0

                            # Для хвойных пород определяем высоту по градациям или среднюю
                            if any([do_05, _05_15, bolee_15]):
                                # Высота определяется по градациям
                                if bolee_15 > 0:
                                    height = 2.0  # >1.5m
                                elif _05_15 > 0:
                                    height = 1.0  # 0.5-1.5m
                                elif do_05 > 0:
                                    height = 0.3  # до 0.5m
                                else:
                                    height = 0.0
                            else:
                                height = breed_info.get('height', 0) or 0
                        else:
                            # Для лиственных пород - обычная плотность и средняя высота
                            density_value = breed_info.get('density', 0)
                            density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                            height = breed_info.get('height', 0) or 0

                        age = breed_info.get('age', 0) or 0
                        diameter = breed_info.get('diameter', 0) or 0

                        # Сбор данных по породе
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {
                                'type': breed_type,
                                'plots': [],
                                'coniferous_zones': {'do_05': 0, '05_15': 0, 'bolee_15': 0} if breed_type == 'coniferous' else None,
                                'diameters': []
                            }

                        # Добавляем данные
                        plot_data = {
                            'density': density,
                            'height': height,
                            'age': age
                        }

                        if breed_type == 'coniferous':
                            plot_data.update({
                                'do_05_density': do_05 / plot_area_ha if plot_area_ha > 0 else 0,
                                '05_15_density': _05_15 / plot_area_ha if plot_area_ha > 0 else 0,
                                'bolee_15_density': bolee_15 / plot_area_ha if plot_area_ha > 0 else 0
                            })

                        breeds_data[breed_name]['plots'].append(plot_data)
                        breeds_data[breed_name]['diameters'].append(diameter)

                        if breed_type == 'coniferous':
                            breeds_data[breed_name]['coniferous_zones']['do_05'] += plot_data['do_05_density']
                            breeds_data[breed_name]['coniferous_zones']['05_15'] += plot_data['05_15_density']
                            breeds_data[breed_name]['coniferous_zones']['bolee_15'] += plot_data['bolee_15_density']

            # Создаем popup с результатами
            content = BoxLayout(orientation='vertical', spacing=10, padding=10)

            # Заголовок результатов с радиусом и адресными данными
            address_parts = []
            if self.current_quarter:
                address_parts.append(f"Квартал: {self.current_quarter}")
            if self.current_plot:
                address_parts.append(f"Выдел: {self.current_plot}")
            if self.current_forestry:
                address_parts.append(f"Лесничество: {self.current_forestry}")
            if self.current_section:
                address_parts.append(f"Участок: {self.current_section}")

            address_text = " | ".join(address_parts) if address_parts else "Адрес не указан"

            header_label = Label(
                text=f'ИТОГИ ПО УЧАСТКУ МОЛОДНЯКОВ\n' +
                     f'{address_text}\n' +
                     f'Радиус участка: {default_radius:.2f} м\n' +
                     f'1 дерево = {trees_per_ha:.0f} тыс.шт./га',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0.5, 0, 1),
                size_hint=(1, None),
                height=100,
                halign='center',
                valign='top'
            )
            header_label.bind(size=lambda *args: setattr(header_label, 'text_size', (header_label.width, None)))
            content.add_widget(header_label)

            scroll = ScrollView(size_hint=(1, None), height=600)
            results_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
            results_layout.bind(minimum_height=results_layout.setter('height'))

            # Коэффициент состава насаждения
            composition_label = Label(
                text='КОЭФФИЦИЕНТ СОСТАВА НАСАЖДЕНИЯ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=30,
                halign='center'
            )
            results_layout.add_widget(composition_label)

            # Расчет коэффициента состава на основе суммарной густоты пород
            total_densities = {}
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    if data['plots'][0].get('type') == 'coniferous':
                        # Для хвойных суммируем густоту по градациям
                        total_density = 0
                        for p in data['plots']:
                            conif_density = (p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0))
                            total_density += conif_density
                    else:
                        # Для лиственных обычная густота
                        total_density = sum(p.get('density', 0) for p in data['plots'])
                    if total_density > 0:
                        total_densities[breed_name] = total_density

            if total_densities:
                # Расчет коэффициентов состава так, чтобы их сумма равнялась 10
                total_all_density = sum(total_densities.values())
                composition_parts = []

                # Сортируем по убыванию плотности
                for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                    if total_all_density > 0:
                        # Коэффициент пропорционален плотности, сумма всех коэффициентов = 10
                        coeff = max(1, round(density / total_all_density * 10))
                    else:
                        coeff = 1
                    breed_letter = self.get_breed_letter(breed_name)
                    composition_parts.append(f"{coeff}{breed_letter}")

                # Корректировка чтобы сумма равнялась 10
                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                iterations = 0
                while total_coeffs != 10 and iterations < 100:
                    if total_coeffs > 10:
                        # Уменьшаем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] -= 1
                    elif total_coeffs < 10:
                        # Увеличиваем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] += 1

                    total_coeffs = sum(coeffs_only)
                    iterations += 1

                # Обновляем composition_parts
                sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
                composition_parts = []
                for i, (breed_name, _) in enumerate(sorted_breeds):
                    if i < len(coeffs_only):
                        breed_letter = self.get_breed_letter(breed_name)
                        composition_parts.append(f"{coeffs_only[i]}{breed_letter}")

                composition_text = ''.join(composition_parts) + "Др"
                composition_result = Label(
                    text=f"Формула состава: {composition_text}",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0, 0, 0, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(composition_result)
            else:
                no_composition = Label(
                    text="Коэффициент состава не определен (недостаточно данных)",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(1, 0, 0, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_composition)

            # Хвойные породы с градациями высот
            coniferous_label = Label(
                text='\nХВОЙНЫЕ ПОРОДЫ - ВЫСОТА ПО ГРАДАЦИЯМ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(coniferous_label)

            has_coniferous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'coniferous' and data['plots']:
                    has_coniferous = True

                    # Средняя густота в градациях
                    zones = data.get('coniferous_zones', {})
                    avg_do_05 = zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0
                    avg_05_15 = zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_bolee_15 = zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0

                    # Средняя высота только по значениям >1.5м
                    avg_heights_over_15 = [p['height'] for p in data['plots'] if p['height'] > 1.5]
                    avg_height_total = sum(avg_heights_over_15) / len(avg_heights_over_15) if avg_heights_over_15 else 0

                    # Средний диаметр
                    avg_diameter = sum(data['diameters']) / len(data['diameters']) if data['diameters'] else 0

                    coniferous_result = Label(
                        text=f"{breed_name}:\n"
                             f"• до 0.5м: {avg_do_05:.1f} шт/га\n"
                             f"• 0.5-1.5м: {avg_05_15:.1f} шт/га\n"
                             f"• >1.5м: {avg_bolee_15:.1f} шт/га\n"
                             f"• средняя высота (>1.5м): {avg_height_total:.1f}м\n"
                             f"• средний диаметр: {avg_diameter:.1f} см",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0, 0.5, 0, 1),
                        size_hint=(1, None),
                        height=120,
                        halign='left',
                        valign='top'
                    )
                    coniferous_result.bind(size=lambda *args: setattr(coniferous_result, 'text_size', (coniferous_result.width, None)))
                    results_layout.add_widget(coniferous_result)

            if not has_coniferous:
                no_coniferous = Label(
                    text="Хвойные породы не найдены",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_coniferous)

            # Лиственные породы - средние высоты и возраст
            deciduous_label = Label(
                text='\nЛИСТВЕННЫЕ ПОРОДЫ - СРЕДНИЕ ПОКАЗАТЕЛИ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(deciduous_label)

            has_deciduous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'deciduous' and data['plots']:
                    has_deciduous = True

                    avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])

                    avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                    avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0

                    avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                    avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                    # Средний диаметр
                    avg_diameter = sum(data['diameters']) / len(data['diameters']) if data['diameters'] else 0

                    deciduous_result = Label(
                        text=f"{breed_name}:\n"
                             f"• Густота: {avg_density:.1f} шт/га\n"
                             f"• Средняя высота: {avg_height:.1f}м\n"
                             f"• Средний возраст: {avg_age:.1f} лет\n"
                             f"• Средний диаметр: {avg_diameter:.1f} см",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0, 0.3, 0.5, 1),
                        size_hint=(1, None),
                        height=100,
                        halign='left',
                        valign='top'
                    )
                    deciduous_result.bind(size=lambda *args: setattr(deciduous_result, 'text_size', (deciduous_result.width, None)))
                    results_layout.add_widget(deciduous_result)

            if not has_deciduous:
                no_deciduous = Label(
                    text="Лиственные породы не найдены",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_deciduous)

            # Средние данные в целом по участку
            overall_label = Label(
                text='\nСРЕДНИЕ ДАННЫЕ В ЦЕЛОМ ПО УЧАСТКУ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(overall_label)

            # Рассчитываем общие средние
            all_densities = []
            all_heights = []
            all_ages = []
            all_diameters = []

            for breed_name, data in breeds_data.items():
                if data['plots']:
                    all_densities.extend([p['density'] for p in data['plots'] if p['density'] > 0])
                    all_heights.extend([p['height'] for p in data['plots'] if p['height'] > 0])
                    all_ages.extend([p['age'] for p in data['plots'] if p['age'] > 0])
                    all_diameters.extend([d for d in data['diameters'] if d > 0])

            avg_overall_density = sum(all_densities) / len(all_densities) if all_densities else 0
            avg_overall_height = sum(all_heights) / len(all_heights) if all_heights else 0
            avg_overall_age = sum(all_ages) / len(all_ages) if all_ages else 0
            avg_overall_diameter = sum(all_diameters) / len(all_diameters) if all_diameters else 0

            overall_result = Label(
                text=f"Средняя густота: {avg_overall_density:.1f} шт/га\n"
                     f"Средняя высота: {avg_overall_height:.1f} м\n"
                     f"Средний возраст: {avg_overall_age:.1f} лет\n"
                     f"Средний диаметр: {avg_overall_diameter:.1f} см",
                font_name='Roboto',
                font_size='14sp',
                color=(0.8, 0.4, 0, 1),
                size_hint=(1, None),
                height=100,
                halign='left',
                valign='top'
            )
            overall_result.bind(size=lambda *args: setattr(overall_result, 'text_size', (overall_result.width, None)))
            results_layout.add_widget(overall_result)



            # Предмет ухода и интенсивность рубки
            care_label = Label(
                text='\nПРЕДМЕТ УХОДА И ИНТЕНСИВНОСТЬ РУБКИ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(care_label)

            # Собираем данные по предмету ухода и рассчитываем интенсивность
            care_data = []
            total_density_all_plots = 0  # Для площадок с предметом ухода
            total_remaining_density = 0
            plot_count_with_care = 0

            # Рассчитываем среднюю густоту по всем площадкам
            total_density_all = 0
            plot_count_all = 0

            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 4 and row[3]:  # Есть данные пород
                        plot_density = 0
                        breeds_text = row[3]
                        if breeds_text:
                            try:
                                breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                for breed_info in breeds_list:
                                    if isinstance(breed_info, dict):
                                        if breed_info.get('type') == 'coniferous':
                                            # Для хвойных суммируем градации
                                            do_05 = breed_info.get('do_05', 0)
                                            _05_15 = breed_info.get('05_15', 0)
                                            bolee_15 = breed_info.get('bolee_15', 0)
                                            plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                        else:
                                            # Для лиственных обычная густота
                                            density = breed_info.get('density', 0)
                                            plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                            except (json.JSONDecodeError, TypeError):
                                pass

                        if plot_density > 0:
                            total_density_all += plot_density
                            plot_count_all += 1

                    if len(row) >= 4 and row[2]:  # Предмет ухода и данные пород
                        care_text = row[2].strip()
                        if care_text:
                            # Рассчитываем общую густоту на этой площадке
                            plot_density = 0
                            breeds_text = row[3]
                            if breeds_text:
                                try:
                                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                    for breed_info in breeds_list:
                                        if isinstance(breed_info, dict):
                                            if breed_info.get('type') == 'coniferous':
                                                # Для хвойных суммируем градации
                                                do_05 = breed_info.get('do_05', 0)
                                                _05_15 = breed_info.get('05_15', 0)
                                                bolee_15 = breed_info.get('bolee_15', 0)
                                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                            else:
                                                # Для лиственных обычная густота
                                                density = breed_info.get('density', 0)
                                                plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                                except (json.JSONDecodeError, TypeError):
                                    pass

                            if plot_density > 0:
                                # Парсим предмет ухода для получения оставляемой густоты
                                remaining_density = self.parse_care_subject_density(care_text)
                                if remaining_density > 0:
                                    care_data.append({
                                        'care_text': care_text,
                                        'plot_density': plot_density,
                                        'remaining_density': remaining_density
                                    })
                                    total_density_all_plots += plot_density
                                    total_remaining_density += remaining_density
                                    plot_count_with_care += 1

            if care_data:
                # Рассчитываем средний предмет ухода
                care_breed_totals = {}
                care_plot_count = 0

                for item in care_data:
                    care_text = item['care_text']
                    breed_densities = self.parse_care_subject_by_breeds(care_text)
                    for breed, density in breed_densities.items():
                        if breed not in care_breed_totals:
                            care_breed_totals[breed] = 0
                        care_breed_totals[breed] += density
                    care_plot_count += 1

                if care_breed_totals and care_plot_count > 0:
                    avg_care_parts = []
                    short_parts = []
                    for breed, total_density in sorted(care_breed_totals.items()):
                        avg_density = total_density / care_plot_count
                        avg_care_parts.append(f"{avg_density * 1000:.0f}шт/га{breed}")
                        short_parts.append(f"{avg_density:.1f}{breed}")
                    avg_care_text = ''.join(avg_care_parts)
                    short_text = ''.join(short_parts).replace('.', ',')

                    care_result = Label(
                        text=f"Средний предмет ухода: {avg_care_text} = {short_text}",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0.2, 0.6, 0.2, 1),
                        size_hint=(1, None),
                        height=40,
                        halign='left',
                        valign='top'
                    )
                    care_result.bind(size=lambda *args: setattr(care_result, 'text_size', (care_result.width, None)))
                    results_layout.add_widget(care_result)

                    # Рассчитываем среднюю интенсивность рубки
                if plot_count_with_care > 0 and plot_count_all > 0:
                    avg_remaining_density = total_remaining_density / plot_count_with_care

                    # Интенсивность рубки = ((общая густота - оставляемая густота) / общая густота) * 100%
                    if avg_overall_density > 0:
                        intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100

                        intensity_result = Label(
                            text=f"Средняя интенсивность рубки: {intensity:.1f}%\n"
                                 f"(было {avg_overall_density:.0f} шт/га, "
                                 f"останется {avg_remaining_density:.0f} шт/га)",
                            font_name='Roboto',
                            font_size='14sp',
                            color=(0.8, 0.2, 0.2, 1),
                            size_hint=(1, None),
                            height=60,
                            halign='left',
                            valign='top'
                        )
                        intensity_result.bind(size=lambda *args: setattr(intensity_result, 'text_size', (intensity_result.width, None)))
                        results_layout.add_widget(intensity_result)
            else:
                no_care = Label(
                    text="Предмет ухода не указан или недостаточно данных для расчета",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_care)

            # Расчет преобладающего типа леса
            forest_types_count = {}
            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 6 and row[5]:  # Тип Леса в row[5]
                        forest_type = row[5].strip()
                        if forest_type:
                            forest_types_count[forest_type] = forest_types_count.get(forest_type, 0) + 1

            predominant_forest_type = ""
            if forest_types_count:
                predominant_forest_type = max(forest_types_count.items(), key=lambda x: x[1])[0]

            # Отображение преобладающего типа леса
            forest_type_label = Label(
                text=f'\nПРЕОБЛАДАЮЩИЙ ТИП ЛЕСА: {predominant_forest_type}',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0.2, 0.4, 0.6, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(forest_type_label)

            # Расчет средних значений по типам леса
            forest_type_stats = {}
            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 6 and row[5]:  # Тип Леса в row[5]
                        forest_type = row[5].strip()
                        if forest_type:
                            if forest_type not in forest_type_stats:
                                forest_type_stats[forest_type] = {
                                    'count': 0,
                                    'total_density': 0,
                                    'total_height': 0,
                                    'total_age': 0,
                                    'valid_density': 0,
                                    'valid_height': 0,
                                    'valid_age': 0
                                }

                            # Расчет плотности для этой строки
                            plot_density = 0
                            breeds_text = row[3]
                            if breeds_text:
                                try:
                                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                    for breed_info in breeds_list:
                                        if isinstance(breed_info, dict):
                                            if breed_info.get('type') == 'coniferous':
                                                do_05 = breed_info.get('do_05', 0)
                                                _05_15 = breed_info.get('05_15', 0)
                                                bolee_15 = breed_info.get('bolee_15', 0)
                                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                            else:
                                                density = breed_info.get('density', 0)
                                                plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                                except (json.JSONDecodeError, TypeError):
                                    pass

                            forest_type_stats[forest_type]['count'] += 1
                            if plot_density > 0:
                                forest_type_stats[forest_type]['total_density'] += plot_density
                                forest_type_stats[forest_type]['valid_density'] += 1

                            # Средняя высота и возраст по породам в этой строке
                            heights = []
                            ages = []
                            if breeds_text:
                                try:
                                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                    for breed_info in breeds_list:
                                        if isinstance(breed_info, dict):
                                            if breed_info.get('height', 0) > 0:
                                                heights.append(breed_info['height'])
                                            if breed_info.get('age', 0) > 0:
                                                ages.append(breed_info['age'])
                                except (json.JSONDecodeError, TypeError):
                                    pass

                            if heights:
                                avg_height = sum(heights) / len(heights)
                                forest_type_stats[forest_type]['total_height'] += avg_height
                                forest_type_stats[forest_type]['valid_height'] += 1

                            if ages:
                                avg_age = sum(ages) / len(ages)
                                forest_type_stats[forest_type]['total_age'] += avg_age
                                forest_type_stats[forest_type]['valid_age'] += 1

            # Отображение средних значений по типам леса
            forest_types_avg_label = Label(
                text='\nСРЕДНИЕ ЗНАЧЕНИЯ ПО ТИПАМ ЛЕСА',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0.4, 0.2, 0.6, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(forest_types_avg_label)

            for forest_type, stats in sorted(forest_type_stats.items()):
                if stats['count'] > 0:
                    avg_density = stats['total_density'] / stats['valid_density'] if stats['valid_density'] > 0 else 0
                    avg_height = stats['total_height'] / stats['valid_height'] if stats['valid_height'] > 0 else 0
                    avg_age = stats['total_age'] / stats['valid_age'] if stats['valid_age'] > 0 else 0

                    is_predominant = forest_type == predominant_forest_type
                    color = (0.8, 0.4, 0.2, 1) if is_predominant else (0.3, 0.3, 0.3, 1)

                    forest_type_avg_result = Label(
                        text=f"{forest_type} ({stats['count']} площадок){' - ПРЕОБЛАДАЮЩИЙ' if is_predominant else ''}:\n"
                             f"• Средняя густота: {avg_density:.1f} шт/га\n"
                             f"• Средняя высота: {avg_height:.1f} м\n"
                             f"• Средний возраст: {avg_age:.1f} лет",
                        font_name='Roboto',
                        font_size='14sp',
                        color=color,
                        size_hint=(1, None),
                        height=80,
                        halign='left',
                        valign='top'
                    )
                    forest_type_avg_result.bind(size=lambda *args: setattr(forest_type_avg_result, 'text_size', (forest_type_avg_result.width, None)))
                    results_layout.add_widget(forest_type_avg_result)

            # Информация о площади участка и площади перечета
            plot_count = len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)])
            total_plot_area_ha = plot_count * plot_area_ha

            plot_area_info = Label(
                text=f"Информация о площади:\n"
                     f"• Радиус пробной площади: {default_radius:.2f} м\n"
                     f"• Площадь одной площадки: {plot_area_ha:.4f} га ({plot_area_m2:.4f} м²)\n"
                     f"• Всего площадок: {plot_count}\n"
                     f"• Совокупная площадь перечета: {total_plot_area_ha:.4f} га ({total_plot_area_ha*10000:.0f} м²)\n"
                     f"• Пример: 1 площадка = {plot_area_ha:.4f} га, значит площадь перечета = {plot_area_m2:.0f} м²",
                font_name='Roboto',
                font_size='12sp',
                color=(0.5, 0.5, 0.5, 1),
                size_hint=(1, None),
                height=120,
                halign='left',
                valign='top'
            )
            plot_area_info.bind(size=lambda *args: setattr(plot_area_info, 'text_size', (plot_area_info.width, None)))
            results_layout.add_widget(plot_area_info)



            scroll.add_widget(results_layout)
            content.add_widget(scroll)

            close_btn = ModernButton(
                text='Закрыть',
                bg_color=get_color_from_hex('#808080'),
                size_hint=(1, None),
                height=50
            )
            content.add_widget(close_btn)

            popup = Popup(
                title="Таксационные показатели молодняков",
                content=content,
                size_hint=(0.95, 0.95)
            )

            close_btn.bind(on_press=popup.dismiss)
            popup.open()

        except Exception as e:
            import traceback
            self.show_error(f"Ошибка расчета таксационных показателей: {str(e)}\n{traceback.format_exc()}")

    def update_totals(self, update_global=True):
        """Обновление строки итогов с поддержкой множественных пород"""
        breed_composition = {}  # Initialize at top
        total_stats = {'density': [], 'height': [], 'age': []}  # Initialize at top

        # Calculate radius and area once
        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)
        if update_global:
            section_data = self.calculate_section_totals()
            composition_text = section_data['composition_text']
            forestry_formulas_text = section_data['forestry_formulas_text']
            total_plots = section_data['total_plots']

    def generate_care_project(self, instance):
        """Генерирует проект ухода в Word документе"""
        try:
            # Сохраняем текущую страницу перед генерацией отчета
            if not self.save_current_page():
                self.show_error("Не удалось сохранить текущую страницу!")
                return

            # Собираем данные из адресной строки
            address_data = {
                'quarter': str(self.current_quarter or ''),
                'plot': str(self.current_plot or ''),
                'section': str(self.current_section or ''),
                'forestry': str(self.current_forestry or ''),
                'plot_area': str(self.plot_area_input or ''),
                'target_purpose': 'Эксплуатационные леса',  # Можно настроить
                'forest_type': 'Смешанный лес'  # Можно настроить
            }

            # Получаем итоговые данные из текущих данных приложения
            total_data = self.get_total_data_from_db()

            # Создаем временный JSON файл с данными для скрипта
            import tempfile
            import subprocess

            temp_data = {
                'address_data': address_data,
                'total_data': total_data
            }

            # Сохраняем данные во временный файл
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
                json.dump(temp_data, f, ensure_ascii=False, indent=2)
                temp_file = f.name

            # Вызываем скрипт fill_word_document.py с параметром
            script_path = os.path.join(os.path.dirname(__file__), 'fill_word_document.py')
            result = subprocess.run([
                sys.executable, script_path, '--data-file', temp_file
            ], capture_output=True, text=False)

            # Удаляем временный файл
            try:
                os.unlink(temp_file)
            except:
                pass

            # Декодируем вывод с обработкой ошибок кодировки
            def decode_output(output_bytes):
                try:
                    if output_bytes is None:
                        return ""
                    if not output_bytes:
                        return ""
                    return output_bytes.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        return output_bytes.decode('cp1251')  # Windows-1251 для русского
                    except UnicodeDecodeError:
                        return output_bytes.decode('utf-8', errors='replace')
                except Exception:
                    return ""

            if result.returncode == 0:
                stdout_text = decode_output(result.stdout)
                self.show_success(f"Проект ухода успешно создан!\n{stdout_text}")
            else:
                stderr_text = decode_output(result.stderr)
                self.show_error(f"Ошибка при создании проекта ухода:\n{stderr_text}")

        except Exception as e:
            self.show_error(f"Ошибка при генерации проекта ухода: {str(e)}")

    def get_total_data_from_db(self):
        """Получает итоговые данные из рассчитанных данных меню Итого"""
        try:
            # Используем рассчитанные данные из меню Итого вместо данных из БД
            # Получаем данные аналогично методу show_total_summary_popup

            default_radius = float(self.current_radius) if self.current_radius else 5.64
            plot_area_ha = 3.14159 * (default_radius ** 2) / 10000

            # Словарь для сбора данных по породам
            breeds_data = {}

            # Обрабатываем все страницы
            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) < 4:
                        continue

                    # Столбец "Порода" в row[3]
                    breeds_text = row[3]
                    if not breeds_text:
                        continue

                    try:
                        breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                    except json.JSONDecodeError:
                        continue

                    for breed_info in breeds_list:
                        if not isinstance(breed_info, dict):
                            continue

                        breed_name = breed_info.get('name', '').strip()
                        if not breed_name:
                            continue

                        breed_type = breed_info.get('type', 'deciduous')
                        density = 0
                        height = None
                        age = None

                        # Расчет густоты и высоты в зависимости от типа породы
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            density = (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0

                            # Для хвойных пород определяем высоту по градациям или среднюю
                            if any([do_05, _05_15, bolee_15]):
                                # Высота определяется по градациям
                                if bolee_15 > 0:
                                    height = 2.0  # >1.5m
                                elif _05_15 > 0:
                                    height = 1.0  # 0.5-1.5m
                                elif do_05 > 0:
                                    height = 0.3  # до 0.5m
                                else:
                                    height = 0.0
                            else:
                                height = breed_info.get('height', 0) or 0
                        else:
                            # Для лиственных пород - обычная плотность и средняя высота
                            density_value = breed_info.get('density', 0)
                            density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                            height = breed_info.get('height', 0) or 0

                        age = breed_info.get('age', 0) or 0
                        diameter = breed_info.get('diameter', 0) or 0

                        # Сбор данных по породе
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {
                                'type': breed_type,
                                'plots': [],
                                'coniferous_zones': {'do_05': 0, '05_15': 0, 'bolee_15': 0} if breed_type == 'coniferous' else None,
                                'diameters': []
                            }

                        # Добавляем данные
                        plot_data = {
                            'density': density,
                            'height': height,
                            'age': age
                        }

                        if breed_type == 'coniferous':
                            plot_data.update({
                                'do_05_density': do_05 / plot_area_ha if plot_area_ha > 0 else 0,
                                '05_15_density': _05_15 / plot_area_ha if plot_area_ha > 0 else 0,
                                'bolee_15_density': bolee_15 / plot_area_ha if plot_area_ha > 0 else 0
                            })

                        breeds_data[breed_name]['plots'].append(plot_data)
                        breeds_data[breed_name]['diameters'].append(diameter)

                        if breed_type == 'coniferous':
                            breeds_data[breed_name]['coniferous_zones']['do_05'] += plot_data['do_05_density']
                            breeds_data[breed_name]['coniferous_zones']['05_15'] += plot_data['05_15_density']
                            breeds_data[breed_name]['coniferous_zones']['bolee_15'] += plot_data['bolee_15_density']

            # Расчет коэффициента состава на основе суммарной густоты пород
            total_densities = {}
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    if data['plots'][0].get('type') == 'coniferous':
                        # Для хвойных суммируем густоту по градациям
                        total_density = 0
                        for p in data['plots']:
                            conif_density = (p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0))
                            total_density += conif_density
                    else:
                        # Для лиственных обычная густота
                        total_density = sum(p.get('density', 0) for p in data['plots'])
                    if total_density > 0:
                        total_densities[breed_name] = total_density

            # Расчет коэффициентов состава
            composition_text = ""
            if total_densities:
                total_all_density = sum(total_densities.values())
                composition_parts = []

                # Сортируем по убыванию плотности
                for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                    if total_all_density > 0:
                        # Коэффициент пропорционален плотности, сумма всех коэффициентов = 10
                        coeff = max(1, round(density / total_all_density * 10))
                    else:
                        coeff = 1
                    breed_letter = self.get_breed_letter(breed_name)
                    composition_parts.append(f"{coeff}{breed_letter}")

                # Корректировка чтобы сумма равнялась 10
                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                iterations = 0
                while total_coeffs != 10 and iterations < 100:
                    if total_coeffs > 10:
                        # Уменьшаем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] -= 1
                    elif total_coeffs < 10:
                        # Увеличиваем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] += 1

                    total_coeffs = sum(coeffs_only)
                    iterations += 1

                # Обновляем composition_parts
                sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
                composition_parts = []
                for i, (breed_name, _) in enumerate(sorted_breeds):
                    if i < len(coeffs_only):
                        breed_letter = self.get_breed_letter(breed_name)
                        composition_parts.append(f"{coeffs_only[i]}{breed_letter}")

                composition_text = ''.join(composition_parts) + "Др"

            # Расчет предмета ухода и интенсивности
            care_data = []
            total_density_all_plots = 0
            total_remaining_density = 0
            plot_count_with_care = 0

            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 4 and row[3]:
                        plot_density = 0
                        breeds_text = row[3]
                        if breeds_text:
                            try:
                                breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                for breed_info in breeds_list:
                                    if isinstance(breed_info, dict):
                                        if breed_info.get('type') == 'coniferous':
                                            do_05 = breed_info.get('do_05', 0)
                                            _05_15 = breed_info.get('05_15', 0)
                                            bolee_15 = breed_info.get('bolee_15', 0)
                                            plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                        else:
                                            density = breed_info.get('density', 0)
                                            plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                            except (json.JSONDecodeError, TypeError):
                                pass

                        if plot_density > 0:
                            total_density_all_plots += plot_density

                    if len(row) >= 4 and row[2]:
                        care_text = row[2].strip()
                        if care_text:
                            plot_density = 0
                            breeds_text = row[3]
                            if breeds_text:
                                try:
                                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                    for breed_info in breeds_list:
                                        if isinstance(breed_info, dict):
                                            if breed_info.get('type') == 'coniferous':
                                                do_05 = breed_info.get('do_05', 0)
                                                _05_15 = breed_info.get('05_15', 0)
                                                bolee_15 = breed_info.get('bolee_15', 0)
                                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                            else:
                                                density = breed_info.get('density', 0)
                                                plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                                except (json.JSONDecodeError, TypeError):
                                    pass

                            if plot_density > 0:
                                remaining_density = self.parse_care_subject_density(care_text)
                                if remaining_density > 0:
                                    care_data.append({
                                        'care_text': care_text,
                                        'plot_density': plot_density,
                                        'remaining_density': remaining_density
                                    })
                                    total_remaining_density += remaining_density
                                    plot_count_with_care += 1

            # Расчет среднего предмета ухода
            care_subject = ""
            intensity = 25.0  # По умолчанию

            if care_data:
                care_breed_totals = {}
                care_plot_count = 0

                for item in care_data:
                    care_text = item['care_text']
                    breed_densities = self.parse_care_subject_by_breeds(care_text)
                    for breed, density in breed_densities.items():
                        if breed not in care_breed_totals:
                            care_breed_totals[breed] = 0
                        care_breed_totals[breed] += density
                    care_plot_count += 1

                if care_breed_totals and care_plot_count > 0:
                    avg_care_parts = []
                    for breed, total_density in sorted(care_breed_totals.items()):
                        avg_density = total_density / care_plot_count
                        avg_care_parts.append(f"{avg_density * 1000:.0f}шт/га{breed}")
                    care_subject = ''.join(avg_care_parts)

                    # Расчет интенсивности
                    if plot_count_with_care > 0:
                        avg_remaining_density = total_remaining_density / plot_count_with_care
                        avg_overall_density = total_density_all_plots / len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)])

                        if avg_overall_density > 0:
                            intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100

            # Расчет средних значений по участку
            all_densities = []
            all_heights = []
            all_ages = []

            for breed_name, data in breeds_data.items():
                if data['plots']:
                    all_densities.extend([p['density'] for p in data['plots'] if p['density'] > 0])
                    all_heights.extend([p['height'] for p in data['plots'] if p['height'] > 0])
                    all_ages.extend([p['age'] for p in data['plots'] if p['age'] > 0])

            avg_overall_density = sum(all_densities) / len(all_densities) if all_densities else 0
            avg_overall_height = sum(all_heights) / len(all_heights) if all_heights else 0
            avg_overall_age = sum(all_ages) / len(all_ages) if all_ages else 0

            # Формируем итоговые данные
            total_data = {
                'page_number': self.current_page,
                'section_name': self.current_section or '',
                'total_composition': composition_text,
                'avg_age': avg_overall_age,
                'avg_density': avg_overall_density,
                'avg_height': avg_overall_height,
                'total_plots': len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)]),
                'composition': composition_text,
                'care_subject': care_subject,
                'intensity': intensity,
                'breeds': []
            }

            # Добавляем данные по породам
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])
                    avg_height = sum(p['height'] for p in data['plots'] if p['height'] > 0) / len([p for p in data['plots'] if p['height'] > 0]) if any(p['height'] > 0 for p in data['plots']) else 0
                    avg_age = sum(p['age'] for p in data['plots'] if p['age'] > 0) / len([p for p in data['plots'] if p['age'] > 0]) if any(p['age'] > 0 for p in data['plots']) else 0

                    breed_data = {
                        'name': breed_name,
                        'type': data['type'],
                        'density': avg_density,
                        'height': avg_height,
                        'age': avg_age
                    }

                    if data['type'] == 'coniferous':
                        zones = data.get('coniferous_zones', {})
                        breed_data.update({
                            'do_05': zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0,
                            '_05_15': zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0,
                            'bolee_15': zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0
                        })

                    total_data['breeds'].append(breed_data)

            return total_data

        except Exception as e:
            print(f"Ошибка получения данных из меню Итого: {e}")
            import traceback
            traceback.print_exc()
            return {}

    def parse_care_subject_density(self, care_text):
        """Парсит предмет ухода и возвращает оставляемую густоту на гектар"""
        if not care_text:
            return 0

        care_text = care_text.strip().upper()

        # Регулярное выражение для поиска чисел и букв
        # Примеры: "3С", "2Б1С", "1Е0.5С" и т.д.
        matches = re.findall(r'(\d+(?:\.\d+)?)([А-ЯA-Z]+)', care_text)

        if not matches:
            return 0

        total_density = 0
        for number_str, breed_code in matches:
            try:
                density = float(number_str)
                total_density += density
            except ValueError:
                continue

        # Предмет ухода показывает сколько деревьев оставить на гектар
        # Например, "3С" значит оставить 3000 сосен на гектар
        return total_density * 1000  # Умножаем на 1000, так как числа обычно означают тысячи деревьев

    def parse_care_subject_by_breeds(self, care_text):
        """Парсит предмет ухода и возвращает словарь {порода: густота в тыс. шт/га}"""
        if not care_text:
            return {}

        care_text = care_text.strip().upper()

        matches = re.findall(r'(\d+(?:\.\d+)?)([А-ЯA-Z]+)', care_text)

        if not matches:
            return {}

        breed_densities = {}
        for number_str, breed_code in matches:
            try:
                density = float(number_str)
                if breed_code not in breed_densities:
                    breed_densities[breed_code] = 0
                breed_densities[breed_code] += density
            except ValueError:
                continue

        return breed_densities

    def _get_current_plot_area_input(self):
        """Получить текущее значение площади участка"""
        # If stored in instance variable
        if hasattr(self, 'plot_area_input') and self.plot_area_input:
            return self.plot_area_input
        return ''
        if total_density > 0:
            summary_parts.append(f"Общая густота: {total_density}")
        if total_height > 0:
            avg_height = total_height / breed_count if breed_count > 0 else 0
            summary_parts.append(f"Средняя высота: {avg_height:.1f}м")
        if total_age > 0:
            avg_age = total_age / breed_count if breed_count > 0 else 0
            summary_parts.append(f"Средний возраст: {avg_age:.1f} лет")

        self.update_totals()
        if total_density > 0:
            summary_parts.append(f"Общая густота: {total_density}")
        if total_height > 0:
            avg_height = total_height / breed_count if breed_count > 0 else 0
            summary_parts.append(f"Средняя высота: {avg_height:.1f}м")
        if total_age > 0:
            avg_age = total_age / breed_count if breed_count > 0 else 0
            summary_parts.append(f"Средний возраст: {avg_age:.1f} лет")

        self.update_totals()

    def get_breed_letter(self, breed_name):
        """Получение первой буквы для коэффициента состава породы"""
        breed_letters = {
            'Сосна': 'С',
            'Ель': 'Е',
            'Пихта': 'П',
            'Кедр': 'К',
            'Лиственница': 'Л',
            'Берёза': 'Б',
            'Осина': 'Ос',
            'Ольха чёрная': 'ОЧ',
            'Ольха серая': 'ОС',
            'Ива': 'И',
            'Ива кустарниковая': 'ИК'
        }

        for full_name, letter in breed_letters.items():
            if full_name.lower() in breed_name.lower():
                return letter

        # Возвращаем первую букву имени породы, если не найдено
        return breed_name[0].upper() if breed_name else 'Н'

    def show_edit_plots_popup(self, instance):
        """Показать popup с списком номеров площадок для редактирования"""
        # Собираем список номеров площадок, где есть данные
        plot_numbers = []
        for row_idx in range(self.rows_per_page):
            # Проверяем, есть ли данные в основных столбцах (кроме №ППР)
            if any(self.inputs[row_idx][col].text.strip() for col in range(1, 6)):
                nn = row_idx + 1  # Номер от 1
                plot_numbers.append(nn)

        if not plot_numbers:
            self.show_error("Нет площадок для редактирования!")
            return

        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text='Выберите площадку для редактирования:',
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        scroll = ScrollView(size_hint=(1, None), height=400)
        plots_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        plots_layout.bind(minimum_height=plots_layout.setter('height'))

        for plot_num in plot_numbers:
            btn = ModernButton(
                text=f'Площадка {plot_num}',
                bg_color=get_color_from_hex('#87CEEB'),
                size_hint=(1, None),
                height=50,
                font_size='16sp'
            )
            btn.bind(on_press=lambda x, num=plot_num: self.edit_plot_popup(num - 1))  # row_index = num - 1
            plots_layout.add_widget(btn)

        scroll.add_widget(plots_layout)
        content.add_widget(scroll)

        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(1, None),
            height=50
        )
        content.add_widget(cancel_btn)

        popup = Popup(
            title="Выбор площадки для редактирования",
            content=content,
            size_hint=(0.8, 0.8)
        )

        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def edit_plot_popup(self, row_index):
        """Открыть popup редактирования для выбранной площадки"""
        MolodnikiTreeDataInputPopup(self, row_index).open()

    def save_to_json(self, instance=None):
        """Сохранение данных в JSON формате"""
        data = {
            'page_data': self.page_data,
            'section': self.current_section,
            'quarter': self.current_quarter,
            'plot': self.current_plot,
            'forestry': self.current_forestry,
            'radius': self.current_radius,
            'export_date': datetime.datetime.now().isoformat()
        }

        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        filename = f"Молодняки_{self.current_section}_{timestamp}.json"
        full_path = os.path.join(self.reports_dir, filename)

        try:
            with open(full_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return f"JSON: {filename}", None
        except Exception as e:
            return None, f"Ошибка сохранения JSON: {str(e)}"

    def save_to_excel_without_dialog(self):
        """Сохранение в Excel без диалога"""
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        filename = f"Молодняки_{self.current_section}_{timestamp}.xlsx"
        full_path = os.path.join(self.reports_dir, filename)

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Молодняки"

            address_parts = []
            if self.current_quarter:
                address_parts.append(f"Квартал: {self.current_quarter}")
            if self.current_plot:
                address_parts.append(f"Выдел: {self.current_plot}")
            if self.current_forestry:
                address_parts.append(f"Лесничество: {self.current_forestry}")
            if self.current_radius:
                address_parts.append(f"Радиус: {self.current_radius} м")

            address_text = " | ".join(address_parts) if address_parts else "Адрес не указан"
            ws['A1'] = f"Адрес: {address_text}"
            ws['A1'].font = openpyxl.styles.Font(bold=True, size=12)

            # Расчет площади перечета
            current_radius = float(self.current_radius) if self.current_radius else 5.64
            plot_area_m2 = 3.14159 * (current_radius ** 2)
            plot_area_ha = plot_area_m2 / 10000
            plot_count = len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)])
            total_plot_area_ha = plot_count * plot_area_ha

            ws['A2'] = f"Площадь перечета: {total_plot_area_ha:.4f} га ({total_plot_area_ha*10000:.0f} м²) - {plot_count} площадок по {plot_area_ha:.4f} га каждая"
            ws['A2'].font = openpyxl.styles.Font(bold=True, size=10)

            ws.append([])

            headers = [
                '№ППР', 'GPS точка', 'Предмет ухода', 'Порода', 'Густота', 'До 0.5м', '0.5-1.5м', '>1.5м', 'Высота', 'Возраст', 'Примечания', 'Тип Леса'
            ]
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_num, value=header)
                cell.font = openpyxl.styles.Font(bold=True)
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

            all_data = []
            for page in sorted(self.page_data.keys()):
                all_data.extend(self.page_data[page])

            current_row = 4
            for row in all_data:
                if any(cell for cell in row[:3] if cell):  # Проверяем, что основные столбцы не пустые
                    try:
                        breeds_data = json.loads(row[3]) if row[3] else []
                    except (json.JSONDecodeError, TypeError):
                        breeds_data = []

                    if isinstance(breeds_data, list) and breeds_data:
                        for breed_info in breeds_data:
                            if isinstance(breed_info, dict):
                                breed_name = breed_info.get('name', 'Неизвестная')
                                density = breed_info.get('density', '')
                                height = breed_info.get('height', '')
                                age = breed_info.get('age', '')

                                # Инициализируем градации
                                do_05 = ''
                                _05_15 = ''
                                bolee_15 = ''

                                if breed_info.get('type') == 'coniferous':
                                    # Для хвойных заполняем градации
                                    do_05 = str(breed_info.get('do_05', ''))
                                    _05_15 = str(breed_info.get('05_15', ''))
                                    bolee_15 = str(breed_info.get('bolee_15', ''))
                                    # Густота оставляем пустой для хвойных
                                    density = ''
                                else:
                                    # Для лиственных оставляем густоту, градации пустые
                                    pass

                                processed_row = [
                                    row[0],  # №ППР
                                    row[1],  # GPS точка
                                    row[2],  # Предмет ухода
                                    breed_name,  # Порода
                                    str(density) if density else '',  # Густота
                                    do_05,  # До 0.5м
                                    _05_15,  # 0.5-1.5м
                                    bolee_15,  # >1.5м
                                    str(height) if height else '',  # Высота
                                    str(age) if age else '',  # Возраст
                                    row[4],  # Примечания
                                    row[5],  # Тип Леса
                                ]
                                ws.append(processed_row)
                                current_row += 1
                    else:
                        # Если нет пород, добавить строку без данных
                        processed_row = [row[0], row[1], row[2], '', '', '', '', '', '', '', row[4], row[5]]
                        ws.append(processed_row)
                        current_row += 1

            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(full_path)
            return f"Excel: {filename}", None
        except Exception as e:
            return None, f"Ошибка сохранения Excel: {str(e)}"

    def save_to_word_without_dialog(self):
        """Сохранение в Word без диалога"""
        try:
            from docx import Document

            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
            filename = f"Молодняки_{self.current_section}_{timestamp}.docx"
            full_path = os.path.join(self.reports_dir, filename)

            doc = Document()
            doc.add_heading(f'Расширенный отчет по молоднякам - Участок {self.current_section}', 0)

            # Расчет площади перечета
            current_radius = float(self.current_radius) if self.current_radius else 5.64
            plot_area_m2 = 3.14159 * (current_radius ** 2)
            plot_area_ha = plot_area_m2 / 10000
            plot_count = len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)])
            total_plot_area_ha = plot_count * plot_area_ha

            # Добавляем информацию о площади перечета
            doc.add_paragraph(f"Площадь перечета: {total_plot_area_ha:.4f} га ({total_plot_area_ha*10000:.0f} м²) - {plot_count} площадок по {plot_area_ha:.4f} га каждая")

            all_data = []
            for page in sorted(self.page_data.keys()):
                all_data.extend(self.page_data[page])

            table = doc.add_table(rows=1, cols=12)
            table.style = 'Table Grid'

            headers = [
                '№ППР', 'GPS точка', 'Предмет ухода', 'Порода', 'Густота', 'До 0.5м', '0.5-1.5м', '>1.5м', 'Высота', 'Возраст', 'Примечания', 'Тип Леса'
            ]
            hdr_cells = table.rows[0].cells
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            for row in all_data:
                if any(cell for cell in row[:3] if cell):  # Проверяем, что основные столбцы не пустые
                    try:
                        breeds_data = json.loads(row[3]) if row[3] else []
                    except (json.JSONDecodeError, TypeError):
                        breeds_data = []

                    if isinstance(breeds_data, list) and breeds_data:
                        for breed_info in breeds_data:
                            if isinstance(breed_info, dict):
                                breed_name = breed_info.get('name', 'Неизвестная')
                                density = breed_info.get('density', '')
                                height = breed_info.get('height', '')
                                age = breed_info.get('age', '')

                                # Инициализируем градации
                                do_05 = ''
                                _05_15 = ''
                                bolee_15 = ''

                                if breed_info.get('type') == 'coniferous':
                                    # Для хвойных заполняем градации
                                    do_05 = str(breed_info.get('do_05', ''))
                                    _05_15 = str(breed_info.get('05_15', ''))
                                    bolee_15 = str(breed_info.get('bolee_15', ''))
                                    # Густота оставляем пустой для хвойных
                                    density = ''
                                else:
                                    # Для лиственных оставляем густоту, градации пустые
                                    pass

                                row_cells = table.add_row().cells
                                row_cells[0].text = str(row[0]) if row[0] else ""  # №ППР
                                row_cells[1].text = str(row[1]) if row[1] else ""  # GPS точка
                                row_cells[2].text = str(row[2]) if row[2] else ""  # Предмет ухода
                                row_cells[3].text = breed_name  # Порода
                                row_cells[4].text = str(density) if density else ""  # Густота
                                row_cells[5].text = do_05  # До 0.5м
                                row_cells[6].text = _05_15  # 0.5-1.5м
                                row_cells[7].text = bolee_15  # >1.5м
                                row_cells[8].text = str(height) if height else ""  # Высота
                                row_cells[9].text = str(age) if age else ""  # Возраст
                                row_cells[10].text = str(row[4]) if row[4] else ""  # Примечания
                                row_cells[11].text = str(row[5]) if row[5] else ""  # Тип Леса
                    else:
                        # Если нет пород, добавить строку без данных
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(row[0]) if row[0] else ""
                        row_cells[1].text = str(row[1]) if row[1] else ""
                        row_cells[2].text = str(row[2]) if row[2] else ""
                        row_cells[3].text = ""
                        row_cells[4].text = ""
                        row_cells[5].text = ""
                        row_cells[6].text = ""
                        row_cells[7].text = ""
                        row_cells[8].text = ""
                        row_cells[9].text = ""
                        row_cells[10].text = str(row[4]) if row[4] else ""
                        row_cells[11].text = str(row[5]) if row[5] else ""

            doc.save(full_path)
            return f"Word: {filename}", None
        except ImportError:
            return None, "Для сохранения в Word установите библиотеку python-docx: pip install python-docx"
        except Exception as e:
            return None, f"Ошибка сохранения Word: {str(e)}"

    def save_all_formats(self, instance=None):
        """Сохранить данные во всех форматах сразу"""
        success_messages = []
        error_messages = []

        # Проверка наличия данных
        if not self.page_data:
            error_messages.append("Нет данных для сохранения!")
            self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
            return

        # Проверка наличия участка
        if not self.current_section:
            error_messages.append("Не указан номер участка!")
            self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
            return

        # Валидация радиуса
        try:
            radius = float(self.current_radius) if self.current_radius else 5.64
            if radius <= 0:
                error_messages.append("Радиус должен быть положительным числом!")
                self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
                return
        except (ValueError, TypeError):
            error_messages.append("Некорректное значение радиуса!")
            self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
            return

        # Автоматически сохраняем текущую страницу перед сохранением отчёта
        if not self.save_current_page():
            error_messages.append("Не удалось сохранить текущую страницу в базу данных!")
            self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
            return

        # Проверка существования папки reports
        if not os.path.exists(self.reports_dir):
            try:
                os.makedirs(self.reports_dir, exist_ok=True)
            except Exception as e:
                error_messages.append(f"Не удалось создать папку reports: {str(e)}")
                self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))
                return

        try:
            # Сохранение в JSON
            result, error = self.save_to_json()
            if result:
                success_messages.append(result)
            else:
                error_messages.append(error)

            # Сохранение в Excel
            result, error = self.save_to_excel_without_dialog()
            if result:
                success_messages.append(result)
            else:
                error_messages.append(error)

            # Сохранение в Word
            result, error = self.save_to_word_without_dialog()
            if result:
                success_messages.append(result)
            else:
                error_messages.append(error)

        except Exception as e:
            import traceback
            error_messages.append(f"Общая ошибка: {str(e)}\n{traceback.format_exc()}")

        if success_messages:
            self.show_success("Файлы сохранены:\n" + "\n".join(success_messages))
        if error_messages:
            self.show_error("Ошибки сохранения:\n" + "\n".join(error_messages))

    def show_edit_breed_popup(self, instance, breed_index, breed_info):
        """Показать popup для редактирования породы"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text=f"Редактирование породы: {breed_info.get('name', '')}",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        fields_layout = GridLayout(cols=2, spacing=5, size_hint=(1, None), height=200)

        breed_type = breed_info.get('type', 'deciduous')
        if breed_type == 'coniferous':
            fields = [
                ('До 0.5м:', 'do_05'),
                ('0.5-1.5м:', '05_15'),
                ('>1.5м:', 'bolee_15'),
                ('Высота (м):', 'height'),
                ('Густота:', 'density'),
                ('Возраст (лет):', 'age')
            ]
        else:
            fields = [
                ('Густота:', 'density'),
                ('Высота (м):', 'height'),
                ('Возраст (лет):', 'age')
            ]

        self.edit_inputs = {}
        for label_text, field_key in fields:
            lbl = Label(text=label_text, font_name='Roboto', size_hint=(None, None), size=(100, 30))
            inp = TextInput(
                multiline=False,
                size_hint=(None, None),
                size=(100, 30),
                text=str(breed_info.get(field_key, ''))
            )
            if field_key in ['density', 'age']:
                inp.input_filter = 'int'
            elif field_key == 'height':
                inp.input_filter = 'float'
            fields_layout.add_widget(lbl)
            fields_layout.add_widget(inp)
            self.edit_inputs[field_key] = inp

        content.add_widget(fields_layout)

        # Кнопки
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title=f"Редактирование породы",
            content=content,
            size_hint=(0.9, 0.8)
        )

        def save_edit(btn):
            for key, inp in self.edit_inputs.items():
                if inp.text.strip():
                    try:
                        if key in ['density', 'age']:
                            breed_info[key] = int(inp.text)
                        elif key == 'height':
                            breed_info[key] = float(inp.text)
                        else:
                            breed_info[key] = float(inp.text)
                    except ValueError:
                        breed_info[key] = inp.text
                else:
                    breed_info[key] = 0 if key in ['density', 'age', 'do_05', '05_15', 'bolee_15'] else 0.0

            breeds_data = self.parse_breeds_data(instance.text)
            if 0 <= breed_index < len(breeds_data):
                breeds_data[breed_index] = breed_info
                instance.text = json.dumps(breeds_data, ensure_ascii=False, indent=2)
                self.update_totals()
                self.show_success("Порода обновлена!")
                if hasattr(self.table_screen, 'popup') and self.table_screen.popup:
                    self.table_screen.popup.dismiss()
                popup.dismiss()

        save_btn.bind(on_press=save_edit)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_custom_breed_popup(self, instance, breed_type):
        """Показать popup для ввода названия другой породы"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите название другой породы",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.custom_breed_input = TextInput(
            hint_text="Название породы",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto'
        )
        content.add_widget(self.custom_breed_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title=f"Ввод {'хвойной' if breed_type == 'coniferous' else 'лиственной'} породы",
            content=content,
            size_hint=(0.8, 0.6)
        )

        def save_custom_breed(btn):
            breed_name = self.custom_breed_input.text.strip()
            if breed_name:
                # Проверяем, не является ли порода запрещенной
                forbidden_breeds = ['семенная', 'культуры', 'подрост']
                if any(forbidden.lower() in breed_name.lower() for forbidden in forbidden_breeds):
                    self.show_error("Эта порода не разрешена для использования!")
                    return
                instance.text = breed_name
                self.show_breed_details_popup(instance, breed_type, breed_name)
                popup.dismiss()
            else:
                self.show_error("Название породы не может быть пустым!")

        save_btn.bind(on_press=save_custom_breed)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def update_coniferous_density(self, instance, value):
        """Автоматический расчет густоты для хвойных пород"""
        if 'density' in self.breed_inputs:
            density_input = self.breed_inputs['density']
            try:
                do_05 = int(self.breed_inputs.get('do_05', TextInput(text='0')).text or '0')
                _05_15 = int(self.breed_inputs.get('05_15', TextInput(text='0')).text or '0')
                bolee_15 = int(self.breed_inputs.get('bolee_15', TextInput(text='0')).text or '0')

                total_density = do_05 + _05_15 + bolee_15
                density_input.text = str(total_density) if total_density > 0 else ''
            except (ValueError, AttributeError):
                pass

    def update_address_label(self):
        """Обновить текст адресной строки"""
        address_parts = []
        if self.current_quarter:
            address_parts.append(f"{self.current_quarter} кв.")
        if self.current_plot:
            address_parts.append(f"{self.current_plot} выд.")
        if self.current_forestry:
            address_parts.append(self.current_forestry)
        if getattr(self, 'current_district_forestry', ''):
            address_parts.append(f"участковое: {self.current_district_forestry}")

        address_text = "Адрес: " + " ".join(address_parts) if address_parts else "Адрес: не указан"
        self.address_label.text = address_text

    def load_existing_data(self):
        """Загружаем существующие данные из базы данных"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        try:
            cursor.execute('''
                SELECT DISTINCT page_number FROM molodniki_data
                WHERE section_name = ?
                ORDER BY page_number
            ''', (self.current_section,))

            page_numbers = [row[0] for row in cursor.fetchall()]

            if page_numbers:
                for page_num in page_numbers:
                    cursor.execute('''
                        SELECT row_index, nn, gps_point, predmet_uhoda, poroda, primechanie, radius
                        FROM molodniki_data
                        WHERE page_number = ? AND section_name = ?
                        ORDER BY row_index
                    ''', (page_num, self.current_section))

                    page_data = []
                    rows_data = cursor.fetchall()

                    for row_idx in range(self.rows_per_page):
                        row_data = ['', '', '', '', '', '']
                        page_data.append(row_data)

                    for row_data in rows_data:
                        row_idx, nn, gps_point, predmet_uhoda, poroda, primechanie, radius = row_data
                        if row_idx < len(page_data):
                            page_data[row_idx] = [
                                str(nn) if nn is not None else '',
                                str(gps_point) if gps_point is not None else '',
                                str(predmet_uhoda) if predmet_uhoda is not None else '',
                                str(poroda) if poroda is not None else '',
                                str(primechanie) if primechanie is not None else '',
                                str(radius) if radius is not None else '',
                            ]

                    self.page_data[page_num] = page_data

                self.current_page = min(page_numbers)
                self.load_page_data()
                self.update_pagination()

        except Exception as e:
            print(f"Error loading existing data: {e}")
        finally:
            conn.close()

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

        # Обновляем итоги по каждой строке
        for row_idx in range(len(self.inputs)):
            self.update_row_total(self.inputs[row_idx][3], self.inputs[row_idx][3].text)

        self.update_totals()

    def clear_table_data(self, instance=None):
        for row in self.inputs:
            for inp in row:
                inp.text = ''
        self.page_data.clear()
        self.update_totals()
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

    def show_quarter_popup(self, instance):
        """Показать popup для ввода квартала"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите номер квартала",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.quarter_input = TextInput(
            hint_text="Номер квартала",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            input_filter='int',
            text=self.current_quarter
        )
        content.add_widget(self.quarter_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=40)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка квартала",
            content=content,
            size_hint=(0.6, 0.5)
        )

        def save_quarter(btn):
            quarter = self.quarter_input.text.strip()
            if quarter:
                self.current_quarter = quarter
                self.update_address_label()
                self.show_success(f"Квартал установлен: {quarter}")
                popup.dismiss()
            else:
                self.show_error("Номер квартала не может быть пустым!")

        save_btn.bind(on_press=save_quarter)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_plot_popup(self, instance):
        """Показать popup для ввода выдела"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите номер выдела",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.plot_input = TextInput(
            hint_text="Номер выдела",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            input_filter='int',
            text=self.current_plot
        )
        content.add_widget(self.plot_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка выдела",
            content=content,
            size_hint=(0.6, 0.5)
        )

        def save_plot(btn):
            plot = self.plot_input.text.strip()
            if plot:
                self.current_plot = plot
                self.update_address_label()
                self.show_success(f"Выдел установлен: {plot}")
                popup.dismiss()
            else:
                self.show_error("Номер выдела не может быть пустым!")

        save_btn.bind(on_press=save_plot)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_forestry_popup(self, instance):
        """Показать popup для ввода лесничества и участкового лесничества"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите название лесничества и участкового лесничества",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        # Поле для лесничества
        forestry_label = Label(
            text="Лесничество:",
            font_name='Roboto',
            size_hint=(1, None),
            height=25,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(forestry_label)

        self.forestry_input = TextInput(
            hint_text="Название лесничества",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            text=self.current_forestry
        )
        content.add_widget(self.forestry_input)

        # Поле для участкового лесничества
        district_forestry_label = Label(
            text="Участковое лесничество:",
            font_name='Roboto',
            size_hint=(1, None),
            height=25,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(district_forestry_label)

        self.district_forestry_input = TextInput(
            hint_text="Название участкового лесничества",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            text=getattr(self, 'current_district_forestry', '')
        )
        content.add_widget(self.district_forestry_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка лесничества",
            content=content,
            size_hint=(0.6, 0.7)
        )

        def save_forestry(btn):
            forestry = self.forestry_input.text.strip()
            district_forestry = self.district_forestry_input.text.strip()
            if forestry:
                self.current_forestry = forestry
                self.current_district_forestry = district_forestry
                self.update_address_label()
                self.show_success(f"Лесничество установлено: {forestry}" + (f", участковое: {district_forestry}" if district_forestry else ""))
                popup.dismiss()
            else:
                self.show_error("Название лесничества не может быть пустым!")

        save_btn.bind(on_press=save_forestry)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_district_forestry_popup(self):
        """Показать popup для ввода участкового лесничества"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите название участкового лесничества",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.district_forestry_input = TextInput(
            hint_text="Название участкового лесничества",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto',
            text=getattr(self, 'current_district_forestry', '')
        )
        content.add_widget(self.district_forestry_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка участкового лесничества",
            content=content,
            size_hint=(0.6, 0.4)
        )

        def save_district_forestry(btn):
            district_forestry = self.district_forestry_input.text.strip()
            self.current_district_forestry = district_forestry
            self.update_address_label()
            if district_forestry:
                self.show_success(f"Участковое лесничество установлено: {district_forestry}")
            else:
                self.show_success("Участковое лесничество очищено")
            popup.dismiss()

        save_btn.bind(on_press=save_district_forestry)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def calculate_section_totals(self):
        """Расчет итогов по всему разделу (все страницы)"""
        breed_composition = {}
        total_stats = {'density': [], 'height': [], 'age': []}
        coniferous_stats = {'do_05': [], '05_15': [], 'bolee_15': [], 'height': [], 'age': []}

        for page_num, page_data in self.page_data.items():
            for row_data in page_data:
                # row_data имеет 6 элементов: nn, gps_point, predmet_uhoda, poroda, primechanie, tip_lesa
                if row_data[2]:  # predmet_uhoda
                    composition = self.parse_composition(row_data[2])
                    for breed, count in composition.items():
                        if breed not in breed_composition:
                            breed_composition[breed] = []
                        breed_composition[breed].append(count)

                breeds_text = row_data[3]  # poroda
                if breeds_text:
                    breeds_data = self.parse_breeds_data(breeds_text)
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'coniferous':
                            # Густота хвойных = сумма градаций
                            coniferous_density = (breed_info.get('do_05', 0) +
                                                breed_info.get('05_15', 0) +
                                                breed_info.get('bolee_15', 0))
                            if coniferous_density > 0:
                                total_stats['density'].append(coniferous_density)
                            else:
                                if 'density' in breed_info and breed_info['density']:
                                    total_stats['density'].append(breed_info['density'])

                        elif 'density' in breed_info and breed_info['density']:
                            total_stats['density'].append(breed_info['density'])

                        if 'height' in breed_info and breed_info['height']:
                            total_stats['height'].append(breed_info['height'])
                        if 'age' in breed_info and breed_info['age']:
                            total_stats['age'].append(breed_info['age'])

        # Рассчитываем остальные итоги

        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)  # Площадь пробной площади в м²

        # Расчет средних по градациям для хвойных по формулам лесного хозяйства на гектар
        coniferous_stats_ha = []
        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                if breeds_data and any(b.get('type') == 'coniferous' for b in breeds_data):
                    coniferous_density_ha = 0
                    height_sum = 0
                    age_sum = 0
                    count = 0
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            coniferous_density_ha += (do_05 * 10000 / plot_area_m2) + (_05_15 * 10000 / plot_area_m2) + (bolee_15 * 10000 / plot_area_m2)
                            if breed_info.get('height'):
                                height_sum += breed_info['height']
                                count += 1
                            if breed_info.get('age'):
                                age_sum += breed_info['age']

                    coniferous_stats_ha.append({
                        'density_ha': coniferous_density_ha if coniferous_density_ha > 0 else 0,
                        'height': height_sum / count if count > 0 else 0,
                        'age': age_sum / count if count > 0 else 0
                    })

        # Итоги по лиственным
        deciduous_stats = []
        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                if breeds_data:
                    deciduous_density_total = 0
                    deciduous_height = []
                    deciduous_age = []
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'deciduous':
                            deciduous_density_total += breed_info.get('density', 0)
                            if breed_info.get('height', 0) > 0:
                                deciduous_height.append(breed_info['height'])
                            if breed_info.get('age', 0) > 0:
                                deciduous_age.append(breed_info['age'])
                    if deciduous_height or deciduous_age:
                        deciduous_density_ha = deciduous_density_total
                        avg_height = sum(deciduous_height) / len(deciduous_height) if deciduous_height else 0
                        avg_age = sum(deciduous_age) / len(deciduous_age) if deciduous_age else 0
                        deciduous_stats.append({'density': deciduous_density_ha, 'height': avg_height, 'age': avg_age})

        # Сводные итоги
        avg_composition = {}
        for breed, counts in breed_composition.items():
            if counts:
                avg_composition[breed] = sum(counts) / len(counts)

        composition_text = ""
        for breed in sorted(avg_composition.keys()):
            count = avg_composition[breed]
            if count > 0:
                composition_text += f"{int(count)}{breed}"

        # Обновляем итоговую строку или возвращаем данные
        # В зависимости от логики приложения
        return {
            'composition_text': composition_text,
            'forestry_formulas_text': forestry_formulas_text if 'forestry_formulas_text' in locals() else "",
            'total_plots': total_plots if 'total_plots' in locals() else 0
        }

        forestry_formulas_text = ""

        # Расчет градаций для хвойных пород
        coniferous_gradiations_stats = {'do_05_ha': [], '05_15_ha': [], 'bolee_15_ha': [], 'height': [], 'age': []}

        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)

        for row_data in [row for page in self.page_data.values() for row in page]:
            breeds_text = row_data[3]
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                for breed_info in breeds_data:
                    if breed_info.get('type') == 'coniferous':
                        do_05_ha = breed_info.get('do_05', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        _05_15_ha = breed_info.get('05_15', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        bolee_15_ha = breed_info.get('bolee_15', 0) * 10000 / plot_area_m2 if plot_area_m2 > 0 else 0
                        height = breed_info.get('height', 0)
                        age = breed_info.get('age', 0)

                        coniferous_gradiations_stats['do_05_ha'].append(do_05_ha)
                        coniferous_gradiations_stats['05_15_ha'].append(_05_15_ha)
                        coniferous_gradiations_stats['bolee_15_ha'].append(bolee_15_ha)
                        if height > 0:
                            coniferous_gradiations_stats['height'].append(height)
                        if age > 0:
                            coniferous_gradiations_stats['age'].append(age)

        if coniferous_gradiations_stats['do_05_ha'] or coniferous_gradiations_stats['05_15_ha'] or coniferous_gradiations_stats['bolee_15_ha']:
            forestry_formulas_text += "Хвойные: "
            gradiations = []
            avg_do_05 = sum(coniferous_gradiations_stats['do_05_ha']) / len(coniferous_gradiations_stats['do_05_ha']) if coniferous_gradiations_stats['do_05_ha'] else 0
            gradiations.append(f"до 0.5м: {avg_do_05:.1f} шт/га")

            avg_05_15 = sum(coniferous_gradiations_stats['05_15_ha']) / len(coniferous_gradiations_stats['05_15_ha']) if coniferous_gradiations_stats['05_15_ha'] else 0
            gradiations.append(f"0.5-1.5м: {avg_05_15:.1f} шт/га")

            avg_bolee_15 = sum(coniferous_gradiations_stats['bolee_15_ha']) / len(coniferous_gradiations_stats['bolee_15_ha']) if coniferous_gradiations_stats['bolee_15_ha'] else 0
            gradiations.append(f">1.5м: {avg_bolee_15:.1f} шт/га")

            forestry_formulas_text += ", ".join(gradiations)

            if coniferous_gradiations_stats['height']:
                avg_height = sum(coniferous_gradiations_stats['height']) / len(coniferous_gradiations_stats['height'])
                forestry_formulas_text += f", высота: {avg_height:.1f}м"
            if coniferous_gradiations_stats['age']:
                avg_age = sum(coniferous_gradiations_stats['age']) / len(coniferous_gradiations_stats['age'])
                forestry_formulas_text += f", возраст: {avg_age:.1f} лет"

        # Лиственные итоги
        if deciduous_stats:
            if forestry_formulas_text:
                forestry_formulas_text += "; "
            forestry_formulas_text += "Лиственные: "
            avg_deciduous_density = sum(d['density'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0
            avg_deciduous_height = sum(d['height'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0
            avg_deciduous_age = sum(d['age'] for d in deciduous_stats) / len(deciduous_stats) if deciduous_stats else 0

            if avg_deciduous_density > 0:
                forestry_formulas_text += f"густота: {avg_deciduous_density:.1f} шт/га "
            if avg_deciduous_height > 0:
                forestry_formulas_text += f"высота: {avg_deciduous_height:.1f}м "
            if avg_deciduous_age > 0:
                forestry_formulas_text += f"возраст: {avg_deciduous_age:.1f} лет"

        return {
            'composition_text': composition_text,
            'forestry_formulas_text': forestry_formulas_text,
            'total_plots': sum(len(page) for page in self.page_data.values() if page)
        }

    def show_total_summary_popup(self, *args, **kwargs):
        """Показать popup со сводными итогами и таксационными расчетами (как в меню таксационные показатели)"""
        try:
            default_radius = float(self.current_radius) if self.current_radius else 1.78
            plot_area_m2_default = 3.14159 * (default_radius ** 2)
            trees_per_ha = 10000 / plot_area_m2_default if plot_area_m2_default > 0 else 0

            # Словарь для сбора данных по породам
            breeds_data = {}

            # Обрабатываем все страницы
            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) < 4:
                        continue

                    # Get radius for this specific plot row
                    plot_radius = default_radius  # always use default radius for consistent calculations

                    plot_area_m2 = 3.14159 * (plot_radius ** 2)
                    plot_area_ha = plot_area_m2 / 10000  # Гектары

                    # Столбец "Порода" в row[3]
                    breeds_text = row[3]
                    if not breeds_text:
                        continue

                    try:
                        breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                    except json.JSONDecodeError:
                        continue

                    for breed_info in breeds_list:
                        if not isinstance(breed_info, dict):
                            continue

                        breed_name = breed_info.get('name', '').strip()
                        if not breed_name:
                            continue

                        breed_type = breed_info.get('type', 'deciduous')
                        density = 0
                        height = None
                        age = None

                        # Расчет густоты и высоты в зависимости от типа породы
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            density = (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0

                            # Для хвойных пород определяем высоту по градациям или среднюю
                            if any([do_05, _05_15, bolee_15]):
                                # Высота определяется по градациям
                                if bolee_15 > 0:
                                    height = 2.0  # >1.5m
                                elif _05_15 > 0:
                                    height = 1.0  # 0.5-1.5m
                                elif do_05 > 0:
                                    height = 0.3  # до 0.5m
                                else:
                                    height = 0.0
                            else:
                                height = breed_info.get('height', 0) or 0
                        else:
                            # Для лиственных пород - обычная плотность и средняя высота
                            density_value = breed_info.get('density', 0)
                            density = density_value / plot_area_ha if plot_area_ha > 0 else 0
                            height = breed_info.get('height', 0) or 0

                        age = breed_info.get('age', 0) or 0
                        diameter = breed_info.get('diameter', 0) or 0

                        # Сбор данных по породе
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {
                                'type': breed_type,
                                'plots': [],
                                'coniferous_zones': {'do_05': 0, '05_15': 0, 'bolee_15': 0} if breed_type == 'coniferous' else None,
                                'diameters': []
                            }

                        # Добавляем данные
                        plot_data = {
                            'density': density,
                            'height': height,
                            'age': age
                        }

                        if breed_type == 'coniferous':
                            plot_data.update({
                                'do_05_density': do_05 / plot_area_ha if plot_area_ha > 0 else 0,
                                '05_15_density': _05_15 / plot_area_ha if plot_area_ha > 0 else 0,
                                'bolee_15_density': bolee_15 / plot_area_ha if plot_area_ha > 0 else 0
                            })

                        breeds_data[breed_name]['plots'].append(plot_data)
                        breeds_data[breed_name]['diameters'].append(diameter)

                        if breed_type == 'coniferous':
                            breeds_data[breed_name]['coniferous_zones']['do_05'] += plot_data['do_05_density']
                            breeds_data[breed_name]['coniferous_zones']['05_15'] += plot_data['05_15_density']
                            breeds_data[breed_name]['coniferous_zones']['bolee_15'] += plot_data['bolee_15_density']

            # Создаем popup с результатами
            content = BoxLayout(orientation='vertical', spacing=10, padding=10)

            # Заголовок результатов с радиусом
            header_label = Label(
                text=f'ИТОГИ ПО УЧАСТКУ МОЛОДНЯКОВ\n' +
                     f'Радиус участка: {default_radius:.2f} м\n' +
                     f'1 дерево = {trees_per_ha:.0f} тыс.шт./га',
                font_name='Roboto',
                font_size='18sp',
                bold=True,
                color=(0, 0.5, 0, 1),
                size_hint=(1, None),
                height=80,
                halign='center',
                valign='top'
            )
            content.add_widget(header_label)

            scroll = ScrollView(size_hint=(1, None), height=600)
            results_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
            results_layout.bind(minimum_height=results_layout.setter('height'))

            # Коэффициент состава насаждения
            composition_label = Label(
                text='КОЭФФИЦИЕНТ СОСТАВА НАСАЖДЕНИЯ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=30,
                halign='center'
            )
            results_layout.add_widget(composition_label)

            # Расчет коэффициента состава на основе суммарной густоты пород
            total_densities = {}
            for breed_name, data in breeds_data.items():
                if data['plots']:
                    if data['plots'][0].get('type') == 'coniferous':
                        # Для хвойных суммируем густоту по градациям
                        total_density = 0
                        for p in data['plots']:
                            conif_density = (p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0))
                            total_density += conif_density
                    else:
                        # Для лиственных обычная густота
                        total_density = sum(p.get('density', 0) for p in data['plots'])
                    if total_density > 0:
                        total_densities[breed_name] = total_density

            if total_densities:
                # Расчет коэффициентов состава так, чтобы их сумма равнялась 10
                total_all_density = sum(total_densities.values())
                composition_parts = []

                # Сортируем по убыванию плотности
                for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                    if total_all_density > 0:
                        # Коэффициент пропорционален плотности, сумма всех коэффициентов = 10
                        coeff = max(1, round(density / total_all_density * 10))
                    else:
                        coeff = 1
                    breed_letter = self.get_breed_letter(breed_name)
                    composition_parts.append(f"{coeff}{breed_letter}")

                # Корректировка чтобы сумма равнялась 10
                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                iterations = 0
                while total_coeffs != 10 and iterations < 100:
                    if total_coeffs > 10:
                        # Уменьшаем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] -= 1
                    elif total_coeffs < 10:
                        # Увеличиваем самый большой коэффициент
                        max_idx = coeffs_only.index(max(coeffs_only))
                        coeffs_only[max_idx] += 1

                    total_coeffs = sum(coeffs_only)
                    iterations += 1

                # Обновляем composition_parts
                sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
                composition_parts = []
                for i, (breed_name, _) in enumerate(sorted_breeds):
                    if i < len(coeffs_only):
                        breed_letter = self.get_breed_letter(breed_name)
                        composition_parts.append(f"{coeffs_only[i]}{breed_letter}")

                composition_text = ''.join(composition_parts) + "Др"
                composition_result = Label(
                    text=f"Формула состава: {composition_text}",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0, 0, 0, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(composition_result)
            else:
                no_composition = Label(
                    text="Коэффициент состава не определен (недостаточно данных)",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(1, 0, 0, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_composition)

            # Хвойные породы с градациями высот
            coniferous_label = Label(
                text='\nХВОЙНЫЕ ПОРОДЫ - ВЫСОТА ПО ГРАДАЦИЯМ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(coniferous_label)

            has_coniferous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'coniferous' and data['plots']:
                    has_coniferous = True

                    # Средняя густота в градациях
                    zones = data.get('coniferous_zones', {})
                    avg_do_05 = zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0
                    avg_05_15 = zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0
                    avg_bolee_15 = zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0

                    # Средняя высота только по значениям >1.5м
                    avg_heights_over_15 = [p['height'] for p in data['plots'] if p['height'] > 1.5]
                    avg_height_total = sum(avg_heights_over_15) / len(avg_heights_over_15) if avg_heights_over_15 else 0

                    # Средний диаметр
                    avg_diameter = sum(data['diameters']) / len(data['diameters']) if data['diameters'] else 0

                    coniferous_result = Label(
                        text=f"{breed_name}:\n"
                             f"• до 0.5м: {avg_do_05:.1f} шт/га\n"
                             f"• 0.5-1.5м: {avg_05_15:.1f} шт/га\n"
                             f"• >1.5м: {avg_bolee_15:.1f} шт/га\n"
                             f"• средняя высота (>1.5м): {avg_height_total:.1f}м\n"
                             f"• средний диаметр: {avg_diameter:.1f} см",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0, 0.5, 0, 1),
                        size_hint=(1, None),
                        height=120,
                        halign='left',
                        valign='top'
                    )
                    coniferous_result.bind(size=lambda *args: setattr(coniferous_result, 'text_size', (coniferous_result.width, None)))
                    results_layout.add_widget(coniferous_result)

            if not has_coniferous:
                no_coniferous = Label(
                    text="Хвойные породы не найдены",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_coniferous)

            # Лиственные породы - средние высоты и возраст
            deciduous_label = Label(
                text='\nЛИСТВЕННЫЕ ПОРОДЫ - СРЕДНИЕ ПОКАЗАТЕЛИ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(deciduous_label)

            has_deciduous = False
            for breed_name, data in sorted(breeds_data.items()):
                if data['type'] == 'deciduous' and data['plots']:
                    has_deciduous = True

                    avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])

                    avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                    avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0

                    avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                    avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                    # Средний диаметр
                    avg_diameter = sum(data['diameters']) / len(data['diameters']) if data['diameters'] else 0

                    deciduous_result = Label(
                        text=f"{breed_name}:\n"
                             f"• Густота: {avg_density:.1f} шт/га\n"
                             f"• Средняя высота: {avg_height:.1f}м\n"
                             f"• Средний возраст: {avg_age:.1f} лет\n"
                             f"• Средний диаметр: {avg_diameter:.1f} см",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0, 0.3, 0.5, 1),
                        size_hint=(1, None),
                        height=100,
                        halign='left',
                        valign='top'
                    )
                    deciduous_result.bind(size=lambda *args: setattr(deciduous_result, 'text_size', (deciduous_result.width, None)))
                    results_layout.add_widget(deciduous_result)

            if not has_deciduous:
                no_deciduous = Label(
                    text="Лиственные породы не найдены",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_deciduous)

            # Средние данные в целом по участку
            overall_label = Label(
                text='\nСРЕДНИЕ ДАННЫЕ В ЦЕЛОМ ПО УЧАСТКУ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(overall_label)

            # Рассчитываем общие средние
            all_densities = []
            all_heights = []
            all_ages = []
            all_diameters = []

            for breed_name, data in breeds_data.items():
                if data['plots']:
                    all_densities.extend([p['density'] for p in data['plots'] if p['density'] > 0])
                    all_heights.extend([p['height'] for p in data['plots'] if p['height'] > 0])
                    all_ages.extend([p['age'] for p in data['plots'] if p['age'] > 0])
                    all_diameters.extend([d for d in data['diameters'] if d > 0])

            avg_overall_density = sum(all_densities) / len(all_densities) if all_densities else 0
            avg_overall_height = sum(all_heights) / len(all_heights) if all_heights else 0
            avg_overall_age = sum(all_ages) / len(all_ages) if all_ages else 0
            avg_overall_diameter = sum(all_diameters) / len(all_diameters) if all_diameters else 0

            overall_result = Label(
                text=f"Средняя густота: {avg_overall_density:.1f} шт/га\n"
                     f"Средняя высота: {avg_overall_height:.1f} м\n"
                     f"Средний возраст: {avg_overall_age:.1f} лет\n"
                     f"Средний диаметр: {avg_overall_diameter:.1f} см",
                font_name='Roboto',
                font_size='14sp',
                color=(0.8, 0.4, 0, 1),
                size_hint=(1, None),
                height=100,
                halign='left',
                valign='top'
            )
            overall_result.bind(size=lambda *args: setattr(overall_result, 'text_size', (overall_result.width, None)))
            results_layout.add_widget(overall_result)



            # Предмет ухода и интенсивность рубки
            care_label = Label(
                text='\nПРЕДМЕТ УХОДА И ИНТЕНСИВНОСТЬ РУБКИ',
                font_name='Roboto',
                font_size='16sp',
                bold=True,
                color=(0, 0, 0, 1),
                size_hint=(1, None),
                height=50,
                halign='center'
            )
            results_layout.add_widget(care_label)

            # Собираем данные по предмету ухода и рассчитываем интенсивность
            care_data = []
            total_density_all_plots = 0  # Для площадок с предметом ухода
            total_remaining_density = 0
            plot_count_with_care = 0

            # Рассчитываем среднюю густоту по всем площадкам
            total_density_all = 0
            plot_count_all = 0

            for page_num, page_rows in self.page_data.items():
                for row in page_rows:
                    if len(row) >= 4 and row[3]:  # Есть данные пород
                        plot_density = 0
                        breeds_text = row[3]
                        if breeds_text:
                            try:
                                breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                for breed_info in breeds_list:
                                    if isinstance(breed_info, dict):
                                        if breed_info.get('type') == 'coniferous':
                                            # Для хвойных суммируем градации
                                            do_05 = breed_info.get('do_05', 0)
                                            _05_15 = breed_info.get('05_15', 0)
                                            bolee_15 = breed_info.get('bolee_15', 0)
                                            plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                        else:
                                            # Для лиственных обычная густота
                                            density = breed_info.get('density', 0)
                                            plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                            except (json.JSONDecodeError, TypeError):
                                pass

                        if plot_density > 0:
                            total_density_all += plot_density
                            plot_count_all += 1

                    if len(row) >= 4 and row[2]:  # Предмет ухода и данные пород
                        care_text = row[2].strip()
                        if care_text:
                            # Рассчитываем общую густоту на этой площадке
                            plot_density = 0
                            breeds_text = row[3]
                            if breeds_text:
                                try:
                                    breeds_list = json.loads(breeds_text) if isinstance(breeds_text, str) else []
                                    for breed_info in breeds_list:
                                        if isinstance(breed_info, dict):
                                            if breed_info.get('type') == 'coniferous':
                                                # Для хвойных суммируем градации
                                                do_05 = breed_info.get('do_05', 0)
                                                _05_15 = breed_info.get('05_15', 0)
                                                bolee_15 = breed_info.get('bolee_15', 0)
                                                plot_density += (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                                            else:
                                                # Для лиственных обычная густота
                                                density = breed_info.get('density', 0)
                                                plot_density += density / plot_area_ha if plot_area_ha > 0 else 0
                                except (json.JSONDecodeError, TypeError):
                                    pass

                            if plot_density > 0:
                                # Парсим предмет ухода для получения оставляемой густоты
                                remaining_density = self.parse_care_subject_density(care_text)
                                if remaining_density > 0:
                                    care_data.append({
                                        'care_text': care_text,
                                        'plot_density': plot_density,
                                        'remaining_density': remaining_density
                                    })
                                    total_density_all_plots += plot_density
                                    total_remaining_density += remaining_density
                                    plot_count_with_care += 1

            if care_data:
                # Рассчитываем средний предмет ухода
                care_breed_totals = {}
                care_plot_count = 0

                for item in care_data:
                    care_text = item['care_text']
                    breed_densities = self.parse_care_subject_by_breeds(care_text)
                    for breed, density in breed_densities.items():
                        if breed not in care_breed_totals:
                            care_breed_totals[breed] = 0
                        care_breed_totals[breed] += density
                    care_plot_count += 1

                if care_breed_totals and care_plot_count > 0:
                    avg_care_parts = []
                    short_parts = []
                    for breed, total_density in sorted(care_breed_totals.items()):
                        avg_density = total_density / care_plot_count
                        avg_care_parts.append(f"{avg_density * 1000:.0f}шт/га{breed}")
                        short_parts.append(f"{avg_density:.1f}{breed}")
                    avg_care_text = ''.join(avg_care_parts)
                    short_text = ''.join(short_parts).replace('.', ',')

                    care_result = Label(
                        text=f"Средний предмет ухода: {avg_care_text} = {short_text}",
                        font_name='Roboto',
                        font_size='14sp',
                        color=(0.2, 0.6, 0.2, 1),
                        size_hint=(1, None),
                        height=40,
                        halign='left',
                        valign='top'
                    )
                    care_result.bind(size=lambda *args: setattr(care_result, 'text_size', (care_result.width, None)))
                    results_layout.add_widget(care_result)

                    # Рассчитываем среднюю интенсивность рубки
                if plot_count_with_care > 0 and plot_count_all > 0:
                    avg_remaining_density = total_remaining_density / plot_count_with_care

                    # Интенсивность рубки = ((общая густота - оставляемая густота) / общая густота) * 100%
                    if avg_overall_density > 0:
                        intensity = ((avg_overall_density - avg_remaining_density) / avg_overall_density) * 100

                        intensity_result = Label(
                            text=f"Средняя интенсивность рубки: {intensity:.1f}%\n"
                                 f"(было {avg_overall_density:.0f} шт/га, "
                                 f"останется {avg_remaining_density:.0f} шт/га)",
                            font_name='Roboto',
                            font_size='14sp',
                            color=(0.8, 0.2, 0.2, 1),
                            size_hint=(1, None),
                            height=60,
                            halign='left',
                            valign='top'
                        )
                        intensity_result.bind(size=lambda *args: setattr(intensity_result, 'text_size', (intensity_result.width, None)))
                        results_layout.add_widget(intensity_result)
            else:
                no_care = Label(
                    text="Предмет ухода не указан или недостаточно данных для расчета",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0.5, 0.5, 0.5, 1),
                    size_hint=(1, None),
                    height=30,
                    halign='center'
                )
                results_layout.add_widget(no_care)

            # Информация о площади участка и площади перечета
            plot_count = len([row for page in self.page_data.values() for row in page if any(cell for cell in row[:3] if cell)])
            total_plot_area_ha = plot_count * plot_area_ha

            plot_area_info = Label(
                text=f"Информация о площади:\n"
                     f"• Радиус пробной площади: {default_radius:.2f} м\n"
                     f"• Площадь одной площадки: {plot_area_ha:.4f} га ({plot_area_m2:.4f} м²)\n"
                     f"• Всего площадок: {plot_count}\n"
                     f"• Совокупная площадь перечета: {total_plot_area_ha:.4f} га ({total_plot_area_ha*10000:.0f} м²)\n"
                     f"• Пример: 1 площадка = {plot_area_ha:.4f} га, значит площадь перечета = {plot_area_m2:.0f} м²",
                font_name='Roboto',
                font_size='12sp',
                color=(0.5, 0.5, 0.5, 1),
                size_hint=(1, None),
                height=120,
                halign='left',
                valign='top'
            )
            plot_area_info.bind(size=lambda *args: setattr(plot_area_info, 'text_size', (plot_area_info.width, None)))
            results_layout.add_widget(plot_area_info)

            scroll.add_widget(results_layout)
            content.add_widget(scroll)

            close_btn = ModernButton(
                text='Закрыть',
                bg_color=get_color_from_hex('#808080'),
                size_hint=(1, None),
                height=50
            )
            content.add_widget(close_btn)

            popup = Popup(
                title="Таксационные показатели молодняков",
                content=content,
                size_hint=(0.95, 0.95)
            )

            close_btn.bind(on_press=popup.dismiss)
            popup.open()

        except Exception as e:
            import traceback
            self.show_error(f"Ошибка расчета таксационных показателей: {str(e)}\n{traceback.format_exc()}")

    def update_totals(self, update_global=True):
        """Обновление строки итогов с поддержкой множественных пород"""
        breed_composition = {}  # Initialize at top
        total_stats = {'density': [], 'height': [], 'age': []}  # Initialize at top

        # Calculate radius and area once
        current_radius = float(self.current_radius) if self.current_radius else 5.64
        plot_area_m2 = 3.14159 * (current_radius ** 2)
        if update_global:
            section_data = self.calculate_section_totals()
            composition_text = section_data['composition_text']
            forestry_formulas_text = section_data['forestry_formulas_text']
            total_plots = section_data['total_plots']
        else:
            # Старая логика по странице
            breed_composition = {}
            total_stats = {'density': [], 'height': [], 'age': []}
            coniferous_stats = {'do_05': [], '05_15': [], 'bolee_15': [], 'height': [], 'age': []}

            for row in self.inputs:
                predmet_text = row[2].text
                if predmet_text:
                    composition = self.parse_composition(predmet_text)
                    for breed, count in composition.items():
                        if breed not in breed_composition:
                            breed_composition[breed] = []
                        breed_composition[breed].append(count)

                breeds_text = row[3].text
                if breeds_text:
                    breeds_data = self.parse_breeds_data(breeds_text)
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'coniferous':
                            # Густота хвойных = сумма градаций
                            coniferous_density = (breed_info.get('do_05', 0) +
                                                breed_info.get('05_15', 0) +
                                                breed_info.get('bolee_15', 0))
                            if coniferous_density > 0:
                                total_stats['density'].append(coniferous_density)
                            else:
                                if 'density' in breed_info and breed_info['density']:
                                    total_stats['density'].append(breed_info['density'])

                            # Сбор данных по градациям для хвойных
                            if breed_info.get('do_05', 0) > 0:
                                coniferous_stats['do_05'].append(breed_info['do_05'])
                            if breed_info.get('05_15', 0) > 0:
                                coniferous_stats['05_15'].append(breed_info['05_15'])
                            if breed_info.get('bolee_15', 0) > 0:
                                coniferous_stats['bolee_15'].append(breed_info['bolee_15'])
                            if breed_info.get('height', 0) > 0:
                                coniferous_stats['height'].append(breed_info['height'])
                            if breed_info.get('age', 0) > 0:
                                coniferous_stats['age'].append(breed_info['age'])
                        elif 'density' in breed_info and breed_info['density']:
                            total_stats['density'].append(breed_info['density'])

                        if 'height' in breed_info and breed_info['height']:
                            total_stats['height'].append(breed_info['height'])
                        if 'age' in breed_info and breed_info['age']:
                            total_stats['age'].append(breed_info['age'])

            # Рассчитываем итоги по странице

            current_radius = float(self.current_radius) if self.current_radius else 1.78
            plot_area_m2 = 3.14159 * (current_radius ** 2)  # Площадь пробной площади в м²

            # Расчет средних по градациям для хвойных по формулам лесного хозяйства на гектар
            coniferous_stats_ha = []
            for row in range(len(self.inputs)):
                row_do_05 = coniferous_stats['do_05'][row] if row < len(coniferous_stats['do_05']) and row < len(coniferous_stats['do_05']) else 0
                row_05_15 = coniferous_stats['05_15'][row] if row < len(coniferous_stats['05_15']) else 0
                row_bolee_15 = coniferous_stats['bolee_15'][row] if row < len(coniferous_stats['bolee_15']) else 0
                row_height = coniferous_stats['height'][row] if row < len(coniferous_stats['height']) else 0
                row_age = coniferous_stats['age'][row] if row < len(coniferous_stats['age']) else 0

                # Рассчитываем густоту на гектар для градаций
                do_05_ha = (row_do_05 * 10000) / plot_area_m2 if plot_area_m2 > 0 else 0
                _05_15_ha = (row_05_15 * 10000) / plot_area_m2 if plot_area_m2 > 0 else 0
                bolee_15_ha = (row_bolee_15 * 10000) / plot_area_m2 if plot_area_m2 > 0 else 0

                coniferous_stats_ha.append({
                    'do_05_ha': do_05_ha,
                    '05_15_ha': _05_15_ha,
                    'bolee_15_ha': bolee_15_ha,
                    'height': row_height,
                    'age': row_age
                })

            avg_coniferous_do_05_ha = sum(d['do_05_ha'] for d in coniferous_stats_ha) / len(coniferous_stats_ha) if coniferous_stats_ha else 0
            avg_coniferous_05_15_ha = sum(d['05_15_ha'] for d in coniferous_stats_ha) / len(coniferous_stats_ha) if coniferous_stats_ha else 0
            avg_coniferous_bolee_15_ha = sum(d['bolee_15_ha'] for d in coniferous_stats_ha) / len(coniferous_stats_ha) if coniferous_stats_ha else 0
            avg_coniferous_height_ha = sum(d['height'] for d in coniferous_stats_ha) / len(coniferous_stats_ha) if coniferous_stats_ha else 0
            avg_coniferous_age_ha = sum(d['age'] for d in coniferous_stats_ha) / len(coniferous_stats_ha) if coniferous_stats_ha else 0

            # Формирование текста для столбца Порода в строке итогов с формулами лесного хозяйства
            forestry_formulas_text = ""

            # Хвойные породы - средние значения по градациям на га
            if coniferous_stats_ha:
                forestry_formulas_text += "Хвойные: "
                gradiations = []
                if avg_coniferous_do_05_ha > 0:
                    gradiations.append(f"до 0.5м: {avg_coniferous_do_05_ha:.1f} шт/га")
                if avg_coniferous_05_15_ha > 0:
                    gradiations.append(f"0.5-1.5м: {avg_coniferous_05_15_ha:.1f} шт/га")
                if avg_coniferous_bolee_15_ha > 0:
                    gradiations.append(f">1.5м: {avg_coniferous_bolee_15_ha:.1f} шт/га")
                if gradiations:
                    forestry_formulas_text += ", ".join(gradiations)
                if avg_coniferous_height_ha > 0:
                    forestry_formulas_text += f", высота: {avg_coniferous_height_ha:.1f}м"
                if avg_coniferous_age_ha > 0:
                    forestry_formulas_text += f", возраст: {avg_coniferous_age_ha:.1f} лет"

            # Лиственные породы - средние значения без градаций
            deciduous_density = []
            deciduous_height = []
            deciduous_age = []

            for row in self.inputs:
                breeds_text = row[3].text
                if breeds_text:
                    breeds_data = self.parse_breeds_data(breeds_text)
                    for breed_info in breeds_data:
                        if breed_info.get('type') == 'deciduous':
                            if breed_info.get('density'):
                                deciduous_density.append(breed_info['density'] * (10000 / plot_area_m2) if plot_area_m2 > 0 else breed_info['density'])
                            if breed_info.get('height'):
                                deciduous_height.append(breed_info['height'])
                            if breed_info.get('age'):
                                deciduous_age.append(breed_info['age'])

            # Рассчитываем средние по лиственным на га
            if deciduous_density or deciduous_height or deciduous_age:
                avg_deciduous_density = sum(deciduous_density) / len(deciduous_density) if deciduous_density else 0
                avg_deciduous_height = sum(deciduous_height) / len(deciduous_height) if deciduous_height else 0
                avg_deciduous_age = sum(deciduous_age) / len(deciduous_age) if deciduous_age else 0

                if forestry_formulas_text:
                    forestry_formulas_text += "; "
                forestry_formulas_text += "Лиственные: "
                parts = []
                if avg_deciduous_density > 0:
                    parts.append(f"густота: {avg_deciduous_density:.1f} шт/га")
                if avg_deciduous_height > 0:
                    parts.append(f"высота: {avg_deciduous_height:.1f}м")
                if avg_deciduous_age > 0:
                    parts.append(f"возраст: {avg_deciduous_age:.1f} лет")
                forestry_formulas_text += ", ".join(parts)

        avg_composition = {}
        for breed, counts in breed_composition.items():
            if counts:
                avg_composition[breed] = sum(counts) / len(counts)

        composition_text = ""
        for breed in sorted(avg_composition.keys()):
            count = avg_composition[breed]
            if count > 0:
                composition_text += f"{int(count)}{breed}"


    def parse_composition(self, text):
        """Парсит текстовое представление состава пород"""
        composition = {}
        if isinstance(text, str):
            matches = re.findall(r'(\d+)([А-ЯA-Z])', text.upper())
            for count, breed in matches:
                try:
                    composition[breed] = int(count)
                except ValueError:
                    pass
        return composition

    def parse_breeds_data(self, breeds_text):
        """Парсит данные пород из текстового поля"""
        if not breeds_text or not isinstance(breeds_text, str):
            return []

        try:
            if isinstance(breeds_text, str) and breeds_text.startswith('['):
                return json.loads(breeds_text)
            elif isinstance(breeds_text, str) and breeds_text.startswith('{'):
                return [json.loads(breeds_text)]
        except (json.JSONDecodeError, TypeError):
            pass

        return []

    def calculate_page_totals(self):
        """Вычисляет итоговые значения для текущей страницы"""
        totals = {
            'composition': '',
            'total_area': 0.0,
            'avg_age': 0.0,
            'avg_density': 0.0,
            'avg_height': 0.0
        }

        breed_composition = {}
        total_stats = {'density': [], 'height': [], 'age': []}
        total_area = 0.0

        for row in self.inputs:
            predmet_text = row[2].text
            if predmet_text:
                composition = self.parse_composition(predmet_text)
                for breed, count in composition.items():
                    if breed not in breed_composition:
                        breed_composition[breed] = []
                    breed_composition[breed].append(count)

            radius = 5.64
            try:
                if row[5].text:
                    radius = float(row[5].text)
            except (ValueError, IndexError):
                pass

            area = 3.14159 * (radius ** 2)
            total_area += area

            breeds_text = row[3].text
            if breeds_text:
                breeds_data = self.parse_breeds_data(breeds_text)
                for breed_info in breeds_data:
                    if breed_info.get('type') == 'coniferous':
                        coniferous_density = (breed_info.get('do_05', 0) +
                                            breed_info.get('05_15', 0) +
                                            breed_info.get('bolee_15', 0))
                        if coniferous_density > 0:
                            total_stats['density'].append(coniferous_density)
                    elif 'density' in breed_info and breed_info['density']:
                        total_stats['density'].append(breed_info['density'])

                    if 'height' in breed_info and breed_info['height']:
                        total_stats['height'].append(breed_info['height'])
                    if 'age' in breed_info and breed_info['age']:
                        total_stats['age'].append(breed_info['age'])

        avg_composition = {}
        for breed, counts in breed_composition.items():
            if counts:
                avg_composition[breed] = sum(counts) / len(counts)

        composition_text = ""
        for breed in sorted(avg_composition.keys()):
            count = avg_composition[breed]
            if count > 0:
                composition_text += f"{int(count)}{breed}"

        totals['composition'] = composition_text
        totals['total_area'] = total_area
        totals['avg_density'] = sum(total_stats['density']) / len(total_stats['density']) if total_stats['density'] else 0.0
        totals['avg_height'] = sum(total_stats['height']) / len(total_stats['height']) if total_stats['height'] else 0.0
        totals['avg_age'] = sum(total_stats['age']) / len(total_stats['age']) if total_stats['age'] else 0.0

        return totals

    def save_current_page(self, instance=None):
        """Сохраняем текущую страницу в базу данных с поддержкой множественных пород"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        try:
            cursor.execute('''
                DELETE FROM molodniki_data
                WHERE page_number = ? AND section_name = ?
            ''', (self.current_page, self.current_section))

            for row_idx, row in enumerate(self.inputs):
                row_data = [inp.text.strip() for inp in row]
                if any(row_data[:5]):
                    radius = 5.64
                    try:
                        if row_data[5]:
                            radius = float(row_data[5])
                    except (ValueError, IndexError):
                        pass

                    cursor.execute('''
                        INSERT INTO molodniki_data
                        (page_number, row_index, nn, gps_point, predmet_uhoda, primechanie, section_name)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        self.current_page,
                        row_idx,
                        row_data[0] or None,
                        row_data[1] or None,
                        row_data[2] or None,
                        row_data[4] or None,
                        self.current_section
                    ))

                    molodniki_data_id = cursor.lastrowid

                    if row_data[3]:
                        breeds_data = self.parse_breeds_data(row_data[3])
                        for breed_info in breeds_data:
                            try:
                                # Validate and convert data types
                                density = int(breed_info.get('density', 0) or 0)
                                height = float(breed_info.get('height', 0.0) or 0.0)
                                age = int(breed_info.get('age', 0) or 0)
                                do_05 = int(breed_info.get('do_05', 0) or 0)
                                _05_15 = int(breed_info.get('05_15', 0) or 0)
                                bolee_15 = int(breed_info.get('bolee_15', 0) or 0)

                                composition_coeff = 0.0
                                if density and radius:
                                    area = 3.14159 * (radius ** 2)
                                    composition_coeff = (density * area) / 10000

                                diameter = float(breed_info.get('diameter', 0.0) or 0.0)

                                cursor.execute('''
                                    INSERT INTO molodniki_breeds
                                    (molodniki_data_id, breed_name, breed_type, do_05, _05_15, bolee_15,
                                     density, height, diameter, age, composition_coefficient)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (
                                    molodniki_data_id,
                                    breed_info.get('name', ''),
                                    breed_info.get('type', 'deciduous'),
                                    do_05,
                                    _05_15,
                                    bolee_15,
                                    density,
                                    height,
                                    diameter,
                                    age,
                                    composition_coeff
                                ))
                            except Exception as e:
                                print(f"Error inserting breed: {e}, skipping this breed")
                                continue

            totals = self.calculate_page_totals()
            cursor.execute('''
                INSERT OR REPLACE INTO molodniki_totals
                (page_number, section_name, total_composition, avg_age, avg_density, avg_height)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                self.current_page,
                self.current_section,
                totals['composition'],
                totals['avg_age'],
                totals['avg_density'],
                totals['avg_height']
            ))

            conn.commit()
            self.show_success("Страница сохранена в базу данных!")
            success = True

        except Exception as e:
            conn.rollback()
            self.show_error(f"Ошибка сохранения: {str(e)}")
            success = False
        finally:
            conn.close()

        page_data = []
        for row in self.inputs:
            page_data.append([inp.text for inp in row])
        self.page_data[self.current_page] = page_data

        return success

    def show_save_dialog(self, instance=None):
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Введите имя файла",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30
        )
        content.add_widget(title_label)

        self.filename_input = TextInput(
            hint_text="Имя файла",
            multiline=False,
            size_hint=(1, None),
            height=40,
            font_name='Roboto'
        )
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        default_name = f"Молодняки_{self.current_section}_{timestamp}"
        self.filename_input.text = default_name
        content.add_widget(self.filename_input)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        self.save_popup = Popup(
            title="Сохранение отчета Excel",
            content=content,
            size_hint=(0.7, 0.5)
        )
        save_btn.bind(on_press=self.save_to_excel)
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
            wb = Workbook()
            ws = wb.active
            ws.title = "Молодняки"

            address_parts = []
            if self.current_quarter:
                address_parts.append(f"Квартал: {self.current_quarter}")
            if self.current_plot:
                address_parts.append(f"Выдел: {self.current_plot}")
            if self.current_forestry:
                address_parts.append(f"Лесничество: {self.current_forestry}")
            if self.current_radius:
                address_parts.append(f"Радиус: {self.current_radius} м")

            address_text = " | ".join(address_parts) if address_parts else "Адрес не указан"
            ws['A1'] = f"Адрес: {address_text}"
            ws['A1'].font = openpyxl.styles.Font(bold=True, size=12)

            ws.append([])

            headers = [
                '№ППР', 'GPS точка', 'Предмет ухода', 'Порода', 'Густота', 'Высота', 'Возраст', 'Примечания', 'Тип Леса'
            ]
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_num, value=header)
                cell.font = openpyxl.styles.Font(bold=True)
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

            all_data = []
            for page in sorted(self.page_data.keys()):
                all_data.extend(self.page_data[page])

            current_row = 4
            for row in all_data:
                if any(cell for cell in row[:3] if cell):  # Проверяем, что основные столбцы не пустые
                    try:
                        breeds_data = json.loads(row[3]) if row[3] else []
                    except (json.JSONDecodeError, TypeError):
                        breeds_data = []

                    if isinstance(breeds_data, list) and breeds_data:
                        for breed_info in breeds_data:
                            if isinstance(breed_info, dict):
                                breed_name = breed_info.get('name', 'Неизвестная')
                                density = breed_info.get('density', '')
                                height = breed_info.get('height', '')
                                age = breed_info.get('age', '')

                                # Для хвойных рассчитываем густоту по градациям
                                if breed_info.get('type') == 'coniferous':
                                    coniferous_density = (breed_info.get('do_05', 0) +
                                                        breed_info.get('05_15', 0) +
                                                        breed_info.get('bolee_15', 0))
                                    if coniferous_density > 0:
                                        density = str(coniferous_density)

                                processed_row = [
                                    row[0],  # №ППР
                                    row[1],  # GPS точка
                                    row[2],  # Предмет ухода
                                    breed_name,  # Порода
                                    str(density) if density else '',  # Густота
                                    str(height) if height else '',  # Высота
                                    str(age) if age else '',  # Возраст
                                    row[4],  # Примечания
                                    row[5],  # Тип Леса
                                ]
                                ws.append(processed_row)
                                current_row += 1
                    else:
                        # Если нет пород, добавить строку без данных
                        processed_row = [row[0], row[1], row[2], '', '', '', '', row[4], row[5]]
                        ws.append(processed_row)
                        current_row += 1

            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(full_path)
            self.save_popup.dismiss()
            self.show_success(f"Файл сохранен: {filename}")
        except Exception as e:
            self.show_error(f"Ошибка: {str(e)}")

    def save_to_word(self, instance):
        try:
            from docx import Document

            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
            filename = f"Молодняки_{self.current_section}_{timestamp}.docx"
            full_path = os.path.join(self.reports_dir, filename)

            doc = Document()
            doc.add_heading(f'Расширенный отчет по молоднякам - Участок {self.current_section}', 0)

            all_data = []
            for page in sorted(self.page_data.keys()):
                all_data.extend(self.page_data[page])

            table = doc.add_table(rows=1, cols=9)
            table.style = 'Table Grid'

            headers = [
                '№ППР', 'GPS точка', 'Предмет ухода', 'Порода', 'Густота', 'Высота', 'Возраст', 'Примечания', 'Тип Леса'
            ]
            hdr_cells = table.rows[0].cells
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            for row in all_data:
                if any(cell for cell in row[:3] if cell):  # Проверяем, что основные столбцы не пустые
                    try:
                        breeds_data = json.loads(row[3]) if row[3] else []
                    except (json.JSONDecodeError, TypeError):
                        breeds_data = []

                    if isinstance(breeds_data, list) and breeds_data:
                        for breed_info in breeds_data:
                            if isinstance(breed_info, dict):
                                breed_name = breed_info.get('name', 'Неизвестная')
                                density = breed_info.get('density', '')
                                height = breed_info.get('height', '')
                                age = breed_info.get('age', '')

                                # Для хвойных рассчитываем густоту по градациям
                                if breed_info.get('type') == 'coniferous':
                                    coniferous_density = (breed_info.get('do_05', 0) +
                                                        breed_info.get('05_15', 0) +
                                                        breed_info.get('bolee_15', 0))
                                    if coniferous_density > 0:
                                        density = str(coniferous_density)

                                row_cells = table.add_row().cells
                                row_cells[0].text = str(row[0]) if row[0] else ""  # №ППР
                                row_cells[1].text = str(row[1]) if row[1] else ""  # GPS точка
                                row_cells[2].text = str(row[2]) if row[2] else ""  # Предмет ухода
                                row_cells[3].text = breed_name  # Порода
                                row_cells[4].text = str(density) if density else ""  # Густота
                                row_cells[5].text = str(height) if height else ""  # Высота
                                row_cells[6].text = str(age) if age else ""  # Возраст
                                row_cells[7].text = str(row[4]) if row[4] else ""  # Примечания
                                row_cells[8].text = str(row[5]) if row[5] else ""  # Тип Леса
                    else:
                        # Если нет пород, добавить строку без данных
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(row[0]) if row[0] else ""
                        row_cells[1].text = str(row[1]) if row[1] else ""
                        row_cells[2].text = str(row[2]) if row[2] else ""
                        row_cells[3].text = ""
                        row_cells[4].text = ""
                        row_cells[5].text = ""
                        row_cells[6].text = ""
                        row_cells[7].text = str(row[4]) if row[4] else ""
                        row_cells[8].text = str(row[5]) if row[5] else ""

            doc.save(full_path)
            self.show_success(f"Word документ сохранен: {filename}")
        except ImportError:
            self.show_error("Для сохранения в Word установите библиотеку python-docx: pip install python-docx")
        except Exception as e:
            self.show_error(f"Ошибка сохранения Word: {str(e)}")

    def aggregate_breeds_data(self, df):
        """Агрегирует данные пород по площадкам из Excel файла с учетом заголовков"""
        # Получаем заголовки из первой строки
        headers = df.iloc[0] if not df.empty else []

        # Находим индексы нужных столбцов
        nn_idx = None
        gps_idx = None
        predmet_idx = None
        breed_name_idx = None
        density_idx = None
        do_05_idx = None
        _05_15_idx = None
        bolee_15_idx = None
        height_idx = None
        age_idx = None
        primechanie_idx = None
        tip_lesa_idx = None

        for i, header in enumerate(headers):
            header_str = str(header).strip().lower()
            if '№ппр' in header_str:
                nn_idx = i
            elif 'gps' in header_str:
                gps_idx = i
            elif 'предмет ухода' in header_str:
                predmet_idx = i
            elif 'порода' in header_str:
                breed_name_idx = i
            elif 'густота' in header_str:
                density_idx = i
            elif 'до 0.5м' in header_str:
                do_05_idx = i
            elif '0.5-1.5м' in header_str:
                _05_15_idx = i
            elif '>1.5м' in header_str or 'выше' in header_str:
                bolee_15_idx = i
            elif 'высота' in header_str:
                height_idx = i
            elif 'возраст' in header_str:
                age_idx = i
            elif 'примечания' in header_str:
                primechanie_idx = i
            elif 'тип леса' in header_str:
                tip_lesa_idx = i

        # Группировка по данным площадки (GPS, Предмет ухода, Примечания, Тип Леса)
        grouped = {}

        print(f"DEBUG: aggregate_breeds_data starting, df shape: {df.shape}")

        # Начинаем с второй строки (после заголовков)
        for index in range(1, len(df)):
            row = df.iloc[index]

            # Извлекаем данные по найденным индексам
            gps = str(row.iloc[gps_idx]) if gps_idx is not None and gps_idx < len(row) else ''
            predmet = str(row.iloc[predmet_idx]) if predmet_idx is not None and predmet_idx < len(row) else ''
            breed_name = str(row.iloc[breed_name_idx]) if breed_name_idx is not None and breed_name_idx < len(row) else ''
            density = str(row.iloc[density_idx]) if density_idx is not None and density_idx < len(row) else ''
            do_05 = str(row.iloc[do_05_idx]) if do_05_idx is not None and do_05_idx < len(row) else ''
            _05_15 = str(row.iloc[_05_15_idx]) if _05_15_idx is not None and _05_15_idx < len(row) else ''
            bolee_15 = str(row.iloc[bolee_15_idx]) if bolee_15_idx is not None and bolee_15_idx < len(row) else ''
            height = str(row.iloc[height_idx]) if height_idx is not None and height_idx < len(row) else ''
            age = str(row.iloc[age_idx]) if age_idx is not None and age_idx < len(row) else ''
            primechanie = str(row.iloc[primechanie_idx]) if primechanie_idx is not None and primechanie_idx < len(row) else ''
            tip_lesa = str(row.iloc[tip_lesa_idx]) if tip_lesa_idx is not None and tip_lesa_idx < len(row) else ''

            print(f"DEBUG: Processing row {index}: breed_name='{breed_name}', primechanie='{primechanie}', tip_lesa='{tip_lesa}'")

            # Ключ группы по уникальной комбинации данных площадки
            key = (str(gps), str(predmet), str(primechanie), str(tip_lesa))

            if key not in grouped:
                grouped[key] = {
                    'gps': gps,
                    'predmet': predmet,
                    'primechanie': primechanie,
                    'tip_lesa': tip_lesa,
                    'breeds': []
                }

            # Пропускаем строки без породы
            if not breed_name.strip() or breed_name in ['nan', 'NaN', '']:
                continue

            # Определяем тип породы
            breed_type = self.determine_breed_type(breed_name)

            # Создаем объект породы
            breed_data = {
                'name': breed_name,
                'type': breed_type
            }

            # Добавляем параметры с проверкой типов
            if density and density not in ['nan', 'NaN', '']:
                try:
                    breed_data['density'] = int(float(density))
                except (ValueError, TypeError):
                    pass

            if do_05 and do_05 not in ['nan', 'NaN', '']:
                try:
                    breed_data['do_05'] = int(float(do_05))
                except (ValueError, TypeError):
                    pass

            if _05_15 and _05_15 not in ['nan', 'NaN', '']:
                try:
                    breed_data['05_15'] = int(float(_05_15))
                except (ValueError, TypeError):
                    pass

            if bolee_15 and bolee_15 not in ['nan', 'NaN', '']:
                try:
                    breed_data['bolee_15'] = int(float(bolee_15))
                except (ValueError, TypeError):
                    pass

            if height and height not in ['nan', 'NaN', '']:
                try:
                    breed_data['height'] = float(height)
                except (ValueError, TypeError):
                    pass

            if age and age not in ['nan', 'NaN', '']:
                try:
                    breed_data['age'] = int(float(age))
                except (ValueError, TypeError):
                    pass

            # Если порода уже есть в списке, добавляем/обновляем параметры
            existing_breed = None
            for b in grouped[key]['breeds']:
                if b['name'] == breed_name:
                    existing_breed = b
                    break

            if existing_breed:
                # Обновляем существующую породу
                for k, v in breed_data.items():
                    if k not in existing_breed or not existing_breed.get(k):
                        existing_breed[k] = v
            else:
                grouped[key]['breeds'].append(breed_data)

        # Формируем финальный список данных с автоматической нумерацией площадок
        result = []
        nn_counter = 1
        for group_key, group_data in grouped.items():
            # Создаем JSON строку для пород
            breeds_json = json.dumps(group_data['breeds'], ensure_ascii=False, indent=2) if group_data['breeds'] else ''

            result.append([
                str(nn_counter),  # Автоматическая нумерация от 1
                group_data['gps'],
                group_data['predmet'],
                breeds_json,  # Данные пород в JSON формате
                group_data['primechanie'],
                group_data['tip_lesa']
            ])
            nn_counter += 1

        return result

    def determine_breed_type(self, breed_name):
        """Определяет тип породы по названию"""
        coniferous_breeds = ['Сосна', 'Ель', 'Пихта', 'Кедр', 'Лиственница']
        deciduous_breeds = ['Берёза', 'Осина', 'Ольха чёрная', 'Ольха серая', 'Ива', 'Ива кустарниковая']

        if any(coniferous.lower() in breed_name.lower() for coniferous in coniferous_breeds):
            return 'coniferous'
        elif any(deciduous.lower() in breed_name.lower() for deciduous in deciduous_breeds):
            return 'deciduous'
        else:
            # По умолчанию считаем лиственными
            return 'deciduous'

    def load_section(self, instance):
        """Показать popup для выбора JSON файла"""
        if not os.path.exists(self.reports_dir):
            self.show_error("Папка reports не найдена!")
            return

        content = FloatLayout()

        title_label = Label(
            text='Выберите файл JSON данных приложения:',
            font_name='Roboto',
            font_size='18sp',
            color=(1, 1, 1, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.95},
            size_hint=(None, None),
            size=(400, 50)
        )

        scroll = ScrollView(size_hint=(0.9, 0.75), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        files_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        files_layout.bind(minimum_height=files_layout.setter('height'))

        # Добавляем кнопку для ручного ввода пути
        manual_input_layout = BoxLayout(orientation='vertical', size_hint_y=None, height=80, spacing=5)
        manual_label = Label(
            text="Или введите полный путь к файлу:",
            font_name='Roboto',
            size_hint=(1, None),
            height=30,
            color=(1, 1, 1, 1)
        )
        self.manual_file_input = TextInput(
            hint_text="Полный путь к JSON файлу",
            multiline=False,
            size_hint=(1, None),
            height=40,
            background_color=(1, 1, 1, 0.8),
            font_name='Roboto'
        )
        manual_input_layout.add_widget(manual_label)
        manual_input_layout.add_widget(self.manual_file_input)
        files_layout.add_widget(manual_input_layout)

        uploader_label = Label(
            text="Доступные JSON файлы:",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=30,
            color=(1, 1, 1, 1)
        )
        files_layout.add_widget(uploader_label)

        # Получаем список JSON файлов
        json_files = [f for f in os.listdir(self.reports_dir) if f.endswith('.json')]
        if not json_files:
            no_files_label = Label(
                text="JSON файлы не найдены в папке reports\nИспользуйте ручной ввод пути выше",
                font_name='Roboto',
                size_hint=(1, None),
                height=50,
                color=(0.8, 0.8, 0.8, 1),
                valign='top'
            )
            no_files_label.bind(size=lambda *args: setattr(no_files_label, 'text_size', (no_files_label.width, None)))
            files_layout.add_widget(no_files_label)
        else:
            for filename in sorted(json_files):
                btn = ModernButton(
                    text=filename,
                    bg_color=get_color_from_hex('#87CEEB'),
                    size_hint=(1, None),
                    height=50,
                    font_size='14sp'
                )
                btn.bind(on_press=lambda x, f=filename: self.select_json_file(os.path.join(self.reports_dir, f)))
                files_layout.add_widget(btn)

        scroll.add_widget(files_layout)

        btn_layout = BoxLayout(
            orientation='horizontal',
            spacing=10,
            size_hint=(1, None),
            height=60,
            pos_hint={'center_x': 0.5, 'center_y': 0.08}
        )
        load_manual_btn = ModernButton(
            text='Загрузить указанный файл',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.35, 1),
            height=60
        )
        load_manual_btn.bind(on_press=self.load_manual_json)
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.35, 1),
            height=60
        )
        cancel_btn.bind(on_press=self.dismiss_json_popup)
        btn_layout.add_widget(load_manual_btn)
        btn_layout.add_widget(cancel_btn)

        content.add_widget(title_label)
        content.add_widget(scroll)
        content.add_widget(btn_layout)

        self.json_popup = Popup(
            title='Загрузка данных приложения',
            content=content,
            size_hint=(0.9, 0.9),
            separator_height=0,
            background_color=(0.5, 0.5, 0.5, 1),
            overlay_color=(0, 0, 0, 0.5)
        )
        self.json_popup.open()

    def load_section_popup(self):
        """Показать popup для выбора JSON файла (вызывается из главного меню)"""
        return self.load_section(None)

    def select_json_file(self, file_path):
        """Обработка выбора JSON файла из списка"""
        try:
            self.load_json_data(file_path)
            self.json_popup.dismiss()
        except Exception as e:
            self.show_error(f"Ошибка загрузки: {str(e)}")

    def load_manual_json(self, instance):
        """Загрузка JSON файла по указанному пути"""
        file_path = self.manual_file_input.text.strip()
        if not file_path:
            self.show_error("Укажите путь к файлу!")
            return

        if not os.path.exists(file_path):
            self.show_error("Файл не найден!")
            return

        try:
            self.load_json_data(file_path)
            self.json_popup.dismiss()
        except Exception as e:
            self.show_error(f"Ошибка загрузки: {str(e)}")

    def dismiss_json_popup(self, instance=None):
        """Закрыть popup выбора файла"""
        if hasattr(self, 'json_popup'):
            self.json_popup.dismiss()

    def load_json_data(self, file_path):
        """Загрузка данных из JSON файла"""
        print(f"DEBUG: Loading JSON file: {file_path}")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.current_section = os.path.splitext(os.path.basename(file_path))[0].replace('.json', '').replace('_приложение', '')
            self.update_section_label()
            self.page_data.clear()

            # Ожидаем, что JSON содержит page_data как словарь
            if isinstance(data, dict) and 'page_data' in data:
                self.page_data = data['page_data']
                print(f"DEBUG: Loaded page_data: {len(self.page_data)} pages")
            else:
                # Старый формат или простая структура
                self.page_data = data if isinstance(data, dict) else {}
                print(f"DEBUG: Loaded data as dict: {len(self.page_data) if isinstance(self.page_data, dict) else 'not dict'}")

            # Проверяем и исправляем формат страницы
            corrected_page_data = {}
            for page_key, page_rows in self.page_data.items():
                if isinstance(page_key, str):
                    try:
                        page_num = int(page_key)
                    except ValueError:
                        continue
                else:
                    page_num = page_key

                if isinstance(page_rows, list):
                    # Убеждаемся, что каждая строка - список из 6 элементов
                    corrected_rows = []
                    for row in page_rows:
                        if isinstance(row, list) and len(row) == 6:
                            corrected_rows.append(row)
                        elif isinstance(row, list):
                            # Дополняем до 6 элементов пустыми строками
                            corrected_row = row + [''] * (6 - len(row))
                            corrected_rows.append(corrected_row[:6])
                        else:
                            continue
                    corrected_page_data[page_num] = corrected_rows
                else:
                    continue

            self.page_data = corrected_page_data

            if self.page_data:
                self.current_page = min(self.page_data.keys())
            else:
                self.current_page = 0

            self.load_page_data()
            self.update_pagination()

            total_plots = sum(len(rows) for rows in self.page_data.values())
            self.show_success(f"Данные приложения успешно загружены! Найдено {total_plots} площадок в {len(self.page_data)} страницах.")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.show_error(f"Ошибка загрузки JSON файла: {str(e)}\n{error_details}")

    def show_radius_popup(self, instance):
        """Показать popup для установки радиуса"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text="Установка радиуса для расчета коэффициента состава",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=40
        )
        content.add_widget(title_label)

        self.radius_input = TextInput(
            hint_text="Радиус (метры)",
            multiline=False,
            size_hint=(1, None),
            height=50,
            font_name='Roboto',
            input_filter='float',
            text=self.current_radius
        )
        content.add_widget(self.radius_input)

        info_label = Label(
            text="Радиус используется для расчета площади круга:\n"
                 "Площадь = π × радиус²\n"
                 "Коэффициент состава = (густота × площадь) / 10000\n"
                 "Радиус применяется автоматически ко всем площадкам",
            font_name='Roboto',
            size_hint=(1, None),
            height=100,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(info_label)

        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint=(1, None), height=50)
        save_btn = ModernButton(
            text='Сохранить радиус',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1),
            height=50
        )
        cancel_btn = ModernButton(
            text='Отмена',
            bg_color=get_color_from_hex('#FF6347'),
            size_hint=(0.5, 1),
            height=50
        )
        btn_layout.add_widget(save_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Настройка радиуса",
            content=content,
            size_hint=(0.8, 0.7)
        )

        def apply_radius(btn):
            try:
                radius = float(self.radius_input.text.strip())
                if radius <= 0:
                    self.show_error("Радиус должен быть положительным числом!")
                    return

                self.current_radius = str(radius)
                self.update_totals()
                self.show_success(f"Радиус {radius} м сохранен для всех расчетов")
                popup.dismiss()

            except ValueError:
                self.show_error("Введите корректное числовое значение радиуса!")

        save_btn.bind(on_press=apply_radius)
        cancel_btn.bind(on_press=popup.dismiss)

        popup.open()

    def show_breed_choice_popup(self, instance, selected_breed):
        """Показать popup с выбором после добавления первой породы"""
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)

        title_label = Label(
            text=f"Порода '{selected_breed}' добавлена!\nВыберите действие:",
            font_name='Roboto',
            bold=True,
            size_hint=(1, None),
            height=60,
            color=(0, 0.5, 0, 1)
        )
        content.add_widget(title_label)

        # Информация о номере породы
        info_label = Label(
            text="Автоматически присвоен номер: 1 порода",
            font_name='Roboto',
            size_hint=(1, None),
            height=30,
            color=(0.3, 0.3, 0.3, 1)
        )
        content.add_widget(info_label)

        btn_layout = BoxLayout(orientation='vertical', spacing=10, size_hint=(1, None), height=120)
        add_more_btn = ModernButton(
            text='Добавить еще породу',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(1, None),
            height=50
        )
        save_exit_btn = ModernButton(
            text='Сохранить и выйти',
            bg_color=get_color_from_hex('#32CD32'),
            size_hint=(1, None),
            height=50
        )
        btn_layout.add_widget(add_more_btn)
        btn_layout.add_widget(save_exit_btn)
        content.add_widget(btn_layout)

        popup = Popup(
            title="Выбор действия",
            content=content,
            size_hint=(0.8, 0.5)
        )

        def add_more_breed(btn):
            popup.dismiss()
            self.show_breed_popup(instance, True)

        def save_and_exit(btn):
            popup.dismiss()
            self.show_success("Данные по площадке сохранены!")

        add_more_btn.bind(on_press=add_more_breed)
        save_exit_btn.bind(on_press=save_and_exit)

        popup.open()

    def update_row_total(self, instance, value):
        """Обновляем итоги по строке"""
        # Обновляем общие итоги страницы при изменении данных
        self.update_totals()

    def update_plot_total(self, instance, value):
        """Обновляем итог по площадке при изменении данных"""
        row_idx = instance.row_index

        breeds_text = self.inputs[row_idx][3].text
        breeds_data = self.parse_breeds_data(breeds_text)

        if not breeds_data:
            return

        total_density = 0
        total_height = 0.0
        total_age = 0
        breed_count = 0
        breed_names = []

        for breed_info in breeds_data:
            breed_count += 1
            breed_name = breed_info.get('name', 'Неизвестная')
            breed_names.append(breed_name)

            if breed_info.get('type') == 'coniferous':
                coniferous_density = (breed_info.get('do_05', 0) +
                                    breed_info.get('05_15', 0) +
                                    breed_info.get('bolee_15', 0))
                if coniferous_density > 0:
                    total_density += coniferous_density
            elif 'density' in breed_info and breed_info['density']:
                total_density += breed_info['density']

            if 'height' in breed_info and breed_info['height']:
                total_height += breed_info['height']
            if 'age' in breed_info and breed_info['age']:
                total_age += breed_info['age']

        # Метод завершен
