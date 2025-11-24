from kivy.uix.screenmanager import Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.properties import StringProperty, ListProperty, BooleanProperty
from kivy.utils import get_color_from_hex
from kivy.clock import Clock
from kivy.graphics import Color, RoundedRectangle
from kivy.animation import Animation
import json
import sqlite3
import os

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

class TaxationPopup(Popup):
    def __init__(self, **kwargs):
        super().__init__(
            title='Таксационные показатели участка',
            size_hint=(0.9, 0.9),
            separator_height=0,
            background_color=(0.5, 0.5, 0.5, 1),
            overlay_color=(0, 0, 0, 0.5),
            **kwargs
        )

        self.content = FloatLayout()

        # Заголовок
        title_label = Label(
            text='Создание таксационных характеристик участка молодняков',
            font_name='Roboto',
            font_size='20sp',
            bold=True,
            color=(1, 1, 1, 1),
            pos_hint={'center_x': 0.5, 'center_y': 0.95},
            size_hint=(None, None),
            size=(500, 50),
            halign='center'
        )

        # Кнопки выбора источника данных
        buttons_layout = BoxLayout(
            orientation='horizontal',
            spacing=20,
            size_hint=(0.8, None),
            height=50,
            pos_hint={'center_x': 0.5, 'center_y': 0.85}
        )

        # Кнопка для текущих данных приложения
        current_btn = ModernButton(
            text='Текущие данные',
            bg_color=get_color_from_hex('#00FF00'),
            size_hint=(0.5, 1)
        )
        current_btn.bind(on_press=self.calculate_from_current)

        # Кнопка для загрузки из файла
        load_btn = ModernButton(
            text='Загрузить из файла',
            bg_color=get_color_from_hex('#0000FF'),
            size_hint=(0.5, 1)
        )
        load_btn.bind(on_press=self.load_from_file)

        buttons_layout.add_widget(current_btn)
        buttons_layout.add_widget(load_btn)

        # Область результатов
        self.results_scroll = ScrollView(
            size_hint=(0.9, 0.7),
            pos_hint={'center_x': 0.5, 'center_y': 0.4}
        )

        self.results_layout = GridLayout(
            cols=1,
            spacing=10,
            size_hint_y=None
        )
        self.results_layout.bind(minimum_height=self.results_layout.setter('height'))

        self.results_scroll.add_widget(self.results_layout)

        # Кнопка закрытия
        close_btn = ModernButton(
            text='Закрыть',
            bg_color=get_color_from_hex('#FF0000'),
            size_hint=(None, None),
            size=(200, 50),
            pos_hint={'center_x': 0.5, 'center_y': 0.1}
        )
        close_btn.bind(on_press=self.dismiss)

        self.content.add_widget(title_label)
        self.content.add_widget(buttons_layout)
        self.content.add_widget(self.results_scroll)
        self.content.add_widget(close_btn)

    def calculate_from_current(self, instance):
        """Расчет из текущих данных в приложении"""
        try:
            from kivy.app import App
            molodniki_screen = App.get_running_app().root.get_screen('molodniki')

            if not molodniki_screen.page_data:
                self.show_error("Нет данных в текущем участке молодняков!")
                return

            self.calculate_taxation_data(molodniki_screen.page_data, molodniki_screen.current_radius)
        except Exception as e:
            self.show_error(f"Ошибка расчета: {str(e)}")

    def load_from_file(self, instance):
        """Загрузка данных из JSON файла"""
        from tkinter import Tk, filedialog
        Tk().withdraw()
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                radius = data.get('radius', '5.64')
                page_data = data.get('page_data', {})

                self.calculate_taxation_data(page_data, radius)
            except Exception as e:
                self.show_error(f"Ошибка загрузки файла: {str(e)}")

    def calculate_taxation_data(self, page_data, radius):
        """Основной расчет таксационных показателей"""
        try:
            radius_m = float(radius) if radius else 5.64
            plot_area_m2 = 3.14159 * (radius_m ** 2)
            plot_area_ha = plot_area_m2 / 10000  # Гектары

            # Словарь для сбора данных по породам
            breeds_data = {}

            # Обрабатываем все страницы
            for page_num, page_rows in page_data.items():
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
                            density = density_value  # Абсолютное количество деревьев, не на га
                            height = breed_info.get('height', 0) or 0

                        age = breed_info.get('age', 0) or 0

                        # Сбор данных по породе
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {
                                'type': breed_type,
                                'plots': [],
                                'coniferous_zones': {'do_05': 0, '05_15': 0, 'bolee_15': 0} if breed_type == 'coniferous' else None
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

                        if breed_type == 'coniferous':
                            breeds_data[breed_name]['coniferous_zones']['do_05'] += plot_data['do_05_density']
                            breeds_data[breed_name]['coniferous_zones']['05_15'] += plot_data['05_15_density']
                            breeds_data[breed_name]['coniferous_zones']['bolee_15'] += plot_data['bolee_15_density']

            # Формируем результаты
            self.display_taxation_results(breeds_data, plot_area_ha)

        except Exception as e:
            import traceback
            self.show_error(f"Ошибка расчета: {str(e)}\n{traceback.format_exc()}")

    def display_taxation_results(self, breeds_data, plot_area_ha):
        """Отображение результатов таксационных расчетов"""
        self.results_layout.clear_widgets()

        # Заголовок результатов с радиусом
        radius_val = float(self.get_radius_from_data()) if self.get_radius_from_data() else 5.64
        header_label = Label(
            text=f'РЕЗУЛЬТАТЫ ТАКСАЦИОННЫХ РАСЧЕТОВ\nРадиус участка: {radius_val:.2f} м',
            font_name='Roboto',
            font_size='18sp',
            bold=True,
            color=(0, 0.5, 0, 1),
            size_hint=(1, None),
            height=60,
            halign='center'
        )
        self.results_layout.add_widget(header_label)

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
        self.results_layout.add_widget(composition_label)

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
            while sum(int(part.rstrip('СБЕПОсЛКОсИОЧ').rstrip('0123456789')) if 'СБЕПОсЛКОсИОЧ' in part else 0 for part in composition_parts) != 10:
                coeffs_only = [int(''.join(filter(str.isdigit, part))) for part in composition_parts]
                total_coeffs = sum(coeffs_only)
                if total_coeffs > 10:
                    # Уменьшаем самый большой коэффициент
                    max_idx = coeffs_only.index(max(coeffs_only))
                    coeffs_only[max_idx] -= 1
                elif total_coeffs < 10:
                    # Увеличиваем самый большой коэффициент
                    max_idx = coeffs_only.index(max(coeffs_only))
                    coeffs_only[max_idx] += 1
                else:
                    break

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
            self.results_layout.add_widget(composition_result)
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
            self.results_layout.add_widget(no_composition)

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
        self.results_layout.add_widget(coniferous_label)

        has_coniferous = False
        for breed_name, data in sorted(breeds_data.items()):
            if data['type'] == 'coniferous' and data['plots']:
                has_coniferous = True

                # Средняя густота в градациях
                zones = data.get('coniferous_zones', {})
                avg_do_05 = zones.get('do_05', 0) / len(data['plots']) if data['plots'] else 0
                avg_05_15 = zones.get('05_15', 0) / len(data['plots']) if data['plots'] else 0
                avg_bolee_15 = zones.get('bolee_15', 0) / len(data['plots']) if data['plots'] else 0

                # Средняя общая высота
                avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                avg_height_total = sum(avg_heights) / len(avg_heights) if avg_heights else 0

                coniferous_result = Label(
                    text=f"{breed_name}:\n"
                         f"• до 0.5м: {avg_do_05:.1f} шт/га\n"
                         f"• 0.5-1.5м: {avg_05_15:.1f} шт/га\n"
                         f"• >1.5м: {avg_bolee_15:.1f} шт/га\n"
                         f"• средняя высота породы: {avg_height_total:.1f}м",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0, 0.5, 0, 1),
                    size_hint=(1, None),
                    height=100,
                    halign='left',
                    valign='top'
                )
                coniferous_result.bind(size=lambda *args: setattr(coniferous_result, 'text_size', (coniferous_result.width, None)))
                self.results_layout.add_widget(coniferous_result)

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
            self.results_layout.add_widget(no_coniferous)

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
        self.results_layout.add_widget(deciduous_label)

        has_deciduous = False
        for breed_name, data in sorted(breeds_data.items()):
            if data['type'] == 'deciduous' and data['plots']:
                has_deciduous = True

                # Средняя густота
                avg_density = sum(p['density'] for p in data['plots']) / len(data['plots'])

                # Средняя высота
                avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0

                # Средний возраст
                avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                deciduous_result = Label(
                    text=f"{breed_name}:\n"
                         f"• Средняя густота: {avg_density:.1f} шт\n"
                         f"• Средняя высота: {avg_height:.1f}м\n"
                         f"• Средний возраст: {avg_age:.1f} лет",
                    font_name='Roboto',
                    font_size='14sp',
                    color=(0, 0.3, 0.5, 1),
                    size_hint=(1, None),
                    height=80,
                    halign='left',
                    valign='top'
                )
                deciduous_result.bind(size=lambda *args: setattr(deciduous_result, 'text_size', (deciduous_result.width, None)))
                self.results_layout.add_widget(deciduous_result)

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
            self.results_layout.add_widget(no_deciduous)

        # Расчет густоты исходя из радиуса
        radius_val = float(self.get_radius_from_data()) if self.get_radius_from_data() else 5.64
        area_m2 = 3.14159 * (radius_val ** 2)
        trees_per_ha = 10000 / area_m2 if area_m2 > 0 else 0  # шт/га на 1 дерево

        density_label = Label(
            text=f"Густота: 1 дерево = {trees_per_ha:.0f} тыс.шт./га",
            font_name='Roboto',
            font_size='14sp',
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=30,
            halign='center'
        )
        self.results_layout.add_widget(density_label)

        # Информация о площади участка
        plot_area_label = Label(
            text=f"Площадь участка: {plot_area_ha:.4f} га (радиус пробной площади: {radius_val:.2f}м)",
            font_name='Roboto',
            font_size='12sp',
            color=(0.5, 0.5, 0.5, 1),
            size_hint=(1, None),
            height=30,
            halign='center'
        )
        self.results_layout.add_widget(plot_area_label)

    def get_radius_from_data(self):
        """Получение радиуса из текущих данных"""
        try:
            from kivy.app import App
            molodniki_screen = App.get_running_app().root.get_screen('molodniki')
            return molodniki_screen.current_radius
        except:
            return '5.64'

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
        """Показать сообщение об ошибке"""
        error_popup = Popup(
            title='Ошибка',
            content=Label(text=message, color=(1, 0, 0, 1)),
            size_hint=(0.6, 0.3)
        )
        error_popup.open()
