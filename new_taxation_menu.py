"""
Таксационное меню — KivyMD
Расчёт таксационных показателей с единым стилем ForestApp
"""
import json
import sqlite3
import os
import traceback

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout

from kivymd.uix.card import MDCard
from kivymd.uix.button import MDButton, MDButtonText, MDIconButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.textfield import MDTextField

from ui_styles import Colors, Spacing, Fonts


class TaxationPopup(Popup):
    def __init__(self, **kwargs):
        super().__init__(
            title='',
            size_hint=(0.92, 0.92),
            separator_height=0,
            background_color=[0,0,0,0],
            **kwargs
        )
        self._build_ui()

    def _build_ui(self):
        content = MDBoxLayout(orientation='vertical', spacing=0)

        # Header
        header = MDCard(
            orientation='horizontal',
            size_hint_y=None, height=dp(56),
            md_bg_color=Colors.SECONDARY,
            padding=[dp(12), dp(8)],
            radius=[0],
            elevation=4,
        )
        header.add_widget(MDLabel(
            text='Таксационные расчёты', font_size='20sp', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            size_hint_x=0.85, valign='middle',
        ))
        close_btn = MDIconButton(
            icon='close', font_size='24sp',
            theme_text_color='Custom', text_color=[1,1,1,1],
            on_release=self.dismiss,
        )
        header.add_widget(close_btn)
        content.add_widget(header)

        # Источники данных
        source_section = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            padding=[dp(16), dp(12)],
            spacing=dp(12),
            md_bg_color=Colors.SURFACE,
        )

        source_section.add_widget(MDLabel(
            text='ИСТОЧНИК ДАННЫХ', font_style='Label', role='small',
            theme_text_color='Secondary', size_hint_y=None, height=dp(20),
        ))

        btn_row = MDBoxLayout(orientation='horizontal', spacing=dp(12), size_hint_y=None, height=dp(56))
        current_btn = MDCard(
            orientation='horizontal', size_hint=(0.5, None), height=dp(48),
            radius=[dp(12)], elevation=2, md_bg_color=Colors.SUCCESS_BG,
            padding=[dp(12), dp(4)], spacing=dp(8),
            on_release=self.calculate_from_current,
        )
        current_btn.add_widget(MDIconButton(
            icon='database', font_size='22sp',
            theme_text_color='Custom', text_color=Colors.PRIMARY,
            on_release=lambda x: None,
        ))
        current_btn.add_widget(MDLabel(
            text='Текущие данные', font_size='14sp', bold=True,
            theme_text_color='Custom', text_color=Colors.PRIMARY,
            adaptive_height=True, valign='middle',
        ))
        btn_row.add_widget(current_btn)

        load_btn = MDCard(
            orientation='horizontal', size_hint=(0.5, None), height=dp(48),
            radius=[dp(12)], elevation=2, md_bg_color=Colors.INFO_BG,
            padding=[dp(12), dp(4)], spacing=dp(8),
            on_release=self.load_from_file,
        )
        load_btn.add_widget(MDIconButton(
            icon='file-upload', font_size='22sp',
            theme_text_color='Custom', text_color=Colors.INFO,
            on_release=lambda x: None,
        ))
        load_btn.add_widget(MDLabel(
            text='Из файла', font_size='14sp', bold=True,
            theme_text_color='Custom', text_color=Colors.INFO,
            adaptive_height=True, valign='middle',
        ))
        btn_row.add_widget(load_btn)
        source_section.add_widget(btn_row)
        content.add_widget(source_section)

        # Результаты
        results_header = MDCard(
            orientation='horizontal',
            size_hint_y=None, height=dp(40),
            md_bg_color=Colors.SURFACE_ALT,
            padding=[dp(16), dp(8)],
            radius=[0],
        )
        results_header.add_widget(MDIconButton(
            icon='chart-box-outline', font_size='20sp',
            theme_text_color='Custom', text_color=Colors.PRIMARY,
            on_release=lambda x: None,
        ))
        results_header.add_widget(MDLabel(
            text='РЕЗУЛЬТАТЫ', font_style='Label', role='small',
            theme_text_color='Primary', bold=True,
            size_hint_x=0.9, valign='middle',
        ))
        content.add_widget(results_header)

        self.results_scroll = MDScrollView(size_hint=(1, 1), bar_width=dp(4))
        self.results_layout = MDBoxLayout(
            orientation='vertical', size_hint_y=None,
            spacing=dp(8), padding=[dp(16), dp(12)],
        )
        self.results_layout.bind(minimum_height=self.results_layout.setter('height'))

        placeholder = MDCard(
            orientation='vertical',
            size_hint=(1, None), height=dp(120),
            padding=dp(24), radius=[dp(12)],
            md_bg_color=[0.97,0.97,0.97,1],
        )
        placeholder.add_widget(MDIconButton(
            icon='calculator-variant', font_size='48sp',
            theme_text_color='Custom', text_color=[0.8,0.8,0.8,1],
            pos_hint={'center_x': 0.5},
            on_release=lambda x: None,
        ))
        placeholder.add_widget(MDLabel(
            text='Выберите источник данных для расчёта',
            font_size='14sp', theme_text_color='Hint',
            halign='center', adaptive_height=True,
        ))
        self.results_layout.add_widget(placeholder)
        self.results_scroll.add_widget(self.results_layout)
        content.add_widget(self.results_scroll)

        self.content = content

    def calculate_from_current(self, instance):
        try:
            molodniki_screen = App.get_running_app().root.get_screen('molodniki')
            if not molodniki_screen.page_data:
                self.show_info('Нет данных в текущем участке молодняков!')
                return
            self.calculate_taxation_data(molodniki_screen.page_data, molodniki_screen.current_radius)
        except Exception as e:
            self.show_info(f'Ошибка расчета: {str(e)}')

    def load_from_file(self, instance):
        from tkinter import Tk, filedialog
        Tk().withdraw()
        file_path = filedialog.askopenfilename(
            filetypes=[('JSON files', '*.json'), ('All files', '*.*')]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                radius = data.get('radius', '5.64')
                page_data = data.get('page_data', {})
                self.calculate_taxation_data(page_data, radius)
            except Exception as e:
                self.show_info(f'Ошибка загрузки файла: {str(e)}')

    def calculate_taxation_data(self, page_data, radius):
        try:
            radius_m = float(radius) if radius else 5.64
            plot_area_m2 = 3.14159 * (radius_m ** 2)
            plot_area_ha = plot_area_m2 / 10000

            breeds_data = {}
            for page_num, page_rows in page_data.items():
                if isinstance(page_num, str) and page_num.isdigit():
                    page_num = int(page_num)
                for row in page_rows:
                    if len(row) < 4:
                        continue
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
                        if breed_type == 'coniferous':
                            do_05 = breed_info.get('do_05', 0)
                            _05_15 = breed_info.get('05_15', 0)
                            bolee_15 = breed_info.get('bolee_15', 0)
                            density = (do_05 + _05_15 + bolee_15) / plot_area_ha if plot_area_ha > 0 else 0
                            if any([do_05, _05_15, bolee_15]):
                                if bolee_15 > 0:
                                    height = 2.0
                                elif _05_15 > 0:
                                    height = 1.0
                                elif do_05 > 0:
                                    height = 0.3
                            else:
                                height = breed_info.get('height', 0) or 0
                        else:
                            density = breed_info.get('density', 0)
                            height = breed_info.get('height', 0) or 0
                        age = breed_info.get('age', 0) or 0
                        if breed_name not in breeds_data:
                            breeds_data[breed_name] = {
                                'type': breed_type, 'plots': [],
                                'coniferous_zones': {'do_05': 0, '05_15': 0, 'bolee_15': 0} if breed_type == 'coniferous' else None
                            }
                        plot_data = {'density': density, 'height': height, 'age': age}
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

            self.display_taxation_results(breeds_data, plot_area_ha)
        except Exception as e:
            self.show_info(f'Ошибка расчета: {str(e)}\n{traceback.format_exc()}')

    def display_taxation_results(self, breeds_data, plot_area_ha):
        self.results_layout.clear_widgets()

        radius_val = 5.64
        try:
            radius_val = float(App.get_running_app().root.get_screen('molodniki').current_radius)
        except Exception:
            pass

        # Карточка с формулой состава
        total_densities = {}
        for breed_name, data in breeds_data.items():
            if data['plots']:
                conif_types = {'do_05_density', '05_15_density', 'bolee_15_density'}
                if any(k in data['plots'][0] for k in conif_types):
                    total_density = sum(p.get('do_05_density', 0) + p.get('05_15_density', 0) + p.get('bolee_15_density', 0) for p in data['plots'])
                else:
                    total_density = sum(p.get('density', 0) for p in data['plots'])
                if total_density > 0:
                    total_densities[breed_name] = total_density

        if total_densities:
            total_all_density = sum(total_densities.values())
            composition_parts = []
            for breed_name, density in sorted(total_densities.items(), key=lambda x: x[1], reverse=True):
                coeff = max(1, round(density / total_all_density * 10)) if total_all_density > 0 else 1
                breed_letter = self.get_breed_letter(breed_name)
                composition_parts.append({'breed': breed_name, 'coeff': coeff, 'letter': breed_letter})

            # Корректировка коэффициентов, чтобы сумма равнялась 10
            sorted_breeds = sorted(total_densities.items(), key=lambda x: x[1], reverse=True)
            for _ in range(20):
                coeffs = [p['coeff'] for p in composition_parts]
                total = sum(coeffs)
                if total == 10:
                    break
                if total > 10:
                    max_idx = coeffs.index(max(coeffs))
                    composition_parts[max_idx]['coeff'] -= 1
                else:
                    max_idx = coeffs.index(max(coeffs))
                    composition_parts[max_idx]['coeff'] += 1

            composition_text = ''.join(f"{p['coeff']}{p['letter']}" for p in composition_parts) + 'Др'

            comp_card = MDCard(
                orientation='vertical', size_hint=(1, None),
                padding=dp(16), spacing=dp(8), radius=[dp(12)],
                elevation=2, md_bg_color=Colors.PRIMARY_LIGHT,
            )
            comp_card.add_widget(MDLabel(
                text=f'Формула состава: {composition_text}',
                font_size='18sp', bold=True,
                theme_text_color='Custom', text_color=[1,1,1,1],
                halign='center', size_hint_y=None, height=dp(32),
            ))
            comp_card.add_widget(MDLabel(
                text=f'Радиус: {radius_val:.2f} м | Площадь: {plot_area_ha:.4f} га',
                font_size='12sp',
                theme_text_color='Custom', text_color=[1,1,1,0.8],
                halign='center', size_hint_y=None, height=dp(24),
            ))
            self.results_layout.add_widget(comp_card)

        # Хвойные породы
        conif_card = MDCard(
            orientation='vertical', size_hint=(1, None),
            padding=dp(12), spacing=dp(8), radius=[dp(12)],
            elevation=1, md_bg_color=Colors.INFO_BG,
        )
        conif_card.add_widget(MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(32), spacing=dp(8)))
        conif_card.children[0].add_widget(MDIconButton(
            icon='pine-tree', font_size='20sp',
            theme_text_color='Custom', text_color=Colors.INFO,
            on_release=lambda x: None,
        ))
        conif_card.children[0].add_widget(MDLabel(
            text='ХВОЙНЫЕ ПОРОДЫ', font_style='Label', role='small', bold=True,
            theme_text_color='Custom', text_color=Colors.INFO,
            adaptive_height=True, valign='middle',
        ))

        has_coniferous = False
        for breed_name, data in sorted(breeds_data.items()):
            if data['type'] == 'coniferous' and data['plots']:
                has_coniferous = True
                zones = data.get('coniferous_zones', {})
                n = len(data['plots'])
                avg_do_05 = zones.get('do_05', 0) / n if n else 0
                avg_05_15 = zones.get('05_15', 0) / n if n else 0
                avg_bolee_15 = zones.get('bolee_15', 0) / n if n else 0
                avg_height_total = sum(p['height'] for p in data['plots'] if p['height'] > 0)
                avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                avg_height = avg_height_total / len(avg_heights) if avg_heights else 0
                avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                breed_card = MDCard(
                    orientation='vertical', size_hint=(1, None),
                    padding=dp(12), spacing=dp(4), radius=[dp(8)],
                    elevation=1, md_bg_color=[0.95,0.97,1,1],
                )
                breed_card.add_widget(MDLabel(
                    text=breed_name, font_size='15sp', bold=True,
                    theme_text_color='Custom', text_color=Colors.INFO,
                    size_hint_y=None, height=dp(24),
                ))
                details = (
                    f'до 0.5м: {avg_do_05:.1f} шт/га\n'
                    f'0.5-1.5м: {avg_05_15:.1f} шт/га\n'
                    f'>1.5м: {avg_bolee_15:.1f} шт/га\n'
                    f'Ср. высота: {avg_height:.1f}м | Ср. возраст: {avg_age:.1f} лет'
                )
                breed_card.add_widget(MDLabel(
                    text=details, font_size='12sp',
                    theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY,
                    adaptive_height=True,
                ))
                conif_card.add_widget(breed_card)

        if not has_coniferous:
            conif_card.add_widget(MDLabel(
                text='Хвойные породы не найдены', font_size='13sp',
                theme_text_color='Hint', halign='center', size_hint_y=None, height=dp(32),
            ))
        self.results_layout.add_widget(conif_card)

        # Лиственные породы
        decid_card = MDCard(
            orientation='vertical', size_hint=(1, None),
            padding=dp(12), spacing=dp(8), radius=[dp(12)],
            elevation=1, md_bg_color=Colors.SUCCESS_BG,
        )
        decid_card.add_widget(MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(32), spacing=dp(8)))
        decid_card.children[0].add_widget(MDIconButton(
            icon='leaf', font_size='20sp',
            theme_text_color='Custom', text_color=Colors.PRIMARY,
            on_release=lambda x: None,
        ))
        decid_card.children[0].add_widget(MDLabel(
            text='ЛИСТВЕННЫЕ ПОРОДЫ', font_style='Label', role='small', bold=True,
            theme_text_color='Custom', text_color=Colors.PRIMARY,
            adaptive_height=True, valign='middle',
        ))

        has_deciduous = False
        for breed_name, data in sorted(breeds_data.items()):
            if data['type'] == 'deciduous' and data['plots']:
                has_deciduous = True
                n = len(data['plots'])
                avg_density = sum(p['density'] for p in data['plots']) / n
                avg_heights = [p['height'] for p in data['plots'] if p['height'] > 0]
                avg_height = sum(avg_heights) / len(avg_heights) if avg_heights else 0
                avg_ages = [p['age'] for p in data['plots'] if p['age'] > 0]
                avg_age = sum(avg_ages) / len(avg_ages) if avg_ages else 0

                breed_card = MDCard(
                    orientation='vertical', size_hint=(1, None),
                    padding=dp(12), spacing=dp(4), radius=[dp(8)],
                    elevation=1, md_bg_color=[0.95,1,0.95,1],
                )
                breed_card.add_widget(MDLabel(
                    text=breed_name, font_size='15sp', bold=True,
                    theme_text_color='Custom', text_color=Colors.PRIMARY,
                    size_hint_y=None, height=dp(24),
                ))
                details = (
                    f'Ср. густота: {avg_density:.1f} шт\n'
                    f'Ср. высота: {avg_height:.1f}м\n'
                    f'Ср. возраст: {avg_age:.1f} лет'
                )
                breed_card.add_widget(MDLabel(
                    text=details, font_size='12sp',
                    theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY,
                    adaptive_height=True,
                ))
                decid_card.add_widget(breed_card)

        if not has_deciduous:
            decid_card.add_widget(MDLabel(
                text='Лиственные породы не найдены', font_size='13sp',
                theme_text_color='Hint', halign='center', size_hint_y=None, height=dp(32),
            ))
        self.results_layout.add_widget(decid_card)

        # Информация
        area_m2 = 3.14159 * (radius_val ** 2)
        trees_per_ha = 10000 / area_m2 if area_m2 > 0 else 0
        info_card = MDCard(
            orientation='vertical', size_hint=(1, None),
            padding=dp(12), spacing=dp(4), radius=[dp(12)],
            elevation=1, md_bg_color=[0.98,0.98,0.98,1],
        )
        info_card.add_widget(MDLabel(
            text=f'Густота: 1 дерево = {trees_per_ha:.0f} тыс.шт./га',
            font_size='13sp', theme_text_color='Custom', text_color=Colors.TEXT_DIM,
            adaptive_height=True,
        ))
        info_card.add_widget(MDLabel(
            text=f'Площадь: {plot_area_ha:.4f} га (радиус: {radius_val:.2f}м)',
            font_size='13sp', theme_text_color='Custom', text_color=Colors.TEXT_DIM,
            adaptive_height=True,
        ))
        self.results_layout.add_widget(info_card)

    def get_breed_letter(self, breed_name):
        breed_letters = {
            'Сосна': 'С', 'Ель': 'Е', 'Пихта': 'П', 'Кедр': 'К',
            'Лиственница': 'Л', 'Берёза': 'Б', 'Осина': 'Ос',
            'Ольха чёрная': 'ОЧ', 'Ольха серая': 'ОС', 'Ива': 'И',
            'Ива кустарниковая': 'ИК',
        }
        for full_name, letter in breed_letters.items():
            if full_name.lower() in breed_name.lower():
                return letter
        return breed_name[0].upper() if breed_name else 'Н'

    def show_info(self, message):
        snack = MDSnackbar(duration=3)
        snack.add_widget(MDSnackbarText(text=message))
        snack.open()


class ModernTaxationPopup(Popup):
    """Альтернативный попап с табличным вводом таксационных данных"""
    def __init__(self, **kwargs):
        super().__init__(
            title='',
            size_hint=(0.95, 0.95),
            separator_height=0,
            **kwargs
        )
        self._build_ui()

    def _build_ui(self):
        content = MDBoxLayout(orientation='vertical', spacing=0)

        header = MDCard(
            orientation='horizontal',
            size_hint_y=None, height=dp(56),
            md_bg_color=Colors.SECONDARY,
            padding=[dp(12), dp(8)],
            radius=[0], elevation=4,
        )
        header.add_widget(MDLabel(
            text='Ввод таксационных показателей', font_size='20sp', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            valign='middle',
        ))
        header.add_widget(MDIconButton(
            icon='close', font_size='24sp',
            theme_text_color='Custom', text_color=[1,1,1,1],
            on_release=self.dismiss,
        ))
        content.add_widget(header)

        form_scroll = MDScrollView(size_hint=(1, 1))
        form = MDBoxLayout(
            orientation='vertical', size_hint_y=None,
            spacing=dp(12), padding=[dp(20), dp(16)],
        )
        form.bind(minimum_height=form.setter('height'))

        fields = [
            ('Номер участка', 'section_input'),
            ('Площадь, га', 'area_input'),
            ('Радиус площадки, м', 'radius_input'),
            ('Количество площадок', 'plots_count'),
            ('Метод таксации', 'method_input'),
        ]
        for label_text, attr_name in fields:
            field_box = MDBoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(4), adaptive_height=True)
            field_box.add_widget(MDLabel(
                text=label_text, font_size='13sp', bold=True,
                theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY,
                size_hint_y=None, height=dp(20),
            ))
            inp = MDTextField(mode='outlined', size_hint_y=None, height=dp(48), font_size='16sp')
            setattr(self, attr_name, inp)
            field_box.add_widget(inp)
            form.add_widget(field_box)

        form_scroll.add_widget(form)
        content.add_widget(form_scroll)

        btn_bar = MDBoxLayout(
            orientation='horizontal', size_hint_y=None, height=dp(56),
            spacing=dp(12), padding=[dp(16), dp(8)],
            md_bg_color=Colors.SURFACE,
        )
        calc_btn = MDButton(style='filled', md_bg_color=Colors.PRIMARY,
                              size_hint=(0.5, None), height=dp(44),
                              on_release=self.calculate)
        calc_btn.add_widget(MDButtonText(text='Рассчитать'))
        btn_bar.add_widget(calc_btn)
        close_calc_btn = MDButton(style='filled', md_bg_color=Colors.DANGER,
                                    size_hint=(0.5, None), height=dp(44),
                                    on_release=self.dismiss)
        close_calc_btn.add_widget(MDButtonText(text='Закрыть'))
        btn_bar.add_widget(close_calc_btn)
        content.add_widget(btn_bar)

        self.content = content

    def calculate(self, instance):
        snack = MDSnackbar(duration=2)
        snack.add_widget(MDSnackbarText(text='✅ Расчёт выполнен (в разработке)'))
        snack.open()
