"""
ForestApp — Единое лесотаксационное приложение
KivyMD (Material Design) с системой иконок и единым UI
"""
import os
import sys
import json
import sqlite3
import datetime
import logging

from kivy.app import App
from kivy.core.window import Window
from kivy.config import Config
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.core.text import LabelBase
from kivy.properties import ObjectProperty, StringProperty, ListProperty
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.utils import get_color_from_hex
from kivy.uix.widget import Widget

from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.card import MDCard
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.dialog import (
    MDDialog,
    MDDialogHeadlineText,
    MDDialogSupportingText,
    MDDialogButtonContainer,
    MDDialogContentContainer,
)
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.label import MDLabel
from kivymd.uix.textfield import MDTextField
from kivymd.uix.list import MDListItem, MDListItemHeadlineText, MDListItemLeadingIcon
from kivymd.uix.appbar import MDTopAppBar, MDTopAppBarLeadingButtonContainer, MDTopAppBarTrailingButtonContainer, MDTopAppBarTitle, MDActionTopAppBarButton
from kivymd.uix.navigationdrawer import MDNavigationDrawer
from kivymd.uix.gridlayout import MDGridLayout

Config.set('graphics', 'width', '480')
Config.set('graphics', 'height', '854')
Config.set('graphics', 'resizable', True)
Config.set('input', 'mouse', 'mouse,multitouch_on_demand')

LabelBase.register(name='Roboto',
                   fn_regular='fonts/Roboto-Medium.ttf',
                   fn_bold='fonts/Roboto-Bold.ttf')

from ui_styles import Colors, Spacing, Fonts, ModernButton
from theme_manager import ThemeManager


def make_raised_btn(text, **kwargs):
    icon = kwargs.pop('icon', None)
    text_kwargs = {}
    for key in ['font_size', 'color', 'theme_text_color', 'text_color', 'bold']:
        if key in kwargs:
            text_kwargs[key] = kwargs.pop(key)
    if 'color' in text_kwargs and 'text_color' not in text_kwargs:
        text_kwargs['text_color'] = text_kwargs.pop('color')
    btn = MDButton(style='filled', **kwargs)
    if icon:
        btn.add_widget(MDButtonIcon(icon=icon))
    btn.add_widget(MDButtonText(text=text, **text_kwargs))
    return btn

def make_outlined_btn(text, **kwargs):
    icon = kwargs.pop('icon', None)
    text_kwargs = {}
    for key in ['font_size', 'color', 'theme_text_color', 'text_color', 'bold']:
        if key in kwargs:
            text_kwargs[key] = kwargs.pop(key)
    if 'color' in text_kwargs and 'text_color' not in text_kwargs:
        text_kwargs['text_color'] = text_kwargs.pop('color')
    btn = MDButton(style='outlined', **kwargs)
    if icon:
        btn.add_widget(MDButtonIcon(icon=icon))
    btn.add_widget(MDButtonText(text=text, **text_kwargs))
    return btn

class MDTopAppBarOld(MDTopAppBar):
    def __init__(self, title='', anchor_title='left', elevation=2, md_bg_color=None,
                 specific_text_color=None, left_action_items=None, right_action_items=None,
                 **kwargs):
        super().__init__(type='small', elevation=elevation, md_bg_color=md_bg_color, **kwargs)
        leading = MDTopAppBarLeadingButtonContainer()
        if left_action_items:
            for icon, callback in left_action_items:
                leading.add_widget(MDActionTopAppBarButton(icon=icon, on_release=callback))
        self.add_widget(leading)
        self.add_widget(MDTopAppBarTitle(text=title))
        trailing = MDTopAppBarTrailingButtonContainer()
        if right_action_items:
            for icon, callback in right_action_items:
                trailing.add_widget(MDActionTopAppBarButton(icon=icon, on_release=callback))
        self.add_widget(trailing)


# Backward-compatible imports for molodniki_extended.py
from core.widgets import AutoCompleteTextInput, TreeDataInputPopup, ExitConfirmPopup

from screens.table_screen import TableScreen
from molodniki_extended import ExtendedMolodnikiTableScreen
from new_taxation_menu import TaxationPopup


class LazyScreenManager(MDScreenManager):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._factories = {}

    def register_factory(self, name, factory):
        self._factories[name] = factory

    def get_screen(self, name):
        if name in self._factories and name not in self.screen_names:
            screen = self._factories[name]()
            if not screen.name:
                screen.name = name
            self.add_widget(screen)
            self._factories.pop(name, None)
        return super().get_screen(name)


class MainMenu(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'main'
        self._ui_built = False
        Clock.schedule_once(lambda dt: self.create_ui(), 0)

    def create_ui(self):
        self._build_ui()
        self._ui_built = True

    def _build_ui(self):
        self.clear_widgets()
        app = App.get_running_app()
        if not app or not hasattr(app, 'theme_manager'):
            return
        theme = app.theme_manager.current_theme
        bg = theme['background'] if theme['type'] == 'color' else (0.15, 0.18, 0.2, 1)
        Window.clearcolor = bg

        main = MDBoxLayout(orientation='vertical')
        scroll = MDScrollView(size_hint=(1, 1), bar_width=dp(4))
        center = MDBoxLayout(
            orientation='vertical', size_hint_y=None,
            spacing=dp(8), padding=[dp(12), dp(8), dp(12), dp(24)],
        )
        center.bind(minimum_height=center.setter('height'))

        logo_card = MDCard(
            orientation='vertical', size_hint=(1, None), height=dp(100),
            padding=dp(16), radius=[dp(16)], elevation=3,
            md_bg_color=Colors.PRIMARY,
        )
        logo_card.add_widget(MDLabel(
            text='ФАНАТЫ ПИХТЫ', font_size='32sp', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            halign='center', size_hint_y=None, height=dp(48),
        ))
        logo_card.add_widget(MDLabel(
            text='Лесное таксационное приложение',
            font_size='13sp', theme_text_color='Custom',
            text_color=[1,1,1,0.9], halign='center',
            size_hint_y=None, height=dp(22),
        ))
        center.add_widget(logo_card)

        center.add_widget(Widget(size_hint_y=None, height=dp(6)))

        # Основные разделы с иконками
        sections = [
            ('ПЕРЕЧЁТНАЯ ВЕДОМОСТЬ', self.show_tally_submenu,
             'file-document-outline', Colors.PRIMARY),
            ('РУМ (МОЛОДНЯКИ)', self.show_molodniki_submenu,
             'seedling', Colors.SECONDARY),
            ('ТАКСАЦИОННЫЕ РАСЧЁТЫ', self.show_taxation,
             'calculator', Colors.ACCENT),
            ('СПРАВОЧНИКИ И НОРМАТИВЫ', self.show_ref_submenu,
             'bookshelf', Colors.INFO),
            ('ОТЧЁТЫ', self.show_reports,
             'file-chart', Colors.BTN_PURPLE),
        ]
        for label, cb, icon, color in sections:
            center.add_widget(self._make_card(icon, label, cb, color))

        center.add_widget(Widget(size_hint_y=None, height=dp(4)))

        sys_card = self._make_card(
            'cog-outline', 'СИСТЕМА',
            self.show_sys_submenu, [0.5, 0.5, 0.5, 1],
        )
        center.add_widget(sys_card)

        center.add_widget(MDLabel(
            text='Качество — Надёжность — Точность',
            font_size='10sp', theme_text_color='Hint',
            halign='center', size_hint_y=None, height=dp(20),
        ))

        scroll.add_widget(center)
        main.add_widget(scroll)
        self.add_widget(main)

    def _make_card(self, icon, text, callback, color):
        card = MDCard(
            orientation='horizontal',
            size_hint=(1, None), height=dp(72),
            padding=[dp(12), dp(8)], spacing=dp(16),
            radius=[dp(12)], elevation=2,
            md_bg_color=[0.18, 0.18, 0.18, 0.95],
            on_release=callback,
        )
        left_bar = MDBoxLayout(
            orientation='vertical',
            size_hint_x=None, width=dp(6),
            md_bg_color=color,
        )
        card.add_widget(left_bar)
        icon_w = MDIconButton(
            icon=icon, font_size='28sp',
            theme_icon_color='Custom',
            icon_color=[1,1,1,1],
            on_release=lambda: None,
        )
        card.add_widget(icon_w)
        label = MDLabel(
            text=text, font_size='17sp', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            adaptive_height=True, valign='middle',
        )
        card.add_widget(label)
        return card

    def _fullscreen_submenu(self, title, items):
        overlay = MDBoxLayout(orientation='vertical')
        sheet = MDCard(
            orientation='vertical',
            size_hint=(1, None), height=dp(600),
            radius=[dp(20), dp(20), 0, 0],
            elevation=8, md_bg_color=[0.15, 0.15, 0.15, 1],
            pos_hint={'x': 0, 'y': 0},
        )
        header = MDBoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(56),
            md_bg_color=Colors.PRIMARY,
            padding=[dp(8), dp(4)],
        )
        header.add_widget(MDLabel(
            text=title, font_size='20sp', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            size_hint_x=0.8, valign='middle',
        ))
        header.add_widget(Widget())
        close_btn = MDIconButton(
            icon='arrow-left', font_size='24sp',
            theme_icon_color='Custom', icon_color=[1,1,1,1],
            on_release=self._dismiss_submenu,
        )
        header.add_widget(close_btn)
        sheet.add_widget(header)
        scroll = MDScrollView(size_hint=(1, 1))
        layout = MDBoxLayout(
            orientation='vertical', size_hint_y=None,
            spacing=dp(4), padding=[dp(12), dp(12)],
        )
        layout.bind(minimum_height=layout.setter('height'))
        for icon, text, cb in items:
            item_card = MDCard(
                orientation='horizontal',
                size_hint=(1, None), height=dp(56),
                padding=[dp(16), dp(8)], spacing=dp(16),
                radius=[dp(12)], elevation=1,
                md_bg_color=[0.25, 0.25, 0.25, 1],
                on_release=lambda x, c=cb: (self._dismiss_submenu(), c(x)),
            )
            ic = MDIconButton(
                icon=icon, font_size='24sp',
                theme_icon_color='Custom', icon_color=[1,1,1,1],
            )
            item_card.add_widget(ic)
            item_card.add_widget(MDLabel(
                text=text, font_size='16sp',
                theme_text_color='Custom', text_color=[1,1,1,1],
                adaptive_height=True, valign='middle',
            ))
            layout.add_widget(item_card)
        scroll.add_widget(layout)
        sheet.add_widget(scroll)
        overlay.add_widget(sheet)
        self._submenu_overlay = overlay
        self.add_widget(overlay)

    def _dismiss_submenu(self, *args):
        if hasattr(self, '_submenu_overlay') and self._submenu_overlay:
            self.remove_widget(self._submenu_overlay)
            self._submenu_overlay = None

    def _go(self, screen):
        App.get_running_app().root.current = screen

    def show_tally_submenu(self, instance):
        self._fullscreen_submenu('Перечётная ведомость', [
            ('plus-circle', 'Новый участок (сплошной)', self.show_new_section_form),
            ('folder-open', 'Загрузить участок (сплошной)', self.show_load_section),
            ('format-list-bulleted', 'Сплошной перечёт', lambda *a: self._go('table')),
            ('table-large', 'Таблица перечёта', lambda *a: self._go('table')),
        ])

    def show_molodniki_submenu(self, instance):
        self._fullscreen_submenu('РУМ (Молодняки)', [
            ('plus-circle', 'Новый участок молодняков', self.show_new_molodniki_section),
            ('folder-open', 'Загрузить участок из БД', self.show_load_molodniki_section),
            ('file-upload-outline', 'Загрузить из JSON', self.show_load_molodniki_json),
            ('file-excel', 'Загрузить из Excel', self.show_load_molodniki_excel),
            ('seedling', 'Расширенная таблица', lambda *a: self._go('molodniki')),
        ])

    def show_ref_submenu(self, instance):
        self._fullscreen_submenu('Справочники и нормативы', [
            ('file-document', 'Нормативы таксации', self.show_normatives_doc),
            ('book-open-variant', 'Справочные таблицы', self.show_normatives_doc),
            ('calculator', 'Таксационные расчёты', self.show_taxation),
        ])

    def show_sys_submenu(self, instance):
        self._fullscreen_submenu('Система', [
            ('palette', 'Темы оформления', self.show_theme_chooser),
            ('information-outline', 'О программе', self.show_about),
            ('exit-run', 'Выход', self.confirm_exit),
        ])

    def show_taxation(self, instance):
        TaxationPopup().open()

    def show_reports(self, instance):
        try:
            os.startfile('reports')
        except Exception:
            self._snack('Папка отчётов не найдена')

    def show_normatives_doc(self, instance):
        try:
            os.startfile('normativs')
        except Exception:
            self._snack('Папка нормативов не найдена')

    def show_about(self, instance):
        dialog = MDDialog(
            MDDialogHeadlineText(text='🌲 Фанаты Пихты'),
            MDDialogSupportingText(
                text='Версия 3.0\n\n'
                     'Лесное таксационное приложение для\n'
                     'учёта лесных насаждений и расчёта\n'
                     'таксационных показателей.\n\n'
                     '© 2025 Все права защищены',
            ),
            MDDialogButtonContainer(
                make_outlined_btn('Закрыть', on_release=lambda x: dialog.dismiss()),
                spacing='8dp',
            ),
        )
        dialog.open()

    def show_theme_chooser(self, instance):
        content = MDBoxLayout(orientation='vertical', spacing=dp(8), padding=dp(16),
                              md_bg_color=[0.15, 0.15, 0.15, 1])
        content.add_widget(MDLabel(text='Выбор темы', font_style='Title', role='large', bold=True,
                                   theme_text_color='Custom', text_color=[1,1,1,1],
                                   size_hint_y=None, height=dp(48)))
        sv = MDScrollView(size_hint=(1,1))
        bl = MDBoxLayout(orientation='vertical', spacing=dp(4), size_hint_y=None)
        bl.bind(minimum_height=bl.setter('height'))
        themes = App.get_running_app().theme_manager.themes
        for idx, theme in enumerate(themes):
            bg = theme['background']
            color = bg if theme['type'] == 'color' else [0.2, 0.2, 0.2, 1]
            card = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(48),
                         padding=[dp(8), dp(4)], spacing=dp(8), radius=[dp(8)],
                         elevation=1, md_bg_color=[0.25, 0.25, 0.25, 1],
                         on_release=lambda x, i=idx: self._select_theme(i))
            indicator = MDCard(size_hint_x=None, width=dp(8), radius=[dp(4)],
                              md_bg_color=color, elevation=0)
            card.add_widget(indicator)
            icon_name = 'check-circle' if idx == App.get_running_app().theme_manager.current_theme_index else 'circle-outline'
            card.add_widget(MDIconButton(icon=icon_name, theme_icon_color='Custom',
                                        icon_color=[1,1,1,1], on_release=lambda: None))
            card.add_widget(MDLabel(text=theme['name'], font_size='15sp',
                                   theme_text_color='Custom', text_color=[1,1,1,1],
                                   adaptive_height=True))
            bl.add_widget(card)
        sv.add_widget(bl)
        content.add_widget(sv)
        content.add_widget(make_raised_btn('Закрыть', size_hint=(1, None), height=dp(44),
                                         md_bg_color=Colors.PRIMARY,
                                         on_release=lambda x: theme_popup.dismiss()))
        theme_popup = Popup(title='', content=content, size_hint=(0.5, 0.55),
                           separator_height=0, background_color=[0,0,0,0.3],
                           overlay_color=[0,0,0,0.3])
        self._theme_popup = theme_popup
        theme_popup.open()

    def _select_theme(self, index):
        app = App.get_running_app()
        app.theme_manager.switch_theme(index)
        app.reload_theme()
        if hasattr(self, '_theme_popup'):
            self._theme_popup.dismiss()
        self._snack('Тема применена')

    def show_new_section_form(self, *args):
        content = MDCard(orientation='vertical', spacing=0, padding=0,
                        radius=[dp(16)], elevation=4, md_bg_color=[0.18, 0.18, 0.18, 0.95])
        hdr = MDBoxLayout(orientation='vertical', padding=[dp(16), dp(12)],
                          size_hint_y=None, height=dp(56), md_bg_color=Colors.PRIMARY)
        hdr.add_widget(MDLabel(text='Новый участок', font_size='20sp', bold=True,
                              theme_text_color='Custom', text_color=[1,1,1,1],
                              size_hint_y=None, height=dp(32)))
        content.add_widget(hdr)
        scroll = MDScrollView(size_hint=(1,1))
        layout = MDBoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None,
                             padding=[dp(20), dp(16)])
        layout.bind(minimum_height=layout.setter('height'))
        fields = [('Номер участка', 'section_number_input'), ('Квартал', 'quarter_input'),
                  ('Выдел', 'plot_input'), ('Лесничество', 'forestry_input'),
                  ('Участковое лесничество', 'district_forestry_input')]
        for label_text, attr_name in fields:
            inp = MDTextField(hint_text=label_text, mode='outlined',
                             size_hint_y=None, height=dp(52), font_size='16sp')
            setattr(self, attr_name, inp)
            layout.add_widget(inp)
        scroll.add_widget(layout)
        content.add_widget(scroll)
        btn_row = MDBoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8), padding=[dp(16), dp(8)])
        btn_row.add_widget(make_raised_btn('Сохранить', md_bg_color=Colors.PRIMARY,
                                         on_release=self.save_section))
        btn_row.add_widget(make_raised_btn('Отмена', md_bg_color=Colors.DANGER,
                                         on_release=lambda x: popup.dismiss()))
        content.add_widget(btn_row)
        popup = Popup(title='', content=content, size_hint=(0.45, 0.55),
                     separator_height=0, background_color=[0,0,0,0.3],
                     overlay_color=[0,0,0,0.3])
        self._section_popup = popup
        popup.open()

    def save_section(self, instance):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS sections
            (id INTEGER PRIMARY KEY AUTOINCREMENT, section_number TEXT,
             quarter TEXT, plot TEXT, forestry TEXT, district_forestry TEXT)''')
        cursor.execute('INSERT INTO sections (section_number, quarter, plot, forestry, district_forestry) VALUES (?,?,?,?,?)',
                      (getattr(self, 'section_number_input', MDTextField()).text,
                       getattr(self, 'quarter_input', MDTextField()).text,
                       getattr(self, 'plot_input', MDTextField()).text,
                       getattr(self, 'forestry_input', MDTextField()).text,
                       getattr(self, 'district_forestry_input', MDTextField()).text))
        conn.commit()
        conn.close()
        self._snack('Участок сохранён')
        if hasattr(self, '_section_popup'):
            self._section_popup.dismiss()

    def show_load_section(self, instance):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        try:
            cursor.execute('SELECT id, section_number, quarter, plot, forestry FROM sections ORDER BY id DESC')
            sections = cursor.fetchall()
        except Exception:
            sections = []
        conn.close()
        if not sections:
            self._snack('Нет сохранённых участков')
            return
        content = MDBoxLayout(orientation='vertical', spacing=dp(8), padding=dp(16),
                              md_bg_color=[0.15, 0.15, 0.15, 1])
        content.add_widget(MDLabel(text='Выберите участок', font_style='Title', role='large', bold=True,
                                   theme_text_color='Custom', text_color=[1,1,1,1],
                                   size_hint_y=None, height=dp(48)))
        sv = MDScrollView(size_hint=(1,1))
        bl = MDBoxLayout(orientation='vertical', spacing=dp(4), size_hint_y=None)
        bl.bind(minimum_height=bl.setter('height'))
        for sec in sections:
            text = f'Участок {sec[1] or "?"}  |  Кв:{sec[2] or "?"}  |  Выд:{sec[3] or "?"}'
            card = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(48),
                         padding=[dp(12), dp(4)], spacing=dp(8), radius=[dp(8)],
                         elevation=1, md_bg_color=[0.25, 0.25, 0.25, 1],
                         on_release=lambda x, sid=sec[0]: self._load_section(sid))
            card.add_widget(MDIconButton(icon='map-marker', theme_icon_color='Custom',
                                        icon_color=[1,1,1,1], on_release=lambda: None))
            card.add_widget(MDLabel(text=text, adaptive_height=True,
                                   theme_text_color='Custom', text_color=[1,1,1,1]))
            bl.add_widget(card)
        sv.add_widget(bl)
        content.add_widget(sv)
        content.add_widget(make_raised_btn('Закрыть', size_hint=(1, None), height=dp(44),
                                         md_bg_color=Colors.PRIMARY,
                                         on_release=lambda x: popup.dismiss()))
        popup = Popup(title='', content=content, size_hint=(0.5, 0.55),
                     separator_height=0, background_color=[0,0,0,0.3],
                     overlay_color=[0,0,0,0.3])
        self._load_popup = popup
        popup.open()

    def _load_section(self, section_id):
        table = App.get_running_app().root.get_screen('table')
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('SELECT section_number FROM sections WHERE id=?', (section_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            table.current_section = result[0]
            table.update_section_label()
            self._snack(f'Участок {result[0]} загружен')
        App.get_running_app().root.current = 'table'
        if hasattr(self, '_load_popup'):
            self._load_popup.dismiss()

    def _ensure_molodniki_sections_columns(self, cursor):
        existing = [row[1] for row in cursor.execute('PRAGMA table_info(molodniki_sections)').fetchall()]
        extra = [
            'quarter TEXT', 'plot TEXT', 'forestry TEXT', 'district_forestry TEXT',
            'radius REAL DEFAULT 5.64', 'plot_area TEXT', 'forest_type TEXT',
            'care_queue TEXT', 'characteristics TEXT', 'care_date TEXT',
            'technology TEXT', 'forest_purpose TEXT',
        ]
        for col_def in extra:
            col_name = col_def.split()[0]
            if col_name not in existing:
                cursor.execute(f'ALTER TABLE molodniki_sections ADD COLUMN {col_def}')

    def show_new_molodniki_section(self, *args):
        content = MDCard(orientation='vertical', spacing=0, padding=0,
                        radius=[dp(16)], elevation=4, md_bg_color=[0.18, 0.18, 0.18, 0.95])
        hdr = MDBoxLayout(orientation='vertical', padding=[dp(16), dp(12)],
                          size_hint_y=None, height=dp(56), md_bg_color=Colors.SECONDARY)
        hdr.add_widget(MDLabel(text='Новый участок молодняков', font_size='20sp', bold=True,
                              theme_text_color='Custom', text_color=[1,1,1,1],
                              size_hint_y=None, height=dp(32)))
        content.add_widget(hdr)
        scroll = MDScrollView(size_hint=(1,1))
        layout = MDBoxLayout(orientation='vertical', spacing=dp(12), size_hint_y=None,
                             padding=[dp(20), dp(16)], adaptive_height=True)
        layout.bind(minimum_height=layout.setter('height'))

        layout.add_widget(MDLabel(text='[b]Адресные данные[/b]', markup=True,
                                 size_hint_y=None, height=dp(28),
                                 theme_text_color='Custom', text_color=[1,1,1,1]))
        address_fields = [
            ('Номер участка', 'mol_section_input'),
            ('Квартал', 'mol_quarter_input'),
            ('Выдел', 'mol_plot_input'),
            ('Лесничество', 'mol_forestry_input'),
            ('Участковое лесничество', 'mol_district_forestry_input'),
            ('Радиус площадки, м', 'mol_radius_input'),
            ('Площадь участка, га', 'mol_plot_area_input'),
            ('Тип леса', 'mol_forest_type_input'),
        ]
        for label_text, attr_name in address_fields:
            inp = MDTextField(hint_text=label_text, mode='outlined',
                             size_hint_y=None, height=dp(52), font_size='16sp')
            if attr_name == 'mol_radius_input':
                inp.text = '5.64'
            setattr(self, attr_name, inp)
            layout.add_widget(inp)

        layout.add_widget(MDLabel(text='[b]Таксационные характеристики[/b]', markup=True,
                                 size_hint_y=None, height=dp(28),
                                 theme_text_color='Custom', text_color=[1,1,1,1]))
        self.mol_care_queue_input = MDTextField(
            hint_text='Очередь рубки (первая/вторая/третья)', mode='outlined',
            size_hint_y=None, height=dp(52), font_size='16sp')
        layout.add_widget(self.mol_care_queue_input)

        layout.add_widget(MDLabel(text='Характеристика деревьев:', font_style='Body', role='small',
                                 size_hint_y=None, height=dp(20),
                                 theme_text_color='Custom', text_color=[1,1,1,1]))
        from kivy.uix.textinput import TextInput
        self.mol_characteristics_input = TextInput(
            hint_text='Лучшие: ..., Вспомогательные: ..., Нежелательные: ...',
            size_hint_y=None, height=dp(72), font_size='15sp',
            background_color=[0.22, 0.22, 0.22, 1], foreground_color=[1,1,1,1],
            hint_text_color=[0.5,0.5,0.5,1])
        layout.add_widget(self.mol_characteristics_input)

        self.mol_care_date_input = MDTextField(
            hint_text='Дата рубки (напр. 2025-06)', mode='outlined',
            size_hint_y=None, height=dp(52), font_size='16sp')
        layout.add_widget(self.mol_care_date_input)

        layout.add_widget(MDLabel(text='Технология ухода:', font_style='Body', role='small',
                                 size_hint_y=None, height=dp(20),
                                 theme_text_color='Custom', text_color=[1,1,1,1]))
        self.mol_technology_input = TextInput(
            hint_text='Описание технологии',
            size_hint_y=None, height=dp(72), font_size='15sp',
            background_color=[0.22, 0.22, 0.22, 1], foreground_color=[1,1,1,1],
            hint_text_color=[0.5,0.5,0.5,1])
        layout.add_widget(self.mol_technology_input)

        self.mol_forest_purpose_input = MDTextField(
            hint_text='Назначение лесов (Эксплуатационные/Защитные)', mode='outlined',
            size_hint_y=None, height=dp(52), font_size='16sp')
        layout.add_widget(self.mol_forest_purpose_input)

        scroll.add_widget(layout)
        content.add_widget(scroll)
        btn_row = MDBoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8), padding=[dp(16), dp(8)])
        btn_row.add_widget(make_raised_btn('Сохранить', md_bg_color=Colors.PRIMARY,
                                         on_release=self.save_molodniki_section))
        btn_row.add_widget(make_raised_btn('Отмена', md_bg_color=Colors.DANGER,
                                         on_release=lambda x: popup.dismiss()))
        content.add_widget(btn_row)
        popup = Popup(title='', content=content, size_hint=(0.55, 0.88),
                     separator_height=0, background_color=[0,0,0,0.3],
                     overlay_color=[0,0,0,0.3])
        self._molodniki_popup = popup
        popup.open()

    def save_molodniki_section(self, instance):
        section_number = getattr(self, 'mol_section_input', MDTextField()).text.strip()
        if not section_number:
            self._snack('Введите номер участка')
            return
        quarter = getattr(self, 'mol_quarter_input', MDTextField()).text.strip()
        plot = getattr(self, 'mol_plot_input', MDTextField()).text.strip()
        forestry = getattr(self, 'mol_forestry_input', MDTextField()).text.strip()
        district_forestry = getattr(self, 'mol_district_forestry_input', MDTextField()).text.strip()
        radius = getattr(self, 'mol_radius_input', MDTextField()).text.strip() or '5.64'
        plot_area = getattr(self, 'mol_plot_area_input', MDTextField()).text.strip()
        forest_type = getattr(self, 'mol_forest_type_input', MDTextField()).text.strip()
        care_queue = getattr(self, 'mol_care_queue_input', MDTextField()).text.strip()
        characteristics = getattr(self, 'mol_characteristics_input', None)
        characteristics = characteristics.text.strip() if characteristics else ''
        care_date = getattr(self, 'mol_care_date_input', MDTextField()).text.strip()
        technology = getattr(self, 'mol_technology_input', None)
        technology = technology.text.strip() if technology else ''
        forest_purpose = getattr(self, 'mol_forest_purpose_input', MDTextField()).text.strip()
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS molodniki_sections
            (id INTEGER PRIMARY KEY AUTOINCREMENT, section_number TEXT,
             quarter TEXT, plot TEXT, forestry TEXT, district_forestry TEXT,
             radius REAL DEFAULT 5.64, plot_area TEXT, forest_type TEXT)''')
        self._ensure_molodniki_sections_columns(cursor)
        cursor.execute('''INSERT OR REPLACE INTO molodniki_sections
            (id, section_number, quarter, plot, forestry, district_forestry,
             radius, plot_area, forest_type, care_queue, characteristics,
             care_date, technology, forest_purpose)
            VALUES ((SELECT id FROM molodniki_sections WHERE section_number=?),
             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                       (section_number, section_number, quarter, plot, forestry,
                        district_forestry, radius, plot_area, forest_type,
                        care_queue, characteristics, care_date, technology, forest_purpose))
        conn.commit()
        conn.close()
        screen = App.get_running_app().root.get_screen('molodniki')
        screen.current_section = section_number
        screen.current_quarter = quarter
        screen.current_plot = plot
        screen.current_forestry = forestry
        screen.current_district_forestry = district_forestry
        screen.current_radius = radius
        screen.plot_area_input = plot_area
        screen.project_data['address']['quarter'] = quarter
        screen.project_data['address']['plot'] = plot
        screen.project_data['address']['forestry'] = forestry
        screen.project_data['address']['district_forestry'] = district_forestry
        screen.project_data['address']['radius'] = radius
        screen.project_data['address']['plot_area'] = plot_area
        if care_queue:
            screen.project_data['details']['care_queue'] = care_queue
        if characteristics:
            screen.project_data['details']['characteristics'] = characteristics
        if care_date:
            screen.project_data['details']['care_date'] = care_date
        if technology:
            screen.project_data['details']['technology'] = technology
        if forest_purpose:
            screen.project_data['details']['forest_purpose'] = forest_purpose
        self._snack(f'Участок {section_number} создан')
        if hasattr(self, '_molodniki_popup'):
            self._molodniki_popup.dismiss()
        App.get_running_app().root.current = 'molodniki'

    def show_load_molodniki_section(self, instance):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        try:
            self._ensure_molodniki_sections_columns(cursor)
            cursor.execute('''SELECT id, section_number, quarter, plot, forestry, radius
                              FROM molodniki_sections ORDER BY id DESC''')
            sections = cursor.fetchall()
        except Exception:
            sections = []
        conn.close()
        if not sections:
            self._snack('Нет сохранённых участков молодняков')
            return
        content = MDBoxLayout(orientation='vertical', spacing=dp(8), padding=dp(16),
                              md_bg_color=[0.15, 0.15, 0.15, 1])
        content.add_widget(MDLabel(text='Выберите участок молодняков', font_style='Title', role='large', bold=True,
                                   theme_text_color='Custom', text_color=[1,1,1,1],
                                   size_hint_y=None, height=dp(48)))
        sv = MDScrollView(size_hint=(1,1))
        bl = MDBoxLayout(orientation='vertical', spacing=dp(4), size_hint_y=None)
        bl.bind(minimum_height=bl.setter('height'))
        for sec in sections:
            sid, snum, q, p, f, r = sec[0], sec[1] or '', sec[2] or '', sec[3] or '', sec[4] or '', sec[5] or ''
            label_text = f'Уч.{snum}  |  Кв:{q}  |  Выд:{p}'
            if f:
                label_text += f'  |  {f}'
            if r:
                label_text += f'  |  R={r}м'
            card = MDCard(orientation='horizontal', size_hint=(1, None), height=dp(52),
                         padding=[dp(12), dp(4)], spacing=dp(8), radius=[dp(8)],
                         elevation=1, md_bg_color=[0.25, 0.25, 0.25, 1],
                         on_release=lambda x, sid=sid: self._load_molodniki_section(sid))
            card.add_widget(MDIconButton(icon='seedling', theme_icon_color='Custom',
                                        icon_color=[1,1,1,1], on_release=lambda: None))
            card.add_widget(MDLabel(text=label_text, adaptive_height=True,
                                   theme_text_color='Custom', text_color=[1,1,1,1]))
            bl.add_widget(card)
        sv.add_widget(bl)
        content.add_widget(sv)
        content.add_widget(make_raised_btn('Закрыть', size_hint=(1, None), height=dp(44),
                                         md_bg_color=Colors.SECONDARY,
                                         on_release=lambda x: popup.dismiss()))
        popup = Popup(title='', content=content, size_hint=(0.5, 0.55),
                     separator_height=0, background_color=[0,0,0,0.3],
                     overlay_color=[0,0,0,0.3])
        popup.open()

    def show_load_molodniki_json(self, *args):
        """Загрузить JSON файл через штатный загрузчик молодняков"""
        screen = App.get_running_app().root.get_screen('molodniki')
        App.get_running_app().root.current = 'molodniki'
        Clock.schedule_once(lambda dt: screen.load_section_popup(), 0.2)

    def show_load_molodniki_excel(self, *args):
        """Загрузить данные из Excel файла"""
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title='Выберите Excel файл',
            filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
            initialdir='reports',
        )
        root.destroy()
        if not file_path:
            return
        screen = App.get_running_app().root.get_screen('molodniki')
        try:
            import pandas as pd
            df = pd.read_excel(file_path)
            screen.page_data.clear()
            section_name = os.path.splitext(os.path.basename(file_path))[0]
            for page_num in range(0, len(df), screen.rows_per_page):
                page = page_num // screen.rows_per_page
                page_data = df.iloc[page_num:page_num+screen.rows_per_page].values.tolist()
                for row in page_data:
                    while len(row) < 29:
                        row.append('')
                screen.page_data[page] = page_data
            screen.current_section = section_name
            screen.current_page = 0
            screen.load_page_data()
            self._snack(f'Данные загружены из {os.path.basename(file_path)}')
            App.get_running_app().root.current = 'molodniki'
        except Exception as e:
            self._snack(f'Ошибка загрузки Excel: {e}')

    def _load_molodniki_section(self, section_id):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()
        try:
            self._ensure_molodniki_sections_columns(cursor)
            cursor.execute('''SELECT section_number, quarter, plot, forestry,
                              district_forestry, radius, plot_area, forest_type,
                              care_queue, characteristics, care_date, technology,
                              forest_purpose
                              FROM molodniki_sections WHERE id=?''', (section_id,))
            row = cursor.fetchone()
        except Exception:
            row = None
        conn.close()
        if not row:
            self._snack('Ошибка загрузки участка')
            return
        snum, q, p, f, df, r, pa, ft = (row[0] or '', row[1] or '', row[2] or '',
                                          row[3] or '', row[4] or '', str(row[5] or '5.64'),
                                          row[6] or '', row[7] or '')
        care_queue = row[8] or '' if len(row) > 8 else ''
        characteristics = row[9] or '' if len(row) > 9 else ''
        care_date = row[10] or '' if len(row) > 10 else ''
        technology = row[11] or '' if len(row) > 11 else ''
        forest_purpose = row[12] or '' if len(row) > 12 else ''
        screen = App.get_running_app().root.get_screen('molodniki')
        screen.current_section = snum
        screen.current_quarter = q
        screen.current_plot = p
        screen.current_forestry = f
        screen.current_district_forestry = df
        screen.current_radius = r
        screen.plot_area_input = pa
        screen.project_data['address']['quarter'] = q
        screen.project_data['address']['plot'] = p
        screen.project_data['address']['forestry'] = f
        screen.project_data['address']['district_forestry'] = df
        screen.project_data['address']['radius'] = r
        screen.project_data['address']['plot_area'] = pa
        if care_queue:
            screen.project_data['details']['care_queue'] = care_queue
        if characteristics:
            screen.project_data['details']['characteristics'] = characteristics
        if care_date:
            screen.project_data['details']['care_date'] = care_date
        if technology:
            screen.project_data['details']['technology'] = technology
        if forest_purpose:
            screen.project_data['details']['forest_purpose'] = forest_purpose
        screen.load_existing_data()
        App.get_running_app().root.current = 'molodniki'
        self._snack(f'Участок {snum} загружен')

    def confirm_exit(self, instance):
        content = MDBoxLayout(orientation='vertical', spacing=15, padding=15, adaptive_height=True,
                             md_bg_color=[0.18, 0.18, 0.18, 0.95])
        content.add_widget(MDLabel(text='Завершить работу?', font_style='Title', role='medium',
                                   halign='center', size_hint_y=None, height=40,
                                   theme_text_color='Custom', text_color=[1,1,1,1]))
        dialog = MDDialog(
            MDDialogHeadlineText(text='Подтверждение'),
            MDDialogContentContainer(content),
            MDDialogButtonContainer(
                make_raised_btn('Выход', md_bg_color=Colors.DANGER,
                               on_release=lambda x: (dialog.dismiss(), App.get_running_app().stop())),
                make_outlined_btn('Отмена', on_release=lambda x: dialog.dismiss()),
                spacing='8dp',
            ),
        )
        dialog.open()

    def _snack(self, message, duration=2.5):
        sn = MDSnackbar(duration=duration)
        sn.add_widget(MDSnackbarText(text=message))
        sn.open()


class ForestApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.theme_manager = ThemeManager()
        theme = self.theme_manager.current_theme
        if theme['type'] == 'color':
            self.theme_cls.primary_palette = theme.get('kivymd_palette', 'Green')
            self.theme_cls.theme_style = theme.get('kivymd_style', 'Light')
            Window.clearcolor = (theme['background'][0], theme['background'][1], theme['background'][2], 1)
        else:
            self.theme_cls.primary_palette = 'Green'
            self.theme_cls.theme_style = 'Light'
            Window.clearcolor = (0.15, 0.18, 0.2, 1)
        self.title = 'Фанаты Пихты — Лесное таксационное приложение'

    def build(self):
        Window.bind(on_key_down=self.on_keyboard)
        sm = LazyScreenManager()
        sm.add_widget(MainMenu())
        sm.add_widget(TableScreen(name='table'))
        sm.add_widget(ExtendedMolodnikiTableScreen(name='molodniki'))
        sm.current = 'main'
        return sm

    def on_keyboard(self, window, key, scancode, codepoint, modifier):
        if key == 27:
            current = self.root.current
            back_map = {
                'table': 'main', 'molodniki': 'main', 'mdol_dashboard': 'main',
                'references': 'main', 'taxation_ai': 'main',
            }
            if current in back_map and back_map[current]:
                self.root.current = back_map[current]
            return True
        return False

    def reload_theme(self):
        theme = self.theme_manager.current_theme
        if theme['type'] == 'color':
            self.theme_cls.primary_palette = theme.get('kivymd_palette', 'Green')
            self.theme_cls.theme_style = theme.get('kivymd_style', 'Light')
            Window.clearcolor = (theme['background'][0], theme['background'][1], theme['background'][2], 1)
        for screen in self.root.screens:
            if hasattr(screen, 'create_ui'):
                screen.create_ui()
            if hasattr(screen, 'update_bg'):
                screen.update_bg()

    def on_pause(self):
        return True

    def on_resume(self):
        pass


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
    try:
        ForestApp().run()
    except Exception as e:
        logging.exception('Fatal error: %s', e)
