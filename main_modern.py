"""
ForestApp - Современное приложение для учёта лесных данных
KivyMD (Material Design) + собственная система стилей
"""
from kivy.core.window import Window
from kivy.config import Config
from kivy.metrics import dp
from kivy.core.text import LabelBase
from kivy.app import App
from kivy.clock import Clock
import os
import json
import sqlite3
import glob
from kivy.uix.textinput import TextInput

Config.set('graphics', 'width', '480')
Config.set('graphics', 'height', '854')
Config.set('graphics', 'resizable', True)
Config.set('input', 'mouse', 'mouse,multitouch_on_demand')

LabelBase.register(name='Roboto',
                   fn_regular='fonts/Roboto-Medium.ttf',
                   fn_bold='fonts/Roboto-Bold.ttf')

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
from kivymd.uix.appbar import MDTopAppBar, MDTopAppBarLeadingButtonContainer, MDTopAppBarTrailingButtonContainer, MDTopAppBarTitle, MDActionTopAppBarButton
from kivymd.uix.navigationdrawer import MDNavigationLayout, MDNavigationDrawer
from kivymd.uix.list import MDListItem, MDListItemHeadlineText, MDListItemLeadingIcon
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.selectioncontrol import MDSwitch
from kivymd.uix.menu import MDDropdownMenu


from screens.dashboard_screen import DashboardScreen
from screens.map_screen import MapScreen
from molodniki_extended import ExtendedMolodnikiTableScreen
from new_taxation_menu import TaxationPopup
from ui_styles import Colors, Spacing, Fonts
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


class MainMenu(MDScreen):
    """Главное меню приложения"""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.theme_manager = ThemeManager()
        Clock.schedule_once(lambda dt: self.create_ui(), 0)

    def create_ui(self):
        self.clear_widgets()

        nav_layout = MDNavigationLayout()

        inner_sm = MDScreenManager()
        main_screen = MDScreen()
        main_layout = MDBoxLayout(orientation='vertical')

        toolbar = MDTopAppBarOld(
            title='Фанаты Пихты',
            elevation=2,
            md_bg_color=Colors.PRIMARY,
            left_action_items=[['menu', lambda x: self.open_nav_drawer()]],
            right_action_items=[['palette', lambda x: self.show_theme_dialog()]],
        )
        main_layout.add_widget(toolbar)

        scroll = MDScrollView()
        content = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            spacing=Spacing.MD,
            padding=[Spacing.LG, Spacing.SM],
            md_bg_color=[0.12, 0.12, 0.12, 1],
        )
        content.bind(minimum_height=content.setter('height'))

        welcome_card = self._create_welcome_card()
        content.add_widget(welcome_card)

        content.add_widget(MDCard(
            size_hint_y=None, height=dp(1),
            md_bg_color=[0,0,0,0.08],
            radius=[0]
        ))

        section_title = MDLabel(
            text='ОСНОВНЫЕ РАЗДЕЛЫ',
            font_style='Label', role='small',
            size_hint_y=None,
            height=dp(24),
            theme_text_color='Custom', text_color=[1,1,1,0.7],
            padding=[Spacing.SM, 0]
        )
        content.add_widget(section_title)

        sections = [
            {'icon': 'seedling', 'title': 'РУМ (Молодняки)',
             'desc': 'Регулируемый уход за молодняками — ввод и расчёт',
             'color': Colors.PRIMARY, 'action': 'molodniki'},
            {'icon': 'calculator', 'title': 'Таксация',
             'desc': 'Расчёт таксационных показателей молодняков',
             'color': Colors.ACCENT, 'action': 'taxation'},
        ]
        for section in sections:
            card = self._create_section_card(section)
            content.add_widget(card)

        content.add_widget(MDCard(
            size_hint_y=None, height=dp(1),
            md_bg_color=[0,0,0,0.08],
            radius=[0]
        ))

        service_title = MDLabel(
            text='СЛУЖЕБНЫЕ',
            font_style='Label', role='small',
            size_hint_y=None,
            height=dp(24),
            theme_text_color='Custom', text_color=[1,1,1,0.7],
            padding=[Spacing.SM, 0]
        )
        content.add_widget(service_title)

        service_actions = [
            {'icon': 'chart-box-outline', 'title': 'Дашборд',
             'desc': 'Графики и статистика по данным',
             'action': 'dashboard'},
            {'icon': 'map', 'title': 'Карта',
             'desc': 'Просмотр участков на карте',
             'action': 'map'},
            {'icon': 'calculator-variant', 'title': 'Калькулятор',
             'desc': 'Расчёт площади, густоты, запаса',
             'action': 'calculator'},
            {'icon': 'file-document-outline', 'title': 'Отчёты',
             'desc': 'Просмотр сохранённых отчетов',
             'action': 'reports'},
            {'icon': 'database-outline', 'title': 'База данных',
             'desc': 'Управление данными участков',
             'action': 'database'},
            {'icon': 'camera', 'title': 'Фото',
             'desc': 'Фотофиксация участков',
             'action': 'photos'},
            {'icon': 'backup-restore', 'title': 'Бекап',
             'desc': 'Резервное копирование и восстановление',
             'action': 'backup'},
            {'icon': 'magnify', 'title': 'Поиск',
             'desc': 'Поиск по всем данным',
             'action': 'search'},
            {'icon': 'compare', 'title': 'Сравнение',
             'desc': 'Сравнить два участка',
             'action': 'compare'},
            {'icon': 'information-outline', 'title': 'О программе',
             'desc': 'Версия 2.0',
             'action': 'about'},
        ]
        for item in service_actions:
            card = self._create_service_card(item)
            content.add_widget(card)

        scroll.add_widget(content)
        main_layout.add_widget(scroll)

        main_screen.add_widget(main_layout)
        inner_sm.add_widget(main_screen)

        nav_layout.add_widget(inner_sm)

        self.nav_drawer = self._create_nav_drawer()
        nav_layout.add_widget(self.nav_drawer)

        self.add_widget(nav_layout)

    def _create_welcome_card(self):
        card = MDCard(
            size_hint_y=None,
            height=dp(140),
            orientation='vertical',
            padding=Spacing.LG,
            radius=[Spacing.RADIUS_LG],
            elevation=2,
            md_bg_color=Colors.PRIMARY_LIGHT
        )
        title = MDLabel(
            text='🌲 Фанаты Пихты',
            font_style='Headline', role='medium',
            theme_text_color='Custom',
            text_color=[1,1,1,1],
            bold=True,
            size_hint_y=None,
            height=dp(48)
        )
        subtitle = MDLabel(
            text='Лесное таксационное приложение\nУчёт, анализ и управление лесными данными',
            font_style='Body', role='medium',
            theme_text_color='Custom',
            text_color=[1, 1, 1, 0.85],
            size_hint_y=None,
            height=dp(56)
        )
        card.add_widget(title)
        card.add_widget(subtitle)
        return card

    def _create_section_card(self, section):
        card = MDCard(
            size_hint_y=None,
            height=dp(88),
            orientation='horizontal',
            padding=Spacing.MD,
            spacing=Spacing.LG,
            radius=[Spacing.RADIUS_LG],
            elevation=1,
            md_bg_color=[0.18, 0.18, 0.18, 0.95],
            ripple_behavior=True,
            on_release=lambda x, s=section.get('screen'), a=section.get('action'): self._handle_action(s, a)
        )

        icon_card = MDCard(
            size_hint_x=None,
            width=dp(56),
            radius=[Spacing.RADIUS_MD],
            md_bg_color=section['color'],
            padding=[0, 0, 0, 0],
            orientation='vertical'
        )
        icon_btn = MDIconButton(
            icon=section['icon'],
            theme_icon_color='Custom',
            icon_color=[1,1,1,1],
            font_size=dp(28),
            pos_hint={'center_x': 0.5, 'center_y': 0.5},
            size_hint=(None, None),
            size=(dp(56), dp(56))
        )
        icon_card.add_widget(icon_btn)
        card.add_widget(icon_card)

        text_layout = MDBoxLayout(
            orientation='vertical',
            spacing=dp(2),
            adaptive_height=True
        )
        title = MDLabel(
            text=section['title'],
            font_style='Title', role='small',
            bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            adaptive_height=True
        )
        desc = MDLabel(
            text=section.get('desc', ''),
            font_style='Body', role='small',
            theme_text_color='Custom', text_color=[1,1,1,0.7],
            adaptive_height=True
        )
        text_layout.add_widget(title)
        text_layout.add_widget(desc)
        card.add_widget(text_layout)

        arrow = MDLabel(
            text='›',
            font_style='Headline', role='large',
            theme_text_color='Custom', text_color=[1,1,1,0.5],
            size_hint_x=None,
            width=dp(24),
            halign='right',
            valign='middle'
        )
        card.add_widget(arrow)

        return card

    def _create_service_card(self, item):
        card = MDCard(
            size_hint_y=None,
            height=dp(72),
            orientation='horizontal',
            padding=Spacing.MD,
            spacing=Spacing.MD,
            radius=[Spacing.RADIUS_LG],
            elevation=1,
            md_bg_color=[0.18, 0.18, 0.18, 0.95],
            ripple_behavior=True,
            on_release=lambda x, a=item['action']: self._handle_action(None, a)
        )
        icon = MDIconButton(
            icon=item['icon'],
            theme_icon_color='Custom',
            icon_color=[1,1,1,1],
            on_release=lambda: None,
        )
        text_layout = MDBoxLayout(
            orientation='vertical',
            spacing=dp(2),
            adaptive_height=True,
            padding=[Spacing.SM, 0, 0, 0]
        )
        title = MDLabel(
            text=item['title'],
            font_style='Body', role='medium',
            bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            adaptive_height=True
        )
        desc = MDLabel(
            text=item.get('desc', ''),
            font_style='Body', role='small',
            theme_text_color='Custom', text_color=[1,1,1,0.7],
            adaptive_height=True
        )
        text_layout.add_widget(title)
        text_layout.add_widget(desc)

        card.add_widget(icon)
        card.add_widget(text_layout)
        return card

    def _create_nav_drawer(self):
        drawer = MDNavigationDrawer(id='nav_drawer')

        # Шапка — центрированная
        header = MDBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=dp(170),
            padding=[Spacing.LG, Spacing.XL, Spacing.LG, Spacing.LG],
            md_bg_color=Colors.PRIMARY,
        )
        icon_label = MDLabel(
            text='🌲',
            font_style='Headline', role='large',
            halign='center',
            theme_text_color='Custom',
            text_color=Colors.TEXT_ON_PRIMARY,
            size_hint_y=None, height=dp(44),
        )
        header.add_widget(icon_label)
        header.add_widget(MDLabel(
            text='Фанаты Пихты',
            font_style='Title', role='large',
            halign='center',
            theme_text_color='Custom',
            text_color=Colors.TEXT_ON_PRIMARY,
            bold=True,
            size_hint_y=None, height=dp(36),
        ))
        header.add_widget(MDLabel(
            text='Лесное таксационное приложение',
            halign='center',
            font_style='Body', role='small',
            theme_text_color='Custom',
            text_color=[1, 1, 1, 0.7],
            size_hint_y=None, height=dp(24),
        ))
        drawer.add_widget(header)

        def make_drawer_item(icon, text, screen=None, action=None):
            def on_click(x):
                self._navigate_drawer(screen)
                if action:
                    self._handle_action(None, action)
            item = MDListItem(
                on_release=on_click,
            )
            item.add_widget(MDListItemLeadingIcon(icon=icon))
            item.add_widget(MDListItemHeadlineText(text=text))
            return item

        drawer.add_widget(make_drawer_item('home', 'Главная', 'main'))

        separator1 = MDBoxLayout(size_hint_y=None, height=dp(1), md_bg_color=[1,1,1,0.08])
        drawer.add_widget(separator1)

        drawer.add_widget(MDLabel(
            text='  ОСНОВНЫЕ', font_style='Label', role='small',
            theme_text_color='Custom', text_color=Colors.ACCENT,
            size_hint_y=None, height=dp(28),
        ))
        drawer.add_widget(make_drawer_item('seedling', 'РУМ (Молодняки)', action='molodniki'))
        drawer.add_widget(make_drawer_item('calculator', 'Таксация', action='taxation'))

        separator2 = MDBoxLayout(size_hint_y=None, height=dp(1), md_bg_color=[1,1,1,0.08])
        drawer.add_widget(separator2)

        drawer.add_widget(MDLabel(
            text='  СЛУЖЕБНЫЕ', font_style='Label', role='small',
            theme_text_color='Custom', text_color=Colors.ACCENT,
            size_hint_y=None, height=dp(28),
        ))
        drawer.add_widget(make_drawer_item('chart-box-outline', 'Дашборд', 'dashboard'))
        drawer.add_widget(make_drawer_item('map', 'Карта', 'map'))
        drawer.add_widget(make_drawer_item('calculator-variant', 'Калькулятор', action='calculator'))
        drawer.add_widget(make_drawer_item('file-document-outline', 'Отчёты', action='reports'))
        drawer.add_widget(make_drawer_item('database-outline', 'База данных', action='database'))
        drawer.add_widget(make_drawer_item('camera', 'Фото', action='photos'))
        drawer.add_widget(make_drawer_item('backup-restore', 'Бекап', action='backup'))
        drawer.add_widget(make_drawer_item('magnify', 'Поиск', action='search'))
        drawer.add_widget(make_drawer_item('compare', 'Сравнение', action='compare'))
        drawer.add_widget(make_drawer_item('information-outline', 'О программе', action='about'))

        return drawer

    def _navigate_drawer(self, screen):
        if hasattr(self, 'nav_drawer') and self.nav_drawer:
            self.nav_drawer.set_state('close')
        if screen:
            App.get_running_app().root.current = screen

    def open_nav_drawer(self):
        if hasattr(self, 'nav_drawer') and self.nav_drawer:
            self.nav_drawer.set_state('toggle')

    def _handle_action(self, screen, action):
        if screen:
            App.get_running_app().root.current = screen
        elif action == 'molodniki':
            self.show_molodniki_section_dialog()
        elif action == 'taxation':
            TaxationPopup().open()
        elif action == 'about':
            self._show_about_dialog()
        elif action == 'theme':
            self.show_theme_dialog()
        elif action == 'database':
            self.show_section_dialog()
        elif action == 'reports':
            self.show_reports_dialog()
        elif action == 'dashboard':
            App.get_running_app().root.current = 'dashboard'
        elif action == 'map':
            App.get_running_app().root.current = 'map'
        elif action == 'calculator':
            self.show_calculator_dialog()
        elif action == 'search':
            self.show_search_dialog()
        elif action == 'backup':
            self.show_backup_dialog()
        elif action == 'compare':
            self.show_comparison_dialog()
        elif action == 'photos':
            self.show_photo_dialog()

    def show_success(self, message):
        snack = MDSnackbar(duration=2.5)
        snack.add_widget(MDSnackbarText(text=f'✅ {message}'))
        snack.open()

    def show_error(self, message):
        snack = MDSnackbar(duration=3.5)
        snack.add_widget(MDSnackbarText(text=f'❌ {message}'))
        snack.open()

    def show_section_dialog(self):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)

        title_label = MDLabel(
            text='Управление участками',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        )
        content.add_widget(title_label)

        self.section_number_input = TextInput(
            hint_text='Номер участка', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.section_number_input)

        self.quarter_input = TextInput(
            hint_text='Квартал', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.quarter_input)

        self.plot_input = TextInput(
            hint_text='Выдел', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.plot_input)

        self.forestry_input = TextInput(
            hint_text='Лесничество', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.forestry_input)

        self.district_forestry_input = TextInput(
            hint_text='Участковое лесничество', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.district_forestry_input)

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        save_btn = make_raised_btn('Сохранить', icon='content-save', size_hint=(0.5, None), height=dp(48),
                                   on_release=lambda x: self.save_section(dialog))
        load_btn = make_outlined_btn('Загрузить', icon='folder-open', size_hint=(0.5, None), height=dp(48),
                                     on_release=lambda x: self.show_load_section_dialog())
        btn_row.add_widget(save_btn)
        btn_row.add_widget(load_btn)
        content.add_widget(btn_row)

        scroll = MDScrollView(size_hint_y=None, height=dp(250))
        list_layout = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT section_number, quarter, plot, forestry, district_forestry FROM sections WHERE section_number IS NOT NULL AND section_number != "" ORDER BY id DESC')
            rows = cursor.fetchall()
            conn.close()
            if rows:
                for row in rows:
                    row_card = MDCard(
                        orientation='horizontal',
                        size_hint_y=None, height=dp(48),
                        padding=[Spacing.MD, 0], spacing=Spacing.MD,
                        md_bg_color=[0.25, 0.25, 0.25, 1],
                        on_release=lambda x, s=row[0]: self.load_saved_section(s),
                    )
                    row_card.add_widget(MDListItemLeadingIcon(icon='folder'))
                    addr_parts = [p for p in (row[0], row[1], row[2]) if p]
                    row_card.add_widget(MDLabel(
                        text=f'{row[0]}  (кв.{row[1]}, выд.{row[2]})' if row[1] or row[2] else row[0],
                        adaptive_height=True, valign='middle',
                    ))
                    list_layout.add_widget(row_card)
        except Exception:
            pass
        scroll.add_widget(list_layout)
        content.add_widget(scroll)

        dialog = MDDialog(
            MDDialogContentContainer(content),
            size_hint=(0.8, None),
        )
        dialog.open()

    def save_section(self, dialog):
        section_number = self.section_number_input.text.strip()
        quarter = self.quarter_input.text.strip()
        plot = self.plot_input.text.strip()
        forestry = self.forestry_input.text.strip()
        district_forestry = self.district_forestry_input.text.strip()
        if not section_number:
            self.show_error('Введите номер участка!')
            return
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('''INSERT OR REPLACE INTO sections
                (section_number, quarter, plot, forestry, district_forestry)
                VALUES (?, ?, ?, ?, ?)''',
                (section_number, quarter, plot, forestry, district_forestry))
            conn.commit()
            conn.close()
            dialog.dismiss()
            molodniki_screen = App.get_running_app().root.get_screen('molodniki')
            molodniki_screen.current_section = section_number
            if hasattr(molodniki_screen, 'update_section_label'):
                molodniki_screen.update_section_label()
            App.get_running_app().root.current = 'molodniki'
            self.show_success(f'Участок {section_number} сохранён!')
        except Exception as e:
            self.show_error(f'Ошибка: {str(e)}')

    def show_load_section_dialog(self):
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT section_number FROM sections WHERE section_number IS NOT NULL AND section_number != "" ORDER BY id DESC')
            sections = cursor.fetchall()
            conn.close()
        except Exception:
            sections = []
        if not sections:
            self.show_error('Нет сохранённых участков!')
            return

        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='Выберите участок', font_style='Title', role='medium', bold=True,
            halign='center', size_hint_y=None, height=dp(36),
        ))
        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        list_layout = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        for section in sections:
            btn_card = MDCard(
                orientation='horizontal', size_hint_y=None, height=dp(48),
                padding=[Spacing.MD, 0], spacing=Spacing.MD,
                md_bg_color=[0.25, 0.25, 0.25, 1],
                on_release=lambda x, s=section[0]: self.load_saved_section(s),
            )
            btn_card.add_widget(MDListItemLeadingIcon(icon='file-excel'))
            btn_card.add_widget(MDLabel(text=section[0], adaptive_height=True, valign='middle'))
            list_layout.add_widget(btn_card)
        scroll.add_widget(list_layout)
        content.add_widget(scroll)
        content.add_widget(make_outlined_btn('Закрыть', size_hint=(1, None), height=dp(48),
                           on_release=lambda x: load_dialog.dismiss()))
        load_dialog = MDDialog(
            MDDialogContentContainer(content),
            size_hint=(0.7, None),
        )
        load_dialog.open()

    def load_saved_section(self, section_number):
        molodniki_screen = App.get_running_app().root.get_screen('molodniki')
        molodniki_screen.current_section = section_number
        if hasattr(molodniki_screen, 'update_section_label'):
            molodniki_screen.update_section_label()
        App.get_running_app().root.current = 'molodniki'
        self.show_success(f'Участок {section_number}')

    def show_photo_dialog(self, section_number=None):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='📸 Фотофиксация', font_style='Title', role='medium', bold=True,
            halign='center', size_hint_y=None, height=dp(36),
        ))

        if not section_number:
            section_number = self.section_number_input.text.strip() if hasattr(self, 'section_number_input') else ''
        if not section_number:
            self.show_error('Сначала укажите номер участка!')
            return

        section_number = str(section_number)
        photo_dir = f'photos/{section_number}'
        os.makedirs(photo_dir, exist_ok=True)

        content.add_widget(MDLabel(
            text=f'Участок: {section_number}', font_size='14sp',
            theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY,
            size_hint_y=None, height=dp(24),
        ))

        scroll = MDScrollView(size_hint_y=None, height=dp(250))
        photo_grid = MDGridLayout(cols=2, spacing=Spacing.SM, size_hint_y=None, adaptive_height=True)
        scroll.add_widget(photo_grid)
        content.add_widget(scroll)

        def refresh_photos():
            photo_grid.clear_widgets()
            if os.path.exists(photo_dir):
                photos = sorted([f for f in os.listdir(photo_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))])
                if not photos:
                    photo_grid.add_widget(MDLabel(
                        text='Нет фотографий', halign='center',
                        theme_text_color='Hint', size_hint_y=None, height=dp(48),
                    ))
                for photo in photos:
                    from kivy.uix.image import Image as KivyImage
                    from kivy.core.image import Image as CoreImage
                    try:
                        photo_card = MDCard(
                            orientation='vertical', size_hint_y=None,
                            padding=dp(2), radius=[dp(6)],
                            md_bg_color=[0.2, 0.2, 0.2, 1],
                        )
                        img = KivyImage(
                            source=os.path.join(photo_dir, photo),
                            size_hint_y=None, height=dp(100),
                            allow_stretch=True, keep_ratio=True,
                        )
                        photo_card.add_widget(img)
                        photo_card.add_widget(MDLabel(
                            text=photo, font_size='9sp', halign='center',
                            theme_text_color='Custom', text_color=Colors.TEXT_DIM,
                            size_hint_y=None, height=dp(18),
                        ))
                        photo_grid.add_widget(photo_card)
                    except Exception:
                        pass

        refresh_photos()

        def take_photo(inst):
            from kivy.uix.filechooser import FileChooserIconView
            from kivy.uix.popup import Popup
            from kivy.uix.boxlayout import BoxLayout
            from shutil import copy2
            import datetime

            fc_content = BoxLayout(orientation='vertical')
            filechooser = FileChooserIconView(
                filters=['*.png', '*.jpg', '*.jpeg', '*.gif'],
                path=os.path.expanduser('~'),
            )
            fc_content.add_widget(filechooser)

            def select(fc):
                if filechooser.selection:
                    src = filechooser.selection[0]
                    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                    ext = os.path.splitext(src)[1]
                    dst = os.path.join(photo_dir, f'photo_{ts}{ext}')
                    copy2(src, dst)
                    fc_popup.dismiss()
                    refresh_photos()
                    self.show_success('Фото добавлено!')

            select_btn = MDButton(style='filled', md_bg_color=Colors.PRIMARY,
                                   size_hint=(1, None), height=dp(48),
                                   on_release=lambda x: select(filechooser))
            select_btn.add_widget(MDButtonText(text='Выбрать'))
            fc_content.add_widget(select_btn)

            fc_popup = Popup(
                title='Выберите фотографию',
                content=fc_content,
                size_hint=(0.9, 0.9),
            )
            fc_popup.open()

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Добавить фото', icon='camera', size_hint=(0.5, None), height=dp(48),
                           on_release=take_photo))
        btn_row.add_widget(make_outlined_btn('Закрыть', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: photo_dialog.dismiss()))
        content.add_widget(btn_row)

        photo_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.85, None))
        photo_dialog.open()

    def show_molodniki_section_dialog(self, instance=None):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='Управление участками молодняков',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))
        self.molodniki_section_input = TextInput(
            hint_text='Введите номер участка молодняков', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(self.molodniki_section_input)

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        save_btn = make_raised_btn('Сохранить', icon='content-save', size_hint=(0.5, None), height=dp(48),
                                   on_release=lambda x: self.save_molodniki_section(molodniki_dialog))
        load_json_btn = make_outlined_btn('Загрузить JSON', icon='file-json', size_hint=(0.5, None), height=dp(48),
                                          on_release=lambda x: self.load_molodniki_json())
        btn_row.add_widget(save_btn)
        btn_row.add_widget(load_json_btn)
        content.add_widget(btn_row)

        content.add_widget(make_outlined_btn('Загрузить из Excel', icon='file-excel', size_hint=(1, None), height=dp(48),
                           on_release=lambda x: self.show_load_molodniki_dialog()))

        molodniki_dialog = MDDialog(
            MDDialogContentContainer(content),
            size_hint=(0.8, None),
        )
        molodniki_dialog.open()

    def save_molodniki_section(self, dialog):
        section_number = self.molodniki_section_input.text.strip()
        if not section_number:
            self.show_error('Введите номер участка!')
            return
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('''INSERT OR REPLACE INTO molodniki_sections (section_number)
                VALUES (?)''', (section_number,))
            conn.commit()
            conn.close()
            dialog.dismiss()
            molodniki_screen = App.get_running_app().root.get_screen('molodniki')
            molodniki_screen.current_section = section_number
            molodniki_screen.update_section_label()
            App.get_running_app().root.current = 'molodniki'
            self.show_success(f'Участок молодняков {section_number} сохранён!')
        except Exception as e:
            self.show_error(f'Ошибка: {str(e)}')

    def load_molodniki_json(self):
        App.get_running_app().root.current = 'molodniki'
        Clock.schedule_once(lambda dt: self._show_molodniki_json_popup(), 0.1)

    def _show_molodniki_json_popup(self):
        molodniki_screen = App.get_running_app().root.get_screen('molodniki')
        if hasattr(molodniki_screen, 'load_section_popup'):
            molodniki_screen.load_section_popup()

    def show_load_molodniki_dialog(self):
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != "" ORDER BY id DESC')
            sections = cursor.fetchall()
            conn.close()
        except Exception:
            sections = []
        if not sections:
            self.show_error('Нет сохранённых участков молодняков!')
            return

        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='Выберите участок молодняков',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))
        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        list_layout = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        for section in sections:
            btn_card = MDCard(
                orientation='horizontal', size_hint_y=None, height=dp(48),
                padding=[Spacing.MD, 0], spacing=Spacing.MD,
                md_bg_color=[0.25, 0.25, 0.25, 1],
                on_release=lambda x, s=section[0]: self.load_saved_molodniki_section(s),
            )
            btn_card.add_widget(MDListItemLeadingIcon(icon='file-excel'))
            btn_card.add_widget(MDLabel(text=section[0], adaptive_height=True, valign='middle'))
            list_layout.add_widget(btn_card)
        scroll.add_widget(list_layout)
        content.add_widget(scroll)
        content.add_widget(make_outlined_btn('Закрыть', size_hint=(1, None), height=dp(48),
                           on_release=lambda x: load_dialog.dismiss()))
        load_dialog = MDDialog(
            MDDialogContentContainer(content),
            size_hint=(0.7, None),
        )
        load_dialog.open()

    def load_saved_molodniki_section(self, section_number):
        molodniki_screen = App.get_running_app().root.get_screen('molodniki')
        files = glob.glob(os.path.join(molodniki_screen.reports_dir, f'Молодняки_расширенный_{section_number}_*.xlsx'))
        if files:
            import pandas as pd
            latest_file = max(files, key=os.path.getctime)
            try:
                df = pd.read_excel(latest_file)
                molodniki_screen.current_section = section_number
                molodniki_screen.update_section_label()
                molodniki_screen.page_data.clear()
                for page_num in range(0, len(df), molodniki_screen.rows_per_page):
                    page = page_num // molodniki_screen.rows_per_page
                    page_data = df.iloc[page_num:page_num+molodniki_screen.rows_per_page].values.tolist()
                    for row in page_data:
                        while len(row) < 29:
                            row.append('')
                    molodniki_screen.page_data[page] = page_data
                molodniki_screen.current_page = 0
                molodniki_screen.load_page_data()
                molodniki_screen.update_pagination()
                self.show_success('Данные молодняков загружены!')
                App.get_running_app().root.current = 'molodniki'
            except Exception as e:
                self.show_error(f'Ошибка загрузки: {str(e)}')
        else:
            molodniki_screen.current_section = section_number
            molodniki_screen.update_section_label()
            App.get_running_app().root.current = 'molodniki'
            self.show_success(f'Участок молодняков {section_number} (новый)')

    def show_calculator_dialog(self):
        from core.forest_calculator import (calculate_area_ha, calculate_trees_per_ha,
                                             calculate_stock, calculate_density)
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='🧮 Калькулятор таксатора',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))

        fields = {}
        calc_params = [
            ('radius', 'Радиус площадки, м', '5.64'),
            ('count', 'Количество деревьев', '1'),
            ('diameter', 'Диаметр средний, см', '20'),
            ('height', 'Высота средняя, м', '15'),
        ]
        for key, label, default in calc_params:
            box = MDBoxLayout(orientation='vertical', size_hint_y=None, adaptive_height=True, spacing=dp(2))
            box.add_widget(MDLabel(text=label, font_size='12sp', size_hint_y=None, height=dp(20),
                                    theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY))
            inp = TextInput(text=default, multiline=False, size_hint_y=None, height=dp(40),
                            input_filter='float')
            fields[key] = inp
            box.add_widget(inp)
            content.add_widget(box)

        result_label = MDLabel(
            text='', font_size='14sp', size_hint_y=None, height=dp(120),
            theme_text_color='Custom', text_color=Colors.ACCENT,
        )
        content.add_widget(result_label)

        def calculate(inst):
            try:
                r = float(fields['radius'].text)
                n = float(fields['count'].text)
                d = float(fields['diameter'].text)
                h = float(fields['height'].text)
                area_ha = calculate_area_ha(r)
                density = calculate_density(r, n) if n > 0 else 0
                stock = calculate_stock(d, h, density) if density > 0 else 0
                tph = calculate_trees_per_ha(r, n)
                result_label.text = (
                    f'Площадь: {area_ha:.4f} га\n'
                    f'Густота: {density:.1f} шт/га\n'
                    f'Деревьев на га: {tph:.0f}\n'
                    f'Запас: {stock:.2f} м³'
                )
            except Exception as e:
                result_label.text = f'Ошибка: {str(e)}'

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Рассчитать', icon='calculator', size_hint=(0.5, None), height=dp(48),
                           on_release=calculate))
        btn_row.add_widget(make_outlined_btn('Закрыть', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: calc_dialog.dismiss()))
        content.add_widget(btn_row)

        calc_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.8, None))
        calc_dialog.open()

    def show_search_dialog(self):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='🔍 Поиск по данным', font_style='Title', role='medium', bold=True,
            halign='center', size_hint_y=None, height=dp(36),
        ))

        search_input = TextInput(
            hint_text='Введите породу, номер участка или дерева',
            multiline=False, size_hint_y=None, height=dp(44),
        )
        content.add_widget(search_input)

        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        results_box = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        scroll.add_widget(results_box)
        content.add_widget(scroll)

        def do_search(inst):
            query = search_input.text.strip()
            if not query:
                return
            results_box.clear_widgets()
            found = 0
            try:
                conn = sqlite3.connect('forest_data.db')
                cursor = conn.cursor()
                for table in ['sections', 'molodniki_sections']:
                    try:
                        cursor.execute(f"SELECT * FROM {table} WHERE section_number LIKE ? LIMIT 20",
                                       (f'%{query}%',))
                        for row in cursor.fetchall():
                            label = f'📁 {table}: {row[1] if len(row) > 1 else row[0]}'
                            results_box.add_widget(MDLabel(
                                text=label, font_size='13sp',
                                theme_text_color='Custom', text_color=Colors.PRIMARY_LIGHT,
                                size_hint_y=None, height=dp(24),
                            ))
                            found += 1
                    except Exception:
                        pass
                conn.close()
            except Exception:
                pass

            if found == 0:
                results_box.add_widget(MDLabel(
                    text='Ничего не найдено', font_size='14sp',
                    theme_text_color='Hint', halign='center', size_hint_y=None, height=dp(48),
                ))
            else:
                results_box.add_widget(MDLabel(
                    text=f'Найдено: {found}', font_size='12sp',
                    theme_text_color='Custom', text_color=Colors.TEXT_DIM,
                    size_hint_y=None, height=dp(20),
                ))

        search_input.bind(on_text_validate=do_search)
        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Искать', icon='magnify', size_hint=(0.5, None), height=dp(48),
                           on_release=do_search))
        btn_row.add_widget(make_outlined_btn('Закрыть', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: search_dialog.dismiss()))
        content.add_widget(btn_row)

        search_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.85, None))
        search_dialog.open()

    def show_backup_dialog(self):
        from core.backup_tools import create_backup, list_backups, restore_backup
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='💾 Резервное копирование',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))

        status_label = MDLabel(
            text='', font_size='13sp', size_hint_y=None, height=dp(24),
            theme_text_color='Custom', text_color=Colors.SUCCESS,
        )
        content.add_widget(status_label)

        def do_backup(inst):
            try:
                path = create_backup()
                status_label.text = f'✅ Бекап создан: {os.path.basename(path)}'
                refresh_list()
            except Exception as e:
                status_label.text = f'❌ Ошибка: {str(e)}'

        def do_restore(path, name):
            try:
                restore_backup(path)
                status_label.text = f'✅ Восстановлено из: {name}'
                refresh_list()
            except Exception as e:
                status_label.text = f'❌ Ошибка: {str(e)}'

        scroll = MDScrollView(size_hint_y=None, height=dp(250))
        list_box = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        scroll.add_widget(list_box)
        content.add_widget(scroll)

        def refresh_list():
            list_box.clear_widgets()
            backups = list_backups()
            if not backups:
                list_box.add_widget(MDLabel(
                    text='Нет бекапов', font_size='13sp', halign='center',
                    theme_text_color='Hint', size_hint_y=None, height=dp(48),
                ))
            for b in backups:
                row = MDCard(
                    orientation='horizontal', size_hint_y=None, height=dp(44),
                    padding=[Spacing.SM, 0], spacing=Spacing.SM,
                    md_bg_color=[0.25, 0.25, 0.25, 1],
                )
                row.add_widget(MDLabel(
                    text=f'{b["created"]} ({b["size"]/1024:.0f} KB)',
                    font_size='12sp', adaptive_height=True, valign='middle',
                    theme_text_color='Custom', text_color=[1,1,1,1],
                ))
                restore_btn = MDIconButton(
                    icon='restore', theme_icon_color='Custom', icon_color=Colors.WARNING,
                    on_release=lambda x, p=b['path'], n=b['name']: do_restore(p, n),
                )
                row.add_widget(restore_btn)
                list_box.add_widget(row)

        refresh_list()

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Создать бекап', icon='backup-restore', size_hint=(0.5, None), height=dp(48),
                           on_release=do_backup))
        btn_row.add_widget(make_outlined_btn('Закрыть', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: backup_dialog.dismiss()))
        content.add_widget(btn_row)

        backup_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.85, None))
        backup_dialog.open()

    def show_comparison_dialog(self):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='📋 Сравнение участков',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))

        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT section_number FROM sections WHERE section_number IS NOT NULL AND section_number != ""')
            sections = [row[0] for row in cursor.fetchall()]
            conn.close()
        except Exception:
            sections = []

        if len(sections) < 2:
            content.add_widget(MDLabel(
                text='Для сравнения нужно минимум 2 участка',
                font_size='14sp', halign='center', theme_text_color='Hint',
                size_hint_y=None, height=dp(48),
            ))
            content.add_widget(make_outlined_btn('Закрыть', size_hint=(1, None), height=dp(48),
                               on_release=lambda x: comp_dialog.dismiss()))
            comp_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.8, None))
            comp_dialog.open()
            return

        left_section = sections[0]
        right_section = sections[1] if len(sections) > 1 else sections[0]

        sel_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))

        left_dropdown_btn = make_outlined_btn(left_section, icon='arrow-left-drop-circle',
                                                size_hint=(0.5, None), height=dp(44))

        def make_dropdown(items, callback, caller):
            menu_items = [{
                'text': item,
                'on_release': lambda x=item: callback(item),
            } for item in items]
            return MDDropdownMenu(caller=caller, items=menu_items)

        left_menu = None
        right_menu = None

        def set_left(val):
            nonlocal left_section
            left_section = val
            left_dropdown_btn.children[1].text = val
            if left_menu:
                left_menu.dismiss()

        def set_right(val):
            nonlocal right_section
            right_section = val
            right_dropdown_btn.children[1].text = val
            if right_menu:
                right_menu.dismiss()

        left_dropdown_btn.on_release = lambda: make_dropdown(sections, set_left, left_dropdown_btn).open()
        right_dropdown_btn = make_outlined_btn(right_section, icon='arrow-right-drop-circle',
                                                size_hint=(0.5, None), height=dp(44))
        right_dropdown_btn.on_release = lambda: make_dropdown(sections, set_right, right_dropdown_btn).open()

        sel_row.add_widget(left_dropdown_btn)
        sel_row.add_widget(right_dropdown_btn)
        content.add_widget(sel_row)

        scroll = MDScrollView(size_hint_y=None, height=dp(300))
        compare_box = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        scroll.add_widget(compare_box)
        content.add_widget(scroll)

        def do_compare(inst):
            compare_box.clear_widgets()
            mol = App.get_running_app().root.get_screen('molodniki')
            left_totals = {}
            if hasattr(mol, 'calculate_totals'):
                left_totals = mol.calculate_totals()
            if not left_totals:
                compare_box.add_widget(MDLabel(text='Нет данных для сравнения', halign='center',
                                                theme_text_color='Hint', size_hint_y=None, height=dp(48)))
                return

            header = MDBoxLayout(orientation='horizontal', spacing=Spacing.SM, size_hint_y=None, height=dp(32))
            header.add_widget(MDLabel(text='Показатель', bold=True, font_size='12sp',
                                       theme_text_color='Custom', text_color=Colors.ACCENT))
            header.add_widget(MDLabel(text=left_section, bold=True, font_size='12sp', halign='center',
                                       theme_text_color='Custom', text_color=Colors.PRIMARY_LIGHT))
            header.add_widget(MDLabel(text=right_section, bold=True, font_size='12sp', halign='center',
                                       theme_text_color='Custom', text_color=Colors.SECONDARY_LIGHT))
            compare_box.add_widget(header)

            compare_rows = [
                ('Запас, м³', left_totals.get('total_stock', 0), left_totals.get('total_stock', 0)),
                ('Площадь, га', f'{left_totals.get("total_area", 0):.2f}', f'{left_totals.get("total_area", 0):.2f}'),
                ('Пород', left_totals.get('species_count', 0), left_totals.get('species_count', 0)),
            ]
            for label, left_val, right_val in compare_rows:
                row = MDBoxLayout(orientation='horizontal', spacing=Spacing.SM, size_hint_y=None, height=dp(24))
                row.add_widget(MDLabel(text=label, font_size='12sp', adaptive_height=True, valign='middle'))
                row.add_widget(MDLabel(text=str(left_val), font_size='12sp', halign='center',
                                       adaptive_height=True, valign='middle',
                                       theme_text_color='Custom', text_color=Colors.PRIMARY_LIGHT))
                row.add_widget(MDLabel(text=str(right_val), font_size='12sp', halign='center',
                                       adaptive_height=True, valign='middle',
                                       theme_text_color='Custom', text_color=Colors.SECONDARY_LIGHT))
                compare_box.add_widget(row)

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Сравнить', icon='compare', size_hint=(0.5, None), height=dp(48),
                           on_release=do_compare))
        btn_row.add_widget(make_outlined_btn('Закрыть', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: comp_dialog.dismiss()))
        content.add_widget(btn_row)

        comp_dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.9, None))
        comp_dialog.open()

    def show_reports_dialog(self):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='Отчёты', font_style='Title', role='medium', bold=True,
            halign='center', size_hint_y=None, height=dp(36),
        ))
        reports_dir = 'reports'
        if not os.path.exists(reports_dir):
            content.add_widget(MDLabel(
                text='Нет сохранённых отчётов', halign='center',
                size_hint_y=None, height=dp(48),
            ))
        else:
            files = [f for f in os.listdir(reports_dir) if f.endswith(('.xlsx', '.docx'))]
            if files:
                scroll = MDScrollView(size_hint_y=None, height=dp(300))
                list_layout = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
                for fname in files:
                    file_card = MDCard(
                        orientation='horizontal', size_hint_y=None, height=dp(48),
                        padding=[Spacing.MD, 0], spacing=Spacing.MD,
                        md_bg_color=[0.25, 0.25, 0.25, 1],
                        on_release=lambda x, fn=fname: self._open_report(fn),
                    )
                    icon = 'file-excel' if fname.endswith('.xlsx') else 'file-word'
                    file_card.add_widget(MDListItemLeadingIcon(icon=icon))
                    file_card.add_widget(MDLabel(text=fname, adaptive_height=True, valign='middle'))
                    list_layout.add_widget(file_card)
                scroll.add_widget(list_layout)
                content.add_widget(scroll)
            else:
                content.add_widget(MDLabel(
                    text='Нет сохранённых отчётов', halign='center',
                    size_hint_y=None, height=dp(48),
                ))
        content.add_widget(make_outlined_btn('Закрыть', size_hint=(1, None), height=dp(48),
                           on_release=lambda x: reports_dialog.dismiss()))
        reports_dialog = MDDialog(
            MDDialogContentContainer(content),
            size_hint=(0.8, None),
        )
        reports_dialog.open()

    def _open_report(self, filename):
        import subprocess
        path = os.path.join('reports', filename)
        try:
            os.startfile(path)
        except Exception:
            try:
                subprocess.Popen(['start', path], shell=True)
            except Exception:
                self.show_error('Не удалось открыть файл')

    def _show_about_dialog(self):
        dialog = MDDialog(
            MDDialogHeadlineText(text='🌲 Фанаты Пихты'),
            MDDialogSupportingText(
                text='Версия 2.0\n\n'
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

    def show_theme_dialog(self):
        """Диалог выбора темы"""
        theme_list = MDBoxLayout(orientation='vertical', spacing=Spacing.SM, adaptive_height=True)
        self.theme_manager = App.get_running_app().theme_manager

        for i in range(self.theme_manager.theme_count):
            theme = self.theme_manager.get_theme(i)
            is_active = i == self.theme_manager.current_theme_index
            theme_name = theme['name'].capitalize()

            item = MDCard(
                orientation='horizontal',
                size_hint_y=None,
                height=dp(48),
                padding=[Spacing.LG, 0],
                spacing=Spacing.MD,
                md_bg_color=[0.25, 0.25, 0.25, 1],
                on_release=lambda x, idx=i: self._apply_theme(idx),
            )
            icon = 'check-circle' if is_active else 'circle-outline'
            icon_widget = MDListItemLeadingIcon(icon=icon)
            label = MDLabel(
                text=f'{"☀️ " if theme["type"] == "color" and theme["name"] == "light" else "🌙 " if theme["type"] == "color" else "🖼️ "}{theme_name}',
                adaptive_height=True,
                valign='middle',
                theme_text_color='Custom', text_color=[1,1,1,1],
            )
            item.add_widget(icon_widget)
            item.add_widget(label)
            theme_list.add_widget(item)

        scroll = MDScrollView(size_hint_y=None, height=dp(250))
        scroll.add_widget(theme_list)

        dialog = MDDialog(
            MDDialogHeadlineText(text='🎨 Тема оформления'),
            MDDialogContentContainer(scroll),
            MDDialogButtonContainer(
                make_outlined_btn('Закрыть', on_release=lambda x: dialog.dismiss()),
                spacing='8dp',
            ),
        )
        dialog.open()

    def _apply_theme(self, index):
        app = App.get_running_app()
        if app.theme_manager.switch_theme(index):
            theme = app.theme_manager.current_theme
            app.theme_cls.theme_style = theme.get('kivymd_style', 'Light')
            app.theme_cls.primary_palette = theme.get('kivymd_palette', 'Green')
            theme_snack = MDSnackbar(duration=2)
            theme_snack.add_widget(MDSnackbarText(text='✅ Тема применена!'))
            theme_snack.open()
            # Пересоздаём UI для применения фона
            self.create_ui()


class ForestApp(MDApp):
    """Основное приложение"""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.theme_manager = ThemeManager()
        theme = self.theme_manager.current_theme
        self.theme_cls.theme_style = theme.get('kivymd_style', 'Light')
        self.theme_cls.primary_palette = theme.get('kivymd_palette', 'Green')

    def build(self):
        self.title = 'Фанаты Пихты'
        self.screen_manager = MDScreenManager()
        self.screen_manager.add_widget(MainMenu(name='main'))
        self.screen_manager.add_widget(ExtendedMolodnikiTableScreen(name='molodniki'))
        self.screen_manager.add_widget(DashboardScreen())
        self.screen_manager.add_widget(MapScreen())
        return self.screen_manager

    def on_start(self):
        self.init_database()

    def init_database(self):
        conn = sqlite3.connect('forest_data.db')
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                section_number TEXT UNIQUE,
                quarter TEXT, plot TEXT, forestry TEXT, district_forestry TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS molodniki_sections (
                id INTEGER PRIMARY KEY AUTOINCREMENT, section_number TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS suggestions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                column_index INTEGER, value TEXT, UNIQUE(column_index, value)
            )
        ''')
        conn.commit()
        conn.close()


if __name__ == '__main__':
    ForestApp().run()
