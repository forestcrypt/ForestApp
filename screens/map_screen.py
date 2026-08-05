import os
import sqlite3
import re

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup

from kivy_garden.mapview import MapView, MapMarker

from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.card import MDCard
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText, MDButtonIcon, MDIconButton
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.dialog import MDDialog, MDDialogContentContainer
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText

from ui_styles import Colors, Spacing


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


COORD_RE = re.compile(r'^(\d{2,3})°(\d{2})′(\d{2})″([NS])\s+(\d{2,3})°(\d{2})′(\d{2})″([EW])$')


def parse_coordinates(text):
    m = COORD_RE.match(text.strip())
    if not m:
        return None
    lat = int(m.group(1)) + int(m.group(2)) / 60 + int(m.group(3)) / 3600
    if m.group(4) == 'S':
        lat = -lat
    lon = int(m.group(5)) + int(m.group(6)) / 60 + int(m.group(7)) / 3600
    if m.group(8) == 'W':
        lon = -lon
    return lat, lon


def format_coords(lat, lon):
    def dec_to_dms(dd, is_lat):
        d = abs(dd)
        deg = int(d)
        minutes = int((d - deg) * 60)
        seconds = round((d - deg - minutes / 60) * 3600)
        letter = 'N' if is_lat and dd >= 0 else 'S' if is_lat else 'E' if dd >= 0 else 'W'
        return f'{deg}°{minutes:02d}′{seconds:02d}″{letter}'
    return f'{dec_to_dms(lat, True)} {dec_to_dms(lon, False)}'


MARKER_SOURCE = None


def get_marker_source():
    global MARKER_SOURCE
    if MARKER_SOURCE:
        return MARKER_SOURCE
    path = 'marker.png'
    if os.path.exists(path):
        MARKER_SOURCE = path
        return path
    try:
        from PIL import Image, ImageDraw
        img = Image.new('RGBA', (32, 32), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([4, 4, 28, 28], fill=(76, 175, 80, 255))
        draw.ellipse([8, 8, 24, 24], fill=(255, 255, 255, 255))
        img.save(path)
        MARKER_SOURCE = path
        return path
    except Exception:
        return ''


class ForestMapMarker(MapMarker):
    def __init__(self, section_number, **kwargs):
        super().__init__(**kwargs)
        self.section_number = section_number


class MapScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'map'
        self.markers = []
        self.default_lat = 55.7558
        self.default_lon = 37.6173
        Clock.schedule_once(lambda dt: self._build(), 0)

    def _build(self):
        self.clear_widgets()
        main = MDBoxLayout(orientation='vertical')

        from main_modern import MDTopAppBarOld
        toolbar = MDTopAppBarOld(
            title='Карта участков',
            elevation=2, md_bg_color=Colors.SECONDARY,
            left_action_items=[['arrow-left', lambda x: self._go_back()]],
            right_action_items=[
                ['plus', lambda x: self._add_coords_dialog()],
                ['refresh', lambda x: self._refresh_map()],
            ],
        )
        main.add_widget(toolbar)

        self.mapview = MapView(
            lat=self.default_lat,
            lon=self.default_lon,
            zoom=10,
        )
        main.add_widget(self.mapview)

        self.add_widget(main)
        Clock.schedule_once(lambda dt: self._refresh_map(), 0.1)

    def _go_back(self):
        App.get_running_app().root.current = 'main'

    def _refresh_map(self):
        self.mapview.clear_widgets()
        self.markers.clear()

        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT section_number FROM sections WHERE section_number IS NOT NULL AND section_number != ""')
            sections1 = [(r[0], '', '', '', '') for r in cursor.fetchall()]
            cursor.execute('SELECT DISTINCT section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != ""')
            sections2 = [(r[0], '', '', '', '') for r in cursor.fetchall()]
            conn.close()
            seen = set()
            sections = []
            for s in sections1 + sections2:
                if s[0] not in seen:
                    seen.add(s[0])
                    sections.append(s)
        except Exception:
            sections = []

        coords_file = 'sections_coords.json'
        saved_coords = {}
        if os.path.exists(coords_file):
            import json
            try:
                with open(coords_file, 'r', encoding='utf-8') as f:
                    saved_coords = json.load(f)
            except Exception:
                pass

        source = get_marker_source()
        for section in sections:
            section_number = section[0]
            if section_number in saved_coords:
                coords = saved_coords[section_number]
                lat, lon = coords.get('lat'), coords.get('lon')
                if lat and lon:
                    marker = ForestMapMarker(
                        section_number=section_number,
                        lat=lat, lon=lon,
                        source=source,
                    )
                    self.mapview.add_widget(marker)
                    self.markers.append(marker)

        if not self.markers and sections:
            self._snack('Нет участков с координатами. Нажмите + чтобы добавить.')

    def _snack(self, message):
        snack = MDSnackbar(duration=2.5)
        snack.add_widget(MDSnackbarText(text=message))
        snack.open()

    def _add_coords_dialog(self, section_number=None):
        content = MDBoxLayout(orientation='vertical', spacing=Spacing.MD, adaptive_height=True)
        content.add_widget(MDLabel(
            text='🗺️ Добавить координаты',
            font_style='Title', role='medium', bold=True, halign='center',
            size_hint_y=None, height=dp(36),
        ))

        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT section_number FROM sections WHERE section_number IS NOT NULL AND section_number != ""')
            s1 = [r[0] for r in cursor.fetchall()]
            cursor.execute('SELECT DISTINCT section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != ""')
            s2 = [r[0] for r in cursor.fetchall()]
            conn.close()
            seen = set()
            sections = []
            for s in s1 + s2:
                if s not in seen:
                    seen.add(s)
                    sections.append(s)
        except Exception:
            sections = []

        if not sections:
            self._snack('Нет участков. Создайте участок в РУМ.')
            return

        section_input = TextInput(
            text=section_number or sections[0],
            multiline=False, size_hint_y=None, height=dp(44),
        )
        content.add_widget(section_input)

        content.add_widget(MDLabel(
            text='Формат: 55°45′30″N 37°37′00″E\nили широта,долгота (55.7558, 37.6173)',
            font_size='12sp', theme_text_color='Custom', text_color=Colors.TEXT_SECONDARY,
            size_hint_y=None, height=dp(36),
        ))

        lat_input = TextInput(
            hint_text='Широта (55.7558)', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        lon_input = TextInput(
            hint_text='Долгота (37.6173)', multiline=False,
            size_hint_y=None, height=dp(44),
        )
        content.add_widget(lat_input)
        content.add_widget(lon_input)

        def save_coords(inst):
            section = section_input.text.strip()
            lat_text = lat_input.text.strip()
            lon_text = lon_input.text.strip()

            if not section or not lat_text or not lon_text:
                self._snack('Заполните все поля')
                return

            try:
                lat = float(lat_text)
                lon = float(lon_text)
            except ValueError:
                self._snack('Некорректные координаты')
                return

            import json
            coords_file = 'sections_coords.json'
            saved = {}
            if os.path.exists(coords_file):
                try:
                    with open(coords_file, 'r', encoding='utf-8') as f:
                        saved = json.load(f)
                except Exception:
                    pass

            saved[section] = {'lat': lat, 'lon': lon, 'label': section}
            with open(coords_file, 'w', encoding='utf-8') as f:
                json.dump(saved, f, ensure_ascii=False, indent=2)

            dialog.dismiss()
            Clock.schedule_once(lambda dt: self._refresh_map(), 0)
            Clock.schedule_once(lambda dt: self._center_on(lat, lon), 0.1)
            self._snack(f'✅ Координаты для участка {section} сохранены')

        btn_row = MDBoxLayout(orientation='horizontal', spacing=Spacing.MD, size_hint_y=None, height=dp(48))
        btn_row.add_widget(make_raised_btn('Сохранить', icon='content-save', size_hint=(0.5, None), height=dp(48),
                           on_release=save_coords))
        btn_row.add_widget(make_outlined_btn('Отмена', size_hint=(0.5, None), height=dp(48),
                           on_release=lambda x: dialog.dismiss()))
        content.add_widget(btn_row)

        dialog = MDDialog(MDDialogContentContainer(content), size_hint=(0.8, None))
        dialog.open()

    def _center_on(self, lat, lon):
        if self.mapview.parent:
            self.mapview.center_on(lat, lon)
            Clock.schedule_once(lambda dt: setattr(self.mapview, 'zoom', 14), 0.1)
