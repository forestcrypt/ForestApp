"""
Theme Manager for ForestApp.
Управление темами: светлая, тёмная, изображения.
"""
import os
import json
from kivy.utils import get_color_from_hex
from ui_styles import Colors


class ThemeManager:
    def __init__(self):
        self.themes = []
        self.current_theme_index = 0
        self.themes_dir = 'themes'
        os.makedirs(self.themes_dir, exist_ok=True)
        self.load_themes()
        self.load_config()

    def load_themes(self):
        self.themes = [{
            'type': 'color', 'name': 'light',
            'background': get_color_from_hex('#F5F5F0'),
            'text_color': '#1C1B1F',
            'kivymd_palette': 'Green', 'kivymd_style': 'Light'
        }, {
            'type': 'color', 'name': 'dark',
            'background': get_color_from_hex('#1C1B1F'),
            'text_color': '#E6E1E5',
            'kivymd_palette': 'Green', 'kivymd_style': 'Dark'
        }]
        for file in os.listdir(self.themes_dir):
            if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                self.themes.append({
                    'type': 'image', 'name': os.path.splitext(file)[0],
                    'background': os.path.join(self.themes_dir, file),
                    'text_color': '#FFFFFF'
                })

    @property
    def current_theme(self):
        return self.themes[self.current_theme_index]

    @property
    def theme_count(self):
        return len(self.themes)

    def get_theme(self, index):
        if 0 <= index < len(self.themes):
            return self.themes[index]
        return self.themes[0]

    def switch_theme(self, index):
        if 0 <= index < len(self.themes):
            self.current_theme_index = index
            self.save_config()
            return True
        return False

    def save_config(self):
        with open('theme_config.json', 'w') as f:
            json.dump({'theme_index': self.current_theme_index, 'themes_dir': self.themes_dir}, f)

    def load_config(self):
        try:
            with open('theme_config.json', 'r') as f:
                config = json.load(f)
                self.current_theme_index = config.get('theme_index', 0)
                self.themes_dir = config.get('themes_dir', 'themes')
        except Exception:
            pass

    def get_text_color(self, theme=None):
        if theme is None:
            theme = self.current_theme
        return get_color_from_hex(theme.get('text_color', '#1C1B1F'))

    def get_background(self, theme=None):
        if theme is None:
            theme = self.current_theme
        return theme['background']
