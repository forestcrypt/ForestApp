import os
import io
import sqlite3
import tempfile
from collections import Counter

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp
from kivy.uix.image import Image
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.scrollview import ScrollView
from kivy.core.image import Image as CoreImage

from kivymd.uix.screen import MDScreen
from kivymd.uix.card import MDCard
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.button import MDButton, MDButtonText
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer, MDDialogContentContainer

from ui_styles import Colors, Spacing


plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['figure.facecolor'] = '#1e1e1e'
plt.rcParams['axes.facecolor'] = '#2d2d2d'
plt.rcParams['axes.edgecolor'] = '#555555'
plt.rcParams['axes.labelcolor'] = '#cccccc'
plt.rcParams['xtick.color'] = '#cccccc'
plt.rcParams['ytick.color'] = '#cccccc'
plt.rcParams['text.color'] = '#ffffff'


class ChartWidget(MDBoxLayout):
    def __init__(self, title, fig, **kwargs):
        super().__init__(orientation='vertical', size_hint_y=None, spacing=Spacing.SM, padding=[Spacing.SM, 0], **kwargs)

        self.add_widget(MDLabel(
            text=title, font_style='Title', role='small', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            size_hint_y=None, height=dp(24),
        ))

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100, bbox_inches='tight', facecolor=fig.get_facecolor())
        plt.close(fig)
        buf.seek(0)

        img = Image(
            texture=CoreImage(buf, ext='png').texture,
            size_hint_y=None,
            height=dp(240),
            allow_stretch=True,
            keep_ratio=True,
        )
        self.add_widget(img)
        self.height = dp(280)


class DashboardScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.name = 'dashboard'
        self.chart_dir = tempfile.mkdtemp(prefix='forestapp_charts_')
        Clock.schedule_once(lambda dt: self._build(), 0)

    def _build(self):
        self.clear_widgets()
        main = MDBoxLayout(orientation='vertical')

        from main_modern import MDTopAppBarOld
        toolbar = MDTopAppBarOld(
            title='Дашборд',
            elevation=2, md_bg_color=Colors.PRIMARY,
            left_action_items=[['arrow-left', lambda x: self._go_back()]],
            right_action_items=[['refresh', lambda x: self._build()]],
        )
        main.add_widget(toolbar)

        scroll = MDScrollView()
        content = MDBoxLayout(
            orientation='vertical', size_hint_y=None,
            spacing=Spacing.MD, padding=[Spacing.MD, Spacing.SM],
        )
        content.bind(minimum_height=content.setter('height'))

        stats = self._gather_stats()
        stats_card = self._make_stats_card(stats)
        content.add_widget(stats_card)

        for chart_widget in self._make_charts(stats):
            content.add_widget(chart_widget)

        scroll.add_widget(content)
        main.add_widget(scroll)
        self.add_widget(main)

    def _go_back(self):
        App.get_running_app().root.current = 'main'

    def _gather_stats(self):
        result = {
            'total_sections': 0, 'total_plots': 0,
            'total_stock': 0, 'total_area': 0,
            'species_counts': {}, 'section_list': [],
            'ages': [], 'heights': [],
        }
        try:
            conn = sqlite3.connect('forest_data.db')
            cursor = conn.cursor()

            cursor.execute('SELECT COUNT(*) FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != ""')
            result['total_sections'] = cursor.fetchone()[0]

            cursor.execute('SELECT DISTINCT section_number FROM molodniki_sections WHERE section_number IS NOT NULL AND section_number != ""')
            result['section_list'] = [r[0] for r in cursor.fetchall()]

            conn.close()
        except Exception:
            pass

        try:
            mol = App.get_running_app().root.get_screen('molodniki')
            if hasattr(mol, 'calculate_section_totals'):
                totals = mol.calculate_section_totals()
                result['total_stock'] = totals.get('total_stock', 0)
                result['total_area'] = totals.get('total_area', 0)
                result['total_plots'] = totals.get('total_plots', 0)

                species_summary = totals.get('species_summary', {})
                for species, data in species_summary.items():
                    result['species_counts'][species] = data.get('count', data.get('area', 0))
                    ages = data.get('ages', [])
                    if ages:
                        result['ages'].extend(ages)
                    heights = data.get('heights', [])
                    if heights:
                        result['heights'].extend(heights)

            if not result['species_counts'] and hasattr(mol, 'page_data'):
                for page_rows in mol.page_data.values():
                    for row in page_rows:
                        if len(row) >= 2 and row[1]:
                            species_key = str(row[1]).strip()
                            if species_key:
                                result['species_counts'][species_key] = result['species_counts'].get(species_key, 0) + 1
        except Exception:
            pass

        return result

    def _make_stats_card(self, stats):
        card = MDCard(
            orientation='vertical', size_hint_y=None,
            padding=Spacing.LG, spacing=Spacing.SM,
            radius=[Spacing.RADIUS_LG], elevation=2,
            md_bg_color=[0.2, 0.2, 0.2, 1],
        )
        card.add_widget(MDLabel(
            text='📊 СТАТИСТИКА', font_style='Title', role='small', bold=True,
            theme_text_color='Custom', text_color=[1,1,1,1],
            size_hint_y=None, height=dp(24),
        ))

        rows = MDGridLayout(cols=3, spacing=Spacing.SM, size_hint_y=None, adaptive_height=True)

        def stat_item(label, value, color):
            box = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(60), spacing=dp(2))
            box.add_widget(MDLabel(
                text=str(value), font_size='24sp', bold=True, halign='center',
                theme_text_color='Custom', text_color=color,
                size_hint_y=None, height=dp(32),
            ))
            box.add_widget(MDLabel(
                text=label, font_size='11sp', halign='center',
                theme_text_color='Custom', text_color=[1,1,1,0.7],
                size_hint_y=None, height=dp(20),
            ))
            return box

        rows.add_widget(stat_item('Участков', stats['total_sections'], Colors.ACCENT))
        rows.add_widget(stat_item('Площадок', stats.get('total_plots', 0), Colors.PRIMARY_LIGHT))
        rows.add_widget(stat_item('Пород', len(stats['species_counts']), Colors.INFO))
        rows.add_widget(stat_item(f'Запас', f'{stats["total_stock"]:.1f} м³', Colors.SECONDARY_LIGHT))
        rows.add_widget(stat_item(f'Площадь', f'{stats["total_area"]:.2f} га', Colors.WARNING))
        rows.add_widget(stat_item('Секций', len(stats['section_list']), Colors.TEXT_DIM))

        card.add_widget(rows)
        card.height = dp(180)
        return card

    def _make_charts(self, stats):
        charts = []

        if stats['species_counts']:
            fig, ax = plt.subplots(figsize=(4, 3))
            species = list(stats['species_counts'].keys())
            counts = list(stats['species_counts'].values())
            colors = plt.cm.Set3([i/len(species) for i in range(len(species))])
            wedges, texts, autotexts = ax.pie(
                counts, labels=None, autopct='%1.0f%%',
                colors=colors, startangle=90,
                textprops={'color': 'white', 'fontsize': 8},
            )
            if species:
                ax.legend(wedges, [f'{s} ({c})' for s, c in zip(species, counts)],
                         loc='lower center', bbox_to_anchor=(0.5, -0.15),
                         ncol=2, fontsize=7, frameon=False, labelcolor='white')
            ax.set_title('Распределение пород', color='white', fontsize=11, pad=10)
            charts.append(ChartWidget('🌳 Распределение по породам', fig))

        if stats['ages']:
            fig, ax = plt.subplots(figsize=(4, 2.5))
            ax.hist(stats['ages'], bins=8, color='#4CAF50', edgecolor='white', alpha=0.8)
            ax.set_xlabel('Возраст, лет', color='#cccccc', fontsize=9)
            ax.set_ylabel('Количество', color='#cccccc', fontsize=9)
            ax.set_title('Распределение возрастов', color='white', fontsize=11)
            ax.tick_params(labelsize=8)
            charts.append(ChartWidget('📏 Распределение возрастов', fig))

        if stats['heights']:
            fig, ax = plt.subplots(figsize=(4, 2.5))
            ax.hist(stats['heights'], bins=8, color='#42A5F5', edgecolor='white', alpha=0.8)
            ax.set_xlabel('Высота, м', color='#cccccc', fontsize=9)
            ax.set_ylabel('Количество', color='#cccccc', fontsize=9)
            ax.set_title('Распределение высот', color='white', fontsize=11)
            ax.tick_params(labelsize=8)
            charts.append(ChartWidget('📐 Распределение высот', fig))

        if not charts:
            no_data = MDBoxLayout(orientation='vertical', size_hint_y=None, height=dp(120))
            no_data.add_widget(MDLabel(
                text='Нет данных для построения графиков.\nДобавьте участки молодняков в разделе РУМ.',
                halign='center', theme_text_color='Hint', size_hint_y=None, height=dp(80),
            ))
            charts.append(no_data)

        return charts
