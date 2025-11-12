from kivy.uix.boxlayout import BoxLayout
from core.base_table import BaseTableScreen
from core.excel_tools import ExcelTemplateHandler
from . import formulas
from .widgets import SpeciesInput, DiameterInput

class MolodnikiTableScreen(BaseTableScreen):
    template_name = 'molodniki'
    version = '1.1'  # Версия шаблона

    def __init__(self, **kwargs):
        self.COLUMN_MAP = {
            'nn': {'type': 'int', 'header': 'NN'},
            'gps': {'type': 'coord', 'header': 'GPS точка'},
            # ... все колонки по шаблону
            'gustomota': {
                'type': 'formula',
                'header': 'Густота',
                'formula': formulas.gustomota,
                'depends': ['do_0.5m', '0.5-1.5m']
            }
        }
        super().__init__(**kwargs)

    def create_widgets(self):
        # Создание элементов по типам данных
        for col in self.COLUMN_MAP.values():
            if col['type'] == 'species':
                self.add_widget(SpeciesInput())
            elif col['type'] == 'diameter':
                self.add_widget(DiameterInput())
            # ... другие типы

    def on_data_change(self, instance, value):
        # Автоматический пересчет формул
        if instance.column in self.formula_dependencies:
            self.recalculate_formulas()

    def recalculate_formulas(self):
        # Применение бизнес-логики из formulas.py
        for row in self.data:
            for col in self.COLUMN_MAP.values():
                if col['type'] == 'formula':
                    row[col['name']] = col['formula'](row)

    def export_to_excel(self):
        handler = ExcelTemplateHandler(
            template_path='templates/molodniki_template.xlsx',
            output_path=self.get_export_path()
        )
        
        # Перенос данных
        handler.fill_data({
            'Лист4': self.prepare_sheet_data(),
            'Расчеты для проекта': self.prepare_calculations()
        })
        
        # Сохранение формул
        handler.preserve_formulas([
            "'Лист4'!AE16",
            "'Лист4'!S20:U20",
            # ... все важные формулы
        ])
        
        handler.save()

    def prepare_sheet_data(self):
        # Преобразование данных к структуре Excel
        return {
            'main_table': self.get_main_table_data(),
            'instructions': self.get_instructions(),
            'metadata': self.get_metadata()
        }