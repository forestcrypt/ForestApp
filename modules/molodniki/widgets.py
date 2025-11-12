from kivy.uix.textinput import TextInput
from kivy.uix.dropdown import DropDown
from core.validators import ForestValidator, ErrorHandler
from kivy.uix.button import Button
from kivy.properties import ListProperty


class ScientificInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bind(text=self.validate)

    def validate(self, instance, value):
        if not ForestValidator.validate_float(value):
            ErrorHandler.show_error("Некорректное числовое значение")
            self.background_color = (1, 0.8, 0.8, 1)

class SpeciesDropdown(DropDown):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.load_species()
    
    def load_species(self):
        from core.database import get_species_list
        for species in get_species_list():
            btn = Button(text=species.name)
            btn.bind(on_release=self.select)
            self.add_widget(btn)

class SpeciesInput(TextInput):
    species = ListProperty(['Сосна', 'Ель', 'Берёза', 'Осина'])
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dropdown = DropDown()
        self.create_dropdown()
        self.bind(focus=self.on_focus)
    
    def create_dropdown(self):
        for species in self.species:
            btn = Button(text=species, size_hint_y=None, height=44)
            btn.bind(on_release=lambda btn: self.select_species(btn.text))
            self.dropdown.add_widget(btn)
    
    def select_species(self, value):
        self.text = value
        self.dropdown.dismiss()
    
    def on_focus(self, instance, value):
        if value: 
            self.dropdown.open(self)

class DiameterInput(TextInput):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.input_filter = 'float'
        self.hint_text = "0.0-150.0 см"
    
    def insert_text(self, substring, from_undo=False):
        if not substring.isdigit() and substring != '.':
            return
        return super().insert_text(substring, from_undo)            