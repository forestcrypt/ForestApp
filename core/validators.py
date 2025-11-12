import re
from kivy.app import App

class ForestValidator:
    @staticmethod
    def validate_diameter(value):
        try:
            return 0 < float(value) < 300  # см
        except:
            return False

    @staticmethod
    def validate_height(value):
        try:
            return 0 < float(value) < 150  # метров
        except:
            return False

    @staticmethod
    def validate_coordinates(value):
        return re.match(r'^\d{2,3}°\d{2}′\d{2}″[NS]\s\d{2,3}°\d{2}′\d{2}″[EW]$', value)

class ErrorHandler:
    @staticmethod
    def show_error(message):
        app = App.get_running_app()
        app.show_error_popup(message)