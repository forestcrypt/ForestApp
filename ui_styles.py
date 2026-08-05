"""
ForestApp Design System.
Единая система стилей: цвета, отступы, типографика, компоненты.
"""
from kivy.utils import get_color_from_hex
from kivy.metrics import dp
from kivy.animation import Animation
from kivy.graphics import Color, RoundedRectangle
from kivy.properties import ListProperty, BooleanProperty
from kivy.uix.button import Button


class Colors:
    PRIMARY = get_color_from_hex('#2E7D32')
    PRIMARY_LIGHT = get_color_from_hex('#4CAF50')
    PRIMARY_DIM = get_color_from_hex('#1B5E20')
    GREEN = get_color_from_hex('#4CAF50')
    SECONDARY = get_color_from_hex('#1565C0')
    SECONDARY_LIGHT = get_color_from_hex('#42A5F5')
    ACCENT = get_color_from_hex('#F9A825')
    ACCENT_LIGHT = get_color_from_hex('#FFD54F')

    SURFACE = get_color_from_hex('#F5F5F5')
    SURFACE_ALT = get_color_from_hex('#E8F5E9')
    SURFACE_CARD = get_color_from_hex('#FFFFFF')
    CARD_BG = get_color_from_hex('#1E1E1E')
    DARK_SURFACE = get_color_from_hex('#1E1E1E')
    SURFACE_HOVER = get_color_from_hex('#C8E6C9')
    BACKGROUND = get_color_from_hex('#FAFAFA')
    OVERLAY = get_color_from_hex('#00000040')

    TEXT = get_color_from_hex('#212121')
    TEXT_SECONDARY = get_color_from_hex('#616161')
    TEXT_DIM = get_color_from_hex('#9E9E9E')
    TEXT_ON_PRIMARY = get_color_from_hex('#FFFFFF')
    TEXT_ON_DARK = get_color_from_hex('#E0E0E0')

    SUCCESS = get_color_from_hex('#2E7D32')
    SUCCESS_BG = get_color_from_hex('#E8F5E9')
    WARNING = get_color_from_hex('#F9A825')
    WARNING_BG = get_color_from_hex('#FFF8E1')
    DANGER = get_color_from_hex('#C62828')
    DANGER_BG = get_color_from_hex('#FFEBEE')
    INFO = get_color_from_hex('#1565C0')
    INFO_BG = get_color_from_hex('#E3F2FD')

    BORDER = get_color_from_hex('#BDBDBD')
    BORDER_LIGHT = get_color_from_hex('#E0E0E0')

    # Semantic button colors
    BTN_PRIMARY = get_color_from_hex('#2E7D32')
    BTN_SUCCESS = get_color_from_hex('#43A047')
    BTN_DANGER = get_color_from_hex('#E53935')
    BTN_WARNING = get_color_from_hex('#FB8C00')
    BTN_INFO = get_color_from_hex('#1E88E5')
    BTN_PURPLE = get_color_from_hex('#7B1FA2')
    BTN_TEAL = get_color_from_hex('#00897B')
    BTN_DARK = get_color_from_hex('#424242')

    DARK = get_color_from_hex('#424242')


class Spacing:
    XS = dp(4)
    SM = dp(8)
    MD = dp(12)
    LG = dp(16)
    XL = dp(20)
    XXL = dp(24)
    XXXL = dp(32)
    SECTION = dp(48)

    BUTTON_HEIGHT = dp(40)
    INPUT_HEIGHT = dp(40)
    HEADER_HEIGHT = dp(56)
    TOOLBAR_HEIGHT = dp(48)
    ROW_HEIGHT = dp(40)
    CARD_HEIGHT = dp(80)

    RADIUS_XS = dp(4)
    RADIUS_SM = dp(6)
    RADIUS_MD = dp(8)
    RADIUS_LG = dp(12)
    RADIUS_XL = dp(16)
    RADIUS_FULL = dp(999)


class Fonts:
    FAMILY = 'Roboto'
    FAMILY_MONO = 'RobotoMono'
    H1 = '28sp'
    H2 = '22sp'
    H3 = '18sp'
    BODY = '14sp'
    BODY_SM = '12sp'
    BODY_XS = '11sp'
    CAPTION = '10sp'
    LABEL = '13sp'


class ModernButton(Button):
    bg_color = ListProperty(Colors.PRIMARY)
    auto_width = BooleanProperty(True)
    no_shadow = BooleanProperty(False)

    def __init__(self, **kwargs):
        self.auto_width = kwargs.pop('auto_width', True)
        self.no_shadow = kwargs.pop('no_shadow', False)
        color_provided = 'color' in kwargs
        size_hint_provided = 'size_hint' in kwargs
        height_provided = 'height' in kwargs
        font_size_provided = 'font_size' in kwargs
        font_name_provided = 'font_name' in kwargs
        padding_provided = 'padding' in kwargs
        super().__init__(**kwargs)
        self.background_color = (0, 0, 0, 0)
        if not font_name_provided:
            self.font_name = Fonts.FAMILY
        if not font_size_provided:
            self.font_size = Fonts.BODY
        if not size_hint_provided:
            self.size_hint = (None, None)
        if not height_provided:
            self.height = Spacing.BUTTON_HEIGHT
        if not padding_provided:
            self.padding = (dp(20), dp(5))
        if not color_provided:
            self.color = Colors.TEXT_ON_PRIMARY

        with self.canvas.before:
            if not self.no_shadow:
                Color(rgba=(0, 0, 0, 0.1))
                self.shadow = RoundedRectangle(
                    pos=(self.x + dp(2), self.y - dp(2)),
                    size=self.size,
                    radius=[Spacing.RADIUS_MD]
                )
            self.bg_color_instruction = Color(rgba=self.bg_color)
            self.background = RoundedRectangle(
                pos=self.pos,
                size=self.size,
                radius=[Spacing.RADIUS_MD]
            )

        self.bind(pos=self.update_graphics, size=self.update_graphics)
        if self.auto_width:
            self.bind(text=self.update_width)
        self.bind(size=self._update_text_size)

    def _update_text_size(self, *args):
        self.text_size = (self.width - self.padding[0] * 2, self.height)

    def update_width(self, instance, value):
        self.width = self.texture_size[0] + dp(60)

    def update_graphics(self, *args):
        self.background.pos = self.pos
        if not self.no_shadow:
            self.shadow.pos = (self.x + dp(2), self.y - dp(2))
            self.shadow.size = self.size
        self.background.size = self.size

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos):
            Animation(rgba=[c * 0.85 for c in self.bg_color], d=0.1).start(self.bg_color_instruction)
        return super().on_touch_down(touch)

    def on_touch_up(self, touch):
        Animation(rgba=self.bg_color, d=0.2).start(self.bg_color_instruction)
        return super().on_touch_up(touch)
