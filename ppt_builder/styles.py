# ppt_builder/styles.py
from pptx.dml.color import RGBColor

class PresentationStyle:
    """A class to hold all design parameters for a presentation."""
    def __init__(self, data):
        palette = data.get('color_palette', {})
        fonts = data.get('font_pairing', {})
        self.primary = self._hex_to_rgb(palette.get('primary', '#0D47A1'))
        self.secondary = self._hex_to_rgb(palette.get('secondary', '#42A5F5'))
        self.background = self._hex_to_rgb(palette.get('background', '#F5F5F5'))
        self.text = self._hex_to_rgb(palette.get('text', '#212121'))
        self.accent = self._hex_to_rgb(palette.get('accent', '#FFC107'))
        self.font_heading = fonts.get('heading', 'Calibri')
        self.font_body = fonts.get('body', 'Arial')
        self.design_concept = data.get('design_concept', 'Professional Presentation')

    @staticmethod
    def _hex_to_rgb(hex_str):
        hex_str = hex_str.lstrip('#')
        return RGBColor.from_string(hex_str)