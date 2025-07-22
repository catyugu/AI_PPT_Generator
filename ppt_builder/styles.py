# ppt_builder/styles.py
import logging

from pptx.dml.color import RGBColor

class PresentationStyle:
    """A class to hold all design parameters for a presentation."""

    def __init__(self, data):
        palette = data.get('color_palette', {})
        fonts = data.get('font_pairing', {})

        # [UPGRADED] This block now handles both dict and list formats
        if isinstance(palette, dict):
            self.primary = self._hex_to_rgb(palette.get('primary', '#0D47A1'))
            self.secondary = self._hex_to_rgb(palette.get('secondary', '#42A5F5'))
            self.background = self._hex_to_rgb(palette.get('background', '#F5F5F5'))
            self.text = self._hex_to_rgb(palette.get('text', '#212121'))
            self.accent = self._hex_to_rgb(palette.get('accent', '#FFC107'))
        elif isinstance(palette, list) and len(palette) >= 5:
            logging.warning("Color palette was a list, not a dict. Assigning colors by index.")
            self.primary = self._hex_to_rgb(palette[0])
            self.secondary = self._hex_to_rgb(palette[1])
            self.background = self._hex_to_rgb(palette[2])
            self.text = self._hex_to_rgb(palette[3])
            self.accent = self._hex_to_rgb(palette[4])
        else:
            logging.error("Invalid color palette format or not found. Using default colors.")
            self.primary = self._hex_to_rgb('#0D47A1')
            self.secondary = self._hex_to_rgb('#42A5F5')
            self.background = self._hex_to_rgb('#F5F5F5')
            self.text = self._hex_to_rgb('#212121')
            self.accent = self._hex_to_rgb('#FFC107')

        self.font_heading = fonts.get('heading', 'Calibri')
        self.font_body = fonts.get('body', 'Arial')
        self.design_concept = data.get('design_concept', 'Professional Presentation')

    @staticmethod
    def _hex_to_rgb(hex_str):
        hex_str = hex_str.lstrip('#')
        return RGBColor.from_string(hex_str)