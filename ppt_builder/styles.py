import logging
from pptx.util import Emu
from pptx.dml.color import RGBColor


class PresentationStyle:
    """
    管理演示文稿的全局设计系统，包括颜色、字体和默认样式。
    """

    def __init__(self, plan: dict):
        """
        从AI生成的计划中初始化设计系统。

        Args:
            plan (dict): 包含 `color_palette` 和 `font_pairing` 的AI计划。
        """
        # [核心优化] 从 plan 中提取设计系统信息
        self.design_info = plan
        self.color_palette = self.design_info.get('color_palette', {})
        self.font_pairing = self.design_info.get('font_pairing', {})

        # 解析颜色，提供备用值
        self.primary = self._hex_to_rgb(self.color_palette.get('primary', '#0D47A1'))
        self.secondary = self._hex_to_rgb(self.color_palette.get('secondary', '#42A5F5'))
        self.background = self._hex_to_rgb(self.color_palette.get('background', '#F5F5F5'))
        self.text_color = self._hex_to_rgb(self.color_palette.get('text', '#333333'))
        self.accent = self._hex_to_rgb(self.color_palette.get('accent', '#FFC107'))

        # 解析字体，提供备用值
        self.heading_font = self.font_pairing.get('heading', 'Arial Black')
        self.body_font = self.font_pairing.get('body', 'Arial')

        # [新功能] 为图表定义一个颜色系列，确保图表颜色与主题匹配
        self.chart_colors = [
            self.accent,
            self.secondary,
            self.primary,
            self._hex_to_rgb(self.color_palette.get('text', '#333333')),
            self._hex_to_rgb('#9E9E9E'),  # 灰色作为补充
            self._hex_to_rgb('#7CB342')  # 绿色作为补充
        ]
        logging.info(f"PresentationStyle initialized with concept: '{self.design_info.get('design_concept', 'N/A')}'")

    def _hex_to_rgb(self, hex_color: str) -> RGBColor:
        """
        将十六进制颜色码转换为RGBColor对象。
        """
        hex_color = hex_color.lstrip('#')
        try:
            return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
        except (ValueError, IndexError):
            logging.warning(f"Invalid hex color '{hex_color}'. Defaulting to black.")
            return RGBColor(0, 0, 0)

    def get_color(self, color_name: str) -> RGBColor:
        """
        从调色板中获取指定名称的颜色。
        """
        return self._hex_to_rgb(self.color_palette.get(color_name, '#000000'))

    def get_chart_color(self, index: int) -> RGBColor:
        """
        为图表系列提供循环的颜色。
        """
        return self.chart_colors[index % len(self.chart_colors)]


def px_to_emu(px):
    """像素到EMU的转换"""
    return Emu(int(px * 9525))


def hex_to_rgb(hex_color):
    """十六进制到RGB的转换"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
