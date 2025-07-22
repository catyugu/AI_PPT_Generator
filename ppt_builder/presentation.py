import logging
from pptx import Presentation
from pptx.dml.color import RGBColor
from ppt_builder.slide_renderer import SlideRenderer
from ppt_builder.styles import PresentationStyle, px_to_emu, hex_to_rgb
from image_service import ImageService


class PresentationBuilder:
    """根据AI生成的计划构建完整的演示文稿。"""

    def __init__(self, plan: dict):
        self.plan = plan
        self.prs = Presentation()
        self.style_manager = PresentationStyle(plan)
        self.image_service = ImageService()

        self.background_image_path = None
        master_data = self.plan.get('master_slide', {})
        if keyword := master_data.get('background', {}).get('image_keyword'):
            logging.info(f"正在为全局背景预生成图片: '{keyword}'")
            self.background_image_path = self.image_service.generate_image(keyword)
            if not self.background_image_path:
                logging.warning(f"无法为关键词 '{keyword}' 生成背景图片。")

        self.slide_renderer = SlideRenderer(
            self.prs,
            self.style_manager,
            self.background_image_path
        )
        logging.info("PresentationBuilder已使用AI方案和StyleManager初始化。")

    def _apply_master_slide_styles(self):
        """
        [已简化] 应用全局母版样式。
        此函数现在只负责处理背景颜色，图片背景由SlideRenderer在每页幻灯片上应用。
        """
        master = self.prs.slide_masters[0]
        master_data = self.plan.get('master_slide', {})
        background_info = master_data.get('background', {})
        fill = master.background.fill

        if self.background_image_path:
            fill.background()
            logging.info("检测到全局背景图片，母版背景将保持透明。")
            return

        try:
            if specific_bg_color_hex := background_info.get('color'):
                fill.solid()
                fill.fore_color.rgb = hex_to_rgb(specific_bg_color_hex)
                logging.info(f"已将母版背景设置为指定颜色: {specific_bg_color_hex}")
            else:
                fill.solid()
                fill.fore_color.rgb = self.style_manager.get_color('background')
                logging.info("已将母版背景设置为调色板中的默认背景色。")
        except Exception as e:
            logging.error(f"应用母版颜色背景时出错: {e}", exc_info=True)
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)

    def build_presentation(self, output_path: str):
        """构建并保存演示文稿。"""
        try:
            self.prs.slide_width = px_to_emu(1280)
            self.prs.slide_height = px_to_emu(720)
            self._apply_master_slide_styles()

            pages = self.plan.get('pages', [])
            total_pages = len(pages)
            for i, page_data in enumerate(pages):
                logging.info(f"--- 正在构建页面 {i + 1}/{total_pages} ---")
                # **[已修复]** 在此处调用 render_slide 时，传入 self.image_service
                self.slide_renderer.render_slide(page_data, self.image_service)

            self.prs.save(output_path)
            logging.info(f"演示文稿已成功保存至 {output_path}")
        except Exception as e:
            logging.error(f"构建演示文稿过程中发生严重错误: {e}", exc_info=True)
            raise
