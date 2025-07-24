import logging
from pptx import Presentation
from pptx.dml.color import RGBColor
from ppt_builder.slide_renderer import SlideRenderer
from ppt_builder.styles import PresentationStyle, px_to_emu, hex_to_rgb
from image_service import ImageService


class PresentationBuilder:
    """
    [已更新] 根据AI生成的计划构建完整的演示文稿，并支持不同宽高比。
    """

    def __init__(self, plan: dict, aspect_ratio: str = "16:9"):
        """
        初始化构建器。
        :param plan: AI生成的JSON方案。
        :param aspect_ratio: 演示文稿的宽高比 ('16:9' 或 '4:3')。
        """
        self.plan = plan
        self.prs = Presentation()
        self.aspect_ratio = aspect_ratio # 存储宽高比
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
        logging.info(f"PresentationBuilder已为 {self.aspect_ratio} 演示文稿初始化。")

    def _apply_master_slide_styles(self):
        """应用全局母版样式。"""
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
            # --- [核心修改] 根据宽高比设置幻灯片尺寸 ---
            if self.aspect_ratio == "4:3":
                self.prs.slide_width = px_to_emu(1024)
                self.prs.slide_height = px_to_emu(768)
                logging.info("已将演示文稿尺寸设置为 4:3 (1024x768)。")
            else: # 默认为 16:9
                self.prs.slide_width = px_to_emu(1280)
                self.prs.slide_height = px_to_emu(720)
                logging.info("已将演示文稿尺寸设置为 16:9 (1280x720)。")

            self._apply_master_slide_styles()

            pages = self.plan.get('pages', [])
            total_pages = len(pages)
            for i, page_data in enumerate(pages):
                logging.info(f"--- 正在构建页面 {i + 1}/{total_pages} ---")
                self.slide_renderer.render_slide(page_data, self.image_service)

            self.prs.save(output_path)
            logging.info(f"演示文稿已成功保存至 {output_path}")
        except Exception as e:
            logging.error(f"构建演示文稿过程中发生严重错误: {e}", exc_info=True)
            raise
