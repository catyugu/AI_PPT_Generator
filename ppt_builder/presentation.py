import logging
from pptx import Presentation
from ppt_builder.slide_renderer import SlideRenderer
from ppt_builder.styles import PresentationStyle
from image_service import ImageService


class PresentationBuilder:
    """
    根据AI生成的计划构建完整的演示文稿。
    """

    def __init__(self, plan: dict):
        """
        初始化构建器。

        Args:
            plan (dict): AI生成的演示文稿计划。
        """
        self.plan = plan
        self.prs = Presentation()
        # [核心优化] 初始化设计风格管理器
        self.style_manager = PresentationStyle(plan)

        self.image_service = ImageService()
        # [核心优化] 将 style_manager 传递给渲染器
        self.slide_renderer = SlideRenderer(self.prs, self.image_service, self.style_manager)
        logging.info("PresentationBuilder initialized with AI plan and StyleManager.")

    def _apply_master_slide_styles(self):
        """
        [新功能] 应用全局母版样式，如背景。
        """
        master = self.prs.slide_masters[0]
        master_data = self.plan.get('master_slide', {})

        # 设置背景
        background_info = master_data.get('background', {})
        if 'color' in background_info:
            try:
                master.background.fill.solid()
                master.background.fill.fore_color.rgb = self.style_manager.get_color('background')
                logging.info(f"Applied solid color background to master slide.")
            except Exception as e:
                logging.error(f"Failed to apply master background color: {e}")
        elif 'image_keyword' in background_info:
            logging.info(f"Searching for master background image with keyword: '{background_info['image_keyword']}'")
            # 背景图可以考虑增加透明度等效果，这里简化为直接使用
            bg_image_path = self.image_service.generate_image(background_info['image_keyword'])
            if bg_image_path:
                try:
                    master.background.fill.stretch(bg_image_path)
                    logging.info(f"Applied image background to master slide: {bg_image_path}")
                except Exception as e:
                    logging.error(f"Failed to apply master background image: {e}")
            else:
                logging.warning("Could not generate background image for master slide.")

        # 可以在此添加页脚、页码等母版元素的逻辑

    def build_presentation(self, output_path: str):
        """
        构建并保存演示文稿。

        Args:
            output_path (str): 输出的PPTX文件路径。
        """
        try:
            # 设置演示文稿尺寸 (16:9)
            self.prs.slide_width = 11614200  # 1280px
            self.prs.slide_height = 6534000  # 720px

            # [新功能] 首先应用母版样式
            self._apply_master_slide_styles()

            pages = self.plan.get('pages', [])
            for i, page_data in enumerate(pages):
                logging.info(f"Building page {i + 1}/{len(pages)}...")
                self.slide_renderer.render_slide(page_data)

            self.prs.save(output_path)
            logging.info(f"Presentation successfully saved to {output_path}")
        except Exception as e:
            logging.error(f"An critical error occurred during presentation building: {e}", exc_info=True)
            raise
