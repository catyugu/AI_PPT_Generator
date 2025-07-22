import logging
from pptx import Presentation
from pptx.dml.color import RGBColor # 为最终的回退方案导入
from ppt_builder.slide_renderer import SlideRenderer
from ppt_builder.styles import PresentationStyle, px_to_emu, hex_to_rgb # 导入所需的 hex_to_rgb
from image_service import ImageService

class PresentationBuilder:
    """根据AI生成的计划构建完整的演示文稿。"""

    def __init__(self, plan: dict):
        self.plan = plan
        self.prs = Presentation()
        self.style_manager = PresentationStyle(plan)
        self.image_service = ImageService()
        self.slide_renderer = SlideRenderer(self.prs, self.image_service, self.style_manager)
        logging.info("PresentationBuilder已使用AI方案和StyleManager初始化。")

    def _apply_master_slide_styles(self):
        """
        [已优化] 应用全局母版样式，如背景。
        优化后的逻辑：
        1. 优先使用 master_slide.background 中直接定义的图片。
        2. 如果无图片，则使用 master_slide.background 中直接定义的颜色。
        3. 如果以上均未定义，则使用 color_palette 中的 'background' 颜色作为默认。
        4. 如果图片生成失败，也会回退到颜色背景。
        """
        master = self.prs.slide_masters[0]
        master_data = self.plan.get('master_slide', {})
        background_info = master_data.get('background', {})
        fill = master.background.fill

        try:
            # 优先级 1: 检查母版是否指定了背景图片
            if keyword := background_info.get('image_keyword'):
                logging.info(f"正在为母版背景搜索图片: '{keyword}'")
                bg_image_path = self.image_service.generate_image(keyword)
                if bg_image_path:
                    fill.stretch(bg_image_path)
                    logging.info(f"已将图片背景应用于母版: {bg_image_path}")
                    return  # 背景设置成功，直接返回

                logging.warning("无法生成母版背景图片，将回退到颜色背景。")

            # 优先级 2: 检查母版是否指定了具体的背景颜色 (也是图片失败时的回退点)
            if specific_bg_color_hex := background_info.get('color'):
                fill.solid()
                # **[已修复]** 直接使用 hex_to_rgb 转换从计划中获取的颜色代码
                fill.fore_color.rgb = hex_to_rgb(specific_bg_color_hex)
                logging.info(f"已将母版背景设置为指定颜色: {specific_bg_color_hex}")
            # 优先级 3: 如果以上均未指定或失败，使用全局调色板的背景色
            else:
                fill.solid()
                fill.fore_color.rgb = self.style_manager.get_color('background')
                logging.info("已将母版背景设置为调色板中的默认背景色。")

        except Exception as e:
            logging.error(f"应用母版背景时出错: {e}", exc_info=True)
            # 终极保险措施：如果一切都失败了，设置为安全的白色背景
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)


    def build_presentation(self, output_path: str):
        """构建并保存演示文稿。"""
        try:
            # 设置演示文稿尺寸为 16:9 (1280x720)
            self.prs.slide_width = px_to_emu(1280)
            self.prs.slide_height = px_to_emu(720)

            # 首先应用母版样式
            self._apply_master_slide_styles()

            pages = self.plan.get('pages', [])
            total_pages = len(pages)
            for i, page_data in enumerate(pages):
                logging.info(f"--- 正在构建页面 {i + 1}/{total_pages} ---")
                self.slide_renderer.render_slide(page_data)

            self.prs.save(output_path)
            logging.info(f"演示文稿已成功保存至 {output_path}")
        except Exception as e:
            logging.error(f"构建演示文稿过程中发生严重错误: {e}", exc_info=True)
            raise