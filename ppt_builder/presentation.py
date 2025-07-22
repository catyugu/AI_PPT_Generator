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
        self.slide_renderer = SlideRenderer(self.prs, self.image_service, self.style_manager)
        logging.info("PresentationBuilder已使用AI方案和StyleManager初始化。")

    def _move_shape_to_background(self, master, shape):
        """**[新功能]** 将一个形状移动到母版形状树的底部（视觉上的最底层）。"""
        try:
            spTree = master.shapes._spTree
            # 将形状的XML元素从当前位置移除，并插入到列表的开头
            spTree.insert(0, spTree.pop(spTree.index(shape.element)))
            logging.info("已成功将背景图片移动到母版底层。")
        except Exception as e:
            logging.error(f"移动形状到背景时出错: {e}", exc_info=True)

    def _apply_master_slide_styles(self):
        """
        [已优化] 应用全局母版样式，如背景。
        现在使用正确的方法添加可拉伸的背景图片。
        """
        master = self.prs.slide_masters[0]
        master_data = self.plan.get('master_slide', {})
        background_info = master_data.get('background', {})

        try:
            # 优先级 1: 检查母版是否指定了背景图片
            if keyword := background_info.get('image_keyword'):
                logging.info(f"正在为母版背景搜索图片: '{keyword}'")
                bg_image_path = self.image_service.generate_image(keyword)
                if bg_image_path:
                    # **[已修复]** 正确的背景图片添加方式：添加为形状并置于底层
                    picture = master.shapes.add_picture(
                        bg_image_path,
                        px_to_emu(0), px_to_emu(0),
                        width=self.prs.slide_width,
                        height=self.prs.slide_height
                    )
                    self._move_shape_to_background(master, picture)
                    return  # 背景图片设置成功，直接返回

                logging.warning("无法生成母版背景图片，将回退到颜色背景。")

            # 优先级 2: 检查母版是否指定了具体的背景颜色 (也是图片失败时的回退点)
            fill = master.background.fill
            if specific_bg_color_hex := background_info.get('color'):
                fill.solid()
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
            master.background.fill.solid()
            master.background.fill.fore_color.rgb = RGBColor(255, 255, 255)

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