import logging
import os

from pptx import Presentation
from ppt_builder import elements
from ppt_builder.styles import PresentationStyle, px_to_emu

ELEMENT_LAYER_ORDER = {
    'image': 0,
    'shape': 1,
    'chart': 2,
    'table': 2,
    'text_box': 3,  # 文本框永远在最上层
    'default': 0  # 为未知类型提供默认层级
}


class SlideRenderer:
    """负责将单页幻灯片的数据渲染到演示文稿中。"""

    def __init__(self, prs: Presentation, style_manager: PresentationStyle, background_image_path: str | None):
        """
        初始化渲染器。
        :param prs: 演示文稿对象。
        :param style_manager: 全局样式管理器。
        :param background_image_path: 全局背景图片的路径，如果无则为None。
        """
        self.prs = prs
        self.style_manager = style_manager
        self.background_image_path = background_image_path
        logging.info("SlideRenderer已使用样式管理器和背景信息初始化。")

    def _add_background_image(self, slide, image_path: str):
        """
        [FIXED] Adds a background image and moves it to the back.
        Now includes checks to prevent errors with invalid paths.
        """
        if not image_path:
            # 静默处理，因为没有全局背景是正常情况
            return
        if not os.path.exists(image_path):
            logging.error(f"Background image not found at path: {image_path}. Skipping.")
            return

        try:
            picture = slide.shapes.add_picture(
                image_path, 0, 0,
                width=self.prs.slide_width,
                height=self.prs.slide_height
            )
            # 将图片移动到最底层
            pic_element = picture.element
            slide.shapes._spTree.remove(pic_element)
            slide.shapes._spTree.insert(0, pic_element)

            logging.info(f"Successfully added background image: {image_path}")
        except Exception as e:
            logging.error(f"An unexpected error occurred while adding background image '{image_path}': {e}",
                          exc_info=True)

    def render_slide(self, slide_data: dict, image_service):
        """根据给定的数据渲染一张幻灯片。"""
        blank_slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(blank_slide_layout)

        # 首先添加全局背景图片（如果存在）
        self._add_background_image(slide, self.background_image_path)

        # ===================== 核心修改：元素排序 =====================
        elements_to_render = slide_data.get('elements', [])

        # 使用我们定义的图层顺序对元素列表进行排序
        # key函数获取每个元素的类型，并在ELEMENT_LAYER_ORDER中查找对应的层级数字
        # get的第二个参数是默认值，确保即使AI生成了未知的元素类型也不会报错
        elements_to_render.sort(
            key=lambda e: ELEMENT_LAYER_ORDER.get(e.get('type'), ELEMENT_LAYER_ORDER['default'])
        )

        logging.info("页面元素已按图层顺序重排，渲染开始...")
        # ===================== 修改结束 ============================

        # 遍历排序后的列表进行渲染
        for element in elements_to_render:
            element_type = element.get('type')
            try:
                if element_type in ['text_box', 'text']:
                    elements.add_text_box(slide, element, self.style_manager)

                elif element_type == 'image':
                    if image_keyword := element.get('image_keyword'):
                        opacity = element.get('style', {}).get('opacity', 1.0)
                        image_path = image_service.generate_image(image_keyword, opacity)
                        if image_path:
                            elements.add_image(slide, image_path, element)
                        else:
                            logging.warning(f"无法为关键词生成图片: '{image_keyword}'。已跳过此元素。")
                    else:
                        logging.warning("图片元素缺少 'image_keyword'，已跳过。")

                elif element_type == 'shape':
                    elements.add_shape(slide, element, self.style_manager)

                elif element_type == 'chart':
                    elements.add_chart(slide, element, self.style_manager)

                elif element_type == 'table':
                    elements.add_table(slide, element, self.style_manager)

                else:
                    logging.warning(f"不支持的元素类型: '{element_type}'。")

            except Exception as e:
                logging.error(f"渲染类型为 '{element_type}' 的元素失败: {e}", exc_info=True)