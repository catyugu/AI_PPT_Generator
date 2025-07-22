import logging
from pptx import Presentation
from ppt_builder import elements
from image_service import ImageService
from ppt_builder.styles import PresentationStyle


class SlideRenderer:
    """
    负责将单页幻灯片的数据渲染到演示文稿中。
    """

    def __init__(self, prs: Presentation, image_service: ImageService, style_manager: PresentationStyle):
        """
        初始化渲染器。

        Args:
            prs (Presentation): python-pptx的Presentation对象。
            image_service (ImageService): 用于获取图片的图片服务。
            style_manager (PresentationStyle): 全局样式管理器。
        """
        self.prs = prs
        self.image_service = image_service
        self.style_manager = style_manager
        logging.info("SlideRenderer initialized with style manager.")

    def render_slide(self, slide_data: dict):
        """
        根据给定的数据渲染一张幻灯片。

        Args:
            slide_data (dict): 包含单页所有元素信息的字典。
        """
        blank_slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(blank_slide_layout)

        for element in slide_data.get('elements', []):
            element_type = element.get('type')
            try:
                # [关键修复] 兼容 'text' 和 'text_box' 两种类型
                if element_type == 'text_box' or element_type == 'text':
                    elements.add_text_box(slide, element, self.style_manager)

                elif element_type == 'image':
                    image_keyword = element.get('image_keyword')
                    if image_keyword:
                        logging.info(f"Generating image for keyword: '{image_keyword}'")
                        image_path = self.image_service.generate_image(image_keyword)
                        if image_path:
                            elements.add_image(slide, image_path, element)
                        else:
                            logging.warning(
                                f"Could not generate image for keyword: '{image_keyword}'. Skipping element.")
                    else:
                        logging.warning("Image element missing 'image_keyword'.")

                elif element_type == 'shape':
                    elements.add_shape(slide, element, self.style_manager)

                elif element_type == 'chart':
                    elements.add_chart(slide, element, self.style_manager)

                elif element_type == 'table':
                    elements.add_table(slide, element, self.style_manager)

                else:
                    logging.warning(f"Unsupported element type: '{element_type}'.")

            except Exception as e:
                logging.error(f"Failed to render element of type '{element_type}': {e}", exc_info=True)
