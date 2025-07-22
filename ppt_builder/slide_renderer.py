import logging
from pptx import Presentation
from ppt_builder import elements
from ppt_builder.styles import PresentationStyle, px_to_emu


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

    def _add_background_image(self, slide):
        """如果存在全局背景图片，则将其添加到当前幻灯片并置于底层。"""
        if not self.background_image_path:
            return

        try:
            # 添加图片，并使其铺满整个幻灯片
            picture = slide.shapes.add_picture(
                self.background_image_path,
                px_to_emu(0), px_to_emu(0),
                width=self.prs.slide_width,
                height=self.prs.slide_height
            )

            # **[核心]** 将图片移动到Z轴的最底层，成为背景
            spTree = slide.shapes._spTree
            spTree.insert(0, spTree.pop(spTree.index(picture.element)))
            logging.info("已在当前页面添加并置底全局背景图片。")

        except Exception as e:
            logging.error(f"在幻灯片上添加背景图片时出错: {e}", exc_info=True)

    def render_slide(self, slide_data: dict, image_service):
        """根据给定的数据渲染一张幻灯片。"""
        blank_slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(blank_slide_layout)

        # **[新逻辑]** 首先添加背景图片
        self._add_background_image(slide)

        for element in slide_data.get('elements', []):
            element_type = element.get('type')
            try:
                if element_type in ['text_box', 'text']:
                    elements.add_text_box(slide, element, self.style_manager)

                elif element_type == 'image':
                    if image_keyword := element.get('image_keyword'):
                        # 注意：ImageService 现在从外部传入
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