# ppt_builder/slide_renderer.py
import logging
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from ppt_builder import elements
from image_service import ImageService # Assuming ImageService is available

class SlideRenderer:
    """
    Renders a single slide based on a detailed JSON plan of elements.
    """
    def __init__(self, prs, image_service: ImageService):
        self.prs = prs
        self.image_service = image_service
        logging.info("SlideRenderer initialized.")

    def render_slide(self, slide_data: dict):
        """
        Creates a new slide and adds elements to it based on slide_data.
        """
        try:
            # Use a blank slide layout for maximum flexibility
            blank_slide_layout = self.prs.slide_layouts[6] # Layout 6 is typically 'Blank'
            slide = self.prs.slides.add_slide(blank_slide_layout)
            logging.info(f"Created new slide: '{slide_data.get('page_title', 'Untitled Slide')}'")

            for element in slide_data.get('elements', []):
                element_type = element.get('type')
                if element_type == 'text_box':
                    elements.add_text_box(slide, element)
                elif element_type == 'image':
                    image_keyword = element.get('image_keyword')
                    if image_keyword:
                        image_path = self.image_service.generate_image(image_keyword)
                        if image_path:
                            elements.add_image(slide, image_path, element)
                        else:
                            logging.warning(f"Could not generate image for keyword: '{image_keyword}'. Skipping image element.")
                    else:
                        logging.warning("Image element missing 'image_keyword'. Skipping.")
                elif element_type == 'shape':
                    elements.add_shape(slide, element)
                elif element_type == 'chart':
                    elements.add_chart(slide, element)
                else:
                    logging.warning(f"Unsupported element type: {element_type}. Skipping element.")
        except Exception as e:
            logging.error(f"Error rendering slide '{slide_data.get('page_title', 'Untitled')}': {e}", exc_info=True)

