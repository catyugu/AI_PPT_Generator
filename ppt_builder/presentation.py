# ppt_builder/presentation.py
import logging
from pptx import Presentation
from pptx.util import Inches
from config import DEFAULT_PPT_WIDTH_INCHES, DEFAULT_PPT_HEIGHT_INCHES
from ppt_builder.styles import PresentationStyle
import ppt_builder.layouts as layouts

# The Layout Registry (or Router)
# This dictionary maps the `page_type` string from the AI to the actual function
# that builds the slide. This is the key to making the system extensible.
LAYOUT_REGISTRY = {
    'cover': layouts.create_cover_slide,
    'section_header': layouts.create_section_header_slide,
    'title_content': layouts.create_title_content_slide,
    'image_text_split': layouts.create_image_text_split_slide,
    'full_bleed_image_quote': layouts.create_full_bleed_image_quote_slide,
    'three_points_icons': layouts.create_three_points_icons_slide,
    'bar_chart': layouts.create_bar_chart_slide,
    'timeline': layouts.create_timeline_slide,
    'process_flow': layouts.create_process_flow_slide,
    'icon_grid': layouts.create_icon_grid_slide,
    'team_intro': layouts.create_team_intro_slide,
    'thank_you': layouts.create_cover_slide,
}


def build_presentation(data: dict, output_path: str):
    """
    Constructs a presentation using a style object and the layout registry.
    """
    if not data or 'pages' not in data:
        logging.error("AI-generated data is incomplete or invalid. Aborting.")
        return

    prs = Presentation()
    prs.slide_width = Inches(DEFAULT_PPT_WIDTH_INCHES)
    prs.slide_height = Inches(DEFAULT_PPT_HEIGHT_INCHES)
    style = PresentationStyle(data)

    for i, page_data in enumerate(data['pages']):
        slide_layout = prs.slide_layouts[6]  # Use a blank layout
        slide = prs.slides.add_slide(slide_layout)

        # Set background color for every slide
        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = style.background

        page_type = page_data.get('page_type')
        layout_func = LAYOUT_REGISTRY.get(page_type)

        if layout_func:
            try:
                logging.info(f"Creating slide {i + 1} with layout: '{page_type}'")
                layout_func(slide, page_data, style, prs)
            except Exception as e:
                logging.error(f"Failed to create slide {i + 1} ('{page_type}'): {e}", exc_info=True)
        else:
            logging.warning(f"No layout function found for page_type '{page_type}'. Using default.")
            layouts.create_title_content_slide(slide, page_data, style, prs)

    try:
        prs.save(output_path)
        logging.info(f"Presentation saved to: {output_path}")
    except Exception as e:
        logging.error(f"Failed to save presentation to '{output_path}': {e}")