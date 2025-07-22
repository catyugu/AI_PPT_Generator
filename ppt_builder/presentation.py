# ppt_builder/presentation.py
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
import logging

from ppt_builder.slide_renderer import SlideRenderer
from image_service import ImageService # Import the ImageService class

class PresentationBuilder:
    """
    Builds a PowerPoint presentation based on a detailed AI-generated plan.
    This builder orchestrates the creation of slides and the addition of elements
    using the SlideRenderer.
    """
    def __init__(self, plan: dict):
        self.plan = plan
        self.prs = Presentation()
        # Set standard slide size (16:9 aspect ratio, roughly 10x7.5 inches at 96 DPI)
        # Default is usually 10x7.5 inches, which corresponds to 914400 EMUs x 685800 EMUs
        # python-pptx uses EMUs (English Metric Units) internally.
        # The AI's pixel coordinates are based on a 914x685 pixel canvas.
        # The default slide size in python-pptx is usually 10 inches x 7.5 inches,
        # which is 914400 EMUs x 685800 EMUs. This aligns well with a 96 DPI assumption
        # (96 pixels/inch * 10 inches = 960 pixels, but PowerPoint's internal pixel
        # mapping for shapes can be a bit different. The AI is using 914x685 as a reference).
        # We'll rely on the default slide size and the px_to_emu conversion in elements.py.

        self.image_service = ImageService() # Initialize ImageService
        self.slide_renderer = SlideRenderer(self.prs, self.image_service)
        logging.info("PresentationBuilder initialized with AI plan.")

    def _apply_global_styles(self):
        """Applies global styles like font pairing and color palette."""
        color_palette = self.plan.get('color_palette', {})
        # This is a simplified application. For full theme control,
        # you'd need to modify the presentation's theme XML.
        # Here, we'll just log the colors for now as direct application
        # to the overall presentation theme is complex without theme modification.
        if color_palette:
            logging.info(f"Applying color palette: {color_palette}")
            # Example: Set background color for all slides (this is a basic approach)
            # This would need to be done per slide or by modifying the master slide.
            # For now, we'll just log this.
            pass

        font_pairing = self.plan.get('font_pairing', {})
        if font_pairing:
            logging.info(f"Applying font pairing: {font_pairing}")
            # Similar to colors, applying global fonts requires theme modification
            # or setting fonts per text box, which is handled by elements.py.
            pass

    def build_presentation(self, output_path: str):
        """
        Builds the presentation by iterating through the plan's pages
        and rendering each slide.
        """
        try:
            self._apply_global_styles()

            pages = self.plan.get('pages', [])
            for i, page_data in enumerate(pages):
                logging.info(f"Building page {i+1}: {page_data.get('page_title', 'Untitled')}")
                self.slide_renderer.render_slide(page_data)

            self.prs.save(output_path)
            logging.info(f"Presentation successfully saved to {output_path}")
        except Exception as e:
            logging.error(f"Error building presentation: {e}", exc_info=True)
            raise # Re-raise to be caught in main.py for user feedback

