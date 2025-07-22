# image_service.py
import logging
import requests
from io import BytesIO
from pexels_api import API as PexelsAPI
import config
import os
import tempfile
from PIL import Image # Import Pillow for image manipulation

class ImageService:
    """
    Handles image fetching from Pexels or a fallback service,
    applies opacity if specified, and saves them to temporary files
    for use in the presentation.
    """
    def __init__(self):
        self.pexels_client = None
        try:
            if config.PEXELS_API_KEY and config.PEXELS_API_KEY != "YOUR_PEXELS_API_KEY_HERE":
                self.pexels_client = PexelsAPI(config.PEXELS_API_KEY)
                logging.info("Pexels client initialized.")
            else:
                logging.warning("Pexels API key not configured. Using placeholders.")
        except Exception as e:
            logging.error(f"Failed to initialize Pexels client: {e}")
            self.pexels_client = None

    def generate_image(self, keyword: str, opacity: float = 1.0) -> str | None:
        """
        Fetches an image from Pexels or a fallback service, applies opacity,
        saves it to a temporary file, and returns the path to that file.
        Opacity should be a float between 0.0 (fully transparent) and 1.0 (fully opaque).
        """
        image_stream = None
        if self.pexels_client:
            try:
                logging.info(f"Searching Pexels for '{keyword}'...")
                search_results = self.pexels_client.search(keyword, page=1, results_per_page=1)
                photos_list = search_results.get('photos', [])

                if photos_list:
                    photo_url = photos_list[0].get('src', {}).get('large2x')
                    if photo_url:
                        response = requests.get(photo_url, timeout=20)
                        response.raise_for_status()
                        image_stream = BytesIO(response.content)

                if not image_stream:
                    logging.warning(f"No Pexels image found for '{keyword}'. Trying fallback.")

            except Exception as e:
                logging.error(f"Pexels search failed for '{keyword}': {e}. Trying fallback.")

        # Fallback to placeholder if Pexels fails or is not configured
        if not image_stream:
            try:
                logging.info(f"Using placeholder image for '{keyword}'.")
                placeholder_url = f"https://placehold.co/1920x1080.png?text={keyword.replace(' ', '+')}&font=lato"
                response = requests.get(placeholder_url, timeout=10)
                response.raise_for_status()
                image_stream = BytesIO(response.content)
            except Exception as e:
                logging.error(f"Fallback placeholder failed for '{keyword}': {e}")
                return None

        if image_stream:
            try:
                # Open image with Pillow
                img = Image.open(image_stream).convert("RGBA") # Ensure it has an alpha channel

                # Apply opacity if less than 1.0
                if opacity < 1.0:
                    alpha = img.split()[-1] # Get the alpha channel
                    alpha = alpha.point(lambda p: p * opacity) # Apply opacity to alpha
                    img.putalpha(alpha) # Put the modified alpha back

                # Create a temporary file to save the image
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                img.save(temp_file, format='PNG') # Save with alpha channel
                temp_file.close()
                logging.info(f"Saved image for '{keyword}' (opacity={opacity}) to temporary file: {temp_file.name}")
                return temp_file.name
            except Exception as e:
                logging.error(f"Error processing or saving image to temporary file: {e}")
                return None
        return None

