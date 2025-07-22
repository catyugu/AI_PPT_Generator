# image_service.py
import logging
import requests
from io import BytesIO
from pexels_api import API as PexelsAPI
import config

# --- Initialize Client ---
try:
    pexels_client = PexelsAPI(config.PEXELS_API_KEY) if config.PEXELS_API_KEY != "YOUR_PEXELS_API_KEY_HERE" else None
    if pexels_client:
        logging.info("Pexels client initialized.")
    else:
        logging.warning("Pexels API key not configured. Using placeholders.")
except Exception as e:
    logging.error(f"Failed to initialize Pexels client: {e}")
    pexels_client = None


def get_image_stream(keyword: str) -> BytesIO | None:
    """
    Fetches an image from Pexels or a fallback service and returns it as a BytesIO stream.
    """
    if pexels_client:
        try:
            logging.info(f"Searching Pexels for '{keyword}'...")
            search_results = pexels_client.search(keyword, page=1, results_per_page=1)
            photos_list = search_results.get('photos', [])

            if photos_list:
                photo_url = photos_list[0].get('src', {}).get('large2x')
                if photo_url:
                    response = requests.get(photo_url, timeout=20)
                    response.raise_for_status()
                    return BytesIO(response.content)

            logging.warning(f"No Pexels image found for '{keyword}'. Trying fallback.")

        except Exception as e:
            logging.error(f"Pexels search failed for '{keyword}': {e}. Trying fallback.")

    # Fallback to placeholder if Pexels fails or is not configured
    try:
        logging.info(f"Using placeholder image for '{keyword}'.")
        placeholder_url = f"https://placehold.co/1920x1080.png?text={keyword.replace(' ', '+')}&font=lato"
        response = requests.get(placeholder_url, timeout=10)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        logging.error(f"Fallback placeholder failed for '{keyword}': {e}")
        return None