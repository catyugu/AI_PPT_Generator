# config.py
import os
from dotenv import load_dotenv

# Load .env file variables into environment
load_dotenv()

# --- Page Number Configuration ---
NUM_PAGES = 16 # <-- Set the desired number of pages here

# --- API Configuration ---
# It's highly recommended to use environment variables for security
ONEAPI_KEY = os.environ.get("ONEAPI_KEY", "YOUR_ONEAPI_KEY_HERE")
ONEAPI_BASE_URL = os.environ.get("ONEAPI_BASE_URL", "http://127.0.0.1:3000/v1")
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "YOUR_PEXELS_API_KEY_HERE")
MODEL_NAME = "glm-4-flash"  # Or use a more powerful model like "gpt-4o"

# --- Output Configuration ---
OUTPUT_DIR = "AI_Generated_PPTs"

# --- Presentation Defaults ---
DEFAULT_PPT_WIDTH_INCHES = 16
DEFAULT_PPT_HEIGHT_INCHES = 9