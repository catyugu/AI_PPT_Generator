# config.py
import os
from dotenv import load_dotenv

# Load .env file variables into environment
load_dotenv()

# --- API Configuration ---
# It's highly recommended to use environment variables for security
ONEAPI_KEY = os.environ.get("ONEAPI_KEY", "YOUR_ONEAPI_KEY_HERE")
ONEAPI_BASE_URL = os.environ.get("ONEAPI_BASE_URL", "http://127.0.0.1:3000/v1")
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "YOUR_PEXELS_API_KEY_HERE")
MODEL_NAME = "gemini-2.5-flash"  # Or use a more powerful model like "gpt-4o"

# --- Output Configuration ---
OUTPUT_DIR = "AI_Generated_PPTs"

def get_env_variable(var_name: str) -> str | None:
    """
    从.env文件或环境变量中获取变量值。
    """
    load_dotenv()
    return os.getenv(var_name)

def get_openai_api_key() -> str | None:
    """
    获取OpenAI API密钥。
    """
    return get_env_variable("OPENAI_API_KEY")

def get_pexels_api_key() -> str | None:
    """
    获取Pexels API密钥。
    """
    return get_env_variable("PEXELS_API_KEY")

