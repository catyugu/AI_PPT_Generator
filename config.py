# config.py
import os
from dotenv import load_dotenv

# 加载 .env 文件变量到环境中
load_dotenv()

# --- API 配置 ---
# 强烈建议使用环境变量以保障安全
ONEAPI_KEY = os.environ.get("ONEAPI_KEY", "YOUR_ONEAPI_KEY_HERE")
ONEAPI_BASE_URL = os.environ.get("ONEAPI_BASE_URL", "http://127.0.0.1:3000/v1")
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY", "YOUR_PEXELS_API_KEY_HERE")
MODEL_NAME = "qwen-plus"  # 您可以换成更强大的模型，如 "gpt-4o"

# --- 输出配置 ---
OUTPUT_DIR = "AI_Generated_PPTs"

# --- Helper Functions (可选，保持清晰) ---
def get_env_variable(var_name: str, default: str = None) -> str | None:
    """从.env文件或环境变量中安全地获取变量值。"""
    load_dotenv()
    return os.getenv(var_name, default)

def get_api_key(key_name: str) -> str | None:
    """获取指定API密钥的通用函数。"""
    return get_env_variable(key_name)