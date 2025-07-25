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

# --- [核心修改] 流水线模型配置 ---
# 为流水线的不同阶段配置不同的“专家”模型
# - designer: 负责顶层创意和视觉风格，建议使用最强大的模型 (如 gpt-4o, claude-3-opus)
# - writer: 负责内容大纲撰写，建议使用知识渊博、逻辑性强的模型 (如 gemini-1.5-pro, gpt-4-turbo)
# - planner: 负责单页布局，任务结构化，可以使用更快速、经济的模型 (如 gemini-1.5-flash, gpt-3.5-turbo)
MODEL_CONFIG = {
    "designer": "glm-4-plus",
    "writer": "glm-4-plus",
    "planner": "glm-4-plus"
}


# --- 输出配置 ---
OUTPUT_DIR = "AI_Generated_PPTs"
ICON_DIR = "assets/icons"

# --- Helper Functions (可选，保持清晰) ---
def get_env_variable(var_name: str, default: str = None) -> str | None:
    """从.env文件或环境变量中安全地获取变量值。"""
    load_dotenv()
    return os.getenv(var_name, default)

def get_api_key(key_name: str) -> str | None:
    """获取指定API密钥的通用函数。"""
    return get_env_variable(key_name)

# [修改] SYSTEM_PROMPT 可以保持原样，或在每个阶段的Prompt中被更具体的内容覆盖
SYSTEM_PROMPT = """
你是一位顶级的AI专家，擅长按要求完成特定任务。
你的任务是根据用户提供的指令，生成一个严格符合格式要求的JSON对象。
"""