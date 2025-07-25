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
    "designer": "gemini-2.5-flash",
    "writer": "gemini-2.5-flash",
    "planner": "glm-4-air",
    "inspector": "glm-4-air",
}
MAX_LAYOUT_RETRIES = 3


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

VALID_ICON_KEYWORDS = ["activity", "airplay", "alert-circle", "alert-octagon", "alert-triangle", "align-center", "align-justify", "align-left", "align-right", "anchor", "aperture", "archive", "arrow-down-circle", "arrow-down-left", "arrow-down-right", "arrow-down", "arrow-left-circle", "arrow-left", "arrow-right-circle", "arrow-right", "arrow-up-circle", "arrow-up-left", "arrow-up-right", "arrow-up", "at-sign", "award", "bar-chart-2", "bar-chart", "battery-charging", "battery", "bell-off", "bell", "bluetooth", "bold", "book-open", "book", "bookmark", "box", "briefcase", "calendar", "camera-off", "camera", "cast", "check-circle", "check-square", "check", "chevron-down", "chevron-left", "chevron-right", "chevron-up", "chevrons-down", "chevrons-left", "chevrons-right", "chevrons-up", "chrome", "circle", "clipboard", "clock", "cloud-drizzle", "cloud-lightning", "cloud-off", "cloud-rain", "cloud-snow", "cloud", "code", "codepen", "codesandbox", "coffee", "columns", "command", "compass", "copy", "corner-down-left", "corner-down-right", "corner-left-down", "corner-left-up", "corner-right-down", "corner-right-up", "corner-up-left", "corner-up-right", "cpu", "credit-card", "crop", "crosshair", "database", "delete", "disc", "divide-circle", "divide-square", "divide", "dollar-sign", "download-cloud", "download", "dribbble", "droplet", "edit-2", "edit-3", "edit", "external-link", "eye-off", "eye", "facebook", "fast-forward", "feather", "figma", "file-minus", "file-plus", "file-text", "file", "film", "filter", "flag", "folder-minus", "folder-plus", "folder", "framer", "frown", "gift", "git-branch", "git-commit", "git-merge", "git-pull-request", "github", "gitlab", "globe", "grid", "hard-drive", "hash", "headphones", "heart", "help-circle", "hexagon", "home", "image", "inbox", "info", "instagram", "italic", "key", "layers", "layout", "life-buoy", "link-2", "link", "linkedin", "list", "loader", "lock", "log-in", "log-out", "mail", "map-pin", "map", "maximize-2", "maximize", "meh", "menu", "message-circle", "message-square", "mic-off", "mic", "minimize-2", "minimize", "minus-circle", "minus-square", "minus", "monitor", "moon", "more-horizontal", "more-vertical", "mouse-pointer", "move", "music", "navigation-2", "navigation", "octagon", "package", "paperclip", "pause-circle", "pause", "pen-tool", "percent", "phone-call", "phone-forwarded", "phone-incoming", "phone-missed", "phone-off", "phone-outgoing", "phone", "pie-chart", "play-circle", "play", "plus-circle", "plus-square", "plus", "pocket", "power", "printer", "radio", "refresh-ccw", "refresh-cw", "repeat", "rewind", "rotate-ccw", "rotate-cw", "rss", "save", "scissors", "search", "send", "server", "settings", "share-2", "share", "shield-off", "shield", "shopping-bag", "shopping-cart", "shuffle", "sidebar", "skip-back", "skip-forward", "slack", "slash", "sliders", "smartphone", "smile", "speaker", "square", "star", "stop-circle", "sun", "sunrise", "sunset", "table", "tablet", "tag", "target", "terminal", "thermometer", "thumbs-down", "thumbs-up", "toggle-left", "toggle-right", "tool", "trash-2", "trash", "trello", "trending-down", "trending-up", "triangle", "truck", "tv", "twitch", "twitter", "type", "umbrella", "underline", "unlock", "upload-cloud", "upload", "user-check", "user-minus", "user-plus", "user-x", "user", "users", "video-off", "video", "voicemail", "volume-1", "volume-2", "volume-x", "volume", "watch", "wifi-off", "wifi", "wind", "x-circle", "x-octagon", "x-square", "x", "youtube", "zap-off", "zap", "zoom-in", "zoom-out"]
