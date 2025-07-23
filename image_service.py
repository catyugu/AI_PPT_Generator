import logging
import requests
from io import BytesIO
from pexels_api import API as PexelsAPI
import config
import tempfile
from PIL import Image
import os
import time  # 引入 time 模块

# 定义临时文件目录
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'temp')


class ImageService:
    """处理图片获取、应用效果（如透明度）并保存为临时文件。"""

    def __init__(self):
        self.pexels_client = None
        pexels_key = config.get_api_key("PEXELS_API_KEY")
        if pexels_key and pexels_key != "YOUR_PEXELS_API_KEY_HERE":
            try:
                self.pexels_client = PexelsAPI(pexels_key)
                logging.info("Pexels客户端初始化成功。")
            except Exception as e:
                logging.error(f"初始化Pexels客户端失败: {e}")
        else:
            logging.warning("未配置Pexels API密钥，将使用占位图片服务。")

    def _fetch_from_pexels(self, keyword: str) -> BytesIO | None:
        """
        [已优化] 从Pexels获取图片，带有重试机制。
        最多重试3次，每次间隔3秒。
        """
        if not self.pexels_client:
            return None

        max_retries = 3
        for attempt in range(max_retries):
            try:
                logging.info(f"正在从Pexels搜索 '{keyword}' (尝试 {attempt + 1}/{max_retries})...")
                search_results = self.pexels_client.search(keyword, page=1, results_per_page=1)
                if photos := search_results.get('photos'):
                    photo_url = photos[0].get('src', {}).get('large2x')
                    if photo_url:
                        response = requests.get(photo_url, timeout=20)
                        response.raise_for_status()
                        logging.info(f"Pexels图片 '{keyword}' 获取成功。")
                        return BytesIO(response.content)

                logging.warning(f"未在Pexels上找到 '{keyword}' 的图片。")
                return None  # 如果搜索成功但没有图片，直接返回None，无需重试

            except Exception as e:
                logging.warning(f"Pexels搜索 '{keyword}' 失败 (尝试 {attempt + 1}): {e}。")
                if attempt < max_retries - 1:
                    logging.info("将在3秒后重试...")
                    time.sleep(3)
                    self.pexels_client = None
                    pexels_key = config.get_api_key("PEXELS_API_KEY")
                    if pexels_key and pexels_key != "YOUR_PEXELS_API_KEY_HERE":
                        try:
                            self.pexels_client = PexelsAPI(pexels_key)
                            logging.info("Pexels客户端初始化成功。")
                        except Exception as e:
                            logging.error(f"初始化Pexels客户端失败: {e}")
                    else:
                        logging.warning("未配置Pexels API密钥，将使用占位图片服务。")
                else:
                    logging.error(f"Pexels搜索 '{keyword}' 在 {max_retries} 次尝试后彻底失败。")

        return None

    def _fetch_from_fallback(self, keyword: str) -> BytesIO | None:
        """从备用服务获取占位图片。"""
        try:
            logging.info(f"正在为 '{keyword}' 使用占位图片。")
            placeholder_url = f"https://placehold.co/1280x720.png?text={keyword.replace(' ', '+')}&font=lato"
            response = requests.get(placeholder_url, timeout=10)
            response.raise_for_status()
            return BytesIO(response.content)
        except Exception as e:
            logging.error(f"获取 '{keyword}' 的占位图片失败: {e}")
            return None

    def generate_image(self, keyword: str, opacity: float = 1.0) -> str | None:
        """
        获取图片，应用透明度，保存到临时文件并返回路径。
        """
        image_stream = self._fetch_from_pexels(keyword) or self._fetch_from_fallback(keyword)

        if not image_stream:
            return None

        try:
            os.makedirs(TEMP_DIR, exist_ok=True)

            img = Image.open(image_stream).convert("RGBA")

            if opacity < 1.0:
                alpha = img.getchannel('A')
                new_alpha = alpha.point(lambda p: p * opacity)
                img.putalpha(new_alpha)

            with tempfile.NamedTemporaryFile(delete=False, suffix='.png', dir=TEMP_DIR) as temp_file:
                img.save(temp_file, format='PNG')
                logging.info(f"已为 '{keyword}' (透明度={opacity}) 生成并保存临时图片: {temp_file.name}")
                return temp_file.name
        except Exception as e:
            logging.error(f"处理或保存图片到临时文件时出错: {e}", exc_info=True)
            return None