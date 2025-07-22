# ai_service.py
import json
import logging
from openai import OpenAI
import config

# --- Initialize Client ---
try:
    ai_client = OpenAI(api_key=config.ONEAPI_KEY, base_url=config.ONEAPI_BASE_URL)
    logging.info("OpenAI client initialized.")
except Exception as e:
    logging.error(f"Failed to initialize OpenAI client: {e}")
    ai_client = None

def generate_presentation_plan(theme: str, num_pages: int) -> dict | None:
    """
    [ENHANCED V3] Generates a PPT plan with flexible, AI-driven layout options.
    """
    if not ai_client:
        logging.error("AI client not initialized. Cannot generate plan.")
        return None

    # V3 Prompt: Introduces a "layout_options" object for dynamic control.
    prompt = f"""
    You are an expert presentation designer that produces structured JSON output. Your task is to create a plan for a presentation on the theme "{theme}".

    **Your response MUST be a single, raw JSON object and nothing else.**

    ---
    **## JSON STRUCTURE REQUIREMENTS**

    1.  **Root Object**: Must contain keys: `design_concept`, `font_pairing`, `color_palette`, `pages`.
    2.  **Language**: All user-facing text (`design_concept`, `title`, `content`) MUST be in **Chinese**.
    3.  **Keywords**: All machine-facing keywords (`image_keyword`, `icon_keyword`) MUST be in **English**.
    4.  **Page Count**: The `pages` list must contain exactly **{num_pages}** page objects.
    5.  **`color_palette`**: Must be a JSON **object** with keys: "primary", "secondary", "background", "text", "accent".
    6.  **`layout_options`**: An **optional** object inside a page object to control style.
        -   `image_position`: "left" | "right"
        -   `split_ratio`: "40/60" | "50/50" | "60/40"
        -   `text_alignment`: "left" | "center" | "right"
        -   `columns`: 1 | 2 | 3
        -   `overlay_opacity`: A number from 0.2 to 0.8

    ---
    **## EXAMPLE JSON OUTPUT**

    ```json
    {{
      "design_concept": "简约科技，数据驱动未来",
      "font_pairing": {{
        "heading": "Dengxian",
        "body": "Microsoft YaHei"
      }},
      "color_palette": {{
        "primary": "#0A3D62",
        "secondary": "#3C6382",
        "background": "#F5F5F5",
        "text": "#333333",
        "accent": "#FDCB6E"
      }},
      "pages": [
        {{
          "page_type": "cover",
          "title": "企业数字化转型解决方案",
          "content": [
            "引领您的业务进入智能时代"
          ],
          "image_keyword": "abstract digital transformation",
          "layout_options": {{
            "text_alignment": "center",
            "overlay_opacity": 0.6
          }}
        }},
        {{
          "page_type": "image_text_split",
          "title": "核心挑战与机遇",
          "content": [
            "数据孤岛：整合现有数据源，打破信息壁垒。",
            "流程效率：自动化重复性任务，提升运营效率。",
            "客户体验：通过数据分析提供个性化服务。"
          ],
          "image_keyword": "business people analyzing data chart",
          "layout_options": {{
            "image_position": "left",
            "split_ratio": "60/40"
          }}
        }},
        {{
          "page_type": "process_flow",
          "title": "我们的实施路径",
          "content": [
            {{ "step": "第一阶段：评估与规划", "description": "全面分析业务现状，制定转型路线图。" }},
            {{ "step": "第二阶段：技术实施", "description": "部署云平台与数据分析工具。" }},
            {{ "step": "第三阶段：优化与迭代", "description": "持续监控效果，根据反馈进行优化。" }}
          ],
          "image_keyword": "process workflow diagram"
        }}
      ]
    }}
    """
    try:
        logging.info(f"Generating flexible {num_pages}-page plan for theme: '{theme}'...")
        chat_completion = ai_client.chat.completions.create(
            model=config.MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.8, # Slightly higher temp for more creative layouts
            response_format={"type": "json_object"}
        )
        content_str = chat_completion.choices[0].message.content
        logging.info("Successfully received AI design plan.")
        return json.loads(content_str)
    except Exception as e:
        logging.error(f"Error calling AI API or parsing JSON: {e}", exc_info=True)
        return None