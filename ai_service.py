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
    6.  **`page_type` Value (CRITICAL)**: The value for `page_type` **MUST** be one of the following exact strings.
        - `cover`
        - `section_header`
        - `title_content`
        - `image_text_split`
        - `full_bleed_image_quote`
        - `three_points_icons`
        - `icon_grid`
        - `timeline`
        - `process_flow`
        - `bar_chart`
        - `team_intro`
        - `thank_you`
        **Do NOT invent any other `page_type` values.** If a specific layout isn't available, use `title_content`.
    7.  **`layout_options`**: An **optional** object inside a page object to control style.
    
    ---
    **## EXAMPLE JSON OUTPUT**
    
    ```json
    {{
      "design_concept": "简约科技，数据驱动未来",
      "font_pairing": {{ "heading": "Dengxian", "body": "Microsoft YaHei" }},
      "color_palette": {{ "primary": "#0A3D62", "secondary": "#3C6382", "background": "#F5F5F5", "text": "#333333", "accent": "#FDCB6E" }},
      "pages": [
        {{
          "page_type": "cover",
          "title": "企业数字化转型解决方案",
          "content": ["引领您的业务进入智能时代"],
          "image_keyword": "abstract digital transformation",
          "layout_options": {{ "text_alignment": "center" }}
        }},
        {{
          "page_type": "image_text_split",
          "title": "核心挑战与机遇",
          "content": ["数据孤岛：整合现有数据源，打破信息壁垒。"],
          "image_keyword": "business people analyzing data chart",
          "layout_options": {{ "image_position": "left", "split_ratio": "60/40" }}
        }},
        {{
          "page_type": "bar_chart",
          "title": "季度销售额对比",
          "content": {{
            "categories": ["第一季度", "第二季度", "第三季度", "第四季度"],
            "series_name": "销售额（万元）",
            "values": [250, 310, 280, 360]
          }},
          "image_keyword": "financial growth chart abstract"
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