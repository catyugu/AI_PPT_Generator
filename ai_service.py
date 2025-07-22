# ai_service.py
import os
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
    [ENHANCED V2] Generates a PPT plan with specified page number and language.
    """
    if not ai_client:
        logging.error("AI client not initialized. Cannot generate plan.")
        return None

    # V2 Prompt: Asks for Chinese text, specified page count, and universal fonts.
    prompt = f"""
    You are an expert presentation designer. Create a plan for a presentation on the theme "{theme}".

    **CRITICAL INSTRUCTIONS**:
    1.  **Language**: All text content (`title`, `content` descriptions, etc.) MUST be in **Chinese**.
    2.  **Image Keywords**: The `image_keyword` value must remain in **English** for better API search results.
    3.  **Page Count**: The presentation must have exactly **{num_pages}** pages.
    4.  **Fonts**: For `font_pairing`, choose universally available fonts. For Chinese, suggest fonts like "Dengxian", "Microsoft YaHei". For English/general use, suggest "Calibri", "Arial", "Times New Roman".

    Provide the following in a strict JSON format. Do not include any text outside the JSON object.
    1.  `design_concept`: A short, inspiring description of the design philosophy (in Chinese).
    2.  `font_pairing`: {{ "heading": "FontName", "body": "FontName" }}.
    3.  `color_palette`: A palette of 5 HEX codes: `primary`, `secondary`, `background`, `text`, `accent`.
    4.  `pages`: A list of exactly **{num_pages}** pages. For each page, provide:
        a.  `page_type`: Choose from: `cover`, `section_header`, `title_content`, `image_text_split`, `full_bleed_image_quote`, `three_points_icons`, `icon_grid`, `timeline`, `process_flow`, `bar_chart`, `team_intro`, `thank_you`.
        b.  `title`: The slide title (in Chinese).
        c.  `content`: Varies by `page_type` (all text in Chinese):
            - For text-based slides: An array of bullet point strings.
            - For `full_bleed_image_quote`: A single quote string.
            - For `icon_grid`, `three_points_icons`: Array of objects, `{{ "icon_keyword": "English keyword", "text": "Chinese text" }}`.
            - For `team_intro`: Array of objects, `{{ "name": "姓名", "title": "职位", "image_keyword": "English keyword for photo" }}`.
            - For `timeline`, `process_flow`: Array of objects, `{{ "step": "第一步", "description": "详细描述" }}`.
            - For `bar_chart`: Object `{{ "categories": ["类别一", "类别二"], "series_name": "系列名称", "values": [10, 20] }}`.
        d.  `image_keyword`: A concise **English** keyword for an image search.
        e.  `design_notes`: (Optional) Specific instructions (in Chinese).

    Return ONLY the raw JSON object.
    """
    try:
        logging.info(f"Generating {num_pages}-page plan for theme: '{theme}' in Chinese...")
        chat_completion = ai_client.chat.completions.create(
            model=config.MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            response_format={"type": "json_object"}
        )
        content_str = chat_completion.choices[0].message.content
        logging.info("Successfully received AI design plan.")
        return json.loads(content_str)
    except Exception as e:
        logging.error(f"Error calling AI API or parsing JSON: {e}", exc_info=True)
        return None