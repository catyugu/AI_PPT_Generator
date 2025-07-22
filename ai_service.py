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
    Generates a PPT plan with flexible, AI-driven element-level layout options.
    The AI defines the exact position, size, type, and style of each element.
    Coordinates (x, y), width, and height are in pixels, relative to a standard
    PowerPoint slide size (914x685 pixels, roughly 10x7.5 inches at 96 DPI).
    """
    if not ai_client:
        logging.error("AI client not initialized. Cannot generate plan.")
        return None

    # New prompt for flexible element-based layout generation
    prompt = f"""
    You are an expert presentation designer that produces structured JSON output. Your task is to create a detailed plan for a presentation on the theme "{theme}".

    **Your response MUST be a single, raw JSON object and nothing else.**

    ---
    **## JSON STRUCTURE REQUIREMENTS**

    1.  **Root Object**: Must contain keys: `design_concept`, `font_pairing`, `color_palette`, `pages`.
    2.  **Language**: All user-facing text (`design_concept`, `page_title`, `content`) MUST be in **Chinese**.
    3.  **Keywords**: All machine-facing keywords (`image_keyword`) MUST be in **English**.
    4.  **Page Count**: The `pages` list must contain exactly **{num_pages}** page objects.
    5.  **`color_palette`**: Must be a JSON **object** with keys: "primary", "secondary", "background", "text", "accent".
    6.  **`pages` array**: Each object in `pages` must have a `page_title` and an `elements` array.
    7.  **`elements` array**: Each object in `elements` represents a single visual component on the slide.
        * **Common properties for all elements**:
            * `type`: (string) **REQUIRED**. Must be one of: `text_box`, `image`, `shape`, `chart`.
            * `x`: (integer) **REQUIRED**. X-coordinate in pixels (from left edge).
            * `y`: (integer) **REQUIRED**. Y-coordinate in pixels (from top edge).
            * `width`: (integer) **REQUIRED**. Width in pixels.
            * `height`: (integer) **REQUIRED**. Height in pixels.
            * `style`: (object) **OPTIONAL**. Contains styling properties.
        * **`text_box` specific properties**:
            * `content`: (string) The text content.
            * `style`: (object) Can include `font_size`, `font_name`, `color` (hex), `alignment` (left, center, right), `bold`, `italic`.
        * **`image` specific properties**:
            * `image_keyword`: (string) Keyword for image generation/selection.
            * `style`: (object) Can include `opacity` (0.0-1.0).
        * **`shape` specific properties**:
            * `shape_type`: (string) **REQUIRED**. Must be one of: `rectangle`, `oval`, `triangle`, `line`.
            * `style`: (object) Can include `fill_color` (hex), `line_color` (hex), `line_width` (points), `opacity` (0.0-1.0).
        * **`chart` specific properties**:
            * `chart_type`: (string) **REQUIRED**. Must be one of: `bar_chart`, `column_chart`, `line_chart`, `pie_chart`.
            * `data`: (object) **REQUIRED**. Structure depends on `chart_type`.
                * For `bar_chart`/`column_chart`/`line_chart`: `{{\"categories\": [\"Cat1\", \"Cat2\"], \"series\": [{{\"name\": \"Series A\", \"values\": [10, 20]}}, {{\"name\": \"Series B\", \"values\": [15, 25]}}]}}`
                * For `pie_chart`: `{{\"labels\": [\"Label1\", \"Label2\"], \"values\": [30, 70]}}`
            * `title`: (string) Chart title.
            * `style`: (object) Can include `data_labels` (boolean), `legend_position` (top, bottom, left, right).

    ---
    **## EXAMPLE JSON OUTPUT (for a 2-page presentation)**

    ```json
    {{
      "design_concept": "未来科技，创新驱动",
      "font_pairing": {{ "heading": "Arial Bold", "body": "Arial" }},
      "color_palette": {{ "primary": "#1A2B3C", "secondary": "#4A6C8E", "background": "#F0F5F9", "text": "#333333", "accent": "#FFD700" }},
      "pages": [
        {{
          "page_title": "封面页：创新未来",
          "elements": [
            {{
              "type": "image",
              "image_keyword": "futuristic city skyline",
              "x": 0, "y": 0, "width": 914, "height": 685,
              "style": {{ "opacity": 0.3 }}
            }},
            {{
              "type": "text_box",
              "content": "未来科技：创新与机遇",
              "x": 100, "y": 200, "width": 714, "height": 100,
              "style": {{ "font_size": 48, "font_name": "Arial Bold", "color": "#1A2B3C", "alignment": "center" }}
            }},
            {{
              "type": "text_box",
              "content": "探索人工智能、大数据和物联网的无限可能",
              "x": 100, "y": 320, "width": 714, "height": 60,
              "style": {{ "font_size": 24, "font_name": "Arial", "color": "#4A6C8E", "alignment": "center" }}
            }},
            {{
              "type": "shape",
              "shape_type": "rectangle",
              "x": 50, "y": 180, "width": 814, "height": 220,
              "style": {{ "fill_color": "#FFFFFF", "opacity": 0.6, "line_color": null }}
            }}
          ]
        }},
        {{
          "page_title": "人工智能核心概念",
          "elements": [
            {{
              "type": "text_box",
              "content": "人工智能：驱动未来",
              "x": 50, "y": 50, "width": 800, "height": 50,
              "style": {{ "font_size": 36, "font_name": "Arial Bold", "color": "#1A2B3C", "alignment": "left" }}
            }},
            {{
              "type": "text_box",
              "content": "人工智能（AI）是计算机科学的一个分支，旨在创建能够执行通常需要人类智能的任务的机器。它涵盖了机器学习、深度学习、自然语言处理等多个子领域。",
              "x": 50, "y": 120, "width": 400, "height": 150,
              "style": {{ "font_size": 18, "font_name": "Arial", "color": "#333333", "alignment": "left" }}
            }},
            {{
              "type": "image",
              "image_keyword": "robot hand touching human hand",
              "x": 480, "y": 100, "width": 400, "height": 300
            }},
            {{
              "type": "shape",
              "shape_type": "oval",
              "x": 600, "y": 450, "width": 200, "height": 100,
              "style": {{ "fill_color": "#FFD700", "opacity": 0.8, "line_color": "#1A2B3C", "line_width": 2 }}
            }},
            {{
              "type": "text_box",
              "content": "AI应用广泛",
              "x": 600, "y": 475, "width": 200, "height": 50,
              "style": {{ "font_size": 16, "font_name": "Arial", "color": "#1A2B3C", "alignment": "center", "bold": true }}
            }}
          ]
        }}
      ]
    }}
    ```
    """
    try:
        logging.info(f"Generating flexible {num_pages}-page plan for theme: '{theme}'...")
        chat_completion = ai_client.chat.completions.create(
            model=config.MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.8,  # Slightly higher temp for more creative layouts
            response_format={"type": "json_object"}
        )
        content_str = chat_completion.choices[0].message.content
        logging.info("Successfully received AI design plan.")
        return json.loads(content_str)
    except Exception as e:
        logging.error(f"Error calling AI API or parsing JSON: {e}", exc_info=True)
        return None
