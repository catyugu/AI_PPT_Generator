# ai_service.py
import json
import logging
from openai import OpenAI
import config

# --- Initialize Client ---
try:
    client = OpenAI(
        api_key=config.ONEAPI_KEY,
        base_url=config.ONEAPI_BASE_URL,
    )
    logging.info("OpenAI client initialized.")
except Exception as e:
    logging.error(f"Failed to initialize OpenAI client: {e}")
    client = None

def _extract_json_from_response(text: str) -> str | None:
    """
    从可能包含额外文本的字符串中提取出JSON对象。
    """
    try:
        start_index = text.find('{')
        end_index = text.rfind('}')
        if start_index != -1 and end_index != -1 and end_index > start_index:
            return text[start_index:end_index+1]
        logging.warning("Could not find a valid JSON object in the AI response.")
        return None
    except Exception as e:
        logging.error(f"Error while extracting JSON: {e}")
        return None


def generate_presentation_plan(theme: str, num_pages: int) -> dict | None:
    """
    使用OpenAI API生成演示文稿的详细JSON计划。
    """
    if not client:
        logging.error("OpenAI client not initialized. Cannot generate presentation plan.")
        return None

    prompt = f"""
    你是一位世界顶级的演示文稿（PPT）设计大师和信息架构专家。你精通平面设计、版式理论、色彩心理学和视觉传达。你的任务是根据用户提供的主题，设计一份兼具专业性、设计感和视觉冲击力的演示文稿方案。
    
    **你的输出必须是一个单一、完整、严格符合以下所有规则的原始JSON对象，禁止包含任何JSON之外的解释性文字、注释或Markdown代码块标记（如 \`\`\`json）。**
    
    ### **第一部分：全局设计系统 (Global Design System)**
    
    这是整个演示文稿的基石，定义了统一的视觉规范。
    
    1. **design_concept**: (字符串) **必须用中文**为本次设计提炼一个高度概括、富有创意的核心设计理念。例如：“深海数据之境”、“都市脉搏与光影”、“墨韵书香”、“赛博朋克霓虹”等。  
    2. **font_pairing**: (对象) 定义全局字体搭配。  
       * heading: (字符串) 标题字体。例如: "Microsoft YaHei Heavy"。  
       * body: (字符串) 正文字体。例如: "Microsoft YaHei"。  
    3. **color_palette**: (对象) 定义一个专业、和谐的色板。  
       * primary: (字符串, Hex) 主色，用于关键元素、标题。  
       * secondary: (字符串, Hex) 辅色，用于次要信息、图表。  
       * background: (字符串, Hex) 背景色。  
       * text: (字符串, Hex) 主要文本颜色。  
       * accent: (字符串, Hex) 点缀色/强调色，用于按钮、图表高亮、特殊标记。  
    4. **master_slide**: (对象) 定义应用于所有页面的“母版”元素。  
       * background: (对象) 定义背景。可以是纯色 {{ "color": "#RRGGBB" }}，也可以是图片 {{ "image_keyword": "英文图片关键词" }}。  
       * footer: (对象, 可选) 页脚。包含 text (如“公司名称 | 内部资料”) 和 style (定义 x, y, width, height, font_size, color 等)。  
       * page_number: (对象, 可选) 页码。包含 style (定义 x, y, width, height, font_size, color 等)。
    
    ### **第二部分：页面详细规划 (Page Details)**
    
    pages 是一个数组，其中每个对象代表一页幻灯片。
    
    * **layout_type**: (字符串) 对本页布局风格的描述。例如: title_slide, image_left_content_right, full_screen_image_with_quote, three_column_comparison, data_chart_summary, process_flow, team_introduction。  
    * **elements**: (数组) 页面上所有视觉元素的集合。
    
    #### **元素 (Element) 定义**
    
    **所有元素都必须包含** type, x, y, width, height 这五个基本属性 (单位: px, 基于 1280x720 的画布)。
    
    1. **text_box**  
       * type: "text_box"  
       * content: (字符串) 文本内容，支持用 \\n 换行。  
       * style: (对象)  
         * font: (对象)  
           * type: (字符串) "heading" 或 "body"，用于调用全局字体。  
           * size: (数字) 字号 (pt)。  
           * color: (字符串, Hex, 可选) 局部覆盖全局文本颜色。  
           * bold: (布尔值, 可选) 是否加粗。  
           * italic: (布尔值, 可选) 是否斜体。  
         * alignment: (字符串, 可选) 对齐方式: "LEFT", "CENTER", "RIGHT"。  
    2. **image**  
       * type: "image"  
       * image_keyword: (字符串) **必须是英文**的图片搜索关键词，越具体越好。  
       * style: (对象, 可选)  
         * opacity: (数字, 0.0-1.0) 透明度。  
         * border: (对象) 边框。包含 color (Hex) 和 width (px)。  
         * crop: (对象, 可选) 裁剪。{{ "shape": "circle" }} 可将图片裁剪为圆形。  
    3. **shape**  
       * type: "shape"  
       * shape_type: (字符串) 形状类型。可选值: rectangle, oval, triangle, star, rounded_rectangle。  
       * style: (对象)  
         * fill_color: (字符串, Hex, 可选) 填充色。  
         * gradient: (对象, 可选, 与fill_color互斥) 渐变填充。  
           * type: "linear" 或 "radial"。  
           * angle: (数字, 仅linear) 渐变角度。  
           * colors: (数组) ["#RRGGBB", "#RRGGBB"]。  
         * border: (对象, 可选) 边框。包含 color (Hex) 和 width (px)。  
    4. **chart**  
       * type: "chart"  
       * chart_type: (字符串) 图表类型。可选值: bar (柱状图), pie (饼图), line (折线图)。  
       * data: (对象)  
         * categories: (字符串数组) 类别轴标签。  
         * series: (对象数组) 每个对象是一组数据系列。  
           * name: (字符串) 系列名称。  
           * values: (数字数组) 数据值。  
    5. **table**  
       * type: "table"  
       * headers: (字符串数组) 表头。  
       * rows: (二维字符串数组) 表格数据。  
       * style: (对象, 可选)  
         * header_color: (字符串, Hex) 表头背景色。  
         * row_colors: (字符串数组, Hex) ["#FFFFFF", "#F0F0F0"] 用于创建斑马条纹。  
         * line_color: (字符串, Hex) 表格线的颜色。
    
    ### **第三部分：输出样例**
    
    #### **样例一：科技商务风PPT**
    
    {{  
      "design_concept": "深蓝智域：数据与洞察",  
      "font_pairing": {{  
        "heading": "Microsoft YaHei Heavy",  
        "body": "Microsoft YaHei"  
      }},  
      "color_palette": {{  
        "primary": "#0D47A1",  
        "secondary": "#1976D2",  
        "background": "#FFFFFF",  
        "text": "#212121",  
        "accent": "#4CAF50"  
      }},  
      "master_slide": {{  
        "background": {{  
          "color": "#FFFFFF"  
        }},  
        "footer": {{  
          "text": "ABC科技有限公司 | 内部资料",  
          "style": {{  
            "x": 60, "y": 680, "width": 500, "height": 20, "font_size": 10, "color": "#757575"  
          }}  
        }}  
      }},  
      "pages": [  
        {{  
          "layout_type": "title_slide",  
          "elements": [  
            {{ "type": "shape", "shape_type": "rectangle", "x": 0, "y": 200, "width": 1280, "height": 320, "style": {{ "fill_color": "#0D47A1" }} }},  
            {{ "type": "text_box", "x": 80, "y": 280, "width": 1120, "height": 100, "content": "2025年度市场增长战略报告", "style": {{ "font": {{ "type": "heading", "size": 48, "color": "#FFFFFF", "bold": true }}, "alignment": "CENTER" }} }},  
            {{ "type": "text_box", "x": 80, "y": 390, "width": 1120, "height": 40, "content": "数据驱动，洞见未来", "style": {{ "font": {{ "type": "body", "size": 22, "color": "#BDBDBD" }}, "alignment": "CENTER" }} }}  
          ]  
        }},  
        {{  
          "layout_type": "data_chart_summary",  
          "elements": [  
            {{ "type": "text_box", "x": 60, "y": 60, "width": 1160, "height": 50, "content": "季度销售额对比分析", "style": {{ "font": {{ "type": "heading", "size": 32, "bold": true }} }} }},  
            {{ "type": "chart", "x": 60, "y": 140, "width": 1160, "height": 500, "chart_type": "bar", "data": {{ "categories": ["第一季度", "第二季度", "第三季度", "第四季度"], "series": [ {{ "name": "2024年", "values": [450, 520, 580, 650] }}, {{ "name": "2025年预测", "values": [500, 590, 670, 750] }} ] }} }}  
          ]  
        }}  
      ]  
    }}
    
    #### **样例二：人文艺术风PPT**
    
    {{  
      "design_concept": "墨韵留白：东方美学的现代诠释",  
      "font_pairing": {{  
        "heading": "SimSun",  
        "body": "KaiTi"  
      }},  
      "color_palette": {{  
        "primary": "#212121",  
        "secondary": "#757575",  
        "background": "#F5F5F5",  
        "text": "#424242",  
        "accent": "#C62828"  
      }},  
      "master_slide": {{  
        "background": {{  
          "image_keyword": "chinese ink wash painting minimalist background"  
        }}  
      }},  
      "pages": [  
        {{  
          "layout_type": "full_screen_image_with_quote",  
          "elements": [  
            {{ "type": "image", "x": 0, "y": 0, "width": 1280, "height": 720, "image_keyword": "lonely boat on calm water black and white", "style": {{ "opacity": 0.8 }} }},  
            {{ "type": "shape", "shape_type": "rectangle", "x": 0, "y": 0, "width": 1280, "height": 720, "style": {{ "gradient": {{ "type": "linear", "angle": 90, "colors": ["#00000000", "#000000CC"] }} }} }},  
            {{ "type": "text_box", "x": 100, "y": 500, "width": 600, "height": 150, "content": "“天地有大美而不言”", "style": {{ "font": {{ "type": "heading", "size": 44, "color": "#FFFFFF", "italic": true }}, "alignment": "LEFT" }} }}  
          ]  
        }},  
        {{  
          "layout_type": "image_left_content_right",  
          "elements": [  
            {{ "type": "image", "x": 80, "y": 120, "width": 450, "height": 450, "image_keyword": "zen stone garden detail", "style": {{ "crop": {{ "shape": "circle" }} }} }},  
            {{ "type": "shape", "shape_type": "oval", "x": 70, "y": 110, "width": 470, "height": 470, "style": {{ "border": {{ "color": "#C62828", "width": 3 }} }} }},  
            {{ "type": "text_box", "x": 600, "y": 180, "width": 600, "height": 80, "content": "静观", "style": {{ "font": {{ "type": "heading", "size": 36, "bold": true }} }} }},  
            {{ "type": "text_box", "x": 600, "y": 280, "width": 600, "height": 300, "content": "在纷繁的世界里，\\n寻一方内心的宁静。\\n一石一木，\\n皆是禅意。", "style": {{ "font": {{ "type": "body", "size": 22, "color": "#424242" }}, "alignment": "LEFT" }} }}  
          ]  
        }}  
      ]  
    }}
    
    ### **最后指令：**  
    现在，请**严格按照以上所有规则**，为主题 “{theme}” 生成一个包含 {num_pages} 页的完整PPT设计方案JSON。
    """

    try:
        logging.info("Generating presentation plan from AI...")
        response = client.chat.completions.create(
            model=config.MODEL_NAME,
            messages=[
                {"role": "system", "content": "You are a world-class presentation designer. Your output must be a single, raw JSON object without any extra text or markdown. You must strictly follow all instructions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
        )

        response_content = response.choices[0].message.content
        if response_content:
            logging.info("Successfully received presentation plan from AI.")
            json_string = _extract_json_from_response(response_content)
            if json_string:
                return json.loads(json_string)
            else:
                logging.error("Could not extract a valid JSON object from the AI's response.")
                logging.debug(f"Raw response was: {response_content}")
                return None
        else:
            logging.error("Received empty response from AI.")
            return None

    except json.JSONDecodeError as e:
        logging.error(f"Failed to parse JSON response from AI: {e}")
        logging.error(f"Raw response was: {response_content}")
        return None
    except Exception as e:
        logging.error(f"An error occurred while communicating with OpenAI: {e}", exc_info=True)
        return None
