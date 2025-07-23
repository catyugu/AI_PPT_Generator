import json
import logging
import re

from openai import OpenAI

from config import ONEAPI_KEY, ONEAPI_BASE_URL, MODEL_NAME

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 初始化OneAPI客户端
try:
    if not ONEAPI_KEY or ONEAPI_KEY == "YOUR_ONEAPI_KEY_HERE":
        raise ValueError("ONEAPI_KEY未在config.py或.env文件中正确设置。")
    client = OpenAI(api_key=ONEAPI_KEY, base_url=ONEAPI_BASE_URL)
    logging.info(f"OpenAI客户端已为OneAPI初始化，目标地址: {ONEAPI_BASE_URL}")
except ValueError as e:
    logging.error(e)
    client = None


def _extract_json_from_response(text: str) -> str | None:
    """从可能包含markdown和注释的字符串中提取并清理JSON对象。"""
    try:
        # 移除Markdown代码块标记和可选的'json'语言标识符
        text = re.sub(r'```json\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'```', '', text)

        # 找到第一个'{'和最后一个'}'，以确定JSON对象的边界
        start_index = text.find('{')
        end_index = text.rfind('}')

        if start_index != -1 and end_index != -1 and end_index > start_index:
            json_block = text[start_index:end_index + 1]
            # 移除行内和行末的JS风格注释
            json_block = re.sub(r'//.*', '', json_block)
            # 移除多行注释 (/* ... */)
            json_block = re.sub(r'/\*.*?\*/', '', json_block, flags=re.DOTALL)
            return json_block
        logging.warning("在AI响应中未找到有效的JSON对象边界。")
        return None
    except Exception as e:
        logging.error(f"提取和清理JSON时出错: {e}")
        return None

def generate_presentation_plan(theme: str, num_pages: int) -> dict | None:
    """使用OneAPI为演示文稿生成详细的JSON计划。"""
    if not client:
        logging.error("OneAPI client not initialized.")
        return None

    # 直接使用从 config.py 导入的模型名称
    logging.info(f"Requesting plan from model '{MODEL_NAME}' via OneAPI...")

    # 使用我们最终优化的、带有“多样性强制”规则的提示词
    prompt = f"""
    你是一位世界顶级的演示文稿（PPT）设计大师和信息架构专家。你精通平面设计、版式理论、色彩心理学和视觉传达。你的任务是根据用户提供的主题，设计一份兼具专业性、设计感和视觉冲击力的演示文稿方案。
    
    **你的输出必须是一个单一、完整、严格符合以下所有规则的原始JSON对象，禁止包含任何JSON之外的解释性文字、注释或Markdown代码块标记（如 ```json）。**
    
    ---
    
    ### **第一部分：全局设计系统 (Global Design System)**
    
    这是整个演示文稿的基石，定义了统一的视觉规范。
    
    1.  **`design_concept`**: (字符串) **必须用中文**为本次设计提炼一个高度概括、富有创意的核心设计理念。例如：“深海数据之境”、“都市脉搏与光影”、“墨韵书香”、“赛博朋克霓虹”等。
    2.  **`font_pairing`**: (对象) 定义全局字体搭配。
        * `heading`: (字符串) 标题字体。例如: "华文细黑"。
        * `body`: (字符串) 正文字体。例如: "宋体"。
    3.  **`color_palette`**: (对象) 定义一个专业、和谐的色板。
        * `primary`: (字符串, Hex) 主色，用于关键元素、标题。
        * `secondary`: (字符串, Hex) 辅色，用于次要信息、图表。
        * `background`: (字符串, Hex) 背景色。
        * `text`: (字符串, Hex) 主要文本颜色。
        * `accent`: (字符串, Hex) 点缀色/强调色，用于按钮、图表高亮、特殊标记。
    4.  **`master_slide`**: (对象) 定义应用于所有页面的“母版”元素。
        * `background`: (对象) 定义背景。可以是纯色 `{{ "color": "#RRGGBB" }}`，也可以是图片 `{{ "image_keyword": "英文图片关键词" }}`。
        * `footer`: (对象, 可选) 页脚。包含 `text` (如“公司名称 | 内部资料”) 和 `style` (定义 `x`, `y`, `width`, `height`, `font_size`, `color` 等)。
        * `page_number`: (对象, 可选) 页码。包含 `style` (定义 `x`, `y`, `width`, `height`, `font_size`, `color` 等)。
    
    ---
    
    ### **第二部分：页面详细规划 (Page Details)**
    
    `pages` 是一个数组，其中每个对象代表一页幻灯片。
    
    * **`layout_type`**: (字符串) 对本页布局风格的描述。例如: `title_slide`, `image_left_content_right`, `full_screen_image_with_quote`, `three_column_comparison`, `data_chart_summary`, `process_flow`, `team_introduction`。
    * **`elements`**: (数组) 页面上所有视觉元素的集合。
    
    #### **元素 (Element) 定义**
    
    **所有元素都必须包含** `type`, `x`, `y`, `width`, `height` 这五个基本属性 (单位: px, 基于 1280x720 的画布)。
    
    1.  **`text_box`**
        * `type`: "text_box"
        * `content`: (字符串) 文本内容，支持用 `\\n` 换行。
        * `style`: (对象)
            * `font`: (对象)
                * `type`: (字符串) "heading" 或 "body"，用于调用全局字体。
                * `size`: (数字) 字号 (pt)。
                * `color`: (字符串, Hex, 可选) 局部覆盖全局文本颜色。
                * `bold`: (布尔值, 可选) 是否加粗。
                * `italic`: (布尔值, 可选) 是否斜体。
            * `alignment`: (字符串, 可选) 对齐方式: "LEFT", "CENTER", "RIGHT"。
    
    2.  **`image`**
        * `type`: "image"
        * `image_keyword`: (字符串) **必须是英文**的图片搜索关键词，越具体越好。
        * `style`: (对象, 可选)
            * `opacity`: (数字, 0.0-1.0) 透明度。
            * `border`: (对象) 边框。包含 `color` (Hex) 和 `width` (px)。
            * `crop`: (对象, 可选) 裁剪。`{{ "shape": "circle" }}` 可将图片裁剪为圆形。
    
    3.  **`shape`**
        * `type`: "shape"
        * `shape_type`: (字符串) 形状类型。可选值: `rectangle`, `oval`, `triangle`, `star`, `rounded_rectangle`。
        * `style`: (对象)
            * `fill_color`: (字符串, Hex, 可选) 填充色。
            * `gradient`: (对象, 可选, 与`fill_color`互斥) 渐变填充。
                * `type`: "linear" 或 "radial"。
                * `angle`: (数字, 仅linear) 渐变角度。
                * `colors`: (数组) `["#RRGGBB", "#RRGGBB"]`。
            * `border`: (对象, 可选) 边框。包含 `color` (Hex) 和 `width` (px)。
    
    4.  **`chart`**
        * `type`: "chart"
        * `chart_type`: (字符串) 图表类型。可选值: `bar` (柱状图), `pie` (饼图), `line` (折线图)。
        * `data`: (对象)
            * `categories`: (字符串数组) 类别轴标签。
            * `series`: (对象数组) 每个对象是一组数据系列。
                * `name`: (字符串) 系列名称。
                * `values`: (数字数组) 数据值。
    
    5.  **`table`**
        * `type`: "table"
        * `headers`: (字符串数组) 表头。
        * `rows`: (二维字符串数组) 表格数据。
        * `style`: (对象, 可选)
            * `header_color`: (字符串, Hex) 表头背景色。
            * `row_colors`: (字符串数组, Hex) `["#FFFFFF", "#F0F0F0"]` 用于创建斑马条纹。
            * `line_color`: (字符串, Hex) 表格线的颜色。
    
    ---
    
    ### **第三部分：多样性与一致性核心准则 (Core Principles for Variety & Consistency)**
    
    为了确保整个演示文稿都保持高水准，你必须遵守以下核心准则：
    
    1.  **布局多样性 (Layout Variety)**: 在生成 `pages` 数组时，**必须**有意识地混合使用多种不同的 `layout_type`。**严禁**连续超过两页使用完全相同的简单布局（如纯文本页面）。请交错使用图文、图表、引用、多栏等复杂布局。
    2.  **视觉元素丰富度 (Visual Richness)**: 除了标题页和结论页，**每一张内容页都应至少包含一个视觉元素** (`image`, `shape`, `chart`, `table`)，以避免页面单调。鼓励使用形状和图片进行创意组合，以增强视觉吸引力。
    3.  **设计系统贯穿始终 (Consistent Design System)**: 你在第一部分定义的 `color_palette` 和 `font_pairing` **必须**被应用到**所有页面**的**所有元素**上。所有颜色和字体都应源自这个全局设计系统，以保证视觉统一性。
    4.  **字体选择的泛用性 (Font Generality)**: 确保你的字体是在大部分电脑上可用的，以免因字体不支持等原因导致无法正常显示。
    5.  **切勿在`content`字段的文本中使用任何Markdown语法**（例如 `**文字**` 或 `*`）。
    6.  **所有的文本样式（如加粗）都必须通过`style`对象中的对应属性（如 `"bold": true`）来定义。**
    7.  **你的PPT页数应该严格与用户要求的页数一致**
    ---
    
    ### **第四部分：输出样例**
    
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
    
    **最后指令：**
    现在，请**回顾并严格遵守以上所有部分的规则**，为主题 **“{theme}”** 生成一个包含 **{num_pages}** 页的完整PPT设计方案JSON。
    """

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system",
                 "content": "You are a world-class presentation designer. Your output must be a single, raw JSON object. You must strictly follow all instructions."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=128000,
            temperature=0.55  # 稍微提高一点温度以增加创意性，但仍保持结构稳定
        )

        response_content = response.choices[0].message.content
        if response_content:
            logging.info("已成功从AI接收到演示文稿方案。")
            json_string = _extract_json_from_response(response_content)
            if json_string:
                # 在加载前移除可能导致错误的尾随逗号
                cleaned_json_string = re.sub(r',\s*([}\]])', r'\1', json_string)
                return json.loads(cleaned_json_string)
            else:
                logging.error("从AI响应中提取JSON失败。")
                return None

        logging.error("AI响应内容为空。")
        return None

    except json.JSONDecodeError as e:
        logging.error(f"JSON解码失败: {e}。原始响应片段: '{json_string[:500]}...'")
        return None
    except Exception as e:
        logging.error(f"与OneAPI通信时发生严重错误: {e}", exc_info=True)
        return None
