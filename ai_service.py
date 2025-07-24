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


def generate_presentation_plan(theme: str, num_pages: int, aspect_ratio: str = "16:9") -> dict | None:
    """使用OneAPI为演示文稿生成详细的JSON计划。"""
    if not client:
        logging.error("OneAPI client not initialized.")
        return None

    # 直接使用从 config.py 导入的模型名称
    logging.info(f"Requesting plan from model '{MODEL_NAME}' via OneAPI...")
    if aspect_ratio == "4:3":
        canvas_width, canvas_height = 1024, 768
    else:  # 默认为 16:9
        canvas_width, canvas_height = 1280, 720

    # [核心修改] 更新了Prompt，赋予AI更灵活的字体控制权
    prompt = f"""
    你是一位你是一位深谙**年轻女性审美**的顶级演示文稿（PPT）设计大师和信息架构专家。你精通平面设计、版式理论、色彩心理学和视觉传达。你的任务是根据用户提供的主题，设计一份兼具专业性、设计感和视觉冲击力的演示文稿方案。

    **你的输出必须是一个单一、完整、严格符合以下所有规则的原始JSON对象，禁止包含任何JSON之外的解释性文字、注释或Markdown代码块标记（如 ```json）。**

    ---

    ### **第一部分：全局设计系统 (Global Design System)**

    这是整个演示文稿的基石，定义了统一的视觉规范。

    1.  **`design_concept`**: (字符串) **必须用中文**为本次设计提炼一个高度概括、富有创意的核心设计理念。例如：“深海数据之境”、“都市脉搏与光影”、“墨韵书香”、“赛博朋克霓虹”等。
    2.  **`font_pairing`**: (对象) 定义全局**默认**字体搭配。当单个文本元素未指定特定字体时，将使用这里的设置。
        * `heading`: (字符串) 默认标题字体。例如: "黑体 (SimHei)"。
        * `body`: (字符串) 默认正文字体。例如: "宋体 (SimSun)"。
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

    **所有元素都必须包含** `type`, `x`, `y`, `width`, `height` 这五个基本属性。
    **重要：所有坐标和尺寸都必须基于一个 {canvas_width}x{canvas_height} 像素的画布进行设计。**

    1.  **`text_box`**
        * `type`: "text_box"
        * `content`: (字符串或字符串数组) **[新功能]** 如果是普通文本，则为字符串，支持用 `\\n` 换行。如果要创建项目符号列表，则**必须**使用字符串数组，数组中每个字符串代表一个列表项。
        * `style`: (对象)
            * `font`: (对象)
                * `name`: (字符串, 可选) **直接指定字体名称** (如 "黑体", "微软雅黑")。**如果提供此项，将优先使用，并忽略下面的 `type` 字段。**
                * `type`: (字符串, 可选) "heading" 或 "body"。如果未指定 `name`，则会根据此类型调用全局 `font_pairing` 中定义的默认字体。
                * `size`: (数字) 字号 (pt)。
                * `color`: (字符串, Hex, 可选) 局部覆盖全局文本颜色。
                * `bold`: (布尔值, 可选) 是否加粗。
                * `italic`: (布尔值, 可选) 是否斜体。
            * `alignment`: (字符串, 可选) 对齐方式: "LEFT", "CENTER", "RIGHT"。

    2.  **`image`**
        * `type`: "image"
        * `image_keyword`: (字符串) **必须是英文**的图片搜索关键词，越具体越好。
        * `style`: (对象, 可选)
            * `opacity`: (数字, 0.0-1.0) 图片整体的不透明度。
            * `border`: (对象) 边框。包含 `color` (Hex) 和 `width` (px)。
            * `crop`: (字符串, 可选) **[新功能]** 裁剪形状。目前唯一支持的值是 **`"circle"`**，用于将图片裁剪为圆形。

    3.  **`shape`**
        * `type`: "shape"
        * `shape_type`: (字符串) 形状类型。可选值: `rectangle`, `oval`, `triangle`, `star`, `rounded_rectangle`。
        * `style`: (对象)
            * `fill_color`: (字符串, Hex, 可选) 填充色。
            * `opacity`: (数字, 0.0-1.0, 可选) 填充色的不透明度。`0.0`为完全不透明, `1.0`为完全透明。
            * `gradient`: (对象, 可选, 与`fill_color`互斥) 渐变填充。
            * `border`: (对象, 可选) 边框。
    ** 重要 ** 以下是一个示例，描述如何使用`text_box`, `image` 和 `shape` 元素进行综合布局
    ```json
        {{
      "layout_type": "title_slide_with_gradient_overlay",
      "elements": [
        {{
          "type": "image",
          "image_keyword": "abstract technology background blue",
          "x": 0, "y": 0, "width": 1280, "height": 720
        }},
        {{
          "type": "shape",
          "shape_type": "rectangle",
          "x": 0, "y": 0, "width": 1280, "height": 720,
          "style": {{
            "gradient": {{
              "angle": 45,
              "colors": ["#0D47A1", "#42A5F5"]
            }},
            "opacity": 0.85
          }}
        }},
        {{
          "type": "text_box",
          "content": "在渐变蒙版上的标题",
          "x": 100, "y": 300, "width": 1080, "height": 120,
          "style": {{
            "font": {{ "size": 60, "bold": true, "color": "#FFFFFF" }},
            "alignment": "CENTER"
          }}
        }}
      ]
    }}
    ```
    4.  **`chart`**
        * `type`: "chart"
        * `title`: (字符串) **[新要求]** 必须为图表提供一个清晰、简洁的标题。
        * `chart_type`: (字符串) 图表类型。可选值: `bar` (柱状图), `pie` (饼图), `line` (折线图)。
        * `data`: (对象)
            * `categories`: (字符串数组) 类别轴标签。
            * `series`: (对象数组) 每个对象是一组数据序列。
                * `name`: (字符串) **必须提供**，将用于图例显示。
                * `values`: (数字数组) 数据值。
        ** 美观的饼图示例: **
        ```json
        {{
          "type": "chart",
          "x": 140, 
          "y": 100, 
          "width": 1000, 
          "height": 550,
          "title": "用户对现有智能产品痛点分布",
          "chart_type": "pie",
          "data": {{
            "categories": [
              "操作复杂",
              "兼容性差",
              "隐私担忧",
              "功能单一",
              "价格过高"
            ],
            "series": [
              {{ 
                "name": "痛点分布", 
                "values": [30, 25, 20, 15, 10]
                }}
            ]
          }}
        }}
        ```
    5.  **`table`**
        * `type`: "table"
        * `headers`: (字符串数组) 表头。
        * `rows`: (二维字符串数组) 表格数据。
        * `style`: (对象, 可选) 定义表头/行颜色等。

    ---

    ### **第三部分：多样性与一致性核心准则 (Core Principles for Variety & Consistency)**

    1.  **布局多样性 (Layout Variety)**: **必须**混合使用多种 `layout_type`。**严禁**连续超过两页使用完全相同的简单布局。
    2.  **视觉元素丰富度 (Visual Richness)**: **每一张内容页都应至少包含一个视觉元素** (`image`, `shape`, `chart`, `table`)。
    3.  **设计系统贯穿始终 (Consistent Design System)**: 全局定义的 `color_palette` 和 `font_pairing` **必须**被应用到所有页面。
    4.  **色彩对比度与可读性 (Color Contrast & Readability)**: **这是一条绝对的、必须遵守的规则。**
        * 当文本被放置在任何有颜色的背景（如形状或图片）之上时，**必须确保文本颜色与背景色之间有足够高的对比度**，以保证内容清晰易读。
        * **具体准则：**
            * 在**深色背景**（如深蓝、黑色、深紫色）上，**必须使用浅色文字**（如白色 `#FFFFFF` 或极浅的灰色）。
            * 在**浅色背景**（如白色、米色、淡黄色）上，**必须使用深色文字**（如黑色 `#000000` 或极深的灰色）。
        * **严格禁止**：将颜色亮度相近的文本和背景放在一起（例如，在深灰色背景上放黑色文字，或在天蓝色背景上放白色文字）。
    5.  **字体策略 (Font Strategy)**:
            * **优先使用推荐字体**: 为了保证最佳兼容性，请**严格从以下列表中选择字体**。这些字体在绝大多数现代操作系统中都可用。
            * **中文推荐**:  **华文新魏 (STXinwei)**, **黑体 (SimHei)**, **华文行楷 (STXingkai)**, **楷体 (KaiTi)**, **等线 (Dengxian)**, **微软雅黑 (Microsoft YaHei)**。
            * **英文推荐**: **Arial**, **Calibri**, **Times New Roman**, **Verdana**, **Georgia**。
            * **创意与兜底**: 你可以为标题、引用等特殊文本使用列表中的字体进行创意组合。对于大段正文，如果没有特别的设计需求，使用 "微软雅黑" 或 "等线" 是最安全的选择。
            * **严格禁止**: 请**绝对不要使用** "思源黑体 (Source Han Sans)", "思源宋体 (Source Han Serif)", "苹方 (PingFang SC)" 或任何其他需要用户额外安装的字体。
    6.  **切勿在`content`字段的文本中使用任何Markdown语法**。
    7.  **所有的文本样式（如加粗）都必须通过`style`对象中的对应属性（如 `"bold": true`）来定义。**
    8.  **你的PPT页数应该严格与用户要求的页数一致**
    9.  **设计质量规则**: 每一页都必须承载明确的信息，严禁创建无实质内容的“过渡页”，也不要在一个页面中只放置一句话格言。
    
    ---

    ### **第四部分：设计风格指南 (Style Guide for Target Audience)**

    为了更好地贴合**中国女大学生**的审美，请遵循以下设计风格：

    1.  **设计理念 (Design Concept)**: 请使用更具诗意和画面感的词语。例如：“夏日橘子汽水”、“莫兰迪的午后”、“赛博蝴蝶梦”、“落日飞行”、“盐系手帐”。
    2.  **色彩运用 (Color Palette)**: 优先考虑**低饱和度的莫兰迪色系、柔和的马卡龙色系或高对比度的艺术撞色**。避免使用高饱和度的商务蓝、红色等传统商业配色，除非主题特别要求。
    3.  **版式布局 (Layout)**: 多采用**留白**，创造呼吸感。尝试不对称布局、图片网格、以及大号字体和图片的创意组合，营造杂志般的视觉效果。
    4.  **字体策略 (Font Strategy)**: **这是设计的灵魂，必须严格遵守。**
        * **核心原则：和谐源于对比与一致性。**
            * **经典搭配**: 在`font_pairing`中，优先选择一个**无衬线字体 (Sans-serif, 如 微软雅黑, 等线, 黑体)**用于正文，搭配一个**有衬线 (Serif, 如 宋体, 楷体) 或有设计感的无衬线字体 (如 华文行楷, 华文新魏)** 用于标题。这种对比清晰易读，且富有美感。
        * **建立清晰的视觉层级**:
            * **主标题 (页面大标题)**: 使用`heading`字体，字号最大 (如 36-48pt)，通常加粗。
            * **副标题/小标题**: 使用`heading`字体，字号中等 (如 24-32pt)，粗细可变。
            * **正文/列表**: 使用`body`字体，字号最小 (如 16-22pt)，使用常规体。
            * **你必须在整个PPT中保持这套层级规则的一致性。**
        * **风格与主题匹配**:
            * 你选择的`font_pairing`**必须**与你的`design_concept`在气质上保持一致。
            * **科技/简约风**: 多用“等线”、“微软雅黑”、“黑体”。
            * **人文/艺术/复古风**: 多用“宋体”、“楷体”、“华文行楷”。
            * **女性/柔美风**: 多用“等线 Light”、“微软雅黑 Light”或“楷体”。
        * **安全字体列表**: 为了保证最佳兼容性，请**严格从以下列表中选择字体**。
            * **中文推荐**:  **华文新魏 (STXinwei)**, **黑体 (SimHei)**, **华文行楷 (STXingkai)**, **楷体 (KaiTi)**, **等线 (Dengxian)**, **微软雅黑 (Microsoft YaHei)**, **宋体 (SimSun)**。
            * **英文推荐**: **Arial**, **Calibri**, **Times New Roman**, **Verdana**, **Georgia**。
            * **严格禁止**: 请**绝对不要使用** "思源黑体", "思源宋体", "苹方" 或任何需要用户额外安装的字体。


    ---

    ### **第五部分：输出样例（已加入新功能）**

    #### **样例一：团队介绍页 (使用圆形裁剪和半透明背景)**

    {{
      "design_concept": "盐系手帐：我们的故事",
      "font_pairing": {{ "heading": "Dengxian", "body": "Dengxian" }},
      "color_palette": {{ "primary": "#4A4A4A", "secondary": "#9B9B9B", "background": "#FDFBF8", "text": "#4A4A4A", "accent": "#F5A623" }},
      "master_slide": {{ "background": {{ "color": "#FDFBF8" }} }},
      "pages": [
        {{
          "layout_type": "team_introduction",
          "elements": [
            {{ "type": "text_box", "x": 100, "y": 80, "width": 1080, "height": 60, "content": "核心团队成员", "style": {{ "font": {{ "type": "heading", "size": 36, "bold": true }}, "alignment": "CENTER" }} }},
            {{ "type": "image", "x": 240, "y": 200, "width": 150, "height": 150, "image_keyword": "professional portrait of a smiling young woman", "style": {{ "crop": "circle" }} }},
            {{ "type": "text_box", "x": 215, "y": 360, "width": 200, "height": 60, "content": "张三\\n产品经理", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }} }},
            {{ "type": "image", "x": 565, "y": 200, "width": 150, "height": 150, "image_keyword": "professional portrait of a smiling young man", "style": {{ "crop": "circle" }} }},
            {{ "type": "text_box", "x": 540, "y": 360, "width": 200, "height": 60, "content": "李四\\n首席设计师", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }} }},
            {{ "type": "image", "x": 890, "y": 200, "width": 150, "height": 150, "image_keyword": "professional portrait of a friendly woman software developer", "style": {{ "crop": "circle" }} }},
            {{ "type": "text_box", "x": 865, "y": 360, "width": 200, "height": 60, "content": "王五\\n后端工程师", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }} }},
            {{ "type": "shape", "shape_type": "rectangle", "x": 0, "y": 550, "width": 1280, "height": 170, "style": {{ "fill_color": "#4A4A4A", "opacity": 0.1 }} }}
          ]
        }}
      ]
    }}

    #### **样例二：功能介绍页 (使用项目符号列表)**

    {{
      "design_concept": "夏日橘子汽水",
      "font_pairing": {{ "heading": "Microsoft YaHei Light", "body": "Microsoft YaHei Light" }},
      "color_palette": {{ "primary": "#FF6B6B", "secondary": "#FFD166", "background": "#FFFFFF", "text": "#4A4A4A", "accent": "#06D6A0" }},
      "master_slide": {{ "background": {{ "color": "#FFFFFF" }} }},
      "pages": [
        {{
          "layout_type": "image_left_content_right",
          "elements": [
            {{ "type": "image", "x": 0, "y": 0, "width": 640, "height": 720, "image_keyword": "vibrant flat illustration of a mobile app interface" }},
            {{ "type": "text_box", "x": 700, "y": 150, "width": 520, "height": 80, "content": "产品核心功能", "style": {{ "font": {{ "type": "heading", "size": 40, "bold": true }} }} }},
            {{ "type": "text_box", "x": 700, "y": 250, "width": 520, "height": 300, 
               "content": [
                 "AI智能规划：一句话生成完整演示文稿。",
                 "丰富的设计模板：覆盖多种行业和场景。",
                 "在线协同编辑：支持团队成员实时修改。",
                 "一键导出分享：轻松获取PPTX或PDF文件。"
               ],
               "style": {{ "font": {{ "type": "body", "size": 22 }}, "alignment": "LEFT" }}
            }}
          ]
        }}
      ]
    }}

    **最后指令：**
    现在，请**回顾并严格遵守以上所有部分的规则**，为主题 **“{theme}”** 生成一个包含 **{num_pages}** 页，宽为{canvas_width}，高为{canvas_height}的完整PPT设计方案JSON。
    """

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system",
                 "content": "You are a world-class presentation designer. Your output must be a single, raw JSON object. You must strictly follow all instructions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.55
        )

        response_content = response.choices[0].message.content
        if response_content:
            logging.info("已成功从AI接收到演示文稿方案。")
            json_string = _extract_json_from_response(response_content)
            if json_string:
                # 移除可能由模型生成的多余的尾随逗号
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