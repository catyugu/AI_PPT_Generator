# ai_service.py
import json
import logging
import re
import time
from typing import List, Optional, Dict, Any

from openai import OpenAI

from config import ONEAPI_KEY, ONEAPI_BASE_URL, MODEL_CONFIG, SYSTEM_PROMPT, MAX_LAYOUT_RETRIES, VALID_ICON_KEYWORDS

# ... (日志、客户端初始化、_extract_json_from_response, _call_ai 保持不变) ...
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
try:
    if not ONEAPI_KEY or ONEAPI_KEY == "YOUR_ONEAPI_KEY_HERE":
        raise ValueError("ONEAPI_KEY未在config.py或.env文件中正确设置。")
    client = OpenAI(api_key=ONEAPI_KEY, base_url=ONEAPI_BASE_URL)
    logging.info(f"OpenAI客户端已为OneAPI初始化，目标地址: {ONEAPI_BASE_URL}")
except ValueError as e:
    logging.error(e)
    client = None


def _extract_json_from_response(text: str) -> str | None:
    # 辅助函数，用于从可能包含杂质的文本中提取出JSON部分
    try:
        # 寻找第一个 '{' 和最后一个 '}'
        start_index = text.find('{')
        end_index = text.rfind('}')
        if start_index != -1 and end_index != -1 and end_index > start_index:
            json_block = text[start_index:end_index + 1]
            return json_block
        logging.warning("在AI响应中未找到有效的JSON对象边界。")
        return None
    except Exception as e:
        logging.error(f"提取JSON时出错: {e}")
        return None


def _call_ai_and_parse_json(prompt: str, model_name: str) -> dict | None:
    # 用于需要直接返回有效JSON的、较为简单的任务
    if not client:
        logging.error("OneAPI client not initialized.")
        return None
    try:
        logging.info(f"正在通过OneAPI调用模型 '{model_name}' 并期望获得JSON...")
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
        )
        response_content = response.choices[0].message.content
        if response_content:
            return json.loads(response_content) # 直接加载，因为开启了JSON模式
        logging.error("AI响应内容为空。")
        return None
    except json.JSONDecodeError as e:
        logging.error(f"JSON解码失败: {e}。原始响应: {response_content[:500]}")
        return None


def _call_ai_and_extract_manually(prompt: str, model_name: str) -> dict | None:
    # 手动提取JSON的后备方法
    if not client:
        logging.error("OneAPI client not initialized.")
        return None
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6
        )
        response_content = response.choices[0].message.content
        if response_content:
            json_string = _extract_json_from_response(response_content)
            if json_string:
                cleaned_json_string = re.sub(r',\s*([}\]])', r'\1', json_string) # 移除悬空逗号
                return json.loads(cleaned_json_string)
        logging.error("AI响应内容为空或提取JSON失败。")
        return None
    except json.JSONDecodeError as e:
        logging.error(f"手动提取后JSON解码失败: {e}。")
        return None
    except Exception as e:
        logging.error(f"与OneAPI通信时发生严重错误: {e}", exc_info=True)
        return None


def _get_raw_ai_response(prompt: str, model_name: str) -> str | None:
    # [新] 这个函数只获取AI的原始字符串响应，不做任何解析
    if not client:
        logging.error("OneAPI client not initialized.")
        return None
    try:
        logging.info(f"正在通过OneAPI调用模型 '{model_name}' 并获取原始字符串响应...")
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6
        )
        return response.choices[0].message.content
    except Exception as e:
        logging.error(f"与OneAPI通信时发生严重错误: {e}", exc_info=True)
        return None
# --- [最终版] 流水线阶段一：AI视觉总监 (配备专属案例) ---
def _generate_design_system(theme: str) -> dict | None:
    """生成PPT的核心设计风格，并提供高度相关的案例供AI学习。"""
    prompt = f"""
    你是一位深谙**年轻女性审美**的顶级演示文稿（PPT）设计大师。你精通平面设计、版式理论、色彩心理学和视觉传达。

    你的任务是为主题 **“{theme}”** 设计一个全局视觉系统。

    **你的输出必须是一个单一、完整、严格符合以下所有规则的原始JSON对象。**

    ### **设计风格指南 (Style Guide for Target Audience)**
    
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
             * **中文推荐**:  **华文新魏**, **黑体**, **华文行楷**, **楷体**, **等线**, **微软雅黑**, **宋体**。
             * **英文推荐**: **Arial**, **Calibri**, **Times New Roman**, **Verdana**, **Georgia**。
             * **严格禁止**: 请**绝对不要使用** "思源黑体", "思源宋体", "苹方" 或任何需要用户额外安装的字体。
    
    ### **输出样例 (参考案例)**
    你必须学习以下成功案例的风格和格式：

    #### 案例1: "盐系手帐"
    {{
      "design_concept": "盐系手帐：我们的故事",
      "font_pairing": {{ "heading": "行楷", "body": "等线" }},
      "color_palette": {{ "primary": "#4A4A4A", "secondary": "#9B9B9B", "background": "#FDFBF8", "text": "#4A4A4A", "accent": "#F5A623" }}
    }}

    #### 案例2: "夏日橘子汽水"
    {{
      "design_concept": "夏日橘子汽水",
      "font_pairing": {{ "heading": "黑体", "body": "等线" }},
      "color_palette": {{ "primary": "#FF6B6B", "secondary": "#FFD166", "background": "#FFFFFF", "text": "#4A4A4A", "accent": "#06D6A0" }}
    }}

    #### 案例3: "赛博蝴蝶梦"
    {{
        "design_concept": "赛博蝴蝶梦",
        "font_pairing": {{ "heading": "行楷", "body": "等线" }},
        "color_palette": {{ "primary": "#7B1FA2", "secondary": "#CE93D8", "background": "#121212", "text": "#FFFFFF", "accent": "#F48FB1" }}
    }}

    ---
    现在，请为主题 **“{theme}”** 生成一个全新的、符合上述所有要求的设计系统JSON。
    """
    model = MODEL_CONFIG.get("designer", "glm-4")
    return _call_ai_and_parse_json(prompt, model)


# --- [最终版] 流水线阶段二：AI信息架构师 (配备专属案例) ---
def _generate_content_outline(theme: str, num_pages: int) -> dict | None:
    """生成PPT的内容大纲，并提供高度相关的案例供AI学习。"""
    prompt = f"""
    你是一位顶级的信息架构专家。你的任务是为主题 **“{theme}”**，规划一个逻辑清晰、内容充实的演示文稿大纲，总页数严格为 **{num_pages}** 页。

    **核心准则**:
    1.  **页数严格**: 你的输出必须不多不少，正好包含 {num_pages} 页的大纲。
    2.  **内容明确**: 每一页都必须承载明确的信息，严禁创建无实质内容的“过渡页”，也不要在一个页面中只放置一句话格言。
    3.  **逻辑连贯**: 从封面到内容再到封底，必须形成一条完整的故事线。

    **你的输出必须是严格的JSON格式，包含一个`pages`键，其值为一个数组。**
    数组中的每个对象代表一页，且必须包含`title` (字符串) 和 `content` (字符串或字符串数组) 两个键。
    -   `title` 是本页的核心标题。
    -   `content` 是本页的核心内容。如果是普通文本，则为字符串，支持用 `\\n` 换行。如果要创建项目符号列表，则**必须**使用字符串数组。

    **JSON输出示例**:

    #### 案例1: 团队介绍页
    {{
        "pages": [
            {{
                "title": "核心团队成员",
                "content": [
                    "张三\\n产品经理",
                    "李四\\n首席设计师",
                    "王五\\n后端工程师"
                ]
            }}
        ]
    }}

    #### 案例2: 功能介绍页
    {{
        "pages": [
            {{
                "title": "产品核心功能",
                "content": [
                    "AI智能规划：一句话生成完整演示文稿。",
                    "丰富的设计模板：覆盖多种行业和场景。",
                    "在线协同编辑：支持团队成员实时修改。",
                    "一键导出分享：轻松获取PPTX或PDF文件。"
                ]
            }}
        ]
    }}

    #### 案例3: 核心优势页
    {{
        "pages": [
            {{
                "title": "我们的核心优势",
                "content": [
                    "**AI强力驱动**\\n基于最新的AI模型，理解能力和生成质量遥遥领先。",
                    "**高度可定制**\\n从颜色到字体，从布局到内容，一切尽在掌握。",
                    "**完全免费开源**\\n我们相信知识共享的力量，欢迎共建。"
                ]
            }}
        ]
    }}
    ---
    现在，请为主题 **“{theme}”** 生成一个全新的、包含 **{num_pages}** 页的完整内容大纲JSON。
    """
    model = MODEL_CONFIG.get("writer", "glm-4")
    return _call_ai_and_parse_json(prompt, model)

# --- [最终版] 流水线阶段三：AI排版导演 (配备专属案例) ---
def _generate_slide_layout(
    page_content: dict,
    design_system: dict,
    all_pages_outline: list,
    page_index: int,
    total_pages: int,
    aspect_ratio: str,
    previous_slide_layout: Optional[Dict] = None,  # <<< 新增参数
) -> str | None:
    """为单页内容设计布局和动画，并提供高度相关的案例及上一页的布局作为上下文。"""
    canvas_width, canvas_height = (1024, 768) if aspect_ratio == "4:3" else (1280, 720)

    full_outline_str = "\n".join(
        [f"- 第{i + 1}页: {p.get('title', '无标题')}" for i, p in enumerate(all_pages_outline)]
    )

    # <<< 新增逻辑：构建上一页的上下文Prompt >>>
    context_prompt = ""
    if page_index > 0 and previous_slide_layout:
        # 为了防止Prompt过长，可以只包含关键信息或整个JSON
        previous_layout_str = json.dumps(previous_slide_layout, ensure_ascii=False, indent=2)
        context_prompt = f"""
            ---
            ### **重要参考：上一页 (第 {page_index} 页) 的最终布局**
            这是你刚刚完成的前一页的布局JSON。请仔细分析它的结构、`layout_type`、图片位置和元素排布。
            
            ```json
            {previous_layout_str}
            本页设计指导原则
            寻求变化，避免重复: 这是最重要的规则！ 请绝对不要与上一页使用完全相同的布局结构。例如，如果上一页是“左图右文”，请在本页中创造性地使用“右图左文”、“全屏大图”、“三栏对比”等不同版式。
            
            保持风格连贯: 布局结构需要变化，但整体的设计感（如颜色、字体、间距）应与全局规范和上一页保持和谐。
            
            聚焦当前任务: 以下才是你需要为当前页面（第 {page_index + 1} 页）设计的具体内容。
            """
    prompt = f"""
       你是一位精通布局、动画和视觉传达的演示文稿排版导演。
       你的任务是为一部拥有完整剧本的演示文稿，设计其中**一页**的详细视觉方案。
       
       {context_prompt}

       ### **剧本大纲 (全局上下文)**
       这是整个演示文稿的故事线，你要设计的页面是其中的一部分：
       {full_outline_str}

       ---
    **全局设计规范 (必须遵守):**
    - **设计理念**: {design_system.get('design_concept', '无')}
    - **字体搭配**: {json.dumps(design_system.get('font_pairing'), ensure_ascii=False)}
    - **全局色板**: {json.dumps(design_system.get('color_palette'), ensure_ascii=False)}

    **当前页面内容 (需要你来设计布局):**
    - **标题**: {page_content.get('title', '无标题')}
    - **核心内容**: {json.dumps(page_content.get('content', ''), ensure_ascii=False)}
    
    **你的输出必须是一个严格的JSON对象，包含`layout_type`, `elements`, 和可选的`animation_sequence`。**

    ---
    ### **核心设计规则 (必须严格遵守)**
    1.  **画布尺寸**: 所有坐标和尺寸都基于 **{canvas_width}x{canvas_height}** 的画布。
    2.  **布局多样性**: 这是第 {page_index + 1} / {total_pages} 页。请考虑上下文，选择一个合适的 `layout_type` (如 `title_slide`, `image_left_content_right", "full_screen_image_with_quote", "three_column_comparison", "data_chart_summary", "team_introduction`)。**避免与可能的上一页布局过于雷同。**
    3.  **视觉元素**: **内容页应至少包含一个视觉元素** (`image`, `shape`, `chart`, `table`, `icon`)。
    4.  **色彩对比度**: **这是一条绝对规则！** 在任何背景上放置文本时，必须确保高对比度。深色背景配浅色文字，浅色背景配深色文字。
    5.  **元素定义**: 严格遵循以下JSON结构。所有需要动画的元素**必须**包含唯一的`id`。

    ---
    #### **元素 (Element) 定义**
     **所有元素都必须包含** `type`, `x`, `y`, `width`, `height` 这五个基本属性。
     此外，所有元素**可以包含** `z_index` (数字, 可选) 属性，用于控制元素的堆叠顺序（图层）。`z_index`值越大，元素越靠上。如果未指定，系统将根据元素类型使用默认层级。
     **重要：所有计划进行动画的元素（例如图片、文本框、形状、图表、表格、图标）都必须包含一个唯一的 `id` (字符串) 属性，例如 `\"id\": \"my_element_id\"`。** 这个ID将用于在 `animation_sequence` 中引用该元素。
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
             * `opacity`: (数字, 0.0-1.0) 图片整体的不透明度， 0.0表示完全透明，1.0表示完全可见。
             * `border`: (对象) 边框。包含 `color` (Hex) 和 `width` (px)。
             * `crop`: (字符串, 可选) **[新功能]** 裁剪形状。目前唯一支持的值是 **`"circle"`**，用于将图片裁剪为圆形。

     3.  **`shape`**
         * `type`: "shape"
         * `shape_type`: (字符串) 形状类型。可选值: `rectangle`, `oval`, `triangle`, `star`, `rounded_rectangle`。
         * `style`: (对象)
             * `fill_color`: (字符串, Hex, 可选) 填充色。
             * `opacity`: (数字, 0.0-1.0, 可选) 填充色的不透明度。0.0表示完全透明，1.0表示完全可见。
             * `gradient`: (对象, 可选, 与`fill_color`互斥) 渐变填充。
             * `border`: (对象, 可选) 边框。

     4.  **`chart`**
         * `type`: "chart"
         * `title`: (字符串) **[新要求]** 必须为图表提供一个清晰、简洁的标题。
         * `chart_type`: (字符串) 图表类型。可选值: `bar` (柱状图), `pie` (饼图), `line` (折线图)。
         * `data`: (对象)
             * `categories`: (字符串数组) 类别轴标签。
             * `series`: (对象数组) 每个对象是一组数据序列。
                 * `name`: (字符串) **必须提供**，将用于图例显示。
                 * `values`: (数字数组) 数据值。

     5.  **`table`**
         * `type`: "table"
         * `headers`: (字符串数组) 表头。
         * `rows`: (二维字符串数组) 表格数据。
         * `style`: (对象, 可选) 定义表头/行颜色等。

     6.  **`icon`** * `type`: "icon"
         * `icon_keyword`: (字符串, 必须) 描述图标核心含义的**一个小写英文关键词**。你**只能**从以下可用列表中选择图标关键词。** 这个关键词将直接用于查找对应的SVG图标文件:
             * [`activity`, `airplay`, `alert-circle`, `alert-octagon`, `alert-triangle`, `align-center`, `align-justify`, `align-left`, `align-right`, `anchor`, `aperture`, `archive`, `arrow-down-circle`, `arrow-down-left`, `arrow-down-right`, `arrow-down`, `arrow-left-circle`, `arrow-left`, `arrow-right-circle`, `arrow-right`, `arrow-up-circle`, `arrow-up-left`, `arrow-up-right`, `arrow-up`, `at-sign`, `award`, `bar-chart-2`, `bar-chart`, `battery-charging`, `battery`, `bell-off`, `bell`, `bluetooth`, `bold`, `book-open`, `book`, `bookmark`, `box`, `briefcase`, `calendar`, `camera-off`, `camera`, `cast`, `check-circle`, `check-square`, `check`, `chevron-down`, `chevron-left`, `chevron-right`, `chevron-up`, `chevrons-down`, `chevrons-left`, `chevrons-right`, `chevrons-up`, `chrome`, `circle`, `clipboard`, `clock`, `cloud-drizzle`, `cloud-lightning`, `cloud-off`, `cloud-rain`, `cloud-snow`, `cloud`, `code`, `codepen`, `codesandbox`, `coffee`, `columns`, `command`, `compass`, `copy`, `corner-down-left`, `corner-down-right`, `corner-left-down`, `corner-left-up`, `corner-right-down`, `corner-right-up`, `corner-up-left`, `corner-up-right`, `cpu`, `credit-card`, `crop`, `crosshair`, `database`, `delete`, `disc`, `divide-circle`, `divide-square`, `divide`, `dollar-sign`, `download-cloud`, `download`, `dribbble`, `droplet`, `edit-2`, `edit-3`, `edit`, `external-link`, `eye-off`, `eye`, `facebook`, `fast-forward`, `feather`, `figma`, `file-minus`, `file-plus`, `file-text`, `file`, `film`, `filter`, `flag`, `folder-minus`, `folder-plus`, `folder`, `framer`, `frown`, `gift`, `git-branch`, `git-commit`, `git-merge`, `git-pull-request`, `github`, `gitlab`, `globe`, `grid`, `hard-drive`, `hash`, `headphones`, `heart`, `help-circle`, `hexagon`, `home`, `image`, `inbox`, `info`, `instagram`, `italic`, `key`, `layers`, `layout`, `life-buoy`, `link-2`, `link`, `linkedin`, `list`, `loader`, `lock`, `log-in`, `log-out`, `mail`, `map-pin`, `map`, `maximize-2`, `maximize`, `meh`, `menu`, `message-circle`, `message-square`, `mic-off`, `mic`, `minimize-2`, `minimize`, `minus-circle`, `minus-square`, `minus`, `monitor`, `moon`, `more-horizontal`, `more-vertical`, `mouse-pointer`, `move`, `music`, `navigation-2`, `navigation`, `octagon`, `package`, `paperclip`, `pause-circle`, `pause`, `pen-tool`, `percent`, `phone-call`, `phone-forwarded`, `phone-incoming`, `phone-missed`, `phone-off`, `phone-outgoing`, `phone`, `pie-chart`, `play-circle`, `play`, `plus-circle`, `plus-square`, `plus`, `pocket`, `power`, `printer`, `radio`, `refresh-ccw`, `refresh-cw`, `repeat`, `rewind`, `rotate-ccw`, `rotate-cw`, `rss`, `save`, `scissors`, `search`, `send`, `server`, `settings`, `share-2`, `share`, `shield-off`, `shield`, `shopping-bag`, `shopping-cart`, `shuffle`, `sidebar`, `skip-back`, `skip-forward`, `slack`, `slash`, `sliders`, `smartphone`, `smile`, `speaker`, `square`, `star`, `stop-circle`, `sun`, `sunrise`, `sunset`, `table`, `tablet`, `tag`, `target`, `terminal`, `thermometer`, `thumbs-down`, `thumbs-up`, `toggle-left`, `toggle-right`, `tool`, `trash-2`, `trash`, `trello`, `trending-down`, `trending-up`, `triangle`, `truck`, `tv`, `twitch`, `twitter`, `type`, `umbrella`, `underline`, `unlock`, `upload-cloud`, `upload`, `user-check`, `user-minus`, `user-plus`, `user-x`, `user`, `users`, `video-off`, `video`, `voicemail`, `volume-1`, `volume-2`, `volume-x`, `volume`, `watch`, `wifi-off`, `wifi`, `wind`, `x-circle`, `x-octagon`, `x-square`, `x`, `youtube`, `zap-off`, `zap`, `zoom-in`, `zoom-out`]
         * **重要**: `x`, `y`, `width`, `height` 仍然是必须的。为保持图标比例，建议将 `width` 和 `height` 设置为相同或相近的小数值（例如，都设为0.05）。
    #### **页面级动画序列 (Animation Sequence) [全新功能]**
    在`pages`的每个页面对象中，现在可以包含一个可选的 `animation_sequence` 数组。

    * `animation_sequence`: (数组, 可选) **定义了本页所有动画的触发顺序**。数组中的每个对象代表一个“单击步骤”。
    * **序列对象的结构**:
        * `element_id`: (字符串, 必须) 引用本页中某个元素的 `id`。
        * `animation`: (对象, 必须) 定义要对该元素执行的动画。

    #### **动画 (Animation) 对象定义**
    * `type`: "fadeIn", "fadeOut", "flyIn", "flyOut"
    * `duration_ms`: (数字, 可选)
    * `direction`: (字符串, 可选)
    
    ---
    ### **动画导演指南 (Animation Director's Guide)**
    - 你现在是一个“导演”。你需要通过 `animation_sequence` 数组来编排整个幻灯片的动画故事线。
    - **思考顺序**: 先布局所有静态元素，并为它们分配好ID。然后，在 `animation_sequence` 中，像写剧本一样，一步一步（一次单击）地定义哪个元素以何种方式出现或消失。
    - **动画类型**: `type` 可为 "fadeIn", "fadeOut", "flyIn", "flyOut"。可附加 `duration_ms` 和 `direction`。
    - **范例**: 一个经典的序列是：主标题 `fadeIn` -> (下一次单击) -> 主图片 `flyIn` -> (下一次单击) -> 要点一 `flyIn`。

    ---
    
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
             * **中文推荐**:  **华文新魏**, **黑体**, **华文行楷**, **楷体**, **等线**, **微软雅黑**, **宋体**。
             * **英文推荐**: **Arial**, **Calibri**, **Times New Roman**, **Verdana**, **Georgia**。
             * **严格禁止**: 请**绝对不要使用** "思源黑体", "思源宋体", "苹方" 或任何需要用户额外安装的字体。


        ---
    ### **输出样例 (参考案例) **
    你必须学习以下**所有**成功案例的布局和动画编排手法，并根据当前页面的内容，创造性地应用它们。

    #### **案例1: 全屏图片引言页 (Full Screen Image with Quote)**
    {{
      "layout_type": "full_screen_image_with_quote",
      "elements": [
        {{ "id": "background_image", "type": "image", "x": 0, "y": 0, "width": {canvas_width}, "height": {canvas_height}, "image_keyword": "tranquil landscape with morning mist", "style": {{ "opacity": 0.8 }}, "z_index": 5 }},
        {{ "id": "quote_text", "type": "text_box", "x": {int(canvas_width*0.1)}, "y": {int(canvas_height*0.4)}, "width": {int(canvas_width*0.8)}, "height": {int(canvas_height*0.2)}, "content": "“设计的本质，是让生活更美好。”", "style": {{ "font": {{ "type": "heading", "size": 48, "color": "#FFFFFF", "bold": true }}, "alignment": "CENTER" }}, "z_index": 20 }},
        {{ "id": "author_text", "type": "text_box", "x": {int(canvas_width*0.5)}, "y": {int(canvas_height*0.55)}, "width": {int(canvas_width*0.4)}, "height": {int(canvas_height*0.1)}, "content": "—— 著名设计师", "style": {{ "font": {{ "type": "body", "size": 22, "color": "#FFFFFF" }}, "alignment": "RIGHT" }}, "z_index": 20 }}
      ],
      "animation_sequence": [
        {{ "element_id": "background_image", "animation": {{ "type": "fadeIn", "duration_ms": 1000 }} }},
        {{ "element_id": "quote_text", "animation": {{ "type": "flyIn", "direction": "fromTop", "duration_ms": 800 }} }}
      ]
    }}

    #### **案例2: 三列对比/要点页 (Three Column Comparison)**
    {{
      "layout_type": "three_column_comparison",
      "elements": [
        {{ "id": "title", "type": "text_box", "x": 100, "y": 60, "width": 1080, "height": 60, "content": "我们的三大核心优势", "style": {{ "font": {{ "type": "heading", "size": 40, "bold": true }}, "alignment": "CENTER" }} }},
        {{ "id": "col1_icon", "type": "icon", "x": 200, "y": 200, "width": 80, "height": 80, "icon_keyword": "cpu"}},
        {{ "id": "col1_title", "type": "text_box", "x": 150, "y": 300, "width": 180, "height": 40, "content": "AI强力驱动", "style": {{ "font": {{ "type": "heading", "size": 24, "bold": true }}, "alignment": "CENTER" }} }},
        {{ "id": "col1_text", "type": "text_box", "x": 150, "y": 350, "width": 180, "height": 200, "content": "基于最新模型，理解能力和生成质量遥遥领先。", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }} }},
        {{ "id": "col2_icon", "type": "icon", "x": 600, "y": 200, "width": 80, "height": 80, "icon_keyword": "settings"}},
        {{ "id": "col2_title", "type": "text_box", "x": 550, "y": 300, "width": 180, "height": 40, "content": "高度可定制", "style": {{ "font": {{ "type": "heading", "size": 24, "bold": true }}, "alignment": "CENTER" }} }},
        {{ "id": "col2_text", "type": "text_box", "x": 550, "y": 350, "width": 180, "height": 200, "content": "从颜色到字体，从布局到内容，一切尽在掌握。", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }} }}
      ],
      "animation_sequence": [
        {{ "element_id": "title", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "col1_icon", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }},
        {{ "element_id": "col2_icon", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }}
      ]
    }}

    #### **案例3: 数据图表总结页 (Data Chart Summary)**
    {{
      "layout_type": "data_chart_summary",
      "elements": [
        {{ "id": "chart_title", "type": "text_box", "x": 100, "y": 80, "width": 1080, "height": 60, "content": "近三年用户增长趋势", "style": {{ "font": {{ "type": "heading", "size": 36, "bold": true }}, "alignment": "LEFT" }} }},
        {{ "id": "main_chart", "type": "chart", "x": 100, "y": 180, "width": 800, "height": 450, 
           "chart_type": "line", "title": "用户增长曲线",
           "data": {{ 
             "categories": ["2022", "2023", "2024"], 
             "series": [
               {{ "name": "App下载量 (万)", "values": [120, 250, 480] }},
               {{ "name": "网站访问量 (万)", "values": [300, 550, 880] }}
             ]
           }}
        }},
        {{ "id": "chart_conclusion", "type": "text_box", "x": 950, "y": 250, "width": 250, "height": 300, "content": "**结论:**\\n- App下载量呈现指数级增长。\\n- 网站流量稳步提升，市场认知度扩大。", "style": {{ "font": {{ "type": "body", "size": 20 }} }} }}
      ],
       "animation_sequence": [
        {{ "element_id": "chart_title", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "main_chart", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "chart_conclusion", "animation": {{ "type": "flyIn", "direction": "fromRight" }} }}
      ]
    }}

    #### **案例4: 流程/时间轴页 (Process/Timeline)**
    {{
        "layout_type": "process_timeline",
        "elements": [
            {{ "id": "proc_title", "type": "text_box", "x": 100, "y": 80, "width": 1080, "height": 60, "content": "我们的四步服务流程", "style": {{ "font": {{ "type": "heading", "size": 36, "bold": true }}, "alignment": "CENTER" }} }},
            {{ "id": "arrow_line", "type": "shape", "shape_type": "rectangle", "x": 200, "y": 358, "width": 880, "height": 4, "style": {{ "fill_color": "#CE93D8", "opacity": 0.5 }}, "z_index": 10}},
            {{ "id": "step1_circle", "type": "shape", "shape_type": "oval", "x": 180, "y": 320, "width": 80, "height": 80, "style": {{ "fill_color": "#7B1FA2" }}, "z_index": 20}},
            {{ "id": "step1_text", "type": "text_box", "x": 155, "y": 420, "width": 130, "height": 60, "content": "**第一步**\\n需求沟通", "style": {{ "font": {{ "type": "body", "size": 18 }}, "alignment": "CENTER" }} }},
            {{ "id": "step2_circle", "type": "shape", "shape_type": "oval", "x": 460, "y": 320, "width": 80, "height": 80, "style": {{ "fill_color": "#7B1FA2" }}, "z_index": 20}}
        ],
        "animation_sequence": [
            {{ "element_id": "proc_title", "animation": {{ "type": "fadeIn"}} }},
            {{ "element_id": "arrow_line", "animation": {{ "type": "fadeIn", "duration_ms": 500 }} }},
            {{ "element_id": "step1_circle", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }},
            {{ "element_id": "step2_circle", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }}
        ]
    }}
    ---
    现在，请为当前页面设计布局和动画，并输出最终的JSON对象。
    """
    model = MODEL_CONFIG.get("planner", "glm-4")
    return _get_raw_ai_response(prompt, model)  # <<< 关键：调用获取原始字符串的函数

def _validate_and_correct_slide_layout(slide_json_str: str, canvas_width: int, canvas_height: int) -> dict | None:
    """
    [V2版] 引入一个更智能、更挑剔的AI首席视觉官，负责审查和升华单页布局JSON。
    """
    prompt = f"""
    你是一位极度挑剔且才华横溢的演示文稿（PPT）首席视觉官(Chief Visual Officer)。你的审美标准极高，对细节的把控到了像素级别。你的任务是接收一个由初级设计师完成的JSON布局草稿，不仅要修复其中明显的错误，更要以你的专业眼光，将它打磨成一件平衡、精致、专业且富有美感的艺术品。

    你的画布尺寸为 {canvas_width}x{canvas_height} 像素。

    ---
    ### **第一阶段：硬性错误修复 (Technical Bug Fixing)**
    你必须首先修复所有技术层面的硬伤。

    1.  **JSON格式校验**: 确保输入是完整且语法正确的JSON。如果存在语法错误（如悬空的逗号），必须修复。
    2.  **图标库校验**: 检查所有 `"type": "icon"` 元素，其 `icon_keyword` **必须**存在于官方列表中。如果无效，从列表中找一个**语义最相近的**替换它。
        * **官方图标库**: {json.dumps(VALID_ICON_KEYWORDS)}
    3.  **动画逻辑校验**: 检查 `animation_sequence`，确保每个 `element_id` 都真实存在于 `elements` 数组中。如果动画指向了不存在的`element_id`，**果断从 `animation_sequence` 数组中删除该动画对象**。

    ---
    ### **第二阶段：专业布局审查与优化 (Professional Layout Review & Enhancement)**
    这是你的核心价值所在。在脑海中建立一个**虚拟网格系统**和**80像素的安全边距**。所有核心内容都应位于安全边距之内。

    #### **A. 元素碰撞与重叠检测 (Collision & Overlap Detection)**
    你必须智能地判断元素重叠是“设计意图”还是“布局事故”。

    * **允许的层叠 (Intentional Layering)**:
        * **特征**: 通常涉及不同类型的元素（如文字在形状上），且`z_index`明确定义了层级关系。
        * **你的行动**: **保留并尊重**这种设计。但要检查上层文本是否清晰可读，如果背景过于花哨导致文字看不清，你需要主动为文字背景添加一个半透明的深色或浅色衬底（一个`shape`元素）来保证可读性。

    * **不允许的碰撞 (Unintentional Collision)**:
        * **特征**: 两个同类元素（如两个文本框）互相重叠；关键信息（如标题）被图片部分遮挡；元素间距过小，显得拥挤不堪。
        * **你的行动**: 必须修正！

    #### **B. 文本溢出预判与修正 (Text Overflow Prediction & Correction)**
    **不要相信原始的 `height` 值！** 你必须基于 `content`, `font.size`, 和 `width` 智能预估文本渲染后的实际高度。

    * **预估方法**: 一个18号字，在400宽的盒子里，每行大约能放20个中文字。用这个方法估算出行数，再乘以 `fontSize * 1.5` 得到渲染高度。
    * **修正策略**:
        1.  **首先，尝试增加 `height`**: 这是最直接的方法。
        2.  **其次，移动下方元素**: 如果增加高度会导致新的碰撞，则**整体下移**被碰撞的元素，保持原有的垂直间距。

    #### **C. 对齐与分布 (Alignment & Distribution)**
    专业感源于秩序。这是**必须执行**的步骤。

    1.  **对齐检查**: 寻找页面中可以建立“对齐线”的元素。例如，左侧的图片和下方的文本块，它们的左边缘是否对齐了？标题和正文是否左对齐？
    2.  **分布检查**: 如果有一组横向或纵向排列的元素（如三个卖点介绍），检查它们之间的**间距是否完全相等**。
    3.  **修正策略**: **果断地修改`x`, `y`坐标**，让元素严格对齐或等距分布。允许你进行±20像素范围内的精细调整。

    ---
    ### **第三阶段：最终美学审查 (Final Aesthetic Review)**
    在完成所有技术和布局修正后，退后一步，像审视一幅画一样审视整个页面。

    1.  **视觉平衡 (Visual Balance)**: 页面的“视觉重量”是否均衡？有没有头重脚轻，或者所有元素都挤在左上角？**你的行动**: 如果不平衡，请大胆地移动元素，甚至可以整体移动一组元素，以达到视觉上的和谐。
    2.  **留白与呼吸感 (White Space & Breathability)**: 页面是否塞得过满？元素之间是否有足够的“呼吸空间”？**你的行动**: 如果过于拥挤，可以适当缩小某些非核心元素（如装饰性形状或图片）的 `width` 和 `height`，或者增加元素间距。
    3.  **视觉焦点 (Focal Point)**: 这一页的视觉焦点是什么？是标题，还是核心图片？这个焦点是否足够突出？**你的行动**: 如果焦点不明确，可以适当**加粗或微调字号**来强调核心标题，或者给核心图片增加一个微妙的边框或阴影效果（通过增加`shape`元素实现）。

    ---
    ### **输出指令**
    * **输入**: 一个可能存在各种问题的JSON草稿字符串。
    * **输出**: **必须且只能**是一个经过你三轮严格审查和精心打磨后，达到出版级别的、完美的、单一的JSON对象。**绝对不要添加任何解释性文字或代码块标记。**

    ---
    **现在，请对以下JSON草稿进行你严苛的审查和修正：**

    ```json
    {slide_json_str}
    ```
    """
    model = MODEL_CONFIG.get("inspector", "glm-4")
    # 质检员必须返回可解析的JSON，所以这里用回 _call_ai_and_parse_json
    return _call_ai_and_parse_json(prompt, model)



def _generate_slide_layout_with_retry(
        page_content: dict,
        design_system: dict,
        all_pages_outline: list,
        page_index: int,
        total_pages: int,
        aspect_ratio: str,
        previous_slide_layout: Optional[Dict] = None,
) -> dict | None:
    """
    [核心重构] 编排“生成草稿->送交质检”的完整流程，并集成重试逻辑。
    """


    canvas_width, canvas_height = (1024, 768) if aspect_ratio == "4:3" else (1280, 720)
    for attempt in range(MAX_LAYOUT_RETRIES):
        logging.info(f"  - 正在处理第 {page_index + 1} 页布局 (尝试 {attempt + 1}/{MAX_LAYOUT_RETRIES})...")
        try:
            # 步骤1: 先生成原始、可能不完美的字符串草稿
            logging.info(f"    - (1/2) AI排版导演正在生成布局草稿...")
            slide_layout_draft_str = _generate_slide_layout(
                page_content,
                design_system,
                all_pages_outline,
                page_index,
                total_pages,
                aspect_ratio,
                previous_slide_layout,
            )

            # 如果第一步就失败了（比如API不通），直接进入下一次重试
            if not slide_layout_draft_str:
                logging.warning(f"    - (1/2) AI排版导演未能生成草稿，将在等待后重试。")
                time.sleep(2 * (attempt + 1))
                continue

            logging.info(f"    - (2/2) 草稿已生成，送往AI首席视觉官进行审查和修正...")
            # 步骤2: 将原始草稿（无论好坏）送交“质检员”进行修正和解析
            corrected_layout = _validate_and_correct_slide_layout(
                slide_layout_draft_str,
                canvas_width,
                canvas_height
            )

            # 成功条件：质检员成功返回了一个有效的、修正后的JSON对象
            if corrected_layout and "elements" in corrected_layout:
                logging.info(f"  - 第 {page_index + 1} 页布局处理成功。")
                return corrected_layout

            logging.warning(f"  - (2/2) AI首席视觉官未能修正或返回有效布局。")

        except Exception as e:
            logging.error(f"  - 处理第 {page_index + 1} 页布局时发生意外异常: {e}", exc_info=True)

        # 如果不是最后一次尝试，则等待一段时间后重试
        if attempt < MAX_LAYOUT_RETRIES - 1:
            wait_time = (attempt + 1) * 2
            logging.info(f"  - 将在 {wait_time} 秒后重试...")
            time.sleep(wait_time)

        logging.error(f"  - 经过 {MAX_LAYOUT_RETRIES} 次尝试后，仍无法为第 {page_index + 1} 页生成有效布局。")
        return None  # 所有尝试均告失败

# --- [最终版] 主函数：编排“拥有案例手册的专家流水线” ---

def generate_presentation_pipeline(theme: str, num_pages: int, aspect_ratio: str = "16:9") -> dict | None:
    """
    [V3 健壮版] 最终版流水线，采用串行上下文感知+草稿送审机制。
    """
    logging.info("--- 开始AI专家流水线生成任务 (V3 - 上下文感知+草稿送审) ---")

    # 阶段一、二保持不变
    design_system = _generate_design_system(theme)
    if not design_system: return None
    content_outline = _generate_content_outline(theme, num_pages)
    if not content_outline or 'pages' not in content_outline: return None

    all_pages_content = content_outline['pages']
    final_pages: List[Optional[Dict[str, Any]]] = [None] * len(all_pages_content)
    canvas_width, canvas_height = (1024, 768) if aspect_ratio == "4:3" else (1280, 720)

    # 串行循环，维持上下文
    previous_slide_layout = None
    for i, page_content in enumerate(all_pages_content):
        # [核心修改] 调用重构后的、包含完整“生成-修正”流程的函数
        final_layout = _generate_slide_layout_with_retry(
            page_content=page_content,
            design_system=design_system,
            all_pages_outline=all_pages_content,
            page_index=i,
            total_pages=len(all_pages_content),
            aspect_ratio=aspect_ratio,
            previous_slide_layout=previous_slide_layout,
        )

        if final_layout:
            final_pages[i] = final_layout
            previous_slide_layout = final_layout  # 更新上下文为最终修正后的版本
        else:
            # 容错逻辑：如果重试多次后仍然失败，则创建一个错误提示页面
            logging.error(f"  - 第 {i + 1} 页布局生成永久失败，将使用错误提示页。")
            error_page = {
                "layout_type": "fallback_error",
                "elements": [{
                    "id": "error_text", "type": "text_box", "content": f"页面 {i + 1}\n生成失败",
                    "x": int(canvas_width * 0.1), "y": int(canvas_height * 0.4),
                    "width": int(canvas_width * 0.8), "height": int(canvas_height * 0.2),
                    "style": {"font": {"type": "heading", "size": 48, "color": "#E53935", "bold": True},
                              "alignment": "CENTER"}
                }]
            }
            final_pages[i] = error_page
            previous_slide_layout = error_page # 即使失败，也更新上下文，避免好坏页交错

    # 组装最终计划
    if all(p is None for p in final_pages):
        logging.error("所有页面布局均生成失败，任务中止。")
        return None

    final_plan = {
        **design_system,
        "master_slide": {"background": {"color": design_system.get("color_palette", {}).get("background", "#FFFFFF")}},
        "pages": [p for p in final_pages if p]
    }
    logging.info("--- AI专家流水线任务全部完成 ---")
    return final_plan

# --- 无缝切换接口 ---
generate_presentation_plan = generate_presentation_pipeline