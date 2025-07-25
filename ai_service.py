# ai_service.py
import json
import logging
import re
from concurrent.futures import as_completed, ThreadPoolExecutor
from typing import List, Optional, Dict, Any

from openai import OpenAI

from config import ONEAPI_KEY, ONEAPI_BASE_URL, MODEL_CONFIG, SYSTEM_PROMPT

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
    # ... (此辅助函数保持不变) ...
    try:
        text = re.sub(r'```json\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'```', '', text)
        start_index = text.find('{')
        end_index = text.rfind('}')
        if start_index != -1 and end_index != -1 and end_index > start_index:
            json_block = text[start_index:end_index + 1]
            json_block = re.sub(r'//.*', '', json_block)
            json_block = re.sub(r'/\*.*?\*/', '', json_block, flags=re.DOTALL)
            return json_block
        logging.warning("在AI响应中未找到有效的JSON对象边界。")
        return None
    except Exception as e:
        logging.error(f"提取和清理JSON时出错: {e}")
        return None


def _call_ai(prompt: str, model_name: str) -> dict | None:
    # ... (此通用调用函数保持不变) ...
    if not client:
        logging.error("OneAPI client not initialized.")
        return None
    try:
        logging.info(f"正在通过OneAPI调用模型 '{model_name}'...")
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            max_tokens=16000,
            temperature=0.6
        )
        response_content = response.choices[0].message.content
        if response_content:
            json_string = _extract_json_from_response(response_content)
            if json_string:
                cleaned_json_string = re.sub(r',\s*([}\]])', r'\1', json_string)
                return json.loads(cleaned_json_string)
        logging.error("AI响应内容为空或提取JSON失败。")
        return None
    except json.JSONDecodeError as e:
        logging.error(f"JSON解码失败: {e}。")
        return None
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
    return _call_ai(prompt, model)


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
    return _call_ai(prompt, model)


# --- [最终版] 流水线阶段三：AI排版导演 (配备专属案例) ---
def _generate_slide_layout(page_content: dict, design_system: dict, page_index: int, total_pages: int,
                           aspect_ratio: str) -> dict | None:
    """为单页内容设计布局和动画，并提供高度相关的案例供AI学习。"""
    canvas_width, canvas_height = (1024, 768) if aspect_ratio == "4:3" else (1280, 720)

    prompt = f"""
    你是一位精通**布局、动画和视觉传达**的演示文稿排版导演。
    你的任务是为**一页**幻灯片进行详细的视觉设计。

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
    2.  **布局多样性**: 这是第 {page_index + 1} / {total_pages} 页。请考虑上下文，选择一个合适的 `layout_type` (如 `title_slide`, `image_left_content_right`, `full_screen_image_with_quote`, `three_column_comparison`, `data_chart_summary`, `team_introduction`)。**避免与可能的上一页布局过于雷同。**
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
    ### **输出样例 (参考案例)**
    你必须学习以下成功案例的布局和动画编排手法。

    #### 案例1: 团队介绍页 (使用圆形裁剪、半透明背景和动画)
    {{
      "layout_type": "team_introduction",
      "elements": [
        {{ "id": "page_title", "type": "text_box", "x": 100, "y": 80, "width": 1080, "height": 60, "content": "核心团队成员", "style": {{ "font": {{ "type": "heading", "size": 36, "bold": true }}, "alignment": "CENTER" }}, "z_index": 20 }},
        {{ "id": "shape_bg", "type": "shape", "shape_type": "rectangle", "x": 0, "y": 550, "width": 1280, "height": 170, "style": {{ "fill_color": "#4A4A4A", "opacity": 0.1 }}, "z_index": 10 }},
        {{ "id": "member1_image", "type": "image", "x": 240, "y": 200, "width": 150, "height": 150, "image_keyword": "professional portrait of a smiling young woman", "style": {{ "crop": "circle" }}, "z_index": 15 }},
        {{ "id": "member1_text", "type": "text_box", "x": 215, "y": 360, "width": 200, "height": 60, "content": "张三\\n产品经理", "style": {{ "font": {{ "type": "body", "size": 16 }}, "alignment": "CENTER" }}, "z_index": 25 }}
      ],
      "animation_sequence": [
        {{ "element_id": "page_title", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "member1_image", "animation": {{ "type": "flyIn", "direction": "fromLeft" }} }}
      ]
    }}

    #### 案例2: 功能介绍页 (左图右文)
    {{
      "layout_type": "image_left_content_right",
      "elements": [
        {{ "id": "main_image", "type": "image", "x": 0, "y": 0, "width": 640, "height": 720, "image_keyword": "vibrant flat illustration of a mobile app interface", "z_index": 10 }},
        {{ "id": "title_text", "type": "text_box", "x": 700, "y": 150, "width": 520, "height": 80, "content": "产品核心功能", "style": {{ "font": {{ "type": "heading", "size": 40, "bold": true }} }}, "z_index": 20 }},
        {{ "id": "features_list", "type": "text_box", "x": 700, "y": 250, "width": 520, "height": 300, 
           "content": [
             "AI智能规划：一句话生成完整演示文稿。",
             "丰富的设计模板：覆盖多种行业和场景。"
           ],
           "style": {{ "font": {{ "type": "body", "size": 22 }}, "alignment": "LEFT" }}, "z_index": 20
        }}
      ],
      "animation_sequence": [
        {{ "element_id": "main_image", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "title_text", "animation": {{ "type": "flyIn", "direction": "fromTop" }} }}
      ]
    }}

    #### 案例3: 演示复杂动画序列和图层
     {{
      "layout_type": "content_reveal_with_layers",
      "elements": [
        {{ "id": "background_image", "type": "image", "x": 0, "y": 0, "width": 1280, "height": 720, "image_keyword": "modern city skyline at sunset", "z_index": 5, "style": {{ "opacity": 0.9 }} }},
        {{ "id": "overlay_shape", "type": "shape", "shape_type": "rounded_rectangle", "x": 100, "y": 120, "width": 500, "height": 450, "style": {{ "fill_color": "#000000", "opacity": 0.4 }}, "z_index": 15 }},
        {{ "id": "title", "type": "text_box", "x": 130, "y": 150, "width": 440, "height": 80, "content": "光影之城", "style": {{ "font": {{ "name": "微软雅黑", "size": 48, "bold": true, "color": "#FFFFFF" }}}}, "z_index": 25 }},
        {{ "id": "point_1", "type": "text_box", "x": 130, "y": 250, "width": 440, "height": 100, "content": "探索都市脉搏与自然光影的交织。", "style": {{ "font": {{ "size": 20, "color": "#FFFFFF" }}}}, "z_index": 25 }},
        {{ "id": "point_2", "type": "text_box", "x": 130, "y": 350, "width": 440, "height": 100, "content": "运用半透明背景块，突出文字信息。", "style": {{ "font": {{ "size": 20, "color": "#FFFFFF" }}}}, "z_index": 25 }}
      ],
      "animation_sequence": [
        {{ "element_id": "background_image", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "overlay_shape", "animation": {{ "type": "flyIn", "direction": "fromLeft" }} }},
        {{ "element_id": "title", "animation": {{ "type": "fadeIn" }} }},
        {{ "element_id": "point_1", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }},
        {{ "element_id": "point_2", "animation": {{ "type": "flyIn", "direction": "fromBottom" }} }}
      ]
     }}
    ---
    现在，请为当前页面设计布局和动画，并输出最终的JSON对象。
    """
    model = MODEL_CONFIG.get("planner", "glm-4")
    return _call_ai(prompt, model)


# --- [最终版] 主函数：编排“拥有案例手册的专家流水线” ---
def generate_presentation_pipeline(theme: str, num_pages: int, aspect_ratio: str = "16:9") -> dict | None:
    """
    通过分阶段、并发执行、带自我修正的AI专家流水线，高效生成完整的、高质量的演示文稿计划。
    """
    logging.info("--- 开始AI专家流水线生成任务 (带并发引擎和纠错层) ---")

    # 阶段一和阶段二保持不变...
    logging.info("[阶段 1/3] 正在由“视觉总监”生成核心设计系统...")
    design_system = _generate_design_system(theme)
    if not design_system:
        logging.error("设计系统生成失败，任务中止。")
        return None
    logging.info("核心设计系统已生成。")

    logging.info("[阶段 2/3] 正在由“信息架构师”生成内容大纲...")
    content_outline = _generate_content_outline(theme, num_pages)
    if not content_outline or 'pages' not in content_outline:
        logging.error("内容大纲生成失败，任务中止。")
        return None
    logging.info("内容大纲已生成。")

    # 阶段三: 并发处理
    logging.info("[阶段 3/3] “排版导演团队”开始并行设计布局与动画...")

    final_pages: List[Optional[Dict[str, Any]]] = [None] * len(content_outline['pages'])

    with ThreadPoolExecutor(max_workers=6) as executor:
        future_to_index = {
            executor.submit(
                _generate_slide_layout,
                page_content,
                design_system,
                i,
                len(content_outline['pages']),
                aspect_ratio
            ): i
            for i, page_content in enumerate(content_outline['pages'])
        }

        for future in as_completed(future_to_index):
            index = future_to_index[future]
            try:
                slide_layout = future.result()
                if slide_layout and "elements" in slide_layout:

                    # --- [核心修正] AI纠错层：清理无效动画 ---
                    # 1. 获取本页所有真实存在的元素ID
                    existing_element_ids = {
                        element['id'] for element in slide_layout['elements'] if 'id' in element
                    }

                    # 2. 过滤动画序列，只保留引用了真实ID的动画
                    if "animation_sequence" in slide_layout:
                        valid_animations = [
                            anim for anim in slide_layout['animation_sequence']
                            if anim.get('element_id') in existing_element_ids
                        ]
                        # 如果有变化，则记录日志
                        if len(valid_animations) < len(slide_layout['animation_sequence']):
                            logging.warning(f"  - 第 {index + 1} 页：已自动清理无效的动画引用。")
                        slide_layout['animation_sequence'] = valid_animations
                    # --- AI纠错层结束 ---

                    logging.info(f"  - 第 {index + 1} 页设计完成。")
                    final_pages[index] = slide_layout
                else:
                    raise ValueError("布局生成失败或未包含 'elements'。")
            except Exception as exc:
                logging.error(f"  - 第 {index + 1} 页设计中发生错误: {exc}")
                final_pages[index] = {
                    "layout_type": "fallback_error",
                    "elements": [{"type": "text_box", "content": f"页面 {index+1} 生成失败", "x": 100, "y": 100, "width": 1080, "height": 100, "style": {"font": {"type": "heading", "size": 36}}}]
                }

    # 组装最终计划
    if all(p is None for p in final_pages):
        logging.error("所有页面布局均生成失败，任务中止。")
        return None

    final_plan = {
        **design_system,
        "master_slide": {
            "background": { "color": design_system.get("color_palette", {}).get("background", "#FFFFFF") }
        },
        "pages": [p for p in final_pages if p] # 过滤掉None值
    }

    logging.info("--- AI专家流水线任务全部完成 ---")
    return final_plan

# --- 无缝切换接口 ---
generate_presentation_plan = generate_presentation_pipeline
