# ai_service.py
import json
import logging
import re
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
            temperature=0.55
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
    你是一位深谙**年轻女性审美**的顶级演示文稿（PPT）设计大师。
    你的任务是为主题 **“{theme}”** 设计一个全局视觉系统。
    你的输出必须是一个单一、完整、严格符合格式的原始JSON对象。

    ### 设计风格指南
    (此处省略了完整的风格指南，以保持代码简洁，但实际应保留上一版本中的所有详细规则)
    ...

    ### **输出样例 (参考案例)**
    你必须学习以下成功案例的风格和格式：

    #### 案例1: "盐系手帐"
    {{
      "design_concept": "盐系手帐：我们的故事",
      "font_pairing": {{ "heading": "Dengxian", "body": "Dengxian" }},
      "color_palette": {{ "primary": "#4A4A4A", "secondary": "#9B9B9B", "background": "#FDFBF8", "text": "#4A4A4A", "accent": "#F5A623" }}
    }}

    #### 案例2: "夏日橘子汽水"
    {{
      "design_concept": "夏日橘子汽水",
      "font_pairing": {{ "heading": "Microsoft YaHei Light", "body": "Microsoft YaHei Light" }},
      "color_palette": {{ "primary": "#FF6B6B", "secondary": "#FFD166", "background": "#FFFFFF", "text": "#4A4A4A", "accent": "#06D6A0" }}
    }}

    #### 案例3: "赛博蝴蝶梦"
    {{
        "design_concept": "赛博蝴蝶梦",
        "font_pairing": {{ "heading": "Dengxian", "body": "Dengxian" }},
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

    ### 核心准则
    (此处省略了完整的核心准则，以保持代码简洁，但实际应保留上一版本中的所有详细规则)
    ...

    ### **输出样例 (参考案例)**
    你必须学习以下成功案例的格式，为每一页生成清晰的`title`和`content`。

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
    {json.dumps(design_system, ensure_ascii=False)}

    **当前页面内容 (需要你来设计布局):**
    {json.dumps(page_content, ensure_ascii=False)}

    ### 核心设计规则
    (此处省略了完整的核心设计规则，以保持代码简洁，但实际应保留上一版本中的所有详细规则)
    ...

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
    通过分阶段的、配备了专属案例的AI专家流水线，生成完整的、高质量的演示文稿计划。
    """
    # ... (此主函数的逻辑保持不变，它负责调用上面的三个新函数) ...
    logging.info("--- 开始AI专家流水线生成任务 ---")

    # 阶段一：AI视觉总监确定风格
    logging.info("[阶段 1/3] 正在由“视觉总监”生成核心设计系统...")
    design_system = _generate_design_system(theme)
    if not design_system:
        logging.error("设计系统生成失败，任务中止。")
        return None
    logging.info("核心设计系统已生成。")

    # 阶段二：AI信息架构师构建大纲
    logging.info("[阶段 2/3] 正在由“信息架构师”生成内容大纲...")
    # 注意：这里需要一个逻辑来合并从AI获取的title和content
    content_outline = _generate_content_outline(theme, num_pages)
    if not content_outline or 'pages' not in content_outline:
        logging.error("内容大纲生成失败，任务中止。")
        return None
    logging.info("内容大纲已生成。")

    # 阶段三：AI排版导演逐页设计
    logging.info("[阶段 3/3] 正在由“排版导演”逐页设计布局与动画...")
    final_pages = []
    total_pages = len(content_outline['pages'])
    for i, page_content in enumerate(content_outline['pages']):
        logging.info(f"  - 正在设计第 {i + 1}/{total_pages} 页...")
        slide_layout = _generate_slide_layout(page_content, design_system, i, total_pages, aspect_ratio)
        if slide_layout:
            page_data = {
                "layout_type": slide_layout.get("layout_type", "default"),
                "elements": slide_layout.get("elements", []),
                "animation_sequence": slide_layout.get("animation_sequence", [])
            }
            final_pages.append(page_data)
        else:
            logging.warning(f"第 {i + 1} 页布局生成失败，已跳过。")
            final_pages.append({
                "layout_type": "fallback_simple",
                "elements": [
                    {"type": "text_box", "content": page_content.get("title", "此页生成失败"), "x": 100, "y": 100,
                     "width": 1080, "height": 100, "style": {"font": {"type": "heading", "size": 36}}}]
            })

    if not final_pages:
        logging.error("所有页面布局均生成失败，任务中止。")
        return None

    # 组装最终的完整计划
    final_plan = {
        **design_system,
        "master_slide": {
            "background": {"color": design_system.get("color_palette", {}).get("background", "#FFFFFF")}
        },
        "pages": final_pages
    }

    logging.info("--- AI专家流水线任务全部完成 ---")
    return final_plan


# --- 无缝切换接口 ---
generate_presentation_plan = generate_presentation_pipeline