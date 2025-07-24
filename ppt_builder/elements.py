import logging
import re
import io

from PIL import Image, ImageOps, ImageDraw
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
# **[修复]** 导入正确的段落对齐枚举
from pptx.enum.text import PP_ALIGN
from ppt_builder.styles import px_to_emu, hex_to_rgb, PresentationStyle

# 形状类型映射
SHAPE_TYPE_MAP = {
    'rectangle': MSO_SHAPE.RECTANGLE,
    'oval': MSO_SHAPE.OVAL,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'star': MSO_SHAPE.STAR_5_POINT,
    'rounded_rectangle': MSO_SHAPE.ROUNDED_RECTANGLE,
}


# **[修复]** 创建一个健壮的对齐方式映射
ALIGNMENT_MAP = {
    'LEFT': PP_ALIGN.LEFT,
    'CENTER': PP_ALIGN.CENTER,
    'RIGHT': PP_ALIGN.RIGHT,
    'JUSTIFY': PP_ALIGN.JUSTIFY,
}


def _crop_to_circle(image_path: str):
    """
    使用Pillow将指定路径的图片处理成圆形，并返回一个内存中的图片流。
    """
    try:
        # 打开图片并转换为RGBA以支持透明度
        with Image.open(image_path) as img:
            img = img.convert("RGBA")

            # 创建一个与原图短边等大的正方形画布
            size = (min(img.size), min(img.size))

            # 创建一个黑色的圆形遮罩 (mask)
            mask = Image.new('L', size, 0)
            draw = ImageDraw.Draw(mask)
            draw.ellipse((0, 0) + size, fill=255)

            # 使用ImageOps.fit将原图无拉伸地裁剪并缩放到正方形
            # 这会居中裁剪，保留最重要的部分
            output = ImageOps.fit(img, mask.size, centering=(0.5, 0.5))

            # 应用圆形遮罩
            output.putalpha(mask)

            # 将处理后的图片保存到内存中的字节流中
            buffer = io.BytesIO()
            output.save(buffer, format='PNG')
            buffer.seek(0)
            return buffer

    except Exception as e:
        logging.error(f"处理图片为圆形时失败: {image_path} - {e}", exc_info=True)
        return None

def _set_shape_transparency(shape, transparency: int):
    """
    设置形状的透明度。
    :param shape: pptx 形状对象
    :param transparency: 透明度值（0-100000）
    """
    try:
        fill = shape.fill
        if fill.type == 1:  # 填充类型为纯色
            fill.fore_color._xColorVal.transparency = transparency
        elif fill.type == 2:  # 填充类型为渐变
            for stop in fill.gradient_stops:
                stop.color._xColorVal.transparency = transparency
    except Exception as e:
        logging.warning(f"设置形状透明度时出错: {e}")
def add_text_box(slide, element_data: dict, style_manager: PresentationStyle):
    """
    [已更新] 添加文本框，实现灵活的字体控制和项目符号列表。
    - 优先使用元素自身指定的字体 (`font.name`)。
    - 如果未指定，则根据 `font.type` ('heading'/'body') 回退到全局默认字体。
    - 兼容处理Markdown风格的加粗。
    - **[新功能]** 如果`content`是列表，则自动生成项目符号列表。
    """
    try:
        x, y, width, height = map(px_to_emu, [
            element_data.get('x', 50), element_data.get('y', 50),
            element_data.get('width', 1180), element_data.get('height', 100)
        ])

        txBox = slide.shapes.add_textbox(x, y, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.clear()

        content = element_data.get('content', '')
        style = element_data.get('style', {})
        font_style = style.get('font', {})

        # --- 字体决策逻辑 ---
        font_name = font_style.get('name')
        if not font_name:
            font_type = font_style.get('type', 'body')
            font_name = style_manager.heading_font if font_type == 'heading' else style_manager.body_font

        default_font_color = hex_to_rgb(font_style['color']) if 'color' in font_style else style_manager.text_color
        font_size_pt = font_style.get('size', 18)
        is_italic_from_json = font_style.get('italic', False)
        is_bold_from_json = font_style.get('bold', False)

        def apply_run_style(run, text_part, is_markdown_bold=False):
            """辅助函数，用于应用样式到文本块"""
            run.text = text_part
            font = run.font
            font.name = font_name
            font.size = Pt(font_size_pt)
            font.italic = is_italic_from_json
            font.color.rgb = default_font_color
            font.bold = is_bold_from_json or is_markdown_bold

        # ================== 新增逻辑：项目符号列表 ==================
        if isinstance(content, list):
            logging.info(f"检测到项目符号列表，共 {len(content)} 项。")
            for i, item_text in enumerate(content):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()

                p.level = 0  # 可根据需要设置缩进级别

                if '**' in item_text:
                    parts = re.split(r'\*\*(.*?)\*\*', item_text)
                    for j, part in enumerate(parts):
                        if not part: continue
                        apply_run_style(p.add_run(), part, is_markdown_bold=(j % 2 == 1))
                else:
                    apply_run_style(p.add_run(), item_text)

                # 为每个段落设置对齐方式
                if alignment_str := style.get('alignment'):
                    p.alignment = ALIGNMENT_MAP.get(alignment_str.upper(), PP_ALIGN.LEFT)
                else:
                    p.alignment = PP_ALIGN.LEFT
        # ================== 原有逻辑：处理单个字符串 ==================
        else:
            p = tf.paragraphs[0]
            if '**' in content:
                parts = re.split(r'\*\*(.*?)\*\*', content)
                for i, part in enumerate(parts):
                    if not part: continue
                    apply_run_style(p.add_run(), part, is_markdown_bold=(i % 2 == 1))
            else:
                apply_run_style(p.add_run(), content)

            # 设置段落对齐
            if alignment_str := style.get('alignment'):
                p.alignment = ALIGNMENT_MAP.get(alignment_str.upper(), PP_ALIGN.LEFT)
            else:
                p.alignment = PP_ALIGN.LEFT

        log_content = str(content)
        logging.info(f"添加文本框 (字体: {font_name}): '{log_content[:30]}...'")
    except Exception as e:
        logging.error(f"添加文本框时出错: {e} | 原始元素数据: {element_data}", exc_info=True)


def add_image(slide, image_path: str, element_data: dict):
    """
    [最终版] 向幻灯片添加图片。
    - 如果 style.crop 为 'circle', 则将图片裁剪为圆形。
    - 否则，执行智能矩形裁剪以适应图框，避免拉伸。
    """
    try:
        if not image_path:
            logging.warning("图片路径为空，跳过添加图片。")
            return

        box_x, box_y, box_width, box_height = [
            element_data.get('x', 0), element_data.get('y', 0),
            element_data.get('width', 1280), element_data.get('height', 720)
        ]
        box_x_emu, box_y_emu, box_width_emu, box_height_emu = map(px_to_emu, [box_x, box_y, box_width, box_height])

        style = element_data.get('style', {})

        if style.get('crop') == 'circle':
            logging.info(f"检测到圆形裁剪请求，正在处理图片: {image_path}")
            circular_image_stream = _crop_to_circle(image_path)
            if circular_image_stream:
                diameter = min(box_width_emu, box_height_emu)
                slide.shapes.add_picture(circular_image_stream, box_x_emu, box_y_emu, width=diameter, height=diameter)
                logging.info(f"成功添加圆形图片: {image_path}")
            else:
                logging.error(f"圆形图片处理失败，无法添加图片: {image_path}")
        else:
            pic = slide.shapes.add_picture(image_path, box_x_emu, box_y_emu, width=box_width_emu, height=box_height_emu)
            with Image.open(image_path) as img:
                img_width_px, img_height_px = img.size
            if img_height_px == 0 or box_height == 0:
                logging.warning(f"图片或图框高度为零，无法计算宽高比，跳过裁剪: {image_path}")
                return
            img_aspect = img_width_px / img_height_px
            box_aspect = box_width / box_height
            if round(img_aspect, 2) != round(box_aspect, 2):
                if img_aspect > box_aspect:
                    crop_ratio = (1 - box_aspect / img_aspect) / 2
                    pic.crop_left = crop_ratio
                    pic.crop_right = crop_ratio
                else:
                    crop_ratio = (1 - img_aspect / box_aspect) / 2
                    pic.crop_top = crop_ratio
                    pic.crop_bottom = crop_ratio
            logging.info(f"从路径添加并智能裁剪矩形图片: {image_path}")

    except FileNotFoundError:
        logging.error(f"图片文件未找到: {image_path}")
    except Exception as e:
        logging.error(f"添加或裁剪图片 {image_path} 时出错: {e}", exc_info=True)


def add_shape(slide, element_data: dict, style_manager: PresentationStyle):
    """
    [已修复] 添加形状并应用样式。
    - 统一处理透明度，确保其对纯色和渐变填充都生效。
    - 增加对无效颜色值的容错处理。
    """
    try:
        x, y, width, height = map(px_to_emu, [
            element_data.get('x', 50), element_data.get('y', 50),
            element_data.get('width', 200), element_data.get('height', 200)
        ])
        shape_type_str = element_data.get('shape_type', 'rectangle').lower()
        shape_type = SHAPE_TYPE_MAP.get(shape_type_str, MSO_SHAPE.RECTANGLE)
        style = element_data.get('style', {})

        shape = slide.shapes.add_shape(shape_type, x, y, width, height)
        fill = shape.fill
        line = shape.line

        # --- 第一步：设置填充类型和颜色 ---
        # 这个逻辑块决定了形状是渐变、纯色还是无填充
        if 'gradient' in style:
            grad_info = style['gradient']
            fill.gradient()
            # 检查并应用渐变角度 (如果提供)
            # 注意: python-pptx 似乎主要支持线性渐变的角度
            if 'angle' in grad_info:
                fill.gradient_angle = grad_info['angle']

            for i, hex_color in enumerate(grad_info.get('colors', [])):
                if i < len(fill.gradient_stops):
                    try:
                        fill.gradient_stops[i].color.rgb = hex_to_rgb(hex_color)
                    except (ValueError, IndexError):
                        logging.warning(f"形状渐变中提供了无效的十六进制颜色 '{hex_color}'。将使用默认黑色。")
                        fill.gradient_stops[i].color.rgb = RGBColor(0, 0, 0)
            logging.info("为形状应用了渐变填充。")

        elif 'fill_color' in style and style['fill_color'] is not None:
            fill.solid()
            try:
                fill.fore_color.rgb = hex_to_rgb(style['fill_color'])
            except (ValueError, IndexError):
                logging.warning(f"形状填充中提供了无效的十六进制颜色 '{style['fill_color']}'。将使用默认黑色。")
                fill.fore_color.rgb = RGBColor(0, 0, 0)
        else:
            fill.background()  # 无填充

        # --- 第二步：统一应用透明度 ---
        # [核心修复] 将透明度逻辑移到填充类型设置之后，使其对所有填充都生效
        opacity = style.get('opacity')
        if isinstance(opacity, (float, int)) and 0 <= opacity <= 1:
            transparency = 1 - opacity
            _set_shape_transparency(shape, 100000 * transparency)
            logging.info(f"为形状应用了透明度: {transparency}")

        # --- 第三步：设置边框样式 ---
        if border_style := style.get('border'):
            try:
                line_color_hex = border_style.get('color', '#000000')
                line.color.rgb = hex_to_rgb(line_color_hex)
                line.width = Pt(border_style.get('width', 1))
            except (ValueError, IndexError):
                logging.warning(f"形状边框中提供了无效的十六进制颜色 '{line_color_hex}'。将使用默认黑色边框。")
                line.color.rgb = RGBColor(0, 0, 0)
                line.width = Pt(border_style.get('width', 1))
        else:
            line.fill.background()  # 无边框

        logging.info(f"添加 {shape_type_str} 形状。")
    except Exception as e:
        logging.error(f"添加形状时发生意外错误: {e}", exc_info=True)


# 在 elements.py 文件中，请用下面的函数替换旧的 add_chart 函数

def add_chart(slide, element_data: dict, style_manager: PresentationStyle):
    """
    [再次优化] 添加图表并进行深度样式化，以实现最佳可读性。
    - 优化数据标签，使其更清晰。
    """
    try:
        x, y, width, height = map(px_to_emu, [
            element_data.get('x', 100), element_data.get('y', 150),
            element_data.get('width', 1080), element_data.get('height', 450)
        ])

        chart_type_str = element_data.get('chart_type', 'bar').upper()
        chart_type_map = {
            'BAR': XL_CHART_TYPE.COLUMN_CLUSTERED,
            'PIE': XL_CHART_TYPE.PIE,
            'LINE': XL_CHART_TYPE.LINE
        }
        chart_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

        chart_data_info = element_data.get('data', {})
        chart_data = ChartData()
        chart_data.categories = chart_data_info.get('categories', [])
        for series_data in chart_data_info.get('series', []):
            chart_data.add_series(series_data.get('name', ''), series_data.get('values', []))

        graphic_frame = slide.shapes.add_chart(chart_type, x, y, width, height, chart_data)
        chart = graphic_frame.chart

        # --- 1. 添加和设置图表标题 ---
        if chart_title_text := element_data.get('title'):
            chart.has_title = True
            chart_title = chart.chart_title
            chart_title.text_frame.text = chart_title_text
            p = chart_title.text_frame.paragraphs[0]
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = style_manager.text_color
            p.font.name = style_manager.heading_font
        else:
            chart.has_title = False

        # --- 2. 美化图例 (Legend) ---
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(12)
            chart.legend.font.color.rgb = style_manager.text_color
            chart.legend.font.name = style_manager.body_font

        # --- 3. 设置数据标签 ---
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels
        data_labels.font.size = Pt(14)  # 稍微增大字号
        data_labels.font.bold = True  # 加粗更清晰
        # 对于饼图，将标签颜色设为白色通常效果更好
        if chart_type == XL_CHART_TYPE.PIE:
            data_labels.font.color.rgb = RGBColor(255, 255, 255)

        # --- 4. 为系列应用主题色 ---
        if hasattr(plot, 'series'):
            for i, series in enumerate(plot.series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = style_manager.get_chart_color(i)
                if chart_type == XL_CHART_TYPE.LINE:
                    series.format.line.color.rgb = style_manager.get_chart_color(i)

        # --- 5. 美化坐标轴 (如果存在) ---
        if chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.LINE]:
            category_axis = chart.category_axis
            category_axis.tick_labels.font.size = Pt(12)
            category_axis.tick_labels.font.color.rgb = style_manager.text_color
            category_axis.tick_labels.font.name = style_manager.body_font

            value_axis = chart.value_axis
            if value_axis.has_major_gridlines:
                value_axis.major_gridlines.format.line.fill.solid()
                value_axis.major_gridlines.format.line.color.rgb = hex_to_rgb("#E0E0E0")
            value_axis.tick_labels.font.size = Pt(12)
            value_axis.tick_labels.font.color.rgb = style_manager.text_color
            value_axis.tick_labels.font.name = style_manager.body_font

        logging.info(f"添加并深度美化了 '{element_data.get('title', '无标题')}' 图表。")

    except Exception as e:
        logging.error(f"添加图表时出错: {e}", exc_info=True)
def add_table(slide, element_data: dict, style_manager: PresentationStyle):
    """添加表格并应用样式。"""
    try:
        x, y, width, height = map(px_to_emu, [
            element_data.get('x', 100), element_data.get('y', 150),
            element_data.get('width', 1080), element_data.get('height', 420)
        ])

        headers = element_data.get('headers', [])
        rows_data = element_data.get('rows', [])
        if not headers or not rows_data:
            logging.warning("表格数据缺少表头或行数据，已跳过。")
            return

        num_rows, num_cols = len(rows_data) + 1, len(headers)
        shape = slide.shapes.add_table(num_rows, num_cols, x, y, width, height)
        table = shape.table

        for c in range(num_cols):
            table.columns[c].width = int(width / num_cols)
        for r in range(num_rows):
            table.rows[r].height = int(height / num_rows)

        style = element_data.get('style', {})
        header_color = hex_to_rgb(style.get('header_color')) if 'header_color' in style else style_manager.primary
        row_colors = [hex_to_rgb(c) for c in style.get('row_colors', [])]

        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_color
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True
            p.font.name = style_manager.heading_font

        for r, row_data in enumerate(rows_data):
            for c, cell_data in enumerate(row_data):
                cell = table.cell(r + 1, c)
                cell.text = str(cell_data)
                if row_colors:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = row_colors[r % len(row_colors)]
                p = cell.text_frame.paragraphs[0]
                p.font.color.rgb = style_manager.text_color
                p.font.name = style_manager.body_font

        logging.info(f"添加了包含 {len(rows_data)} 行的表格。")
    except Exception as e:
        logging.error(f"添加表格时出错: {e}", exc_info=True)