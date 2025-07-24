import logging
import re

from PIL import Image
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
def add_text_box(slide, element_data: dict, style_manager: PresentationStyle):
    """
    [已更新] 添加文本框，实现灵活的字体控制。
    - 优先使用元素自身指定的字体 (`font.name`)。
    - 如果未指定，则根据 `font.type` ('heading'/'body') 回退到全局默认字体。
    - 兼容处理Markdown风格的加粗。
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

        p = tf.paragraphs[0]
        content = element_data.get('content', '')
        style = element_data.get('style', {})
        font_style = style.get('font', {})

        # --- [核心修改] 字体决策逻辑 ---
        # 1. 优先从元素 style 中获取 `font.name`
        font_name = font_style.get('name')
        if not font_name:
            # 2. 如果没有，则根据 `font.type` 回退到 style_manager 中的默认字体
            font_type = font_style.get('type', 'body')  # 默认为 'body'
            font_name = style_manager.heading_font if font_type == 'heading' else style_manager.body_font
        # --- [核心修改结束] ---

        # 优先使用局部颜色，否则使用全局文本颜色
        default_font_color = hex_to_rgb(font_style['color']) if 'color' in font_style else style_manager.text_color

        # 处理Markdown加粗 `**text**`
        if '**' in content:
            parts = re.split(r'\*\*(.*?)\*\*', content)
            for i, part in enumerate(parts):
                if not part: continue

                run = p.add_run()
                run.text = part
                font = run.font

                font.name = font_name  # 应用决策后的字体
                font.size = Pt(font_style.get('size', 18))
                font.italic = font_style.get('italic', False)
                font.color.rgb = default_font_color

                is_bold_from_json = font_style.get('bold', False)
                is_bold_from_markdown = (i % 2 == 1)
                font.bold = is_bold_from_json or is_bold_from_markdown
        else:
            # 标准文本处理
            run = p.add_run()
            run.text = content
            font = run.font
            font.name = font_name  # 应用决策后的字体
            font.size = Pt(font_style.get('size', 18))
            font.bold = font_style.get('bold', False)
            font.italic = font_style.get('italic', False)
            font.color.rgb = default_font_color

        # 设置段落对齐
        if alignment_str := style.get('alignment'):
            p.alignment = ALIGNMENT_MAP.get(alignment_str.upper(), PP_ALIGN.LEFT)
        else:
            p.alignment = PP_ALIGN.LEFT

        logging.info(f"添加文本框 (字体: {font_name}): '{content[:30]}...'")
    except Exception as e:
        logging.error(f"添加文本框时出错: {e} | 原始元素数据: {element_data}", exc_info=True)


def add_image(slide, image_path: str, element_data: dict):
    """
    [已更新] 添加图片并智能裁剪以适应图框，避免拉伸。
    该函数会保持图片的原始宽高比，通过裁剪多余部分来填充指定的图框区域。
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

        # 初始时，将图片添加到幻灯片上，尺寸与图框相同（这会导致暂时拉伸）
        pic = slide.shapes.add_picture(image_path, box_x_emu, box_y_emu, width=box_width_emu, height=box_height_emu)

        # --- 核心裁剪逻辑 ---
        # 读取图片原始尺寸
        with Image.open(image_path) as img:
            img_width_px, img_height_px = img.size

        # 计算宽高比
        # 避免除以零的错误
        if img_height_px == 0 or box_height == 0:
            logging.warning(f"图片或图框高度为零，无法计算宽高比，跳过裁剪: {image_path}")
            return

        img_aspect = img_width_px / img_height_px
        box_aspect = box_width / box_height

        # 比较宽高比，决定如何裁剪
        # 使用 round 避免浮点数精度问题
        if round(img_aspect, 2) != round(box_aspect, 2):
            if img_aspect > box_aspect:
                # 图片比图框更“宽”，需要裁剪左右两边
                # 新的裁剪比例计算公式: (1 - 目标宽高比 / 图片宽高比) / 2
                crop_ratio = (1 - box_aspect / img_aspect) / 2
                pic.crop_left = crop_ratio
                pic.crop_right = crop_ratio
                pic.crop_top = 0
                pic.crop_bottom = 0
            else:
                # 图片比图框更“高”，需要裁剪上下两边
                crop_ratio = (1 - img_aspect / box_aspect) / 2
                pic.crop_top = crop_ratio
                pic.crop_bottom = crop_ratio
                pic.crop_left = 0
                pic.crop_right = 0

        logging.info(f"从路径添加并智能裁剪图片: {image_path}")

    except FileNotFoundError:
        logging.error(f"图片文件未找到: {image_path}")
    except Exception as e:
        logging.error(f"添加或裁剪图片 {image_path} 时出错: {e}", exc_info=True)



def add_shape(slide, element_data: dict, style_manager: PresentationStyle):
    """
    [已更新] 添加形状并应用样式，增加对无效颜色值的容错处理。
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

        # 填充样式
        if 'gradient' in style:
            grad_info = style['gradient']
            fill.gradient()
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

        # 边框样式
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

def add_chart(slide, element_data: dict, style_manager: PresentationStyle):
    """添加图表并使用主题颜色进行样式化。"""
    try:
        x, y, width, height = map(px_to_emu, [
            element_data.get('x', 100), element_data.get('y', 150),
            element_data.get('width', 1080), element_data.get('height', 450)
        ])

        chart_type_str = element_data.get('chart_type', 'bar').upper()
        chart_type_map = {'BAR': XL_CHART_TYPE.COLUMN_CLUSTERED, 'PIE': XL_CHART_TYPE.PIE, 'LINE': XL_CHART_TYPE.LINE}
        chart_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

        chart_data_info = element_data.get('data', {})
        chart_data = ChartData()
        chart_data.categories = chart_data_info.get('categories', [])
        for series_data in chart_data_info.get('series', []):
            chart_data.add_series(series_data.get('name', ''), series_data.get('values', []))

        graphic_frame = slide.shapes.add_chart(chart_type, x, y, width, height, chart_data)
        chart = graphic_frame.chart

        # **[关键优化]** 使用样式管理器为图表系列应用主题颜色
        if hasattr(chart.plots[0], 'series'):
            for i, series in enumerate(chart.plots[0].series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = style_manager.get_chart_color(i)

        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels
        data_labels.font.size = Pt(10)
        data_labels.font.color.rgb = style_manager.text_color
        logging.info(f"添加已应用主题样式的 {chart_type_str} 图表。")
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

        # 设置列宽和行高 (可以根据需要进行更复杂的计算)
        for c in range(num_cols):
            table.columns[c].width = int(width / num_cols)
        for r in range(num_rows):
            table.rows[r].height = int(height / num_rows)

        style = element_data.get('style', {})
        header_color = hex_to_rgb(style.get('header_color')) if 'header_color' in style else style_manager.primary
        row_colors = [hex_to_rgb(c) for c in style.get('row_colors', [])]

        # 填充表头
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_color
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)  # 白色文字
            p.font.bold = True
            p.font.name = style_manager.heading_font

        # 填充数据行
        for r, row_data in enumerate(rows_data):
            for c, cell_data in enumerate(row_data):
                cell = table.cell(r + 1, c)
                cell.text = str(cell_data)
                # 应用斑马条纹
                if row_colors:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = row_colors[r % len(row_colors)]
                p = cell.text_frame.paragraphs[0]
                p.font.color.rgb = style_manager.text_color
                p.font.name = style_manager.body_font

        logging.info(f"添加了包含 {len(rows_data)} 行的表格。")
    except Exception as e:
        logging.error(f"添加表格时出错: {e}", exc_info=True)