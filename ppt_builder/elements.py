import logging
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.table import Table
from ppt_builder.styles import px_to_emu, hex_to_rgb, PresentationStyle

# 形状类型映射
shape_type_map = {
    'rectangle': MSO_SHAPE.RECTANGLE,
    'oval': MSO_SHAPE.OVAL,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'star': MSO_SHAPE.STAR_5_POINT,
    'rounded_rectangle': MSO_SHAPE.ROUNDED_RECTANGLE,
}


def add_text_box(slide, element_data: dict, style_manager: PresentationStyle):
    x, y, width, height = map(px_to_emu,
                              [element_data['x'], element_data['y'], element_data['width'], element_data['height']])
    style = element_data.get('style', {})

    try:
        txBox = slide.shapes.add_textbox(x, y, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = element_data.get('content', '')

        # 应用样式
        font = p.font
        font_style = style.get('font', {})
        font.name = style_manager.body_font if font_style.get('type') == 'body' else style_manager.heading_font
        font.size = Pt(font_style.get('size', 18))
        font.bold = font_style.get('bold', False)
        font.italic = font_style.get('italic', False)
        font.color.rgb = hex_to_rgb(
            font_style.get('color', '#000000')) if 'color' in font_style else style_manager.text_color

        logging.info(f"Added text box: '{p.text[:20]}...'")
    except Exception as e:
        logging.error(f"Error adding text box: {e}", exc_info=True)


def add_image(slide, image_path: str, element_data: dict):
    x, y, width, height = map(px_to_emu,
                              [element_data['x'], element_data['y'], element_data['width'], element_data['height']])
    try:
        slide.shapes.add_picture(image_path, x, y, width, height)
        logging.info(f"Added image from path: {image_path}")
    except Exception as e:
        logging.error(f"Error adding image {image_path}: {e}", exc_info=True)


def add_shape(slide, element_data: dict, style_manager: PresentationStyle):
    x, y, width, height = map(px_to_emu,
                              [element_data['x'], element_data['y'], element_data['width'], element_data['height']])
    shape_type_str = element_data.get('shape_type', 'rectangle').lower()
    shape_type = shape_type_map.get(shape_type_str, MSO_SHAPE.RECTANGLE)
    style = element_data.get('style', {})

    try:
        shape = slide.shapes.add_shape(shape_type, x, y, width, height)
        fill = shape.fill
        line = shape.line

        if 'gradient' in style:
            grad_info = style['gradient']
            fill.gradient()
            for i, hex_color in enumerate(grad_info.get('colors', [])):
                if i < len(fill.gradient_stops):
                    fill.gradient_stops[i].color.rgb = hex_to_rgb(hex_color)
            logging.info("Applied gradient fill to shape.")
        elif 'fill_color' in style and style['fill_color'] is not None:
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(style['fill_color'])
        else:
            fill.background()

        if 'border_color' in style and style['border_color'] is not None:
            line.color.rgb = hex_to_rgb(style['border_color'])
            line.width = Pt(style.get('border_width', 1))
        else:
            line.fill.background()

        logging.info(f"Added {shape_type_str} shape.")
    except Exception as e:
        logging.error(f"Error adding shape: {e}", exc_info=True)


def add_chart(slide, element_data: dict, style_manager: PresentationStyle):
    x, y, width, height = map(px_to_emu,
                              [element_data['x'], element_data['y'], element_data['width'], element_data['height']])
    chart_type_str = element_data.get('chart_type', 'bar').upper()
    chart_type_map = {'BAR': XL_CHART_TYPE.BAR_CLUSTERED, 'PIE': XL_CHART_TYPE.PIE, 'LINE': XL_CHART_TYPE.LINE}
    chart_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.BAR_CLUSTERED)

    try:
        chart_data_info = element_data.get('data', {})
        chart_data = ChartData()
        chart_data.categories = chart_data_info.get('categories', [])
        for series_data in chart_data_info.get('series', []):
            chart_data.add_series(series_data.get('name', ''), series_data.get('values', []))

        graphic_frame = slide.shapes.add_chart(chart_type, x, y, width, height, chart_data)
        chart = graphic_frame.chart

        plot = chart.plots[0]
        if hasattr(plot, 'series'):
            for i, series in enumerate(plot.series):
                color = style_manager.get_chart_color(i)
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = color

        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

        logging.info(f"Added styled {chart_type_str} chart.")
    except Exception as e:
        logging.error(f"Error adding chart: {e}", exc_info=True)


def add_table(slide, element_data: dict, style_manager: PresentationStyle):
    # [关键修复] 使用 .get() 提供默认值，防止因AI未提供坐标而产生KeyError
    x_px = element_data.get('x', 100)
    y_px = element_data.get('y', 150)
    width_px = element_data.get('width', 1080)
    height_px = element_data.get('height', 420)

    if 'x' not in element_data:
        logging.warning("Table element missing coordinates. Using default position and size.")

    x, y, width, height = map(px_to_emu, [x_px, y_px, width_px, height_px])

    headers = element_data.get('headers', [])
    rows_data = element_data.get('rows', [])

    if not headers or not rows_data:
        logging.warning("Table data missing headers or rows. Skipping.")
        return

    num_rows = len(rows_data) + 1
    num_cols = len(headers)

    try:
        shape = slide.shapes.add_table(num_rows, num_cols, x, y, width, height)
        table = shape.table

        table.columns[0].width = width
        for c in range(num_cols):
            table.columns[c].width = int(width / num_cols)
        for r in range(num_rows):
            table.rows[r].height = int(height / num_rows)

        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = style_manager.primary
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True
            p.font.name = style_manager.heading_font

        for r, row_data in enumerate(rows_data):
            for c, cell_data in enumerate(row_data):
                cell = table.cell(r + 1, c)
                cell.text = str(cell_data)
                p = cell.text_frame.paragraphs[0]
                p.font.color.rgb = style_manager.text_color
                p.font.name = style_manager.body_font

        logging.info(f"Added table with {num_rows - 1} rows.")
    except Exception as e:
        logging.error(f"Error adding table: {e}", exc_info=True)
