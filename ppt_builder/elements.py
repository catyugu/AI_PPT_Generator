# ppt_builder/elements.py
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData, ChartData

import logging

# Standard slide dimensions in Emu for conversion from pixels (914x685 pixels at 96 DPI)
# 1 inch = 914400 Emu. Standard slide is 10 inches x 7.5 inches.
# 1 pixel = 9525 Emu (914400 Emu / 96 DPI)
PIXEL_TO_EMU = 9525


def px_to_emu(px: int) -> Emu:
    """Converts pixels to Emu."""
    return Emu(px * PIXEL_TO_EMU)


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Converts a hex color string (e.g., #RRGGBB) to RGBColor."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def add_text_box(slide, element_data: dict):
    """Adds a text box to the slide based on element_data."""
    x = px_to_emu(element_data['x'])
    y = px_to_emu(element_data['y'])
    width = px_to_emu(element_data['width'])
    height = px_to_emu(element_data['height'])
    content = element_data.get('content', '')
    style = element_data.get('style', {})

    try:
        txBox = slide.shapes.add_textbox(x, y, width, height)
        tf = txBox.text_frame
        tf.text = content
        tf.word_wrap = True  # Enable word wrap by default

        # Apply text box specific styles
        p = tf.paragraphs[0]
        font = p.font
        font.size = Pt(style.get('font_size', 18))
        font.name = style.get('font_name', 'Microsoft YaHei')
        font.bold = style.get('bold', False)
        font.italic = style.get('italic', False)

        if 'color' in style:
            font.color.rgb = hex_to_rgb(style['color'])

        # Corrected: Use PP_ALIGN for horizontal text alignment
        alignment = style.get('alignment', 'left')
        if alignment == 'center':
            p.alignment = PP_ALIGN.CENTER
        elif alignment == 'right':
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT  # Default to left

        logging.info(f"Added text box at ({element_data['x']},{element_data['y']}) with content: '{content[:50]}...'")
    except Exception as e:
        logging.error(f"Error adding text box: {e}", exc_info=True)


def add_image(slide, image_path: str, element_data: dict):
    """Adds an image to the slide based on element_data and image_path."""
    x = px_to_emu(element_data['x'])
    y = px_to_emu(element_data['y'])
    width = px_to_emu(element_data['width'])
    height = px_to_emu(element_data['height'])
    # Opacity is now handled by image_service, so no direct styling here
    # style = element_data.get('style', {}) # No longer needed for opacity here

    try:
        pic = slide.shapes.add_picture(image_path, x, y, width, height)
        # The warning about opacity has been removed as it's now handled by ImageService
        logging.info(f"Added image at ({element_data['x']},{element_data['y']}) from path: {image_path}")
    except Exception as e:
        logging.error(f"Error adding image: {e}", exc_info=True)


def add_shape(slide, element_data: dict):
    """Adds a shape to the slide based on element_data."""
    x = px_to_emu(element_data['x'])
    y = px_to_emu(element_data['y'])
    width = px_to_emu(element_data['width'])
    height = px_to_emu(element_data['height'])
    shape_type_str = element_data.get('shape_type', 'rectangle').lower()
    style = element_data.get('style', {})

    shape_type_map = {
        'rectangle': MSO_SHAPE.RECTANGLE,
        'oval': MSO_SHAPE.OVAL,
        'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
        'line': MSO_SHAPE['LINE'],  # Changed from MSO_SHAPE.LINE to MSO_SHAPE['LINE']
    }

    if shape_type_str not in shape_type_map:
        logging.warning(f"Unsupported shape type: {shape_type_str}. Defaulting to rectangle.")
        shape_type = MSO_SHAPE.RECTANGLE
    else:
        shape_type = shape_type_map[shape_type_str]

    try:
        shape = slide.shapes.add_shape(shape_type, x, y, width, height)

        fill = shape.fill
        line = shape.line

        # Apply fill style
        if 'fill_color' in style and style['fill_color'] is not None:
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(style['fill_color'])
            if 'opacity' in style:
                fill.fore_color.alpha = int(style['opacity'] * 100)
        else:
            fill.background()  # No fill

        # Apply line style
        if 'line_color' in style and style['line_color'] is not None:
            line.fill.solid()
            line.fill.fore_color.rgb = hex_to_rgb(style['line_color'])
            line.width = Pt(style.get('line_width', 1))
        else:
            line.fill.background()  # No line

        logging.info(f"Added {shape_type_str} shape at ({element_data['x']},{element_data['y']})")
    except Exception as e:
        logging.error(f"Error adding shape: {e}", exc_info=True)


def add_chart(slide, element_data: dict):
    """Adds a chart to the slide based on element_data."""
    x = px_to_emu(element_data['x'])
    y = px_to_emu(element_data['y'])
    width = px_to_emu(element_data['width'])
    height = px_to_emu(element_data['height'])
    chart_type_str = element_data.get('chart_type', 'column_chart').lower()
    data = element_data.get('data', {})
    title = element_data.get('title', '')
    style = element_data.get('style', {})

    chart_type_map = {
        'bar_chart': XL_CHART_TYPE.BAR_CLUSTERED,
        'column_chart': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'line_chart': XL_CHART_TYPE.LINE,
        'pie_chart': XL_CHART_TYPE.PIE,
    }

    if chart_type_str not in chart_type_map:
        logging.warning(f"Unsupported chart type: {chart_type_str}. Defaulting to column chart.")
        chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
    else:
        chart_type = chart_type_map[chart_type_str]

    chart_data = CategoryChartData()
    if chart_type in [XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.LINE]:
        chart_data.categories = data.get('categories', [])
        for series_data in data.get('series', []):
            chart_data.add_series(series_data.get('name', ''), series_data.get('values', []))
    elif chart_type == XL_CHART_TYPE.PIE:
        # Pie charts handle data differently, typically one series
        chart_data = ChartData()
        chart_data.categories = data.get('labels', [])
        chart_data.add_series(title, data.get('values', []))

    try:
        graphic_frame = slide.shapes.add_chart(
            chart_type, x, y, width, height, chart_data
        )
        chart = graphic_frame.chart
        chart.has_title = True
        chart.chart_title.text_frame.text = title

        # Apply chart specific styles
        chart.has_data_labels = style.get('data_labels', False)

        # Complete implementation for legend positioning
        legend_position_str = style.get('legend_position')
        if legend_position_str:
            chart.has_legend = True
            legend_position_map = {
                'bottom': XL_LEGEND_POSITION.BOTTOM,
                'corner': XL_LEGEND_POSITION.CORNER,  # Often top-right
                'left': XL_LEGEND_POSITION.LEFT,
                'right': XL_LEGEND_POSITION.RIGHT,
                'top': XL_LEGEND_POSITION.TOP,
            }
            if legend_position_str.lower() in legend_position_map:
                chart.legend.position = legend_position_map[legend_position_str.lower()]
            else:
                logging.warning(f"Unsupported legend position: {legend_position_str}. Defaulting to standard behavior.")
        else:
            chart.has_legend = False  # Ensure legend is off if not specified

        logging.info(f"Added {chart_type_str} chart at ({element_data['x']},{element_data['y']}) with title: '{title}'")
    except Exception as e:
        logging.error(f"Error adding chart: {e}", exc_info=True)
