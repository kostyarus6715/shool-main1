import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
from PyQt6.QtGui import QFontMetrics, QFont

def load_excel_data(file_path):
    """Загружает данные из Excel файла и возвращает список словарей."""
    data_list = []
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]  
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data = dict(zip(headers, row))
            data_list.append(data)
    except Exception as e:
        raise ValueError(f"Ошибка при загрузке данных из Excel файла: {e}")

    return data_list

def hex_to_rgb(hex_color):
    """Преобразует HEX цвет в формат RGB (значения от 0 до 1)"""
    hex_color = hex_color.lstrip('#')  
    return (
        int(hex_color[0:2], 16) / 255.0, 
        int(hex_color[2:4], 16) / 255.0, 
        int(hex_color[4:6], 16) / 255.0
    )

def create_certificate(
        data,
        output_path,
        template_file_path, 
        font_file_path,
        line_spacing,  
        x_position, 
        y_position, 
        selected_columns,  
        show_headers,  
        column_font_sizes,
        column_spacing,
        column_colors
    ):

    font_name = "CustomFont"
    pdfmetrics.registerFont(TTFont(font_name, font_file_path))  

    reader = PdfReader(template_file_path)  
    writer = PdfWriter()

    output_pdf = BytesIO()
    c = canvas.Canvas(output_pdf, pagesize=A4)  

    current_y = y_position
    max_line_width = 500  

    for column, include in selected_columns.items():  
        if include and column in data:  
            text_value = str(data[column]) if data[column] is not None else ""

            if show_headers.get(column, False):
                text = f"{column}: {text_value}"
            else:
                text = text_value

            font_size = column_font_sizes.get(column, 12)  
            c.setFont(font_name, font_size)

            column_color = column_colors.get(column, "#000000")  
            rgb_color = hex_to_rgb(column_color)  
            c.setFillColorRGB(*rgb_color)  

            lines = wrap_text(text, font_name, font_size, max_line_width)  
            for line in lines:
                c.drawString(x_position, current_y, line)  
                current_y -= line_spacing  

            if column in column_spacing:
                current_y -= column_spacing[column]  

            if current_y < 50:  
                c.showPage()  
                c.setFont(font_name, font_size)  
                c.setFillColorRGB(*rgb_color)  
                current_y = y_position  

    c.save()
    output_pdf.seek(0)

    overlay_reader = PdfReader(output_pdf)
    overlay_page = overlay_reader.pages[0]

    template_page = reader.pages[0]
    template_page.merge_page(overlay_page)  

    writer.add_page(template_page)

    with open(output_path, "wb") as f:
        writer.write(f)

def wrap_text(text, font_name, font_size, max_width):
    """Функция для переноса текста в строки, чтобы он вмещался в указанный предел по ширине"""
    font = QFont(font_name, font_size)
    font_metrics = QFontMetrics(font)

    words = text.split()
    lines = []
    current_line = ""
    for word in words:

        if font_metrics.horizontalAdvance(current_line + " " + word) < max_width:
            current_line += " " + word if current_line else word
        else:
            lines.append(current_line)  
            current_line = word  
    if current_line:
        lines.append(current_line)  
    return lines