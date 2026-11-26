import os
from docx import Document
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def _ensure_dir():
    if not os.path.exists("output"):
        os.makedirs("output")

def export_to_docx(items, filename="output/employees.docx"):
    _ensure_dir()
    doc = Document()
    doc.add_heading('Штатное расписание сотрудников', 0)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'ID'
    hdr_cells[1].text = 'ФИО'
    hdr_cells[2].text = 'Должность'
    hdr_cells[3].text = 'Отдел'
    hdr_cells[4].text = 'Оклад'
    
    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item[0])
        row_cells[1].text = str(item[1])
        row_cells[2].text = str(item[2])
        row_cells[3].text = str(item[3])
        row_cells[4].text = str(item[4])
        
    doc.save(filename)
    return os.path.abspath(filename)

def export_to_xlsx(items, filename="output/employees.xlsx"):
    _ensure_dir()
    wb = Workbook()
    ws = wb.active
    ws.title = "Сотрудники"
    headers = ['ID', 'ФИО', 'Должность', 'Отдел', 'Оклад']
    ws.append(headers)
    for item in items:
        ws.append(list(item))
    wb.save(filename)
    return os.path.abspath(filename)

def export_to_pdf(items, filename="output/employees.pdf"):
    _ensure_dir()
    
    # Подключаем шрифт Arial для поддержки кириллицы
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'C:/Windows/Fonts/arial.ttf'))
        font_name = 'Arial'
    except:
        font_name = 'Helvetica'
        
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    
    c.setFont(font_name, 14)
    c.drawString(50, height - 50, "Список сотрудников")
    
    c.setFont(font_name, 10)
    y = height - 80
    c.drawString(50, y, "ID")
    c.drawString(80, y, "ФИО")
    c.drawString(250, y, "Должность")
    c.drawString(400, y, "Оклад")
    y -= 20
    
    for item in items:
        if y < 50:
            c.showPage()
            c.setFont(font_name, 10)
            y = height - 50
        c.drawString(50, y, str(item[0]))
        c.drawString(80, y, str(item[1])[:20])
        c.drawString(250, y, str(item[2])[:20])
        c.drawString(400, y, str(item[4]))
        y -= 20
    c.save()
    return os.path.abspath(filename)