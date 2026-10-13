"""
Скрипт для сборки официального отчета в формате DOCX, DOC и PDF без знаков подчеркивания в именах.
"""
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

GITHUB_REPO = "https://github.com/Ganger228-vibe/samara-university-python-intro"
COLAB_URL = "https://colab.research.google.com/github/Ganger228-vibe/samara-university-python-intro/blob/main/notebook.ipynb"
SCREENSHOTS_DIR = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\screenshots"
OUT_DOCX = r"c:\graviti\для  учебы автоматизации\Отчет Введение в программирование Кириакиди.docx"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_hyperlink(paragraph, url, text, color="0969DA", underline=True, font_size_pt=11):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    sz_val = int(font_size_pt * 2)
    new_run = parse_xml(f'''
        <w:r {nsdecls("w")}>
            <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                <w:color w:val="{color}"/>
                <w:sz w:val="{sz_val}"/>
                <w:szCs w:val="{sz_val}"/>
                {"<w:u w:val='single'/>" if underline else ""}
            </w:rPr>
            <w:t>{text}</w:t>
        </w:r>
    ''')
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_code_block(doc, code_str):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.2)
    set_cell_background(cell, "F6F8FA")
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="D0D7DE"/>
            <w:left w:val="single" w:sz="18" w:space="0" w:color="0969DA"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="D0D7DE"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="D0D7DE"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    
    run = p.add_run(code_str.strip())
    run.font.name = "Consolas"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(2)
    p_after.paragraph_format.space_after = Pt(4)

TASKS_IO = [
    {
        "num": 1,
        "title": "Периметр квадрата",
        "task_desc": "Дана сторона квадрата a. Найти его периметр P = 4 × a.",
        "file": "io1.py",
        "img": "io1.png",
        "explanation": (
            "Считываем значение стороны a с консоли и преобразуем во вещественный тип float. "
            "Периметр квадрата вычисляется умножением стороны на 4 (p = 4 * a) и выводится с помощью print()."
        ),
        "code": '''a = float(input())
p = 4 * a
print(p)'''
    },
    {
        "num": 2,
        "title": "Площадь квадрата",
        "task_desc": "Дана сторона квадрата a. Найти его площадь S = a².",
        "file": "io2.py",
        "img": "io2.png",
        "explanation": (
            "Вводим сторону квадрата. Операция возведения в степень в Python записывается как **. "
            "Площадь квадрата рассчитывается как a ** 2 и выводится на экран."
        ),
        "code": '''a = float(input())
s = a ** 2
print(s)'''
    },
    {
        "num": 3,
        "title": "Площадь и периметр прямоугольника",
        "task_desc": "Даны стороны прямоугольника a и b. Найти его площадь S = a × b и периметр P = 2 × (a + b).",
        "file": "io3.py",
        "img": "io3.png",
        "explanation": (
            "Считываются две стороны прямоугольника a и b. "
            "Площадь находится умножением сторон (a * b), а периметр — удвоенной суммой 2 * (a + b). "
            "Скобки гарантируют выполнение сложения до умножения."
        ),
        "code": '''a = float(input())
b = float(input())
s = a * b
p = 2 * (a + b)
print(s)
print(p)'''
    },
    {
        "num": 4,
        "title": "Длина окружности по диаметру",
        "task_desc": "Дан диаметр окружности d. Найти ее длину L = π × d. В качестве значения π использовать 3.14.",
        "file": "io4.py",
        "img": "io4.png",
        "explanation": (
            "Константа pi задается равной 3.14 строго по условию. "
            "Считывается диаметр d (float). Длина окружности вычисляется по формуле l = pi * d."
        ),
        "code": '''pi = 3.14
d = float(input())
l = pi * d
print(l)'''
    },
    {
        "num": 5,
        "title": "Объем и площадь поверхности куба",
        "task_desc": "Дана длина ребра куба a. Найти объем куба V = a³ и площадь его поверхности S = 6 × a².",
        "file": "io5.py",
        "img": "io5.png",
        "explanation": (
            "Считывается ребро куба a. Объем куба вычисляется возведением в третью степень (a ** 3), "
            "а площадь полной поверхности — умножением площади одной грани на 6: 6 * (a ** 2)."
        ),
        "code": '''a = float(input())
v = a ** 3
s = 6 * (a ** 2)
print(v)
print(s)'''
    },
    {
        "num": 6,
        "title": "Объем и площадь поверхности прямоугольного параллелепипеда",
        "task_desc": "Даны длины ребер a, b, c прямоугольного параллелепипеда. Найти его объем V = a × b × c и площадь поверхности S = 2 × (a × b + b × c + a × c).",
        "file": "io6.py",
        "img": "io6.png",
        "explanation": (
            "Считываются три измерения: a, b и c. "
            "Объем равен их произведению a * b * c. "
            "Площадь поверхности вычисляется как удвоенная сумма попарных произведений смежных граней: 2 * (a * b + b * c + a * c)."
        ),
        "code": '''a = float(input())
b = float(input())
c = float(input())
v = a * b * c
s = 2 * (a * b + b * c + a * c)
print(v)
print(s)'''
    },
    {
        "num": 7,
        "title": "Длина окружности и площадь круга",
        "task_desc": "Найти длину окружности L и площадь круга S заданного радиуса R: L = 2 × π × R, S = π × R². В качестве значения π использовать 3.14.",
        "file": "io7.py",
        "img": "io7.png",
        "explanation": (
            "Задается число pi = 3.14 и считывается радиус r. "
            "Длина окружности рассчитывается по формуле 2 * pi * r, а площадь круга — pi * (r ** 2)."
        ),
        "code": '''pi = 3.14
r = float(input())
l = 2 * pi * r
s = pi * (r ** 2)
print(l)
print(s)'''
    },
    {
        "num": 8,
        "title": "Среднее арифметическое двух чисел",
        "task_desc": "Даны два числа a и b. Найти их среднее арифметическое: (a + b) / 2.",
        "file": "io8.py",
        "img": "io8.png",
        "explanation": (
            "Считываются два вещественных числа. "
            "Среднее арифметическое находится суммированием чисел и делением результата на 2: (a + b) / 2."
        ),
        "code": '''a = float(input())
b = float(input())
avg = (a + b) / 2
print(avg)'''
    },
    {
        "num": 9,
        "title": "Среднее геометрическое двух неотрицательных чисел",
        "task_desc": "Даны два неотрицательных числа a и b. Найти их среднее геометрическое: √(a × b).",
        "file": "io9.py",
        "img": "io9.png",
        "explanation": (
            "Считываются два неотрицательных числа a и b. "
            "Извлечение квадратного корня реализуется возведением произведения в степень 0.5: (a * b) ** 0.5."
        ),
        "code": '''a = float(input())
b = float(input())
geom = (a * b) ** 0.5
print(geom)'''
    },
    {
        "num": 10,
        "title": "Операции над квадратами двух ненулевых чисел",
        "task_desc": "Даны два ненулевых числа. Найти сумму, разность, произведение и частное их квадратов.",
        "file": "io10.py",
        "img": "io10.png",
        "explanation": (
            "Считываются ненулевые числа a и b. "
            "Квадраты сохраняются в переменные a2 = a ** 2 и b2 = b ** 2. "
            "Последовательно вычисляются и выводятся сумма (a2 + b2), разность (a2 - b2), произведение (a2 * b2) и частное (a2 / b2)."
        ),
        "code": '''a = float(input())
b = float(input())
a2 = a ** 2
b2 = b ** 2
print(a2 + b2)
print(a2 - b2)
print(a2 * b2)
print(a2 / b2)'''
    }
]

TASKS_IF = [
    {
        "num": 1,
        "title": "Модификация положительного целого числа",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; в противном случае не изменять его. Вывести полученное число.",
        "file": "if1.py",
        "img": "if1.png",
        "explanation": (
            "Считывается целое число n через int(input()). "
            "Используется неполное ветвление if n > 0:. Если число положительное, выполняется n += 1. "
            "В противном случае значение n не изменяется."
        ),
        "code": '''n = int(input())
if n > 0:
    n += 1
print(n)'''
    },
    {
        "num": 2,
        "title": "Двухвариантное изменение целого числа",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; в противном случае вычесть из него 2. Вывести полученное число.",
        "file": "if2.py",
        "img": "if2.png",
        "explanation": (
            "Используется конструкция if - else. "
            "Если число n положительное (n > 0), к нему прибавляется 1 (n += 1). "
            "В противном случае (n <= 0) из него вычитается 2 (n -= 2)."
        ),
        "code": '''n = int(input())
if n > 0:
    n += 1
else:
    n -= 2
print(n)'''
    },
    {
        "num": 3,
        "title": "Трёхвариантное ветвление",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; если отрицательным, то вычесть из него 2; если нулевым, то заменить его на 10. Вывести полученное число.",
        "file": "if3.py",
        "img": "if3.png",
        "explanation": (
            "Для трех вариантов исхода применяется цепочка if - elif - else. "
            "Если n > 0, выполняется n += 1; если n < 0, выполняется n -= 2; если n == 0, число заменяется на 10 (n = 10)."
        ),
        "code": '''n = int(input())
if n > 0:
    n += 1
elif n < 0:
    n -= 2
else:
    n = 10
print(n)'''
    },
    {
        "num": 4,
        "title": "Подсчет количества положительных чисел из трех",
        "task_desc": "Даны три целых числа. Найти количество положительных чисел в исходном наборе.",
        "file": "if4.py",
        "img": "if4.png",
        "explanation": (
            "Инициализируется счетчик count = 0. "
            "Каждое из трех чисел a, b, c независимо проверяется оператором if на условие строго больше нуля. "
            "При истинности условия счетчик увеличивается на 1."
        ),
        "code": '''a = int(input())
b = int(input())
c = int(input())
count = 0
if a > 0:
    count += 1
if b > 0:
    count += 1
if c > 0:
    count += 1
print(count)'''
    },
    {
        "num": 5,
        "title": "Подсчет положительных и отрицательных чисел из трех",
        "task_desc": "Даны три целых числа. Найти количество положительных и количество отрицательных чисел в исходном наборе.",
        "file": "if5.py",
        "img": "if5.png",
        "explanation": (
            "Создаются два счетчика: pos = 0 и neg = 0. "
            "В цикле for x in (a, b, c) каждое число проверяется: если x > 0, увеличивается pos; если x < 0 — neg. "
            "Нулевые значения корректно игнорируются."
        ),
        "code": '''a = int(input())
b = int(input())
c = int(input())
pos = 0
neg = 0
for x in (a, b, c):
    if x > 0:
        pos += 1
    elif x < 0:
        neg += 1
print(pos)
print(neg)'''
    }
]

def build_document():
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.59)
    
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x24, 0x29, 0x2F)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    
    # ------------------ ТИТУЛЬНЫЙ ЛИСТ ------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ\n")
    r.font.size = Pt(10)
    r.bold = True
    r2 = p.add_run("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ\nВЫСШЕГО ОБРАЗОВАНИЯ\n«САМАРСКИЙ НАЦИОНАЛЬНЫЙ ИССЛЕДОВАТЕЛЬСКИЙ УНИВЕРСИТЕТ\nИМЕНИ АКАДЕМИКА С.П. КОРОЛЕВА»\n")
    r2.font.size = Pt(10)
    r2.bold = True
    r3 = p.add_run("(Самарский университет)\n\nКафедра математики и бизнес-информатики\n")
    r3.font.size = Pt(11)
    
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(70)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_otchet = p_title.add_run("ОТЧЕТ ПО ПРАКТИЧЕСКИМ ЗАДАНИЯМ\n")
    r_otchet.bold = True
    r_otchet.font.size = Pt(18)
    r_otchet.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
    
    r_subj = p_title.add_run("по дисциплине «Введение в программирование»\n\n")
    r_subj.font.size = Pt(14)
    r_subj.bold = True
    
    r_topic = p_title.add_run("Тема: «Ввод и вывод данных, оператор присваивания. Условный оператор»\n")
    r_topic.font.size = Pt(13)
    r_topic.italic = True
    
    p_div2 = doc.add_paragraph()
    p_div2.paragraph_format.space_before = Pt(80)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_meta1 = p_meta.add_run("Выполнил:\nстудент группы 7251\n")
    r_meta1.font.size = Pt(12)
    r_meta2 = p_meta.add_run("Кириакиди Дмитрий Андреевич\n\n")
    r_meta2.font.size = Pt(12)
    r_meta2.bold = True
    r_meta3 = p_meta.add_run("Проверил:\nстарший преподаватель\n")
    r_meta3.font.size = Pt(12)
    r_meta4 = p_meta.add_run("Аксинин Владимир Иванович\n")
    r_meta4.font.size = Pt(12)
    r_meta4.bold = True
    
    p_div3 = doc.add_paragraph()
    p_div3.paragraph_format.space_before = Pt(100)
    
    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_city = p_city.add_run("Самара, 2026")
    r_city.font.size = Pt(12)
    
    doc.add_page_break()
    
    # ------------------ ВВЕДЕНИЕ ------------------
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("ВВЕДЕНИЕ И ИНФОРМАЦИЯ О РЕШЕНИИ")
    r_h1.font.name = "Times New Roman"
    r_h1.font.size = Pt(16)
    r_h1.bold = True
    r_h1.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
    
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.first_line_indent = Inches(0.49)
    p_intro.add_run(
        "Настоящий отчет содержит программную реализацию, консольные снимки экрана и подробный "
        "пошаговый разбор 15 заданий по программированию на языке Python 3: 10 заданий из темы "
        "«Ввод и вывод данных, оператор присваивания» (задания 1–10, программы io1.py - io10.py) и "
        "5 заданий из темы «Условный оператор» (задания 1–5, программы if1.py - if5.py). "
        "Все программы написаны в академическом минималистичном стиле без избыточных текстовых приглашений ко вводу."
    )
    
    p_links_head = doc.add_paragraph()
    r_lh = p_links_head.add_run("Ссылки на репозиторий и интерактивное решение:")
    r_lh.bold = True
    
    p_link1 = doc.add_paragraph(style='List Bullet')
    p_link1.add_run("GitHub-репозиторий со всеми исходниками: ")
    add_hyperlink(p_link1, GITHUB_REPO, GITHUB_REPO)
    
    p_link2 = doc.add_paragraph(style='List Bullet')
    p_link2.add_run("Интерактивный запуск в Google Colab: ")
    add_hyperlink(p_link2, COLAB_URL, "Открыть notebook.ipynb в Google Colab")
    
    doc.add_paragraph()
    
    # ------------------ РАЗДЕЛ 1 ------------------
    h_sec1 = doc.add_heading(level=1)
    r_s1 = h_sec1.add_run("РАЗДЕЛ 1. ВВОД И ВЫВОД ДАННЫХ, ОПЕРАТОР ПРИСВАИВАНИЯ")
    r_s1.font.name = "Times New Roman"
    r_s1.font.size = Pt(15)
    r_s1.bold = True
    r_s1.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
    
    p_s1_note = doc.add_paragraph()
    r_note = p_s1_note.add_run("Примечание к разделу: ")
    r_note.bold = True
    p_s1_note.add_run("Все входные и выходные данные в заданиях этой группы являются вещественными числами (тип float).")
    
    for task in TASKS_IO:
        h_task = doc.add_heading(level=2)
        r_ht = h_task.add_run(f"Задание 1.{task['num']}. {task['title']}")
        r_ht.font.name = "Times New Roman"
        r_ht.font.size = Pt(13)
        r_ht.bold = True
        r_ht.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
        
        p_desc = doc.add_paragraph()
        r_dl = p_desc.add_run("Условие: ")
        r_dl.bold = True
        p_desc.add_run(task["task_desc"])
        
        p_url = doc.add_paragraph()
        r_ul = p_url.add_run("Ссылка на решение (GitHub): ")
        r_ul.bold = True
        file_url = f"{GITHUB_REPO}/blob/main/{task['file']}"
        add_hyperlink(p_url, file_url, f"{task['file']}")
        
        p_code_lbl = doc.add_paragraph()
        r_cl = p_code_lbl.add_run("Исходный код программы:")
        r_cl.italic = True
        add_code_block(doc, task["code"])
        
        img_path = os.path.join(SCREENSHOTS_DIR, task["img"])
        if os.path.exists(img_path):
            p_img_lbl = doc.add_paragraph()
            r_il = p_img_lbl.add_run("Результат выполнения программы (скриншот консоли):")
            r_il.italic = True
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(6)
            doc.add_picture(img_path, width=Inches(5.5))
        
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.first_line_indent = Inches(0.35)
        r_el = p_exp.add_run("Разбор решения: ")
        r_el.bold = True
        p_exp.add_run(task["explanation"])
        
        doc.add_paragraph()
    
    doc.add_page_break()
    
    # ------------------ РАЗДЕЛ 2 ------------------
    h_sec2 = doc.add_heading(level=1)
    r_s2 = h_sec2.add_run("РАЗДЕЛ 2. УСЛОВНЫЙ ОПЕРАТОР")
    r_s2.font.name = "Times New Roman"
    r_s2.font.size = Pt(15)
    r_s2.bold = True
    r_s2.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
    
    p_s2_note = doc.add_paragraph()
    r_note2 = p_s2_note.add_run("Примечание к разделу: ")
    r_note2.bold = True
    p_s2_note.add_run("Все входные числа являются целыми (тип int). Обработка условий выполняется операторами if, elif, else.")
    
    for task in TASKS_IF:
        h_task = doc.add_heading(level=2)
        r_ht = h_task.add_run(f"Задание 2.{task['num']}. {task['title']}")
        r_ht.font.name = "Times New Roman"
        r_ht.font.size = Pt(13)
        r_ht.bold = True
        r_ht.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
        
        p_desc = doc.add_paragraph()
        r_dl = p_desc.add_run("Условие: ")
        r_dl.bold = True
        p_desc.add_run(task["task_desc"])
        
        p_url = doc.add_paragraph()
        r_ul = p_url.add_run("Ссылка на решение (GitHub): ")
        r_ul.bold = True
        file_url = f"{GITHUB_REPO}/blob/main/{task['file']}"
        add_hyperlink(p_url, file_url, f"{task['file']}")
        
        p_code_lbl = doc.add_paragraph()
        r_cl = p_code_lbl.add_run("Исходный код программы:")
        r_cl.italic = True
        add_code_block(doc, task["code"])
        
        img_path = os.path.join(SCREENSHOTS_DIR, task["img"])
        if os.path.exists(img_path):
            p_img_lbl = doc.add_paragraph()
            r_il = p_img_lbl.add_run("Результат выполнения программы (скриншот консоли):")
            r_il.italic = True
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(6)
            doc.add_picture(img_path, width=Inches(5.5))
        
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.first_line_indent = Inches(0.35)
        r_el = p_exp.add_run("Разбор решения: ")
        r_el.bold = True
        p_exp.add_run(task["explanation"])
        
        doc.add_paragraph()
        
    doc.add_page_break()
    
    # ------------------ ЗАКЛЮЧЕНИЕ ------------------
    h_conc = doc.add_heading(level=1)
    r_conc = h_conc.add_run("ЗАКЛЮЧЕНИЕ И СВОДНАЯ ТАБЛИЦА")
    r_conc.font.name = "Times New Roman"
    r_conc.font.size = Pt(16)
    r_conc.bold = True
    r_conc.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
    
    p_conc = doc.add_paragraph()
    p_conc.paragraph_format.first_line_indent = Inches(0.49)
    p_conc.add_run(
        "В ходе работы освоены ключевые конструкции языка Python: считывание числовых значений, "
        "явное приведение типов (float и int), работа со степенями и арифметикой, организация условий "
        "через if, elif, else. Программы реализованы компактно и прозрачно, все результаты проверены на тестовых наборах данных."
    )
    
    doc.add_paragraph()
    p_tbl_lbl = doc.add_paragraph()
    r_tbl_lbl = p_tbl_lbl.add_run("Сводная таблица ссылок на выполненные задания:")
    r_tbl_lbl.bold = True
    
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    hdr_cells = tbl.rows[0].cells
    hdr_cells[0].width = Inches(0.8)
    hdr_cells[1].width = Inches(3.2)
    hdr_cells[2].width = Inches(2.2)
    
    headers = ["№", "Наименование задания", "Файл на GitHub"]
    for i, h_text in enumerate(headers):
        set_cell_background(hdr_cells[i], "0969DA")
        p_c = hdr_cells[i].paragraphs[0]
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_c = p_c.add_run(h_text)
        r_c.bold = True
        r_c.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r_c.font.size = Pt(10.5)
    
    all_tasks = [("1." + str(t["num"]), t["title"], t["file"]) for t in TASKS_IO] + \
                [("2." + str(t["num"]), t["title"], t["file"]) for t in TASKS_IF]
                
    for row_idx, (t_no, t_name, t_file) in enumerate(all_tasks):
        row = tbl.add_row()
        c0, c1, c2 = row.cells
        c0.width = Inches(0.8)
        c1.width = Inches(3.2)
        c2.width = Inches(2.2)
        
        bg_col = "F6F8FA" if row_idx % 2 == 1 else "FFFFFF"
        set_cell_background(c0, bg_col)
        set_cell_background(c1, bg_col)
        set_cell_background(c2, bg_col)
        
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run(t_no).font.size = Pt(10)
        
        p1 = c1.paragraphs[0]
        p1.add_run(t_name).font.size = Pt(10)
        
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_hyperlink(p2, f"{GITHUB_REPO}/blob/main/{t_file}", t_file, font_size_pt=10)
        
    doc.save(OUT_DOCX)
    print(f"DOCX сохранен: {OUT_DOCX}")
    
    # Also save inside lab1_python_intro directory
    second_path = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\Отчет Введение в программирование Кириакиди.docx"
    doc.save(second_path)
    print(f"Копия сохранена: {second_path}")

if __name__ == "__main__":
    build_document()
