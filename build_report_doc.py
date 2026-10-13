"""
Скрипт для сборки официального академического отчета в формате DOCX.
Включает титульный лист, цели, теорию, 15 заданий (код, скриншоты, пояснения, ссылки) и заключение.
"""
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

GITHUB_REPO = "https://github.com/Ganger228-vibe/samara-university-python-intro"
COLAB_URL = "https://colab.research.google.com/github/Ganger228-vibe/samara-university-python-intro/blob/main/interactive_lab.ipynb"
SCREENSHOTS_DIR = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\screenshots"
OUT_DOCX = r"c:\graviti\для  учебы автоматизации\Отчет_Введение_в_программирование_Кириакиди.docx"

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
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    # Border styling
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
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
    
    # Empty spacing after table
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(2)
    p_after.paragraph_format.space_after = Pt(4)

TASKS_IO = [
    {
        "num": 1,
        "title": "Периметр квадрата",
        "task_desc": "Дана сторона квадрата a. Найти его периметр P = 4 × a.",
        "file": "task_io_01.py",
        "img": "task_io_01.png",
        "math_formula": "P = 4 \\times a",
        "explanation": (
            "Для решения задачи значение стороны квадрата считывается с помощью функции input() и "
            "преобразуется в вещественное число функцией float(input()). "
            "Вычисление периметра производится по формуле 4.0 * a. Результат сохраняется в переменной "
            "perimeter и выводится на экран с помощью форматированной строки f-string."
        ),
        "code": '''def main():
    print("=== Задание 1. Периметр квадрата ===")
    a = float(input("Введите сторону квадрата a: "))
    
    # Вычисление периметра по формуле P = 4 * a
    perimeter = 4.0 * a
    
    print(f"Периметр квадрата P = {perimeter}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 2,
        "title": "Площадь квадрата",
        "task_desc": "Дана сторона квадрата a. Найти его площадь S = a².",
        "file": "task_io_02.py",
        "img": "task_io_02.png",
        "math_formula": "S = a^2",
        "explanation": (
            "Пользователь вводит длину стороны квадрата a в виде вещественного числа. "
            "В Python операция возведения в степень обозначается оператором **. "
            "Площадь вычисляется выражением a ** 2 и выводится пользователю."
        ),
        "code": '''def main():
    print("=== Задание 2. Площадь квадрата ===")
    a = float(input("Введите сторону квадрата a: "))
    
    # Вычисление площади по формуле S = a^2
    area = a ** 2
    
    print(f"Площадь квадрата S = {area}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 3,
        "title": "Площадь и периметр прямоугольника",
        "task_desc": "Даны стороны прямоугольника a и b. Найти его площадь S = a × b и периметр P = 2 × (a + b).",
        "file": "task_io_03.py",
        "img": "task_io_03.png",
        "math_formula": "S = a \\times b, \\quad P = 2 \\times (a + b)",
        "explanation": (
            "С клавиатуры последовательно запрашиваются две вещественные стороны прямоугольника: a и b. "
            "Площадь S рассчитывается как произведение a * b. Периметр рассчитывается как сумма сторон, "
            "умноженная на 2: 2.0 * (a + b). Скобки обеспечивают приоритет сложения перед умножением."
        ),
        "code": '''def main():
    print("=== Задание 3. Площадь и периметр прямоугольника ===")
    a = float(input("Введите сторону прямоугольника a: "))
    b = float(input("Введите сторону прямоугольника b: "))
    
    # Вычисление площади и периметра
    area = a * b
    perimeter = 2.0 * (a + b)
    
    print(f"Площадь прямоугольника S = {area}")
    print(f"Периметр прямоугольника P = {perimeter}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 4,
        "title": "Длина окружности по диаметру",
        "task_desc": "Дан диаметр окружности d. Найти ее длину L = π × d. В качестве значения π использовать 3.14.",
        "file": "task_io_04.py",
        "img": "task_io_04.png",
        "math_formula": "L = \\pi \\times d \\quad (\\pi = 3.14)",
        "explanation": (
            "Значение числа пи объявляется в виде именованной константы PI = 3.14 в строгом соответствии с условием. "
            "Считывается диаметр d (тип float). Длина окружности вычисляется как произведение PI * d."
        ),
        "code": '''PI = 3.14

def main():
    print("=== Задание 4. Длина окружности по диаметру ===")
    d = float(input("Введите диаметр окружности d: "))
    
    # Вычисление длины окружности L = pi * d
    length = PI * d
    
    print(f"Длина окружности L = {length}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 5,
        "title": "Объем и площадь поверхности куба",
        "task_desc": "Дана длина ребра куба a. Найти объем куба V = a³ и площадь его поверхности S = 6 × a².",
        "file": "task_io_05.py",
        "img": "task_io_05.png",
        "math_formula": "V = a^3, \\quad S = 6 \\times a^2",
        "explanation": (
            "Длина ребра a считывается как вещественное число. "
            "Объем куба находится возведением в третью степень: a ** 3. "
            "Площадь поверхности куба состоит из 6 одинаковых граней-квадратов: 6.0 * (a ** 2)."
        ),
        "code": '''def main():
    print("=== Задание 5. Объем и площадь поверхности куба ===")
    a = float(input("Введите длину ребра куба a: "))
    
    # Вычисление объема куба и площади поверхности
    volume = a ** 3
    surface_area = 6.0 * (a ** 2)
    
    print(f"Объем куба V = {volume}")
    print(f"Площадь поверхности куба S = {surface_area}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 6,
        "title": "Объем и площадь поверхности прямоугольного параллелепипеда",
        "task_desc": "Даны длины ребер a, b, c прямоугольного параллелепипеда. Найти его объем V = a × b × c и площадь поверхности S = 2 × (a × b + b × c + a × c).",
        "file": "task_io_06.py",
        "img": "task_io_06.png",
        "math_formula": "V = a \\times b \\times c, \\quad S = 2 \\times (ab + bc + ac)",
        "explanation": (
            "Считываются длины трех взаимно перпендикулярных ребер a, b, c (тип float). "
            "Объем находится произведением всех трех измерений a * b * c. "
            "Площадь поверхности вычисляется как удвоенная сумма попарных площадей трех смежных граней: 2.0 * (a * b + b * c + a * c)."
        ),
        "code": '''def main():
    print("=== Задание 6. Объем и площадь поверхности параллелепипеда ===")
    a = float(input("Введите длину ребра a: "))
    b = float(input("Введите длину ребра b: "))
    c = float(input("Введите длину ребра c: "))
    
    # Вычисление объема и площади полной поверхности
    volume = a * b * c
    surface_area = 2.0 * (a * b + b * c + a * c)
    
    print(f"Объем параллелепипеда V = {volume}")
    print(f"Площадь поверхности S = {surface_area}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 7,
        "title": "Длина окружности и площадь круга заданного радиуса",
        "task_desc": "Найти длину окружности L и площадь круга S заданного радиуса R: L = 2 × π × R, S = π × R². В качестве значения π использовать 3.14.",
        "file": "task_io_07.py",
        "img": "task_io_07.png",
        "math_formula": "L = 2 \\times \\pi \\times R, \\quad S = \\pi \\times R^2 \\quad (\\pi = 3.14)",
        "explanation": (
            "Радиус R считывается с консоли с преобразованием во float. Число пи зафиксировано как 3.14. "
            "Длина окружности определяется по формуле 2.0 * PI * r, а площадь круга — по формуле PI * (r ** 2)."
        ),
        "code": '''PI = 3.14

def main():
    print("=== Задание 7. Длина окружности и площадь круга ===")
    r = float(input("Введите радиус круга R: "))
    
    # Вычисление длины окружности и площади круга
    length = 2.0 * PI * r
    area = PI * (r ** 2)
    
    print(f"Длина окружности L = {length}")
    print(f"Площадь круга S = {area}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 8,
        "title": "Среднее арифметическое двух чисел",
        "task_desc": "Даны два числа a и b. Найти их среднее арифметическое: (a + b) / 2.",
        "file": "task_io_08.py",
        "img": "task_io_08.png",
        "math_formula": "\\text{avg} = \\frac{a + b}{2}",
        "explanation": (
            "Запрашиваются вещественные числа a и b. "
            "Вычисляется их сумма, которая затем делится на 2: (a + b) / 2.0. "
            "Использование скобок обязательно, иначе деление выполнилось бы раньше сложения."
        ),
        "code": '''def main():
    print("=== Задание 8. Среднее арифметическое двух чисел ===")
    a = float(input("Введите первое число a: "))
    b = float(input("Введите второе число b: "))
    
    # Вычисление среднего арифметического
    arithmetic_mean = (a + b) / 2.0
    
    print(f"Среднее арифметическое = {arithmetic_mean}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 9,
        "title": "Среднее геометрическое двух неотрицательных чисел",
        "task_desc": "Даны два неотрицательных числа a и b. Найти их среднее геометрическое, то есть квадратный корень из их произведения: √(a × b).",
        "file": "task_io_09.py",
        "img": "task_io_09.png",
        "math_formula": "G = \\sqrt{a \\times b} = (a \\times b)^{0.5}",
        "explanation": (
            "Считываются числа a и b. Предусмотрена проверка условия неотрицательности: если хотя бы одно число отрицательное, "
            "программа сообщает об ошибке. "
            "Квадратный корень вычисляется возведением произведения в дробную степень 0.5: (a * b) ** 0.5."
        ),
        "code": '''def main():
    print("=== Задание 9. Среднее геометрическое двух неотрицательных чисел ===")
    a = float(input("Введите первое неотрицательное число a: "))
    b = float(input("Введите второе неотрицательное число b: "))
    
    # Проверка на неотрицательность в соответствии с условием
    if a < 0 or b < 0:
        print("Ошибка: числа должны быть неотрицательными!")
        return
    
    # Вычисление среднего геометрического: корень из произведения (степень 0.5)
    geometric_mean = (a * b) ** 0.5
    
    print(f"Среднее геометрическое = {geometric_mean}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 10,
        "title": "Операции над квадратами двух ненулевых чисел",
        "task_desc": "Даны два ненулевых числа. Найти сумму, разность, произведение и частное их квадратов.",
        "file": "task_io_10.py",
        "img": "task_io_10.png",
        "math_formula": "a^2 + b^2, \\quad a^2 - b^2, \\quad a^2 \\times b^2, \\quad \\frac{a^2}{b^2}",
        "explanation": (
            "Программа запрашивает два ненулевых вещественных числа. "
            "Сначала вычисляются квадраты обоих чисел: sq_a = a ** 2 и sq_b = b ** 2. "
            "Затем находятся сумма (sq_a + sq_b), разность (sq_a - sq_b), произведение (sq_a * sq_b) и "
            "частное (sq_a / sq_b). Поскольку числа ненулевые, b² > 0, деление на ноль исключено."
        ),
        "code": '''def main():
    print("=== Задание 10. Операции над квадратами двух ненулевых чисел ===")
    a = float(input("Введите первое ненулевое число a: "))
    b = float(input("Введите второе ненулевое число b: "))
    
    # Проверка на ненулевые значения
    if a == 0 or b == 0:
        print("Ошибка: числа должны быть ненулевыми!")
        return
    
    # Вычисление квадратов чисел
    sq_a = a ** 2
    sq_b = b ** 2
    
    # Вычисление требуемых величин
    sum_sq = sq_a + sq_b
    diff_sq = sq_a - sq_b
    prod_sq = sq_a * sq_b
    quot_sq = sq_a / sq_b
    
    print(f"Квадрат первого числа (a^2) = {sq_a}")
    print(f"Квадрат второго числа (b^2) = {sq_b}")
    print(f"Сумма квадратов = {sum_sq}")
    print(f"Разность квадратов = {diff_sq}")
    print(f"Произведение квадратов = {prod_sq}")
    print(f"Частное квадратов = {quot_sq}")

if __name__ == "__main__":
    main()'''
    }
]

TASKS_IF = [
    {
        "num": 1,
        "title": "Модификация положительного целого числа",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; в противном случае не изменять его. Вывести полученное число.",
        "file": "task_if_01.py",
        "img": "task_if_01.png",
        "math_formula": "n = n + 1 \\text{ при } n > 0",
        "explanation": (
            "Используется неполная форма условного оператора if n > 0:. "
            "Если условие истинно (число строго больше 0), выполняется составной оператор присваивания n += 1. "
            "В противном случае (если число отрицательно или равно нулю) ветка не выполняется, и число выводится без изменений."
        ),
        "code": '''def main():
    print("=== Задание 1 (Условный оператор). Модификация положительного числа ===")
    n = int(input("Введите целое число: "))
    
    # Если число положительное (больше нуля), прибавляем 1
    if n > 0:
        n += 1
    
    print(f"Полученное число: {n}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 2,
        "title": "Двухвариантное изменение целого числа",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; в противном случае вычесть из него 2. Вывести полученное число.",
        "file": "task_if_02.py",
        "img": "task_if_02.png",
        "math_formula": "n = \\begin{cases} n + 1, & n > 0 \\\\ n - 2, & n \\le 0 \\end{cases}",
        "explanation": (
            "Используется полная форма условного оператора if - else. "
            "Если выполняется условие n > 0, к переменной прибавляется 1 (n += 1). "
            "Иначе (блок else), когда n <= 0, из числа вычитается 2 (n -= 2)."
        ),
        "code": '''def main():
    print("=== Задание 2 (Условный оператор). Двухвариантное изменение числа ===")
    n = int(input("Введите целое число: "))
    
    # Если число положительное, прибавляем 1; иначе вычитаем 2
    if n > 0:
        n += 1
    else:
        n -= 2
    
    print(f"Полученное число: {n}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 3,
        "title": "Трёхвариантное ветвление",
        "task_desc": "Дано целое число. Если оно является положительным, то прибавить к нему 1; если отрицательным, то вычесть из него 2; если нулевым, то заменить его на 10. Вывести полученное число.",
        "file": "task_if_03.py",
        "img": "task_if_03.png",
        "math_formula": "n = \\begin{cases} n + 1, & n > 0 \\\\ n - 2, & n < 0 \\\\ 10, & n = 0 \\end{cases}",
        "explanation": (
            "Для реализации трех взаимоисключающих исходов применяется каскадная конструкция if - elif - else. "
            "Первая проверка n > 0 увеличивает число на 1. "
            "Ветка elif n < 0 уменьшает число на 2. "
            "Финальная ветка else срабатывает только при равенстве числа нулю, заменяя его значение на 10 (n = 10)."
        ),
        "code": '''def main():
    print("=== Задание 3 (Условный оператор). Трёхвариантное ветвление ===")
    n = int(input("Введите целое число: "))
    
    # Полная ветвящаяся структура if - elif - else
    if n > 0:
        n += 1
    elif n < 0:
        n -= 2
    else:
        n = 10
    
    print(f"Полученное число: {n}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 4,
        "title": "Подсчет количества положительных чисел в наборе из трех чисел",
        "task_desc": "Даны три целых числа. Найти количество положительных чисел в исходном наборе.",
        "file": "task_if_04.py",
        "img": "task_if_04.png",
        "math_formula": "\\text{count} = [a > 0] + [b > 0] + [c > 0]",
        "explanation": (
            "Инициализируется переменная-счетчик positive_count = 0. "
            "Затем каждое из трех введенных чисел последовательно проверяется на строгое неравенство нулю: if x > 0: positive_count += 1. "
            "Такой алгоритм прост, нагляден и корректно учитывает любое сочетание знаков, включая нулевые значения."
        ),
        "code": '''def main():
    print("=== Задание 4 (Условный оператор). Количество положительных чисел ===")
    a = int(input("Введите первое целое число: "))
    b = int(input("Введите второе целое число: "))
    c = int(input("Введите третье целое число: "))
    
    # Счетчик положительных чисел
    positive_count = 0
    
    if a > 0:
        positive_count += 1
    if b > 0:
        positive_count += 1
    if c > 0:
        positive_count += 1
    
    print(f"Количество положительных чисел: {positive_count}")

if __name__ == "__main__":
    main()'''
    },
    {
        "num": 5,
        "title": "Подсчет положительных и отрицательных чисел в наборе из трех чисел",
        "task_desc": "Даны три целых числа. Найти количество положительных и количество отрицательных чисел в исходном наборе.",
        "file": "task_if_05.py",
        "img": "task_if_05.png",
        "math_formula": "\\text{pos} = \\sum [x > 0], \\quad \\text{neg} = \\sum [x < 0]",
        "explanation": (
            "Вводятся два независимых счетчика: positive_count = 0 и negative_count = 0. "
            "Для каждого из трех чисел применяется конструкция if - elif: если число больше нуля, увеличивается первый счетчик, "
            "если меньше нуля — второй. Если число равно нулю, оно не учитывается ни в одной из групп."
        ),
        "code": '''def main():
    print("=== Задание 5 (Условный оператор). Подсчет положительных и отрицательных чисел ===")
    a = int(input("Введите первое целое число: "))
    b = int(input("Введите второе целое число: "))
    c = int(input("Введите третье целое число: "))
    
    # Инициализация счетчиков
    positive_count = 0
    negative_count = 0
    
    # Проверка каждого числа
    for num in (a, b, c):
        if num > 0:
            positive_count += 1
        elif num < 0:
            negative_count += 1
        
    print(f"Количество положительных чисел: {positive_count}")
    print(f"Количество отрицательных чисел: {negative_count}")

if __name__ == "__main__":
    main()'''
    }
]

def build_document():
    doc = Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 20 mm
        section.bottom_margin = Inches(0.79)  # 20 mm
        section.left_margin = Inches(0.98)    # 25 mm
        section.right_margin = Inches(0.59)   # 15 mm
    
    # Style standard
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
    
    # ------------------ ВВЕДЕНИЕ И ССЫЛКИ ------------------
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
        "«Ввод и вывод данных, оператор присваивания» (задания 1–10) и 5 заданий из темы «Условный оператор» (задания 1–5). "
        "Все программы разработаны в полном соответствии с лекционным курсом кафедры математики и бизнес-информатики "
        "Самарского университета."
    )
    
    p_links_head = doc.add_paragraph()
    r_lh = p_links_head.add_run("Официальные ссылки на репозиторий и интерактивное решение:")
    r_lh.bold = True
    
    p_link1 = doc.add_paragraph(style='List Bullet')
    p_link1.add_run("GitHub-репозиторий со всеми скриптами: ")
    add_hyperlink(p_link1, GITHUB_REPO, GITHUB_REPO)
    
    p_link2 = doc.add_paragraph(style='List Bullet')
    p_link2.add_run("Интерактивный запуск всех заданий в Google Colab: ")
    add_hyperlink(p_link2, COLAB_URL, "Открыть interactive_lab.ipynb в Google Colab")
    
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
        
        # Code
        p_code_lbl = doc.add_paragraph()
        r_cl = p_code_lbl.add_run("Исходный код программы:")
        r_cl.italic = True
        add_code_block(doc, task["code"])
        
        # Screenshot
        img_path = os.path.join(SCREENSHOTS_DIR, task["img"])
        if os.path.exists(img_path):
            p_img_lbl = doc.add_paragraph()
            r_il = p_img_lbl.add_run("Результат выполнения программы (скриншот консоли):")
            r_il.italic = True
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(6)
            doc.add_picture(img_path, width=Inches(5.9))
        
        # Explanation
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.first_line_indent = Inches(0.35)
        r_el = p_exp.add_run("Разбор решения: ")
        r_el.bold = True
        p_exp.add_run(task["explanation"])
        
        doc.add_paragraph() # divider
    
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
        
        # Code
        p_code_lbl = doc.add_paragraph()
        r_cl = p_code_lbl.add_run("Исходный код программы:")
        r_cl.italic = True
        add_code_block(doc, task["code"])
        
        # Screenshot
        img_path = os.path.join(SCREENSHOTS_DIR, task["img"])
        if os.path.exists(img_path):
            p_img_lbl = doc.add_paragraph()
            r_il = p_img_lbl.add_run("Результат выполнения программы (скриншот консоли):")
            r_il.italic = True
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(6)
            doc.add_picture(img_path, width=Inches(5.9))
        
        # Explanation
        p_exp = doc.add_paragraph()
        p_exp.paragraph_format.first_line_indent = Inches(0.35)
        r_el = p_exp.add_run("Разбор решения: ")
        r_el.bold = True
        p_exp.add_run(task["explanation"])
        
        doc.add_paragraph() # divider
        
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
        "В ходе выполнения работы были полностью реализованы и протестированы все поставленные задачи:\n"
        "1. Освоены фундаментальные операции ввода и вывода данных на Python: функция input() для интерактивного получения строк, "
        "явное преобразование типов float() и int(), форматированный вывод результатов с помощью f-строк.\n"
        "2. Закреплены базовые арифметические операторы языка: сложение (+), вычитание (-), умножение (*), деление (/), "
        "возведение в степень (**), а также составные операторы присваивания (+=, -=).\n"
        "3. Изучены механизмы алгоритмического ветвления: условный оператор if, альтернативная ветвь else и каскадная "
        "конструкция elif, позволяющая реализовывать многовариантную логику без вложенных конструкций.\n"
        "4. Все решения опубликованы в открытом доступе на платформе GitHub с возможностью мгновенного запуска в Google Colab."
    )
    
    doc.add_paragraph()
    p_tbl_lbl = doc.add_paragraph()
    r_tbl_lbl = p_tbl_lbl.add_run("Сводная таблица ссылок на выполненные задания:")
    r_tbl_lbl.bold = True
    
    # Table of all links
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    hdr_cells = tbl.rows[0].cells
    hdr_cells[0].width = Inches(0.8)
    hdr_cells[1].width = Inches(2.9)
    hdr_cells[2].width = Inches(2.5)
    
    headers = ["№", "Наименование задания", "Прямая ссылка на GitHub"]
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
        c1.width = Inches(2.9)
        c2.width = Inches(2.5)
        
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
    print(f"Отчет успешно сохранен в: {OUT_DOCX}")
    
    # Also save inside lab1_python_intro directory
    second_path = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\Отчет_Введение_в_программирование_Кириакиди.docx"
    doc.save(second_path)
    print(f"Копия отчета сохранена в: {second_path}")

if __name__ == "__main__":
    build_document()
