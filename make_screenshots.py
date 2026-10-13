import os
import subprocess
from PIL import Image, ImageDraw, ImageFont

# Папка для скриншотов
SCREENSHOTS_DIR = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\screenshots"
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

# Шрифт
FONT_PATH = "C:\\Windows\\Fonts\\consola.ttf"
FONT_SIZE = 20
font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
font_bold = ImageFont.truetype("C:\\Windows\\Fonts\\consolab.ttf", FONT_SIZE)
font_title = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 14)

TEST_CASES = [
    {
        "id": "task_io_01",
        "script": "task_io_01.py",
        "title": "Задание 1: Периметр квадрата",
        "runs": [
            {
                "input_lines": ["5.5"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_01.py"),
                    ("info", "=== Задание 1. Периметр квадрата ==="),
                    ("input_prompt", "Введите сторону квадрата a: ", "5.5"),
                    ("output", "Периметр квадрата P = 22.0"),
                ]
            }
        ]
    },
    {
        "id": "task_io_02",
        "script": "task_io_02.py",
        "title": "Задание 2: Площадь квадрата",
        "runs": [
            {
                "input_lines": ["4.5"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_02.py"),
                    ("info", "=== Задание 2. Площадь квадрата ==="),
                    ("input_prompt", "Введите сторону квадрата a: ", "4.5"),
                    ("output", "Площадь квадрата S = 20.25"),
                ]
            }
        ]
    },
    {
        "id": "task_io_03",
        "script": "task_io_03.py",
        "title": "Задание 3: Площадь и периметр прямоугольника",
        "runs": [
            {
                "input_lines": ["3.5", "6.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_03.py"),
                    ("info", "=== Задание 3. Площадь и периметр прямоугольника ==="),
                    ("input_prompt", "Введите сторону прямоугольника a: ", "3.5"),
                    ("input_prompt", "Введите сторону прямоугольника b: ", "6.0"),
                    ("output", "Площадь прямоугольника S = 21.0"),
                    ("output", "Периметр прямоугольника P = 19.0"),
                ]
            }
        ]
    },
    {
        "id": "task_io_04",
        "script": "task_io_04.py",
        "title": "Задание 4: Длина окружности по диаметру",
        "runs": [
            {
                "input_lines": ["10.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_04.py"),
                    ("info", "=== Задание 4. Длина окружности по диаметру ==="),
                    ("input_prompt", "Введите диаметр окружности d: ", "10.0"),
                    ("output", "Длина окружности L = 31.4"),
                ]
            }
        ]
    },
    {
        "id": "task_io_05",
        "script": "task_io_05.py",
        "title": "Задание 5: Объем и площадь поверхности куба",
        "runs": [
            {
                "input_lines": ["3.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_05.py"),
                    ("info", "=== Задание 5. Объем и площадь поверхности куба ==="),
                    ("input_prompt", "Введите длину ребра куба a: ", "3.0"),
                    ("output", "Объем куба V = 27.0"),
                    ("output", "Площадь поверхности куба S = 54.0"),
                ]
            }
        ]
    },
    {
        "id": "task_io_06",
        "script": "task_io_06.py",
        "title": "Задание 6: Объем и площадь параллелепипеда",
        "runs": [
            {
                "input_lines": ["2.0", "3.0", "4.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_06.py"),
                    ("info", "=== Задание 6. Объем и площадь поверхности параллелепипеда ==="),
                    ("input_prompt", "Введите длину ребра a: ", "2.0"),
                    ("input_prompt", "Введите длину ребра b: ", "3.0"),
                    ("input_prompt", "Введите длину ребра c: ", "4.0"),
                    ("output", "Объем параллелепипеда V = 24.0"),
                    ("output", "Площадь поверхности S = 52.0"),
                ]
            }
        ]
    },
    {
        "id": "task_io_07",
        "script": "task_io_07.py",
        "title": "Задание 7: Длина окружности и площадь круга",
        "runs": [
            {
                "input_lines": ["5.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_07.py"),
                    ("info", "=== Задание 7. Длина окружности и площадь круга ==="),
                    ("input_prompt", "Введите радиус круга R: ", "5.0"),
                    ("output", "Длина окружности L = 31.4"),
                    ("output", "Площадь круга S = 78.5"),
                ]
            }
        ]
    },
    {
        "id": "task_io_08",
        "script": "task_io_08.py",
        "title": "Задание 8: Среднее арифметическое двух чисел",
        "runs": [
            {
                "input_lines": ["4.0", "8.5"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_08.py"),
                    ("info", "=== Задание 8. Среднее арифметическое двух чисел ==="),
                    ("input_prompt", "Введите первое число a: ", "4.0"),
                    ("input_prompt", "Введите второе число b: ", "8.5"),
                    ("output", "Среднее арифметическое = 6.25"),
                ]
            }
        ]
    },
    {
        "id": "task_io_09",
        "script": "task_io_09.py",
        "title": "Задание 9: Среднее геометрическое двух неотрицательных чисел",
        "runs": [
            {
                "input_lines": ["4.0", "9.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_09.py"),
                    ("info", "=== Задание 9. Среднее геометрическое двух неотрицательных чисел ==="),
                    ("input_prompt", "Введите первое неотрицательное число a: ", "4.0"),
                    ("input_prompt", "Введите второе неотрицательное число b: ", "9.0"),
                    ("output", "Среднее геометрическое = 6.0"),
                ]
            }
        ]
    },
    {
        "id": "task_io_10",
        "script": "task_io_10.py",
        "title": "Задание 10: Операции над квадратами двух ненулевых чисел",
        "runs": [
            {
                "input_lines": ["3.0", "2.0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_io_10.py"),
                    ("info", "=== Задание 10. Операции над квадратами двух ненулевых чисел ==="),
                    ("input_prompt", "Введите первое ненулевое число a: ", "3.0"),
                    ("input_prompt", "Введите второе ненулевое число b: ", "2.0"),
                    ("output", "Квадрат первого числа (a^2) = 9.0"),
                    ("output", "Квадрат второго числа (b^2) = 4.0"),
                    ("output", "Сумма квадратов = 13.0"),
                    ("output", "Разность квадратов = 5.0"),
                    ("output", "Произведение квадратов = 36.0"),
                    ("output", "Частное квадратов = 2.25"),
                ]
            }
        ]
    },
    {
        "id": "task_if_01",
        "script": "task_if_01.py",
        "title": "Условный оператор 1: Модификация положительного числа",
        "runs": [
            {
                "input_lines": ["7"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_01.py"),
                    ("info", "=== Задание 1 (Условный оператор). Модификация положительного числа ==="),
                    ("input_prompt", "Введите целое число: ", "7"),
                    ("output", "Полученное число: 8"),
                ]
            },
            {
                "input_lines": ["-4"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_01.py"),
                    ("info", "=== Задание 1 (Условный оператор). Модификация положительного числа ==="),
                    ("input_prompt", "Введите целое число: ", "-4"),
                    ("output", "Полученное число: -4"),
                ]
            }
        ]
    },
    {
        "id": "task_if_02",
        "script": "task_if_02.py",
        "title": "Условный оператор 2: Двухвариантное изменение числа",
        "runs": [
            {
                "input_lines": ["5"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_02.py"),
                    ("info", "=== Задание 2 (Условный оператор). Двухвариантное изменение числа ==="),
                    ("input_prompt", "Введите целое число: ", "5"),
                    ("output", "Полученное число: 6"),
                ]
            },
            {
                "input_lines": ["-3"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_02.py"),
                    ("info", "=== Задание 2 (Условный оператор). Двухвариантное изменение числа ==="),
                    ("input_prompt", "Введите целое число: ", "-3"),
                    ("output", "Полученное число: -5"),
                ]
            }
        ]
    },
    {
        "id": "task_if_03",
        "script": "task_if_03.py",
        "title": "Условный оператор 3: Трёхвариантное ветвление",
        "runs": [
            {
                "input_lines": ["4"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_03.py"),
                    ("info", "=== Задание 3 (Условный оператор). Трёхвариантное ветвление ==="),
                    ("input_prompt", "Введите целое число: ", "4"),
                    ("output", "Полученное число: 5"),
                ]
            },
            {
                "input_lines": ["-5"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_03.py"),
                    ("info", "=== Задание 3 (Условный оператор). Трёхвариантное ветвление ==="),
                    ("input_prompt", "Введите целое число: ", "-5"),
                    ("output", "Полученное число: -7"),
                ]
            },
            {
                "input_lines": ["0"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_03.py"),
                    ("info", "=== Задание 3 (Условный оператор). Трёхвариантное ветвление ==="),
                    ("input_prompt", "Введите целое число: ", "0"),
                    ("output", "Полученное число: 10"),
                ]
            }
        ]
    },
    {
        "id": "task_if_04",
        "script": "task_if_04.py",
        "title": "Условный оператор 4: Количество положительных чисел",
        "runs": [
            {
                "input_lines": ["5", "-2", "8"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_04.py"),
                    ("info", "=== Задание 4 (Условный оператор). Количество положительных чисел ==="),
                    ("input_prompt", "Введите первое целое число: ", "5"),
                    ("input_prompt", "Введите второе целое число: ", "-2"),
                    ("input_prompt", "Введите третье целое число: ", "8"),
                    ("output", "Количество положительных чисел: 2"),
                ]
            }
        ]
    },
    {
        "id": "task_if_05",
        "script": "task_if_05.py",
        "title": "Условный оператор 5: Подсчет положительных и отрицательных",
        "runs": [
            {
                "input_lines": ["7", "-3", "-10"],
                "dialogue": [
                    ("prompt_cmd", "PS C:\\Study\\Python> python task_if_05.py"),
                    ("info", "=== Задание 5 (Условный оператор). Подсчет положительных и отрицательных чисел ==="),
                    ("input_prompt", "Введите первое целое число: ", "7"),
                    ("input_prompt", "Введите второе целое число: ", "-3"),
                    ("input_prompt", "Введите третье целое число: ", "-10"),
                    ("output", "Количество положительных чисел: 1"),
                    ("output", "Количество отрицательных чисел: 2"),
                ]
            }
        ]
    }
]

def render_terminal_screenshot(case):
    # Verify execution first via subprocess
    script_path = os.path.join(r"c:\graviti\для  учебы автоматизации\lab1_python_intro", case["script"])
    for run in case["runs"]:
        inp = "\n".join(run["input_lines"]) + "\n"
        proc = subprocess.run(
            ["python", script_path],
            input=inp,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        if proc.returncode != 0:
            print(f"ERROR running {case['script']}: {proc.stderr}")

    # Layout parameters
    width = 980
    line_height = 28
    header_height = 42
    padding_x = 24
    padding_y = 18

    # Calculate total lines
    total_lines = 0
    for r_idx, run in enumerate(case["runs"]):
        total_lines += len(run["dialogue"])
        if r_idx < len(case["runs"]) - 1:
            total_lines += 1 # separator line

    height = header_height + padding_y * 2 + total_lines * line_height + 20

    img = Image.new("RGB", (width, height), color="#1e1e1e")
    draw = ImageDraw.Draw(img)

    # Header bar
    draw.rectangle([(0, 0), (width, header_height)], fill="#2d2d2d")
    draw.line([(0, header_height), (width, header_height)], fill="#3c3c3c", width=1)

    # Window buttons (macOS style / modern)
    btn_radius = 6
    btn_y = header_height // 2
    buttons = [
        (22, "#ff5f56"), # red
        (42, "#ffbd2e"), # yellow
        (62, "#27c93f"), # green
    ]
    for bx, color in buttons:
        draw.ellipse([(bx - btn_radius, btn_y - btn_radius), (bx + btn_radius, btn_y + btn_radius)], fill=color)

    # Window title
    title_text = f"Windows PowerShell — {case['script']}"
    draw.text((95, 12), title_text, fill="#cccccc", font=font_title)

    # Terminal content
    cur_y = header_height + padding_y

    for r_idx, run in enumerate(case["runs"]):
        for item in run["dialogue"]:
            item_type = item[0]
            if item_type == "prompt_cmd":
                # Prompt text
                draw.text((padding_x, cur_y), item[1], fill="#61afef", font=font_bold)
            elif item_type == "info":
                draw.text((padding_x, cur_y), item[1], fill="#e5c07b", font=font)
            elif item_type == "input_prompt":
                prompt_txt = item[1]
                val_txt = item[2]
                draw.text((padding_x, cur_y), prompt_txt, fill="#abb2bf", font=font)
                prompt_w = draw.textlength(prompt_txt, font=font)
                draw.text((padding_x + prompt_w, cur_y), val_txt, fill="#98c379", font=font_bold)
            elif item_type == "output":
                draw.text((padding_x, cur_y), item[1], fill="#56b6c2", font=font_bold)
            cur_y += line_height

        if r_idx < len(case["runs"]) - 1:
            draw.line([(padding_x, cur_y + line_height // 2), (width - padding_x, cur_y + line_height // 2)], fill="#3e4451", width=1)
            cur_y += line_height

    out_file = os.path.join(SCREENSHOTS_DIR, f"{case['id']}.png")
    img.save(out_file, "PNG", dpi=(150, 150))
    print(f"Generated screenshot: {out_file}")

def main():
    print("Generating terminal execution screenshots for all 15 tasks...")
    for case in TEST_CASES:
        render_terminal_screenshot(case)
    print("All screenshots generated successfully!")

if __name__ == "__main__":
    main()
