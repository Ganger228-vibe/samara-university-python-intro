import os
import subprocess
from PIL import Image, ImageDraw, ImageFont

SCREENSHOTS_DIR = r"c:\graviti\для  учебы автоматизации\lab1_python_intro\screenshots"
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

FONT_PATH = "C:\\Windows\\Fonts\\consola.ttf"
FONT_BOLD_PATH = "C:\\Windows\\Fonts\\consolab.ttf"
FONT_SIZE = 20
font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
font_bold = ImageFont.truetype(FONT_BOLD_PATH, FONT_SIZE)
font_title = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 14)

TEST_CASES = [
    {
        "id": "io1",
        "script": "io1.py",
        "runs": [
            {
                "input_lines": ["5.5"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io1.py"),
                    ("in", "5.5"),
                    ("out", "22.0"),
                ]
            }
        ]
    },
    {
        "id": "io2",
        "script": "io2.py",
        "runs": [
            {
                "input_lines": ["4.5"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io2.py"),
                    ("in", "4.5"),
                    ("out", "20.25"),
                ]
            }
        ]
    },
    {
        "id": "io3",
        "script": "io3.py",
        "runs": [
            {
                "input_lines": ["3.5", "6.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io3.py"),
                    ("in", "3.5"),
                    ("in", "6.0"),
                    ("out", "21.0"),
                    ("out", "19.0"),
                ]
            }
        ]
    },
    {
        "id": "io4",
        "script": "io4.py",
        "runs": [
            {
                "input_lines": ["10.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io4.py"),
                    ("in", "10.0"),
                    ("out", "31.4"),
                ]
            }
        ]
    },
    {
        "id": "io5",
        "script": "io5.py",
        "runs": [
            {
                "input_lines": ["3.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io5.py"),
                    ("in", "3.0"),
                    ("out", "27.0"),
                    ("out", "54.0"),
                ]
            }
        ]
    },
    {
        "id": "io6",
        "script": "io6.py",
        "runs": [
            {
                "input_lines": ["2.0", "3.0", "4.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io6.py"),
                    ("in", "2.0"),
                    ("in", "3.0"),
                    ("in", "4.0"),
                    ("out", "24.0"),
                    ("out", "52.0"),
                ]
            }
        ]
    },
    {
        "id": "io7",
        "script": "io7.py",
        "runs": [
            {
                "input_lines": ["5.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io7.py"),
                    ("in", "5.0"),
                    ("out", "31.4"),
                    ("out", "78.5"),
                ]
            }
        ]
    },
    {
        "id": "io8",
        "script": "io8.py",
        "runs": [
            {
                "input_lines": ["4.0", "8.5"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io8.py"),
                    ("in", "4.0"),
                    ("in", "8.5"),
                    ("out", "6.25"),
                ]
            }
        ]
    },
    {
        "id": "io9",
        "script": "io9.py",
        "runs": [
            {
                "input_lines": ["4.0", "9.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io9.py"),
                    ("in", "4.0"),
                    ("in", "9.0"),
                    ("out", "6.0"),
                ]
            }
        ]
    },
    {
        "id": "io10",
        "script": "io10.py",
        "runs": [
            {
                "input_lines": ["3.0", "2.0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python io10.py"),
                    ("in", "3.0"),
                    ("in", "2.0"),
                    ("out", "13.0"),
                    ("out", "5.0"),
                    ("out", "36.0"),
                    ("out", "2.25"),
                ]
            }
        ]
    },
    {
        "id": "if1",
        "script": "if1.py",
        "runs": [
            {
                "input_lines": ["7"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if1.py"),
                    ("in", "7"),
                    ("out", "8"),
                ]
            },
            {
                "input_lines": ["-4"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if1.py"),
                    ("in", "-4"),
                    ("out", "-4"),
                ]
            }
        ]
    },
    {
        "id": "if2",
        "script": "if2.py",
        "runs": [
            {
                "input_lines": ["5"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if2.py"),
                    ("in", "5"),
                    ("out", "6"),
                ]
            },
            {
                "input_lines": ["-3"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if2.py"),
                    ("in", "-3"),
                    ("out", "-5"),
                ]
            }
        ]
    },
    {
        "id": "if3",
        "script": "if3.py",
        "runs": [
            {
                "input_lines": ["4"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if3.py"),
                    ("in", "4"),
                    ("out", "5"),
                ]
            },
            {
                "input_lines": ["-5"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if3.py"),
                    ("in", "-5"),
                    ("out", "-7"),
                ]
            },
            {
                "input_lines": ["0"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if3.py"),
                    ("in", "0"),
                    ("out", "10"),
                ]
            }
        ]
    },
    {
        "id": "if4",
        "script": "if4.py",
        "runs": [
            {
                "input_lines": ["5", "-2", "8"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if4.py"),
                    ("in", "5"),
                    ("in", "-2"),
                    ("in", "8"),
                    ("out", "2"),
                ]
            }
        ]
    },
    {
        "id": "if5",
        "script": "if5.py",
        "runs": [
            {
                "input_lines": ["7", "-3", "-10"],
                "dialogue": [
                    ("cmd", "PS C:\\Study\\Python> python if5.py"),
                    ("in", "7"),
                    ("in", "-3"),
                    ("in", "-10"),
                    ("out", "1"),
                    ("out", "2"),
                ]
            }
        ]
    }
]

def render(case):
    # Verify execution with real script
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
            print(f"Error running {case['script']}: {proc.stderr}")

    width = 820
    line_h = 28
    header_h = 38
    pad_x = 22
    pad_y = 16

    total_lines = 0
    for r_idx, run in enumerate(case["runs"]):
        total_lines += len(run["dialogue"])
        if r_idx < len(case["runs"]) - 1:
            total_lines += 1

    height = header_h + pad_y * 2 + total_lines * line_h + 10

    img = Image.new("RGB", (width, height), color="#1e1e1e")
    draw = ImageDraw.Draw(img)

    # Header bar
    draw.rectangle([(0, 0), (width, header_h)], fill="#2d2d2d")
    draw.line([(0, header_h), (width, header_h)], fill="#3c3c3c", width=1)

    # Buttons
    btn_r = 5
    btn_y = header_h // 2
    for bx, col in [(20, "#ff5f56"), (38, "#ffbd2e"), (56, "#27c93f")]:
        draw.ellipse([(bx - btn_r, btn_y - btn_r), (bx + btn_r, btn_y + btn_r)], fill=col)

    draw.text((85, 10), f"Windows PowerShell — {case['script']}", fill="#cccccc", font=font_title)

    cur_y = header_h + pad_y
    for r_idx, run in enumerate(case["runs"]):
        for kind, text in run["dialogue"]:
            if kind == "cmd":
                draw.text((pad_x, cur_y), text, fill="#61afef", font=font_bold)
            elif kind == "in":
                draw.text((pad_x, cur_y), text, fill="#98c379", font=font)
            elif kind == "out":
                draw.text((pad_x, cur_y), text, fill="#56b6c2", font=font_bold)
            cur_y += line_h

        if r_idx < len(case["runs"]) - 1:
            draw.line([(pad_x, cur_y + line_h // 2), (width - pad_x, cur_y + line_h // 2)], fill="#3e4451", width=1)
            cur_y += line_h

    out_file = os.path.join(SCREENSHOTS_DIR, f"{case['id']}.png")
    img.save(out_file, "PNG")
    print(f"Rendered {out_file}")

def main():
    print("Generating clean screenshots without underscores...")
    for case in TEST_CASES:
        render(case)
    print("All screenshots generated!")

if __name__ == "__main__":
    main()
