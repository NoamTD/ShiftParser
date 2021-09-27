import os
import platform
import re
import subprocess
import typing
from pathlib import Path

import PySimpleGUI as sg
import pytesseract
from pyexcel_ods3 import save_data

sg.theme('SandyBeach')     

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sept", "oct", "nov", "dec"]


class Shift:
    def __init__(self, line: str):
        split = re.split(r"\s+|-|(?<=pm):", line.strip().lower())
        if len(split) != 4:
            raise Exception("bad line: " + line)
        self.month = to_month(split[0])
        self.day = split[1]
        self.start = split[2].lower()
        self.end = split[3].lower()


def to_month(month: str):
    matches = [m.capitalize() for m in months if month.lower().lstrip().startswith(m)]
    if len(matches) > 0:
        return matches[0]
    return ""


def to_suffix(day):
    if 4 <= day <= 20 or 24 <= day <= 30:
        suffix = "th"
    else:
        suffix = ["st", "nd", "rd"][day % 10 - 1]

    return suffix


def to_sheet(shifts: typing.List[Shift]):
    sheet = []
    for s in shifts:
        num = re.match(r"^\d+", s.day).group()
        day = num + to_suffix(int(num))
        sheet.append([s.month + " " + day, s.start, s.end])
    return sheet


def create_window():
    layout = [
        [sg.Text('Shifts Image'), sg.In(), sg.FileBrowse(key="file")],
        [sg.Text(key='output_file')],
        [sg.Button('Generate'), sg.Cancel()]
    ]
    window = sg.Window("Shift Table Generator", layout)
    return window


def main():
    window = create_window()
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
            
        window['output_file'].Update("")

        path = values['file']
        res: str = pytesseract.image_to_string(path)
        lines = [x for x in res.splitlines() if to_month(x) != ""]
        shifts = [Shift(line) for line in lines]
        sheet = {
            "data": to_sheet(shifts)
        }
        output_path = generate_output_path(path)
        save_data(output_path, sheet)
        window['output_file'].Update(f"Data written to: {output_path}")
        
    window.close()


def generate_output_path(path):
    return str(Path(path).parent.absolute().joinpath(Path(path).stem + '.ods'))


if __name__ == '__main__':
    main()
