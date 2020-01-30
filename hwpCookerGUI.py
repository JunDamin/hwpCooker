import os
import time
import PySimpleGUI as sg
import pandas as pd
import xlwings as xw
from hwpCookerModel import generate_hml
import hwpTool as ht


"""

아래는 GUI다.

"""


sg.change_look_and_feel("TanBlue")

layout = [
    [sg.Text("자동입력용 데이터 선택")],
    [sg.Input(size=(70, 5)), sg.FileBrowse(button_text="데이터선택")],
    [sg.Button(button_text="데이터열기", key="OpenData")],
    [sg.Text("한글 탬플릿 선택")],
    [sg.Input(size=(70, 5)), sg.FileBrowse(button_text="탬플릿선택")],
    [sg.Button(button_text="탬플릿열기", key="OpenTemp")],
    [
        sg.InputCombo(("소수점없음", "첫째 자리", "둘째 자리"), size=(20, 1), default_value="소수점없음"),
        sg.Checkbox("천단위 구분", size=(10, 1), default=True),
    ],
    [sg.Output(size=(80, 20))],
    [
        sg.Button(button_text="굽기", key="OK"),
        sg.Button(button_text="종료", key="Cancel"),
        sg.Button("테스트페이지", key="Test"),
    ],
]

window = sg.Window("빵틀 v.0.2", icon="icon\\email.ico", layout=layout)

openText = """
======================================================================
데이터 파일과 탬플릿 파일을 고르고 굽기를 선택하면 엑셀 파일의 데이터를 한글 탬플릿에 입력해 줍니다.

환경

1. 아래아한글이 설치되어 있어야 합니다.
2. 엑셀 첫번째 행에 "파일명"이라는 값이 있어야 합니다.

주요
1. 엑셀 첫번째 행 값을 탬플릿에서 찾아 바꾸어줍니다.
2. 소수점 자리수를 조정할 수 있습니다.
3. 천단위 콤마를 고를 수 있습니다.
======================================================================
"""

endText = """
======================================================================
{}개 생성 작업이 완료 되었습니다.
======================================================================
"""


window.read(timeout=10)
print(openText)

while True:
    event, values = window.read()
    if event in (None, "Cancel"):
        break

    if event == "OpenTemp":
        ht.open_hwp_file(values[1])

    if event == "OpenData":
        xw.Book(values[0])

    if event == "Test":

        """ 기능을 리팩토링해야 함. 그리고 해당 파일을 열 수 있도록 코드도 정리해야 함"""
        if os.path.splitext(values[0])[1] in [".xls", ".xlsx"]:
            data = pd.read_excel(values[0])
        else:
            print("오류: 데이터 파일은 엑셀형식만 가능합니다.")
            continue

        # 변환자료 생성 위치
        output_path = ht.check_output_path(os.path.split(values[0])[0])

        # 자리수 입력 변환
        digit_check = list(("소수점없음", "첫째 자리", "둘째 자리"))
        digit = digit_check.index(values[2])

        if os.path.splitext(values[1])[1] in [".hwp"]:
            hml_address = ht.convert_to_hml(values[1])
        else:
            print("오류: 탬플릿 파일은 hwp형식만 가능합니다.")
            continue

        file_addr_list = generate_hml(
            hml_address,
            data.iloc[:1],
            output_path,
            name="파일명",
            belowZeroDigit=digit,
            thousand=values[3],
        )
        time.sleep(1)
        ht.convert_to_hwps(output_path)
        os.remove(hml_address)

        ht.open_hwp_file(file_addr_list[0])

    if event == "OK":
        if os.path.splitext(values[0])[1] in [".xls", ".xlsx"]:
            data = pd.read_excel(values[0])
        else:
            print("오류: 데이터 파일은 엑셀형식만 가능합니다.")
            continue

        # 변환자료 생성 위치
        output_path = ht.check_output_path(os.path.split(values[0])[0])

        # 자리수 입력 변환
        digit_check = list(("소수점없음", "첫째 자리", "둘째 자리"))
        digit = digit_check.index(values[2])

        if os.path.splitext(values[1])[1] in [".hwp"]:
            hml_address = ht.convert_to_hml(values[1])
        else:
            print("오류: 탬플릿 파일은 hwp형식만 가능합니다.")
            continue

        generate_hml(
            hml_address,
            data,
            output_path,
            name="파일명",
            belowZeroDigit=digit,
            thousand=values[3],
        )
        time.sleep(1)
        ht.convert_to_hwps(output_path)
        os.remove(hml_address)
        print(endText.format(len(data)))

window.close()
