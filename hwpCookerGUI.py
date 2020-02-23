import os
import time
import PySimpleGUI as sg
import pandas as pd
from hwpCookerModel import generate_hml
import hwpTool as ht

"""
데이터를 입력받고 생성된 HML 파일을 HWP로 바꾼다.
"""


def validate_values(values, digit_check):

    """ 데이터 형식을 실행하기 전에 검토하는 기능임 """

    if os.path.splitext(values[0])[1] in [".xls", ".xlsx", "xlsm"]:
        data = pd.read_excel(values[0])
    else:
        print("오류: 데이터 파일은 엑셀형식만 가능합니다.")
        return None

    # 변환자료 생성 위치
    output_path = ht.check_output_path(os.path.split(values[0])[0])

    # 자리수 입력 변환

    digit = digit_check.index(values[2])

    if os.path.splitext(values[1])[1] in [".hwp"]:
        template_hml_addr = ht.convert_to_hml(values[1])
    else:
        print("오류: 탬플릿 파일은 hwp형식만 가능합니다.")
        return None

    return data, output_path, digit, template_hml_addr


"""

아래는 GUI구성

"""

digit_check = list(("소수점없음", "첫째 자리", "둘째 자리", "셋째 자리", "넷째 자리", "다섯째 자리"))

sg.change_look_and_feel("TanBlue")

layout = [
    [sg.Text("자동입력용 데이터 선택")],
    [sg.Input(size=(70, 5)), sg.FileBrowse(button_text="데이터선택")],
    [sg.Button(button_text="데이터열기", key="OpenData")],
    [sg.Text("한글 탬플릿 선택")],
    [sg.Input(size=(70, 5)), sg.FileBrowse(button_text="탬플릿선택")],
    [sg.Button(button_text="탬플릿열기", key="OpenTemp")],
    [
        sg.InputCombo((digit_check), size=(20, 1), default_value="소수점없음"),
        sg.Checkbox("천단위 구분", size=(10, 1), default=True),
    ],
    [sg.Output(size=(80, 20))],
    [sg.ProgressBar(1000, orientation="h", size=(48.5, 20), key="progbar")],
    [
        sg.Button(button_text="굽기", key="OK"),
        sg.Button("테스트페이지", key="Test"),
        sg.Button(button_text="결과물 폴더 열기", pad=((300, 0), 3), key="OutputFolder"),
        sg.Button(button_text="종료", pad=((5, 0), 3), key="Cancel"),
    ],
]


window = sg.Window("빵틀 v.0.2", icon="icon\\stove.ico", layout=layout)

openText = """
======================================================================
데이터 파일과 탬플릿 파일을 고르고 굽기를 선택하면 엑셀 파일의 데이터를 한글 탬플릿에 입력해 줍니다.

환경

1. 아래아한글이 설치되어 있어야 합니다.
2. 엑셀 첫번째 행에 "__파일명__"이라는 셀이 있어야 합니다.

주요
1. 엑셀 첫번째 시트의 첫번째 행 값을 탬플릿에서 찾아 바꾸어줍니다.
2. 소수점 자리수를 조정할 수 있습니다.
3. 천단위 콤마를 고를 수 있습니다.


주의사항
1. 꼭 결과물을 확인하시고 생성된 파일을 사용해주세요.
2. 파일명이 너무 길 경우 작동이 안될 수 있습니다.
======================================================================
"""

endText = """
======================================================================
{}개 생성 작업이 완료 되었습니다.
결과물에 이상이 없는지 꼭 확인 후 사용해주세요.
======================================================================
"""


window.read(timeout=10)
print(openText)
col_name = "__파일명__"

while True:
    event, values = window.read()
    if event in (None, "Cancel"):
        break

    if event == "OpenTemp":
        ht.open_hwp_file(values[1])

    if event == "OpenData":
        os.startfile(values[0])

    if event == "OutputFolder":
        output_path = ht.check_output_path(os.path.split(values[0])[0])
        if not output_path:
            print("파일을 먼저 지정해주세요.")
        else:
            os.startfile(output_path)

    if event == "Test":

        if validate_values(values, digit_check):
            data, output_path, digit, template_hml_addr = validate_values(
                values, digit_check
            )
            print("[테스트 페이지 인쇄]")

            hml_addr = generate_hml(
                template_hml_addr,
                data.iloc[0],
                output_path,
                name=col_name,
                belowZeroDigit=digit,
                thousand=values[3],
            )
            time.sleep(0.2)
            os.remove(template_hml_addr)
            # 출력하기
            print(endText.format(1))
            window.refresh()
            # 생성된 파일 열기
            ht.open_hwp_file(hml_addr)

    if event == "OK":

        if validate_values(values, digit_check):
            data, output_path, digit, template_hml_addr = validate_values(
                values, digit_check
            )

            file_num = len(data)
            progress = 0

            for i in data.index:
                hml_addr = generate_hml(
                    template_hml_addr,
                    data.loc[i],
                    output_path,
                    name=col_name,
                    belowZeroDigit=digit,
                    thousand=values[3],
                )
                window.refresh()
                progress += 1
                window["progbar"].update_bar(int(progress / file_num * 1000))
                ht.convert_to_hwp(hml_addr)
            # Template 삭제
            os.remove(template_hml_addr)
            print(endText.format(len(data)))

window.close()
