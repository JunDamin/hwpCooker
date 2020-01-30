# -*- coding: utf-8 -*-
"""
HWP Api를 파이썬에서 쓰기 쉽도록 정리한 클래스이다. 

"""

import win32com.client as wc
import os


# 클래스 생성


class HwpApi(object):
    def __init__(self):
        self.client()

    def client(self, __value__=True):
        self.HWPAUTO = wc.gencache.EnsureDispatch("HWPFrame.HwpObject.1")

    # 함수정의하기

    # 끼워넣기 함수
    def HwpInsertFileFuction(self, InsertFilePath):
        HwpInsertFile = self.HWPAUTO.CreateAction("InsertFile")  # 끼어넣기 액션을 생성
        HwpInsertFilePara = HwpInsertFile.CreateSet()  # 끼어넣기 액션에 대한 파라미터 세트 생성
        HwpInsertFile.GetDefault(HwpInsertFilePara)  # 끼어넣기 액션에 대한 디폴트 파라미터 세트 지정
        HwpInsertFilePara.SetItem(
            "FileName", InsertFilePath
        )  # 파라미터 세트에 어떤 파라미터에 어떤 데이터를 넣는지 지정
        HwpInsertFilePara.SetItem("KeepSection", 1)
        HwpInsertFilePara.SetItem("KeepCharshape", 1)
        HwpInsertFilePara.SetItem("KeepParashape", 1)
        HwpInsertFilePara.SetItem("KeepStyle", 1)
        return HwpInsertFile.Execute(HwpInsertFilePara)  # 실행

    # 붙여넣기 함수
    def hwpSave(self, file_address, save_ext=".hwp"):
        """설명을 달아야지"""
        format_set = {".hwp": "HWP", ".hml": "HWPML2X"}
        file_set = os.path.splitext(file_address)
        file_path = os.path.split(file_address)[0]
        # 파일 저장하기
        save_file_address = os.path.join(file_path, file_set[0] + save_ext)
        HwpFileSave = self.HWPAUTO.CreateAction("FileSave")
        HwpFileSavePara = HwpFileSave.CreateSet()
        HwpFileSave.GetDefault(HwpFileSavePara)
        HwpFileSavePara.SetItem("FileName", save_file_address)
        HwpFileSavePara.SetItem("Format", format_set[save_ext])
        HwpFileSave.Execute(HwpFileSavePara)
        # 파일 닫기
        self.HWPAUTO.Run("FileClose")
        return save_file_address

    def hwpOpen(self, file_address):
        format_set = {".hwp": "HWP", ".hml": "HWPML2X"}
        file_path = os.path.split(file_address)[0]
        file_set = os.path.splitext(file_address)
        if not file_path:
            file_path = os.getcwd()
            file_address = os.path.join(file_path, file_address)
        HwpFileOpen = self.HWPAUTO.CreateAction("FileOpen")
        HwpFileOpenPara = HwpFileOpen.CreateSet()
        HwpFileOpen.GetDefault(HwpFileOpenPara)
        HwpFileOpenPara.SetItem("FileName", file_address)
        HwpFileOpenPara.SetItem("Format", format_set[file_set[1]])
        HwpFileOpen.Execute(HwpFileOpenPara)

    # 찾아 바꾸기 함수

    def hwpFindReplace(self, old_string, new_string):
        HwpFindReplace = self.HWPAUTO.CreateAction("AllReplace")  # 끼어넣기 액션을 생성
        HwpFindReplacePara = HwpFindReplace.CreateSet()  # 끼어넣기 액션에 대한 파라미터 세트 생성
        HwpFindReplace.GetDefault(HwpFindReplacePara)  # 끼어넣기 액션에 대한 디폴트 파라미터 세트 지정
        HwpFindReplacePara.SetItem("MatchCase", 0)
        HwpFindReplacePara.SetItem("AllWordForms", 0)
        HwpFindReplacePara.SetItem("SeveralWords", 0)
        HwpFindReplacePara.SetItem("UseWildCards", 0)
        HwpFindReplacePara.SetItem("WholeWordOnly", 0)
        HwpFindReplacePara.SetItem("AutoSpell", 1)  # 파라미터 세트에 어떤 파라미터에 어떤 데이터를 넣는지 지정
        HwpFindReplacePara.SetItem("Direction", 0)
        HwpFindReplacePara.SetItem("IgnoreFindString", 0)
        HwpFindReplacePara.SetItem("IgnoreReplaceString", 0)
        HwpFindReplacePara.SetItem("FindString", old_string)
        HwpFindReplacePara.SetItem("ReplaceString", new_string)
        HwpFindReplacePara.SetItem("ReplaceMode", 1)
        HwpFindReplacePara.SetItem("IgnoreMessage", 1)
        HwpFindReplacePara.SetItem("HanjaFromHangul", 0)
        HwpFindReplacePara.SetItem("FindJaso", 0)
        HwpFindReplacePara.SetItem("FindRegExp", 0)
        HwpFindReplacePara.SetItem("FindStyle", "")
        HwpFindReplacePara.SetItem("ReplaceStyle", "")
        HwpFindReplacePara.SetItem("FindType", 1)
        return HwpFindReplace.Execute(HwpFindReplacePara)  # 실행

    def hwpFileClose(self):
        self.HWPAUTO.Run("FileClose")

    def hwpQuit(self):
        self.HWPAUTO.Quit()

    def hwpSaveAs(self, file_address, save_ext=".hwp"):
        """설명을 달아야지"""
        # 기본 세팅
        format_set = {".hwp": "HWP", ".hml": "HWPML2X"}
        file_set = os.path.splitext(file_address)
        file_path = os.path.split(file_address)[0]
        if not file_path:
            file_path = os.getcwd()
        # 파일 저장하기
        save_file_address = os.path.join(file_path, file_set[0] + save_ext)
        HwpFileSaveAs = self.HWPAUTO.CreateAction("FileSaveAs")
        HwpFileSaveAsPara = HwpFileSaveAs.CreateSet()
        HwpFileSaveAs.GetDefault(HwpFileSaveAsPara)
        HwpFileSaveAsPara.SetItem("FileName", save_file_address)
        HwpFileSaveAsPara.SetItem("Format", format_set[save_ext])
        HwpFileSaveAs.Execute(HwpFileSaveAsPara)
        # 파일 닫기
        return save_file_address

    def Goto(self, n):
        """설명을 달아야지"""
        HwpGoto = self.HWPAUTO.CreateAction("Goto")
        HwpGotoPara = HwpGoto.CreateSet()
        HwpGoto.GetDefault(HwpGotoPara)
        HwpGotoPara.SetItem("DialogResult", n)
        HwpGotoPara.SetItem("SetSelectionIndex", 2)
        return HwpGoto.Execute(HwpGotoPara)

    def Delete(self):
        return self.HWPAUTO.Run("Delete")

    def MoveSelTopLevelBegin(self):
        return self.HWPAUTO.Run("MoveSelTopLevelBegin")

    def hwpSaveAs_S(self, file_address, save_ext=".hwp"):
        """설명을 달아야지"""
        # 기본 세팅
        format_set = {".hwp": "HWP", ".hml": "HWPML2X"}
        file_set = os.path.splitext(file_address)
        file_path = os.path.split(file_address)[0]

        # 파일 저장하기
        save_file_address = os.path.join(file_path, file_set[0] + save_ext)
        HwpFileSaveAs_S = self.HWPAUTO.CreateAction("FileSaveAs_S")
        HwpFileSaveAs_SPara = HwpFileSaveAs_S.CreateSet()
        HwpFileSaveAs_S.GetDefault(HwpFileSaveAs_SPara)
        HwpFileSaveAs_SPara.SetItem("FileName", save_file_address)
        HwpFileSaveAs_SPara.SetItem("Format", format_set[save_ext])
        HwpFileSaveAs_S.Execute(HwpFileSaveAs_SPara)
        return save_file_address
