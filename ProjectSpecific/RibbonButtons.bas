Attribute VB_Name = "RibbonButtons"
' ������������ ������� �� ������ �����
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons �Attribution-ShareAlike� 4.0

Option Explicit

' ���������� ������� � ��������
Public Sub GetMoodle2WordTabLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonTabLabel
End Sub
' ���������� ������� � ������
Public Sub GetMoodle2WordGroupLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonGroupLabel
End Sub
' ���������� ������� � ������
Public Sub GetMoodle2WordExecLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonImportLabel
End Sub
' ���������� ������� � ������
Public Sub GetMoodle2WordAboutLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonAboutLabel
End Sub
' ���������� ������� ������ �����
Public Sub Moodle2WordExec(ByVal Control As IRibbonControl)
    Moodle2Word.Moodle2Word
End Sub
' ���������� ������� ������ �����
Public Sub Moodle2WordShowAboutDlg(ByVal Control As IRibbonControl)
    AboutForm.Show
End Sub

