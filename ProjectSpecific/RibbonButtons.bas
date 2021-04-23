Attribute VB_Name = "RibbonButtons"
' Обработчкики нажатий на кнопки ленты
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons «Attribution-ShareAlike» 4.0

Option Explicit

' Возвращает подпись к закладке
Public Sub GetMoodle2WordTabLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonTabLabel
End Sub
' Возвращает подпись к группе
Public Sub GetMoodle2WordGroupLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonGroupLabel
End Sub
' Возвращает подпись к кнопке
Public Sub GetMoodle2WordExecLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonImportLabel
End Sub
' Возвращает подпись к кнопке
Public Sub GetMoodle2WordAboutLabel(ByRef Control As IRibbonControl, ByRef returnVal)
    returnVal = strRibbonAboutLabel
End Sub
' Обработчик события кнопки ленты
Public Sub Moodle2WordExec(ByVal Control As IRibbonControl)
    Moodle2Word.Moodle2Word
End Sub
' Обработчки события кнопки ленты
Public Sub Moodle2WordShowAboutDlg(ByVal Control As IRibbonControl)
    AboutForm.Show
End Sub

