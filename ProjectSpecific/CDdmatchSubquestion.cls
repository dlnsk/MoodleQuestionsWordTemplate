VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDdmatchSubquestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing subquestion of moodle drag-n-drop matching question
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pSubquestion As CHTML
Private pAnswer As CHTML

Private Sub Class_Initialize()
    Set pSubquestion = New CHTML
    Set pAnswer = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pSubquestion = Nothing
    Set pAnswer = Nothing
End Sub

Public Property Get Subquestion() As CHTML
    Set Subquestion = pSubquestion
End Property

Public Property Let Subquestion(ByVal NewValue As CHTML)
    Set pSubquestion = NewValue
End Property

Public Property Get Answer() As CHTML
    Set Answer = pAnswer
End Property

Public Property Set Answer(ByVal NewValue As CHTML)
    Set pAnswer = NewValue
End Property

Public Sub Export(ByRef QuestionElement As MSXML2.IXMLDOMElement)
    Dim XMLSubquestionElement As IXMLDOMElement
    Dim AnswerText As String
    Dim MD5 As String
    
    pSubquestion.Export QuestionElement, "subquestion", XMLSubquestionElement
    ' if DDMatch answer contains picture we should also invisibilly save files hash, so that same pictures with same name look as same answer and different pictures with same name look as different answers
    If pAnswer.Files.Count > 0 Then
        MD5 = pAnswer.Files.GetMD5
        AnswerText = Answer.Text
        Answer.Text = Answer.Text & "<!--" & MD5 & "-->"
        pAnswer.Export XMLSubquestionElement, "answer"
        Answer.Text = AnswerText
    Else
        pAnswer.Export XMLSubquestionElement, "answer"
    End If
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim RegExp As Object
    
    Set pSubquestion = New CHTML
    pSubquestion.Import Element
    
    Set pAnswer = New CHTML
    pAnswer.Import Element.SelectSingleNode("answer")
    
    If pAnswer.Files.Count > 0 Then
        Set RegExp = CreateObject("VBScript.RegExp")
        RegExp.Global = True
        RegExp.MultiLine = True
        RegExp.Pattern = "<!--[0123456789ABCDEF]{32}-->$"
        pAnswer.Text = RegExp.Replace(pAnswer.Text, "")
    End If
End Sub
