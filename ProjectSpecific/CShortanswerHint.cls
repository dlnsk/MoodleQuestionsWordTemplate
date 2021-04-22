VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShortanswerHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing hint in moodle shortanswer and numerical  question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pHint As CHTML

Private Sub Class_Initialize()
    Set pHint = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pHint = Nothing
End Sub

Public Property Get Text() As String
    Text = pHint.Text
End Property

Public Property Let Text(ByVal NewValue As String)
    pHint.Text = NewValue
End Property

Public Sub AddFile(Data As String, Filename As String)
    pHint.AddFile Data, Filename
End Sub

Public Sub Export(ByRef QuestionElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLHintElement As IXMLDOMElement
    Dim XMLElement As IXMLDOMElement
    
    Set DOM = QuestionElement.OwnerDocument
    
    pHint.Export QuestionElement, "hint", XMLHintElement
End Sub
