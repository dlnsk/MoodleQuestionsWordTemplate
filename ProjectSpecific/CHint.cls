VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing hint in most of moodle questions (multichoice, matching ddmatch, ordering)
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pHint As CHTML
Private pShownumcorrect As Boolean
Private pClearwrong As Boolean

Private Sub Class_Initialize()
    Set pHint = New CHTML
    pShownumcorrect = False
    pClearwrong = False
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

Public Property Get Shownumcorrect() As Boolean
    Shownumcorrect = pShownumcorrect
End Property

Public Property Let Shownumcorrect(ByVal NewValue As Boolean)
    pShownumcorrect = NewValue
End Property

Public Property Get Clearwrong() As Boolean
    Clearwrong = pClearwrong
End Property

Public Property Let Clearwrong(ByVal NewValue As Boolean)
    pClearwrong = NewValue
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
    
    If pShownumcorrect Then
        Set XMLElement = DOM.createElement("shownumcorrect")
        XMLHintElement.appendChild XMLElement
    End If
    
    If pClearwrong Then
        Set XMLElement = DOM.createElement("clearwrong")
        XMLHintElement.appendChild XMLElement
    End If
End Sub

