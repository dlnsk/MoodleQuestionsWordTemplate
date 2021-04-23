VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMatchingSubquestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing subquestion of moodle matching question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pSubquestion As CHTML
Private pAnswer As String

Private Sub Class_Initialize()
    Set pSubquestion = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pSubquestion = Nothing
End Sub

Public Property Get Subquestion() As CHTML
    Set Subquestion = pSubquestion
End Property

Public Property Let Subquestion(ByVal NewValue As CHTML)
    Set pSubquestion = NewValue
End Property

Public Property Get Answer() As String
    Answer = pAnswer
End Property

Public Property Let Answer(ByVal NewValue As String)
    pAnswer = NewValue
End Property

Public Sub Export(ByRef QuestionElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLSubquestionElement As IXMLDOMElement
    Dim XMLAnswerElement As IXMLDOMElement
    Dim XMLTextElement As IXMLDOMElement
    Dim XMLText As IXMLDOMText

    pSubquestion.Export QuestionElement, "subquestion", XMLSubquestionElement
    
    Set DOM = QuestionElement.OwnerDocument
    Set XMLAnswerElement = DOM.createElement("answer")
    XMLSubquestionElement.appendChild XMLAnswerElement
    Set XMLTextElement = DOM.createElement("text")
    XMLAnswerElement.appendChild XMLTextElement
    Set XMLText = DOM.createTextNode(pAnswer)
    XMLTextElement.appendChild XMLText
End Sub
