VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMultichoiceAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing answer of moodle multichoice question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pFraction As Double
Private pAnswer As CHTML
Private pFeedback As CHTML

Private Sub Class_Initialize()
    pFraction = 0
    Set pAnswer = New CHTML
    Set pFeedback = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pAnswer = Nothing
    Set pFeedback = Nothing
End Sub

Public Property Get Fraction() As Double
    Fraction = pFraction
End Property

Public Property Let Fraction(ByVal NewValue As Double)
    pFraction = NewValue
End Property

Public Property Get Answer() As CHTML
    Set Answer = pAnswer
End Property

Public Property Set Answer(ByVal NewValue As CHTML)
    Set pAnswer = NewValue
End Property

Public Property Get feedback() As CHTML
    Set feedback = pFeedback
End Property

Public Property Let feedback(ByVal NewValue As CHTML)
    Set pFeedback = NewValue
End Property

Public Sub Export(ByRef QuestionElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLAnswerElement As IXMLDOMElement
    Dim XMLAttr As IXMLDOMAttribute

    pAnswer.Export QuestionElement, "answer", XMLAnswerElement
    
    Set DOM = QuestionElement.OwnerDocument
    Set XMLAttr = DOM.createAttribute("fraction")
    XMLAttr.NodeValue = Replace(Format(pFraction, "###0.00000"), ",", ".")
    XMLAnswerElement.setAttributeNode XMLAttr
    
    pFeedback.Export XMLAnswerElement, "feedback"
End Sub
