VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShortanswerAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing answer of moodle shortanswer and truefalse questions
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pFraction As Double
Private pAnswer As String
Private pFeedback As CHTML

Private Sub Class_Initialize()
    pFraction = 100
    Set pFeedback = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pFeedback = Nothing
End Sub

Public Property Get Fraction() As Double
    Fraction = pFraction
End Property

Public Property Let Fraction(ByVal NewValue As Double)
    pFraction = NewValue
End Property

Public Property Get Text() As String
    Text = pAnswer
End Property

Public Property Let Text(ByVal NewValue As String)
    pAnswer = NewValue
End Property

Public Property Get Feedback() As CHTML
    Set Feedback = pFeedback
End Property

Public Property Let Feedback(ByVal NewValue As CHTML)
    Set pFeedback = NewValue
End Property

Public Sub Export(ByRef QuestionElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLElement1 As MSXML2.IXMLDOMElement
    Dim XMLElement2 As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute

    Set DOM = QuestionElement.OwnerDocument
    Set XMLElement1 = DOM.createElement("answer")
    QuestionElement.appendChild XMLElement1
    
    Set XMLAttr = DOM.createAttribute("fraction")
    XMLAttr.NodeValue = Replace(Format(pFraction, "###0.00000"), ",", ".")
    XMLElement1.setAttributeNode XMLAttr
    
    Set XMLAttr = DOM.createAttribute("format")
    XMLAttr.NodeValue = "moodle_auto_format"
    XMLElement1.setAttributeNode XMLAttr
    
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(pAnswer)
    XMLElement2.appendChild XMLText
    
    pFeedback.Export XMLElement1, "feedback"
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Separator As String
    Separator = Replace(Format(0, "0.0"), "0", "")
    
    pFraction = CDbl(Replace(Element.SelectSingleNode("@fraction").NodeValue, ".", Separator))
    
    pAnswer = Element.SelectSingleNode("text/text()").NodeValue
    
    Set pFeedback = New CHTML
    pFeedback.Import Element.SelectSingleNode("feedback")
End Sub

