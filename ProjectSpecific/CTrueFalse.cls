VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTrueFalse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle multichoice question
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' Copyright 2015 Петр Молоков (PBMolokov@gmail.com)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements CQuestion

Private pName As String
Private pQuestiontext As CHTML
Private pGeneralfeedback As CHTML
Private pDefaultgrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pAnswerTrue As CShortanswerAnswer
Private pAnswerFalse As CShortanswerAnswer

Private Sub Class_Initialize()
    pName = "-"
    Set pQuestiontext = New CHTML
    Set pGeneralfeedback = New CHTML
    pDefaultgrade = 1
    pPenalty = 1
    pHidden = False
    Set pAnswerTrue = New CShortanswerAnswer
    pAnswerTrue.Text = "true"
    pAnswerTrue.Fraction = 100
    Set pAnswerFalse = New CShortanswerAnswer
    pAnswerFalse.Text = "false"
    pAnswerFalse.Fraction = 0
End Sub

Private Sub Class_Terminate()
    Set pQuestiontext = Nothing
    Set pGeneralfeedback = Nothing
    Set pAnswerTrue = Nothing
    Set pAnswerFalse = Nothing
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal NewValue As String)
    pName = NewValue
End Property

Public Property Get QuestionText() As CHTML
    Set QuestionText = pQuestiontext
End Property

Public Property Set QuestionText(ByVal NewValue As CHTML)
    Set pQuestiontext = NewValue
End Property

Public Property Get Generalfeedback() As CHTML
    Set Generalfeedback = pGeneralfeedback
End Property

Public Property Set Generalfeedback(ByVal NewValue As CHTML)
    Set pGeneralfeedback = NewValue
End Property

Public Property Get Defaultgrade() As Double
    Defaultgrade = pDefaultgrade
End Property

Public Property Let Defaultgrade(ByVal NewValue As Double)
    pDefaultgrade = NewValue
End Property

'Public Property Get Penalty() As Double
'    Penalty = pPenalty
'End Property
'
'Public Property Let Penalty(ByVal NewValue As Double)
'    pPenalty = NewValue
'End Property
'
'Public Property Get Hidden() As Boolean
'    Hidden = pHidden
'End Property
'
'Public Property Let Hidden(ByVal NewValue As Boolean)
'    pHidden = NewValue
'End Property

Public Property Get Answer() As Boolean
    If pAnswerTrue.Fraction = 100 Then
        Answer = True
    Else
        Answer = False
    End If
End Property

Public Property Let Answer(ByVal NewValue As Boolean)
    If NewValue = True Then
        pAnswerTrue.Fraction = 100
        pAnswerFalse.Fraction = 0
    Else
        pAnswerTrue.Fraction = 0
        pAnswerFalse.Fraction = 100
    End If
End Property

Public Property Get Truefeedback() As CHTML
    Set Truefeedback = pAnswerTrue.Feedback
End Property

Public Property Set Truefeedback(ByVal NewValue As CHTML)
    Set pAnswerTrue.Feedback = NewValue
End Property

Public Property Get Falsefeedback() As CHTML
    Set Falsefeedback = pAnswerFalse.Feedback
End Property

Public Property Set Falsefeedback(ByVal NewValue As CHTML)
    Set pAnswerFalse.Feedback = NewValue
End Property

Public Sub CQuestion_Export(ByRef DOM As MSXML2.DOMDocument60)
    Dim XMLRootElement As IXMLDOMElement
    Dim XMLQuestionElement As IXMLDOMElement
    Dim XMLElement1 As IXMLDOMElement
    Dim XMLElement2 As IXMLDOMElement
    Dim XMLText As IXMLDOMText
    Dim XMLAttr As IXMLDOMAttribute
    
    Set XMLRootElement = DOM.DocumentElement
    
    Set XMLQuestionElement = DOM.createElement("question")
    XMLRootElement.appendChild XMLQuestionElement
    Set XMLAttr = DOM.createAttribute("type")
    XMLAttr.NodeValue = "truefalse"
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLElement1 = DOM.createElement("name")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(pName)
    XMLElement2.appendChild XMLText
    
    pQuestiontext.Export XMLQuestionElement, "questiontext"
    
    pGeneralfeedback.Export XMLQuestionElement, "generalfeedback"
    
    Set XMLElement1 = DOM.createElement("defaultgrade")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(Replace(Format(pDefaultgrade, "###0.0000000"), ",", "."))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("penalty")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(Replace(Format(pPenalty, "###0.0000000"), ",", "."))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("hidden")
    XMLQuestionElement.appendChild XMLElement1
    If pHidden Then
        Set XMLText = DOM.createTextNode(1)
    Else
        Set XMLText = DOM.createTextNode(0)
    End If
    XMLElement1.appendChild XMLText
            
    pAnswerTrue.Export XMLQuestionElement
    pAnswerFalse.Export XMLQuestionElement
End Sub

Public Sub CQuestion_Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Answer As CMultichoiceAnswer
    Dim Hint As CHint
    Dim Separator As String
    Separator = Replace(Format(0, "0.0"), "0", "")
    Dim TmpAnswer As CShortanswerAnswer

    pName = Element.SelectSingleNode("name/text/text()").NodeValue

    Set pQuestiontext = New CHTML
    pQuestiontext.Import Element.SelectSingleNode("questiontext")

    Set pGeneralfeedback = New CHTML
    pGeneralfeedback.Import Element.SelectSingleNode("generalfeedback")

    pDefaultgrade = CDbl(Replace(Element.SelectSingleNode("defaultgrade/text()").NodeValue, ".", Separator))

    pPenalty = CDbl(Replace(Element.SelectSingleNode("penalty/text()").NodeValue, ".", Separator))

    If Element.SelectSingleNode("hidden/text()").NodeValue = "1" Then
        pHidden = True
    Else
        pHidden = False
    End If

    Set pAnswerTrue = New CShortanswerAnswer
    pAnswerTrue.Import Element.SelectSingleNode("answer[1]")

    Set pAnswerFalse = New CShortanswerAnswer
    pAnswerFalse.Import Element.SelectSingleNode("answer[2]")
    
    If pAnswerTrue.Text = "false" Then
        Set TmpAnswer = pAnswerTrue
        Set pAnswerTrue = pAnswerFalse
        Set pAnswerFalse = TmpAnswer
        Set TmpAnswer = Nothing
    End If
End Sub

