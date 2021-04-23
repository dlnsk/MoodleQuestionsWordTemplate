VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMultichoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle multichoice question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements CQuestion

Public Enum MultichoiceAnswernumberingOptions
    numbering_abc
    numbering_ABCD
    numbering_123
    numbering_iii
    numbering_IIII
    numbering_none
End Enum

Private pName As String
Private pQuestiontext As CHTML
Private pGeneralfeedback As CHTML
Private pDefaultgrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pSingleanswer As Boolean
Private pShuffleanswers As Boolean
Private pAnswernumbering As Long
Private pCorrectfeedback As CHTML
Private pPartiallycorrectfeedback As CHTML
Private pIncorrectfeedback As CHTML
Private pShownumcorrect As Boolean
Private pAnswers As Collection
Private pHints As Collection

Private Sub Class_Initialize()
    pName = "-"
    Set pQuestiontext = New CHTML
    Set pGeneralfeedback = New CHTML
    pDefaultgrade = 1
    pPenalty = 0.3333333
    pHidden = False
    pSingleanswer = True
    pShuffleanswers = True
    pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_abc
    Set pCorrectfeedback = New CHTML
    Set pPartiallycorrectfeedback = New CHTML
    Set pIncorrectfeedback = New CHTML
    pShownumcorrect = 1
    Set pAnswers = New Collection
    Set pHints = New Collection
End Sub

Private Sub Class_Terminate()
    Set pQuestiontext = Nothing
    Set pGeneralfeedback = Nothing
    Set pCorrectfeedback = Nothing
    Set pPartiallycorrectfeedback = Nothing
    Set pIncorrectfeedback = Nothing
    Set pAnswers = Nothing
    Set pHints = Nothing
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

Public Property Get Penalty() As Double
    Penalty = pPenalty
End Property

Public Property Let Penalty(ByVal NewValue As Double)
    pPenalty = NewValue
End Property
'
'Public Property Get Hidden() As Boolean
'    Hidden = pHidden
'End Property
'
'Public Property Let Hidden(ByVal NewValue As Boolean)
'    pHidden = NewValue
'End Property

Public Property Get Singleanswer() As Boolean
    Singleanswer = pSingleanswer
End Property

Public Property Let Singleanswer(ByVal NewValue As Boolean)
    pSingleanswer = NewValue
End Property

Public Property Get Shuffleanswers() As Boolean
    Shuffleanswers = pShuffleanswers
End Property

Public Property Let Shuffleanswers(ByVal NewValue As Boolean)
    pShuffleanswers = NewValue
End Property

Public Property Get Answernumbering() As Long
    Answernumbering = pAnswernumbering
End Property

Public Property Let Answernumbering(ByVal NewValue As Long)
    pAnswernumbering = NewValue
End Property

Public Property Get Correctfeedback() As CHTML
    Set Correctfeedback = pCorrectfeedback
End Property

Public Property Set Correctfeedback(ByVal NewValue As CHTML)
    Set pCorrectfeedback = NewValue
End Property

Public Property Get Partiallycorrectfeedback() As CHTML
    Set Partiallycorrectfeedback = pPartiallycorrectfeedback
End Property

Public Property Set Partiallycorrectfeedback(ByVal NewValue As CHTML)
    Set pPartiallycorrectfeedback = NewValue
End Property

Public Property Get Incorrectfeedback() As CHTML
    Set Incorrectfeedback = pIncorrectfeedback
End Property

Public Property Set Incorrectfeedback(ByVal NewValue As CHTML)
    Set pIncorrectfeedback = NewValue
End Property

Public Property Get Shownumcorrect() As Boolean
    Shownumcorrect = pShownumcorrect
End Property

Public Property Let Shownumcorrect(ByVal NewValue As Boolean)
    pShownumcorrect = NewValue
End Property

Public Property Get Answers() As Collection
    Set Answers = pAnswers
End Property

Public Property Set Answers(ByVal NewValue As Collection)
    Set pAnswers = NewValue
End Property

Public Property Get Hints() As Collection
    Set Hints = pHints
End Property

Public Property Set Hints(ByVal NewValue As Collection)
    Set pHints = NewValue
End Property

Public Sub CQuestion_Export(ByRef DOM As MSXML2.DOMDocument60)
    Dim XMLRootElement As IXMLDOMElement
    Dim XMLQuestionElement As IXMLDOMElement
    Dim XMLElement1 As IXMLDOMElement
    Dim XMLElement2 As IXMLDOMElement
    Dim XMLText As IXMLDOMText
    Dim XMLAttr As IXMLDOMAttribute
    Dim Answer As CMultichoiceAnswer
    Dim Hint As CHint
    
    Set XMLRootElement = DOM.DocumentElement
    
    Set XMLQuestionElement = DOM.createElement("question")
    XMLRootElement.appendChild XMLQuestionElement
    Set XMLAttr = DOM.createAttribute("type")
    XMLAttr.NodeValue = "multichoice"
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
    
    Set XMLElement1 = DOM.createElement("single")
    XMLQuestionElement.appendChild XMLElement1
    If pSingleanswer Then
        Set XMLText = DOM.createTextNode("true")
    Else
        Set XMLText = DOM.createTextNode("false")
    End If
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("shuffleanswers")
    XMLQuestionElement.appendChild XMLElement1
    If pShuffleanswers Then
        Set XMLText = DOM.createTextNode("true")
    Else
        Set XMLText = DOM.createTextNode("false")
    End If
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("answernumbering")
    XMLQuestionElement.appendChild XMLElement1
    Select Case pAnswernumbering
        Case MultichoiceAnswernumberingOptions.numbering_123
            Set XMLText = DOM.createTextNode("123")
        Case MultichoiceAnswernumberingOptions.numbering_abc
            Set XMLText = DOM.createTextNode("abc")
        Case MultichoiceAnswernumberingOptions.numbering_ABCD
            Set XMLText = DOM.createTextNode("ABCD")
        Case MultichoiceAnswernumberingOptions.numbering_iii
            Set XMLText = DOM.createTextNode("iii")
        Case MultichoiceAnswernumberingOptions.numbering_IIII
            Set XMLText = DOM.createTextNode("IIII")
        Case MultichoiceAnswernumberingOptions.numbering_none
            Set XMLText = DOM.createTextNode("none")
    End Select
    XMLElement1.appendChild XMLText
    
    pCorrectfeedback.Export XMLQuestionElement, "correctfeedback"
    pPartiallycorrectfeedback.Export XMLQuestionElement, "partiallycorrectfeedback"
    pIncorrectfeedback.Export XMLQuestionElement, "incorrectfeedback"
    
    If pShownumcorrect Then
        Set XMLElement1 = DOM.createElement("shownumcorrect")
        XMLQuestionElement.appendChild XMLElement1
    End If
    
    For Each Answer In pAnswers
        Answer.Export XMLQuestionElement
    Next Answer
    
    For Each Hint In pHints
        Hint.Export XMLQuestionElement
    Next Hint
End Sub

Public Sub CQuestion_Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Answer As CMultichoiceAnswer
    Dim Hint As CHint
    Dim Separator As String
    Separator = Replace(Format(0, "0.0"), "0", "")
    
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
    
    If Element.SelectSingleNode("single/text()").NodeValue = "true" Then
        pSingleanswer = True
    Else
        pSingleanswer = False
    End If
    
    If Element.SelectSingleNode("shuffleanswers/text()").NodeValue = "true" Then
        pShuffleanswers = True
    Else
        pShuffleanswers = False
    End If
    
    Select Case Element.SelectSingleNode("answernumbering/text()").NodeValue
        Case "123"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_123
        Case "abc"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_abc
        Case "ABCD"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_ABCD
        Case "iii"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_iii
        Case "IIII"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_IIII
        Case "none"
            pAnswernumbering = MultichoiceAnswernumberingOptions.numbering_none
        Case Else
            Err.Raise vbObjectError + 512 + 1, "CMultichoice", "Unknown answernumbering value"
    End Select
    
    Set pCorrectfeedback = New CHTML
    pCorrectfeedback.Import Element.SelectSingleNode("correctfeedback")
    
    Set pPartiallycorrectfeedback = New CHTML
    pPartiallycorrectfeedback.Import Element.SelectSingleNode("partiallycorrectfeedback")
    
    Set pIncorrectfeedback = New CHTML
    pIncorrectfeedback.Import Element.SelectSingleNode("incorrectfeedback")
    
    If Not Element.SelectSingleNode("shownumcorrect") Is Nothing Then
        pShownumcorrect = True
    Else
        pShownumcorrect = False
    End If
    
    Set pAnswers = New Collection
    Set Nodes = Element.SelectNodes("answer")
    For Each Node In Nodes
        Set Answer = New CMultichoiceAnswer
        Answer.Import Node
        pAnswers.Add Answer
    Next

    Set pHints = New Collection
    Set Nodes = Element.SelectNodes("hint")
    For Each Node In Nodes
        Set Hint = New CHint
        Hint.Import Node
        pHints.Add Hint
    Next
End Sub

