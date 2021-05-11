VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion
Implements IExportable

Private pName As String
Private pQType As String
Private pQuestionText As CHTML
Private pGeneralFeedback As CHTML
Private pDefaultGrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pAnswers As Collection
Private pHints As Collection


Private Sub Class_Initialize()
    pName = "-"
    pQType = "-"
    Set pQuestionText = New CHTML
    Set pGeneralFeedback = New CHTML
    pDefaultGrade = 1
    pPenalty = 0.3333333
    pHidden = False
    Set pAnswers = New Collection
    Set pHints = New Collection
End Sub

Private Sub Class_Terminate()
    Set pQuestionText = Nothing
    Set pGeneralFeedback = Nothing
    Set pAnswers = Nothing
    Set pHints = Nothing
End Sub

Public Sub Export(ByRef DOM As MSXML2.DOMDocument60)
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
    XMLAttr.NodeValue = pQType
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLElement1 = DOM.createElement("name")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(RangeToHTML.TidyHTMLFormatng(pName))
    XMLElement2.appendChild XMLText
    
    pQuestionText.Export XMLQuestionElement, "questiontext"
    
    pGeneralFeedback.Export XMLQuestionElement, "generalfeedback"
    
    Set XMLElement1 = DOM.createElement("defaultgrade")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(Replace(Format(pDefaultGrade, "###0.0000000"), ",", "."))
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
    
    For Each Answer In pAnswers
        Answer.Export XMLQuestionElement
    Next Answer
    
    For Each Hint In pHints
        Hint.Export XMLQuestionElement
    Next Hint
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Answer As CMultichoiceAnswer
    Dim Hint As CHint
    Dim Separator As String
    Separator = Replace(Format(0, "0.0"), "0", "")
    
    pName = Element.SelectSingleNode("name/text/text()").NodeValue
    
    Set pQuestionText = New CHTML
    pQuestionText.Import Element.SelectSingleNode("questiontext")
    
    Set pGeneralFeedback = New CHTML
    pGeneralFeedback.Import Element.SelectSingleNode("generalfeedback")
    
    pDefaultGrade = CDbl(Replace(Element.SelectSingleNode("defaultgrade/text()").NodeValue, ".", Separator))
    
    pPenalty = CDbl(Replace(Element.SelectSingleNode("penalty/text()").NodeValue, ".", Separator))
    
    If Element.SelectSingleNode("hidden/text()").NodeValue = "1" Then
        pHidden = True
    Else
        pHidden = False
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

Public Sub IExportable_Export(ByRef DOM As MSXML2.DOMDocument60)
    Export DOM
End Sub

Public Sub IExportable_Import(ByRef Element As MSXML2.IXMLDOMElement)
    Import Element
End Sub



'=================== Class's own properties folowing interface IQuestion ========================
' Properties of this class a little different from real classes CMultichoice etc. by
' combination of Get/Set/Let/ByRef

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal NewValue As String)
    pName = NewValue
End Property

Public Property Get QType() As String
    Name = pName
End Property

Public Property Let QType(ByVal NewValue As String)
    pName = NewValue
End Property

Public Property Get QuestionText() As CHTML
    Set QuestionText = pQuestionText
End Property

Public Property Let QuestionText(ByRef NewValue As CHTML)
    Set pQuestionText = NewValue
End Property

Public Property Get GeneralFeedback() As CHTML
    Set GeneralFeedback = pGeneralFeedback
End Property

Public Property Let GeneralFeedback(ByRef NewValue As CHTML)
    Set pGeneralFeedback = NewValue
End Property

Public Property Get DefaultGrade() As Double
    DefaultGrade = pDefaultGrade
End Property

Public Property Let DefaultGrade(ByVal NewValue As Double)
    pDefaultGrade = NewValue
End Property

Public Property Get Penalty() As Double
    Penalty = pPenalty
End Property

Public Property Let Penalty(ByVal NewValue As Double)
    pPenalty = NewValue
End Property

Public Property Get Hidden() As Boolean
    Hidden = pHidden
End Property

Public Property Let Hidden(ByVal NewValue As Boolean)
    pHidden = NewValue
End Property

Public Property Get Answers() As Collection
    Set Answers = pAnswers
End Property

Public Property Let Answers(ByRef NewValue As Collection)
    Set pAnswers = NewValue
End Property

Public Property Get Hints() As Collection
    Set Hints = pHints
End Property

Public Property Let Hints(ByRef NewValue As Collection)
    Set pHints = NewValue
End Property

'======== Implementing interface IQuestion ============================================
' This part is identical in all classes which use IQuestion interface

Public Property Get IQuestion_Name() As String
    IQuestion_Name = Name
End Property

Public Property Let IQuestion_Name(ByVal NewValue As String)
    Name = NewValue
End Property

Public Property Get IQuestion_QType() As String
    IQuestion_QType = QType
End Property

Public Property Let IQuestion_QType(ByVal NewValue As String)
    QType = NewValue
End Property

Public Property Get IQuestion_QuestionText() As CHTML
    Set IQuestion_QuestionText = QuestionText
End Property

Public Property Let IQuestion_QuestionText(ByRef NewValue As CHTML)
    QuestionText = NewValue
End Property

Public Property Get IQuestion_GeneralFeedback() As CHTML
    Set IQuestion_GeneralFeedback = GeneralFeedback
End Property

Public Property Let IQuestion_GeneralFeedback(ByRef NewValue As CHTML)
    GeneralFeedback = NewValue
End Property

Public Property Get IQuestion_DefaultGrade() As Double
    IQuestion_DefaultGrade = DefaultGrade
End Property

Public Property Let IQuestion_DefaultGrade(ByVal NewValue As Double)
    DefaultGrade = NewValue
End Property

Public Property Get IQuestion_Penalty() As Double
    IQuestion_Penalty = Penalty
End Property

Public Property Let IQuestion_Penalty(ByVal NewValue As Double)
    Penalty = NewValue
End Property

Public Property Get IQuestion_Hidden() As Boolean
    IQuestion_Hidden = Hidden
End Property

Public Property Let IQuestion_Hidden(ByVal NewValue As Boolean)
    Hidden = NewValue
End Property

Public Property Get IQuestion_Answers() As Collection
    Set IQuestion_Answers = Answers
End Property

Public Property Let IQuestion_Answers(ByRef NewValue As Collection)
    Answers = NewValue
End Property

Public Property Get IQuestion_Hints() As Collection
    Set IQuestion_Hints = Hints
End Property

Public Property Let IQuestion_Hints(ByRef NewValue As Collection)
    Hints = NewValue
End Property

