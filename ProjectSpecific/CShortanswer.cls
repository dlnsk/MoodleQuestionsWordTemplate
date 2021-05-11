VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShortanswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle shortanswer question
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion
Implements IExportable

Private pQuestion As CQuestion
Private pUsecase As Boolean

Private Sub Class_Initialize()
    Set pQuestion = New CQuestion
    pQuestion.QType = "shortanswer"
    pUsecase = False
End Sub

Private Sub Class_Terminate()
    Set pQuestion = Nothing
End Sub

Public Property Get Usecase() As Boolean
    Usecase = pUsecase
End Property

Public Property Let Usecase(ByVal NewValue As Boolean)
    pUsecase = NewValue
End Property


Public Sub IExportable_Export(ByRef DOM As MSXML2.DOMDocument60)
    Dim XMLRootElement As IXMLDOMElement
    Dim XMLQuestionElement As IXMLDOMElement
    Dim XMLElement1 As IXMLDOMElement
    Dim XMLElement2 As IXMLDOMElement
    Dim XMLText As IXMLDOMText
    Dim XMLAttr As IXMLDOMAttribute
    Dim Answer As CShortanswerAnswer
    Dim Hint As CShortanswerHint
    
    Set XMLRootElement = DOM.DocumentElement
    
    Set XMLQuestionElement = DOM.createElement("question")
    XMLRootElement.appendChild XMLQuestionElement
    Set XMLAttr = DOM.createAttribute("type")
    XMLAttr.NodeValue = "shortanswer"
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
    
    Set XMLElement1 = DOM.createElement("usecase")
    XMLQuestionElement.appendChild XMLElement1
    If pUsecase Then
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

Public Sub IExportable_Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Answer As CShortanswerAnswer
    Dim Hint As CShortanswerHint
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
    
    If Element.SelectSingleNode("usecase/text()").NodeValue = "1" Then
        pUsecase = True
    Else
        pUsecase = False
    End If
    
    Set pAnswers = New Collection
    Set Nodes = Element.SelectNodes("answer")
    For Each Node In Nodes
        Set Answer = New CShortanswerAnswer
        Answer.Import Node
        pAnswers.Add Answer
    Next

    Set pHints = New Collection
    Set Nodes = Element.SelectNodes("hint")
    For Each Node In Nodes
        Set Hint = New CShortanswerHint
        Hint.Import Node
        pHints.Add Hint
    Next
End Sub



'=================== Class's own properties folowing interface IQuestion ========================
' Properties of this class and other real question a little different from CQuestion etc. by
' combination of Get/Set/Let/ByRef

Public Property Get Name() As String
    Name = pQuestion.Name
End Property

Public Property Let Name(ByVal NewValue As String)
    pQuestion.Name = NewValue
End Property

Public Property Get QType() As String
    QType = pQuestion.QType
End Property

Public Property Let QType(ByVal NewValue As String)
    pQuestion.QType = NewValue
End Property

Public Property Get QuestionText() As CHTML
    Set QuestionText = pQuestion.QuestionText
End Property

Public Property Let QuestionText(ByRef NewValue As CHTML)
    pQuestion.QuestionText = NewValue
End Property

Public Property Get GeneralFeedback() As CHTML
    Set GeneralFeedback = pQuestion.GeneralFeedback
End Property

Public Property Let GeneralFeedback(ByRef NewValue As CHTML)
    pQuestion.GeneralFeedback = NewValue
End Property

Public Property Get DefaultGrade() As Double
    DefaultGrade = pQuestion.DefaultGrade
End Property

Public Property Let DefaultGrade(ByVal NewValue As Double)
    pQuestion.DefaultGrade = NewValue
End Property

Public Property Get Penalty() As Double
    Penalty = pQuestion.Penalty
End Property

Public Property Let Penalty(ByVal NewValue As Double)
    pQuestion.Penalty = NewValue
End Property

Public Property Get Hidden() As Boolean
    Hidden = pQuestion.Hidden
End Property

Public Property Let Hidden(ByVal NewValue As Boolean)
    pQuestion.Hidden = NewValue
End Property

Public Property Get Answers() As Collection
    Set Answers = pQuestion.Answers
End Property

Public Property Let Answers(ByRef NewValue As Collection)
    pQuestion.Answers = NewValue
End Property

Public Property Get Hints() As Collection
    Set Hints = pQuestion.Hints
End Property

Public Property Let Hints(ByRef NewValue As Collection)
    pQuestion.Hints = NewValue
End Property

'======== Implementing interface IQuestion ============================================
' All properies defined below are identical in all real question's classes.

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
