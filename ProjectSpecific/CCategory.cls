VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle question category
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion
Implements IExportable

Private pQuestion As CQuestion

Private Sub Class_Initialize()
    Set pQuestion = New CQuestion
    pQuestion.QType = "category"
End Sub

Private Sub Class_Terminate()
    Set pQuestion = Nothing
End Sub


Public Sub IExportable_Export(ByRef DOM As MSXML2.DOMDocument60)
    Dim XMLRootElement As MSXML2.IXMLDOMElement
    Dim XMLQuestionElement As MSXML2.IXMLDOMElement
    Dim XMLCategoryElement As MSXML2.IXMLDOMElement
    Dim XMLTextElement As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute
    
    Set XMLRootElement = DOM.DocumentElement
    
    Set XMLQuestionElement = DOM.createElement("question")
    XMLRootElement.appendChild XMLQuestionElement
    
    Set XMLAttr = DOM.createAttribute("type")
    XMLAttr.NodeValue = "category"
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLCategoryElement = DOM.createElement("category")
    XMLQuestionElement.appendChild XMLCategoryElement
    
    Set XMLTextElement = DOM.createElement("text")
    XMLCategoryElement.appendChild XMLTextElement
    
    Set XMLText = DOM.createTextNode(RangeToHTML.TidyHTMLFormatng(pQuestion.Name))
    XMLTextElement.appendChild XMLText
End Sub

Public Sub IExportable_Import(ByRef Element As MSXML2.IXMLDOMElement)
    pQuestion.Name = Element.SelectSingleNode("category/text/text()").NodeValue
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
