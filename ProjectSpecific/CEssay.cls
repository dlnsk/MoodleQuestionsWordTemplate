VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEssay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle essay question
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion
Implements IExportable

Public Enum EssayReponseformatOptions
    reponseformat_editor
    reponseformat_editorfilepicker
    reponseformat_plain
    reponseformat_monospaced
End Enum

Public Enum EssayResponsefieldlinesOptions
    responsefieldlines_5 = 5
    responsefieldlines_10 = 10
    responsefieldlines_15 = 15
    responsefieldlines_20 = 20
    responsefieldlines_25 = 25
    responsefieldlines_30 = 30
    responsefieldlines_35 = 35
    responsefieldlines_40 = 40
End Enum

Public Enum Essay¿ttachmentsOptions
    attachments_0 = 0
    attachments_1 = 1
    attachments_2 = 2
    attachments_3 = 3
    attachments_unlimited = -1
End Enum

Private pQuestion As CQuestion
Private pResponseformat As Long
Private pResponsefieldlines As Long
Private pAttachments As Long
Private pGraderinfo As CHTML
Private pResponsetemplate As CHTML

Private Sub Class_Initialize()
    Set pQuestion = New CQuestion
    pQuestion.QType = "essay"
    pResponseformat = EssayReponseformatOptions.reponseformat_editor
    pResponsefieldlines = EssayResponsefieldlinesOptions.responsefieldlines_15
    pAttachments = Essay¿ttachmentsOptions.attachments_0
    Set pGraderinfo = New CHTML
    Set pResponsetemplate = New CHTML
End Sub

Private Sub Class_Terminate()
    Set pGraderinfo = Nothing
    Set pResponsetemplate = Nothing
    Set pQuestion = Nothing
End Sub

Public Property Get Responseformat() As Long
    Responseformat = pResponseformat
End Property

Public Property Let Responseformat(ByVal NewValue As Long)
    pResponseformat = NewValue
End Property

Public Property Get Responsefieldlines() As Long
    Responsefieldlines = pResponsefieldlines
End Property

Public Property Let Responsefieldlines(ByVal NewValue As Long)
    pResponsefieldlines = NewValue
End Property

Public Property Get Attachments() As Long
    Attachments = pAttachments
End Property

Public Property Let Attachments(ByVal NewValue As Long)
    pAttachments = NewValue
End Property

Public Property Get Graderinfo() As CHTML
    Set Graderinfo = pGraderinfo
End Property

Public Property Set Graderinfo(ByVal NewValue As CHTML)
    Set pGraderinfo = NewValue
End Property

Public Property Get Responsetemplate() As CHTML
    Set Responsetemplate = pResponsetemplate
End Property

Public Property Let Responsetemplate(ByVal NewValue As CHTML)
    Set pResponsetemplate = NewValue
End Property

Public Sub IExportable_Export(ByRef DOM As MSXML2.DOMDocument60)
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
    XMLAttr.NodeValue = "essay"
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLElement1 = DOM.createElement("name")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(RangeToHTML.TidyHTMLFormatng(pName))
    XMLElement2.appendChild XMLText
    
    QuestionText.Export XMLQuestionElement, "questiontext"
    
    GeneralFeedback.Export XMLQuestionElement, "generalfeedback"
    
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
    
    Set XMLElement1 = DOM.createElement("responseformat")
    XMLQuestionElement.appendChild XMLElement1
    Select Case pResponseformat
        Case EssayReponseformatOptions.reponseformat_editor
            Set XMLText = DOM.createTextNode("editor")
        Case EssayReponseformatOptions.reponseformat_editorfilepicker
            Set XMLText = DOM.createTextNode("editorfilepicker")
        Case EssayReponseformatOptions.reponseformat_monospaced
            Set XMLText = DOM.createTextNode("monospaced")
        Case EssayReponseformatOptions.reponseformat_plain
            Set XMLText = DOM.createTextNode("plain")
    End Select
    XMLElement1.appendChild XMLText

    Set XMLElement1 = DOM.createElement("responsefieldlines")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(CStr(pResponsefieldlines))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("attachments")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(CStr(pAttachments))
    XMLElement1.appendChild XMLText
    
    pGraderinfo.Export XMLQuestionElement, "graderinfo"
    
    pResponsetemplate.Export XMLQuestionElement, "responsetemplate"
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
    
    Select Case Element.SelectSingleNode("responseformat/text()").NodeValue
        Case "editor"
            pResponseformat = EssayReponseformatOptions.reponseformat_editor
        Case "editorfilepicker"
            pResponseformat = EssayReponseformatOptions.reponseformat_editorfilepicker
        Case "monospaced"
            pResponseformat = EssayReponseformatOptions.reponseformat_monospaced
        Case "plain"
            pResponseformat = EssayReponseformatOptions.reponseformat_plain
        Case Else
            Err.Raise vbObjectError + 512 + 1, "CEssay", "Unknown responseformat value"
    End Select
    
    pResponsefieldlines = CLng(Element.SelectSingleNode("responsefieldlines/text()").NodeValue)
    
    pAttachments = CLng(Element.SelectSingleNode("attachments/text()").NodeValue)

    Set pGraderinfo = New CHTML
    pGraderinfo.Import Element.SelectSingleNode("graderinfo")
    
    Set pResponsetemplate = New CHTML
    pResponsetemplate.Import Element.SelectSingleNode("responsetemplate")
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
