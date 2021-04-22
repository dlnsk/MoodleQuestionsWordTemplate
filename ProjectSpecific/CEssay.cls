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
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements CQuestion

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

Private pName As String
Private pQuestiontext As CHTML
Private pGeneralfeedback As CHTML
Private pDefaultgrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pResponseformat As Long
Private pResponsefieldlines As Long
Private pAttachments As Long
Private pGraderinfo As CHTML
Private pResponsetemplate As CHTMLwoFiles

Private Sub Class_Initialize()
    pName = "-"
    Set pQuestiontext = New CHTML
    Set pGeneralfeedback = New CHTML
    pDefaultgrade = 1
    pPenalty = 0.3333333
    pHidden = False
    pResponseformat = EssayReponseformatOptions.reponseformat_editor
    pResponsefieldlines = EssayResponsefieldlinesOptions.responsefieldlines_15
    pAttachments = Essay¿ttachmentsOptions.attachments_0
    Set pGraderinfo = New CHTML
    Set pResponsetemplate = New CHTMLwoFiles
End Sub

Private Sub Class_Terminate()
    Set pQuestiontext = Nothing
    Set pGeneralfeedback = Nothing
    Set pGraderinfo = Nothing
    Set pResponsetemplate = Nothing
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

Public Property Get Responsetemplate() As String
    Responsetemplate = pResponsetemplate.Text
End Property

Public Property Let Responsetemplate(ByVal NewValue As String)
    pResponsetemplate.Text = NewValue
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
    XMLAttr.NodeValue = "essay"
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLElement1 = DOM.createElement("name")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(pName)
    XMLElement2.appendChild XMLText
    
    QuestionText.Export XMLQuestionElement, "questiontext"
    
    Generalfeedback.Export XMLQuestionElement, "generalfeedback"
    
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


