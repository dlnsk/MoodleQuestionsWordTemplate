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
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements CQuestion

Private pCategoryName As String

Public Property Get Name() As String
    Name = pCategoryName
End Property

Public Property Let Name(ByVal NewValue As String)
    pCategoryName = NewValue
End Property

Public Sub CQuestion_Export(ByRef DOM As MSXML2.DOMDocument60)
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
    
    Set XMLText = DOM.createTextNode(pCategoryName)
    XMLTextElement.appendChild XMLText
End Sub
