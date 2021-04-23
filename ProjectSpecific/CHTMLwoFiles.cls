VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHTMLwoFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle html textarea which does not allow to attach files (currently used only in responsetemplate in essay question).
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pText As String

Public Property Get Text() As String
    Text = pText
End Property

Public Property Let Text(ByVal NewValue As String)
    pText = NewValue
End Property

Public Sub Export(ByRef ParentElement As MSXML2.IXMLDOMElement, ElementName As String, Optional CreatedElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLElement1 As MSXML2.IXMLDOMElement
    Dim XMLElement2 As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute
    
    Set DOM = ParentElement.OwnerDocument
    Set XMLElement1 = DOM.createElement(ElementName)
    ParentElement.appendChild XMLElement1
    Set XMLAttr = DOM.createAttribute("format")
    XMLAttr.NodeValue = "html"
    XMLElement1.setAttributeNode XMLAttr
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(pText)
    XMLElement2.appendChild XMLText
    Set CreatedElement = XMLElement1
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Format As String
    Format = Element.SelectSingleNode("@format").NodeValue
    
    Select Case Format
        Case "html"
            If Not Element.SelectSingleNode("text/text()") Is Nothing Then
                pText = Element.SelectSingleNode("text/text()").NodeValue
            Else
                pText = ""
            End If
        Case "moodle_auto_format"
            If Not Element.SelectSingleNode("text/text()") Is Nothing Then
                pText = Element.SelectSingleNode("text/text()").NodeValue
            Else
                pText = ""
            End If
        Case Else
            Err.Raise vbObjectError + 512 + 1, "CHTML", "Unsupported text format '" + Format + "'"
    End Select
End Sub
