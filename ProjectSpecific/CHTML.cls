VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle html textarea with attached files.
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pText As String
Private pFiles As CFilesCollection

Private Sub Class_Initialize()
    Set pFiles = New CFilesCollection
End Sub

Private Sub Class_Terminate()
    Set pFiles = Nothing
End Sub

Public Property Get Text() As String
    Text = pText
End Property

Public Property Let Text(ByVal NewValue As String)
    pText = NewValue
End Property

Public Property Get Files() As CFilesCollection
    Set Files = pFiles
End Property

Public Property Set Files(ByVal NewValue As CFilesCollection)
    Set pFiles = NewValue
End Property

Public Sub AddFile(Data As String, Filename As String)
    pFiles.Add Data, Filename
End Sub

Public Sub Export(ByRef ParentElement As MSXML2.IXMLDOMElement, ElementName As String, Optional CreatedElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLElement1 As MSXML2.IXMLDOMElement
    Dim XMLElement2 As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute
    Dim nodeCDATA As MSXML2.IXMLDOMCDATASection
    
    Set DOM = ParentElement.OwnerDocument
    Set XMLElement1 = DOM.createElement(ElementName)
    ParentElement.appendChild XMLElement1
    Set XMLAttr = DOM.createAttribute("format")
    XMLAttr.NodeValue = "html"
    XMLElement1.setAttributeNode XMLAttr
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    If pText <> "" Then
        Set nodeCDATA = DOM.createCDATASection(CleanString(pText))
        XMLElement2.appendChild nodeCDATA
    End If
    
'    Set XMLText = DOM.createTextNode(pText)
'    XMLElement2.appendChild XMLText
    pFiles.Export XMLElement1
    Set CreatedElement = XMLElement1
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Format As String
    Format = Element.SelectSingleNode("@format").NodeValue
    
    Select Case Format
        Case "moodle_auto_format", "html", "plain_text"
            If Not Element.SelectSingleNode("text/text()") Is Nothing Then
                pText = Element.SelectSingleNode("text/text()").NodeValue
            Else
                pText = ""
            End If
        Case Else
            Err.Raise vbObjectError + 512 + 1, "CHTML", "Unsupported text format '" + Format + "'"
    End Select
    
    Set pFiles = New CFilesCollection
    pFiles.Import Element
End Sub
