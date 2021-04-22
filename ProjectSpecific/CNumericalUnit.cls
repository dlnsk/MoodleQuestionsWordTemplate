VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNumericalUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing unit in moodle numerical question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pMultiplier As Double
Private pName As String

Public Property Get Multiplier() As Double
    Multiplier = pMultiplier
End Property

Public Property Let Multiplier(ByVal NewValue As Double)
    pMultiplier = NewValue
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal NewValue As String)
    pName = NewValue
End Property

Public Sub Export(ByRef UnitsElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLElement1 As MSXML2.IXMLDOMElement
    Dim XMLElement2 As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute

    Set DOM = UnitsElement.OwnerDocument
    Set XMLElement1 = DOM.createElement("unit")
    UnitsElement.appendChild XMLElement1
    
    Set XMLElement2 = DOM.createElement("multiplier")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(Replace(Format(pMultiplier, "###0.00000"), ",", "."))
    XMLElement2.appendChild XMLText
    
    Set XMLElement2 = DOM.createElement("unit_name")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(pName)
    XMLElement2.appendChild XMLText
End Sub

