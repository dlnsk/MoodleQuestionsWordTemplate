VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuestionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pQuestions As Collection

Private Sub Class_Initialize()
    Set pQuestions = New Collection
End Sub
Private Sub Class_Terminate()
    Set pQuestions = Nothing
End Sub
Public Sub Add(Question As CQuestion)
    pQuestions.Add Question
End Sub
Public Sub Export(XMLFilename As String)
    Dim DOM As MSXML2.DOMDocument60
    Dim Question As CQuestion
    
    Set DOM = CreateDOM
    For Each Question In pQuestions
        Question.Export DOM
    Next
    DOM.Save XMLFilename
End Sub
Private Function CreateDOM()
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLRootElement As MSXML2.IXMLDOMElement
    Dim XMLElement As MSXML2.IXMLDOMElement
    Dim XMLAttr As MSXML2.IXMLDOMAttribute
    
    Set DOM = New MSXML2.DOMDocument60
    Set XMLRootElement = DOM.createElement("quiz")
    DOM.appendChild XMLRootElement
    Set CreateDOM = DOM
End Function
