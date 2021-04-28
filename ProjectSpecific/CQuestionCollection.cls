VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuestionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing collection of moodle questions
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

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

Public Function Item(Index) As CQuestion
    Set Item = pQuestions.Item(Index)
End Function

Public Function Count() As Long
    Count = pQuestions.Count
End Function

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
    DOM.preserveWhiteSpace = True
    Set XMLRootElement = DOM.createElement("quiz")
    DOM.appendChild XMLRootElement
    Set CreateDOM = DOM
End Function

Public Sub Import(XMLFilename As String)
    Dim DOM As MSXML2.DOMDocument60
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Element As MSXML2.IXMLDOMElement
    Dim Question As Object
    Dim QuestionType As String
        
    Set DOM = New MSXML2.DOMDocument60
    DOM.Load XMLFilename
    Set Nodes = DOM.SelectNodes("/quiz/question")
    For Each Node In Nodes
        If Node.NodeType = NODE_ELEMENT Then
            Set Element = Node
            Set Question = GetQuestion(Element)
            If Not Question Is Nothing Then
                pQuestions.Add Question
            End If
        End If
    Next
End Sub

Private Function GetQuestion(Element As MSXML2.IXMLDOMElement) As Object
    Dim QuestionType As String
    Dim Question As CQuestion
        
    QuestionType = Element.getAttribute("type")
    Select Case QuestionType
        Case "category"
            Set Question = New CCategory
            Question.Import Element
            Set GetQuestion = Question
        Case "description"
            Set Question = New CDescription
            Question.Import Element
            Set GetQuestion = Question
        Case "ddmatch"
            Set Question = New CDdmatch
            Question.Import Element
            Set GetQuestion = Question
        Case "essay"
            Set Question = New CEssay
            Question.Import Element
            Set GetQuestion = Question
        Case "matching"
            Set Question = New CMatching
            Question.Import Element
            Set GetQuestion = Question
        Case "multichoice"
            Set Question = New CMultichoice
            Question.Import Element
            Set GetQuestion = Question
        Case "numerical"
            Set Question = New CNumerical
            Question.Import Element
            Set GetQuestion = Question
        Case "order"
            Set Question = New COrder
            Question.Import Element
            Set GetQuestion = Question
        Case "shortanswer"
            Set Question = New CShortanswer
            Question.Import Element
            Set GetQuestion = Question
        Case "truefalse"
            Set Question = New CTrueFalse
            Question.Import Element
            Set GetQuestion = Question
        Case Else
            Set GetQuestion = Nothing
    End Select
    
End Function
