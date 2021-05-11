VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNumerical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle numeriacal question
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion

Public Enum NumericalUnitroleOptions
    Unitrole_onlynumerical = 3
    Unitrole_manynumerical = 0
    unitrole_unitgraded = 1
End Enum

Public Enum NumericalUnitgradingtypeOptions
    unitgradingtype_none = 0
    unitgradingtype_FractionOfResponseGrade = 1
    unitgradingtype_FractionOfQuestionGrade = 2
End Enum

Public Enum NumericalShowunitsOptions
    showunits_text = 0
    showunits_choice = 1
    showunits_select = 2
    showunits_none = 3
End Enum

Public Enum NumericalUnitpositionOptions
    unitposition_right = 0
    unitposition_left = 1
End Enum

Private pQuestion As CQuestion
Private pName As String
Private pQuestionText As CHTML
Private pGeneralFeedback As CHTML
Private pDefaultGrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pUnits As Collection
Private pUnitpenalty As Double
Private pUnitrole As Long
Private pUnitgradingtype As Long
Private pShowunits As Long
Private pUnitposition  As Long
Private pAnswers As Collection
Private pHints As Collection

Private Sub Class_Initialize()
    Set pQuestion = New CQuestion
    pName = "-"
    Set pQuestionText = New CHTML
    Set pGeneralFeedback = New CHTML
    pDefaultGrade = 1
    pPenalty = 0.3333333
    pHidden = False
    Set pUnits = New Collection
    Unitrole = NumericalUnitroleOptions.Unitrole_onlynumerical
    pUnitpenalty = 0.1
    pUnitposition = NumericalUnitpositionOptions.unitposition_right
    Set pAnswers = New Collection
    Set pHints = New Collection
End Sub

Private Sub Class_Terminate()
    Set pQuestionText = Nothing
    Set pGeneralFeedback = Nothing
    Set pAnswers = Nothing
    Set pHints = Nothing
    Set pQuestion = Nothing
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal NewValue As String)
    pName = NewValue
End Property

Public Property Get QuestionText() As CHTML
    Set QuestionText = pQuestionText
End Property

Public Property Set QuestionText(ByVal NewValue As CHTML)
    Set pQuestionText = NewValue
End Property

Public Property Get GeneralFeedback() As CHTML
    Set GeneralFeedback = pGeneralFeedback
End Property

Public Property Set GeneralFeedback(ByVal NewValue As CHTML)
    Set pGeneralFeedback = NewValue
End Property

Public Property Get DefaultGrade() As Double
    DefaultGrade = pDefaultGrade
End Property

Public Property Let DefaultGrade(ByVal NewValue As Double)
    pDefaultGrade = NewValue
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

Public Property Get Units() As Collection
    Set Units = pUnits
End Property

Public Property Set Units(ByVal NewValue As Collection)
    Set pUnits = NewValue
End Property

Public Property Get Unitrole() As Long
    Unitrole = pUnitrole
End Property

Public Property Let Unitrole(ByVal NewValue As Long)
    pUnitrole = NewValue
    Select Case pUnitrole
        Case NumericalUnitroleOptions.Unitrole_onlynumerical
            pShowunits = NumericalShowunitsOptions.showunits_none
            pUnitgradingtype = NumericalUnitgradingtypeOptions.unitgradingtype_none
        Case NumericalUnitroleOptions.Unitrole_manynumerical
            pShowunits = NumericalShowunitsOptions.showunits_text
            pUnitgradingtype = NumericalUnitgradingtypeOptions.unitgradingtype_none
        Case NumericalUnitroleOptions.unitrole_unitgraded
            If pShowunits = NumericalShowunitsOptions.showunits_none Then
                pShowunits = NumericalShowunitsOptions.showunits_text
            End If
            If pUnitgradingtype = NumericalUnitgradingtypeOptions.unitgradingtype_none Then
                pUnitgradingtype = NumericalUnitgradingtypeOptions.unitgradingtype_FractionOfResponseGrade
            End If
    End Select
End Property

Public Property Get Unitgradingtype() As Long
    Unitgradingtype = pUnitgradingtype
End Property

Public Property Let Unitgradingtype(ByVal NewValue As Long)
    pUnitgradingtype = NewValue
    Unitrole = Unitrole 'calling Unitrole Let procedure to verify that unit grading settings are consistent
End Property

Public Property Get Unitpenalty() As Double
    Unitpenalty = pUnitpenalty
End Property

Public Property Let Unitpenalty(ByVal NewValue As Double)
    pUnitpenalty = NewValue
End Property

Public Property Get Showunits() As Long
    Showunits = pShowunits
End Property

Public Property Let Showunits(ByVal NewValue As Long)
    pShowunits = NewValue
    Unitrole = Unitrole 'calling Unitrole Let procedure to verify that unit grading settings are consistent
End Property

Public Property Get Unitposition() As Long
    Unitposition = pUnitposition
End Property

Public Property Let Unitposition(ByVal NewValue As Long)
    pUnitposition = NewValue
End Property

Public Property Get Answers() As Collection
    Set Answers = pAnswers
End Property

Public Property Set Answers(ByVal NewValue As Collection)
    Set pAnswers = NewValue
End Property

Public Property Get Hints() As Collection
    Set Hints = pHints
End Property

Public Property Set Hints(ByVal NewValue As Collection)
    Set pHints = NewValue
End Property

Public Function CleanText(Text As String)
    CleanText = RangeToHTML.TidyHTMLFormatng(Text)
End Function

Public Sub IQuestion_Export(ByRef DOM As MSXML2.DOMDocument60)
    Dim XMLRootElement As IXMLDOMElement
    Dim XMLQuestionElement As IXMLDOMElement
    Dim XMLElement1 As IXMLDOMElement
    Dim XMLElement2 As IXMLDOMElement
    Dim XMLText As IXMLDOMText
    Dim XMLAttr As IXMLDOMAttribute
    Dim Unit As CNumericalUnit
    Dim Answer As CNumericalAnswer
    Dim Hint As CShortanswerHint
    
    Set XMLRootElement = DOM.DocumentElement
    
    Set XMLQuestionElement = DOM.createElement("question")
    XMLRootElement.appendChild XMLQuestionElement
    Set XMLAttr = DOM.createAttribute("type")
    XMLAttr.NodeValue = "numerical"
    XMLQuestionElement.setAttributeNode XMLAttr
    
    Set XMLElement1 = DOM.createElement("name")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLElement2 = DOM.createElement("text")
    XMLElement1.appendChild XMLElement2
    Set XMLText = DOM.createTextNode(CleanText(pName))
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
    
    Set XMLElement1 = DOM.createElement("unitgradingtype")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(CStr(pUnitgradingtype))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("unitpenalty")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(Replace(Format(pUnitpenalty, "###0.0000000"), ",", "."))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("showunits")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(CStr(pShowunits))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("unitsleft")
    XMLQuestionElement.appendChild XMLElement1
    Set XMLText = DOM.createTextNode(CStr(pUnitposition))
    XMLElement1.appendChild XMLText
    
    Set XMLElement1 = DOM.createElement("units")
    XMLQuestionElement.appendChild XMLElement1
    For Each Unit In pUnits
        Unit.Export XMLElement1
    Next Unit
    
    For Each Answer In pAnswers
        Answer.Export XMLQuestionElement
    Next Answer
    
    For Each Hint In pHints
        Hint.Export XMLQuestionElement
    Next Hint
End Sub

Public Sub IQuestion_Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Answer As CNumericalAnswer
    Dim Hint As CShortanswerHint
    Dim Unit As CNumericalUnit
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
    
    Set pUnits = New Collection
    Set Nodes = Element.SelectNodes("units/unit")
    For Each Node In Nodes
        Set Unit = New CNumericalUnit
        Unit.Import Node
        pUnits.Add Unit
    Next

    pUnitpenalty = CDbl(Replace(Element.SelectSingleNode("unitpenalty/text()").NodeValue, ".", Separator))
    
    pUnitgradingtype = CLng(Element.SelectSingleNode("unitgradingtype/text()").NodeValue)
    
    pShowunits = CLng(Element.SelectSingleNode("showunits/text()").NodeValue)

    pUnitposition = CLng(Element.SelectSingleNode("unitsleft/text()").NodeValue)
    
    If pUnitgradingtype <> NumericalUnitgradingtypeOptions.unitgradingtype_none Then
       pUnitrole = NumericalUnitroleOptions.unitrole_unitgraded
    Else
        If pShowunits = NumericalShowunitsOptions.showunits_none Then
            pUnitrole = NumericalUnitroleOptions.Unitrole_onlynumerical
        Else
            pUnitrole = NumericalUnitroleOptions.Unitrole_manynumerical
        End If
    End If
    Unitrole = Unitrole
    
    Set pAnswers = New Collection
    Set Nodes = Element.SelectNodes("answer")
    For Each Node In Nodes
        Set Answer = New CNumericalAnswer
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


