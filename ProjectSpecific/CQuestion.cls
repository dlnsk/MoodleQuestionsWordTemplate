VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Implements IQuestion

Private pName As String
Private pQuestionText As CHTML
Private pGeneralFeedback As CHTML
Private pDefaultGrade As Double
Private pPenalty As Double
Private pHidden As Boolean
Private pAnswers As Collection
Private pHints As Collection

Private Sub Class_Initialize()
    pName = "-"
    Set pQuestionText = New CHTML
    Set pGeneralFeedback = New CHTML
    pDefaultGrade = 1
    pPenalty = 0.3333333
    pHidden = False
    Set pAnswers = New Collection
    Set pHints = New Collection
End Sub

Private Sub Class_Terminate()
    Set pQuestionText = Nothing
    Set pGeneralFeedback = Nothing
    Set pAnswers = Nothing
    Set pHints = Nothing
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

Public Property Get Hidden() As Boolean
    Hidden = pHidden
End Property

Public Property Let Hidden(ByVal NewValue As Boolean)
    pHidden = NewValue
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

Private Sub IQuestion_Export(DOM As MSXML2.IXMLDOMDocument3)

End Sub

Private Sub IQuestion_Import(Element As MSXML2.IXMLDOMElement)

End Sub
