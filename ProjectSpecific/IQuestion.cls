VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing moodle question. Is used as interface for implementing
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Public Property Get Name() As String
End Property

Public Property Let Name(ByVal NewValue As String)
End Property

Public Property Get QType() As String
End Property

Public Property Let QType(ByVal NewValue As String)
End Property

Public Property Get QuestionText() As CHTML
End Property

Public Property Let QuestionText(ByRef NewValue As CHTML)
End Property

Public Property Get GeneralFeedback() As CHTML
End Property

Public Property Let GeneralFeedback(ByRef NewValue As CHTML)
End Property

Public Property Get DefaultGrade() As Double
End Property

Public Property Let DefaultGrade(ByVal NewValue As Double)
End Property

Public Property Get Penalty() As Double
End Property

Public Property Let Penalty(ByVal NewValue As Double)
End Property

Public Property Get Hidden() As Boolean
End Property

Public Property Let Hidden(ByVal NewValue As Boolean)
End Property

Public Property Get Answers() As Collection
End Property

Public Property Let Answers(ByRef NewValue As Collection)
End Property

Public Property Get Hints() As Collection
End Property

Public Property Let Hints(ByRef NewValue As Collection)
End Property
