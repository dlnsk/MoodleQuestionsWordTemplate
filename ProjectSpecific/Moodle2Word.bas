Attribute VB_Name = "Moodle2Word"
' ������ ��� ������� �������� ������� �� ������� MoodleXML � ������� ���-����
' Copyright 2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons �Attribution-ShareAlike� 4.0

Option Explicit

' �������� ���������. ������� ����� �������� c ������� �� ����� � ������� MoodleXML
Public Sub Moodle2Word()
    Dim Filename As String
    Dim FilenameDialog As FileDialog
    Dim I As Integer
    Dim Questions As CQuestionCollection
    Dim Doc As Word.Document
    
'    On Error GoTo error_handler
    
    Set FilenameDialog = Application.FileDialog(msoFileDialogOpen)
    For I = 1 To FilenameDialog.Filters.Count
        If FilenameDialog.Filters.Item(I).Extensions = "*.xml" Then
            FilenameDialog.FilterIndex = I
        End If
    Next
    If Not FilenameDialog.Show Then
        Exit Sub
    End If
    Filename = FilenameDialog.SelectedItems(1)
    
    Set Questions = New CQuestionCollection
    Questions.Import Filename
    
    Set Doc = Documents.Add
    AppendQuestions Doc, Questions
    
    MsgBox strLoadFinished
    
    Exit Sub
error_handler:
    MsgBox strCreateGeneralError
End Sub

' ��������� ��������� ������ � ����� ���������. ���������� �������� ��������� �� ����������� �������
Private Function AppendText(ByRef Doc As Word.Document, ByVal Text As String) As Word.Range
    Dim Range As Word.Range
    
    Set Range = Doc.Range
    Range.Collapse wdCollapseEnd
    Range.InsertAfter Text
    Range.End = Doc.Range.End
    Range.Paragraphs.SpaceBefore = 0
    Range.Paragraphs.SpaceBeforeAuto = False
    Range.Paragraphs.SpaceAfter = 0
    Range.Paragraphs.SpaceAfterAuto = False
    Range.Bold = False
    Range.Italic = False
    Set AppendText = Range
End Function

' ��������� HTML-�������� � ����� ���������
Private Function AppendHTML2(ByRef Doc As Word.Document, ByRef HTML As CHTML)
    Dim Range As Word.Range
    
    HTMLToClipboard.HTMLToClipboard HTML
    Set Range = Doc.Range
    Range.Collapse wdCollapseEnd
    Range.Paste
    Range.End = Doc.Range.End
    Range.Paragraphs.SpaceBefore = 0
    Range.Paragraphs.SpaceBeforeAuto = False
    Range.Paragraphs.SpaceAfter = 0
    Range.Paragraphs.SpaceAfterAuto = False
    Set AppendHTML2 = Range
End Function

' ��������� HTML-�������� � ����� ���������. ���������� �������� ��������� �� ����������� �������
' ��� ���������, ��������� ��� HTML-����� ����������� ������ �������
Private Function AppendHTML(ByRef Doc As Word.Document, ByRef HTML As CHTML) As Word.Range
    Dim RegExp As Object

    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    RegExp.Pattern = "<(""[^""]*""|'[^']*'|[^'"">])*>"
    If RegExp.Test(HTML.Text) Then
        Set AppendHTML = AppendHTML2(Doc, HTML)
    Else
        Set AppendHTML = AppendText(Doc, HTML.Text)
        Doc.Paragraphs.Add
    End If
End Function

' ��������� ������� � ��������
Private Sub AppendQuestions(ByRef Doc As Word.Document, ByRef Questions As CQuestionCollection)
    Dim RegExp As Object
    Dim CategoryName As String
    Dim Category As CCategory
    Dim Categories As Collection
    Dim CategoryWithKey() As Variant
    Dim CategoryQuestions As Collection
    Dim I As Long
    Dim j As Long
    Dim QuestionNumber As Long
    
    ' ������� � ���������� ����� ���� ����������. ������������ �� �� ����������
    CategoryName = ""
    Set Categories = New Collection
    
    For I = 1 To Questions.Count
        If LCase(Typename(Questions.Item(I))) = "ccategory" Then
            Set RegExp = CreateObject("VBScript.RegExp")
            RegExp.Global = True
            RegExp.MultiLine = True
            RegExp.Pattern = "^\$[\s\S]*\$\/"
            Set Category = Questions.Item(I)
            CategoryName = RegExp.Replace(Category.Name, "")
        Else
            If CollectionKeyExists(Categories, "key" & CategoryName) Then
                Set CategoryQuestions = Categories.Item("key" & CategoryName)(1)
            Else
                Set CategoryQuestions = New Collection
                Categories.Add Array(CategoryName, CategoryQuestions), "key" & CategoryName
            End If
            CategoryQuestions.Add Questions.Item(I)
        End If
    Next
    
    ' ��������� ��������� �� ��������
    Set Categories = GetSortedCollection(Categories)
    QuestionNumber = 1
    ' ������������ ���������, � ����� ��� ������� �� ���� ���������
    For I = 1 To Categories.Count
        CategoryName = Categories.Item(I)(0)
        AppendCategory Doc, CategoryName
        Set CategoryQuestions = Categories.Item(I)(1)
        For j = 1 To CategoryQuestions.Count
            AppendQuestion Doc, CategoryQuestions.Item(j), QuestionNumber
            QuestionNumber = QuestionNumber + 1
        Next
    Next
End Sub

Private Sub AppendCategory(ByRef Doc As Word.Document, CategoryName As String)
    Dim RegExp As Object
    Dim CategoriesList As String
    Dim Categories() As String
    Dim I As Long
    Static LastCategories As Collection
    Dim LastCategoriesHit As Boolean
    Dim Range As Word.Range
    
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    RegExp.Pattern = "^\$[\s\S]*\$/"
    CategoriesList = RegExp.Replace(CategoryName, "")
    
    Categories = Split(Replace(CategoriesList, "//", "!@#DOUBLESLASH#@!"), "/")
    For I = 0 To UBound(Categories)
        Categories(I) = Replace(Categories(I), "!@#DOUBLESLASH#@!", "/")
    Next
    LastCategoriesHit = False
    If LastCategories Is Nothing Then
        Set LastCategories = New Collection
    End If
    For I = 0 To UBound(Categories)
        If I > LastCategories.Count - 1 Then
            LastCategoriesHit = True
        ElseIf Categories(I) <> LastCategories.Item(I + 1) Then
            LastCategoriesHit = True
        End If
        If LastCategoriesHit Then
            Set Range = AppendText(Doc, Categories(I))
            Doc.Paragraphs.Last.Style = "��������� " & CStr(I + 1)
            Doc.Paragraphs.Add
            LastCategoriesHit = True
        End If
    Next
    Set LastCategories = New Collection
    For I = 0 To UBound(Categories)
        LastCategories.Add Categories(I)
    Next
End Sub

' ��������� ������ � ��������
Private Sub AppendQuestion(ByRef Doc As Word.Document, ByRef Question As Object, QuestionNumber As Long)
    Dim QuestionType As String
    Dim RegExp As Object
    Dim Range As Word.Range
    
    QuestionType = LCase(Typename(Question))
    Select Case QuestionType
        Case "cddmatch"
            AppendDdmatch Doc, Question, QuestionNumber
        Case "cessay"
            AppendEssay Doc, Question, QuestionNumber
        Case "cmatching"
            AppendMatching Doc, Question, QuestionNumber
        Case "cmultichoice"
            AppendMultichoice Doc, Question, QuestionNumber
        Case "cnumerical"
            AppendNumerical Doc, Question, QuestionNumber
        Case "corder"
            AppendOrder Doc, Question, QuestionNumber
        Case "cshortanswer"
            AppendShortanswer Doc, Question, QuestionNumber
        Case "ctruefalse"
            AppendTruefalse Doc, Question, QuestionNumber
    End Select

    Set Range = Doc.Range
    Range.Collapse wdCollapseEnd
    Range.Select
End Sub

' ��������� � ����� ��������� ����� �������
' QuestionNumber - ����� �������.
' QuestionType - ��������� ������������� ���� �������
' QuestionName - �������� �������
' QuestionText - ����� �������
Private Sub AppendQuestionText(ByRef Doc As Word.Document, QuestionNumber As Long, QuestionType As String, QuestionGrade As Double, QuestionText As CHTML)
    Dim Text As String
    Dim Range As Word.Range
    
    Text = CStr(QuestionNumber) & ". " & QuestionType & ". ������: " & QuestionGrade
    Set Range = AppendText(Doc, Text)
    Range.Bold = True
    Range.Italic = False
    Range.Paragraphs.SpaceBefore = 12
    
    Doc.Paragraphs.Add
    AppendHTML Doc, QuestionText
End Sub

Private Sub AppendDdmatch(ByRef Doc As Word.Document, ByRef Question As CDdmatch, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "�� �������������"
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ��������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Incorrectfeedback
    End If

    For I = 1 To Question.Subquestions.Count
        AppendDdmatchSubquestion Doc, Question.Subquestions.Item(I), I
    Next
End Sub

Private Sub AppendDdmatchSubquestion(ByRef Doc As Word.Document, ByRef Subquestion As CDdmatchSubquestion, SubquestionNumber As Long)
    Dim Range As Word.Range
    
    If Subquestion.Subquestion.Text <> "" Then
        Set Range = AppendText(Doc, "��������� � " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Subquestion
        Set Range = AppendText(Doc, "����� �� ��������� � " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Answer
    Else
        Set Range = AppendText(Doc, "�������� �����. ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Answer
    End If
End Sub

Private Sub AppendEssay(ByRef Doc As Word.Document, ByRef Question As CEssay, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "����"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Graderinfo.Text <> "" Then
        Set Range = AppendText(Doc, "���������� ��� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Graderinfo
    End If

End Sub

Private Sub AppendMatching(ByRef Doc As Word.Document, ByRef Question As CMatching, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "�� �������������"
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ��������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Incorrectfeedback
    End If

    For I = 1 To Question.Subquestions.Count
        AppendMatchingSubquestion Doc, Question.Subquestions.Item(I), I
    Next
End Sub

Private Sub AppendMatchingSubquestion(ByRef Doc As Word.Document, ByRef Subquestion As CMatchingSubquestion, SubquestionNumber As Long)
    Dim Range As Word.Range
    
    If Subquestion.Subquestion.Text <> "" Then
        Set Range = AppendText(Doc, "��������� � " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Subquestion
        Set Range = AppendText(Doc, "����� �� ��������� � " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendText Doc, Subquestion.Answer
        Doc.Paragraphs.Add
    Else
        Set Range = AppendText(Doc, "�������� �����. ")
        Range.Bold = True
        Range.Italic = True
        AppendText Doc, Subquestion.Answer
        Doc.Paragraphs.Add
    End If
End Sub

Private Sub AppendMultichoice(ByRef Doc As Word.Document, ByRef Question As CMultichoice, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    If Question.Singleanswer Then
        QuestionType = "������������� �����. ���� ������� ������"
    Else
        QuestionType = "������������� �����. ��������� ��������� ������"
    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ��������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Incorrectfeedback
    End If

    For I = 1 To Question.Answers.Count
        AppendMultichoiceAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendMultichoiceAnswer(ByRef Doc As Word.Document, ByRef Answer As CMultichoiceAnswer)
    Dim Range As Word.Range
    
    If Answer.Fraction = 100 Then
        Set Range = AppendText(Doc, "������ �����. ")
    ElseIf Answer.Fraction <= 0 Then
        Set Range = AppendText(Doc, "�������� �����. ")
    Else
        Set Range = AppendText(Doc, "�������� ������ ����� (" & CStr(Round(Answer.Fraction)) & "%). ")
    End If
    Range.Bold = True
    Range.Italic = True
    Set Range = AppendHTML(Doc, Answer.Answer)
    
    If Answer.Feedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Answer.Feedback
    End If
End Sub

Private Sub AppendNumerical(ByRef Doc As Word.Document, ByRef Question As CNumerical, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "�������� �����"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    For I = 1 To Question.Answers.Count
        AppendNumericalAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendNumericalAnswer(ByRef Doc As Word.Document, ByRef Answer As CNumericalAnswer)
    Dim Range As Word.Range
    
    If Answer.Answer <> "*" And Answer.Fraction >= 0 Then
        If Answer.Fraction = 100 Then
            Set Range = AppendText(Doc, "������ �����. ")
        Else
            Set Range = AppendText(Doc, "�������� ������ ����� (" & CStr(Round(Answer.Fraction)) & "%). ")
        End If
        Range.Bold = True
        Range.Italic = True
        Set Range = AppendText(Doc, CStr(Answer.Answer))
        Range.Bold = False
        Range.Italic = False
        If Answer.Tolerance > 0 Then
            Set Range = AppendText(Doc, "�" & CStr(Answer.Tolerance))
            Range.Bold = False
            Range.Italic = False
        End If
        Doc.Paragraphs.Add
    End If
    
    If Answer.Answer <> "*" And Answer.Fraction >= 0 Then
        If Answer.Feedback.Text <> "" Then
            Set Range = AppendText(Doc, "����������� � ������: ")
            Range.Bold = False
            Range.Italic = True
            AppendHTML Doc, Answer.Feedback
        End If
    Else
        If Answer.Feedback.Text <> "" Then
            Set Range = AppendText(Doc, "����������� � ��������� ������: ")
            Range.Bold = False
            Range.Italic = True
            AppendHTML Doc, Answer.Feedback
        End If
    End If
End Sub

Private Sub AppendOrder(ByRef Doc As Word.Document, ByRef Question As COrder, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "�� ��������������"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ��������� ������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Incorrectfeedback
    End If

    For I = 1 To Question.Subquestions.Count
        AppendOrderSubquestion Doc, Question.Subquestions.Item(I)
    Next
End Sub

Private Sub AppendOrderSubquestion(ByRef Doc As Word.Document, ByRef Subquestion As COrderSubquestion)
    Dim Range As Word.Range
    
    Set Range = AppendText(Doc, "����� �" & Subquestion.Order & ". ")
    Range.Bold = True
    Range.Italic = True
    Set Range = AppendHTML(Doc, Subquestion.Subquestion)
End Sub

Private Sub AppendShortanswer(ByRef Doc As Word.Document, ByRef Question As CShortanswer, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    If Question.Usecase Then
        QuestionType = "�������� �����. ��� ����� ��������"
    Else
        QuestionType = "�������� �����. � ������ ��������"
    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    For I = 1 To Question.Answers.Count
        AppendShortanswerAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendShortanswerAnswer(ByRef Doc As Word.Document, ByRef Answer As CShortanswerAnswer)
    Dim Range As Word.Range
    
    If Answer.Text <> "*" And Answer.Fraction >= 0 Then
        If Answer.Fraction = 100 Then
            Set Range = AppendText(Doc, "������ �����. ")
        Else
            Set Range = AppendText(Doc, "�������� ������ ����� (" & CStr(Round(Answer.Fraction)) & "%). ")
        End If
        Range.Bold = True
        Range.Italic = True
        Set Range = AppendText(Doc, Answer.Text)
        Range.Bold = False
        Range.Italic = False
        Doc.Paragraphs.Add
    End If
    
    If Answer.Text <> "*" And Answer.Fraction >= 0 Then
        If Answer.Feedback.Text <> "" Then
            Set Range = AppendText(Doc, "����������� � ������: ")
            Range.Bold = False
            Range.Italic = True
            AppendHTML Doc, Answer.Feedback
        End If
    Else
        If Answer.Feedback.Text <> "" Then
            Set Range = AppendText(Doc, "����������� � ��������� ������: ")
            Range.Bold = False
            Range.Italic = True
            AppendHTML Doc, Answer.Feedback
        End If
    End If
End Sub

Private Sub AppendTruefalse(ByRef Doc As Word.Document, ByRef Question As CTrueFalse, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    If Question.Answer = True Then
        QuestionType = "�����/�������. ������ �����������"
    Else
        QuestionType = "�����/�������. �������� �����������"
    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
'    Set Range = AppendText(Doc, "�����: ")
'    Range.Bold = True
'    Range.Italic = True
'    If Question.Answer = True Then
'        AppendText Doc, "�����"
'    Else
'        AppendText Doc, "�������"
'    End If
'    Doc.Paragraphs.Add
'
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����� ����������� � �������: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Truefeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������ " & Chr(171) & "�����" & Chr(187) & ": ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Truefeedback
    End If

    If Question.Falsefeedback.Text <> "" Then
        Set Range = AppendText(Doc, "����������� � ������ " & Chr(171) & "�������" & Chr(187) & ": ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Truefeedback
    End If
End Sub

Private Function CollectionKeyExists(Col As Collection, Key As Variant) As Boolean
    Dim Item As Variant
    
    On Error GoTo Error
    CollectionKeyExists = True
    Item = Col.Item(Key)
    Exit Function
    
Error:
    CollectionKeyExists = False
End Function

Private Function GetSortedCollection(Col As Collection) As Collection
    Dim I As Long
    Dim Result As Collection
    
    Set Result = New Collection
    For I = 1 To Col.Count
        Result.Add Col.Item(I), "key" & I
    Next
    QuickSort Result, 1, Result.Count
    Set GetSortedCollection = Result
End Function

Private Sub QuickSort(Col As Collection, Lo As Long, Hi As Long)
    Dim MiddleElement As Variant
    Dim TempElement As Variant
    Dim TempLow As Long
    Dim TempHi As Long
    
    TempLow = Lo
    TempHi = Hi
    MiddleElement = Col.Item("key" & ((Lo + Hi) \ 2))(0)
    Do While TempLow <= TempHi
        Do While Col.Item("key" & TempLow)(0) < MiddleElement And TempLow < Hi
            TempLow = TempLow + 1
        Loop
        Do While MiddleElement < Col.Item("key" & TempHi)(0) And TempHi > Lo
            TempHi = TempHi - 1
        Loop
        If TempLow <= TempHi Then
            If TempLow < TempHi Then
                ' We cannot replace collection element directly, so we delete element and then create element with same key.
                ' This does not works with numbers, because in this case collection are automatically reordered, so we use string keys.
                TempElement = Col.Item("key" & TempLow)
                Col.Remove "key" & TempLow
                Col.Add Col.Item("key" & TempHi), "key" & TempLow
                Col.Remove "key" & TempHi
                Col.Add TempElement, "key" & TempHi
            End If
            TempLow = TempLow + 1
            TempHi = TempHi - 1
        End If
    Loop
    If Lo < TempHi Then QuickSort Col, Lo, TempHi
    If TempLow < Hi Then QuickSort Col, TempLow, Hi
End Sub
