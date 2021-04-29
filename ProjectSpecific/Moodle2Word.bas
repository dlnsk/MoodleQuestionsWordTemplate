Attribute VB_Name = "Moodle2Word"
' Модуль для импорта тестовых заданий из формата MoodleXML в формата АСТ-Тест
' Copyright 2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons «Attribution-ShareAlike» 4.0

Option Explicit

' Основная процедура. Создает новый документ c тестами из файла в формате MoodleXML
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
    
    Set Doc = ActiveDocument
    AppendQuestions Doc, Questions
    
    MsgBox strLoadFinished
    
    Exit Sub
error_handler:
    MsgBox strCreateGeneralError
End Sub

' Вставляет текстовую строку в конец документа. Возвращает фрагмент документа со вставленным текстом
Private Function AppendText(ByRef Doc As Word.Document, ByVal Text As String) As Word.Range
    Dim Range As Word.Range
    
    Set Range = Doc.Range
    Range.Collapse wdCollapseEnd
    Range.InsertAfter Text
    Range.End = Doc.Range.End
    Set AppendText = Range
End Function

' Вставляет HTML-фрагмент в конец документа
Private Function AppendHTML2(ByRef Doc As Word.Document, ByRef HTML As CHTML)
    Dim Range As Word.Range
    Dim rangeStart As Long
    Dim docNew As Word.Document
    Dim p As Variant
    
    
    HTMLToClipboard.HTMLToClipboard HTML
    
    'Создаем новый документ для обработки полученных после конвертации из html данных
    Set docNew = Documents.Add
    docNew.Range.Paste
    For Each p In docNew.Paragraphs
        If Len(p.Range.Text) = 1 Then
            p.Range.Delete
        End If
    Next
    If docNew.Paragraphs.Count > 1 Then
        Set Range = docNew.Range
        Range.End = docNew.Paragraphs(docNew.Paragraphs.Count - 1).Range.End
        With Range.Find
            .Text = "^p"
            .Replacement.Text = "^l"
            .Forward = True
            .Wrap = wdFindStop
        End With
        Range.Find.Execute Replace:=wdReplaceAll
    End If
    docNew.Range.Copy
    docNew.Close SaveChanges:=wdDoNotSaveChanges
    
    Set Range = Doc.Range
    Range.Collapse wdCollapseEnd
    Range.Paste
    
    Set AppendHTML2 = Range
End Function

' Вставляет HTML-фрагмент в конец документа. Возвращает фрагмент документа со вставленным текстом
' Для ускорения, фрагменты без HTML-тэгов вставляются прямым текстом
Private Function AppendHTML(ByRef Doc As Word.Document, ByRef HTML As CHTML, Style As String) As Word.Range
    Dim RegExp As Object
    Dim Range As Word.Range

    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    RegExp.Pattern = "<(""[^""]*""|'[^']*'|[^'"">])*>"
    If RegExp.Test(HTML.Text) Then
        Set Range = AppendHTML2(Doc, HTML)
    Else
        Set Range = AppendText(Doc, HTML.Text)
        Doc.Paragraphs.Add
    End If
    Range.Style = Style
    Set AppendHTML = Range
End Function

' Вставляет вопросы в документ
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
    
    ' Вопросы в экспортном файле идут вперемешку. Раскладываем их по категориям
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
    
    ' Сортируем категории по алфавиту
    Set Categories = GetSortedCollection(Categories)
    QuestionNumber = 1
    ' Экспортируем категорию, а затем все вопросы из этой категории
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
    Set Range = AppendText(Doc, CategoryName)
    Doc.Paragraphs.Last.Style = GIFT.STYLE_CATEGORY
    Doc.Paragraphs.Add
End Sub

' Вставляет вопрос в документ
Private Sub AppendQuestion(ByRef Doc As Word.Document, ByRef Question As Object, QuestionNumber As Long)
    Dim QuestionType As String
    Dim RegExp As Object
    Dim Range As Word.Range
    
    QuestionType = LCase(Typename(Question))
    Select Case QuestionType
        Case "cdescription"
            AppendDescription Doc, Question, QuestionNumber
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

' Добавляет в конец документа текст вопроса
' QuestionNumber - Номер вопроса.
' QuestionType - текстовое представление типа вопроса
' QuestionName - Название вопроса
' QuestionText - Текст вопроса
Private Sub AppendQuestionText(ByRef Doc As Word.Document, QuestionNumber As Long, QuestionType As String, QuestionGrade As Double, QuestionText As CHTML, Style As String)
    Dim Text As String
    Dim Range As Word.Range
    
'    Text = CStr(QuestionNumber) & ". " & QuestionType & ". Оценка: " & QuestionGrade
'    Set Range = AppendText(Doc, Text)
'    Range.Bold = True
'    Range.Italic = False
'    Range.Paragraphs.SpaceBefore = 12
'
'    Doc.Paragraphs.Add
    Set Range = AppendHTML(Doc, QuestionText, Style)
End Sub

Private Sub AppendDdmatch(ByRef Doc As Word.Document, ByRef Question As CDdmatch, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "На сопоставление"
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Общий комментарий к вопросу: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Комментарий к верному ответу: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Комментарий к неверному ответу: ")
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
        Set Range = AppendText(Doc, "Подвопрос № " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Subquestion
        Set Range = AppendText(Doc, "Ответ на подвопрос № " & CStr(SubquestionNumber) & ". ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Answer
    Else
        Set Range = AppendText(Doc, "Неверный ответ. ")
        Range.Bold = True
        Range.Italic = True
        AppendHTML Doc, Subquestion.Answer
    End If
End Sub

Private Sub AppendEssay(ByRef Doc As Word.Document, ByRef Question As CEssay, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "Эссе"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_ESSAYQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If

    If Question.Responsetemplate.Text <> "" Then
        AppendHTML Doc, Question.Responsetemplate, GIFT.STYLE_RIGHT_ANSWER
    End If

    If Question.Graderinfo.Text <> "" Then
        AppendHTML Doc, Question.Graderinfo, GIFT.STYLE_WRONG_ANSWER
    End If

End Sub

Private Sub AppendMatching(ByRef Doc As Word.Document, ByRef Question As CMatching, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "На сопоставление"
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_MATCHINGQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If

'    If Question.Correctfeedback.Text <> "" Then
'        Set Range = AppendText(Doc, "Комментарий к верному ответу: ")
'        Range.Bold = False
'        Range.Italic = True
'        AppendHTML Doc, Question.Correctfeedback
'    End If
'
'    If Question.Incorrectfeedback.Text <> "" Then
'        Set Range = AppendText(Doc, "Комментарий к неверному ответу: ")
'        Range.Bold = False
'        Range.Italic = True
'        AppendHTML Doc, Question.Incorrectfeedback
'    End If

    For I = 1 To Question.Subquestions.Count
        AppendMatchingSubquestion Doc, Question.Subquestions.Item(I), I
    Next
End Sub

Private Sub AppendMatchingSubquestion(ByRef Doc As Word.Document, ByRef Subquestion As CMatchingSubquestion, SubquestionNumber As Long)
    Dim Range As Word.Range
    
    Set Range = AppendHTML(Doc, Subquestion.Subquestion, GIFT.STYLE_LEFT_PAIR)
    Set Range = AppendText(Doc, Subquestion.Answer)
    Range.Style = GIFT.STYLE_RIGHT_PAIR
    Doc.Paragraphs.Add
End Sub

Private Sub AppendDescription(ByRef Doc As Word.Document, ByRef Question As CDescription, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_DESCRIPTIONQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If
End Sub

Private Sub AppendMultichoice(ByRef Doc As Word.Document, ByRef Question As CMultichoice, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
'    If Question.Singleanswer Then
'        QuestionType = "Множественный выбор. Один вариант ответа"
'    Else
'        QuestionType = "Множественный выбор. Несколько вариантов ответа"
'    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_MULTIPLECHOICEQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If

'    If Question.Correctfeedback.Text <> "" Then
'        Set Range = AppendText(Doc, "Комментарий к верному ответу: ")
'        Range.Bold = False
'        Range.Italic = True
'        AppendHTML Doc, Question.Correctfeedback
'    End If
'
'    If Question.Incorrectfeedback.Text <> "" Then
'        Set Range = AppendText(Doc, "Комментарий к неверному ответу: ")
'        Range.Bold = False
'        Range.Italic = True
'        AppendHTML Doc, Question.Incorrectfeedback
'    End If

    For I = 1 To Question.Answers.Count
        AppendMultichoiceAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendMultichoiceAnswer(ByRef Doc As Word.Document, ByRef Answer As CMultichoiceAnswer)
    Dim Range As Word.Range
    Dim FractionRange As Word.Range
    
    If Answer.Fraction <> 100 And Answer.Fraction <> 0 Then
        Set FractionRange = AppendText(Doc, CStr(Round(Answer.Fraction)) & "%")
        FractionRange.End = FractionRange.End - 1
    End If
    
    If Answer.Fraction > 0 Then
        AppendHTML Doc, Answer.Answer, GIFT.STYLE_RIGHT_ANSWER
    Else
        AppendHTML Doc, Answer.Answer, GIFT.STYLE_WRONG_ANSWER
    End If
    
    If Answer.Fraction <> 100 And Answer.Fraction <> 0 Then
        FractionRange.Style = GIFT.STYLE_ANSWERWEIGHT
    End If
    
    If Answer.Feedback.Text <> "" Then
        AppendHTML Doc, Answer.Feedback, GIFT.STYLE_FEEDBACK
    End If
End Sub

Private Sub AppendNumerical(ByRef Doc As Word.Document, ByRef Question As CNumerical, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "Числовой ответ"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_NUMERICALQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If

    For I = 1 To Question.Answers.Count
        AppendNumericalAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendNumericalAnswer(ByRef Doc As Word.Document, ByRef Answer As CNumericalAnswer)
    Dim Range As Word.Range
    Dim FractionRange As Word.Range
    Dim FinalAnswer As String
    
    If Answer.Fraction <> 100 And Answer.Fraction <> 0 Then
        Set FractionRange = AppendText(Doc, CStr(Round(Answer.Fraction)) & "%")
        FractionRange.End = FractionRange.End - 1
    End If
    
    If Answer.Fraction > 0 Then
        FinalAnswer = CStr(Answer.Answer)
    Else
        FinalAnswer = Strings.strIgnore
    End If
    
    If Answer.Tolerance > 0 Then
        FinalAnswer = FinalAnswer & ":" & CStr(Answer.Tolerance)
    End If
    Set Range = AppendText(Doc, FinalAnswer)
    If Answer.Fraction > 0 Then
        Range.Style = GIFT.STYLE_RIGHT_ANSWER
    Else
        Range.Style = GIFT.STYLE_WRONG_ANSWER
    End If
    
    If Answer.Fraction <> 100 And Answer.Fraction <> 0 Then
        FractionRange.Style = GIFT.STYLE_ANSWERWEIGHT
    End If
    Doc.Paragraphs.Add

    If Answer.Feedback.Text <> "" Then
        AppendHTML Doc, Answer.Feedback, GIFT.STYLE_FEEDBACK
    End If
End Sub

Private Sub AppendOrder(ByRef Doc As Word.Document, ByRef Question As COrder, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    QuestionType = "На упорядочивание"
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText
        
    If Question.Generalfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Общий комментарий к вопросу: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Generalfeedback
    End If

    If Question.Correctfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Комментарий к верному ответу: ")
        Range.Bold = False
        Range.Italic = True
        AppendHTML Doc, Question.Correctfeedback
    End If

    If Question.Incorrectfeedback.Text <> "" Then
        Set Range = AppendText(Doc, "Комментарий к неверному ответу: ")
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
    
    Set Range = AppendText(Doc, "Ответ №" & Subquestion.Order & ". ")
    Range.Bold = True
    Range.Italic = True
    Set Range = AppendHTML(Doc, Subquestion.Subquestion)
End Sub

Private Sub AppendShortanswer(ByRef Doc As Word.Document, ByRef Question As CShortanswer, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    
    ' Регистр пока не используется
    If Question.Usecase Then
        QuestionType = "Короткий ответ. Без учета регистра"
    Else
        QuestionType = "Короткий ответ. С учетом регистра"
    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=GIFT.STYLE_SHORTANSWERQ
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If

    For I = 1 To Question.Answers.Count
        AppendShortanswerAnswer Doc, Question.Answers.Item(I)
    Next
End Sub

Private Sub AppendShortanswerAnswer(ByRef Doc As Word.Document, ByRef Answer As CShortanswerAnswer)
    Dim FractionRange As Word.Range
    Dim Range As Word.Range
    
    If Answer.Fraction <> 100 Then
        Set FractionRange = AppendText(Doc, CStr(Round(Answer.Fraction)) & "%")
        FractionRange.End = FractionRange.End - 1
    End If
    
    Set Range = AppendText(Doc, Answer.Text)
    Range.Style = GIFT.STYLE_RIGHT_ANSWER
    
    If Answer.Fraction <> 100 Then
        FractionRange.Style = GIFT.STYLE_ANSWERWEIGHT
    End If
    Doc.Paragraphs.Add
    
    If Answer.Feedback.Text <> "" Then
        AppendHTML Doc, Answer.Feedback, GIFT.STYLE_FEEDBACK
    End If
End Sub

Private Sub AppendTruefalse(ByRef Doc As Word.Document, ByRef Question As CTrueFalse, QuestionNumber As Long)
    Dim QuestionType As String
    Dim Range As Word.Range
    Dim I As Long
    Dim IgnoreMessage As CHTML
    
    If Question.Answer = True Then
        QuestionType = GIFT.STYLE_TRUESTATEMENT
    Else
        QuestionType = GIFT.STYLE_FALSESTATEMENT
    End If
    
    AppendQuestionText Doc:=Doc, QuestionNumber:=QuestionNumber, QuestionType:=QuestionType, _
        QuestionGrade:=Question.Defaultgrade, QuestionText:=Question.QuestionText, Style:=QuestionType
        
    If Question.Generalfeedback.Text <> "" Then
        AppendHTML Doc, Question.Generalfeedback, GIFT.STYLE_FEEDBACK
    End If
    
    Set IgnoreMessage = New CHTML
    IgnoreMessage.Text = strIgnore

    If Question.TrueFeedback.Text <> "" Then
        AppendHTML Doc, IgnoreMessage, GIFT.STYLE_RIGHT_ANSWER
        AppendHTML Doc, Question.TrueFeedback, GIFT.STYLE_FEEDBACK
    End If

    If Question.FalseFeedback.Text <> "" Then
        AppendHTML Doc, IgnoreMessage, GIFT.STYLE_WRONG_ANSWER
        AppendHTML Doc, Question.FalseFeedback, GIFT.STYLE_FEEDBACK
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
