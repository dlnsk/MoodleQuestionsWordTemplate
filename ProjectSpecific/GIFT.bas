Attribute VB_Name = "GIFT"
' The MIT License
' Copyright (c) 2005 Mikko Rusama
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the "Software"), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
' to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
' BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
' GIFT Converter
' Version 1.9 Rusian, Updated 24.03.2009
' Author: Mikko Rusama (mikko.rusama@iki.fi)
' Translator and follower: Dmitry Pupinin (dlnsk@mail.ru)
'
' A macro for converting a Word document with questions to the native GIFT questionnaire
' format supported by Moodle (www.moodle.org). Questions are defined as different
' Word styles; style definitions are below.
'
' Supported question types are:
'  1. Multiple Choice Question
'  2. Matching Question
'  3. Short Answer Question
'  4. True-False Question (statements)
'  5. Numerical Question
'  6. Missing Word Question (only 1 right answer supported)
'  7. Description
'  8. Essay
'
' Question feedback is also supported as well as weighted answers for Multiple Choice
' Questions.
'
' Copyright 2004- SoberIT, Helsinki University of Technology
'
' Changes:
'  08.10.2004 Fixed decimal converter bug. Bug replaced commas with dots in the question choices.
'  17.05.2005 Menu created, translated to Russian (Dmitry Pupinin)
'  09.08.2005 Added support for multiline questions and answers (Dmitry Pupinin)
'  13.10.2005 Added support for MathType formula translator (remove MathType style) (Dmitry Pupinin)
'             Bug fixed - symbol "\" didn't escape (Dmitry Pupinin)
'  18.11.2005 Modified procedure for MathType support (Dmitry Pupinin)
'  22.12.2005 Weights will be removed only if some paragraphs was select (Dmitry Pupinin)
'  13.06.2006 Supporting inline images: VBA part (Alexey Karpenko), server side scripts (Dmitry Pupinin)
'  24.03.2009 Support Descriptions and Essay (Dmitry Pupinin)
'             Allow commenting questions
'             Allow tolerance in Numerical
'             Allow comments for wrong answers in Numerical
'  31.10.2014 Шаблон модифицирован для работы с Moodle 2.6+, рисунки вставляются по тексту в формате base64 (Molokov Petr)
'  17.02.2015 Шаблон модифицирован для экспорта формата Moodle XML (Molokov Petr)

'Option Explicit

'********************************************************
' Style definitions. The styles defined below are used in the conversion.
'********************************************************

' General purpose styles.
Public Const STYLE_FEEDBACK = "09. Комментарий"
Public Const STYLE_ANSWERWEIGHT = "0. ВесОтвета"
Const STYLE_NORMAL = "Обычный"
Const STYLE_MATHTYPE = "MTConvertedEquation"
Const STYLE_PARAGRAPH_DEFAULT = "Основной шрифт абзаца"

' Styles for multiple choice questions
Public Const STYLE_MULTIPLECHOICEQ = "06. ВопрМножВыбор"
Public Const STYLE_RIGHT_ANSWER = "06.1 ВерныйОтвет"
Public Const STYLE_WRONG_ANSWER = "06.2 НеверныйОтвет"

' Styles for matching pair questions
Public Const STYLE_MATCHINGQ = "03. ВопрНаСопоставление"
Public Const STYLE_DDMATCHQ = "03. ВопрНаСопоставлениеПеретаскиванием"
Public Const STYLE_LEFT_PAIR = "03.1 Утверждение"
Public Const STYLE_RIGHT_PAIR = "03.2 ОтветНаУтвержд"

' Styles for true-false questions
Public Const STYLE_TRUESTATEMENT = "02.1 ВерноеУтвержд"
Public Const STYLE_FALSESTATEMENT = "02.2 НеверноеУтвержд"

' Style for short answer question
Public Const STYLE_SHORTANSWERQ = "05. ВопрКороткийОтв"

' Style for numerical question
Public Const STYLE_NUMERICALQ = "04. ВопрЧисловой"

' Style for missing word question
Const STYLE_MISSINGWORDQ = "08. ВопрПропущСлово"
Const STYLE_BLANK_WORD = "08.1 Пропуск"

' Style for essay question
Public Const STYLE_ESSAYQ = "07. ВопрЭссе"

' Style for description question
Public Const STYLE_DESCRIPTIONQ = "00. Описание"

' Style for category
Public Const STYLE_CATEGORY = "01. Категория"

'********************************************************
' GIFT strings
'********************************************************
Const COPYRIGHT = "Copyright (c) 2004 Mikko Rusama, перевод и доработка Дмитрий Пупынин 2005-2010гг. Модификация Молоков Петр 2014-2015 г."
Const START_QUESTION_COMMENT = "Начало вопроса"
Const CATEGORY_COMMENT = "Выбираем категорию"

'********************************************************
' General constants and variable definitions
'********************************************************
#If VBA7 Then
    Declare PtrSafe Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal _
    lpBuffer As String) As Long
#Else
    Declare Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal _
    lpBuffer As String) As Long
#End If

'Номер курса
Dim course_id As String

'Адрес сайта
Dim sait As String

'Уникальный номер теста
Dim quiz_id As String

'Временная папка
Dim temp As String

'Папка для рисунков от теста
Dim outputPath As String

'Кодировка
Dim encoding As Long

Dim fname As String


' saves the current question type
Dim QuestionType As String

' Prefix for the filename
Const FILE_PREFIX = "Moodle_Questions_"

Public Type TParaType
    Para As Variant
    StyleName As String
    Processed As Boolean
'    InTable As Boolean
'    LeftIndent As Double
End Type

Public Type TParseText
    Text As String
    Name As String
    Shuffleanswers As Boolean
    Defaultgrade As Double
    file As Boolean 'для эссе
End Type

Public Type TNumericalAnswer
    Answer As Variant
    Tolerance As Double
    Fraction As Double
End Type

Public Type TMultichoiceAnswer
    Answer As CHTML
    Fraction As Double
    Singleanswer As Boolean
End Type

Public Type TShortanswerAnswer
    Text As String
    Fraction As Double
End Type

Public Type TMatchingSubquestion
    Subquestion As CHTML
    Answer As String
End Type

Dim paraStyles() As TParaType
Dim paraLast As Integer
'Dim Levels(10) As Double
'Dim lastLevel As Integer

'********************************************************
' GIFT question tags
'********************************************************
Const TAG_CATEGORY = "$CATEGORY: $course$/"
Const TAG_QUESTION_START = " {"
Const TAG_QUESTION_END = "}"
Const TAG_TRUE_CHOICE = "T"
Const TAG_FALSE_CHOICE = "F"
Const TAG_RIGHT_ANSWER = "="
Const TAG_NUMERICAL_QUESTION = "#"
Const TAG_WRONG_ANSWER = "~"
Const TAG_MATCHINGQ_ARROW = " -> "
Const TAG_WEIGHTED_ANSWER = "~%"
Const TAG_WEIGHTED_NUM_ANSWER = "=%"
Const TAG_FEEDBACK = "#"
Const TAG_FEEDBACK2 = "####"

Dim Questions As CQuestionCollection
    

Public Sub About()
    ufAbout.Show vbModal
End Sub


Public Sub ConvertFromACT()
Dim qName As String, qStr As String, TXT As String
Dim aRange As Range
Dim I As Integer
Dim formatIsOK As Boolean

    formatIsOK = False
    For Each Para In ActiveDocument.Paragraphs
        If Left(Para.Range.Text, 2) = "I:" Then
            formatIsOK = True
            Exit For
        End If
    Next Para
    If Not formatIsOK Then
        MsgBox "Формат содержимого не соответствует формату ACT!" + vbCr + _
               "Пожалуйста, поместите в документ вопросы в формате ACT и повторите конвертацию.", vbExclamation, "Ошибка!"
        Exit Sub
    End If
    
    RemoveFormatting
            ActiveDocument.Content.Find.Execute _
            FindText:="\{\{[0-9]@\}\}", ReplaceWith:="", MatchWildcards:=True, _
            Format:=False, Replace:=wdReplaceAll

    For Each Para In ActiveDocument.Paragraphs
        If Left(Para.Range.Text, 2) <> "I:" Then
            Para.Range.Delete
        Else
            Exit For
        End If
    Next Para

    For Each Para In ActiveDocument.Paragraphs
        prefix = Left(Para.Range.Text, 2)
        If prefix = "I:" Then
            TXT = Para.Range.Text
            TXT = Trim(right(TXT, Len(TXT) - 2))
            qName = Left(TXT, Len(TXT) - 1)
            Para.Range.Delete
        End If
        If prefix = "Q:" Then
            If Left(Para.Next.Range.Text, 2) = "S:" Then
                ActiveDocument.Range(Para.Next.Range.start, Para.Next.Range.start + 2).Delete
                ReplaceLineBreaks Para.Range
                prefix = "S:"
            Else
                ActiveDocument.Range(Para.Range.start, Para.Range.start + 2).Delete
                Para.Range.InsertBefore "M: "
            End If
        End If
        If prefix = "S:" Then
            Set nextPara = Para
            Do
                Set nextPara = nextPara.Next
                TXT = Left(nextPara.Range.Text, 2)
            Loop While TXT <> "+:" And TXT <> "-:" And TXT <> "1:" And TXT <> "L1"
            pStart = Para.Range.start
            pEnd = nextPara.Previous.Range.End
            Set aRange = ActiveDocument.Range(pStart, pEnd - 1)
            aRange.Select
            ReplaceLineBreaks aRange
            ActiveDocument.Range(aRange.start, aRange.start + 2).Delete
            If qName <> "" Then
                aRange.InsertBefore "M: ::" + qName + ":: "
                aRange.Select
            Else
                aRange.InsertBefore "M: "
            End If
        End If
    Next Para
    
    
    
    paraLast = -1
    ReDim paraStyles(100)
    For Each Para In ActiveDocument.Paragraphs
        paraLast = paraLast + 1
        Set paraStyles(paraLast).Para = Para
        If paraLast >= UBound(paraStyles) Then
            ReDim Preserve paraStyles(UBound(paraStyles) + 50)
        End If
    Next Para

    I = 0
    While I <= paraLast
        If Left(paraStyles(I).Para.Range.Text, 2) = "M:" Then
            j = I + 1
            wrongPresent = False
            qStyle = ""
            Do While Left(paraStyles(j).Para.Range.Text, 2) <> "M:"
                Set aRange = paraStyles(j).Para.Range
                aRange.Select
                
                If Left(aRange.Text, 2) = "+:" Then
                    aRange.Style = STYLE_RIGHT_ANSWER
                    With aRange.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = "#$#"
                        .Replacement.Text = "*"
                        .Execute Replace:=wdReplaceAll
                    End With

                    ActiveDocument.Range(aRange.start, aRange.start + 2).Delete
                End If
                
                If Left(aRange.Text, 2) = "-:" Then
                    aRange.Style = STYLE_WRONG_ANSWER
                    ActiveDocument.Range(aRange.start, aRange.start + 2).Delete
                    wrongPresent = True
                    qStyle = STYLE_MULTIPLECHOICEQ
                End If
                
                numAnsw = val(Left(aRange.Text, 3))
                If numAnsw > 0 Then
                    aRange.Style = STYLE_LEFT_PAIR
                    qStyle = STYLE_MATCHINGQ
                    If numAnsw < 10 Then
                        ActiveDocument.Range(aRange.start, aRange.start + 2).Delete
                    Else
                        ActiveDocument.Range(aRange.start, aRange.start + 3).Delete
                    End If
                    InsertAfterRange Trim(str(numAnsw)), STYLE_RIGHT_PAIR, aRange
                End If
                
                If Left(aRange.Text, 3) = "L1:" Then
                    firstLeft = j
                    qStyle = STYLE_MATCHINGQ
                End If
                
                If Left(aRange.Text, 1) = "L" Then
                    aRange.Style = STYLE_LEFT_PAIR
                    ActiveDocument.Range(aRange.start, aRange.start + 3).Delete
                End If
                If Left(aRange.Text, 1) = "R" Then
                    aRange.Style = STYLE_RIGHT_PAIR
                    ActiveDocument.Range(aRange.start, aRange.start + 3).Delete
                    Set aRange = paraStyles(j).Para.Range
                    aRange.Copy
                    aRange.Style = STYLE_NORMAL
                    Set aRange = paraStyles(firstLeft).Para.Range
                    aRange.Collapse Direction:=wdCollapseEnd
                    aRange.Paste
                    firstLeft = firstLeft + 1
                End If
                
                j = j + 1
                If j > paraLast Then
                    Exit Do
                End If
            Loop
            
            ActiveDocument.Range(paraStyles(I).Para.Range.start, paraStyles(I).Para.Range.start + 2).Delete
            If qStyle <> "" Then
                paraStyles(I).Para.Range.Style = qStyle
                If qStyle = STYLE_MULTIPLECHOICEQ Then
                    SetAnswerWeights paraStyles(I).Para.Range
                End If
            Else
                paraStyles(I).Para.Range.Style = STYLE_SHORTANSWERQ

                With paraStyles(I).Para.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "###"
                    .Replacement.Text = "_______"
                    .Execute Replace:=wdReplaceAll
                End With

            End If
    
        End If
        I = I + 1
    Wend
    
    For Each Para In ActiveDocument.Paragraphs
        If Para.Range.Style.NameLocal = STYLE_NORMAL Then
            Para.Range.Delete
        End If
    Next Para
End Sub


Sub ReplaceLineBreaks(ByRef aRange As Range)
    With aRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p"
        .Replacement.Text = "^l"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub ExamineExportToGIFT()

    RemoveMathTypeFormat ' remove MathType style because it disturb to make the conversion (dlnsk)
        
    ' Before conversion, document is checked for errors
    If CheckQuestionnaire = True Then
        MsgBox "Ошибок в структуре файла не найдено.", _
                vbInformation, "GIFT конвертер. " & COPYRIGHT
    End If
End Sub

' Main method. Converts the Word document to GIFT format.
Sub ExportToGIFT()
Dim saveDialog As Dialog

    Randomize
    
    'Make sure the document is saved before continuing
    If ActiveDocument.Saved = False Then
        MsgBox "Пожалуйста, перед продолжением сохраните этот документ в формате Word (*.doc, *.docx)!", _
                vbExclamation, "GIFT конвертер. " & COPYRIGHT
        Set saveDialog = Dialogs(wdDialogFileSaveAs)
        saveDialog.AddToMru = True
        ' Cancel pressed -> Exit
        If saveDialog.Show = 0 Then
            ' StatusBar = "Not saved"
            Exit Sub
        End If
    End If
    
    StatusBar = "Происходит преобразование в формат GIFT, пожалуйста, подождите..."
     
    RemoveMathTypeFormat ' remove MathType style because it disturb to make the conversion (dlnsk)
        
    ' Before conversion, document is checked for errors
    If CheckQuestionnaire = True Then
        
'        uExport.Show vbModal
'        If uExport.ok.Value = False Then Exit Sub
'
'        If uExport.cbUnicode.Value = True Then
'            encoding = 65001
'        Else
'            encoding = 1251
'        End If
        
        encoding = 65001
        StartConvert
     
    Else
        MsgBox "Пожалуйста, исправьте имеющиеся ошибки перед преобразованием.", vbCritical, "Ошибка"
    End If
    
    StatusBar = "Готово!"

End Sub

' Add Description Question to the end of the active document
Sub AddCategory()
    AddParagraphOfStyle STYLE_CATEGORY, "Укажите здесь название категории для следующих вопросов"
End Sub

' Add Description Question to the end of the active document
Sub AddDescriptionQ()
    AddParagraphOfStyle STYLE_DESCRIPTIONQ, "Укажите здесь описание теста. Это будет выглядеть как вопрос, на который не нужно отвечать"
End Sub

' Add Essay Question to the end of the active document
Sub AddEssayQ()
    AddParagraphOfStyle STYLE_ESSAYQ, "Напишите здесь задание для эссе"
End Sub

' Add Multiple Choice Question to the end of the active document
Sub AddMultipleChoiceQ()
    AddParagraphOfStyle STYLE_MULTIPLECHOICEQ, "Напишите здесь вопрос с множественным выбором"
End Sub

' Add Matching Question to the end of the active document
Sub AddMatchingQ()
    AddParagraphOfStyle STYLE_MATCHINGQ, "Напишите здесь вопрос на соответствие"
End Sub

' Add Numerical Question to the end of the active document
Sub AddNumericalQ()
    AddParagraphOfStyle STYLE_NUMERICALQ, "Напишите здесь числовой вопрос"
End Sub


' Add Short Answer Question to the end of the active document
Sub AddShortAnswerQ()
    AddParagraphOfStyle STYLE_SHORTANSWERQ, "Напишите здесь вопрос с коротким ответом в открытой форме"
End Sub

' Add Missing Word Question
Sub AddMissingWordQ()
    AddParagraphOfStyle STYLE_MISSINGWORDQ, "Напишите здесь вопрос с пропущенным словом. Не забудьте указать, какое слово должно быть скрыто!"
End Sub

' Add feedback
Sub AddQuestionFeedback()
    If Not (Selection.Range.Style = STYLE_LEFT_PAIR Or _
       Selection.Range.Style = STYLE_RIGHT_PAIR Or _
       Selection.Range.Style = STYLE_FEEDBACK Or _
       Selection.Range.Style = STYLE_CATEGORY) Then
        InsertAfterRange "Напишите здесь комментарий.", _
                         STYLE_FEEDBACK, Selection.Paragraphs(1).Range
    Else
        MsgBox "Нельзя добавить комментарий к этому элементу. ", vbExclamation
    End If
End Sub

' Add a true statement of the true-false question
Sub AddTrueStatement()
    AddParagraphOfStyle STYLE_TRUESTATEMENT, "Вопрос с выбором 'Верно/Неверно': напишите здесь ВЕРНОЕ утверждение"
End Sub

' Add a false statement of the true-false question
Sub AddFalseStatement()
    AddParagraphOfStyle STYLE_FALSESTATEMENT, "Вопрос с выбором 'Верно/Неверно': напишите здесь НЕВЕРНОЕ утверждение"
End Sub

' Marks the right answer
Public Sub MarkTrueAnswer()
    If Selection.Range.Style = STYLE_WRONG_ANSWER Then
        Selection.Range.Style = STYLE_RIGHT_ANSWER
    ElseIf Selection.Range.Style = STYLE_RIGHT_ANSWER Then
        Selection.Range.Style = STYLE_WRONG_ANSWER
    End If
End Sub

' Add a new paragraph with a specified style and text
' Inserted text is selected
Private Sub AddParagraphOfStyle(aStyle, Text)
Dim myRange As Range
    Set myRange = ActiveDocument.Content
    With myRange
        .EndOf Unit:=wdParagraph, Extend:=wdMove
        .InsertParagraphBefore
        .Move Unit:=wdParagraph, Count:=1
        .Style = aStyle
        .InsertBefore Text
        .Select
    End With
End Sub


' Special Characters ~ = # { } control the operation of the Moodle's GIFT filter and
' cannot be used as a normal text within questions. However, if you want to use one
' of these characters, for example to show a mathematical formula in a question, you need
' to escape the control characters, i.e. putting a backslash (\) before a control
' character.
Private Sub EscapeControlCharacters()
    With ActiveDocument.Content.Find
        ' need to escape "\" as written in documentation! (dlnsk)
        .Execute FindText:="\", ReplaceWith:="\\", _
        Format:=False, Replace:=wdReplaceAll
        
        .Execute FindText:="~", ReplaceWith:="\~", _
        Format:=False, Replace:=wdReplaceAll
        
        .Execute FindText:="=", ReplaceWith:="\=", _
        Format:=False, Replace:=wdReplaceAll

        .Execute FindText:="#", ReplaceWith:="\#", _
        Format:=False, Replace:=wdReplaceAll
    
        .Execute FindText:="{", ReplaceWith:="\{", _
        Format:=False, Replace:=wdReplaceAll
        
        .Execute FindText:="}", ReplaceWith:="\}", _
        Format:=False, Replace:=wdReplaceAll
        
        'Replace with "\n " (with space) for difference with \ne in TeX formulas (dlnsk)
        '.Execute FindText:="^l", ReplaceWith:="\n ", _;
        .Execute FindText:="^l", ReplaceWith:="<br>", _
        Format:=False, Replace:=wdReplaceAll
        
        
        'Амперсант
        .Text = "&"
        .Replacement.Text = "&amp;"
        .Execute Replace:=wdReplaceAll
        
' Щербина вставил:
        .Execute FindText:="***lt;", ReplaceWith:="&lt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:="***nolt;", ReplaceWith:="<", _
        Format:=False, Replace:=wdReplaceAll
    
        .Execute FindText:="***gt;", ReplaceWith:="&gt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:="***nogt;", ReplaceWith:=">", _
        Format:=False, Replace:=wdReplaceAll
' -------------
        
        'Нижний индекс
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = True
        .Text = ""

        .Font.Subscript = True
        .Replacement.Font.Subscript = False

        .Replacement.Text = "<sub>^&</sub>"
        .Execute Replace:=wdReplaceAll
        
        'Верхний индекс
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Font.Superscript = True
        .Replacement.Font.Superscript = False
        
        .Replacement.Text = "<sup>^&</sup>"
        .Execute Replace:=wdReplaceAll
        
        'Наклонный шрифт
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Font.Italic = True
        .Replacement.Font.Italic = False
        
        .Replacement.Text = "<i>^&</i>"
        .Execute Replace:=wdReplaceAll
        
        'Полужирный шрифт
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Font.Bold = True
        .Replacement.Font.Bold = False
        
        .Replacement.Text = "<b>^&</b>"
        .Execute Replace:=wdReplaceAll
        
        'Подчеркнутый шрифт
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Font.Underline = True
        .Replacement.Font.Underline = False
        
        .Replacement.Text = "<u>^&</u>"
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
        
        'Греческий алфавит и спец символы (избранное)
        .Text = ChrW(945)
        .Replacement.Text = "&alpha;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(946)
        .Replacement.Text = "&beta;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(947)
        .Replacement.Text = "&gamma;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(951)
        .Replacement.Text = "&eta;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(955)
        .Replacement.Text = "&lambda;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(956)
        .Replacement.Text = "&mu;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(957)
        .Replacement.Text = "&nu;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(960)
        .Replacement.Text = "&pi;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(961)
        .Replacement.Text = "&rho;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(965)
        .Replacement.Text = "&upsilon;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(966)
        .Replacement.Text = "&phi;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(969)
        .Replacement.Text = "&omega;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(176)
        .Replacement.Text = "&deg;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(177)
        .Replacement.Text = "&plusmn;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(8736)
        .Replacement.Text = "&ang;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(8776)
        .Replacement.Text = "&asymp;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(8800)
        .Replacement.Text = "&ne;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(8804)
        .Replacement.Text = "&le;"
        .Execute Replace:=wdReplaceAll
        .Text = ChrW(8805)
        .Replacement.Text = "&ge;"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' Moodle requires a dot (.) as a decimal separator. Thus, all comma separators need to
' be converted.
Private Sub ConvertDecimalSeparator(ByVal aRange As Range)
    aRange.Find.Execute FindText:=",", ReplaceWith:=".", _
    Format:=False, Replace:=wdReplaceAll
End Sub

' Remove all formatting from the document
Private Sub RemoveFormatting()
    Selection.WholeStory
    Selection.ClearFormatting
End Sub

' Remove style of MathType formula from the document (dlnsk)
Private Sub RemoveMathTypeFormat()
Dim sty As Style
    For Each sty In ActiveDocument.Styles
        If sty.NameLocal = STYLE_MATHTYPE Then
            sty.Delete
            Exit Sub
        End If
    Next sty
End Sub

' Count the number of paragraphs having the specified
' style in the defined range
Function CountStylesInRange(aStyle, startPoint, endPoint) As Integer
Dim counter As Integer, endP As Integer
Dim aRange As Range
   
   Set aRange = ActiveDocument.Range(start:=startPoint, End:=endPoint)
   endP = aRange.End  'store end point
   counter = 0
   With aRange.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        Do While .Execute(Wrap:=wdFindStop) = True
            If aRange.End > endP Then
                Exit Do
            Else
                counter = counter + 1    ' Increment Counter.
            End If
        Loop
        
    End With
    CountStylesInRange = counter
End Function

' Checks every paragraph in the document and defines Moodle tags
' accordingly. Empty paragraphs are deleted as well as paragraphs
' that are specified with an unknown/illegal style.
Private Sub ConvertToGIFT()
Dim I As Integer, curQ As Integer
Dim numv As Integer
Dim ncut As String

    EscapeControlCharacters ' escape all special characters
    
    numv = 1
    ncut = ""
    For I = 0 To paraLast
        
        If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Then
            numv = 1
        End If
        If paraStyles(I).StyleName = STYLE_NUMERICALQ Then
            numv = 2
        End If
        If paraStyles(I).StyleName = STYLE_MISSINGWORDQ Then
            FindBlanks (paraStyles(I).Para.Range)
            paraStyles(I).Processed = True
        Else
            If (paraStyles(I).StyleName = STYLE_RIGHT_ANSWER Or paraStyles(I).StyleName = STYLE_WRONG_ANSWER) And _
                        StyleFound(STYLE_ANSWERWEIGHT, paraStyles(I).Para.Range) = True Then
                If numv = 1 Then
                    InsertTextBeforeRange TAG_WEIGHTED_ANSWER, paraStyles(I).Para.Range
                Else
                    InsertTextBeforeRange TAG_WEIGHTED_NUM_ANSWER, paraStyles(I).Para.Range
                End If
            paraStyles(I).Processed = True
            End If
        End If
    Next I
    
    RemoveFormatting
    
    I = 0
    While I <= paraLast
        If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Or _
                paraStyles(I).StyleName = STYLE_MATCHINGQ Or _
                paraStyles(I).StyleName = STYLE_NUMERICALQ Or _
                paraStyles(I).StyleName = STYLE_SHORTANSWERQ Or _
                paraStyles(I).StyleName = STYLE_TRUESTATEMENT Or _
                paraStyles(I).StyleName = STYLE_FALSESTATEMENT Or _
                paraStyles(I).StyleName = STYLE_MISSINGWORDQ Or _
                paraStyles(I).StyleName = STYLE_DESCRIPTIONQ Or _
                paraStyles(I).StyleName = STYLE_ESSAYQ Or _
                paraStyles(I).StyleName = STYLE_CATEGORY Or _
                paraStyles(I).StyleName = STYLE_BLANK_WORD Then

            curQ = I
            
            strComment = "// " & START_QUESTION_COMMENT & ": " & paraStyles(I).StyleName & vbCr
            
            If paraStyles(I).StyleName = STYLE_CATEGORY Then    'Description havn't body but any other have
                strComment = "// " & CATEGORY_COMMENT & vbCr
            End If
            
            If I = 0 Then
                paraStyles(I).Para.Range.InsertBefore strComment
            Else
                If paraStyles(I - 1).StyleName <> STYLE_DESCRIPTIONQ And paraStyles(I - 1).StyleName <> STYLE_CATEGORY And paraStyles(I - 1).StyleName <> STYLE_MISSINGWORDQ Then
                    If ncut = "" Then
                        paraStyles(I).Para.Range.InsertBefore TAG_QUESTION_END & vbCr
                    Else
                        paraStyles(I).Para.Range.InsertBefore ncut & TAG_QUESTION_END & vbCr
                        ncut = ""
                    End If
                End If
                paraStyles(I).Para.Range.InsertBefore vbCr & strComment
            End If
            
            If paraStyles(I).StyleName = STYLE_CATEGORY Then
                InsertTextBeforeRange TAG_CATEGORY, paraStyles(I).Para.Range
            End If
            
            If paraStyles(I).StyleName <> STYLE_DESCRIPTIONQ And paraStyles(I).StyleName <> STYLE_CATEGORY And paraStyles(I).StyleName <> STYLE_MISSINGWORDQ Then    'Description havn't body but any other have
                InsertAfterBeforeCR TAG_QUESTION_START, paraStyles(I).Para.Range
            End If
            
            If paraStyles(I).StyleName = STYLE_NUMERICALQ Then
                InsertAfterBeforeCR TAG_NUMERICAL_QUESTION, paraStyles(I).Para.Range
            End If
            
            If paraStyles(curQ).StyleName = STYLE_TRUESTATEMENT Then
                InsertAfterBeforeCR TAG_TRUE_CHOICE, paraStyles(I).Para.Range
            End If
            If paraStyles(curQ).StyleName = STYLE_FALSESTATEMENT Then
                InsertAfterBeforeCR TAG_FALSE_CHOICE, paraStyles(I).Para.Range
            End If
'            If paraStyles(curQ).StyleName = STYLE_MISSINGWORDQ Then
'                FindBlanks (paraStyles(curQ).Para.Range)
'            End If

            If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                'InsertTextBeforeRange TAG_FEEDBACK2, paraStyles(i + 1).Para.Range
                ncut = TAG_FEEDBACK2 & paraStyles(I + 1).Para.Range.Text
                paraStyles(I + 1).Para.Range.Delete
                I = I + 1
            End If
            
        Else
            With paraStyles(I)
                ' Wrong answer found
                If .StyleName = STYLE_WRONG_ANSWER And (Not .Processed) Then
                    InsertTextBeforeRange TAG_WRONG_ANSWER, .Para.Range
                    
                ' Right answer found
                ElseIf .StyleName = STYLE_RIGHT_ANSWER And (Not .Processed) Then
''                    ' Weighted answer found
''                    If StyleFound(STYLE_ANSWERWEIGHT, .Para.Range) = True Then
''                        InsertTextBeforeRange TAG_WEIGHTED_ANSWER, .Para.Range
'                    ' Answer of the numerical question found
'                    If paraStyles(curQ).StyleName = STYLE_NUMERICALQ Then
'                        InsertTextBeforeRange TAG_RIGHT_NUMERICAL_ANSWER, .Para.Range
'                    ' Answer of the multiple choice question found
'                    Else
'                        InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
'                    End If
                    InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
                    
                ' left pair of the matching question
                ElseIf .StyleName = STYLE_LEFT_PAIR Then
                    InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
                    InsertAfterBeforeCR TAG_MATCHINGQ_ARROW, .Para.Range
                    
                ' right pair of the matching question
                ElseIf .StyleName = STYLE_RIGHT_PAIR Then
                    ' Do nothing
                    
                ' Question feedback
                ElseIf .StyleName = STYLE_FEEDBACK Then
                   InsertTextBeforeRange TAG_FEEDBACK, .Para.Range
                End If
            End With
        End If
        I = I + 1
    Wend
    If paraStyles(I - 1).StyleName <> STYLE_MISSINGWORDQ And paraStyles(I - 1).StyleName <> STYLE_DESCRIPTIONQ Then
        'InsertAfterBeforeCR vbCr & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                    If ncut = "" Then
                        InsertAfterBeforeCR vbCr & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                    Else
                        InsertAfterBeforeCR vbCr & ncut & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                        ncut = ""
                    End If
        
        
    End If
End Sub

' Check if the specified style is found in the range
Function StyleFound(aStyle, aRange As Range) As Boolean
   With aRange.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        .Execute Wrap:=wdFindStop
    End With
    'MsgBox "Style: " & aStyle & " Found: " & aRange.Find.Found
    StyleFound = aRange.Find.Found
End Function

' Removes answer weights from the selection
Public Sub RemoveAnswerWeightsFromTheSelection()
    If Selection.start <> Selection.End Then
        With Selection.Find
            .ClearFormatting
            .Style = STYLE_ANSWERWEIGHT
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Format = True
            .Execute Replace:=wdReplaceAll
        End With

        
    Else
        MsgBox "Необходимо выделить абзацы, в которых будут убраны веса!", vbExclamation, "Ошибка!"
    End If
End Sub

' Checks the questionnaire.
' Returns true if everyhing is fine, otherwise false
Function CheckQuestionnaire() As Boolean
Dim I As Integer

    paraLast = -1
    ReDim paraStyles(100)
    For Each Para In ActiveDocument.Paragraphs
        ' Delete empty paragraphs
        If Para.Range = vbCr And Para.Range.Style.NameLocal <> STYLE_LEFT_PAIR Then
            Para.Range.Delete ' delete all empty paragraphs
        Else
            paraLast = paraLast + 1
            Set paraStyles(paraLast).Para = Para
            paraStyles(paraLast).StyleName = Para.Range.Style.NameLocal
'            paraStyles(paraLast).LeftIndent = Para.LeftIndent

            If paraLast = UBound(paraStyles) Then
                ReDim Preserve paraStyles(UBound(paraStyles) + 50)
            End If
        End If
    Next Para
    
    I = 0
    CheckQuestionnaire = True
    While I <= paraLast
        If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Or _
            paraStyles(I).StyleName = STYLE_MATCHINGQ Or _
            paraStyles(I).StyleName = STYLE_NUMERICALQ Or _
            paraStyles(I).StyleName = STYLE_SHORTANSWERQ Or _
            paraStyles(I).StyleName = STYLE_TRUESTATEMENT Or _
            paraStyles(I).StyleName = STYLE_FALSESTATEMENT Or _
            paraStyles(I).StyleName = STYLE_MISSINGWORDQ Or _
            paraStyles(I).StyleName = STYLE_DESCRIPTIONQ Or _
            paraStyles(I).StyleName = STYLE_ESSAYQ Or _
            paraStyles(I).StyleName = STYLE_CATEGORY Or _
            paraStyles(I).StyleName = STYLE_BLANK_WORD Then
          I = CheckQuestion(I) - 1
          If I = -1 Then
            CheckQuestionnaire = False
            Exit Function
          End If
        End If
        I = I + 1
    Wend
End Function

Function CheckQuestion(qIndex As Integer) As Integer
Dim current, rightsCount As Integer

    haveWrongAnswer = False
    rightsCount = 0
    
    current = qIndex + 1
    If paraStyles(current).StyleName = STYLE_FEEDBACK And _
            paraStyles(qIndex).StyleName <> STYLE_CATEGORY Then
        current = current + 1
    End If
    
    While current <= paraLast
        If paraStyles(current).StyleName = STYLE_FEEDBACK And _
           paraStyles(current - 1).StyleName = STYLE_FEEDBACK Then
                paraStyles(current).Para.Range.Select
                MsgBox "Ошибка! Не допускается несколько комментариев подряд!", vbExclamation, "Ошибка!"
                CheckQuestion = 0
                Exit Function
        End If
        
        If paraStyles(current).StyleName = STYLE_MULTIPLECHOICEQ Or _
            paraStyles(current).StyleName = STYLE_MATCHINGQ Or _
            paraStyles(current).StyleName = STYLE_NUMERICALQ Or _
            paraStyles(current).StyleName = STYLE_SHORTANSWERQ Or _
            paraStyles(current).StyleName = STYLE_TRUESTATEMENT Or _
            paraStyles(current).StyleName = STYLE_FALSESTATEMENT Or _
            paraStyles(current).StyleName = STYLE_MISSINGWORDQ Or _
            paraStyles(current).StyleName = STYLE_DESCRIPTIONQ Or _
            paraStyles(current).StyleName = STYLE_ESSAYQ Or _
            paraStyles(current).StyleName = STYLE_CATEGORY Or _
            paraStyles(current).StyleName = STYLE_BLANK_WORD Then     'Достигли следующего вопроса, значит в этом все в порядке
                If paraStyles(qIndex).StyleName = STYLE_MULTIPLECHOICEQ And rightsCount = 0 Then
                    paraStyles(qIndex).Para.Range.Select
                    MsgBox "Ошибка! Не выбрано ни одного правильного ответа в этом вопросе!", vbExclamation, "Ошибка!"
                    CheckQuestion = 0
                Else
                    CheckQuestion = current
                End If
                Exit Function
        End If
        
        If paraStyles(qIndex).StyleName = STYLE_CATEGORY Then
                paraStyles(current).Para.Range.Select
                MsgBox "Ошибка! Данный элемент не допустим после выбора категории!", vbExclamation, "Ошибка!"
                CheckQuestion = 0
                Exit Function
        End If
        
        If paraStyles(qIndex).StyleName = STYLE_MULTIPLECHOICEQ Then
            If Not (paraStyles(current).StyleName = STYLE_RIGHT_ANSWER Or _
                paraStyles(current).StyleName = STYLE_WRONG_ANSWER Or _
                paraStyles(current).StyleName = STYLE_ANSWERWEIGHT Or _
                paraStyles(current).StyleName = STYLE_FEEDBACK) Then
               paraStyles(current).Para.Range.Select
               MsgBox "Ошибка! Недопустимый элемент:" + paraStyles(current).StyleName, vbExclamation, "Ошибка!"
               CheckQuestion = 0
               Exit Function
            End If
            If paraStyles(current).StyleName = STYLE_RIGHT_ANSWER Then
                rightsCount = rightsCount + 1
            End If
        End If
        
        If paraStyles(qIndex).StyleName = STYLE_MATCHINGQ Then
            If Not (paraStyles(current).StyleName = STYLE_LEFT_PAIR Or _
                paraStyles(current).StyleName = STYLE_RIGHT_PAIR Or _
                paraStyles(current).StyleName = STYLE_CATEGORY) Then
               paraStyles(current).Para.Range.Select
               MsgBox "Ошибка! Недопустимый элемент: " + paraStyles(current).StyleName, vbExclamation, "Ошибка!"
               CheckQuestion = 0
               Exit Function
            Else
               If (paraStyles(current).StyleName = STYLE_LEFT_PAIR And _
                                Not (paraStyles(current - 1).StyleName = STYLE_RIGHT_PAIR Or _
                                     paraStyles(current - 1).StyleName = STYLE_FEEDBACK Or _
                                     paraStyles(current - 1).StyleName = STYLE_MATCHINGQ) _
                        Or _
                        paraStyles(current).StyleName = STYLE_RIGHT_PAIR And _
                                Not (paraStyles(current - 1).StyleName = STYLE_LEFT_PAIR Or _
                                     paraStyles(current - 1).StyleName = STYLE_FEEDBACK Or _
                                     paraStyles(current - 1).StyleName = STYLE_MATCHINGQ)) Then
                    paraStyles(current).Para.Range.Select
                    MsgBox "Ошибка! Утверждения и ответы должны чередоваться!" & vbCr & _
                           "Однако, если вы хотите добавить лишний ответ можно сделать Утверждение пустым.", vbExclamation, "Ошибка!"
                    CheckQuestion = 0
                    Exit Function
                End If
            End If
        End If
        
        If paraStyles(qIndex).StyleName = STYLE_SHORTANSWERQ Then 'paraStyles(qIndex).StyleName = STYLE_NUMERICALQ Or
            If Not (paraStyles(current).StyleName = STYLE_RIGHT_ANSWER Or _
                paraStyles(current).StyleName = STYLE_ANSWERWEIGHT Or _
                paraStyles(current).StyleName = STYLE_FEEDBACK) Then
               paraStyles(current).Para.Range.Select
               MsgBox "Ошибка! Недопустимый элемент:" + paraStyles(current).StyleName, vbExclamation, "Ошибка!"
               CheckQuestion = 0
               Exit Function
            End If
        End If
        
        If paraStyles(qIndex).StyleName = STYLE_NUMERICALQ Then
            If Not (paraStyles(current).StyleName = STYLE_RIGHT_ANSWER Or _
                    paraStyles(current).StyleName = STYLE_WRONG_ANSWER Or _
                    paraStyles(current).StyleName = STYLE_ANSWERWEIGHT Or _
                    paraStyles(current).StyleName = STYLE_FEEDBACK) Then
               paraStyles(current).Para.Range.Select
               MsgBox "Ошибка! Недопустимый элемент:" + paraStyles(current).StyleName, vbExclamation, "Ошибка!"
               CheckQuestion = 0
               Exit Function
            End If
            If paraStyles(current).StyleName = STYLE_WRONG_ANSWER Then
                If current < paraLast And paraStyles(current + 1).StyleName <> STYLE_FEEDBACK Then
                    paraStyles(current).Para.Range.Select
                    MsgBox "Ошибка! Неверный ответ допускается только если после него находится комментарий!" + vbCr + _
                           "Сам ответ игнорируется, а комментарий будет использоваться для всех неверных ответов.", vbExclamation, "Ошибка!"
                    CheckQuestion = 0
                    Exit Function
                End If
                If haveWrongAnswer = True Then
                    paraStyles(current).Para.Range.Select
                    MsgBox "Ошибка! Допускается только один неверный ответ!", vbExclamation, "Ошибка!"
                    CheckQuestion = 0
                    Exit Function
                End If
                haveWrongAnswer = True
            End If
            If paraStyles(current).StyleName = STYLE_RIGHT_ANSWER And haveWrongAnswer = True Then
                paraStyles(current).Para.Range.Select
                MsgBox "Ошибка! Последним должен быть неверный ответ!", vbExclamation, "Ошибка!"
                CheckQuestion = 0
                Exit Function
            End If
        End If
        
        current = current + 1
    Wend
    CheckQuestion = current
End Function

' Check the numeric answer. Note, not checking all valid GIFT formats.
Function CheckNumericAnswer(aRange As Range) As Boolean
Dim Response As Integer
    
    CheckNumericAnswer = True ' By default OK
    
    ' Search for the error margin separator
        aRange.Find.Execute FindText:=":", Format:=False
    
    If aRange.Find.Found = False And IsNumeric(aRange) = False Then
            aRange.Select
            Response = MsgBox("Это верный числовой ответ?" & vbCr & _
                       "Ваш ответ: " & aRange, vbYesNo, "Ошибка?")
            If Response = vbNo Then CheckNumericAnswer = False
    End If
End Function

' Inserts text before the specified range. A new paragraph is inserted.
Sub InsertTextBeforeRange(ByVal Text As String, ByVal aRange As Range)
    With aRange
'         .Style = STYLE_NORMAL
        .InsertBefore Text
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Вырезание коментария из текста
Sub CutFeedbackQuestion(ByVal aRange As Range)
    With aRange
        .Cut
        .Move Unit:=wdParagraph, Count:=-1
    End With

End Sub

' Inserts text having trailing VbCr before the range
Sub InsertQuestionEndTag(endPoint As Integer)
Dim aRange As Range

    Set aRange = ActiveDocument.Range(endPoint - 1, endPoint)
    With aRange
        .InsertBefore vbCr & TAG_QUESTION_END
        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Inserts text at the end of the paragraph before the trailing VbCr
Sub InsertAfterBeforeCR(ByVal Text As String, ByVal aRange As Range)
    aRange.End = aRange.End - 1 ' insert text before cr
    With aRange
        .InsertAfter Text
'        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With
End Sub

' Inserts text at the end of the paragraph
Sub InsertAfterRange(ByVal Text As String, aStyle, ByVal aRange As Range)
    With aRange
        .EndOf Unit:=wdParagraph, Extend:=wdMove
        .InsertParagraphBefore
        .Move Unit:=wdParagraph, Count:=-1
        .Style = aStyle
        .InsertBefore Text
        .Select
    End With
End Sub

' Set the answer weights of multiple choice questions.
Public Sub SetAnswerWeights(Optional aRange As Range = Empty)
    Dim startPoint As Long, endPoint As Long, rightCount As Long, wrongCount As Long
    Dim rightScore As Double, wrongScore As Double
    Dim QuestionRange As Range, tmpRange As Range
    
    Set tmpRange = aRange
    If Not IsObjectValid(aRange) Then     'Argument not passed (menu item pressed)
        Set aRange = Selection.Range
    End If
    
    If aRange.Style = STYLE_MULTIPLECHOICEQ Or aRange.Style = STYLE_NUMERICALQ Then
        startPoint = aRange.Paragraphs(1).Range.start
        rightCount = 0
        wrongCount = 0
        aRange.Move Unit:=wdParagraph, Count:=1
        
        Do While aRange.Style = STYLE_RIGHT_ANSWER Or _
              aRange.Style = STYLE_WRONG_ANSWER Or _
              aRange.Style = STYLE_FEEDBACK Or _
              aRange.Style = STYLE_ANSWERWEIGHT
                          
            'Delete empty paragraphs
            If aRange.Paragraphs(1).Range = vbCr Then
                aRange.Paragraphs(1).Range.Delete ' delete all empty paragraphs
            ' Remove old answer weights
            ElseIf aRange.Style = STYLE_ANSWERWEIGHT Then
                With aRange.Find
                    .ClearFormatting
                    .Style = STYLE_ANSWERWEIGHT
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Format = True
                    .Execute Replace:=wdReplaceOne
                End With

            End If
            
            ' Count the number of right and wrong answers
            If aRange.Style = STYLE_RIGHT_ANSWER Then
                rightCount = rightCount + 1
            ElseIf aRange.Style = STYLE_WRONG_ANSWER Then
                wrongCount = wrongCount + 1
            End If
            
           If aRange.Paragraphs(1).Range.End = ActiveDocument.Range.End Then
                endPoint = aRange.Paragraphs(1).Range.End
                Exit Do
           Else
                aRange.Move Unit:=wdParagraph, Count:=1
                endPoint = aRange.Paragraphs(1).Range.start
           End If
        Loop
        
        Set QuestionRange = ActiveDocument.Range(startPoint, endPoint)
        
        If rightCount < 1 Then
            QuestionRange.Select
            MsgBox "Не определен верный ответ.", vbExclamation, "Ошибка!"
'        ElseIf rightCount = 1 Then
'            If Not IsObjectValid(tmpRange) Then     'Argument not passed (menu item pressed)
'                MsgBox "Нет необходимости указывать вес, если верным является только один ответ.", vbInformation, "Информация"
'            End If
'            Exit Sub 'No need add weight for one right answer
        Else
            ' Calculate the right and wrong scores
            rightScore = Round(100 / rightCount, 3)
            ' MODIFY the default scoring principle for wrong answers if necessary
            wrongScore = -rightScore
            
            AddAnswerWeights QuestionRange, rightScore, wrongScore
        End If
    Else
        MsgBox "Установите курсор в абзац с вопросом типа 'Множественный выбор' или 'Числовой'" & vbCr & _
               "и попробуйте еще раз.", vbExclamation, "Ошибка!"
        ' Find the previous paragraph having the style of multiple choice question.
        With aRange.Find
            .ClearFormatting
            .Text = ""
            .Style = STYLE_MULTIPLECHOICEQ
            .Forward = False
            .Format = True
            .MatchCase = False
            .Execute
        End With

    End If
End Sub

' Marks the blank word
Public Sub MarkBlankWord()
Dim aRange As Range, lastChar As Range
Dim endPos As Integer
    
    Set lastChar = ActiveDocument.Range(start:=Selection.Words(1).End - 1, End:=Selection.Words(1).End)
    If lastChar = " " Then
        endPos = Selection.Words(1).End - 1
    Else
        endPos = Selection.Words(1).End
    End If
    Set aRange = ActiveDocument.Range(start:=Selection.Words(1).start, End:=endPos)
    If Selection.Words(1).Style = STYLE_BLANK_WORD Then
        aRange.Select
        Selection.ClearFormatting
    Else
        'RTrim(ActiveDocument.Words(1)).Style = STYLE_BLANK_WORD
        aRange.Style = STYLE_BLANK_WORD
    End If
End Sub

' Find all the
Private Sub FindBlanks(aRange As Range)
Dim endPoint As Long
    'Set aRange = Selection.Paragraphs(1).Range
    
    
    'ActiveDocument.Content
    endPoint = aRange.End
    With aRange.Find
        .ClearFormatting
        .Style = STYLE_BLANK_WORD
        If .Execute(FindText:="", Forward:=True, Format:=True) = True Then
            With .Parent
                .InsertBefore "{="
                .InsertAfter "}"
                .Move Unit:=wdWord, Count:=1
            End With
        End If
    End With

End Sub

' Insert answer weights
Private Sub AddAnswerWeights(ByVal aRange As Range, rightScore, wrongScore)
Dim Para As Paragraph

    ' Check each paragraph at a time and specify needed tags
    For Each Para In aRange.Paragraphs
        ' Check if empty paragraph
        If Para.Range = vbCr Then
            Para.Range.Delete ' delete all empty paragraphs
        ElseIf Para.Range.Style = STYLE_RIGHT_ANSWER Then
            InsertAnswerWeight rightScore, Para.Range
        ElseIf Para.Range.Style = STYLE_WRONG_ANSWER Then
            InsertAnswerWeight wrongScore, Para.Range
        End If
    Next Para

End Sub

' Inserts text at the end of the chapter before the trailing VbCr
Private Sub InsertAnswerWeight(Score, ByVal aRange As Range)
Dim startPoint As Long, scoreString As String
Dim newRange As Range

    startPoint = aRange.start
    scoreString = "" & Score & "%"
    aRange.InsertBefore scoreString
    Set newRange = ActiveDocument.Range(start:=startPoint, End:=startPoint + Len(scoreString))
    newRange.Style = STYLE_ANSWERWEIGHT
    'Moodle requires that the decimal separator is dot, not comma.
    ConvertDecimalSeparator newRange
    
End Sub

'Создаем временную папку и каталоги в ней
Sub create_temp_path()
    
Const BufferLength& = 256
Dim Result&, Buffer$, Dirname$
    
    Buffer$ = Space$(BufferLength)
    Result& = GetTempPath(BufferLength&, Buffer$)
    Dirname$ = Left$(Buffer$, Result&)
    
    temp = Dirname$
    temp = temp & getrandom()
    
    'Каталог найден, удаляем его
    If Dir(temp, vbDirectory) <> "" Then
        RmDir temp
    End If
    MkDir temp
    
    outputPath = temp & "\release"
    MkDir outputPath
    
End Sub

'Генерит уникальный номер
Function getrandom() As String
    getrandom = Replace(str(Round(6553 * Rnd) + 1) & "_" & str(Round(6553 * Rnd) + 1) & "_" & str(Round(6553 * Rnd) + 1), " ", "")
End Function

Public Sub StartConvert()
Dim curPath As String
Dim curName As String, fname As String, fnameXML As String
Dim start As Integer

' Щербина вставил:
    With ActiveDocument.Content.Find
        .Execute FindText:="\\<", ReplaceWith:="\***lt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:="\<", ReplaceWith:="***nolt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:="<", ReplaceWith:="***lt;", _
        Format:=False, Replace:=wdReplaceAll
    
        .Execute FindText:="\\>", ReplaceWith:="\***gt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:="\>", ReplaceWith:="***nogt;", _
        Format:=False, Replace:=wdReplaceAll
        .Execute FindText:=">", ReplaceWith:="***gt;", _
        Format:=False, Replace:=wdReplaceAll
    End With
' -------------

    'Сохраняем текущею папку
    curPath = ActiveDocument.Path
    
    'Сохраняем название документа
    curName = ActiveDocument.Name
    
    'Создаем временную папку
    Call create_temp_path
'    Call create_quiz_id
    
    ChangeFileOpenDirectory temp

    Dim s As InlineShape
    For Each s In ActiveDocument.InlineShapes
        s.AlternativeText = getrandom
    Next
    
    ActiveDocument.WebOptions.AllowPNG = True
    ActiveDocument.WebOptions.PixelsPerInch = 96
    
    ActiveDocument.SaveAs Filename:="temp_html.htm", Fileformat:= _
        wdFormatFilteredHTML, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
    
    ActiveDocument.SaveAs Filename:="temp_doc.doc", Fileformat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
    
    parseHTML temp, "temp_html.htm", outputPath
    'ConvertToGIFT
    ConvertToMoodleXML
    
    RemoveFormatting
    
   ' ActiveDocument.SaveAs filename:=(curPath & "\gift_format.txt"), FileFormat:= _
        wdFormatText, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, encoding:=encoding
    
    start = InStrRev(curName, ".")
    If start > 0 Then
        fname = Left(curName, start) & "gift_format.txt"
        fnameXML = Left(curName, start) & "MoodleXML.xml"
    End If
            
        
    ActiveDocument.SaveAs Filename:=(curPath & "\" & fname), Fileformat:= _
        wdFormatText, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, encoding:=encoding
    
    ActiveDocument.SaveAs Filename:="temp_doc.doc", Fileformat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
    
    
    ChangeFileOpenDirectory outputPath & "\"
    
 '   start = InStrRev(curName, ".")
 '   If start > 0 Then
 '       fname = Left(curName, start) & "zip"
 '   End If
    
    'Shell ("pkzipc.exe -add -attr=all -dir -nozip """ & curPath & "\" & fname & """ *.*")

    'XML =====================================================
    Questions.Export curPath & "\" & fnameXML
   
    MsgBox "Конвертация закончена, было создано два файла:" & vbCr & curPath & "\" & fname & vbCr & curPath & "\" & fnameXML
    
    Dim Doc As Document
    Set Doc = ActiveDocument
 
    ChangeFileOpenDirectory curPath
    Documents.Open curName
    Doc.Close SaveChanges:=False
'    RmDir temp

End Sub

'Возвращяет содержимое тэга alt
Function getProperty(Name As String, ByVal str As String) As String
Dim start As Integer
   
    start = InStr(str, Name & "=")
    If start > 0 Then
        str = Mid(str, start + 5, Len(str))
        str = Left(str, InStr(str, """") - 1)
    Else
        getProperty = ""
    End If
    
    getProperty = str
    
End Function

Sub parseHTML(m_inputDir As String, m_Filename As String, m_outputDir As String)
Dim adoc As Document, Doc As Document
    
    Set adoc = ActiveDocument
    Set Doc = Documents.Open(Filename:=(m_inputDir & "\" & m_Filename), Format:=wdOpenFormatText, ReadOnly:=True)

'Для Word 2013
    ActiveWindow.View.ReadingLayout = False

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<img"
        .Replacement.Text = "^p<img"
        .Forward = True
        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ">"
        .Replacement.Text = ">^p"
        .Forward = True
        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
    End With

    Selection.Find.Execute Replace:=wdReplaceAll

    Dim p As Paragraph
    Dim str As String
    Dim src As String
    Dim alt As String
    Dim fname As String
    Dim start As Integer
    
    For Each p In Doc.Paragraphs
        str = p.Range.Text
        If UCase(Mid(str, 1, 4)) = "<IMG" Then
            
            src = getProperty("src", str)
            src = Replace(src, "/", "\")
            
            'Выдергиваем имя файла
            start = InStrRev(src, "\")
            If start > 0 Then
                fname = right(src, Len(src) - start)
            End If
            
            FileCopy m_inputDir & "\" & src, m_outputDir & "\" & fname
            
            'Выдергиваем alt
            alt = getProperty("alt", str)
            
            'Заменяем картинки тегами
            Dim s As InlineShape
            For Each s In adoc.InlineShapes
                If s.AlternativeText = alt Then
                    s.Select
                    Selection.Range.Text = imgBase64(m_outputDir & "\" & fname)
                    Exit For
                End If
            Next
        End If
    Next
    Doc.Close SaveChanges:=False
    
  
End Sub

Private Function ReadFile(inFileName$) As String
    Dim outString$
    Debug.Print FileLen(inFileName)
    outString = String(FileLen(inFileName), " ")
    Dim fNum As Long
    fNum = FreeFile
    Open inFileName For Binary As #fNum
    Get #fNum, , outString
    Close #fNum
    ReadFile = outString
End Function

Private Function Base64Encode(InData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
   
    
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(InData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(InData, I, 1)) + _
      &H100 * MyASC(Mid(InData, I + 1, 1)) + MyASC(Mid(InData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(InData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Public Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

Private Function imgBase64(Filename$) As String
    Dim Extension
    Extension = right$(Filename, Len(Filename) - InStrRev(Filename, "."))
    'MsgBox extension
    imgBase64 = "<img src=""data:image/" & Extension & ";base64," & Base64Encode(ReadFile(Filename)) & """ />"
End Function


Private Sub ConvertToMoodleXML()
Dim I As Integer, curQ As Integer
Dim numv As Integer
Dim ncut As String
    Dim LastStyle As String
    Dim LastStyleQ As String 'для эссе
    Dim QDescription As CDescription
    Dim QEssay As CEssay
    Dim QTrueFalse As CTrueFalse
    Dim QNumerical As CNumerical
    Dim ANumerical As CNumericalAnswer
    Dim QMultichoice As CMultichoice
    Dim AMultichoice As CMultichoiceAnswer
    Dim QShortanswer As CShortanswer
    Dim AShortanswer  As CShortanswerAnswer
    Dim QMatching As CMatching
    Dim AMatching  As CMatchingSubquestion
    Dim Text As String
    Dim Feedback As String
    Dim TXT As TParseText
    Dim Answ As TNumericalAnswer
    Dim AnswM As TMultichoiceAnswer
    Dim AnswS As TShortanswerAnswer
    Dim AnswMS As TMatchingSubquestion
    Dim AnswEnd As Boolean
    Dim TFAnswer As Integer 'для truefalse: 1 - добавлен true, 2 - добавлен false, 3 - оба, 0 - нет ответов.
    
       
    

    
    Set Questions = New CQuestionCollection
    


    EscapeControlCharacters ' escape all special characters
    
    numv = 1
    ncut = ""
    For I = 0 To paraLast
        If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Then
            numv = 1
        End If
        If paraStyles(I).StyleName = STYLE_NUMERICALQ Then
            numv = 2
        End If
        If paraStyles(I).StyleName = STYLE_MISSINGWORDQ Then
            FindBlanks (paraStyles(I).Para.Range)
            paraStyles(I).Processed = True
        Else
            If (paraStyles(I).StyleName = STYLE_RIGHT_ANSWER Or paraStyles(I).StyleName = STYLE_WRONG_ANSWER) And _
                        StyleFound(STYLE_ANSWERWEIGHT, paraStyles(I).Para.Range) = True Then
                If numv = 1 Then
                    InsertTextBeforeRange TAG_WEIGHTED_ANSWER, paraStyles(I).Para.Range
                Else
                    InsertTextBeforeRange TAG_WEIGHTED_NUM_ANSWER, paraStyles(I).Para.Range
                End If
            paraStyles(I).Processed = True
            End If
        End If
    Next I
    
    RemoveFormatting

    I = 0
    While I <= paraLast
        If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Or _
                paraStyles(I).StyleName = STYLE_MATCHINGQ Or _
                paraStyles(I).StyleName = STYLE_NUMERICALQ Or _
                paraStyles(I).StyleName = STYLE_SHORTANSWERQ Or _
                paraStyles(I).StyleName = STYLE_TRUESTATEMENT Or _
                paraStyles(I).StyleName = STYLE_FALSESTATEMENT Or _
                paraStyles(I).StyleName = STYLE_MISSINGWORDQ Or _
                paraStyles(I).StyleName = STYLE_DESCRIPTIONQ Or _
                paraStyles(I).StyleName = STYLE_ESSAYQ Or _
                paraStyles(I).StyleName = STYLE_CATEGORY Or _
                paraStyles(I).StyleName = STYLE_BLANK_WORD Then
            curQ = I
            
            '=====================================
            If LastStyle = STYLE_NUMERICALQ Then
                If AnswEnd = True Then QNumerical.Answers.Add ANumerical
                Questions.Add QNumerical
            ElseIf LastStyle = STYLE_ESSAYQ Then
                Questions.Add QEssay
            ElseIf LastStyle = STYLE_MULTIPLECHOICEQ Then
                If AnswEnd = True Then QMultichoice.Answers.Add AMultichoice
                Questions.Add QMultichoice
            ElseIf LastStyle = STYLE_SHORTANSWERQ Then
                If AnswEnd = True Then QShortanswer.Answers.Add AShortanswer
                Questions.Add QShortanswer
            ElseIf LastStyle = STYLE_MATCHINGQ Then
                'If AnswEnd = True Then QMatching.Answers.Add AMatching
                Questions.Add QMatching
            ElseIf LastStyle = STYLE_TRUESTATEMENT Or LastStyle = STYLE_FALSESTATEMENT Then
                If LastStyle = STYLE_TRUESTATEMENT Then
                    QTrueFalse.Answer = True
                Else
                    QTrueFalse.Answer = False
                End If
                Questions.Add QTrueFalse
            End If
            
            AnswEnd = False
         
            LastStyle = ""
            LastStyleQ = ""
            TFAnswer = 0
            '=====================================
            
            strComment = "// " & START_QUESTION_COMMENT & ": " & paraStyles(I).StyleName & vbCr
            
            If paraStyles(I).StyleName = STYLE_CATEGORY Then    'Description havn't body but any other have
                strComment = "// " & CATEGORY_COMMENT & vbCr
            End If
            
            If I = 0 Then
                paraStyles(I).Para.Range.InsertBefore strComment
            Else
                If paraStyles(I - 1).StyleName <> STYLE_DESCRIPTIONQ And paraStyles(I - 1).StyleName <> STYLE_CATEGORY And paraStyles(I - 1).StyleName <> STYLE_MISSINGWORDQ Then
                    If ncut = "" Then
                        paraStyles(I).Para.Range.InsertBefore TAG_QUESTION_END & vbCr
                    Else
                        paraStyles(I).Para.Range.InsertBefore ncut & TAG_QUESTION_END & vbCr
                        ncut = ""
                    End If
                End If
                paraStyles(I).Para.Range.InsertBefore vbCr & strComment
            End If
            
            If paraStyles(I).StyleName = STYLE_CATEGORY Then
                Questions.Add SetCategory(paraStyles(I).Para.Range) 'XML Устанавливаем категорию
                InsertTextBeforeRange TAG_CATEGORY, paraStyles(I).Para.Range
            End If
            
            'feedback
            Feedback = ""
            '===XML вопросы описание===
                If paraStyles(I).StyleName = STYLE_DESCRIPTIONQ Then
                    LastStyle = STYLE_DESCRIPTIONQ
                    Text = paraStyles(I).Para.Range
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                    
                    Set QDescription = New CDescription
                    Set QDescription.QuestionText = GetCHTML(Text)
                    Set QDescription.Generalfeedback = GetCHTML(Feedback)
                    Text = QDescription.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QDescription.Name = TXT.Name
                    QDescription.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    Questions.Add QDescription
                End If
            '===Конец описание===
            
            '===XML пропущенное слово===
                If paraStyles(I).StyleName = STYLE_MISSINGWORDQ Then
                    LastStyle = STYLE_MISSINGWORDQ
                    Text = paraStyles(I).Para.Range
                
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                    
                    
                    Set AShortanswer = New CShortanswerAnswer
                    AShortanswer.Fraction = 100
                    
                    Text = Trim(Text)
                    Text = Replace(Text, "\\", "&&&slesh&&&")
                    Text = Replace(Text, "\", "")
                    Text = Replace(Text, "&&&slesh&&&", "\\")
                    
                    If InStr(1, Text, "{=", vbTextCompare) > 0 And InStr(3, Text, "}", vbTextCompare) > 0 Then
                        AShortanswer.Text = Mid(Text, InStr(1, Text, "{=", vbTextCompare) + 2, InStr(InStr(1, Text, "{=", vbTextCompare) + 2, Text, "}", vbTextCompare) - InStr(1, Text, "{=", vbTextCompare) - 2)
                        'MsgBox AShortanswer.Text
                        Text = Replace(Text, "{=" + AShortanswer.Text + "}", "_____", 1, 1, vbTextCompare)
                    Else
                        MsgBox "Ошибка, не найден пропуск в вопросе."
                    End If
                    
                
                    Set QShortanswer = New CShortanswer
                    Set QShortanswer.QuestionText = GetCHTML(Text)
                    Set QShortanswer.Generalfeedback = GetCHTML(Feedback)
                    Text = QShortanswer.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QShortanswer.Name = TXT.Name
                    QShortanswer.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QShortanswer.Defaultgrade = TXT.Defaultgrade

                    QShortanswer.Answers.Add AShortanswer
                    Questions.Add QShortanswer
                End If
            '===Конец пропущенное слово===

            
            
            
            
            If paraStyles(I).StyleName <> STYLE_DESCRIPTIONQ And paraStyles(I).StyleName <> STYLE_CATEGORY And paraStyles(I).StyleName <> STYLE_MISSINGWORDQ Then    'Description havn't body but any other have
            '===XML вопросы Эссе===
                If paraStyles(I).StyleName = STYLE_ESSAYQ Then
                    LastStyle = STYLE_ESSAYQ
                    Text = paraStyles(I).Para.Range
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                    
                    Set QEssay = New CEssay
                    Set QEssay.QuestionText = GetCHTML(Text)
                    Set QEssay.Generalfeedback = GetCHTML(Feedback)
                    Text = QEssay.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QEssay.Name = TXT.Name
                    QEssay.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QEssay.Defaultgrade = TXT.Defaultgrade
                    If TXT.file = True Then QEssay.Attachments = 1
                    'Questions.Add QEssay
                End If
            '===Конец Эссе===
            
            '===XML вопросы truefalse===
                If paraStyles(I).StyleName = STYLE_TRUESTATEMENT Or paraStyles(I).StyleName = STYLE_FALSESTATEMENT Then
                    LastStyle = paraStyles(I).StyleName
                    Text = paraStyles(I).Para.Range
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                    
                    Set QTrueFalse = New CTrueFalse
                    Set QTrueFalse.QuestionText = GetCHTML(Text)
                    Set QTrueFalse.Generalfeedback = GetCHTML(Feedback)
                    Text = QTrueFalse.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QTrueFalse.Name = TXT.Name
                    QTrueFalse.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QTrueFalse.Defaultgrade = TXT.Defaultgrade
                End If
            '===Конец truefalse===
            
            '===XML вопросы числовые===
                If paraStyles(I).StyleName = STYLE_NUMERICALQ Then
                    LastStyle = STYLE_NUMERICALQ
                    Text = paraStyles(I).Para.Range
                
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                
                    Set QNumerical = New CNumerical
                    Set QNumerical.QuestionText = GetCHTML(Text)
                    Set QNumerical.Generalfeedback = GetCHTML(Feedback)
                    Text = QNumerical.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QNumerical.Name = TXT.Name
                    QNumerical.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QNumerical.Defaultgrade = TXT.Defaultgrade
                End If
            '===Конец числовые вопросы===
            '===XML мультивыбор===
                If paraStyles(I).StyleName = STYLE_MULTIPLECHOICEQ Then
                    LastStyle = STYLE_MULTIPLECHOICEQ
                    Text = paraStyles(I).Para.Range
                
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                
                    Set QMultichoice = New CMultichoice
                    Set QMultichoice.QuestionText = GetCHTML(Text)
                    Set QMultichoice.Generalfeedback = GetCHTML(Feedback)
                    Text = QMultichoice.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QMultichoice.Name = TXT.Name
                    QMultichoice.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QMultichoice.Defaultgrade = TXT.Defaultgrade
                    QMultichoice.Singleanswer = True
                    QMultichoice.Shuffleanswers = TXT.Shuffleanswers
                End If
            '===Конец мультивыбор===
            '===XML короткий ответ===
                If paraStyles(I).StyleName = STYLE_SHORTANSWERQ Then
                    LastStyle = STYLE_SHORTANSWERQ
                    Text = paraStyles(I).Para.Range
                
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                
                    Set QShortanswer = New CShortanswer
                    Set QShortanswer.QuestionText = GetCHTML(Text)
                    Set QShortanswer.Generalfeedback = GetCHTML(Feedback)
                    Text = QShortanswer.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QShortanswer.Name = TXT.Name
                    QShortanswer.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QShortanswer.Defaultgrade = TXT.Defaultgrade
                End If
            '===Конец короткий ответ===
            
            
            '===XML на сопоставление===QMatching.Subquestions.Add AMatching
                If paraStyles(I).StyleName = STYLE_MATCHINGQ Then
                    LastStyle = STYLE_MATCHINGQ
                    Text = paraStyles(I).Para.Range
                
                    If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                        Feedback = paraStyles(I + 1).Para.Range
                        'MsgBox feedback
                    End If
                
                    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                
                    Set QMatching = New CMatching
                    Set QMatching.QuestionText = GetCHTML(Text)
                    Set QMatching.Generalfeedback = GetCHTML(Feedback)
                    Text = QMatching.QuestionText.Text
                    Text = Mid(Text, 4, Len(Text) - 7)
                    TXT = ParseText(Text)
                    QMatching.Name = TXT.Name
                    QMatching.QuestionText.Text = "<p>" + TXT.Text + "</p>"
                    QMatching.Defaultgrade = TXT.Defaultgrade
                    QMatching.Shuffleanswers = TXT.Shuffleanswers
                End If
            '===Конец сопоставление===
            
                InsertAfterBeforeCR TAG_QUESTION_START, paraStyles(I).Para.Range
            End If
            
            If paraStyles(I).StyleName = STYLE_NUMERICALQ Then
                InsertAfterBeforeCR TAG_NUMERICAL_QUESTION, paraStyles(I).Para.Range
            End If
            
            If paraStyles(curQ).StyleName = STYLE_TRUESTATEMENT Then
                InsertAfterBeforeCR TAG_TRUE_CHOICE, paraStyles(I).Para.Range
            End If
            
            If paraStyles(curQ).StyleName = STYLE_FALSESTATEMENT Then
                InsertAfterBeforeCR TAG_FALSE_CHOICE, paraStyles(I).Para.Range
            End If
'            If paraStyles(curQ).StyleName = STYLE_MISSINGWORDQ Then
'                FindBlanks (paraStyles(curQ).Para.Range)
'            End If

            If paraStyles(I + 1).StyleName = STYLE_FEEDBACK Then
                'InsertTextBeforeRange TAG_FEEDBACK2, paraStyles(i + 1).Para.Range
                ncut = TAG_FEEDBACK2 & paraStyles(I + 1).Para.Range.Text
                paraStyles(I + 1).Para.Range.Delete
                I = I + 1
            End If
            
        Else
            With paraStyles(I)
                'Вопросы с весами
                If .Processed = True Then
                    If LastStyle = STYLE_NUMERICALQ Then
                        If AnswEnd = True Then QNumerical.Answers.Add ANumerical
                        AnswEnd = True
                        Set ANumerical = New CNumericalAnswer
                        Answ = ParseNumericalAnswer(.Para.Range)
                        ANumerical.Answer = Answ.Answer
                        ANumerical.Tolerance = Answ.Tolerance
                        ANumerical.Fraction = Answ.Fraction
                    End If
                    If LastStyle = STYLE_MULTIPLECHOICEQ Then
                        If AnswEnd = True Then QMultichoice.Answers.Add AMultichoice
                        AnswEnd = True
                        Set AMultichoice = New CMultichoiceAnswer
                        AnswM = ParseMultichoiceAnswer(.Para.Range)
                        Set AMultichoice.Answer = AnswM.Answer
                        AMultichoice.Fraction = AnswM.Fraction
                        If QMultichoice.Singleanswer = True And AnswM.Singleanswer = False Then
                            QMultichoice.Singleanswer = False
                        End If
                    End If
                'Конец цикла вопросов с весами
                
                ' Wrong answer found
                ElseIf .StyleName = STYLE_WRONG_ANSWER And (Not .Processed) Then
                    If LastStyle = STYLE_NUMERICALQ Then
                        If AnswEnd = True Then QNumerical.Answers.Add ANumerical
                        AnswEnd = True
                        Set ANumerical = New CNumericalAnswer
                        ANumerical.Answer = "*"
                        ANumerical.Fraction = 0
                    End If
                    If LastStyle = STYLE_MULTIPLECHOICEQ Then
                        If AnswEnd = True Then QMultichoice.Answers.Add AMultichoice
                        AnswEnd = True
                        Set AMultichoice = New CMultichoiceAnswer
                        AnswM = ParseMultichoiceAnswer("~" + .Para.Range)
                        Set AMultichoice.Answer = AnswM.Answer
                        AMultichoice.Fraction = AnswM.Fraction
                        If QMultichoice.Singleanswer = True And AnswM.Singleanswer = False Then
                            QMultichoice.Singleanswer = False
                        End If
                    End If
                    
                    If LastStyle = STYLE_ESSAYQ Then 'эссе
                        'LastStyleQ = STYLE_WRONG_ANSWER
                        Text = .Para.Range
                            If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1) 'Убираем абзац
                            Text = Trim(Text)
                            Text = Replace(Text, "\\", "&&&slesh&&&")
                            Text = Replace(Text, "\", "")
                            Text = Replace(Text, "&&&slesh&&&", "\\")
                            QEssay.Responsetemplate = GetCHTML(Text)
                    End If
                    If LastStyle = STYLE_TRUESTATEMENT Or LastStyle = STYLE_FALSESTATEMENT Then
                        LastStyleQ = STYLE_WRONG_ANSWER
                    End If
                    
                            
                    InsertTextBeforeRange TAG_WRONG_ANSWER, .Para.Range

                    
                ' Right answer found
                ElseIf .StyleName = STYLE_RIGHT_ANSWER And (Not .Processed) Then
''                    ' Weighted answer found
''                    If StyleFound(STYLE_ANSWERWEIGHT, .Para.Range) = True Then
''                        InsertTextBeforeRange TAG_WEIGHTED_ANSWER, .Para.Range
'                    ' Answer of the numerical question found
'                    If paraStyles(curQ).StyleName = STYLE_NUMERICALQ Then
'                        InsertTextBeforeRange TAG_RIGHT_NUMERICAL_ANSWER, .Para.Range
'                    ' Answer of the multiple choice question found
'                    Else
'                        InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
'                    End If
                    If LastStyle = STYLE_NUMERICALQ Then
                        If AnswEnd = True Then QNumerical.Answers.Add ANumerical
                        AnswEnd = True
                        Set ANumerical = New CNumericalAnswer
                        Answ = ParseNumericalAnswer(.Para.Range)
                        ANumerical.Answer = Answ.Answer
                        ANumerical.Tolerance = Answ.Tolerance
                        ANumerical.Fraction = Answ.Fraction
                    End If
                    If LastStyle = STYLE_MULTIPLECHOICEQ Then
                        If AnswEnd = True Then QMultichoice.Answers.Add AMultichoice
                        AnswEnd = True
                        Set AMultichoice = New CMultichoiceAnswer
                        AnswM = ParseMultichoiceAnswer("=" + .Para.Range)
                        Set AMultichoice.Answer = AnswM.Answer
                        AMultichoice.Fraction = AnswM.Fraction
                        If QMultichoice.Singleanswer = True And AnswM.Singleanswer = False Then
                            QMultichoice.Singleanswer = False
                        End If
                    End If
                    If LastStyle = STYLE_SHORTANSWERQ Then
                        If AnswEnd = True Then QShortanswer.Answers.Add AShortanswer
                        AnswEnd = True
                        Set AShortanswer = New CShortanswerAnswer
                        AnswS = ParseShortanswerAnswer(.Para.Range)
                        AShortanswer.Text = AnswS.Text
                        AShortanswer.Fraction = AnswS.Fraction
                    End If
                    If LastStyle = STYLE_ESSAYQ Then 'эссе
                        'LastStyleQ = STYLE_RIGHT_ANSWER
                        Set QEssay.Graderinfo = GetCHTML(.Para.Range)
                    End If
                    If LastStyle = STYLE_TRUESTATEMENT Or LastStyle = STYLE_FALSESTATEMENT Then
                        LastStyleQ = STYLE_RIGHT_ANSWER
                    End If

                    InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
                    
                ' left pair of the matching question
                ElseIf .StyleName = STYLE_LEFT_PAIR Then
                
                    If LastStyle = STYLE_MATCHINGQ Then 'XML
                        'If AnswEnd = True Then QMatching.Answers.Add AMatching
                        AnswEnd = True
                        
                        Set AMatching = New CMatchingSubquestion
                        AnswMS = ParseMatchingSubquestion(.Para.Range, "")
                        AMatching.Subquestion = AnswMS.Subquestion
                    End If 'End XML
                
                    InsertTextBeforeRange TAG_RIGHT_ANSWER, .Para.Range
                    InsertAfterBeforeCR TAG_MATCHINGQ_ARROW, .Para.Range
                    
                ' right pair of the matching question
                ElseIf .StyleName = STYLE_RIGHT_PAIR Then
                    If LastStyle = STYLE_MATCHINGQ Then 'XML
                        AnswMS = ParseMatchingSubquestion("", .Para.Range)
                        AMatching.Answer = AnswMS.Answer
                        QMatching.Subquestions.Add AMatching
                        AnswEnd = False
                    End If
                
                
                    ' Do nothing
                    
                ' Question feedback
                ElseIf .StyleName = STYLE_FEEDBACK Then
                    'XML
                    Feedback = .Para.Range
                    If Len(Feedback) > 0 Then Feedback = Left(Feedback, Len(Feedback) - 1)
                    
                    If LastStyle = STYLE_NUMERICALQ Then 'Комментарий для ответом числового вопроса
                        ANumerical.Feedback = GetCHTML(Feedback)
                        QNumerical.Answers.Add ANumerical
                        AnswEnd = False
                    End If
                    If LastStyle = STYLE_MULTIPLECHOICEQ Then 'Комментарий для ответов мультивыбор
                        AMultichoice.Feedback = GetCHTML(Feedback)
                        QMultichoice.Answers.Add AMultichoice
                        AnswEnd = False
                    End If
                    If LastStyle = STYLE_SHORTANSWERQ Then 'Комментарий для ответов короткий выбор
                        AShortanswer.Feedback = GetCHTML(Feedback)
                        QShortanswer.Answers.Add AShortanswer
                        AnswEnd = False
                    End If
                    If LastStyle = STYLE_TRUESTATEMENT Or LastStyle = STYLE_FALSESTATEMENT Then 'truefalse
                        If LastStyleQ = STYLE_RIGHT_ANSWER Then
                            QTrueFalse.TrueFeedback = GetCHTML(Feedback)
                        ElseIf LastStyleQ = STYLE_WRONG_ANSWER Then
                            QTrueFalse.FalseFeedback = GetCHTML(Feedback)
                        End If
                    End If
                    
                    
                    InsertTextBeforeRange TAG_FEEDBACK, .Para.Range
                End If
            End With
        End If
        I = I + 1
    Wend
    
    '=============Дублировать сверху!!!!========================
    If LastStyle = STYLE_NUMERICALQ Then
        If AnswEnd = True Then QNumerical.Answers.Add ANumerical
        Questions.Add QNumerical
    ElseIf LastStyle = STYLE_ESSAYQ Then
        Questions.Add QEssay
    ElseIf LastStyle = STYLE_MULTIPLECHOICEQ Then
        If AnswEnd = True Then QMultichoice.Answers.Add AMultichoice
        Questions.Add QMultichoice
    ElseIf LastStyle = STYLE_SHORTANSWERQ Then
        If AnswEnd = True Then QShortanswer.Answers.Add AShortanswer
        Questions.Add QShortanswer
    ElseIf LastStyle = STYLE_MATCHINGQ Then
        Questions.Add QMatching
    ElseIf LastStyle = STYLE_TRUESTATEMENT Or LastStyle = STYLE_FALSESTATEMENT Then
        If LastStyleQ = STYLE_RIGHT_ANSWER Then
            QTrueFalse.TrueFeedback = GetCHTML(Feedback)
        ElseIf LastStyleQ = STYLE_WRONG_ANSWER Then
            QTrueFalse.FalseFeedback = GetCHTML(Feedback)
        End If
        Questions.Add QTrueFalse
    End If
    '=====================================
    
    If paraStyles(I - 1).StyleName <> STYLE_MISSINGWORDQ And paraStyles(I - 1).StyleName <> STYLE_DESCRIPTIONQ Then
        'InsertAfterBeforeCR vbCr & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                    If ncut = "" Then
                        InsertAfterBeforeCR vbCr & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                    Else
                        InsertAfterBeforeCR vbCr & ncut & TAG_QUESTION_END, paraStyles(paraLast).Para.Range
                        ncut = ""
                    End If
        
        
    End If
End Sub

Private Function SetCategory(Name) As CCategory
    Set SetCategory = New CCategory
    SetCategory.Name = Left(Name, Len(Name) - 1)
End Function


Private Function GetCHTML(Text) As CHTML

    Dim HTML As CHTML
    
    Set HTML = New CHTML

    Text = Trim(Text)
    Text = Replace(Text, "\\", "&&&slesh&&&")
    Text = Replace(Text, "\", "")
    Text = Replace(Text, "&&&slesh&&&", "\")

    'Поиск картинок <img src="data:image/gif;base64,  " />
    Dim nfile As Integer 'Количество файлов
    Dim Pos As Long 'Запоминаем позицию поиска
    Dim imgData As String
    Dim Extension As String
    nfile = 0
    Do While InStr(1, Text, "<img src=""data:image/", vbTextCompare) > 0
        nfile = nfile + 1
        Pos = InStr(Pos + 1, Text, "<img src=""data:image/", vbTextCompare)
        'MsgBox Pos
        Extension = Mid(Text, Pos + 21, InStr(Pos, Text, ";base64,", vbTextCompare) - Pos - 21)
        Pos = InStr(Pos + 1, Text, ";base64,", vbTextCompare)
        imgData = Mid(Text, Pos + 8, InStr(Pos, Text, """ />", vbTextCompare) - Pos - 8)
        HTML.AddFile imgData, "img" + str(nfile) + "." + Extension
        Text = Replace(Text, "<img src=""data:image/" + Extension + ";base64," + imgData + """ />", "<img src=""@@PLUGINFILE@@/img" + str(nfile) + "." + Extension + """ />")
    Loop
    If Len(Text) > 0 Then HTML.Text = "<p>" + Text + "</p>"
    Set GetCHTML = HTML
End Function


Private Function ParseText(Text) As TParseText

    Dim TXT As TParseText
    
    TXT.Defaultgrade = 1
    TXT.Shuffleanswers = True
    TXT.file = False
    
    'Text = Trim(Text)
    'Text = Replace(Text, "\\", "&&&slesh&&&")
    'Text = Replace(Text, "\", "")
    'Text = Replace(Text, "&&&slesh&&&", "\")
    
    If InStr(1, Text, "[no_shuffle]", vbTextCompare) > 0 Then
        TXT.Shuffleanswers = False
        Text = Replace(Text, "[no_shuffle]", "")
    ElseIf InStr(1, Text, "[shuffle]", vbTextCompare) > 0 Then
        TXT.Shuffleanswers = True
        Text = Replace(Text, "[shuffle]", "")
    ElseIf InStr(1, Text, "[file]", vbTextCompare) > 0 Then
        TXT.file = True
        Text = Replace(Text, "[file]", "")
    End If
    'Поиск веса вопроса
    If InStr(1, Text, "[", vbTextCompare) > 0 And InStr(InStr(1, Text, "[", vbTextCompare) + 1, Text, "]", vbTextCompare) > 0 Then
        If IsNumeric(Mid(Text, InStr(1, Text, "[", vbTextCompare) + 1, InStr(InStr(1, Text, "[", vbTextCompare) + 1, Text, "]", vbTextCompare) - InStr(1, Text, "[", vbTextCompare) - 1)) Then
            TXT.Defaultgrade = CDbl(Mid(Text, InStr(1, Text, "[", vbTextCompare) + 1, InStr(InStr(1, Text, "[", vbTextCompare) + 1, Text, "]", vbTextCompare) - InStr(1, Text, "[", vbTextCompare) - 1))
            Text = Replace(Text, "[" & Mid(Text, InStr(1, Text, "[", vbTextCompare) + 1, InStr(InStr(1, Text, "[", vbTextCompare) + 1, Text, "]", vbTextCompare) - InStr(1, Text, "[", vbTextCompare) - 1) & "]", "")
            'MsgBox str(TXT.Defaultgrade)
        End If
    End If
    Text = Trim(Text)
    
    '====Поиск названия вопроса====================
    TXT.Name = ""
    If InStr(1, Text, "::", vbTextCompare) = 1 Then
        If InStr(3, Text, "::", vbTextCompare) > 0 Then
            TXT.Name = Mid(Text, 3, InStr(3, Text, "::", vbTextCompare) - 3)
            Text = Trim(right(Text, Len(Text) - InStr(3, Text, "::", vbTextCompare) - 2))
            'MsgBox name + vbCr + text
        End If
    End If
    If InStr(1, Text, ";;", vbTextCompare) = 1 Then
        If InStr(3, Text, "::", vbTextCompare) > 0 Then
            TXT.Name = Mid(Text, 3, InStr(3, Text, "::", vbTextCompare) - 3)
            Text = Replace(Text, "::", "", 2, 1)
            Text = Trim(right(Text, Len(Text) - 1))
            'MsgBox name + vbCr + text
        End If
    End If
    If TXT.Name = "" Then
        If Len(Text) < 100 Then
            TXT.Name = Text
        Else
            TXT.Name = Left(Text, 100) & "..."
        End If
    End If

    TXT.Text = Text
    
    ParseText = TXT
End Function

Private Function ParseNumericalAnswer(Text) As TNumericalAnswer
    Dim NumericalAnswer As TNumericalAnswer
    NumericalAnswer.Tolerance = 0
    NumericalAnswer.Fraction = 100
    
    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1) 'Убираем абзац
    Text = Trim(Text)
    Text = Replace(Text, "\\", "&&&slesh&&&")
    Text = Replace(Text, "\", "")
    Text = Replace(Text, "&&&slesh&&&", "\")
    '====Поиск веса ответа====================
    If InStr(1, Text, "=%", vbTextCompare) = 1 And InStr(3, Text, "%", vbTextCompare) > 0 Then
        'MsgBox Mid(text, 3, InStr(3, text, "%", vbTextCompare) - 3)
        NumericalAnswer.Fraction = CDbl(Mid(Text, 3, InStr(3, Text, "%", vbTextCompare) - 3))
        If InStr(InStr(3, Text, "%", vbTextCompare) + 1, Text, ":", vbTextCompare) > 0 Then
            'MsgBox Mid(text, InStr(3, text, "%", vbTextCompare) + 1, InStr(InStr(3, text, "%", vbTextCompare) + 1, text, ":", vbTextCompare) - InStr(3, text, "%", vbTextCompare) - 1)
            'MsgBox Right(text, Len(text) - InStr(InStr(3, text, "%", vbTextCompare) + 1, text, ":", vbTextCompare))
            NumericalAnswer.Answer = CVar(Mid(Text, InStr(3, Text, "%", vbTextCompare) + 1, InStr(InStr(3, Text, "%", vbTextCompare) + 1, Text, ":", vbTextCompare) - InStr(3, Text, "%", vbTextCompare) - 1))
            NumericalAnswer.Tolerance = CDbl(right(Text, Len(Text) - InStr(InStr(3, Text, "%", vbTextCompare) + 1, Text, ":", vbTextCompare)))
        Else
            'MsgBox Right(text, Len(text) - InStr(3, text, "%", vbTextCompare))
            NumericalAnswer.Answer = CVar(right(Text, Len(Text) - InStr(3, Text, "%", vbTextCompare)))
        End If
    Else
        NumericalAnswer.Answer = CVar(Text)
    End If
    
    ParseNumericalAnswer = NumericalAnswer
End Function


Private Function ParseMultichoiceAnswer(Text) As TMultichoiceAnswer
    Dim MultichoiceAnswer As TMultichoiceAnswer
    MultichoiceAnswer.Fraction = 0
    MultichoiceAnswer.Singleanswer = True
    
    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1) 'Убираем абзац
    Text = Trim(Text)
    'Text = Replace(Text, "\\", "&&&slesh&&&")
    'Text = Replace(Text, "\", "")
    'Text = Replace(Text, "&&&slesh&&&", "\")
    '====Поиск веса ответа====================
    If InStr(1, Text, "~%", vbTextCompare) = 1 And InStr(3, Text, "%", vbTextCompare) > 0 Then
        'MsgBox Mid(text, 3, InStr(3, text, "%", vbTextCompare) - 3)
        MultichoiceAnswer.Fraction = CDbl(Mid(Text, 3, InStr(3, Text, "%", vbTextCompare) - 3))
        'MsgBox Right(text, Len(text) - InStr(3, text, "%", vbTextCompare))
        Set MultichoiceAnswer.Answer = GetCHTML(right(Text, Len(Text) - InStr(3, Text, "%", vbTextCompare)))
        MultichoiceAnswer.Singleanswer = False
    ElseIf InStr(1, Text, "=", vbTextCompare) = 1 Then
        Set MultichoiceAnswer.Answer = GetCHTML(right(Text, Len(Text) - 1))
        MultichoiceAnswer.Fraction = 100
        MultichoiceAnswer.Singleanswer = True
    ElseIf InStr(1, Text, "~", vbTextCompare) = 1 Then
        Set MultichoiceAnswer.Answer = GetCHTML(right(Text, Len(Text) - 1))
        MultichoiceAnswer.Fraction = 0
        MultichoiceAnswer.Singleanswer = True
    End If
    
    ParseMultichoiceAnswer = MultichoiceAnswer
End Function




Private Function ParseShortanswerAnswer(Text) As TShortanswerAnswer
    Dim ShortanswerAnswer As TShortanswerAnswer
    ShortanswerAnswer.Fraction = 100
    
    If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1) 'Убираем абзац
    Text = Trim(Text)
    Text = Replace(Text, "\\", "&&&slesh&&&")
    Text = Replace(Text, "\", "")
    Text = Replace(Text, "&&&slesh&&&", "\")
    '====Поиск веса ответа====================
    If InStr(1, Text, "%", vbTextCompare) = 1 And InStr(2, Text, "%", vbTextCompare) > 0 Then
        'MsgBox Mid(text, 3, InStr(3, text, "%", vbTextCompare) - 3)
        If IsNumeric(Mid(Text, 2, InStr(2, Text, "%", vbTextCompare) - 2)) Then
            ShortanswerAnswer.Fraction = CDbl(Mid(Text, 2, InStr(2, Text, "%", vbTextCompare) - 2))
            ShortanswerAnswer.Text = right(Text, Len(Text) - InStr(2, Text, "%", vbTextCompare))
        End If
    Else
        ShortanswerAnswer.Text = Text
    End If
    ParseShortanswerAnswer = ShortanswerAnswer
End Function



Private Function ParseMatchingSubquestion(Question, Answer) As TMatchingSubquestion
    Dim MatchingSubquestion As TMatchingSubquestion


    'Обработка вопроса
    If Len(Question) > 0 Then Question = Left(Question, Len(Question) - 1) 'Убираем абзац
    If Len(Question) > 0 Then
        Set MatchingSubquestion.Subquestion = GetCHTML(Question)
    Else
        Set MatchingSubquestion.Subquestion = New CHTML
    End If
    'Обработка ответа
    If Len(Answer) > 0 Then Answer = Left(Answer, Len(Answer) - 1) 'Убираем абзац
    Answer = Trim(Answer)
    Answer = Replace(Answer, "\\", "&&&slesh&&&")
    Answer = Replace(Answer, "\", "")
    Answer = Replace(Answer, "&&&slesh&&&", "\")
    
    MatchingSubquestion.Answer = Answer

    ParseMatchingSubquestion = MatchingSubquestion
End Function


