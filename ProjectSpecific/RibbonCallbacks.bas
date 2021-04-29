Attribute VB_Name = "RibbonCallbacks"
Option Explicit '����������� ������ ���������� ���� ���������� � �����

'import-act (�������: button, �������: onAction), 2010+
Private Sub RibbonImportACT(control As IRibbonControl)
    GIFT.ConvertFromACT
End Sub

'import-xml (�������: button, �������: onAction), 2010+
Private Sub RibbonImportXML(control As IRibbonControl)
    Moodle2Word.Moodle2Word
End Sub

'add-�ategory (�������: button, �������: onAction), 2010+
Private Sub RibbonAddCategory(control As IRibbonControl)
    GIFT.AddCategory
End Sub

'add-description (�������: button, �������: onAction), 2010+
Private Sub RibbonAddDescriptionQ(control As IRibbonControl)
    GIFT.AddDescriptionQ
End Sub

'add-true-stat (�������: button, �������: onAction), 2010+
Private Sub RibbonAddTrueStatement(control As IRibbonControl)
    GIFT.AddTrueStatement
End Sub

'add-false-stat (�������: button, �������: onAction), 2010+
Private Sub RibbonAddFalseStatement(control As IRibbonControl)
    GIFT.AddFalseStatement
End Sub

'add-matching (�������: button, �������: onAction), 2010+
Private Sub RibbonAddMatchingQ(control As IRibbonControl)
    GIFT.AddMatchingQ
End Sub

'add-numerical (�������: button, �������: onAction), 2010+
Private Sub RibbonAddNumericalQ(control As IRibbonControl)
    GIFT.AddNumericalQ
End Sub

'add-short (�������: button, �������: onAction), 2010+
Private Sub RibbonAddShortAnswerQ(control As IRibbonControl)
    GIFT.AddShortAnswerQ
End Sub

'add-multiply (�������: button, �������: onAction), 2010+
Private Sub RibbonAddMultipleChoiceQ(control As IRibbonControl)
    GIFT.AddMultipleChoiceQ
End Sub

'add-essay (�������: button, �������: onAction), 2010+
Private Sub RibbonAddEssayQ(control As IRibbonControl)
    GIFT.AddEssayQ
End Sub

'add-missing-q (�������: button, �������: onAction), 2010+
Private Sub RibbonAddMissingWordQ(control As IRibbonControl)
    GIFT.AddMissingWordQ
End Sub

'add-gap (�������: button, �������: onAction), 2010+
Private Sub RibbonMarkBlankWord(control As IRibbonControl)
    GIFT.MarkBlankWord
End Sub

'change-true-false (�������: button, �������: onAction), 2010+
Private Sub RibbonExchangeTrueFalse(control As IRibbonControl)
    GIFT.MarkTrueAnswer
End Sub

'set-weight (�������: button, �������: onAction), 2010+
Private Sub RibbonSetAswersWeight(control As IRibbonControl)
    GIFT.SetAnswerWeights
End Sub

'remove-weight (�������: button, �������: onAction), 2010+
Private Sub RibbonRemoveAnswersWeight(control As IRibbonControl)
    GIFT.RemoveAnswerWeightsFromTheSelection
End Sub

'add-comment (�������: button, �������: onAction), 2010+
Private Sub RibbonAddComment(control As IRibbonControl)
    GIFT.AddQuestionFeedback
End Sub

'check-export (�������: button, �������: onAction), 2010+
Private Sub RibbonCheckStructure(control As IRibbonControl)
    GIFT.ExamineExportToGIFT
End Sub

'export-gift (�������: button, �������: onAction), 2010+
Private Sub RibbonExportGIFT(control As IRibbonControl)
    GIFT.ExportToGIFT
End Sub

'export-xml (�������: button, �������: onAction), 2010+
Private Sub RibbonExportXML(control As IRibbonControl)
    GIFT.ExportToGIFT
End Sub

'about (�������: button, �������: onAction), 2010+
Private Sub RibbonShowAbout(control As IRibbonControl)
    ufAbout.Show
End Sub

