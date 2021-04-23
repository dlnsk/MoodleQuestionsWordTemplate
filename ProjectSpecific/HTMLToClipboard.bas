Attribute VB_Name = "HTMLToClipboard"
' Library for copying MoodleXML html elements to Word's clipboard ready for paste to document
' Copyright 2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons «Attribution-ShareAlike» 4.0

Option Explicit

Private ConverterDocument As Word.Document

Public Sub HTMLToClipboard(ByRef HTML As CHTML)
    Dim Doc As Word.Document
    Dim FilenameBase  As String
    Dim FilenameDoc As String
    Dim FilenameHTML As String
    Dim FilenameHTMLFiles As String
    Dim Separator As String
    Dim InlineShape As Word.InlineShape
    Dim Shape As Word.Shape

    FilenameBase = TempFilenameBase
    FilenameDoc = TempFilename("rtf", FilenameBase)
    FilenameHTML = TempFilename("html", FilenameBase)
    FilenameHTMLFiles = TempFilename("files", FilenameBase)
        
    WriteHTML HTML.Text, FilenameHTML, FilenameBase & ".files"
    CreateDir FilenameHTMLFiles
    WritePictures HTML.Files, FilenameHTMLFiles
    
    Set Doc = Application.Documents.Open(Filename:=FilenameHTML, Format:=WdOpenFormat.wdOpenFormatWebPages, Visible:=False)
    
    ' ÷òîáû ðèñóíêè ñîõðàíÿëèñü ñ ôàéëîì
    For Each InlineShape In Doc.InlineShapes
        InlineShape.LinkFormat.SavePictureWithDocument = True
    Next
    For Each Shape In Doc.Shapes
        Shape.LinkFormat.SavePictureWithDocument = True
    Next
    
    Doc.Range.Copy
    
    Doc.Close SaveChanges:=False
    RemoveFile FilenameHTML
    RemoveDir FilenameHTMLFiles
End Sub

Private Function WriteHTML(HTML As String, Filename As String, PluginfileSubstitution As String)
    Dim FullHTML As String
    Dim RegExp As Object
    
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    RegExp.Pattern = "@@PLUGINFILE@@"
    HTML = RegExp.Replace(HTML, PluginfileSubstitution)

    FullHTML = ""
    FullHTML = FullHTML & "<html>"
    FullHTML = FullHTML & "<head>"
    FullHTML = FullHTML & "<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">"
    FullHTML = FullHTML & "<title>HTMLTOWORD</title>"
    FullHTML = FullHTML & "</head>"
    FullHTML = FullHTML & "<body>"
    FullHTML = FullHTML & "<div>"
    FullHTML = FullHTML & HTML
    FullHTML = FullHTML & "</div>"
    FullHTML = FullHTML & "</body>"
    FullHTML = FullHTML & "</html>"
    
    FullHTML = UTF8.EncodeUTF8(FullHTML)
    WriteFileString FullHTML, Filename
End Function

Private Sub WritePictures(FileCollection As CFilesCollection, Dirname As String)
    Dim I As Long
    Dim Filename As String
    Dim file As String
    
    For I = 1 To FileCollection.Count
        Filename = FileCollection.Filename(I)
        file = FileCollection.file(I)
        WriteFileByte Base64.Base64Decode(file), Dirname & "\" & Filename
        WriteFileByte Base64.Base64Decode(file), Dirname & "\" & URLEncode(Filename, True)
        WriteFileByte Base64.Base64Decode(file), Dirname & "\" & URLEncode(Filename, False)
    Next
End Sub

Private Sub WriteFileString(Filedata As String, Filename As String)
    Dim FileNumber As Integer

    FileNumber = FreeFile
    Open Filename For Output As #FileNumber
    Print #FileNumber, Filedata
    Close #FileNumber
End Sub

Private Sub WriteFileByte(Filedata() As Byte, Filename As String)
    Dim FileNumber As Integer
    Dim WritePosistion As Long

    FileNumber = FreeFile
    Open Filename For Binary Access Write As #FileNumber
    WritePosistion = 1
    Put #FileNumber, WritePosistion, Filedata
    Close #FileNumber
End Sub

Private Function CreateDir(Dirname As String)
    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Dirname) Then
        FSO.CreateFolder Dirname
    End If
    Set FSO = Nothing
End Function

Private Function RemoveDir(Dirname As String)
    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Dirname) Then
        FSO.DeleteFolder Dirname
    End If
    Set FSO = Nothing
End Function

Private Function RemoveFile(Filename As String)
    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(Filename) Then
        FSO.DeleteFile Filename
    End If
    Set FSO = Nothing
End Function

Private Function TempFilenameBase() As String
    Dim Source As String
    Dim Result As String
    Dim I As Long

    Randomize
    Source = "abcdefghijklmnopqrstuvwxyz0123456789"
    Result = ""
    For I = 1 To 8
        Result = Result & Mid$(Source, Int(Rnd() * Len(Source) + 1), 1)
    Next
    TempFilenameBase = Result
End Function

Private Function TempFilename(Extention As String, Optional Base As String = "") As String
    Dim Source As String
    Dim Result As String
    Dim I As Long

    If Base = "" Then
        Base = TempFilenameBase()
    End If
    TempFilename = Environ("TMP") & "\" & Base & "." & Extention
End Function

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    StringVal = EncodeUTF8(StringVal)
    Dim Space As String
    If SpaceAsPlus Then
        Space = "+"
    Else
        Space = "%20"
    End If
    
    Dim StringLen As Long
    StringLen = Len(StringVal)

    Dim Result As String
    Result = ""
    
    Dim I As Integer
    Dim Char As String
    Dim CharCode As Integer
    Dim CharCode1 As Integer
    Dim CharCode2 As Integer
    For I = 1 To StringLen
      Char = Mid$(StringVal, I, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
            Result = Result & Char
        Case 32
            Result = Result & Space
        Case 0 To 15
            Result = Result & "%0" & Hex(CharCode)
        Case Else
          Result = Result & "%" & Hex(CharCode)
      End Select
    Next I
    URLEncode = Result
End Function

