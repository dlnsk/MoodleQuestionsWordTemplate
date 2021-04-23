Attribute VB_Name = "RangeToHTML"
' Library for convertion Word's Range object into html-formatted test and collection of base64 encoded images
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://creativecommons.org/licenses/by-sa/4.0/ Creative Commons «Attribution-ShareAlike» 4.0

Option Explicit

Private ConverterDocument As Word.Document

Public Function RangeToHTML(ByRef Range As Word.Range)
    Dim HTML As CHTML
    Dim Text As String
    Dim Files As CFilesCollection
    
    Set HTML = New CHTML
    RangeToHTML2 Range, Text, Files
    HTML.Text = Text
    Set HTML.Files = Files
    Set RangeToHTML = HTML
End Function

Private Sub RangeToHTML2(ByRef Range As Word.Range, HTML As String, Images As CFilesCollection)
    Dim Doc As Word.Document
    Dim FilenameBase  As String
    Dim FilenameDoc As String
    Dim FilenameHTML As String
    Dim FilenameHTMLFiles As String
    Dim Separator As String
    
    Range.Select
    Range.Copy
    Selection.Collapse
    Set Doc = Range.Application.Documents.Add(Visible:=False)
    Doc.Range.Paste
    FilenameBase = TempFilenameBase
    FilenameHTML = TempFilename("html", FilenameBase)
    FilenameHTMLFiles = TempFilename("files", FilenameBase)
    Doc.WebOptions.AllowPNG = True
    Doc.WebOptions.PixelsPerInch = 96
    Doc.WebOptions.encoding = msoEncodingUTF8
    Separator = Mid$(CStr(Format(0, "fixed")), 2, 1)
    If CDbl(Replace(Application.Version, ".", Separator)) >= 14 Then
        Doc.SaveAs2 Filename:=FilenameHTML, Fileformat:=WdSaveFormat.wdFormatFilteredHTML, AddToRecentFiles:=False
    Else
        Doc.SaveAs Filename:=FilenameHTML, Fileformat:=WdSaveFormat.wdFormatFilteredHTML, AddToRecentFiles:=False
    End If
    Doc.Close SaveChanges:=False
    
    HTML = GetHTML(FilenameHTML, FilenameBase & ".files")
    Set Images = GetPictures(FilenameHTMLFiles)
    
    RemoveFile FilenameHTML
    RemoveDir FilenameHTMLFiles
End Sub

Private Function GetHTML(Filename As String, PluginfileSubstitution As String) As String
    Dim RegExp As Object
    Dim Matches As Object
    Dim SubMatches As Object
    Dim HTML As String

    HTML = UTF8Decode(ReadFileString(Filename))
    
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    
    RegExp.Pattern = "<body.*?>[\s\S]*?<div.*?>([\s\S]*)</div>[\s\S]*?</body>"
    Set Matches = RegExp.Execute(HTML)
    HTML = Matches(0).SubMatches(0)
    RegExp.Pattern = "<(.*)\sclass=MsoNormal>"
    HTML = RegExp.Replace(HTML, "<$1>")
    RegExp.Pattern = "<a\s+name=""OLE_LINK[0-9]+"">([\s\S]*?)</a>"
    HTML = RegExp.Replace(HTML, "$1")
    RegExp.Pattern = Replace(PluginfileSubstitution, ".", "\.")
    HTML = RegExp.Replace(HTML, "@@PLUGINFILE@@")
    GetHTML = Trim(HTML)
End Function

Private Function GetPictures(Dirname As String) As CFilesCollection
    Dim FSO As Object
    Dim Folder As Object
    Dim file As Object
    Dim FileContent As String
     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set GetPictures = New CFilesCollection
    If FSO.FolderExists(Dirname) Then
        Set Folder = FSO.GetFolder(Dirname)
        For Each file In Folder.Files
            GetPictures.Add Base64.Base64Encode(ReadFileByte(Dirname & "\" & file.Name)), file.Name
        Next
    End If
    Set Folder = Nothing
    Set file = Nothing
    Set FSO = Nothing
End Function

Function UTF8Decode(ByVal sStr As String)
    Dim l As Long
    Dim sUTF8 As String
    Dim iChar As Integer
    Dim iChar2 As Integer
    Dim iChar3 As Integer
    
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                l = l + 1
            Else
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                iChar3 = Asc(Mid(sStr, l + 2, 1))
                sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                l = l + 2
            End If
        Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8Decode = sUTF8
End Function

Private Function ReadFileString(Filename As String) As String
    Dim FileNumber As Integer

    FileNumber = FreeFile
    Open Filename For Input As #FileNumber
    If LOF(FileNumber) > 0 Then
        ReadFileString = Input$(LOF(FileNumber), FileNumber)
    End If
    Close #FileNumber
End Function

Private Function ReadFileByte(Filename As String) As Byte()
    Dim FileNumber As Integer

    FileNumber = FreeFile
    Open Filename For Binary Access Read As #FileNumber
    If LOF(FileNumber) > 0 Then
        ReadFileByte = InputB(LOF(FileNumber), FileNumber)
    End If
    Close #FileNumber
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
