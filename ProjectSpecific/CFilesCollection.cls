VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFilesCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing list of base64 encoded files.
' Is used instead of simple Collection, because it is not efficent to extract keys in foreach loops
' It also calculates MD5 hash of files, it's needed for DDMatch Subquestion export
' Copyright 2014-2015 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private pFiles As Collection
Private pFilenames As Collection

Private Sub Class_Initialize()
    Set pFiles = New Collection
    Set pFilenames = New Collection
End Sub

Private Sub Class_Terminate()
    Set pFiles = Nothing
    Set pFilenames = Nothing
End Sub

Public Sub Add(Data As String, Filename As String)
    pFiles.Add Data
    pFilenames.Add Filename
End Sub

Public Function file(Index) As String
    file = pFiles.Item(Index)
End Function

Public Function Filename(Index) As String
    Filename = pFilenames.Item(Index)
End Function

Public Function Count() As Long
    Count = pFilenames.Count
End Function
' Returns MD5(MD5(FirstFileName) & MD5(FirstFile) & MD5(SecondFile) & MD5(SecondFileName))
Public Function GetMD5() As String
    Dim I As Long
    Dim MD5List As String
    Dim Filenames As Collection
    Dim Files As Collection
    
    GetSortedCollections Filenames, Files
    MD5List = ""
    For I = 1 To pFilenames.Count
        MD5List = MD5List & MD5.MD5_string(Filenames.Item("key" & I))
        MD5List = MD5List & MD5.MD5_string(Files.Item("key" & I))
    Next
    GetMD5 = MD5.MD5_string(MD5List)
    Set Filenames = Nothing
    Set Files = Nothing
End Function

Private Sub GetSortedCollections(Filenames As Collection, Files As Collection)
    Dim I As Long
    
    'When we delete element from collection, where keys are numbers it breaks collection order. So we use collection, where keys are strings
    Set Filenames = New Collection
    Set Files = New Collection
    For I = 1 To pFilenames.Count
        Filenames.Add pFilenames.Item(I), "key" & I
        Files.Add pFiles.Item(I), "key" & I
    Next
    QuickSort Filenames, Files, 1, Filenames.Count
End Sub

Private Sub QuickSort(Filenames As Collection, Files As Collection, Lo As Long, Hi As Long)
    Dim MiddleElement As String
    Dim TempFiledata As String
    Dim TempFilename As String
    Dim TempLow As Long
    Dim TempHi As Long
    
    TempLow = Lo
    TempHi = Hi
    MiddleElement = Filenames.Item("key" & ((Lo + Hi) \ 2))
    Do While TempLow <= TempHi
        Do While Filenames.Item("key" & TempLow) < MiddleElement And TempLow < Hi
            TempLow = TempLow + 1
        Loop
        Do While MiddleElement < Filenames.Item("key" & TempHi) And TempHi > Lo
            TempHi = TempHi - 1
        Loop
        If TempLow <= TempHi Then
            If TempLow < TempHi Then
                TempFilename = Filenames.Item("key" & TempLow)
                TempFiledata = Files.Item("key" & TempLow)
                ' We cannot replace collection element directly, so we delete element and then create element with same key.
                ' This does not works with numbers, because in this case collection are automatically reordered, so we use string keys.
                Filenames.Remove "key" & TempLow
                Filenames.Add Filenames.Item("key" & TempHi), "key" & TempLow
                Files.Remove "key" & TempLow
                Files.Add Files.Item("key" & TempHi), "key" & TempLow
                Filenames.Remove "key" & TempHi
                Filenames.Add TempFilename, "key" & TempHi
                Files.Remove "key" & TempHi
                Files.Add TempFiledata, "key" & TempHi
            End If
            TempLow = TempLow + 1
            TempHi = TempHi - 1
        End If
    Loop
    If Lo < TempHi Then QuickSort Filenames, Files, Lo, TempHi
    If TempLow < Hi Then QuickSort Filenames, Files, TempLow, Hi
End Sub

Public Sub Export(ByRef ParentElement As MSXML2.IXMLDOMElement)
    Dim DOM As MSXML2.DOMDocument60
    Dim XMLElement As MSXML2.IXMLDOMElement
    Dim XMLText As MSXML2.IXMLDOMText
    Dim XMLAttr As MSXML2.IXMLDOMAttribute
    Dim I As Long
    
    Set DOM = ParentElement.OwnerDocument
    For I = 1 To pFilenames.Count
        Set XMLElement = DOM.createElement("file")
        ParentElement.appendChild XMLElement
        Set XMLAttr = DOM.createAttribute("name")
        XMLAttr.NodeValue = pFilenames.Item(I)
        XMLElement.setAttributeNode XMLAttr
        Set XMLAttr = DOM.createAttribute("path")
        XMLAttr.NodeValue = "/"
        XMLElement.setAttributeNode XMLAttr
        Set XMLAttr = DOM.createAttribute("encoding")
        XMLAttr.NodeValue = "base64"
        XMLElement.setAttributeNode XMLAttr
        Set XMLText = DOM.createTextNode(pFiles.Item(I))
        XMLElement.appendChild XMLText
    Next I
End Sub

Public Sub Import(ByRef Element As MSXML2.IXMLDOMElement)
    Dim Nodes As MSXML2.IXMLDOMNodeList
    Dim Node As MSXML2.IXMLDOMNode
    Dim Filedata As String
    Dim Filename As String
    
    Set pFiles = New Collection
    Set pFilenames = New Collection
    Set Nodes = Element.SelectNodes("file")
    For Each Node In Nodes
        Filedata = Node.SelectSingleNode("text()").NodeValue
        Filename = Node.SelectSingleNode("@name").NodeValue
        Add Filedata, Filename
    Next
End Sub
