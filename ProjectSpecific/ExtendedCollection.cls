VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtendedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class collection with ability to take keys
' Copyright 2014 Vadim Dvorovenko (Vadimon@mail.ru)
' License http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later

Option Explicit

Private m_Collection As Collection

Private Sub Class_Initialize()
    Set m_Collection = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Collection = Nothing
End Sub

Public Function Add(Item, Optional ByVal Key As String) As Boolean
    m_Collection.Add Array(Key, Item), Key
End Function

Public Sub Clear()
    Set m_Collection = New Collection
End Sub

Public Function Count() As Long
    Count = m_Collection.Count
End Function
Public Function Item(Index)
Attribute Item.VB_UserMemId = 0
    If VarType(m_Collection(Index)(1)) <> vbObject Then
        Item = m_Collection(Index)
    Else
        Set Item = m_Collection(Index)(1)
    End If
End Function
Public Function Key(ByVal Index As Long) As String
    Key = m_Collection.Item(Index)(0)
End Function
Public Sub Remove(Index)
    m_Collection.Remove Index
End Sub

