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

Public Function CleanText(Text As String)
    Dim RegExp As Object
    
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Global = True
    RegExp.MultiLine = True
    RegExp.Pattern = "<[^>]*>"
    
    CleanText = CleanString(RegExp.Replace(Text, ""))
End Function

