Attribute VB_Name = "UTF8"
' http://habrahabr.ru/post/138173/

Option Explicit

Public Function EncodeUTF8(s)
    Dim I, c, utfc, b1, b2, b3
    
    For I = 1 To Len(s)
        c = ToLong(AscW(Mid(s, I, 1)))
 
        If c < 128 Then
            utfc = Chr(c)
        ElseIf c < 2048 Then
            b1 = c Mod &H40
            b2 = (c - b1) / &H40
            utfc = Chr(&HC0 + b2) & Chr(&H80 + b1)
        ElseIf c < 65536 And (c < 55296 Or c > 57343) Then
            b1 = c Mod &H40
            b2 = ((c - b1) / &H40) Mod &H40
            b3 = (c - b1 - (&H40 * b2)) / &H1000
            utfc = Chr(&HE0 + b3) & Chr(&H80 + b2) & Chr(&H80 + b1)
        Else
            ' Младший или старший суррогат UTF-16
            utfc = Chr(&HEF) & Chr(&HBF) & Chr(&HBD)
        End If

        EncodeUTF8 = EncodeUTF8 + utfc
    Next
End Function

Private Function ToLong(intVal)
    If intVal < 0 Then
        ToLong = CLng(intVal) + &H10000
    Else
        ToLong = CLng(intVal)
    End If
End Function

Public Function DecodeUTF8(s)
    Dim I, c, n, b1, b2, b3

    I = 1
    Do While I <= Len(s)
        c = Asc(Mid(s, I, 1))
        If (c And &HC0) = &HC0 Then
            n = 1
            Do While I + n <= Len(s)
                If (Asc(Mid(s, I + n, 1)) And &HC0) <> &H80 Then
                    Exit Do
                End If
                n = n + 1
            Loop
            If n = 2 And ((c And &HE0) = &HC0) Then
                b1 = Asc(Mid(s, I + 1, 1)) And &H3F
                b2 = c And &H1F
                c = b1 + b2 * &H40
            ElseIf n = 3 And ((c And &HF0) = &HE0) Then
                b1 = Asc(Mid(s, I + 2, 1)) And &H3F
                b2 = Asc(Mid(s, I + 1, 1)) And &H3F
                b3 = c And &HF
                c = b3 * &H1000 + b2 * &H40 + b1
            Else
                ' Символ больше U+FFFF или неправильная последовательность
                c = &HFFFD
            End If
            s = Left(s, I - 1) + ChrW(c) + Mid(s, I + n)
        ElseIf (c And &HC0) = &H80 Then
            ' Неожидаемый продолжающий байт
            s = Left(s, I - 1) + ChrW(&HFFFD) + Mid(s, I + 1)
        End If
        I = I + 1
    Loop
    DecodeUTF8 = s
End Function
