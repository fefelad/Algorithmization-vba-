```vba
Option Explicit

Sub ex1()
    Dim i As Integer, j As Integer
    Dim stroka As String

    stroka = ""

    For i = 1 To 9
        For j = 1 To 9
            stroka = stroka & i & "x" & j & "=" & (i * j) & vbTab
        Next j
        stroka = stroka & vbCrLf
    Next i

    MsgBox stroka
End Sub
```

```vba
Sub ex2()
    Dim n As Integer, m As Integer
    Dim i As Integer, j As Integer
    Dim count As Integer
    Dim stroka As String

    n = InputBox("Строк:")
    m = InputBox("Столбцов:")

    Randomize
    count = 0
    stroka = ""

    For i = 1 To n
        For j = 1 To m
            Dim x As Integer
            x = Int(Rnd * 10)
            stroka = stroka & x & " "
            If x = 0 Then count = count + 1
        Next j
        stroka = stroka & vbCrLf
    Next i

    MsgBox stroka & vbCrLf & "Количество нулей: " & count
End Sub
```

```vba
Sub ex3()
    Dim n As Integer, m As Integer
    Dim i As Integer, j As Integer
    Dim x As Integer
    Dim max As Integer, min As Integer
    Dim stroka As String
    Dim first As Boolean

    n = InputBox("Строк:")
    m = InputBox("Столбцов:")

    Randomize
    stroka = ""
    first = True

    For i = 1 To n
        For j = 1 To m
            x = Int(Rnd * 21) - 10 ' случайное число от -10 до 10
            stroka = stroka & x & " "

            If first Then
                max = x
                min = x
                first = False
            Else
                If x > max Then max = x
                If x < min Then min = x
            End If
        Next j
        stroka = stroka & vbCrLf
    Next i

    MsgBox stroka & vbCrLf & "Максимум: " & max & vbCrLf & "Минимум: " & min
End Sub
```
