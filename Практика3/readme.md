## ex2

```vba
Option Explicit

Sub ex2()
    Dim num As Double
    num = Val(InputBox("Введите целое число:"))

    Select Case num
        Case 0 To 9
            MsgBox "Однозначное число"
        Case 10 To 99
            MsgBox "Двузначное число"
        Case 100 To 999
            MsgBox "Трёхзначное число"
        Case 1000 To 9999
            MsgBox "Четырёхзначное число"
        Case 10000 To 99999
            MsgBox "Пятизначное число"
        Case 100000 To 999999
            MsgBox "Шестизначное число"
        Case 1000000 To 9999999
            MsgBox "Семизначное число"
    End Select
End Sub
```

## ex3

```vba
Sub ex3()

    Dim number As Integer
    Dim P As Double
    Dim S As Double
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double
    Dim PI As Double

    PI = 3.14

    number = InputBox("Выберите фигуру:" & vbCrLf & _
                      "1 - Треугольник" & vbCrLf & _
                      "2 - Прямоугольник" & vbCrLf & _
                      "3 - Квадрат" & vbCrLf & _
                      "4 - Круг" & vbCrLf & _
                      "5 - Трапеция" & vbCrLf & _
                      "6 - Параллелограмм")

    Select Case number
        Case 1
            A = InputBox("Введите сторону треугольника:")
            P = 3 * A
            S = (A ^ 2 * Sqr(3)) / 4

        Case 2
            A = InputBox("Введите сторону A:")
            B = InputBox("Введите сторону B:")
            P = 2 * (A + B)
            S = A * B

        Case 3
            A = InputBox("Введите сторону квадрата:")
            P = 4 * A
            S = A ^ 2

        Case 4
            A = InputBox("Введите радиус круга:")
            P = 2 * PI * A
            S = PI * A ^ 2

        Case 5
            A = InputBox("Введите сторону трапеции:")
            B = InputBox("Введите боковую сторону:")
            C = InputBox("Введите сторону верхнего основания:")
            D = InputBox("Введите высоту:")
            P = A + B + 2 * C
            S = (A + B) / 2 * D

        Case 6
            A = InputBox("Введите сторону A:")
            B = InputBox("Введите сторону B:")
            P = 4 * A
            S = A * B

    End Select

    MsgBox "Периметр: " & P & vbCrLf & _
           "Площадь: " & S

End Sub
```

## ex4

```vba
Sub ex4()
    Dim number As String
    Dim prefix As String
    Dim operator As String

    number = InputBox("Введите десятизначный номер телефона:")
    prefix = Left(number, 3)

    Select Case prefix
        Case "910", "911", "912", "913", "914", "915", "916", "917", "918", "919"
            operator = "МТС"

        Case "903", "905", "906", "909", "960", "961", "962", "963", "964", "965", "966", "967", "968", "969"
            operator = "Билайн"

        Case "920", "921", "922", "923", "924", "925", "926", "927", "928", "929", "930", "931", "932", "933", "934", "936", "937", "938", "939"
            operator = "МегаФон"

        Case "900", "901", "902", "904", "908", "950", "951", "952", "953", "955", "956", "958", "977", "978", "991", "992", "993", "994", "995", "996", "997", "999"
            operator = "Tele2"

        Case Else
            operator = "Оператор не определён"
    End Select

    MsgBox "Номер: " & number & vbCrLf & _
           "Оператор: " & operator
End Sub
```
