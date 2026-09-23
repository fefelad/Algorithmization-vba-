## ex1

```vba
Sub ex1()

    Dim x As Double
    Dim y As Double

    x = InputBox("Введите x от 0 до 1")

    Select Case x

        Case Is < 0
            MsgBox "x вне диапазона"

        Case Is < 0.2
            y = 1 + Log(1 + x)

        Case 0.2 To 0.8
            y = (1 + Sqr(x)) / (1 + x)

        Case Is <= 1
            y = 2 * Exp(-2 * x)

        Case Else
            MsgBox "x вне диапазона"
            Exit Sub

    End Select

    MsgBox "y = " & y

End Sub
```

---

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
    End Select
End Sub
```

---

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

---

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

---

## ex12

```vba
Sub ex12()

    Dim x As Double
    Dim y As Double

    x = InputBox("введите x от -2 до 2")

    Select Case x

        Case Is < -2
            MsgBox "x вне диапозона"
            Exit Sub

        Case -2 To -1
            y = (1 + Abs(x)) / ((1 + x + x ^ 2) ^ (1 / 3))

        Case Is < 0
            y = 2 * Log(1 + x ^ 2) + (1 + Cos(x) ^ 4) / (2 + x)

        Case 0 To 2
            y = (1 + x) ^ (3 / 5)

        Case Else
            MsgBox "x вне диапозона"
            Exit Sub

    End Select

    MsgBox "y = " & y

End Sub
```

---

## ex13

```vba
Sub ex13()

    Dim x As Double
    Dim y As Double

    x = InputBox("Введите x от -2 до 2")

    Select Case x

        Case Is < -2
            MsgBox "x вне диапазона"
            Exit Sub

        Case -2 To 0
            y = (1 + x) / ((1 + x ^ 2) ^ (1 / 3))

        Case Is < 1
            y = -x + 2 * Exp(-2 * x)

        Case 1 To 2
            y = Abs(2 - x) ^ (1 / 3)

        Case Else
            MsgBox "x вне диапазона"
            Exit Sub

    End Select

    MsgBox "y = " & y

End Sub
```

---

## ex14

```vba
Sub ex14()

    Dim x As Double
    Dim y As Double

    x = InputBox("Введите x от -2 до 2")

    Select Case x

        Case Is < 0
            y = 3 * x + Sqr(1 + x ^ 2)

        Case 0 To 1
            y = 2 * Cos(x) * Exp(-2 * x)

        Case Else
            y = 2 * Sin(3 * x)

    End Select

    MsgBox "y = " & y

End Sub
```

---

## ex15

```vba
Sub ex15()

    Dim x As Double
    Dim y As Double

    x = InputBox("Введите x от -1.7 до 1.5")

    Select Case x

        Case Is < 0
            y = (1 + x + x ^ 2) / (1 + x ^ 3)

        Case Is < 1
            y = Sqr(1 + (2 * x) / (1 + x ^ 2))

        Case 1 To 1.5
            y = 2 * Abs(0.5 + Sin(x))

    End Select

    MsgBox "y = " & y

End Sub
```

---

## ex16

```vba
Sub ex16()
    Dim x As Double
    Dim y As Double

    x = InputBox("Введите значение x от -1.5 до 1.8:")

    Select Case x

        Case Is < 0
            y = 1 + (3 + x) / (1 + x ^ 2)

        Case Is < 1
            y = Sqr(1 + (1 - x) ^ 2)

        Case Else
            y = (1 + x) / (1 + Cos(x) ^ 2)

    End Select
    MsgBox "y = " & y
End Sub
```

---

## ex17

```vba
Sub ex17()
    Dim x As Double
    Dim y As Double

    x = InputBox("Введите значение x от -1.4 до 1.9:")
    Select Case x

        Case Is < 0
            y = (1 + 2 * x) / (1 + x ^ 2)

        Case Is < 1
            y = Sin(x) ^ 2 * Sqr(1 + x)

        Case 1 To 1.9
            y = Sin(x) ^ 2 * Exp(0.2 * x)

    End Select

    MsgBox "y = " & y

End Sub
```

---

## ex18

```vba
Sub ex18()
    Dim x As Double
    Dim y As Double

    x = InputBox("Введите значение x от -1.4 до 1.4:")

    Select Case x

        Case Is < 0
            y = Abs(x) / (1 + x) * Exp(-2 * x)

        Case Is < 1
            y = Sqr(1 + x)

        Case 1 To 1.4
            y = (1 + Sin(x)) / (1 + x) + 3 * x

    End Select

    MsgBox "y = " & y

End Sub
```
