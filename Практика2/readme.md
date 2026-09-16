```vb
Option Explicit

Sub ex1()

    Dim A As Double
    Dim B As Double
    Dim A1 As Double
    Dim B1 As Double

    A = InputBox("Введите A:")
    B = InputBox("Введите B:")

    If A > B Then
        A1 = B
        B1 = A
    Else
        A1 = A
        B1 = B
    End If

    MsgBox "Прежнее значение A: " & A & vbCrLf & _
           "Полученное значение A: " & A1 & vbCrLf & _
           "Прежнее значение B: " & B & vbCrLf & _
           "Полученное значение B: " & B1

End Sub
```

## Задание 2

```vb
Sub ex2()

    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double
    Dim arithmeticMean As Double
    Dim geometricMean As Double

    A = InputBox("Введите A:")
    B = InputBox("Введите B:")
    C = InputBox("Введите C:")
    D = InputBox("Введите D:")

    arithmeticMean = (A + B + C + D) / 4
    geometricMean = (A * B * C * D) ^ (1 / 4)

    If A * B > C * D Then
        MsgBox "Среднее арифметическое: " & arithmeticMean
    Else
        MsgBox "Среднее геометрическое: " & geometricMean
    End If

End Sub
```

## Задание 3

```vb
Sub ex3()

    Dim A As Double
    Dim B As Double
    Dim C As Double

    A = InputBox("Введите сторону A:")
    B = InputBox("Введите сторону B:")
    C = InputBox("Введите сторону C:")

    If (A + B > C) And (A + C > B) And (B + C > A) Then
        MsgBox "С данными сторонами можно построить треугольник"
    Else
        MsgBox "С данными сторонами нельзя построить треугольник"
    End If

End Sub
```

## Задание 4

```vb
Sub ex4()

    Dim choice As Integer
    Dim number(1 To 5) As Double
    Dim i As Integer
    Dim max3 As Double
    Dim max4 As Double
    Dim max5 As Double

    choice = MsgBox("Использовать случайные числа?" & vbCrLf & _
                    "Да - случайные" & vbCrLf & _
                    "Нет - ввести вручную", vbYesNo)

    Randomize

    For i = 1 To 5

        If choice = vbYes Then
            number(i) = Int(Rnd * 100)
        Else
            number(i) = InputBox("Введите число " & i & ":")
        End If

    Next i

    max3 = number(1)

    If number(2) > max3 Then
        max3 = number(2)
    End If

    If number(3) > max3 Then
        max3 = number(3)
    End If

    max4 = max3

    If number(4) > max4 Then
        max4 = number(4)
    End If

    max5 = max4

    If number(5) > max5 Then
        max5 = number(5)
    End If

    MsgBox "Максимум из 3 чисел: " & max3 & vbCrLf & _
           "Максимум из 4 чисел: " & max4 & vbCrLf & _
           "Максимум из 5 чисел: " & max5

End Sub
```

## Задание 5

```vb
Sub ex5()

    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim diffAB As Double
    Dim diffBC As Double
    Dim diffCA As Double
    Dim minDiff As Double

    A = InputBox("Введите A:")
    B = InputBox("Введите B:")
    C = InputBox("Введите C:")

    diffAB = Abs(A - B)
    diffBC = Abs(B - C)
    diffCA = Abs(C - A)

    minDiff = diffAB

    If diffBC < minDiff Then
        minDiff = diffBC
    End If

    If diffCA < minDiff Then
        minDiff = diffCA
    End If

    MsgBox "Наименьшая разность: " & minDiff

End Sub
```

## Задание 6

```vb
Sub ex6()

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
                      "6 - Ромб")

    If number = 1 Then

        A = InputBox("Введите сторону треугольника:")
        P = 3 * A
        S = (A ^ 2 * Sqr(3)) / 4

    ElseIf number = 2 Then

        A = InputBox("Введите сторону A:")
        B = InputBox("Введите сторону B:")
        P = 2 * (A + B)
        S = A * B

    ElseIf number = 3 Then

        A = InputBox("Введите сторону квадрата:")
        P = 4 * A
        S = A ^ 2

    ElseIf number = 4 Then

        A = InputBox("Введите радиус круга:")
        P = 2 * PI * A
        S = PI * A ^ 2

    ElseIf number = 5 Then

        A = InputBox("Введите верхнее основание:")
        B = InputBox("Введите нижнее основание:")
        C = InputBox("Введите боковую сторону:")
        D = InputBox("Введите высоту:")

        P = A + B + 2 * C
        S = (A + B) / 2 * D

    ElseIf number = 6 Then

        A = InputBox("Введите сторону ромба:")
        B = InputBox("Введите высоту ромба:")

        P = 4 * A
        S = A * B

    Else

        MsgBox "Неверный номер фигуры!"
        Exit Sub

    End If

    MsgBox "Периметр: " & P & vbCrLf & _
           "Площадь: " & S

End Sub
```

## Задание 7

```vb
Sub ex7()

    Dim monthNumber As Integer
    Dim season As String
    Dim days As Integer

    monthNumber = InputBox("Введите номер месяца от 1 до 12:")

    If monthNumber < 1 Or monthNumber > 12 Then
        MsgBox "Номер месяца должен быть от 1 до 12"
        Exit Sub
    End If

    If monthNumber = 12 Or monthNumber = 1 Or monthNumber = 2 Then
        season = "Зима"
    ElseIf monthNumber >= 3 And monthNumber <= 5 Then
        season = "Весна"
    ElseIf monthNumber >= 6 And monthNumber <= 8 Then
        season = "Лето"
    Else
        season = "Осень"
    End If

    If monthNumber = 2 Then
        days = 28
    ElseIf monthNumber = 4 Or monthNumber = 6 Or monthNumber = 9 Or monthNumber = 11 Then
        days = 30
    Else
        days = 31
    End If

    MsgBox "Время года: " & season & vbCrLf & _
           "Количество дней: " & days

End Sub
```

## Задание 8

```vb
Sub ex8()

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim q1 As Integer
    Dim q2 As Integer

    x1 = InputBox("Точка 1. Введите X:")
    y1 = InputBox("Точка 1. Введите Y:")

    x2 = InputBox("Точка 2. Введите X:")
    y2 = InputBox("Точка 2. Введите Y:")

    If x1 > 0 And y1 > 0 Then
        q1 = 1
    ElseIf x1 < 0 And y1 > 0 Then
        q1 = 2
    ElseIf x1 < 0 And y1 < 0 Then
        q1 = 3
    ElseIf x1 > 0 And y1 < 0 Then
        q1 = 4
    End If

    If x2 > 0 And y2 > 0 Then
        q2 = 1
    ElseIf x2 < 0 And y2 > 0 Then
        q2 = 2
    ElseIf x2 < 0 And y2 < 0 Then
        q2 = 3
    ElseIf x2 > 0 And y2 < 0 Then
        q2 = 4
    End If

    MsgBox "Точка 1 находится в " & q1 & "-й четверти" & vbCrLf & _
           "Точка 2 находится в " & q2 & "-й четверти"

End Sub
```
