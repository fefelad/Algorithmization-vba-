# VBA Examples

## ex1 — Сумма ряда 1 + 1/2² + 1/3³ + … + 1/nⁿ

```vba
Sub ex1()
    Dim n As Integer, i As Integer
    Dim sum As Double

    n = InputBox("Введите n:")
    sum = 0
    i = 1

    Do While i <= n
        sum = sum + 1 / (i ^ i)
        i = i + 1
    Loop

    MsgBox "Сумма ряда: " & sum
End Sub
```

## ex2 — Факториал числа N

```vba
Sub ex2()
    Dim n As Integer, i As Integer
    Dim fact As Double

    n = InputBox("Введите число N:")
    fact = 1
    i = 1

    Do While i <= n
        fact = fact * i
        i = i + 1
    Loop

    MsgBox "Факториал " & n & "! = " & fact
End Sub
```

## ex3 — Суммы чисел, кратных 3 и 4

```vba
Sub ex3()
    Dim n As Integer, i As Integer
    Dim chislo As Integer
    Dim sum3 As Double, sum4 As Double
    Dim result As Double

    Randomize
    n = InputBox("Введите количество чисел N:")

    sum3 = 0: sum4 = 0
    i = 1

    Do While i <= n
        chislo = Int(Rnd * 100) + 1 ' случайное число от 1 до 100

        If chislo Mod 3 = 0 Then sum3 = sum3 + chislo
        If chislo Mod 4 = 0 Then sum4 = sum4 + chislo

        i = i + 1
    Loop

    If sum3 > sum4 Then
        result = (sum3 + sum4) / 2
        MsgBox "Сумма кратных 3 > суммы кратных 4" & vbNewLine & _
               "Среднее арифметическое: " & result
    Else
        result = Sqr(sum3 * sum4)
        MsgBox "Сумма кратных 3 <= суммы кратных 4" & vbNewLine & _
               "Среднее геометрическое: " & result
    End If
End Sub
```

## ex4 — Максимум и минимум среди N чисел

```vba
Sub ex4()
    Dim n As Integer, i As Integer
    Dim chislo As Integer
    Dim max As Integer, min As Integer
    Dim stroka As String

    Randomize
    n = InputBox("Введите количество чисел N:")

    i = 1
    stroka = ""

    Do
        chislo = Int(Rnd * 100) + 1 ' случайное число от 1 до 100
        stroka = stroka & chislo & " "

        If i = 1 Then
            max = chislo
            min = chislo
        Else
            If chislo > max Then max = chislo
            If chislo < min Then min = chislo
        End If

        i = i + 1
    Loop Until i > n

    MsgBox "Числа: " & stroka & vbNewLine & _
           "Максимум: " & max & vbNewLine & _
           "Минимум: " & min
End Sub
```
