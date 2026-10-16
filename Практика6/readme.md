## ex1

```vba
Option Explicit

Sub ex1()
    Dim A(0 To 14) As Integer
    Dim i As Integer
    Dim estDva As Boolean
    Dim count As Integer
    Dim sum As Double, k As Integer
    Dim startArr As String, endArr As String

    Randomize
    startArr = ""
    estDva = False

    For i = 0 To 14
        A(i) = Int(Rnd * 26)
        startArr = startArr & A(i) & " "
        If A(i) = 2 Then estDva = True
    Next i

    If estDva Then
        count = 0
        For i = 0 To 14
            If A(i) Mod 2 = 0 Then
                A(i) = 0
                count = count + 1
            End If
        Next i
        endArr = ""
        For i = 0 To 14
            endArr = endArr & A(i) & " "
        Next i
        MsgBox "Начальный: " & startArr & vbCrLf & _
               "Есть 2. Чётные заменены нулями." & vbCrLf & _
               "Кол-во замен: " & count & vbCrLf & _
               "Конечный: " & endArr
    Else
        sum = 0: k = 0
        For i = 0 To 14
            If A(i) Mod 2 <> 0 Then
                sum = sum + A(i)
                k = k + 1
            End If
        Next i
        If k > 0 Then
            MsgBox "Начальный: " & startArr & vbCrLf & _
                   "Двойки нет. Среднее нечётных: " & sum / k
        Else
            MsgBox "Начальный: " & startArr & vbCrLf & _
                   "Двойки нет. Нечётных элементов нет."
        End If
    End If
End Sub
```

---

## ex2

```vba
Sub ex2()
    Dim M(0 To 14) As Double
    Dim i As Integer
    Dim sum As Double, proizv As Double
    Dim sredArifm As Double, sredGeom As Double
    Dim raznost As Double
    Dim sumKv As Double
    Dim startArr As String, endArr As String

    Randomize
    sum = 0: proizv = 1
    startArr = ""

    For i = 0 To 14
        M(i) = Int(Rnd * 20 + 1)
        startArr = startArr & M(i) & " "
        sum = sum + M(i)
        proizv = proizv * M(i)
    Next i

    sredArifm = sum / 15
    sredGeom = proizv ^ (1 / 15)
    raznost = Abs(sredArifm - sredGeom)

    If Int(raznost) Mod 2 = 0 Then
        For i = 0 To 14
            If i Mod 2 = 0 Then M(i) = 0
        Next i
        endArr = ""
        For i = 0 To 14
            endArr = endArr & M(i) & " "
        Next i
        MsgBox "Начальный: " & startArr & vbCrLf & _
               "Разность = " & raznost & " (кратна 2)" & vbCrLf & _
               "Конечный: " & endArr
    Else
        sumKv = 0
        For i = 0 To 14
            sumKv = sumKv + M(i) ^ 2
        Next i
        MsgBox "Начальный: " & startArr & vbCrLf & _
               "Разность = " & raznost & " (не кратна 2)" & vbCrLf & _
               "Сумма квадратов: " & sumKv
    End If
End Sub
```

---

## ex3

```vba
Sub ex3()
    Dim C(0 To 14) As Double
    Dim i As Integer
    Dim sum As Double
    Dim sred As Double
    Dim otkl As Double
    Dim maxOtk As Double

    Randomize
    sum = 0
    For i = 0 To 14
        C(i) = Int(Rnd * 50 + 1)
        sum = sum + C(i)
    Next i

    sred = sum / 15
    maxOtk = 0

    For i = 0 To 14
        otkl = Abs(C(i) - sred)
        If otkl > maxOtk Then maxOtk = otkl
    Next i

    MsgBox "Среднее арифметическое: " & sred & vbCrLf & _
           "Максимальное отклонение: " & maxOtk
End Sub
```

---

## ex4

```vba
Sub ex4()
    Dim A(0 To 11) As Double
    Dim i As Integer
    Dim sum As Double
    Dim sred As Double
    Dim startArr As String, endArr As String

    Randomize
    sum = 0
    startArr = ""

    For i = 0 To 11
        A(i) = Int(Rnd * 50 + 1)
        startArr = startArr & A(i) & " "
        sum = sum + A(i)
    Next i

    sred = sum / 12

    For i = 0 To 11
        If A(i) < sred Then
            A(i) = 0
        ElseIf A(i) > sred Then
            A(i) = 1
        End If
    Next i

    endArr = ""
    For i = 0 To 11
        endArr = endArr & A(i) & " "
    Next i

    MsgBox "Начальный: " & startArr & vbCrLf & _
           "Среднее: " & sred & vbCrLf & _
           "Конечный: " & endArr
End Sub
```

---

## ex5

```vba
Sub ex5()
    Dim X(0 To 7) As Integer
    Dim Y(0 To 7) As Integer
    Dim i As Integer
    Dim startX As String, startY As String
    Dim endX As String, endY As String

    Randomize
    startX = "": startY = ""

    For i = 0 To 7
        X(i) = Int(Rnd * 50 + 1)
        Y(i) = Int(Rnd * 50 + 1)
        startX = startX & X(i) & " "
        startY = startY & Y(i) & " "
    Next i

    For i = 0 To 7
        If X(i) Mod 5 = 0 Then
            Y(i) = Y(i) * 5
        Else
            Y(i) = 0
        End If
    Next i

    endX = "": endY = ""
    For i = 0 To 7
        endX = endX & X(i) & " "
        endY = endY & Y(i) & " "
    Next i

    MsgBox "Начальный X: " & startX & vbCrLf & _
           "Начальный Y: " & startY & vbCrLf & _
           "Конечный X: " & endX & vbCrLf & _
           "Конечный Y: " & endY
End Sub
```

---

## ex6

```vba
Sub ex6()
    Dim M(0 To 14) As Integer
    Dim i As Integer
    Dim count As Integer
    Dim sumKrat As Double
    Dim startArr As String, endArr As String

    Randomize
    startArr = ""
    count = 0

    For i = 0 To 14
        M(i) = Int(Rnd * 16)
        startArr = startArr & M(i) & " "
        If M(i) >= 3 And M(i) <= 11 Then count = count + 1
    Next i

    If count < 4 Then
        sumKrat = 0
        For i = 0 To 14
            If count <> 0 And M(i) Mod count = 0 Then
                sumKrat = sumKrat + M(i)
            End If
        Next i
        MsgBox "Начальный: " & startArr & vbCrLf & _
               "Элементов 3..11: " & count & " (< 4)" & vbCrLf & _
               "Сумма кратных: " & sumKrat
    Else
        For i = 0 To 14
            If count <> 0 And M(i) Mod count <= 2 Then
                M(i) = M(i) ^ 2
            End If
        Next i
        endArr = ""
        For i = 0 To 14
            endArr = endArr & M(i) & " "
        Next i
        MsgBox "Начальный: " & startArr & vbCrLf & _
               "Элементов 3..11: " & count & " (>= 4)" & vbCrLf & _
               "Конечный: " & endArr
    End If
End Sub
```

---

## ex7

```vba
Sub ex7()
    Dim X(0 To 14) As Double
    Dim i As Integer
    Dim proizv As Double
    Dim sredGeom As Double
    Dim otkl As Double
    Dim maxOtk As Double
    Dim startArr As String, endArr As String

    Randomize
    proizv = 1
    startArr = ""

    For i = 0 To 14
        X(i) = Int(Rnd * 20 + 1)
        startArr = startArr & X(i) & " "
        proizv = proizv * X(i)
    Next i

    sredGeom = proizv ^ (1 / 15)
    maxOtk = 0

    For i = 0 To 14
        otkl = Abs(X(i) - sredGeom)
        If otkl > maxOtk Then maxOtk = otkl
    Next i

    If maxOtk Mod 2 = 0 Then
        For i = 2 To 14 Step 3
            X(i) = maxOtk
        Next i
    Else
        For i = 4 To 14 Step 5
            X(i) = X(i) ^ 2
        Next i
    End If

    endArr = ""
    For i = 0 To 14
        endArr = endArr & X(i) & " "
    Next i

    MsgBox "Начальный: " & startArr & vbCrLf & _
           "Среднее геом.: " & sredGeom & vbCrLf & _
           "Макс. отклонение: " & maxOtk & vbCrLf & _
           "Конечный: " & endArr
End Sub
```

---

## ex8

```vba
Sub ex8()
    Dim X(0 To 9) As Double
    Dim Novy() As Double
    Dim i As Integer, j As Integer
    Dim R As Double
    Dim startArr As String, endArr As String

    Randomize
    R = InputBox("Введите число R:")

    startArr = ""
    For i = 0 To 9
        X(i) = Rnd * 50 - 25
        startArr = startArr & Round(X(i), 2) & " "
    Next i

    ReDim Novy(0 To 30)
    j = 0
    Novy(j) = X(0)
    j = j + 1

    For i = 1 To 9
        If (X(i - 1) > 0 And X(i) > 0) Or (X(i - 1) < 0 And X(i) < 0) Then
            Novy(j) = R
            j = j + 1
        End If
        Novy(j) = X(i)
        j = j + 1
    Next i

    endArr = ""
    For i = 0 To j - 1
        endArr = endArr & Round(Novy(i), 2) & " "
    Next i

    MsgBox "Начальный: " & startArr & vbCrLf & _
           "Число R: " & R & vbCrLf & _
           "Новый массив: " & endArr
End Sub
```

---

## ex9

```vba
Sub ex9()
    Dim K(0 To 7) As Integer
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim startArr As String, endArr As String

    Randomize
    startArr = ""
    For i = 0 To 7
        K(i) = Int(Rnd * 50 + 1)
        startArr = startArr & K(i) & " "
    Next i

    For i = 0 To 6
        For j = 0 To 6 - i
            If K(j) < K(j + 1) Then
                temp = K(j)
                K(j) = K(j + 1)
                K(j + 1) = temp
            End If
        Next j
    Next i

    endArr = ""
    For i = 0 To 7
        endArr = endArr & K(i) & " "
    Next i

    MsgBox "Начальный: " & startArr & vbCrLf & _
           "Массив по убыванию: " & endArr
End Sub
```

---

## ex10

```vba
Sub ex10()
    Dim B(0 To 14) As Integer
    Dim i As Integer
    Dim pervy As Integer, vtoroy As Integer
    Dim naydeno As Integer
    Dim countPar As Integer
    Dim temp As Integer
    Dim startArr As String, endArr As String

    Randomize
    startArr = ""
    For i = 0 To 14
        B(i) = Int(Rnd * 50 + 1)
        startArr = startArr & B(i) & " "
    Next i

    pervy = -1: vtoroy = -1
    naydeno = 0
    For i = 0 To 14
        If B(i) Mod 2 = 0 Then
            naydeno = naydeno + 1
            If naydeno = 1 Then pervy = i
            If naydeno = 2 Then
                vtoroy = i
                Exit For
            End If
        End If
    Next i

    If naydeno = 2 Then
        temp = B(pervy)
        B(pervy) = B(vtoroy)
        B(vtoroy) = temp
    End If

    countPar = 0
    For i = 0 To 13
        If B(i) Mod 2 = 0 And B(i + 1) Mod 2 = 0 Then
            countPar = countPar + 1
        End If
    Next i

    For i = 1 To 14 Step 2
        B(i) = countPar
    Next i

    endArr = ""
    For i = 0 To 14
        endArr = endArr & B(i) & " "
    Next i

    MsgBox "Начальный: " & startArr & vbCrLf & _
           "Кол-во пар чётных: " & countPar & vbCrLf & _
           "Конечный: " & endArr
End Sub
```
