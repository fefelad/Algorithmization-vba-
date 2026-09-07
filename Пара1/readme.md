Option Explicit

Sub ex1()
Dim inputNumber As Double
Dim v As Double
Dim s As Double

    inputNumber = InputBox("Введите длину ребра куба: ")
    v = inputNumber ^ 3
    s = 6 * inputNumber ^ 2
    MsgBox "Объем куба = " & v
    MsgBox "Площадь поверхности = " & s

End Sub

Sub ex2()
Dim syrie As Double
Dim pryazha As Double
Dim othody As Double
Dim poteri As Double

    syrie = 12 * 1000

    pryazha = syrie * 0.93
    othody = syrie * 0.06
    poteri = syrie * 0.01

    MsgBox "Пряжа: " & pryazha
    MsgBox "Отходы: " & othody
    MsgBox "Потери: " & poteri

End Sub

Sub ex3()

End Sub

Sub ex4()
Dim x1 As Double
Dim y1 As Double

    Dim x2 As Double
    Dim y2 As Double

    Dim distance As Double

    x1 = InputBox("Введите X1:")
    y1 = InputBox("Введите Y1:")

    x2 = InputBox("Введите X2:")
    y2 = InputBox("Введите Y2:")

    distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

    MsgBox "Ðàññòîÿíèå = " & distance

End Sub

Sub ex5()
Dim A As Double
Dim B As Double
Dim C As Double
Dim summa As Double

    A = InputBox("Введите A:")
    B = InputBox("Введите B:")
    C = InputBox("Введите C:")

    summa = A + B - C

    MsgBox "Процент A: " & A / summa * 100 & "%"
    MsgBox "Процент B: " & B / summa * 100 & "%"
    MsgBox "Процент C: " & C / summa * 100 & "%"

End Sub

Sub ex6()
Dim productionPerMinute As Integer
Dim machines As Integer
Dim hours As Integer
Dim minutes As Integer
Dim total As Integer

    productionPerMinute = 7
    machines = 3
    hours = 6


    minutes = hours * 60
    total = productionPerMinute * machines * minutes

    MsgBox "Всего тарелок: " & total

End Sub
