### CICLOS

El CICLO **for** sirve para repetir una tarea las veces que sean indicadas
ejemplo: 

´´´´

    for a = 1 to 10
      msgbox "hola"
    
    next a

´´´´
Esto quiere decir que mostrara 10 veces el **hola** en pantalla.
 
## Ejercicio

´´´´

 Sub estudiante()
    For x = 1 To 4
    n = "SI"
    b = "NO"
    d = InputBox("usted va aportar ? (SI o NO)")
    If d = n Then
        a = Int(InputBox("cuanto va aportar? "))
        f = f + 1
        w = w + a
    Else
        h = h + 1
    End If
    If a > 10000 Then
        e = e + 1
        End If
    Next x
    p = w / f
    MsgBox "total de aportes: " & w
    MsgBox "el promedio es: " & p
    MsgBox "estudiante que aportaron: " & f
    MsgBox "estudiantes que no apoertaron: " & h
    MsgBox "estudiantes que dieron más de 10,000: " & e
    
End Sub

´´´´

´´´´


    Sub nombres()
    For x = 2 To 21
    p = xd.Cells(x, 2)
    ñ = Mid(p, 1, 2)

    e = xd.Cells(x, 3)
    k = Int(Len(e))
    b = Mid(e, k - 1, 2)

    c = xd.Cells(x, 1)
    q = Mid(c, 1, 2)
    
    xd.Cells(x, 4) = ñ & b & q
    
    Next x
    
    End Sub

´´´´
