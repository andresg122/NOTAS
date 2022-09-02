# NOTAS
registro de notas y clases
### PARADICMAS ESTRUCTURADOS 
#### algoritmo 
Tiene que tener un principio y un final.
 para introducir texto que se vea reflejado es escribe de la siguiente manera:
 escriba "texto"
 o tambien puedes escribir
 Esc "texto"
 el codigo seria:
~~~
Msgbox "texto"
~~~
para hacer una variable nom,num
el codigo seria:
~~~
nom "el valor de la variable"
~~~

### IMPUBOX
se ultiliza para deperle informacion a un usuario y quede guardada en una variable 
el codigo es:
~~~
a = inputbox("texto")
~~~

si la informacion que daran es en numero poner al princio **int**
quedaria de la siguiente manera:

~~~
int(inputbox("texto"))
~~~
### CONDICIONAL SI
~~~
If 1 < 9 then 
msgbox "texto"
else 
msgbox "texto"
~~~
#### actividades con la condicional **si**
<a href="https://ibb.co/8KGvcHK"><img src="https://i.ibb.co/8KGvcHK/a.jpg" alt="a" border="0"></a>
<a href="https://ibb.co/dJtzMDj"><img src="https://i.ibb.co/dJtzMDj/aa.jpg" alt="aa" border="0"></a>

diagrama de flujo 

<a href="https://ibb.co/d0tygPp"><img src="https://i.ibb.co/d0tygPp/222.jpg" alt="222" border="0"></a>

### FUNCIONES 
~~~
function a (r,t)
c =  r * t
a = c
end function
~~~

### SIGLOS 
El siglo **for** sirve para repetir una tarea las veces que sean indicadas
ejemlo: 
~~~
for a = 1 to 10
  msgbox "hola"
next a
 ~~~
 esto quiere decir que mostrara 10 veces el **hola** en pantalla.
 
 ejercicio
 ~~~
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
~~~

~~~
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
~~~
