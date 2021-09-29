# septiembre 20 2021

el instructor nos hablo de que en excel hay algo llamado funciones, para manejar la hoja de calculo o manipularla un poco, que puedo crear una funcion personal, ponerle el nombre que yo quiera
hay una funcion llamada AVERAGE que saca el promedo, esta funcion divide, resta y multiplica, se puede hacer lo mismo que con una formula.
se trabajaria con Cells que al momento de ejecutar en visual basic en mi hoja de calculo aparecera la informacion 


## ejemplo en excel

```
Sub prueba10
    datos.Cells(3,1) = form.Cells(6,4)
    datos.Cells(3,2) = form.Cells(8,4)
    MsgBox "registro almacenado"
    form.Cells(6,4) = Empty
    form.Cells(8,4) = ""
End Sub
```

```
Function promediosena(a,b,c)
    promedio = (a + b + c) / 3
End Function
```

## tarea

hacer una funcion que trabaje 
=misnotas(6,3,9,8,6) las notas que yo quiera
la funcion debe mostrar gano, perdio 
si el promedio es mayor a 7 
y un formulario que registre en la hoja dos

```
Sub ejemplo()
    Hoja2.Cells(3, 2) = Hoja1.Cells(6, 4)
    Hoja2.Cells(3, 3) = Hoja1.Cells(8, 4)
    Hoja2.Cells(3, 4) = Hoja1.Cells(10, 4)
    Hoja2.Cells(3, 5) = Hoja1.Cells(12, 4)
    MsgBox "registro almacenado"
    Hoja1.Cells(6, 4) = Empty
    Hoja1.Cells(8, 4) = Empty
    Hoja1.Cells(10, 4) = Empty
    Hoja1.Cells(12, 4) = Empty
End Sub
```

```
Function misnotas(f, g, h, i, j)
    promedio = (f + g + h + i + j) / 5
    If (promedio < 7) Then
        misnotas = " reprobo " & promedio & " con la nota "
    Else
        misnotas = " aprobo " & promedio & " con la nota "
    End If
End Function
```