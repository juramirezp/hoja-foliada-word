## Versión manual del procedimiento
---

````vb
Sub Impr_Consecutivo()

Dim Message As String, Title As String, Default As String, NumCopies As Long
Dim Rng1 As Range

' Establece el aviso.
Message = "Ingrese el número de copias que quiere imprimir"
' Establece el título.
Title = "Imprimir"
' Establece el valor predeterminado.
Default = "1"

' Despliega el mensaje, título y valor predeterminado.
NumCopies = Val(InputBox(Message, Title, Default))

' Se ingresa el valor inicial del contador.
' Se debe modificar en cada envio a impresión
SerialNumber = 70000

Set Rng1 = ActiveDocument.Bookmarks("SerialNumber").Range
Counter = 0

While Counter < NumCopies
Rng1.Delete
Rng1.Text = SerialNumber

Set myfont = New Font
myfont.Bold = False
'Inserta aquí el tipo de fuente que deseas
myfont.Name = "Calibri"
'Inserta aquí el tamaño de fuente que deseas
myfont.Size = "10"
Rng1.Font = myfont

ActiveDocument.PrintOut
SerialNumber = SerialNumber + 1
Counter = Counter + 1
Wend

End Sub
````