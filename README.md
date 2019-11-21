# Como imprimir hojas foliadas en MS Word usando Macro
###### Solución obtenida de https://microsoft.public.es.office2000.narkive.com/6auUnqMk/incluir-un-consecutivo-en-un-documento
---

1. Cuando un programa, por ejemplo Word, no ofrece opciones que buscas,
se puede lograr desarrollar estas opciones con un poco de código. En
este caso, usas Visual Basic Application (VBA). Todas las aplicaciones
de Office te permiten hacerlo. Este código se debe escribir en una ventana aparte, que
abres siempre con las teclas ALT-F11

2. Abre tu documento. Ten en mente que hay tres pasos aquí: 
   - Crear un marcador en el documento 
   - Crear un archivo texto donde se guardará la secuencia 
   - Escribir tu código (copiar/pegar) en el editor VBA.

3. **MARCADOR**: En el documento de Word ubícate en el sitio que quieres que
vaya el consecutivo. Selecciona del menú Insertar > Marcador. Aquí debes
poner nombre a tu marcador, debido a que el código posterior va a
incluir el nombre de este marcador, deberás ponerle el nombre
SerialNumber; finalmente haces clic en Agregar. No verás nada especial,
pues es un marcador oculto.

4. **ARCHIVO**: Debido a que el nombre y su ubicación ya se determinan en el
próximo código, debes crear un vacío archivo consecutivo.txt en la raíz
de tu disco duro: c:\consecutivo.txt.

5. **CODIGO**: Con ALT+F11 abres el editor de VBA. Veamos esta ventana a
grandes rasgos: A la izquierda verás el panel de Proyectos, es muy
parecido a la lista de carpetas cuando abres Explorador de WIndows.
Observa que el cursor está en "This Document", que cuelga de la carpeta
"Microsoft Word Objetos", lo cual es correcto. Entonces, vas a crear una
carpeta especial para poner el código.

6. Del menú de esta ventana del editor de VBA, selecciona Insertar >
Módulo.

7. Observa que automáticamente te ha creado una carpeta "Módulos", y te
ha creado un objeto en esta carpeta: "Módulo1". Observa que a la derecha
se ha abierto una ventana, es lo que va a contener este nuevo objeto
"Módulo1". Es aquí donde debes pegar el código.

````vb
Sub Impr_Consecutivo()
Dim Message As String, Title As String, Default As String, NumCopies As
Long
Dim Rng1 As Range

' Establece el aviso.
Message = "Ingrese el número de copias que quiere imprimir"
' Establece el título.
Title = "Imprimir"
' Establece el valor predeterminado.
Default = "1"

' Despliega el mensaje, título y valor predeterminado.
NumCopies = Val(InputBox(Message, Title, Default))
SerialNumber = System.PrivateProfileString("C:\Consecutivo.Txt", _
"MacroSettings", "SerialNumber")

If SerialNumber = "" Then
SerialNumber = 1
End If

Set Rng1 = ActiveDocument.Bookmarks("SerialNumber").Range
Counter = 0

While Counter < NumCopies
Rng1.Delete
Rng1.Text = SerialNumber
ActiveDocument.PrintOut
SerialNumber = SerialNumber + 1
Counter = Counter + 1
Wend

'Guarda el próximo número en el archivo Consecutivo.txt listo para su
próximo uso.
System.PrivateProfileString("C:\Settings.txt", "MacroSettings", _
"SerialNumber") = SerialNumber

'Recrea el marcador listo para su próximo uso.
With ActiveDocument.Bookmarks
.Add Name:="SerialNumber", Range:=Rng1
End With

ActiveDocument.Save
End Sub 
````

8. Ve al menú Archivo > Cerrar y Volver a Word. Esto cerrará el editor
de VBA para regresar a tu documento.
9.  Guarda el documento, para que se guarde la macro.
10.  Para ejecutar la macro, puedes ir a menú Herramienta > Macro >
Macros y Ejecutar la macro a la que he denominado Impr_Consecutivo.
11. También puedes asignar un botón o una combinación de teclas a esa
macro para ejecutar más rápido. Sobre esto puedes encontrar más
información en la propia ayuda de Word.

**Nota:** Se adjunta tambien una version editada del codigo en donde la asignacion de los valores se realiza de manera manual.