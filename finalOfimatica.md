
## exportar tabla a ACCES base de datos

```
Sub ExportarDatosAClientesX()
    Dim conn As Object
    Dim hoja As Worksheet
    Dim i As Integer, sql As String

    ' Conexi�n a Access
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Base.accdb"

    ' Crear tabla ClientesX si no existe
    On Error Resume Next
    conn.Execute "CREATE TABLE ClientesX (ID AUTOINCREMENT PRIMARY KEY, Nombre TEXT(50), Correo TEXT(50), Telefono TEXT(20))"
    On Error GoTo 0

    ' Borrar registros existentes en ClientesX
    On Error Resume Next
    conn.Execute "DELETE FROM ClientesX"
    On Error GoTo 0

    ' Referencia a la hoja activa
    Set hoja = ThisWorkbook.Sheets("Datos") ' Aseg�rate que se llama as�
    i = 2 ' Empieza en la fila 2 (asumiendo fila 1 tiene encabezados)

    ' Insertar registros
    Do While hoja.Cells(i, 1).Value <> ""
        sql = "INSERT INTO ClientesX (Nombre, Correo, Telefono) VALUES ('" & _
              hoja.Cells(i, 1).Value & "', '" & hoja.Cells(i, 2).Value & "', '" & hoja.Cells(i, 3).Value & "')"
        conn.Execute sql
        i = i + 1
    Loop

    conn.Close
    MsgBox "Datos exportados a la tabla ClientesX (registros anteriores borrados)"
End Sub

```

## importar a exel desde la base de datos ACCES

```
Sub ImportarDesdeClientesX()
    Dim conn As Object, rs As Object
    Dim hoja As Worksheet, i As Integer

    ' Conexi�n a Access
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Base.accdb"

    ' Consulta a la tabla ClientesX
    rs.Open "SELECT Nombre, Correo, Telefono FROM ClientesX", conn

    ' Verificar o crear hoja "Datos"
    On Error Resume Next
    Set hoja = ThisWorkbook.Sheets("Datos")
    If hoja Is Nothing Then
        Set hoja = ThisWorkbook.Sheets.Add
        hoja.Name = "Datos"
    End If
    On Error GoTo 0

    hoja.Cells.ClearContents

    ' Encabezados
    hoja.Cells(1, 1).Value = "Nombre"
    hoja.Cells(1, 2).Value = "Correo"
    hoja.Cells(1, 3).Value = "Telefono"

    ' Importar datos
    i = 2
    Do While Not rs.EOF
        hoja.Cells(i, 1).Value = rs.Fields("Nombre").Value
        hoja.Cells(i, 2).Value = rs.Fields("Correo").Value
        hoja.Cells(i, 3).Value = rs.Fields("Telefono").Value
        i = i + 1
        rs.MoveNext
    Loop

    rs.Close: conn.Close
    MsgBox "Datos importados desde ClientesX"
End Sub


```

## Crear una base de datos Access desde Exel - visual Basic 

```
Sub CrearBaseAccessConDatos()
    Dim cat As Object
    Dim conn As Object
    Dim ruta As String
    
    ruta = "C:\Users\Matute\Documents\Base.accdb" ' Ajusta la ruta seg�n tu PC
    
    ' Crear archivo Access
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ruta
    
    ' Crear tabla Clientes
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ruta
    conn.Execute "CREATE TABLE Clientes (ID AUTOINCREMENT PRIMARY KEY, Nombre TEXT(50), Correo TEXT(50), Telefono TEXT(20))"
    
    ' Insertar 10 registros de ejemplo
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Ana Torres','ana.torres@email.com','999111222')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Luis Garc�a','luis.garcia@email.com','999222333')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Mar�a L�pez','maria.lopez@email.com','999333444')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Carlos P�rez','carlos.perez@email.com','999444555')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Sof�a Ram�rez','sofia.ramirez@email.com','999555666')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Jorge Castillo','jorge.castillo@email.com','999666777')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Elena Rojas','elena.rojas@email.com','999777888')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Pedro S�nchez','pedro.sanchez@email.com','999888999')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Luc�a Fern�ndez','lucia.fernandez@email.com','999999000')"
    conn.Execute "INSERT INTO Clientes (Nombre, Correo, Telefono) VALUES ('Miguel D�az','miguel.diaz@email.com','988123456')"
    
    conn.Close
    
    MsgBox "Base.accdb creada con tabla Clientes y 10 registros de ejemplo"
End Sub

```

## generar reporte en word desde Exel - Visual basic

```
Sub GenerarReporteWordConTabla()
    Dim wordApp As Object, doc As Object
    Dim hoja As Worksheet
    Dim i As Integer, totalFilas As Integer

    ' Inicializar Word
    Set wordApp = CreateObject("Word.Application")
    Set doc = wordApp.Documents.Add
    wordApp.Visible = True

    ' Referencia a la hoja
    Set hoja = ThisWorkbook.Sheets("Datos")

    ' Contar filas con datos
    i = 2
    Do While hoja.Cells(i, 1).Value <> ""
        i = i + 1
    Loop
    totalFilas = i - 2 ' N�mero de registros

    ' Insertar t�tulo
    doc.Content.Paragraphs(1).Range.Text = "Reporte de Clientes"
    doc.Content.Paragraphs(1).Range.Font.Bold = True
    doc.Content.Paragraphs(1).Range.Font.Size = 14
    doc.Content.Paragraphs.Add

    ' Insertar tabla
    Dim tabla As Object
    Set tabla = doc.Tables.Add(doc.Paragraphs(2).Range, totalFilas + 1, 3)

    ' Encabezados
    tabla.Cell(1, 1).Range.Text = "Nombre"
    tabla.Cell(1, 2).Range.Text = "Correo"
    tabla.Cell(1, 3).Range.Text = "Tel�fono"

    ' Formato de encabezado
    tabla.Rows(1).Range.Font.Bold = True
    tabla.Rows(1).Shading.BackgroundPatternColor = -603914241 ' gris claro

    ' Insertar datos
    For i = 1 To totalFilas
        tabla.Cell(i + 1, 1).Range.Text = hoja.Cells(i + 1, 1).Value
        tabla.Cell(i + 1, 2).Range.Text = hoja.Cells(i + 1, 2).Value
        tabla.Cell(i + 1, 3).Range.Text = hoja.Cells(i + 1, 3).Value
    Next i

    MsgBox "Reporte generado en Word con tabla distribuida"
End Sub


```

## conectar evento de boton click con una funcion destro de un modulo

```
Private Sub CommandButton1_Click()
    Call funcionNamePrintHola
End Sub

```

## mandar por correo Outlook

```
Sub Enviar_Correo_Outlook_Seguro()

    On Error Resume Next   ' evita que el error bloquee todo

    Dim appOutlook As Object
    Dim correo As Object

    Set appOutlook = CreateObject("Outlook.Application")

    If appOutlook Is Nothing Then
        MsgBox "Outlook no está disponible en este equipo.", vbExclamation
        Exit Sub
    End If

    Set correo = appOutlook.CreateItem(0)

    With correo
        .To = "ejemplo@gmail.com"
        .Subject = "Reporte Automático"
        .Body = "Correo generado desde Excel usando Visual Basic."
        .Display   ' SOLO MOSTRAR
    End With

    MsgBox "Correo creado correctamente"

End Sub
```
