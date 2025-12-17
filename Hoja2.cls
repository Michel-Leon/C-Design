'====================== Hoja Codigos =============
Public Codigo As String

Public Sub Buscar_codigo(CodigoBuscado As String)

    Dim hojaDestino As Worksheet
    Dim hojaDatos As Worksheet
    Dim nombreHoja As Variant
    Dim celdaCodigo As Range
    Dim filaCodigo As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    Dim Descripcion As String
    Dim hojasDatos As Variant
    
    ' Definir las hojas donde buscar
    hojasDatos = Array("BD") ' Agrega más hojas si es necesario: Array("BD", "Hoja2", "Hoja3")
    
    Set hojaDestino = ThisWorkbook.Sheets("Codigos")
    encontrado = False

    If CodigoBuscado = "" Then
        MsgBox "No se ha generado ningún código de envolvente.", vbExclamation
        Exit Sub
    End If

    ' Buscar en cada hoja
    For Each nombreHoja In hojasDatos
        Set hojaDatos = ThisWorkbook.Sheets(nombreHoja)
        
        ' Buscar en la columna D el codigoBuscado
        Set celdaCodigo = hojaDatos.Range("D:D").Find(What:=CodigoBuscado, LookIn:=xlValues, LookAt:=xlWhole)

        If Not celdaCodigo Is Nothing Then
            filaCodigo = celdaCodigo.Row
            Descripcion = hojaDatos.Cells(filaCodigo, "Q").Value
            encontrado = True
            Exit For
        End If
    Next nombreHoja

    If Not encontrado Then
        MsgBox "Código no encontrado en las hojas de datos.", vbExclamation
        Exit Sub
    End If

    ' Escribir EN EL AREA DE TRABAJO
    With hojaDestino
        .Range("CX145").Value = CodigoBuscado          ' Código
        .Range("CX149").Value = Descripcion            ' Descripción
        .Range("CX153").Value = "Falta definir ruta"   ' Ruta de plano
    End With

    MsgBox "Código encontrado", vbInformation
    
End Sub

Public Sub Leer_Codigo()
    Codigo = Me.Codigo_Dibujo.Text
    'MsgBox "Código leído: " & Codigo, vbInformation
    Call Buscar_codigo(Codigo)
End Sub

