Attribute VB_Name = "CeseGuardar"
Sub GUARDAR()
    'Libro de CESE Abierto
    ActiveWorkbook.Save
    Dim tienda As String
    tienda = ""
    Sheets("PareoMarcajes").Select
    'Seleccionar tienda
    tienda = Cells(12, 5).Value
    'msg = MsgBox("tienda" + tienda, vbOKOnly, "Prueba")
    'Abrir el libro donde guardar, segun tienda
    Dim tda As Boolean
    tda = False
    Select Case tienda
        Case "500035-Maestro Chacarilla"
            Workbooks.Open Filename:="D:\500035 CHACARILLA\INFO RRHH Chacarilla\02 Ceses Chacarilla\FORMATO DE CESE 2017 Chacarilla.xlsx"
            Sheets("Ceses Chacarilla").Select
            
        Case "500037-Maestro Pueblo Libre"
            Workbooks.Open Filename:="D:\500037 PUEBLO LIBRE\INFO RRHH Pueblo Libre\02 Ceses Pueblo Libre\FORMATO DE CESE 2017 Pueblo Libre.xlsx"
            Sheets("Ceses Pueblo Libre").Select
            
        Case "500039-Maestro Ate"
            Workbooks.Open Filename:="D:\500039 ATE\INFO RRHH Ate\02 Ceses Ate\FORMATO DE CESE 2017 Ate.xlsx"
            Sheets("Ceses Ate").Select
        
        Case "500047-Maestro Trujillo"
            Workbooks.Open Filename:="D:\500047 TRUJILLO\INFO RRHH Trujillo\02 Ceses Trujillo\FORMATO DE CESE 2017 Trujillo.xlsx"
            Sheets("Ceses Trujillo").Select
            
        Case "500058-Maestro San Luis"
            Workbooks.Open Filename:="D:\500058 SAN LUIS\INFO RRHH San Luis\02 Ceses San Luis\FORMATO DE CESE 2017 San Luis.xlsx"
            Sheets("Ceses San Luis").Select
            
        Case Else
            MsgBox "Información tienda de apoyo " + tienda, vbOKOnly, "Guardar CESE"
            Workbooks.Open Filename:="D:\ECA - Varios\FORMATO DE CESE 2017 Apoyo.xlsx"
            Sheets("Ceses Tiendas").Select
            tda = True
    End Select
    'Selecciona la primera celda de datos
    Range("B7").Select
    'Busca la ultima celda de datos
    Range("B7").End(xlDown).Select
    Dim NumFila, NumColumna As Integer
    NumFila = ActiveCell.Row
    NumColumna = ActiveCell.Column
    'Selecciona la celda siguiente a la ultima con datos
    Cells(NumFila + 1, NumColumna).Select
    Selection.RowHeight = 17
    Windows("Proyecto CESE 3.xlsm").Activate
    Sheets("CESE").Select
    'Rango a copiar
    Range("A9:R9").Select
    Selection.Copy
    
    Select Case tienda
        Case "500035-Maestro Chacarilla"
            Windows("FORMATO DE CESE 2017 Chacarilla.xlsx").Activate
            
        Case "500037-Maestro Pueblo Libre"
            Windows("FORMATO DE CESE 2017 Pueblo Libre.xlsx").Activate
            
        Case "500039-Maestro Ate"
            Windows("FORMATO DE CESE 2017 Ate.xlsx").Activate
            
        Case "500047-Maestro Trujillo"
            Windows("FORMATO DE CESE 2017 Trujillo.xlsx").Activate
            
        Case "500058-Maestro San Luis"
            Windows("FORMATO DE CESE 2017 San Luis.xlsx").Activate
            
        Case Else
            Windows("FORMATO DE CESE 2017 Apoyo.xlsx").Activate
    End Select
    
    'Pega los datos
    Cells(NumFila + 1, NumColumna).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    'Agrego nombre de tienda de apoyo
    If (tda = True) Then
        Cells(NumFila + 1, NumColumna + 18).Select
        ActiveCell.FormulaR1C1 = tienda
    End If
    'Ultima celda con datos
    Cells(NumFila + 1, NumColumna).Select
    'Guarda y cierra el libro
    ActiveWorkbook.Save
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
    '-------
    ActiveWindow.Close
    'Regresa al libro CESE
    Windows("Proyecto CESE 3.xlsm").Activate
    Sheets("CESE").Select
    Range("A9").Select
    'Genera una copia del CESE y lo guarda
    Sheets("CESE").Copy
    Dim nombre As String
    nombre = Cells(15, 3).Value
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & "" & nombre & ".xlsx"
    Sheets("CESE").Select
    Sheets("CESE").Name = nombre
    'Borra datos no necesarios
    Range("B12:J15").Select
    Selection.ClearContents
    'Borra las imagenes
    Range("A9").Select
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Picture 3282")).Select
    Selection.Delete
    'Guarda y cierra
    ActiveWorkbook.Save
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
    '-------
    ActiveWorkbook.Close
    'Regresa al Libro CESE 3
    Windows("Proyecto CESE 3.xlsm").Activate
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
    '-------
    Sheets("CESE").Select
End Sub

