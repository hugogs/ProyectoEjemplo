Attribute VB_Name = "Guardar"
Sub Guardar()
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
        'Case "500035-Maestro Chacarilla"
            'Workbooks.Open Filename:="D:\500035 CHACARILLA\INFO RRHH Chacarilla\02 Ceses Chacarilla\FORMATO DE CESE 2017 Chacarilla.xlsx"
            'Sheets("Ceses Chacarilla").Select
            
        Case "500002-HC San Miguel"
            Workbooks.Open Filename:="D:\500002 SAN MIGUEL\INFO RRHH San Miguel\02 Ceses San miguel\FORMATO DE CESE 2018 San Miguel.xlsx"
            Sheets("Ceses San Miguel").Select
            
        Case "500005-HC Angamos"
            Workbooks.Open Filename:="D:\500005 ANGAMOS\INFO RRHH Angamos\02 Ceses Angamos\FORMATO DE CESE 2018 Angamos.xlsx"
            Sheets("Ceses Angamos").Select
        
        Case "500010-HC Lima Centro"
            Workbooks.Open Filename:="D:\500010 LIMA CENTRO\INFO RRHH Lima Centro\02 Ceses Lima Centro\FORMATO DE CESE 2018 Lima Centro.xlsx"
            Sheets("Ceses Lima Centro").Select
            
        Case "500026-HC San Juan de Lurigancho"
            Workbooks.Open Filename:="D:\500026 SAN JUAN DE LURIGANCHO\INFO RRHH San Juan de Lurigancho\02 Ceses San Juan de Lurigancho\FORMATO DE CESE 2018 San Juan de Lurigancho.xlsx"
            Sheets("Ceses San Juan de Lurigancho").Select
            
        Case Else
            MsgBox "Información tienda de apoyo " + tienda, vbOKOnly, "Guardar CESE"
            Workbooks.Open Filename:="D:\ECA - Varios\FORMATO DE CESE Apoyo.xlsx"
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
        'Case "500035-Maestro Chacarilla"
            'Windows("FORMATO DE CESE 2017 Chacarilla.xlsx").Activate
            
        Case "500002-HC San Miguel"
            Windows("FORMATO DE CESE 2018 San Miguel.xlsx").Activate
            
        Case "500005-HC Angamos"
            Windows("FORMATO DE CESE 2018 Angamos.xlsx").Activate
            
        Case "500010-HC Lima Centro"
            Windows("FORMATO DE CESE 2018 Lima Centro.xlsx").Activate
            
        Case "500026-HC San Juan de Lurigancho"
            Windows("FORMATO DE CESE 2018 San Juan de Lurigancho.xlsx").Activate
            
        Case Else
            Windows("FORMATO DE CESE Apoyo.xlsx").Activate
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

