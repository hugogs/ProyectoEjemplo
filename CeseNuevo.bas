Attribute VB_Name = "CeseNuevo"
Sub NUEVO()
    'Borrar datos de PareoMarcajes
    Sheets("PareoMarcajes").Select
    Cells.Select
    Selection.ClearContents
    Cells.Select
    'Borra datos de HorasExtras
    Sheets("HorasExtras").Select
    Range("A1:Q120,R1:R120").Select
    Selection.ClearContents
    Range("I:Q,S:W").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    'Borra datos de DATOS
    Sheets("DATOS").Select
    Range("W3:W43").Select
    Selection.ClearContents
    Range("I:Q,V:V,Y:Y,AA:AA,AC:AC").Select
    Selection.EntireColumn.Hidden = True
    Range("A3").Select
    'Borra Datos de CESE
    Sheets("CESE").Select
    'Elimina comentarios anteriores
    Range("F9:O9").Select
    Selection.ClearComments
    Range("A9:R9").Select
    Selection.ClearContents
    Range("A9").Select
End Sub
Sub CESE()
    Call DNI_Texto
    Call Datos
    Call UltimoDia
    Call DatosFeriados
    Call Datos_Tard_SalTempranas
    Call DatosFaltas
    Call BorraComentariosConCero
    Sheets("CESE").Select
    Range("A9").Select
End Sub
Sub DNI_Texto()
    Sheets("CESE").Select
    Range("A9").Select
    'DNI a texto
    Selection.TextToColumns Destination:=Range("A9"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    Range("A9").Select
End Sub
Sub Datos()
    Sheets("CESE").Select
    Range("I9:O9").Select
    Selection.ClearContents
  
    Range("I9").Select
    ActiveCell.FormulaR1C1 = "=HorasExtras!R[-3]C[16]"
    Range("J9").Select
    ActiveCell.FormulaR1C1 = "=HorasExtras!R[-3]C[16]"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "=HorasExtras!R[-3]C[16]"
    Range("L9").Select
    ActiveCell.FormulaR1C1 = "=HorasExtras!R[-3]C[16]"
    
    Range("M9").Select
    ActiveCell.FormulaR1C1 = "=DATOS!R[34]C[11]"
    Range("N9").Select
    ActiveCell.FormulaR1C1 = "=DATOS!R[34]C[12]"
    Range("O9").Select
    ActiveCell.FormulaR1C1 = "=DATOS!R[34]C[13]+DATOS!R[34]C[15]"
    
    Range("I9:O9").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
End Sub
Sub UltimoDia()
'ULTIMO DIA DE TRABAJO
Sheets("CESE").Select
'Elimina comentarios anteriores
Cells(9, 6).Select
Selection.ClearComments
'Copia dato de ultimo dia de trabajo
Sheets("DATOS").Select
textoUltTrabajo = "Datos:" & Chr(10)
Dim i, j As Integer
Dim cond1, cond2, cond3 As Boolean
'----ULTIMO DIA DE TRABAJO----
'Verificacion si existe dato: ULTIMO
Range("W2").Select
cond1 = False
For i = 0 To 40
    If ActiveCell.Value = "ULTIMO" Then
        cond1 = True
        Exit For
    End If
    ActiveCell.Offset(1, 0).Select
Next i
'Si Existe datos, agrego lo siguiente
If cond1 = True Then
    Range("W2").Select
    textoUltTrabajo = textoUltTrabajo
    For i = 0 To 40
        If ActiveCell.Value = "ULTIMO" Then
            textoUltTrabajo = textoUltTrabajo & "Último día de marcación:" & Chr(10) & Cells(2 + i, 4).Value & Chr(10)
            Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
End If
'----DESCANSO DE LA SEMANA----
'Verificacion si existe dato: DESCANSO
Range("W2").Select
cond2 = False
For j = 0 To 40
    If ActiveCell.Value = "DESCANSO" Then
        cond2 = True
        Exit For
    End If
    ActiveCell.Offset(1, 0).Select
Next j
'Si Existe datos, agrego lo siguiente
If cond2 = True Then
    Range("W2").Select
    textoUltTrabajo = textoUltTrabajo
    For j = 0 To 40
        If ActiveCell.Value = "DESCANSO" Then
            textoUltTrabajo = textoUltTrabajo & "Día Libre Semanal:" & Chr(10) & Cells(2 + j, 4).Value
            Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next j
End If
'----ULTIMO DIA DE VACACIONES----
'Verificacion si existe dato: VACACIONES
Range("W2").Select
cond3 = False
For j = 0 To 40
    If ActiveCell.Value = "VACACIONES" Then
        cond3 = True
        Exit For
    End If
    ActiveCell.Offset(1, 0).Select
Next j
'Si Existe datos, agrego lo siguiente
If cond3 = True Then
    Range("W2").Select
    textoUltTrabajo = textoUltTrabajo
    For j = 0 To 40
        If ActiveCell.Value = "VACACIONES" Then
            textoUltTrabajo = textoUltTrabajo & "Último día de vacaciones:" & Chr(10) & Cells(2 + j, 4).Value
            Exit For
        End If
        ActiveCell.Offset(1, 0).Select
    Next j
End If

'Se agrega como comentario de celda
Sheets("CESE").Select
Cells(9, 6).Select
ActiveCell.AddComment.Text Text:=textoUltTrabajo
End Sub
Sub Datos_Tard_SalTempranas()
'DATOS DE TARDANZAS
Sheets("CESE").Select
'Elimina comentarios anteriores
Cells(9, 15).Select
Selection.ClearComments
'Busca los datos a copiar
Sheets("DATOS").Select
textoTard = "Corresponde a:"
'Recorre y valida celdas con datos a copiar
If Range("AB43").Value <> 0 Then
    Dim t As Integer
    textoTard = textoTard & Chr(10) & "*Tardanzas:" & Chr(10)
    For t = 0 To 39
        If (Cells(3 + t, 28).Value <> "") Then
            textoTard = textoTard & Cells(3 + t, 28).Value & Chr(10)
        End If
    Next t
End If
'Recorre y valida celdas con datos a copiar
If Range("AD43").Value <> 0 Then
    Dim s As Integer
    textoTard = textoTard & Chr(10) & "*Salidas Tempranas:" & Chr(10)
    For s = 0 To 39
        If (Cells(3 + s, 30).Value <> "") Then
            textoTard = textoTard & Cells(3 + s, 30).Value & Chr(10)
        End If
    Next s
End If
'Inserta los datos encontrados como comentario
Sheets("CESE").Select
Cells(9, 15).Select
ActiveCell.AddComment.Text Text:=textoTard
Range("A9").Select
End Sub
Sub DatosFaltas()
'DATOS DE INASISTENCIAS
Sheets("CESE").Select
'Elimina comentarios anteriores
Cells(9, 14).Select
Selection.ClearComments
'Busca los datos a copiar
Sheets("DATOS").Select
textoFalt = "Corresponde a:" & Chr(10)
'Recorre y valida celdascon datos a copiar
Dim f As Integer
For f = 0 To 39
    If (Cells(3 + f, 26).Value <> "") Then
        textoFalt = textoFalt & Cells(3 + f, 26).Value & Chr(10)
    End If
Next f
'Inserta los datos encontrados como comentario
Sheets("CESE").Select
Cells(9, 14).Select
ActiveCell.AddComment.Text Text:=textoFalt
Range("A9").Select
End Sub
Sub DatosFeriados()
'DATOS DE INASISTENCIAS
Sheets("CESE").Select
'Elimina comentarios anteriores
Cells(9, 12).Select
Selection.ClearComments
'Busca los datos a copiar
Sheets("HorasExtras").Select
textoFer = "Corresponde al:" & Chr(10)
'Recorre y valida celdascon datos a copiar
Dim fe As Integer
For fe = 0 To 99
    If (Cells(4 + fe, 19).Value <> "") Then
        textoFer = textoFer & Cells(4 + fe, 19).Value & Chr(10)
    End If
Next fe
'Inserta los datos encontrados como comentario
Sheets("CESE").Select
Cells(9, 12).Select
ActiveCell.AddComment.Text Text:=textoFer
Range("A9").Select
End Sub
Sub BorraComentariosConCero()
Sheets("CESE").Select
Dim z As Integer
For z = 0 To 6
    Cells(9, 9 + z).Select
    If (Cells(9, 9 + z).Value = "0") Then
    Selection.ClearComments
    End If
Next z
Range("A9").Select
End Sub
