Attribute VB_Name = "CeseNuevo"
Sub NUEVO()
    'Borrar datos de PareoMarcajes
    Sheets("PareoMarcajes").Select
    Cells.Select
    Selection.ClearContents
    Cells.Select
    'Borra datos de HorasExtras
    Sheets("HorasExtras").Select
    Range("A1:Q120,R4:R120").Select
    Selection.ClearContents
    Range("I:Q,S:W").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    'Borra datos de DATOS
    Sheets("DATOS").Select
    Range("W1").Select
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
    Call Datos
    Call UltimoDia
    Call DatosFeriados
    Call Datos_Tard_SalTempranas
    Call DatosFaltas
    Call BorraComentariosConCero
    Sheets("CESE").Select
    Range("A9").Select
    
End Sub
Sub Datos()
    Sheets("CESE").Select
    'Range("A9:R9").Select
    'Selection.ClearComments
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
    ActiveCell.FormulaR1C1 = "=DATOS!R[29]C[11]"
    Range("N9").Select
    ActiveCell.FormulaR1C1 = "=DATOS!R[29]C[12]"
    Range("O9").Select
    ActiveCell.FormulaR1C1 = "=DATOS!R[29]C[13]+DATOS!R[29]C[15]"
    
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
textoUltTrabajo = "Último día de marcación:" & Chr(10) & Cells(1, 23).Value
'Inserta el último dia de trabajo como comentario
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
textoTard = "Corresponde a:" & Chr(10) & "*Tardanzas:" & Chr(10)
'Recorre y valida celdas con datos a copiar
Dim t As Integer
For t = 0 To 34
    If (Cells(3 + t, 28).Value <> "") Then
        textoTard = textoTard & Cells(3 + t, 28).Value & Chr(10)
    End If
Next t
textoTard = textoTard & Chr(10) & "*Salidas Tempranas:" & Chr(10)
'Recorre y valida celdas con datos a copiar
Dim s As Integer
For s = 0 To 34
    If (Cells(3 + s, 30).Value <> "") Then
        textoTard = textoTard & Cells(3 + s, 30).Value & Chr(10)
    End If
Next s
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
For f = 0 To 34
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
