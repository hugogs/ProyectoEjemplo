Attribute VB_Name = "Dist01Y02"
Sub Dist_OC_1()
    'Elaborado por Hugo Garcia S.
    'V1.3 21.12.17
    Call Dist_OC_A
    Call Dist_OC_B
    'Confirmacion de pregunta
    Dim respSod, respMa As Byte
    Do
        respSod = MsgBox("¿La EMPRESA a trabajar es SODIMAC?", vbYesNo + vbQuestion + vbDefaultButton1, "Decisión de continuar...")
        If respSod = vbYes Then
            Call guardarArchivo("Sodimac")
            Call Dist_OC_C_S
        Else
            respMa = MsgBox("¿La EMPRESA a trabajar es MAESTRO?", vbYesNo + vbQuestion + vbDefaultButton1, "Decisión de continuar...")
            If respMa = vbYes Then
                Call guardarArchivo("Maestro")
                Call Dist_OC_C_M
            End If
        End If
    Loop While (respSod = vbNo And respMa = vbNo)
End Sub
Sub Dist_OC_A()
    'Formato a la hoja
    Cells.Select
    With Selection.Font
    .Name = "Calibri"
    .Size = 11
    End With
    'Elimino comentarios en caso de existir
    Range("A1:L1").ClearComments
    'Agrego titulos a la hoja
    Range("A1").FormulaLocal = "OC"
    Range("B1").FormulaLocal = "Línea"
    Range("C1").FormulaLocal = "Artículo"
    Range("D1").FormulaLocal = "Descripción"
    Range("E1").FormulaLocal = "UDM"
    Range("F1").FormulaLocal = "Cantidad"
    Range("G1").FormulaLocal = "Cuenta Cargo"
    Range("H1").FormulaLocal = "CC"
    Range("I1").FormulaLocal = "Tienda"
    Range("J1").FormulaLocal = "Importe"
    Range("K1").FormulaLocal = "Divisa"
    Range("L1").FormulaLocal = "Entregado"
    Range("L1").Select
    'Agrego leyendo en comentario en la celda L1
    ActiveCell.AddComment ("Completo = Todo" & Chr(10) & "Parcial = Indicar cantidad atendida" & Chr(10) & "Pendiente = No despachado")
    'Negrita al titulo
    Range("A1:L1").Select
    Selection.Font.Bold = True
    Range("H2").Select
End Sub
Sub Dist_OC_B()
    Application.Calculation = xlCalculationManual
    Range("H2").Select
    'Inserto formula =Extrae(G2,16,5)"
    ActiveCell.Formula = "=MID(G2,16,5)"
    'Declaro varibles
    Dim NroFilaFin, NroColumnaFin As Integer
    'Selecciono la ultima celda con datos de la columna F
    Range("H2").End(xlDown).Select
    'Obtengo el total de datos (fila y columna) a trabajar
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'Copio formula
    Range("H2").Select
    Selection.Copy
    'agrego formula a todo las filas con datos
    Range(Cells(3, 8), Cells(NroFilaFin, NroColumnaFin)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    'Se actualiza las formulas
    Application.Calculation = xlCalculationAutomatic
    'Copia y pega como valores
    Range(Cells(2, 8), Cells(NroFilaFin, NroColumnaFin)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub Dist_OC_C_S()
    'Declaro varibles
    Dim NroFilaFin, NroColumnaFin, cant, i As Integer
    'Ultimo dato de columna
    Range("H2").End(xlDown).Select
    'Captura de fila y columna
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'conteo de datos
    cant = Range(Cells(2, 8), Cells(NroFilaFin, NroColumnaFin)).Count
    'color rojo a datos observados
    Range("H2").Select
    For i = 1 To cant
        Select Case Cells(1 + i, 8).Value
            Case "50200"
            Case "50400"
            Case "50500"
            Case "50600"
            Case "50800"
            Case "51000"
            Case "51100"
            Case "51200"
            Case "51400"
            Case "51500"
            Case "51600"
            Case "51700"
            Case "51800"
            Case "52000"
            Case "52100"
            Case "52300"
            Case "52400"
            Case "52500"
            Case "52600"
            Case "52700"
            Case "52900"
            Case "53000"
            Case "53100"
            Case "53200"
            Case "57600"
            Case Else
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End Select
        ActiveCell.Offset(1, 0).Select
    Next i
    Range("Q1").Select
    ActiveCell = "SODIMAC"
    Range("H1").Select
End Sub
Sub Dist_OC_C_M()
    'Declaro varibles
    Dim NroFilaFin, NroColumnaFin, cant, i As Integer
    'Ultimo dato de columna
    Range("H2").End(xlDown).Select
    'Captura de fila y columna
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'conteo de datos
    cant = Range(Cells(2, 8), Cells(NroFilaFin, NroColumnaFin)).Count
    'color rojo a datos observados
    Range("H2").Select
    For i = 1 To cant
        Select Case Cells(1 + i, 8).Value
            Case "53501"
            Case "53502"
            Case "53503"
            Case "53504"
            Case "53505"
            Case "53506"
            Case "53507"
            Case "53508"
            Case "53509"
            Case "53510"
            Case "53511"
            Case "53512"
            Case "53513"
            Case "53514"
            Case "53515"
            Case "53516"
            Case "53517"
            Case "53518"
            Case "53519"
            Case "53520"
            Case "53521"
            Case "53522"
            Case "53523"
            Case "53524"
            Case "53525"
            Case "53526"
            Case "53527"
            Case "53528"
            Case "53529"
            Case Else
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End Select
        ActiveCell.Offset(1, 0).Select
    Next i
    Range("Q1").Select
    ActiveCell = "MAESTRO"
    Range("H1").Select
End Sub
Sub Dist_OC_2()
    'Elaborado por Hugo Garcia S.
    'V1.2 270417
    Call Dist_OC_2A
    If (Range("Q1").Value = "SODIMAC") Then
        Call Dist_OC_2B_S
    Else
        Call Dist_OC_2B_M
    End If
    Call Dist_OC_2C
    Call Dist_OC_3
End Sub
Sub Dist_OC_2A()
    'Declaro varibles
    Dim NroFilaFin, NroColumnaFin, cant, i As Integer
    'Ultimo dato de columna
    Range("H2").End(xlDown).Select
    'Captura de fila y columna
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'Quitar color a celdas
    Range(Cells(2, 8), Cells(NroFilaFin, NroColumnaFin)).Interior.ColorIndex = 0

    Range(Cells(2, 8), Cells(NroFilaFin, NroColumnaFin)).Select
    Selection.TextToColumns Destination:=Range("H2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    'Autoajuste de columna D
    Range("D1").Select
    Columns("D:D").EntireColumn.AutoFit
    'Oculto columna G
    Range("G:G").Select
    Selection.EntireColumn.Hidden = True
    Application.CutCopyMode = False
    Range("H2").Select
    'Ordenar por Centro de Costo(CC)
    Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Header:=xlYes, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    'Subtotal por CC
    Selection.Subtotal GroupBy:=8, Function:=xlSum, TotalList:=Array(10), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    Range("H2").Select
End Sub
Sub Dist_OC_2B_S()
    'Tiendas Sodimac
    Range("O1").FormulaLocal = "Total 50200"
    Range("O2").FormulaLocal = "Total 50400"
    Range("O3").FormulaLocal = "Total 50500"
    Range("O4").FormulaLocal = "Total 50600"
    Range("O5").FormulaLocal = "Total 50800"
    Range("O6").FormulaLocal = "Total 51000"
    Range("O7").FormulaLocal = "Total 51100"
    Range("O8").FormulaLocal = "Total 51200"
    Range("O9").FormulaLocal = "Total 51400"
    Range("O10").FormulaLocal = "Total 51500"
    Range("O11").FormulaLocal = "Total 51600"
    Range("O12").FormulaLocal = "Total 51700"
    Range("O13").FormulaLocal = "Total 51800"
    Range("O14").FormulaLocal = "Total 52000"
    Range("O15").FormulaLocal = "Total 52100"
    Range("O16").FormulaLocal = "Total 52300"
    Range("O17").FormulaLocal = "Total 52400"
    Range("O18").FormulaLocal = "Total 52500"
    Range("O19").FormulaLocal = "Total 52600"
    Range("O20").FormulaLocal = "Total 52700"
    Range("O21").FormulaLocal = "Total 52900"
    Range("O22").FormulaLocal = "Total 53000"
    Range("O23").FormulaLocal = "Total 53100"
    Range("O24").FormulaLocal = "Total 53200"
    Range("O25").FormulaLocal = "Total 57600"
    
    Range("P1").FormulaLocal = "San Miguel"
    Range("P2").FormulaLocal = "Megaplaza"
    Range("P3").FormulaLocal = "Angamos"
    Range("P4").FormulaLocal = "Atocongo"
    Range("P5").FormulaLocal = "Javier Prado"
    Range("P6").FormulaLocal = "Lima Centro"
    Range("P7").FormulaLocal = "Tr. Mansiche"
    Range("P8").FormulaLocal = "Chiclayo I"
    Range("P9").FormulaLocal = "Jockey Plaza"
    Range("P10").FormulaLocal = "Canta Callao"
    Range("P11").FormulaLocal = "Ica Mall"
    Range("P12").FormulaLocal = "Tr. Los Jardines"
    Range("P13").FormulaLocal = "Bellavista"
    Range("P14").FormulaLocal = "Piura"
    Range("P15").FormulaLocal = "Arequipa"
    Range("P16").FormulaLocal = "Chimbote"
    Range("P17").FormulaLocal = "Cajamarca"
    Range("P18").FormulaLocal = "Santa Anita"
    Range("P19").FormulaLocal = "San Juan Lurigancho"
    Range("P20").FormulaLocal = "Huacho"
    Range("P21").FormulaLocal = "Cañete"
    Range("P22").FormulaLocal = "VES"
    Range("P23").FormulaLocal = "Sullana"
    Range("P24").FormulaLocal = "Chiclayo II"
    Range("P25").FormulaLocal = "Huancayo"
End Sub
Sub Dist_OC_2B_M()
    'Tiendas Maestro
    Range("O1").FormulaLocal = "Total 53501"
    Range("O2").FormulaLocal = "Total 53502"
    Range("O3").FormulaLocal = "Total 53503"
    Range("O4").FormulaLocal = "Total 53504"
    Range("O5").FormulaLocal = "Total 53505"
    Range("O6").FormulaLocal = "Total 53506"
    Range("O7").FormulaLocal = "Total 53507"
    Range("O8").FormulaLocal = "Total 53508"
    Range("O9").FormulaLocal = "Total 53509"
    Range("O10").FormulaLocal = "Total 53510"
    Range("O11").FormulaLocal = "Total 53511"
    Range("O12").FormulaLocal = "Total 53512"
    Range("O13").FormulaLocal = "Total 53513"
    Range("O14").FormulaLocal = "Total 53514"
    Range("O15").FormulaLocal = "Total 53515"
    Range("O16").FormulaLocal = "Total 53516"
    Range("O17").FormulaLocal = "Total 53517"
    Range("O18").FormulaLocal = "Total 53518"
    Range("O19").FormulaLocal = "Total 53519"
    Range("O20").FormulaLocal = "Total 53520"
    Range("O21").FormulaLocal = "Total 53521"
    Range("O22").FormulaLocal = "Total 53522"
    Range("O23").FormulaLocal = "Total 53523"
    Range("O24").FormulaLocal = "Total 53524"
    Range("O25").FormulaLocal = "Total 53525"
    Range("O26").FormulaLocal = "Total 53526"
    Range("O27").FormulaLocal = "Total 53527"
    Range("O28").FormulaLocal = "Total 53528"
    Range("O29").FormulaLocal = "Total 53529"
    
    Range("P1").FormulaLocal = "Chacarilla"
    Range("P2").FormulaLocal = "Surquillo"
    Range("P3").FormulaLocal = "Pueblo Libre"
    Range("P4").FormulaLocal = "Chorrillos"
    Range("P5").FormulaLocal = "Ate"
    Range("P6").FormulaLocal = "Arequipa I"
    Range("P7").FormulaLocal = "Naranjal"
    Range("P8").FormulaLocal = "Colonial"
    Range("P9").FormulaLocal = "Callao"
    Range("P10").FormulaLocal = "Independencia"
    Range("P11").FormulaLocal = "Piura"
    Range("P12").FormulaLocal = "Chiclayo"
    Range("P13").FormulaLocal = "Trujillo"
    Range("P14").FormulaLocal = "Huancayo"
    Range("P15").FormulaLocal = "Cuzco"
    Range("P16").FormulaLocal = "Ica"
    Range("P17").FormulaLocal = "VES"
    Range("P18").FormulaLocal = "Arequipa II"
    Range("P19").FormulaLocal = "San Luis"
    Range("P20").FormulaLocal = "Tacna"
    Range("P21").FormulaLocal = "Barrios Altos"
    Range("P22").FormulaLocal = "Comas"
    Range("P23").FormulaLocal = "Cajamarca"
    Range("P24").FormulaLocal = "Sullana"
    Range("P25").FormulaLocal = "Chincha"
    Range("P26").FormulaLocal = "Puente Piedra"
    Range("P27").FormulaLocal = "SJM"
    Range("P28").FormulaLocal = "Huacho"
    Range("P29").FormulaLocal = "Ventanilla"
End Sub
Sub Dist_OC_2C()
    Range("I2").Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(H2,$O$1:$P$29,2,0),"""")"
    'Declaro varibles
    Dim NroFilaFin, NroColumnaFin As Integer
    'Selecciono la ultima celda con datos de la columna
    Range("H2").End(xlDown).Select
    'Obtengo el total de datos (fila y columna) a trabajar
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    'Copio formula
    Range("I2").Select
    Selection.Copy
    'agrego formula a todo las filas con datos
    Range(Cells(3, 9), Cells(NroFilaFin, NroColumnaFin + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    'Copia y pega como valores
    Range(Cells(2, 9), Cells(NroFilaFin, NroColumnaFin + 1)).Select
    Selection.Font.Bold = True
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O1:P29").ClearContents
    'Range("Q1").ClearContents
    Range(Cells(2, 10), Cells(NroFilaFin, NroColumnaFin + 2)).Select
    Selection.NumberFormat = "0.00"
    Columns("A:F").EntireColumn.AutoFit
    Columns("H:M").EntireColumn.AutoFit
    Range("A1").Select
End Sub
Sub Dist_OC_3()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim tienda As String
    tienda = Range("Q1").Value
    
    'abrir el archivo SD.xlsx MP.xlsx:
    If (Range("Q1").Value = "SODIMAC") Then
        Workbooks.Open Filename:="D:\Existencias Tiendas - Adquisiciones\Compras SOD-MP\00.Macro\SD.xlsx"
    Else
        Workbooks.Open Filename:="D:\Existencias Tiendas - Adquisiciones\Compras SOD-MP\00.Macro\MP.xlsx"
    End If
    
    Windows("Pedido.xlsx").Activate
    
    'validacion de datos
    Dim NroFilaFin, NroColumnaFin, cant, i As Integer
    'Ultimo dato de columna
    Range("H1").End(xlDown).Select
    'Captura de fila y columna
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    
    'inserta la formula en la hojas master de E hasta I
    Range("L2").Formula = "=IFERROR(ROUND(J2/F2,2),"""")"
    Range("N2").Formula = "=IFERROR(L2-M2,"""")"
    
    If (tienda = "SODIMAC") Then
        Range("M2").Formula = "=IFERROR(VLOOKUP(D2,[SD.xlsx]Master!$E$5:$G$5000,3,0),"""")"
        Range("O2").Formula = "=IFERROR(VLOOKUP(D2,[SD.xlsx]Master!$E$5:$I$5000,5,0),"""")"
    Else
        Range("M2").Formula = "=IFERROR(VLOOKUP(D2,[MP.xlsx]Master!$E$5:$G$5000,3,0),"""")"
        Range("O2").Formula = "=IFERROR(VLOOKUP(D2,[MP.xlsx]Master!$E$5:$I$5000,5,0),"""")"
    End If
    
    Range("L2:O2").Copy
    'agrego formula a todo las filas con datos
    Range(Cells(3, 12), Cells(NroFilaFin - 2, NroColumnaFin + 7)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("A1:O1").AutoFilter
    
    'Borra la validacion
    Range("Q1").ClearContents
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    If (tienda = "SODIMAC") Then
         Windows("SD.xlsx").Activate
         ActiveWorkbook.Close
    Else
        Windows("MP.xlsx").Activate
        ActiveWorkbook.Close
    End If
    Range("O1").Select
    ActiveWorkbook.Save
End Sub
Sub guardarArchivo(ByVal m As String)
    Dim marca, mesSig As String
    Dim mes, año As Integer
    
    'Valido si la carpeta existe
    año = Year(Date)
    mes = Month(Date)
    marca = m

    Select Case mes
        Case 1
        mesSig = "Febrero"
        Case 2
        mesSig = "Marzo"
        Case 3
        mesSig = "Abril"
        Case 4
        mesSig = "Mayo"
        Case 5
        mesSig = "Junio"
        Case 6
        mesSig = "Julio"
        Case 7
        mesSig = "Agosto"
        Case 8
        mesSig = "Septiembre"
        Case 9
        mesSig = "Octubre"
        Case 10
        mesSig = "Noviembre"
        Case 11
        mesSig = "Diciembre"
        Case 12
        mesSig = "Enero"
    End Select
    
    'Si es el mes diciembre aumenta en un año mas
    If (mesSig = "Enero") Then
        año = año + 1
    End If
    
    'Verifico que existe la carpeta con el nombre del año, caso contrario la crea
    Path = "D:\" & año
    If Dir(Path, vbDirectory) = "" Then
        MkDir Path
    End If
    
    'Verifico que existe la marca, caso contrario la crea
    If (marca = "Sodimac") Then
        Path1 = "D:\" & año & "\" & marca
        If Dir(Path1, vbDirectory) = "" Then
            MkDir Path1
        End If
    Else
        Path1 = "D:\" & año & "\" & marca
        If Dir(Path1, vbDirectory) = "" Then
            MkDir Path1
        End If
    End If
    
    'Verifico que existe la carpeta con el nombre del mes, caso contrario la crea
    Path2 = "D:\" & año & "\" & marca & "\" & mesSig
    If Dir(Path2, vbDirectory) = "" Then
        MkDir Path2
    End If
    
    'Ubico el directorio en donde se guardara
    ChDir "D:\" & año & "\" & marca & "\" & mesSig
    
    'Genero la aplicaion de guardar como
    Application.GetSaveAsFilename
End Sub
