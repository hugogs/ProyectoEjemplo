Attribute VB_Name = "GuardarArchivo"
Sub GuardarArchivo()
    Dim marca, mesSig As String
    Dim mes, año As Integer
    'Valido si la carpeta existe (ByVal m As String)
    año = Year(Date)
    mes = Month(Date)
    marca = m
    marca = "Maestro" 'A DESACTIVAR
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
