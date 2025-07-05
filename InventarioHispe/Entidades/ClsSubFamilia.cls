VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSubFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strCodigoExterno2           As String
Private strCodigoFamilia            As String
Private strDescripcion              As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstSubFamilia             As ADODB.Recordset
Private strSQLSubFamilia          As String


Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let CodigoExterno(ByVal value As String)
    strCodigoExterno = value
End Property

Public Property Get CodigoExterno() As String
    CodigoExterno = strCodigoExterno
End Property

Public Property Let CodigoExterno2(ByVal value As String)
    strCodigoExterno2 = value
End Property

Public Property Get CodigoExterno2() As String
    CodigoExterno2 = strCodigoExterno2
End Property

Public Property Let CodigoFamilia(ByVal value As String)
    strCodigoFamilia = value
End Property

Public Property Get CodigoFamilia() As String
    CodigoFamilia = strCodigoFamilia
End Property

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property



Public Property Let Estado(ByVal value As Boolean)
    bolEstado = value
End Property

Public Property Get Estado() As Boolean
    Estado = bolEstado
End Property



'Propiedad Fecha de Registro de Color de Bien
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Color de Bien
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property



Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    strCodigoExterno = vbNullString
    strCodigoExterno2 = vbNullString
    strCodigoFamilia = vbNullString
    strDescripcion = vbNullString
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerSubFamilia() As Boolean
    On Error GoTo errObtenerSubFamilia
    
    Set rstSubFamilia = New ADODB.Recordset
    
    strSQLSubFamilia = vbNullString
    strSQLSubFamilia = strSQLSubFamilia & "SELECT "
    strSQLSubFamilia = strSQLSubFamilia & "* "
    strSQLSubFamilia = strSQLSubFamilia & "FROM "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
    strSQLSubFamilia = strSQLSubFamilia & "WHERE "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSubFamilia.EOF Then
        strCodigo = Trim(rstSubFamilia!F7CODCON & "")
        strCodigoExterno = Trim(rstSubFamilia!CodExterno & "")
        strCodigoExterno2 = Trim(rstSubFamilia!CodExterno2 & "")
        strCodigoFamilia = Trim(rstSubFamilia!F7NIVEL01 & "")
        strDescripcion = Trim(rstSubFamilia!F7DESCON & "")
        
        bolEstado = CBool(rstSubFamilia!Estado)
        
        strFecReg = Trim(rstSubFamilia!FecReg & "")
        strUsuReg = Trim(rstSubFamilia!UsuReg & "")
        strFecMod = Trim(rstSubFamilia!FecMod & "")
        strUsuMod = Trim(rstSubFamilia!UsuMod & "")
        
        obtenerSubFamilia = True
    Else
        obtenerSubFamilia = False
    End If
    
    rstSubFamilia.Close
    
    Set rstSubFamilia = Nothing
    
    Exit Function
errObtenerSubFamilia:
    obtenerSubFamilia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: ObtenerSubFamilia"
    
    Err.Clear
End Function

Public Sub obtenerConfigSubFamilia()
    On Error GoTo errObtenerConfigSubFamilia
    
    Set rstSubFamilia = New ADODB.Recordset
    
    strSQLSubFamilia = vbNullString
    strSQLSubFamilia = strSQLSubFamilia & "SELECT "
    strSQLSubFamilia = strSQLSubFamilia & "* "
    strSQLSubFamilia = strSQLSubFamilia & "FROM "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
    strSQLSubFamilia = strSQLSubFamilia & "WHERE "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSubFamilia.EOF Then
        strCodigo = Trim(rstSubFamilia!F7CODCON & "")
        strCodigoExterno = Trim(rstSubFamilia!CodExterno & "")
        strCodigoExterno2 = Trim(rstSubFamilia!CodExterno2 & "")
        strCodigoFamilia = Trim(rstSubFamilia!F7NIVEL01 & "")
        strDescripcion = Trim(rstSubFamilia!F7DESCON & "")
        
        bolEstado = CBool(rstSubFamilia!Estado)
        
        strFecReg = Trim(rstSubFamilia!FecReg & "")
        strUsuReg = Trim(rstSubFamilia!UsuReg & "")
        strFecMod = Trim(rstSubFamilia!FecMod & "")
        strUsuMod = Trim(rstSubFamilia!UsuMod & "")
    End If
    
    rstSubFamilia.Close
    
    Set rstSubFamilia = Nothing
    
    Exit Sub
errObtenerConfigSubFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: ObtenerSubFamilia"
    
    Err.Clear
End Sub

Public Function generarCodigoSubFamilia() As String
    On Error GoTo errGenerarCodigoSubFamilia

    Set rstSubFamilia = New ADODB.Recordset

    strSQLSubFamilia = vbNullString
    strSQLSubFamilia = strSQLSubFamilia & "SELECT TOP 1 "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON AS CODIGO "
    strSQLSubFamilia = strSQLSubFamilia & "FROM "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
    strSQLSubFamilia = strSQLSubFamilia & "ORDER BY "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON DESC"

    If rstSubFamilia.State = 1 Then rstSubFamilia.Close

    rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    If Not rstSubFamilia.EOF Then
        generarCodigoSubFamilia = Format(Val(Trim(rstSubFamilia!Codigo & "")) + 1, "0000")
    Else
        generarCodigoSubFamilia = "0001"
    End If

    rstSubFamilia.Close

    Set rstSubFamilia = Nothing

    strSQLSubFamilia = vbNullString

    Exit Function
errGenerarCodigoSubFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: GenerarCodigoSubFamilia"

    generarCodigoSubFamilia = vbNullString

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstSubFamilia = New ADODB.Recordset
    
    strSQLSubFamilia = vbNullString
    strSQLSubFamilia = strSQLSubFamilia & "SELECT "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02.* "
    strSQLSubFamilia = strSQLSubFamilia & "FROM "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
    strSQLSubFamilia = strSQLSubFamilia & "WHERE "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSubFamilia.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstSubFamilia.Close
    
    strSQLSubFamilia = vbNullString
    
    Set rstSubFamilia = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarSubFamilia() As Boolean
    On Error GoTo errGuardarSubFamilia
    
    strSQLSubFamilia = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoSubFamilia
        End If
        
        strSQLSubFamilia = strSQLSubFamilia & "INSERT INTO SF7NIVEL02("
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON, CODEXTERNO, CODEXTERNO2, F7NIVEL01, "
        strSQLSubFamilia = strSQLSubFamilia & "F7DESCON, ESTADO, FECREG, USUREG"
        strSQLSubFamilia = strSQLSubFamilia & ") "
        strSQLSubFamilia = strSQLSubFamilia & "VALUES("
        strSQLSubFamilia = strSQLSubFamilia & "'" & strCodigo & "', "
        strSQLSubFamilia = strSQLSubFamilia & "'" & strCodigoExterno & "', "
        strSQLSubFamilia = strSQLSubFamilia & "'" & strCodigoExterno2 & "', "
        strSQLSubFamilia = strSQLSubFamilia & "'" & strCodigoFamilia & "', "
        strSQLSubFamilia = strSQLSubFamilia & "'" & strDescripcion & "', "
        strSQLSubFamilia = strSQLSubFamilia & IIf(bolEstado, "TRUE", "FALSE") & ", "
        strSQLSubFamilia = strSQLSubFamilia & "CVDATE('" & strFecReg & "'), "
        strSQLSubFamilia = strSQLSubFamilia & "'" & strUsuReg & "'"
        strSQLSubFamilia = strSQLSubFamilia & ")"
    Else
        strSQLSubFamilia = strSQLSubFamilia & "UPDATE "
        strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
        strSQLSubFamilia = strSQLSubFamilia & "SET "
        strSQLSubFamilia = strSQLSubFamilia & "CODEXTERNO = '" & strCodigoExterno & "', "
        strSQLSubFamilia = strSQLSubFamilia & "CODEXTERNO2 = '" & strCodigoExterno2 & "', "
        strSQLSubFamilia = strSQLSubFamilia & "F7NIVEL01 = '" & strCodigoFamilia & "', "
        strSQLSubFamilia = strSQLSubFamilia & "F7DESCON = '" & strDescripcion & "', "
        strSQLSubFamilia = strSQLSubFamilia & "ESTADO = " & IIf(bolEstado, "TRUE, ", "FALSE, ")
        strSQLSubFamilia = strSQLSubFamilia & "FECMOD = CVDATE('" & strFecMod & "'), "
        strSQLSubFamilia = strSQLSubFamilia & "USUMOD = '" & strUsuMod & "' "
        strSQLSubFamilia = strSQLSubFamilia & "WHERE "
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLSubFamilia
    
    guardarSubFamilia = True
    
    strSQLSelectAlter = strSQLSubFamilia
    strSQLSubFamilia = vbNullString
    
    Exit Function
errGuardarSubFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: GuardarSubFamilia"
    
    guardarSubFamilia = False
    
    Err.Clear
End Function

Public Function eliminarSubFamilia() As Boolean
    On Error GoTo errEliminarSubFamilia
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F5UBICACIO", "IF5PLA", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarSubFamilia = False
        
        Exit Function
    End If
    
    strSQLSubFamilia = vbNullString
    strSQLSubFamilia = strSQLSubFamilia & "DELETE FROM "
    strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
    strSQLSubFamilia = strSQLSubFamilia & "WHERE "
    strSQLSubFamilia = strSQLSubFamilia & "F7CODCON = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLSubFamilia
    
    eliminarSubFamilia = True
    
    strSQLSelectAlter = strSQLSubFamilia
    strSQLSubFamilia = vbNullString
    
    Exit Function
errEliminarSubFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: EliminarSubFamilia"
    
    eliminarSubFamilia = False
    
    Err.Clear
End Function

Public Sub listarSubFamilia(ByVal comboList As Object)
    On Error GoTo errListarSubFamilia
    
    Set rstSubFamilia = New ADODB.Recordset
    
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        strSQLSubFamilia = vbNullString
        strSQLSubFamilia = strSQLSubFamilia & "SELECT "
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON, "
        strSQLSubFamilia = strSQLSubFamilia & "F7DESCON "
        strSQLSubFamilia = strSQLSubFamilia & "FROM "
        strSQLSubFamilia = strSQLSubFamilia & "MAESTROS.SF7NIVEL02 "
        strSQLSubFamilia = strSQLSubFamilia & "WHERE "
        strSQLSubFamilia = strSQLSubFamilia & "(REPLACE('CODEXTERNO', ' ', '') + '') <> '' AND "
        strSQLSubFamilia = strSQLSubFamilia & "ESTADO = 1 "
            
            If strCodigoFamilia <> vbNullString Then
                strSQLSubFamilia = strSQLSubFamilia & "AND F7NIVEL01 = '" & strCodigoFamilia & "' "
            End If
            
        strSQLSubFamilia = strSQLSubFamilia & "ORDER BY "
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON"
        
        rstSubFamilia.Open strSQLSubFamilia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
        strSQLSubFamilia = vbNullString
        strSQLSubFamilia = strSQLSubFamilia & "SELECT "
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON, "
        strSQLSubFamilia = strSQLSubFamilia & "F7DESCON "
        strSQLSubFamilia = strSQLSubFamilia & "FROM "
        strSQLSubFamilia = strSQLSubFamilia & "SF7NIVEL02 "
        strSQLSubFamilia = strSQLSubFamilia & "WHERE "
        strSQLSubFamilia = strSQLSubFamilia & "TRIM(CODEXTERNO & '') <> '' AND "
        strSQLSubFamilia = strSQLSubFamilia & "ESTADO = TRUE "
    
            If strCodigoFamilia <> vbNullString Then
                strSQLSubFamilia = strSQLSubFamilia & "AND F7NIVEL01 = '" & strCodigoFamilia & "' "
            End If
    
        strSQLSubFamilia = strSQLSubFamilia & "ORDER BY "
        strSQLSubFamilia = strSQLSubFamilia & "F7CODCON"
        
        rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    comboList.Clear
    
    If Not rstSubFamilia.EOF Then
        comboList.AddItem "(*) - Todos" & Space(150)
        
        Do While Not rstSubFamilia.EOF
            comboList.AddItem Trim(rstSubFamilia!F7DESCON & "") & Space(150) & Trim(rstSubFamilia!F7CODCON & "")
            
            rstSubFamilia.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstSubFamilia.Close
    strSQLSubFamilia = vbNullString
    
    Set rstSubFamilia = Nothing
    
    Exit Sub
errListarSubFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: ListarSubFamilia"
    
    Err.Clear
End Sub

Public Sub listarSubFamiliaSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarSubFamiliaSoloSeleccion
    
    Set rstSubFamilia = New ADODB.Recordset
    
    strSQLSubFamilia = "SELECT F7CODCON, F7DESCON " & _
                        "FROM SF7NIVEL02 " & _
                        "WHERE ESTADO = TRUE " & _
                        "ORDER BY F7CODCON"
    
    If rstSubFamilia.State = 1 Then rstSubFamilia.Close
    
    rstSubFamilia.Open strSQLSubFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSubFamilia.EOF Then
        comboList.Clear
        
        Do While Not rstSubFamilia.EOF
            comboList.AddItem Trim(rstSubFamilia!Descripcion & "") & Space(50) & Trim(rstSubFamilia!Codigo & "")
            
            rstSubFamilia.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstSubFamilia.Close
    strSQLSubFamilia = vbNullString
    
    Set rstSubFamilia = Nothing
    
    Exit Sub
errListarSubFamiliaSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSubFamilia: ListarSubFamiliaSoloSeleccion"
    
    Err.Clear
End Sub


