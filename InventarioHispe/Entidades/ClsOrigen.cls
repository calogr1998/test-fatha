VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOrigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strDescripcion              As String
Private strTipoMovimiento           As String
Private bolRegistrarCosto           As Boolean
Private bolTieneAlmacenDestino      As Boolean
Private strCodAyudaProducto         As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstOrigen             As ADODB.Recordset
Private strSQLOrigen          As String

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

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let TipoMovimiento(ByVal value As String)
    strTipoMovimiento = value
End Property

Public Property Get TipoMovimiento() As String
    TipoMovimiento = strTipoMovimiento
End Property

Public Property Let RegistrarCosto(ByVal value As Boolean)
    bolRegistrarCosto = value
End Property

Public Property Get RegistrarCosto() As Boolean
    RegistrarCosto = bolRegistrarCosto
End Property

Public Property Let TieneAlmacenDestino(ByVal value As Boolean)
    bolTieneAlmacenDestino = value
End Property

Public Property Get TieneAlmacenDestino() As Boolean
    TieneAlmacenDestino = bolTieneAlmacenDestino
End Property
'CodAyudaProducto
Public Property Let CodigoAyudaProducto(ByVal value As String)
    strCodAyudaProducto = value
End Property

Public Property Get CodigoAyudaProducto() As String
    CodigoAyudaProducto = strCodAyudaProducto
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
    strDescripcion = vbNullString
    strTipoMovimiento = vbNullString
    bolRegistrarCosto = False
    bolTieneAlmacenDestino = False
    strCodAyudaProducto = vbNullString
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerOrigen() As Boolean
    On Error GoTo errObtenerOrigen
    
    Set rstOrigen = New ADODB.Recordset
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        strSQLOrigen = vbNullString
        strSQLOrigen = strSQLOrigen & "SELECT "
        strSQLOrigen = strSQLOrigen & "* "
        strSQLOrigen = strSQLOrigen & "FROM "
        strSQLOrigen = strSQLOrigen & "MAESTROS.SF1ORIGENES "
        strSQLOrigen = strSQLOrigen & "WHERE "
        strSQLOrigen = strSQLOrigen & "F1CODORI = '" & strCodigo & "'"
        
        If rstOrigen.State = 1 Then rstOrigen.Close
        
        rstOrigen.Open strSQLOrigen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
        strSQLOrigen = vbNullString
        strSQLOrigen = strSQLOrigen & "SELECT "
        strSQLOrigen = strSQLOrigen & "* "
        strSQLOrigen = strSQLOrigen & "FROM "
        strSQLOrigen = strSQLOrigen & "SF1ORIGENES "
        strSQLOrigen = strSQLOrigen & "WHERE "
        strSQLOrigen = strSQLOrigen & "F1CODORI = '" & strCodigo & "'"
        
        If rstOrigen.State = 1 Then rstOrigen.Close
        
        rstOrigen.Open strSQLOrigen, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not rstOrigen.EOF Then
        strCodigo = Trim(rstOrigen!F1CODORI & "")
        strCodigoExterno = Trim(rstOrigen!F1CODORIEXTERNO & "")
        strDescripcion = Trim(rstOrigen!F1NOMORI & "")
        
        strTipoMovimiento = Trim(rstOrigen!F1TIPMOV & "")
        bolRegistrarCosto = IIf(Trim(rstOrigen!F1COSTO & "") = "*", True, False)
        bolTieneAlmacenDestino = IIf(Trim(rstOrigen!CODALMDES & "") = "*", True, False)
        strCodAyudaProducto = Trim(rstOrigen!CODAYUDAPRODUCTO & "")
        
        bolEstado = CBool(rstOrigen!Estado)
        
        obtenerOrigen = True
    Else
        obtenerOrigen = False
    End If
    
    rstOrigen.Close
    
    Set rstOrigen = Nothing
    
    Exit Function
errObtenerOrigen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: ObtenerOrigen"
    
    obtenerOrigen = False
    
    Err.Clear
End Function

Public Sub obtenerConfigOrigen()
    On Error GoTo errObtenerConfigOrigen
    
    Set rstOrigen = New ADODB.Recordset
    
    strSQLOrigen = vbNullString
    strSQLOrigen = strSQLOrigen & "SELECT "
    strSQLOrigen = strSQLOrigen & "* "
    strSQLOrigen = strSQLOrigen & "FROM "
    strSQLOrigen = strSQLOrigen & "SF1ORIGENES "
    strSQLOrigen = strSQLOrigen & "WHERE "
    strSQLOrigen = strSQLOrigen & "F1CODORI = '" & strCodigo & "'"
    
    If rstOrigen.State = 1 Then rstOrigen.Close
    
    rstOrigen.Open strSQLOrigen, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrigen.EOF Then
        strCodigo = Trim(rstOrigen!F1CODORI & "")
        strCodigoExterno = Trim(rstOrigen!F1CODORIEXTERNO & "")
        strDescripcion = Trim(rstOrigen!F1NOMORI & "")
        
        strTipoMovimiento = Trim(rstOrigen!F1TIPMOV & "")
        bolRegistrarCosto = IIf(Trim(rstOrigen!F1COSTO & "") = "*", True, False)
        bolTieneAlmacenDestino = IIf(Trim(rstOrigen!CODALMDES & "") = "*", True, False)
        strCodAyudaProducto = Trim(rstOrigen!CODAYUDAPRODUCTO & "")
        
        bolEstado = CBool(rstOrigen!Estado)
    End If
    
    rstOrigen.Close
    
    Set rstOrigen = Nothing
    
    Exit Sub
errObtenerConfigOrigen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: ObtenerConfigOrigen"
    
    Err.Clear
End Sub

Public Function generarCodigoOrigen() As String
    On Error GoTo errGenerarCodigoOrigen
    
    Set rstOrigen = New ADODB.Recordset
    
    strSQLOrigen = vbNullString
    strSQLOrigen = strSQLOrigen & "SELECT "
    strSQLOrigen = strSQLOrigen & "TOP 1 MID(F1CODORI, " & Len(strTipoMovimiento) + 1 & ") AS CODIGO "
    strSQLOrigen = strSQLOrigen & "FROM "
    strSQLOrigen = strSQLOrigen & "SF1ORIGENES "
    strSQLOrigen = strSQLOrigen & "WHERE "
    strSQLOrigen = strSQLOrigen & "LEFT(F1CODORI, " & Len(strTipoMovimiento) & ") = '" & strTipoMovimiento & "' "
    strSQLOrigen = strSQLOrigen & "ORDER BY "
    strSQLOrigen = strSQLOrigen & "F1CODORI DESC"
    
    If rstOrigen.State = 1 Then rstOrigen.Close
    
    rstOrigen.Open strSQLOrigen, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrigen.EOF Then
        generarCodigoOrigen = strTipoMovimiento & Format(Val(rstOrigen!Codigo & "") + 1, "000")
    Else
        generarCodigoOrigen = strTipoMovimiento & "001"
    End If
    
    rstOrigen.Close
    
    Set rstOrigen = Nothing
    
    strSQLOrigen = vbNullString
    
    Exit Function
errGenerarCodigoOrigen:
    generarCodigoOrigen = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: GenerarCodigoOrigen"
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstOrigen = New ADODB.Recordset
    
    strSQLOrigen = vbNullString
    strSQLOrigen = strSQLOrigen & "SELECT "
    strSQLOrigen = strSQLOrigen & "SF1ORIGENES.* "
    strSQLOrigen = strSQLOrigen & "FROM "
    strSQLOrigen = strSQLOrigen & "SF1ORIGENES "
    strSQLOrigen = strSQLOrigen & "WHERE "
    strSQLOrigen = strSQLOrigen & "F1CODORI = '" & strCodigo & "'"
    
    If rstOrigen.State = 1 Then rstOrigen.Close
    
    rstOrigen.Open strSQLOrigen, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrigen.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstOrigen.Close
    
    strSQLOrigen = vbNullString
    
    Set rstOrigen = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarOrigen() As Boolean
    On Error GoTo errGuardarOrigen
    
    ''cnn_dbbancos.BeginTrans
    
    If Not verificarExistencia Then
        strCodigo = generarCodigoOrigen
        
        strSQLOrigen = vbNullString
        strSQLOrigen = strSQLOrigen & "INSERT INTO SF1ORIGENES("
        strSQLOrigen = strSQLOrigen & "F1CODORI, F1CODORIEXTERNO, F1NOMORI, F1TIPMOV, "
        strSQLOrigen = strSQLOrigen & "F1COSTO, CODALMDES, CODAYUDAPRODUCTO, ESTADO, "
        strSQLOrigen = strSQLOrigen & "FECREG, USUREG) "
        strSQLOrigen = strSQLOrigen & "VALUES("
        strSQLOrigen = strSQLOrigen & "'" & strCodigo & "', "
        strSQLOrigen = strSQLOrigen & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "NULL") & ", "
        strSQLOrigen = strSQLOrigen & "'" & strDescripcion & "', "
        strSQLOrigen = strSQLOrigen & "'" & strTipoMovimiento & "', "
        strSQLOrigen = strSQLOrigen & "'" & IIf(bolRegistrarCosto, "*", vbNullString) & "', "
        strSQLOrigen = strSQLOrigen & "'" & IIf(bolTieneAlmacenDestino, "*", vbNullString) & "', "
        strSQLOrigen = strSQLOrigen & "'" & strCodAyudaProducto & "', "
        strSQLOrigen = strSQLOrigen & IIf(bolEstado, "TRUE", "FALSE") & ", "
        strSQLOrigen = strSQLOrigen & "CVDATE('" & strFecReg & "'), "
        strSQLOrigen = strSQLOrigen & "'" & strUsuReg & "')"
    Else
        strSQLOrigen = vbNullString
        strSQLOrigen = strSQLOrigen & "UPDATE "
        strSQLOrigen = strSQLOrigen & "SF1ORIGENES "
        strSQLOrigen = strSQLOrigen & "SET "
        strSQLOrigen = strSQLOrigen & "F1CODORIEXTERNO = " & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "NULL") & ", "
        strSQLOrigen = strSQLOrigen & "F1NOMORI = '" & strDescripcion & "', "
        strSQLOrigen = strSQLOrigen & "F1COSTO = '" & IIf(bolRegistrarCosto, "*", vbNullString) & "', "
        strSQLOrigen = strSQLOrigen & "CODALMDES = '" & IIf(bolTieneAlmacenDestino, "*", vbNullString) & "', "
        strSQLOrigen = strSQLOrigen & "CODAYUDAPRODUCTO = '" & strCodAyudaProducto & "', "
        strSQLOrigen = strSQLOrigen & "ESTADO = " & IIf(bolEstado, "TRUE, ", "FALSE, ")
        strSQLOrigen = strSQLOrigen & "FECMOD = CVDATE('" & strFecMod & "'), "
        strSQLOrigen = strSQLOrigen & "USUMOD = '" & strUsuMod & "' "
        strSQLOrigen = strSQLOrigen & "WHERE "
        strSQLOrigen = strSQLOrigen & "F1CODORI = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLOrigen
    
    ''cnn_dbbancos.CommitTrans
    
    guardarOrigen = True
    
    strSQLSelectAlter = strSQLOrigen
    strSQLOrigen = vbNullString
    
    Exit Function
errGuardarOrigen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: GuardarOrigen"
    
    ''cnn_dbbancos.RollbackTrans
    
    guardarOrigen = False
    
    Err.Clear
End Function

Public Function eliminarOrigen() As Boolean
    On Error GoTo errEliminarOrigen
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F1CODORI", "IF4VALES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarOrigen = False
        
        Exit Function
    End If
    
    ''cnn_dbbancos.BeginTrans
    
    strSQLOrigen = "DELETE FROM SF1ORIGENES " & _
                        "WHERE F1CODORI = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLOrigen
    
    ''cnn_dbbancos.CommitTrans
    
    eliminarOrigen = True
    
    strSQLSelectAlter = strSQLOrigen
    strSQLOrigen = vbNullString
    
    Exit Function
errEliminarOrigen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrigen: EliminarOrigen"
    
    ''cnn_dbbancos.RollbackTrans
    
    eliminarOrigen = False
    
    Err.Clear
End Function
