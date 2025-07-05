VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strBase                     As String
Private strCuentaContable           As String
Private strDescripcion              As String
Private strTipoGasto                As String
Private strMoneda                   As String
Private strGrupoFlujo               As String


Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstGasto             As ADODB.Recordset
Private strSQLGasto          As String

Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Base(ByVal value As String)
    strBase = value
End Property

Public Property Get Base() As String
    Base = strBase
End Property

Public Property Let CuentaContable(ByVal value As String)
    strCuentaContable = value
End Property

Public Property Get CuentaContable() As String
    CuentaContable = strCuentaContable
End Property

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property
'TipoGasto
Public Property Let TipoGasto(ByVal value As String)
    strTipoGasto = value
End Property

Public Property Get TipoGasto() As String
    TipoGasto = strTipoGasto
End Property
'Moneda
Public Property Let Moneda(ByVal value As String)
    strMoneda = value
End Property

Public Property Get Moneda() As String
    Moneda = strMoneda
End Property
'GrupoFlujo
Public Property Let GrupoFlujo(ByVal value As String)
    strGrupoFlujo = value
End Property

Public Property Get GrupoFlujo() As String
    GrupoFlujo = strGrupoFlujo
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
    strBase = vbNullString
    strCuentaContable = vbNullString
    strDescripcion = vbNullString
    strTipoGasto = vbNullString
    strMoneda = vbNullString
    strGrupoFlujo = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerGasto() As Boolean
    On Error GoTo errObtenerGasto
    
    Set rstGasto = New ADODB.Recordset
    
    strSQLGasto = vbNullString
    strSQLGasto = strSQLGasto & "SELECT "
    strSQLGasto = strSQLGasto & "* "
    strSQLGasto = strSQLGasto & "FROM "
    strSQLGasto = strSQLGasto & "BF9GIN "
    strSQLGasto = strSQLGasto & "WHERE "
    strSQLGasto = strSQLGasto & "CODIGO = '" & strCodigo & "'"
    
    If rstGasto.State = 1 Then rstGasto.Close
    
    rstGasto.Open strSQLGasto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstGasto.EOF Then
        strCodigo = Trim(rstGasto!Codigo & "")
        strBase = Trim(rstGasto!Base & "")
        strCuentaContable = Trim(rstGasto!CUENTA & "")
        
        strDescripcion = Trim(rstGasto!nombre & "")
        strTipoGasto = Trim(rstGasto!Tipo & "")
        strMoneda = Trim(rstGasto!Moneda & "")
        strGrupoFlujo = Trim(rstGasto!GrupoFlujo & "")
        
        obtenerGasto = True
    Else
        obtenerGasto = False
    End If
    
    rstGasto.Close
    
    Set rstGasto = Nothing
    
    Exit Function
errObtenerGasto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: ObtenerGasto"
    
    obtenerGasto = False
    
    Err.Clear
End Function

Public Sub obtenerConfigGasto()
    On Error GoTo errObtenerConfigGasto
    
    Set rstGasto = New ADODB.Recordset
    
    strSQLGasto = vbNullString
    strSQLGasto = strSQLGasto & "SELECT "
    strSQLGasto = strSQLGasto & "* "
    strSQLGasto = strSQLGasto & "FROM "
    strSQLGasto = strSQLGasto & "BF9GIN "
    strSQLGasto = strSQLGasto & "WHERE "
    strSQLGasto = strSQLGasto & "CODIGO = '" & strCodigo & "'"
    
    If rstGasto.State = 1 Then rstGasto.Close
    
    rstGasto.Open strSQLGasto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstGasto.EOF Then
        strCodigo = Trim(rstGasto!Codigo & "")
        strBase = Trim(rstGasto!Base & "")
        strCuentaContable = Trim(rstGasto!CUENTA & "")
        
        strDescripcion = Trim(rstGasto!nombre & "")
        strTipoGasto = Trim(rstGasto!Tipo & "")
        strMoneda = Trim(rstGasto!Moneda & "")
        strGrupoFlujo = Trim(rstGasto!GrupoFlujo & "")
    End If
    
    rstGasto.Close
    
    Set rstGasto = Nothing
    
    Exit Sub
errObtenerConfigGasto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: ObtenerConfigGasto"
    
    Err.Clear
End Sub

Public Function generarCodigoGasto() As String
    On Error GoTo errGenerarCodigoGasto
    
    Set rstGasto = New ADODB.Recordset
    
    strSQLGasto = vbNullString
    strSQLGasto = strSQLGasto & "SELECT "
    strSQLGasto = strSQLGasto & "TOP 1 CODIGO "
    strSQLGasto = strSQLGasto & "FROM "
    strSQLGasto = strSQLGasto & "BF9GIN "
    strSQLGasto = strSQLGasto & "WHERE "
    strSQLGasto = strSQLGasto & "ISNUMERIC(CODIGO) "
    strSQLGasto = strSQLGasto & "ORDER BY "
    strSQLGasto = strSQLGasto & "CODIGO DESC"
    
    If rstGasto.State = 1 Then rstGasto.Close
    
    rstGasto.Open strSQLGasto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstGasto.EOF Then
        generarCodigoGasto = Format(Val(rstGasto!Codigo & "") + 1, "0000")
    Else
        generarCodigoGasto = "0001"
    End If
    
    rstGasto.Close
    
    Set rstGasto = Nothing
    
    strSQLGasto = vbNullString
    
    Exit Function
errGenerarCodigoGasto:
    generarCodigoGasto = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: GenerarCodigoGasto"
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstGasto = New ADODB.Recordset
    
    strSQLGasto = vbNullString
    strSQLGasto = strSQLGasto & "SELECT "
    strSQLGasto = strSQLGasto & "BF9GIN.* "
    strSQLGasto = strSQLGasto & "FROM "
    strSQLGasto = strSQLGasto & "BF9GIN "
    strSQLGasto = strSQLGasto & "WHERE "
    strSQLGasto = strSQLGasto & "CODIGO = '" & strCodigo & "'"
    
    If rstGasto.State = 1 Then rstGasto.Close
    
    rstGasto.Open strSQLGasto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstGasto.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstGasto.Close
    
    strSQLGasto = vbNullString
    
    Set rstGasto = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarGasto() As Boolean
    On Error GoTo errGuardarGasto
    
    ''cnn_dbbancos.BeginTrans
    
    If Not verificarExistencia Then
        strCodigo = generarCodigoGasto
        
        strSQLGasto = vbNullString
        strSQLGasto = strSQLGasto & "INSERT INTO BF9GIN("
        strSQLGasto = strSQLGasto & "CODIGO, BASE, CUENTA, NOMBRE, "
        strSQLGasto = strSQLGasto & "TIPO, MONEDA, GRUPOFLUJO) "
        strSQLGasto = strSQLGasto & "VALUES("
        strSQLGasto = strSQLGasto & "'" & strCodigo & "', "
        strSQLGasto = strSQLGasto & "'" & strBase & "', "
        strSQLGasto = strSQLGasto & "'" & strCuentaContable & "', "
        strSQLGasto = strSQLGasto & "'" & strDescripcion & "', "
        strSQLGasto = strSQLGasto & "'" & strTipoGasto & "', "
        strSQLGasto = strSQLGasto & "'" & strMoneda & "', "
        strSQLGasto = strSQLGasto & "'" & strGrupoFlujo & "'"
        strSQLGasto = strSQLGasto & ")"
    Else
        strSQLGasto = vbNullString
        strSQLGasto = strSQLGasto & "UPDATE "
        strSQLGasto = strSQLGasto & "BF9GIN "
        strSQLGasto = strSQLGasto & "SET "
        strSQLGasto = strSQLGasto & "BASE = '" & strBase & "', "
        strSQLGasto = strSQLGasto & "CUENTA = '" & strCuentaContable & "', "
        strSQLGasto = strSQLGasto & "TIPO = '" & strTipoGasto & "', "
        strSQLGasto = strSQLGasto & "MONEDA = '" & strMoneda & "', "
        strSQLGasto = strSQLGasto & "GRUPOFLUJO = '" & strGrupoFlujo & "', "
        strSQLGasto = strSQLGasto & "WHERE "
        strSQLGasto = strSQLGasto & "CODIGO = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLGasto
    
    ''cnn_dbbancos.CommitTrans
    
    guardarGasto = True
    
    strSQLSelectAlter = strSQLGasto
    strSQLGasto = vbNullString
    
    Exit Function
errGuardarGasto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: GuardarGasto"
    
    ''cnn_dbbancos.RollbackTrans
    
    guardarGasto = False
    
    Err.Clear
End Function

Public Function eliminarGasto() As Boolean
    On Error GoTo errEliminarGasto
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F3GASTO", "REGISMOV", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarGasto = False
        
        Exit Function
    End If
    
    ''cnn_dbbancos.BeginTrans
    
    strSQLGasto = "DELETE FROM BF9GIN " & _
                        "WHERE CODIGO = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLGasto
    
    ''cnn_dbbancos.CommitTrans
    
    eliminarGasto = True
    
    strSQLSelectAlter = strSQLGasto
    strSQLGasto = vbNullString
    
    Exit Function
errEliminarGasto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsGasto: EliminarGasto"
    
    ''cnn_dbbancos.RollbackTrans
    
    eliminarGasto = False
    
    Err.Clear
End Function


