VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String
Private strAbreviatura              As String
Private bolUMParaExportacion        As Boolean

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstUM             As ADODB.Recordset
Private strSQLUM          As String


Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let Abreviatura(ByVal value As String)
    strAbreviatura = value
End Property

Public Property Get Abreviatura() As String
    Abreviatura = strAbreviatura
End Property
'UMParaExportacion
Public Property Let UMParaExportacion(ByVal value As Boolean)
    bolUMParaExportacion = value
End Property

Public Property Get UMParaExportacion() As Boolean
    UMParaExportacion = bolUMParaExportacion
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
    strDescripcion = vbNullString
    strAbreviatura = vbNullString
    bolUMParaExportacion = False
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerUM() As Boolean
    On Error GoTo errObtenerUM
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT "
    strSQLUM = strSQLUM & "* "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "WHERE "
    strSQLUM = strSQLUM & "F7CODMED = '" & strCodigo & "'"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        strCodigo = Trim(rstUM!f7codmed & "")
        strDescripcion = Trim(rstUM!F7NOMMED & "")
        strAbreviatura = Trim(rstUM!F7SIGMED & "")
        bolUMParaExportacion = CBool(rstUM!F7IMPORT)
        
        bolEstado = CBool(rstUM!Estado)
        
        strFecReg = Trim(rstUM!FecReg & "")
        strUsuReg = Trim(rstUM!UsuReg & "")
        strFecMod = Trim(rstUM!FecMod & "")
        strUsuMod = Trim(rstUM!UsuMod & "")
        
        obtenerUM = True
    Else
        obtenerUM = False
    End If
    
    rstUM.Close
    
    Set rstUM = Nothing
    
    Exit Function
errObtenerUM:
    obtenerUM = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: ObtenerUM"
    
    Err.Clear
End Function

Public Sub obtenerConfigUM()
    On Error GoTo errObtenerConfigUM
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT "
    strSQLUM = strSQLUM & "* "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "WHERE "
    strSQLUM = strSQLUM & "F7CODMED = '" & strCodigo & "'"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        strCodigo = Trim(rstUM!f7codmed & "")
        strDescripcion = Trim(rstUM!F7NOMMED & "")
        strAbreviatura = Trim(rstUM!F7SIGMED & "")
        bolUMParaExportacion = CBool(rstUM!F7IMPORT)
        
        bolEstado = CBool(rstUM!Estado)
        
        strFecReg = Trim(rstUM!FecReg & "")
        strUsuReg = Trim(rstUM!UsuReg & "")
        strFecMod = Trim(rstUM!FecMod & "")
        strUsuMod = Trim(rstUM!UsuMod & "")
    End If
    
    rstUM.Close
    
    Set rstUM = Nothing
    
    Exit Sub
errObtenerConfigUM:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: ObtenerUM"
    
    Err.Clear
End Sub

Public Function generarCodigoUM() As String
    On Error GoTo errGenerarCodigoUM
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT TOP 1 "
    strSQLUM = strSQLUM & "CODIGOMEDIDA "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "ORDER BY "
    strSQLUM = strSQLUM & "CODIGOMEDIDA DESC"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        generarCodigoUM = Format(Val(Trim(rstUM!Codigo & "")) + 1, "0000")
    Else
        generarCodigoUM = "0001"
    End If
    
    rstUM.Close
    
    Set rstUM = Nothing
    
    strSQLUM = vbNullString
    
    Exit Function
errGenerarCodigoUM:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: GenerarCodigoUM"

    generarCodigoUM = vbNullString

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS.* "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "WHERE "
    strSQLUM = strSQLUM & "F7CODMED = '" & strCodigo & "'"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstUM.Close
    
    strSQLUM = vbNullString
    
    Set rstUM = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarUM() As Boolean
    On Error GoTo errGuardarUM
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoUM
        End If
        
        strSQLUM = vbNullString
        strSQLUM = strSQLUM & "INSERT INTO MAESTROS.EF7MEDIDAS("
        strSQLUM = strSQLUM & "F7CODMED, F7NOMMED, F7SIGMED, "
        strSQLUM = strSQLUM & "F7IMPORT, ESTADO, FECREG, USUREG"
        strSQLUM = strSQLUM & ") "
        strSQLUM = strSQLUM & "VALUES("
        strSQLUM = strSQLUM & "'" & strCodigo & "', '" & strDescripcion & "', '" & strAbreviatura & "', "
        strSQLUM = strSQLUM & IIf(bolUMParaExportacion, "1", "0") & ", "
        strSQLUM = strSQLUM & IIf(bolEstado, "1", "0") & ", "
        strSQLUM = strSQLUM & "'" & strFecReg & "', '" & strUsuReg & "'"
        strSQLUM = strSQLUM & ")"
    Else
        strSQLUM = vbNullString
        strSQLUM = strSQLUM & "UPDATE "
        strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
        strSQLUM = strSQLUM & "SET "
        strSQLUM = strSQLUM & "F7NOMMED = '" & strDescripcion & "', "
        strSQLUM = strSQLUM & "F7SIGMED = '" & strAbreviatura & "', "
        strSQLUM = strSQLUM & "F7IMPORT = " & IIf(bolUMParaExportacion, "1", "0") & ", "
        strSQLUM = strSQLUM & "ESTADO = " & IIf(bolEstado, "1", "0") & ", "
        strSQLUM = strSQLUM & "USUMOD = '" & strUsuMod & "', "
        strSQLUM = strSQLUM & "FECMOD = '" & strFecMod & "' "
        strSQLUM = strSQLUM & "WHERE "
        strSQLUM = strSQLUM & "F7CODMED = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLUM
    
    guardarUM = True
    
    strSQLSelectAlter = strSQLUM
    strSQLUM = vbNullString
    
    Exit Function
errGuardarUM:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: GuardarUM"
    
    guardarUM = False
    
    Err.Clear
End Function

Public Function eliminarUM() As Boolean
    On Error GoTo errEliminarUM
    
    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F7CODMED", "MAESTROS.IF5PLA", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarUM = False
        
        Exit Function
    End If
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "DELETE FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "WHERE "
    strSQLUM = strSQLUM & "F7CODMED = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLUM
    
    eliminarUM = True
    
    strSQLSelectAlter = strSQLUM
    strSQLUM = vbNullString
    
    Exit Function
errEliminarUM:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: EliminarUM"
    
    eliminarUM = False
    
    Err.Clear
End Function

Public Sub listarUM(ByVal comboList As Object)
    On Error GoTo errListarUM
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT "
    strSQLUM = strSQLUM & "CODIGOMEDIDA, "
    strSQLUM = strSQLUM & "F7NOMMED "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "ORDER BY "
    strSQLUM = strSQLUM & "CODIGOMEDIDA"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        comboList.Clear
        comboList.AddItem "(*) - Todos"
        
        Do While Not rstUM.EOF
            comboList.AddItem Trim(rstUM!F7NOMMED & "") & Space(150) & Trim(rstUM!CODIGOMEDIDA & "")
            
            rstUM.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstUM.Close
    strSQLUM = vbNullString
    
    Set rstUM = Nothing
    
    Exit Sub
errListarUM:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: ListarUM"
    
    Err.Clear
End Sub

Public Sub listarUMSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarUMSoloSeleccion
    
    Set rstUM = New ADODB.Recordset
    
    strSQLUM = vbNullString
    strSQLUM = strSQLUM & "SELECT "
    strSQLUM = strSQLUM & "CODIGOMEDIDA, "
    strSQLUM = strSQLUM & "F7NOMMED "
    strSQLUM = strSQLUM & "FROM "
    strSQLUM = strSQLUM & "MAESTROS.EF7MEDIDAS "
    strSQLUM = strSQLUM & "WHERE "
    strSQLUM = strSQLUM & "ESTADO = 1 "
    strSQLUM = strSQLUM & "ORDER BY "
    strSQLUM = strSQLUM & "CODIGOMEDIDA"
    
    If rstUM.State = 1 Then rstUM.Close
    
    rstUM.Open strSQLUM, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUM.EOF Then
        comboList.Clear
        
        Do While Not rstUM.EOF
            comboList.AddItem Trim(rstUM!F7NOMMED & "") & Space(50) & Trim(rstUM!CODIGOMEDIDA & "")
            
            rstUM.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstUM.Close
    
    strSQLUM = vbNullString
    
    Set rstUM = Nothing
    
    Exit Sub
errListarUMSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUM: ListarUMSoloSeleccion"
    
    Err.Clear
End Sub


