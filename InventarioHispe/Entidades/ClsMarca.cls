VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String
Private strOrigen                   As String
Private strMarcaObc                 As String


Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstMarca             As ADODB.Recordset
Private strSQLMarca          As String

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

Public Property Let Origen(ByVal value As String)
    strOrigen = value
End Property

Public Property Get Origen() As String
    Origen = strOrigen
End Property

Public Property Let MarcaObc(ByVal value As String)
    strMarcaObc = value
End Property

Public Property Get MarcaObc() As String
    MarcaObc = strMarcaObc
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
    strOrigen = vbNullString
    strMarcaObc = vbNullString
    
    bolEstado = False
        
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerMarca() As Boolean
    On Error GoTo errObtenerMarca
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = vbNullString
    strSQLMarca = strSQLMarca & "SELECT "
    strSQLMarca = strSQLMarca & "* "
    strSQLMarca = strSQLMarca & "FROM "
    strSQLMarca = strSQLMarca & "EF2MARCAS "
    strSQLMarca = strSQLMarca & "WHERE "
    strSQLMarca = strSQLMarca & "F2CODMAR = '" & strCodigo & "'"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        strCodigo = Trim(rstMarca!f2codmar & "")
        strDescripcion = Trim(rstMarca!f2desmar & "")
        strOrigen = Trim(rstMarca!F2ORIGEN & "")
        strMarcaObc = Trim(rstMarca!MARCA_OBC & "")
        
        obtenerMarca = True
    Else
        obtenerMarca = False
    End If
    
    rstMarca.Close
    
    Set rstMarca = Nothing
    
    Exit Function
errObtenerMarca:
    obtenerMarca = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: ObtenerMarca"
    
    Err.Clear
End Function

Public Sub obtenerConfigMarca()
    On Error GoTo errObtenerConfigMarca
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = vbNullString
    strSQLMarca = strSQLMarca & "SELECT "
    strSQLMarca = strSQLMarca & "* "
    strSQLMarca = strSQLMarca & "FROM "
    strSQLMarca = strSQLMarca & "EF2MARCAS "
    strSQLMarca = strSQLMarca & "WHERE "
    strSQLMarca = strSQLMarca & "F2CODMAR = '" & strCodigo & "'"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        strCodigo = Trim(rstMarca!f2codmar & "")
        strDescripcion = Trim(rstMarca!f2desmar & "")
        strOrigen = Trim(rstMarca!F2ORIGEN & "")
        strMarcaObc = Trim(rstMarca!MARCA_OBC & "")
    End If
    
    rstMarca.Close
    
    Set rstMarca = Nothing
    
    Exit Sub
errObtenerConfigMarca:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: ObtenerMarca"
    
    Err.Clear
End Sub

Public Function generarCodigoMarca() As String
    On Error GoTo errGenerarCodigoMarca
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = "SELECT TOP 1 F2CODMAR " & _
                    "FROM EF2MARCAS " & _
                    "ORDER BY F2CODMAR DESC"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        generarCodigoMarca = Format(Val(Trim(rstMarca!Codigo & "")) + 1, "000")
    Else
        generarCodigoMarca = "001"
    End If
    
    rstMarca.Close
    
    Set rstMarca = Nothing
    
    strSQLMarca = vbNullString
    
    Exit Function
errGenerarCodigoMarca:
    generarCodigoMarca = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: GenerarCodigoMarca"
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = "SELECT EF2MARCAS.* " & _
                    "FROM EF2MARCAS " & _
                    "WHERE F2CODMAR = '" & strCodigo & "'"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstMarca.Close
    
    strSQLMarca = vbNullString
    
    Set rstMarca = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarMarca() As Boolean
    On Error GoTo errGuardarMarca
    
    strSQLMarca = vbNullString
    
    If Not verificarExistencia Then
        strCodigo = generarCodigoMarca
        
        strSQLMarca = strSQLMarca & "INSERT INTO EF2MARCAS("
        strSQLMarca = strSQLMarca & "F2CODMAR, F2DESMAR, F2ORIGEN, MARCA_OBC"
        strSQLMarca = strSQLMarca & ") "
        strSQLMarca = strSQLMarca & "VALUES("
        strSQLMarca = strSQLMarca & "'" & strCodigo & "', "
        strSQLMarca = strSQLMarca & "'" & strDescripcion & "', "
        strSQLMarca = strSQLMarca & "'" & strOrigen & "', "
        strSQLMarca = strSQLMarca & "'" & strMarcaObc & "'"
        strSQLMarca = strSQLMarca & ")"
    Else
        strSQLMarca = strSQLMarca & "UPDATE "
        strSQLMarca = strSQLMarca & "EF2MARCAS "
        strSQLMarca = strSQLMarca & "SET "
        strSQLMarca = strSQLMarca & "F2DESMAR = '" & strDescripcion & "', "
        strSQLMarca = strSQLMarca & "F2ORIGEN = '" & strOrigen & "', "
        strSQLMarca = strSQLMarca & "MARCA_OBC = '" & strMarcaObc & "' "
        strSQLMarca = strSQLMarca & "WHERE "
        strSQLMarca = strSQLMarca & "F2CODMAR = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLMarca
    
    guardarMarca = True
    
    strSQLSelectAlter = strSQLMarca
    strSQLMarca = vbNullString
    
    Exit Function
errGuardarMarca:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: GuardarMarca"
    
    guardarMarca = False
    
    Err.Clear
End Function

Public Function eliminarMarca() As Boolean
    On Error GoTo errEliminarMarca
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F2CODMAR", "IF3ORDEN", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarMarca = False
        
        Exit Function
    End If
    
    strSQLMarca = vbNullString
    strSQLMarca = strSQLMarca & "DELETE FROM "
    strSQLMarca = strSQLMarca & "EF2MARCAS "
    strSQLMarca = strSQLMarca & "WHERE "
    strSQLMarca = strSQLMarca & "F2CODMAR = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLMarca
    
    eliminarMarca = True
    
    strSQLSelectAlter = strSQLMarca
    strSQLMarca = vbNullString
    
    Exit Function
errEliminarMarca:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: EliminarMarca"
    
    eliminarMarca = False
    
    Err.Clear
End Function

Public Sub listarMarca(ByVal comboList As Object)
    On Error GoTo errListarMarca
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = vbNullString
    strSQLMarca = strSQLMarca & "SELECT "
    strSQLMarca = strSQLMarca & "F2CODMAR, F2DESMAR "
    strSQLMarca = strSQLMarca & "FROM "
    strSQLMarca = strSQLMarca & "EF2MARCAS "
    strSQLMarca = strSQLMarca & "ORDER BY "
    strSQLMarca = strSQLMarca & "F2CODMAR"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        comboList.Clear
        comboList.AddItem "(*) - Todos"
        
        Do While Not rstMarca.EOF
            comboList.AddItem Trim(rstMarca!f2desmar & "") & Space(150) & Trim(rstMarca!f2codmar & "")
            
            rstMarca.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstMarca.Close
    
    strSQLMarca = vbNullString
    
    Set rstMarca = Nothing
    
    Exit Sub
errListarMarca:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: ListarMarca"
    
    Err.Clear
End Sub

Public Sub listarMarcaSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarMarcaSoloSeleccion
    
    Set rstMarca = New ADODB.Recordset
    
    strSQLMarca = vbNullString
    strSQLMarca = strSQLMarca & "SELECT "
    strSQLMarca = strSQLMarca & "F2CODMAR, F2DESMAR "
    strSQLMarca = strSQLMarca & "FROM "
    strSQLMarca = strSQLMarca & "EF2MARCAS "
    strSQLMarca = strSQLMarca & "ORDER BY "
    strSQLMarca = strSQLMarca & "F2CODMAR"
    
    If rstMarca.State = 1 Then rstMarca.Close
    
    rstMarca.Open strSQLMarca, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstMarca.EOF Then
        comboList.Clear
        
        Do While Not rstMarca.EOF
            comboList.AddItem Trim(rstMarca!f2desmar & "") & Space(50) & Trim(rstMarca!f2codmar & "")
            
            rstMarca.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstMarca.Close
    
    strSQLMarca = vbNullString
    
    Set rstMarca = Nothing
    
    Exit Sub
errListarMarcaSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsMarca: ListarMarcaSoloSeleccion"
    
    Err.Clear
End Sub


