VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strDescripcion              As String
Private strDireccion                As String
Private strRucAlmacen               As String
Private strCentroCosto              As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstAlmacen             As ADODB.Recordset
Private strSQLAlmacen          As String


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

Public Property Let Direccion(ByVal value As String)
    strDireccion = value
End Property

Public Property Get Direccion() As String
    Direccion = strDireccion
End Property

Public Property Let RucAlmacen(ByVal value As String)
    strRucAlmacen = value
End Property

Public Property Get RucAlmacen() As String
    RucAlmacen = strRucAlmacen
End Property

Public Property Let CentroCosto(ByVal value As String)
    strCentroCosto = value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
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
    strDireccion = vbNullString
    strRucAlmacen = vbNullString
    strCentroCosto = vbNullString
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerAlmacen() As Boolean
    On Error GoTo errObtenerAlmacen
    
    Set rstAlmacen = New ADODB.Recordset
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "* "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "WHERE "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM = '" & strCodigo & "'"
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAlmacen.EOF Then
        strCodigo = Trim(rstAlmacen!f2codalm & "")
        strCodigoExterno = Trim(rstAlmacen!F2CODALMEXTERNO & "")
        strDescripcion = Trim(rstAlmacen!F2NOMALM & "")
        
        strDireccion = Trim(rstAlmacen!F2DIRALM & "")
        strRucAlmacen = Trim(rstAlmacen!F2RUCALM & "")
        strCentroCosto = Trim(rstAlmacen!F4CENTRO & "")
        
        obtenerAlmacen = True
    Else
        obtenerAlmacen = False
    End If
    
    rstAlmacen.Close
    
    Set rstAlmacen = Nothing
    
    Exit Function
errObtenerAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: ObtenerAlmacen"
    
    obtenerAlmacen = False
    
    Err.Clear
End Function

Public Sub obtenerConfigAlmacen()
    On Error GoTo errObtenerConfigAlmacen
    
    Set rstAlmacen = New ADODB.Recordset
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "* "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "WHERE "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM = '" & strCodigo & "'"
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAlmacen.EOF Then
        strCodigo = Trim(rstAlmacen!f2codalm & "")
        strCodigoExterno = Trim(rstAlmacen!F2CODALMEXTERNO & "")
        strDescripcion = Trim(rstAlmacen!F2NOMALM & "")
        
        strDireccion = Trim(rstAlmacen!F2DIRALM & "")
        strRucAlmacen = Trim(rstAlmacen!F2RUCALM & "")
        strCentroCosto = Trim(rstAlmacen!F4CENTRO & "")
    End If
    
    rstAlmacen.Close
    
    Set rstAlmacen = Nothing
    
    Exit Sub
errObtenerConfigAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: ObtenerConfigAlmacen"
    
    Err.Clear
End Sub

Public Function generarCodigoAlmacen() As String
    On Error GoTo errGenerarCodigoAlmacen
    
    Set rstAlmacen = New ADODB.Recordset

    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "TOP 1 F2CODALM AS CODIGO "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "ORDER BY "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM DESC"

    If rstAlmacen.State = 1 Then rstAlmacen.Close

    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly

    If Not rstAlmacen.EOF Then
        generarCodigoAlmacen = Format(Val(rstAlmacen!Codigo & "") + 1, "000")
    Else
        generarCodigoAlmacen = "001"
    End If

    rstAlmacen.Close

    Set rstAlmacen = Nothing

    strSQLAlmacen = vbNullString

    Exit Function
errGenerarCodigoAlmacen:
    generarCodigoAlmacen = vbNullString

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: GenerarCodigoAlmacen"

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstAlmacen = New ADODB.Recordset
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES.* "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "WHERE "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM = '" & strCodigo & "'"
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAlmacen.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstAlmacen.Close
    
    strSQLAlmacen = vbNullString
    
    Set rstAlmacen = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarAlmacen() As Boolean
    On Error GoTo errGuardarAlmacen
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoAlmacen
        End If
        
        strSQLAlmacen = vbNullString
        strSQLAlmacen = strSQLAlmacen & "INSERT INTO MAESTROS.EF2ALMACENES("
        strSQLAlmacen = strSQLAlmacen & "F2CODALM, F2CODALMEXTERNO, F2NOMALM, F2DIRALM, "
        strSQLAlmacen = strSQLAlmacen & "F2RUCALM, F4CENTRO) "
        strSQLAlmacen = strSQLAlmacen & "VALUES("
        strSQLAlmacen = strSQLAlmacen & "'" & strCodigo & "', "
        strSQLAlmacen = strSQLAlmacen & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "''") & ", "
        strSQLAlmacen = strSQLAlmacen & "'" & strDescripcion & "', "
        strSQLAlmacen = strSQLAlmacen & "'" & strDireccion & "', "
        strSQLAlmacen = strSQLAlmacen & "'" & strRucAlmacen & "', "
        strSQLAlmacen = strSQLAlmacen & "'" & strCentroCosto & "'"
        strSQLAlmacen = strSQLAlmacen & ")"
    Else
        strSQLAlmacen = vbNullString
        strSQLAlmacen = strSQLAlmacen & "UPDATE "
        strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
        strSQLAlmacen = strSQLAlmacen & "SET "
        strSQLAlmacen = strSQLAlmacen & "F2CODALMEXTERNO = " & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "NULL") & ", "
        strSQLAlmacen = strSQLAlmacen & "F2NOMALM = '" & strDescripcion & "', "
        strSQLAlmacen = strSQLAlmacen & "F2DIRALM = '" & strDireccion & "', "
        strSQLAlmacen = strSQLAlmacen & "F2RUCALM = '" & strRucAlmacen & "', "
        strSQLAlmacen = strSQLAlmacen & "F4CENTRO = '" & strCentroCosto & "' "
        strSQLAlmacen = strSQLAlmacen & "WHERE "
        strSQLAlmacen = strSQLAlmacen & "F2CODALM = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLAlmacen
    
    guardarAlmacen = True
    
    strSQLSelectAlter = strSQLAlmacen
    strSQLAlmacen = vbNullString
    
    Exit Function
errGuardarAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: GuardarAlmacen"
    
    guardarAlmacen = False
    
    Err.Clear
End Function

Public Function eliminarAlmacen() As Boolean
    On Error GoTo errEliminarAlmacen
    
    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F2CODALM", "PROCESOS.IF4VALES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarAlmacen = False
        
        Exit Function
    End If
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "DELETE FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "WHERE "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLAlmacen
    
    eliminarAlmacen = True
    
    strSQLSelectAlter = strSQLAlmacen
    strSQLAlmacen = vbNullString
    
    Exit Function
errEliminarAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: EliminarAlmacen"
    
    eliminarAlmacen = False
    
    Err.Clear
End Function

Public Sub listarAlmacen(ByVal comboList As Object, _
                            ByVal bolIncluirOpcionTodos As Boolean)
                            
    On Error GoTo errListarAlmacen
    
    Set rstAlmacen = New ADODB.Recordset
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM, "
    strSQLAlmacen = strSQLAlmacen & "F2NOMALM "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "ORDER BY "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM"
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstAlmacen.EOF Then
        If bolIncluirOpcionTodos Then
            comboList.AddItem "(*) - Todos" & Space(150)
        End If
        
        Do While Not rstAlmacen.EOF
            comboList.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(150) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    strSQLAlmacen = vbNullString
    
    Set rstAlmacen = Nothing
    
    Exit Sub
errListarAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: ListarAlmacen"
    
    Err.Clear
End Sub

Public Sub listarAlmacenSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarAlmacenSoloSeleccion
    
    Set rstAlmacen = New ADODB.Recordset
    
    strSQLAlmacen = vbNullString
    strSQLAlmacen = strSQLAlmacen & "SELECT "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM, "
    strSQLAlmacen = strSQLAlmacen & "F2NOMALM "
    strSQLAlmacen = strSQLAlmacen & "FROM "
    strSQLAlmacen = strSQLAlmacen & "MAESTROS.EF2ALMACENES "
    strSQLAlmacen = strSQLAlmacen & "ORDER BY "
    strSQLAlmacen = strSQLAlmacen & "F2CODALM"
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open strSQLAlmacen, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAlmacen.EOF Then
        comboList.Clear
        
        Do While Not rstAlmacen.EOF
            comboList.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(50) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    strSQLAlmacen = vbNullString
    
    Set rstAlmacen = Nothing
    
    Exit Sub
errListarAlmacenSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsAlmacen: ListarAlmacenSoloSeleccion"
    
    Err.Clear
End Sub

