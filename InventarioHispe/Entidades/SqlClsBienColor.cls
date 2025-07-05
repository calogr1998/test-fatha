VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsBienColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strDescripcion              As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstBienColor             As ADODB.Recordset
Private strSQLBienColor          As String



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
    
    bolEstado = False
        
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerBienColor() As Boolean
    On Error GoTo errObtenerBienColor
    
    Set rstBienColor = New ADODB.Recordset
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT "
    strSQLBienColor = strSQLBienColor & "* "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "CODIGO = '" & strCodigo & "'"
    
    If rstBienColor.State = 1 Then rstBienColor.Close
    
    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienColor.EOF Then
        strCodigo = Trim(rstBienColor!Codigo & "")
        strCodigoExterno = Trim(rstBienColor!CodExterno & "")
        strDescripcion = Trim(rstBienColor!Descripcion & "")
        
        bolEstado = CBool(rstBienColor!Estado)
        
        obtenerBienColor = True
    Else
        obtenerBienColor = False
    End If
    
    rstBienColor.Close
    
    Set rstBienColor = Nothing
    
    Exit Function
errObtenerBienColor:
    obtenerBienColor = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: ObtenerBienColor"
    
    Err.Clear
End Function

Public Sub obtenerConfigBienColor()
    On Error GoTo errObtenerConfigBienColor
    
    Set rstBienColor = New ADODB.Recordset
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT "
    strSQLBienColor = strSQLBienColor & "* "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "CODIGO = '" & strCodigo & "'"
    
    If rstBienColor.State = 1 Then rstBienColor.Close
    
    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienColor.EOF Then
        strCodigo = Trim(rstBienColor!Codigo & "")
        strCodigoExterno = Trim(rstBienColor!CodExterno & "")
        strDescripcion = Trim(rstBienColor!Descripcion & "")
        
        bolEstado = CBool(rstBienColor!Estado & "")
    End If
    
    rstBienColor.Close
    
    Set rstBienColor = Nothing
    
    Exit Sub
errObtenerConfigBienColor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: ObtenerBienColor"
    
    Err.Clear
End Sub

Public Function generarCodigoBienColor() As String
    On Error GoTo errGenerarCodigoBienColor

    Set rstBienColor = New ADODB.Recordset

    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT TOP 1 "
    strSQLBienColor = strSQLBienColor & "CODIGO "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "ORDER BY "
    strSQLBienColor = strSQLBienColor & "CODIGO DESC"

    If rstBienColor.State = 1 Then rstBienColor.Close

    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly

    If Not rstBienColor.EOF Then
        generarCodigoBienColor = Format(Val(Trim(rstBienColor!Codigo & "")) + 1, "0000")
    Else
        generarCodigoBienColor = "0001"
    End If

    rstBienColor.Close

    Set rstBienColor = Nothing

    strSQLBienColor = vbNullString

    Exit Function
errGenerarCodigoBienColor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: GenerarCodigoBienColor"

    generarCodigoBienColor = vbNullString

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstBienColor = New ADODB.Recordset
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR.* "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "CODIGO = '" & strCodigo & "'"
    
    If rstBienColor.State = 1 Then rstBienColor.Close
    
    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienColor.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstBienColor.Close
    
    strSQLBienColor = vbNullString
    
    Set rstBienColor = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarBienColor() As Boolean
    On Error GoTo errGuardarBienColor
    
    strSQLBienColor = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoBienColor
        End If
        
        strSQLBienColor = strSQLBienColor & "INSERT INTO MAESTROS.EF2BIENCOLOR("
        strSQLBienColor = strSQLBienColor & "CODIGO, CODEXTERNO, DESCRIPCION, "
        strSQLBienColor = strSQLBienColor & "ESTADO, FECREG, USUREG"
        strSQLBienColor = strSQLBienColor & ") "
        strSQLBienColor = strSQLBienColor & "VALUES("
        strSQLBienColor = strSQLBienColor & "'" & strCodigo & "', "
        strSQLBienColor = strSQLBienColor & "'" & strCodigoExterno & "', "
        strSQLBienColor = strSQLBienColor & "'" & strDescripcion & "', "
        strSQLBienColor = strSQLBienColor & IIf(bolEstado, "1", "0") & ", "
        strSQLBienColor = strSQLBienColor & "'" & strFecReg & "', "
        strSQLBienColor = strSQLBienColor & "'" & strUsuReg & "'"
        strSQLBienColor = strSQLBienColor & ")"
    Else
        strSQLBienColor = strSQLBienColor & "UPDATE "
        strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
        strSQLBienColor = strSQLBienColor & "SET "
        strSQLBienColor = strSQLBienColor & "CODEXTERNO = '" & strCodigoExterno & "', "
        strSQLBienColor = strSQLBienColor & "DESCRIPCION = '" & strDescripcion & "', "
        strSQLBienColor = strSQLBienColor & "ESTADO = " & IIf(bolEstado, "1, ", "0, ")
        strSQLBienColor = strSQLBienColor & "FECMOD = '" & strFecMod & "', "
        strSQLBienColor = strSQLBienColor & "USUMOD = '" & strUsuMod & "' "
        strSQLBienColor = strSQLBienColor & "WHERE "
        strSQLBienColor = strSQLBienColor & "CODIGO = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLBienColor
    
    guardarBienColor = True
    
    strSQLSelectAlter = strSQLBienColor
    strSQLBienColor = vbNullString
    
    Exit Function
errGuardarBienColor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: GuardarBienColor"
    
    guardarBienColor = False
    
    Err.Clear
End Function

Public Function eliminarBienColor() As Boolean
    On Error GoTo errEliminarBienColor
    
    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "CODCOLOR", "PROCESOS.IF3ORDEN", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarBienColor = False
        
        Exit Function
    End If
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "DELETE FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "CODIGO = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLBienColor
    
    eliminarBienColor = True
    
    strSQLSelectAlter = strSQLBienColor
    strSQLBienColor = vbNullString
    
    Exit Function
errEliminarBienColor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: EliminarBienColor"
    
    eliminarBienColor = False
    
    Err.Clear
End Function

Public Sub listarBienColor(ByVal comboList As Object)
    On Error GoTo errListarBienColor
    
    Set rstBienColor = New ADODB.Recordset
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT "
    strSQLBienColor = strSQLBienColor & "CODIGO, "
    strSQLBienColor = strSQLBienColor & "DESCRIPCION "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "ESTADO = 1 "
    strSQLBienColor = strSQLBienColor & "ORDER BY "
    strSQLBienColor = strSQLBienColor & "CODIGO"
    
    If rstBienColor.State = 1 Then rstBienColor.Close
    
    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstBienColor.EOF Then
        comboList.AddItem "(*) - Todos"
        
        Do While Not rstBienColor.EOF
            comboList.AddItem Trim(rstBienColor!Descripcion & "") & Space(150) & Trim(rstBienColor!Codigo & "")
            
            rstBienColor.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstBienColor.Close
    
    strSQLBienColor = vbNullString
    
    Set rstBienColor = Nothing
    
    Exit Sub
errListarBienColor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: ListarBienColor"
    
    Err.Clear
End Sub

Public Sub listarBienColorSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarBienColorSoloSeleccion
    
    Set rstBienColor = New ADODB.Recordset
    
    strSQLBienColor = vbNullString
    strSQLBienColor = strSQLBienColor & "SELECT "
    strSQLBienColor = strSQLBienColor & "CODIGO, "
    strSQLBienColor = strSQLBienColor & "DESCRIPCION "
    strSQLBienColor = strSQLBienColor & "FROM "
    strSQLBienColor = strSQLBienColor & "MAESTROS.EF2BIENCOLOR "
    strSQLBienColor = strSQLBienColor & "WHERE "
    strSQLBienColor = strSQLBienColor & "ESTADO = 1 "
    strSQLBienColor = strSQLBienColor & "ORDER BY "
    strSQLBienColor = strSQLBienColor & "CODIGO"
    
    If rstBienColor.State = 1 Then rstBienColor.Close
    
    rstBienColor.Open strSQLBienColor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstBienColor.EOF Then
        Do While Not rstBienColor.EOF
            comboList.AddItem Trim(rstBienColor!Descripcion & "") & Space(50) & Trim(rstBienColor!Codigo & "")
            
            rstBienColor.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstBienColor.Close
    
    strSQLBienColor = vbNullString
    
    Set rstBienColor = Nothing
    
    Exit Sub
errListarBienColorSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienColor: ListarBienColorSoloSeleccion"
    
    Err.Clear
End Sub
