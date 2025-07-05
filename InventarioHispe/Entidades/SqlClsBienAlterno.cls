VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsBienAlterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigoBien               As String
Private strCodigoBienAlterno        As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter           As String

Private rstBienAlterno                As ADODB.Recordset
Private strSQLBienAlterno             As String

Public Property Let CodigoBien(ByVal value As String)
    strCodigoBien = value
End Property

Public Property Get CodigoBien() As String
    CodigoBien = strCodigoBien
End Property

Public Property Let CodigoBienAlterno(ByVal value As String)
    strCodigoBienAlterno = value
End Property

Public Property Get CodigoBienAlterno() As String
    CodigoBienAlterno = strCodigoBienAlterno
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
    strCodigoBien = vbNullString
    strCodigoBienAlterno = vbNullString
    
    bolEstado = False
        
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerBienAlterno() As Boolean
    On Error GoTo errObtenerBienAlterno
    
    Set rstBienAlterno = New ADODB.Recordset
    
    strSQLBienAlterno = vbNullString
    strSQLBienAlterno = strSQLBienAlterno & "SELECT "
    strSQLBienAlterno = strSQLBienAlterno & "* "
    strSQLBienAlterno = strSQLBienAlterno & "FROM "
    strSQLBienAlterno = strSQLBienAlterno & "EF2BIENALTERNO "
    strSQLBienAlterno = strSQLBienAlterno & "WHERE "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO = '" & strCodigoBien & "' AND "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPROALTERNO = '" & strCodigoBienAlterno & "'"
    
    If rstBienAlterno.State = 1 Then rstBienAlterno.Close
    
    rstBienAlterno.Open strSQLBienAlterno, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienAlterno.EOF Then
        strCodigoBien = Trim(rstBienAlterno!f5codpro & "")
        strCodigoBienAlterno = Trim(rstBienAlterno!F5CODPROALTERNO & "")
        
        bolEstado = CBool(rstBienAlterno!Estado)
        
        obtenerBienAlterno = True
    Else
        obtenerBienAlterno = False
    End If
    
    rstBienAlterno.Close
    
    Set rstBienAlterno = Nothing
    
    Exit Function
errObtenerBienAlterno:
    obtenerBienAlterno = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienAlterno: ObtenerBienAlterno"
    
    Err.Clear
End Function

Public Sub obtenerConfigBienAlterno()
    On Error GoTo errObtenerConfigBienAlterno
    
    Set rstBienAlterno = New ADODB.Recordset
    
    strSQLBienAlterno = vbNullString
    strSQLBienAlterno = strSQLBienAlterno & "SELECT "
    strSQLBienAlterno = strSQLBienAlterno & "* "
    strSQLBienAlterno = strSQLBienAlterno & "FROM "
    strSQLBienAlterno = strSQLBienAlterno & "EF2BIENALTERNO "
    strSQLBienAlterno = strSQLBienAlterno & "WHERE "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO = '" & strCodigoBien & "' AND "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPROALTERNO = '" & strCodigoBienAlterno & "'"
    
    If rstBienAlterno.State = 1 Then rstBienAlterno.Close
    
    rstBienAlterno.Open strSQLBienAlterno, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienAlterno.EOF Then
        strCodigoBien = Trim(rstBienAlterno!f5codpro & "")
        strCodigoBienAlterno = Trim(rstBienAlterno!F5CODPROALTERNO & "")
        
        bolEstado = CBool(rstBienAlterno!Estado)
    End If
    
    rstBienAlterno.Close
    
    Set rstBienAlterno = Nothing
    
    Exit Sub
errObtenerConfigBienAlterno:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienAlterno: ObtenerBienAlterno"
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstBienAlterno = New ADODB.Recordset
    
    strSQLBienAlterno = vbNullString
    strSQLBienAlterno = strSQLBienAlterno & "SELECT "
    strSQLBienAlterno = strSQLBienAlterno & "EF2BIENALTERNO.* "
    strSQLBienAlterno = strSQLBienAlterno & "FROM "
    strSQLBienAlterno = strSQLBienAlterno & "EF2BIENALTERNO "
    strSQLBienAlterno = strSQLBienAlterno & "WHERE "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO = '" & strCodigoBien & "' AND "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPROALTERNO = '" & strCodigoBienAlterno & "'"
    
    If rstBienAlterno.State = 1 Then rstBienAlterno.Close
    
    rstBienAlterno.Open strSQLBienAlterno, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienAlterno.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstBienAlterno.Close
    
    strSQLBienAlterno = vbNullString
    
    Set rstBienAlterno = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienAlterno: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarBienAlterno() As Boolean
    On Error GoTo errGuardarBienAlterno
    
    strSQLBienAlterno = vbNullString
    
    If Not verificarExistencia Then
        strSQLBienAlterno = strSQLBienAlterno & "INSERT INTO EF2BIENALTERNO("
        strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO, F5CODPROALTERNO, ESTADO, FECREG, USUREG) "
        strSQLBienAlterno = strSQLBienAlterno & "VALUES("
        strSQLBienAlterno = strSQLBienAlterno & "'" & strCodigoBien & "', "
        strSQLBienAlterno = strSQLBienAlterno & "'" & strCodigoBienAlterno & "', "
        strSQLBienAlterno = strSQLBienAlterno & IIf(bolEstado, "TRUE", "FALSE") & ", "
        strSQLBienAlterno = strSQLBienAlterno & "CVDATE('" & strFecReg & "'), "
        strSQLBienAlterno = strSQLBienAlterno & "'" & strUsuReg & "')"
    Else
        strSQLBienAlterno = strSQLBienAlterno & "UPDATE "
        strSQLBienAlterno = strSQLBienAlterno & "EF2BIENALTERNO "
        strSQLBienAlterno = strSQLBienAlterno & "SET "
        strSQLBienAlterno = strSQLBienAlterno & "ESTADO = " & IIf(bolEstado, "TRUE, ", "FALSE, ")
        strSQLBienAlterno = strSQLBienAlterno & "FECMOD = CVDATE('" & strFecMod & "'), "
        strSQLBienAlterno = strSQLBienAlterno & "USUMOD = '" & strUsuMod & "' "
        strSQLBienAlterno = strSQLBienAlterno & "WHERE "
        strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO = '" & strCodigoBien & "' AND "
        strSQLBienAlterno = strSQLBienAlterno & "F5CODPROALTERNO = '" & strCodigoBienAlterno & "'"
    End If
    
    cnn_dbbancos.Execute strSQLBienAlterno
    
    guardarBienAlterno = True
    
    strSQLBienAlterno = vbNullString
    
    Exit Function
errGuardarBienAlterno:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienAlterno: GuardarBienAlterno"
    
    guardarBienAlterno = False
    
    Err.Clear
End Function

Public Function eliminarBienAlterno() As Boolean
    On Error GoTo errEliminarBienAlterno
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F5CODPROORIGINAL", "IF3VALES", strCodigoBien & "' AND F5CODPRO = '" & strCodigoBienAlterno, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarBienAlterno = False
        
        Exit Function
    End If
    
    strSQLBienAlterno = vbNullString
    strSQLBienAlterno = strSQLBienAlterno & "DELETE "
    strSQLBienAlterno = strSQLBienAlterno & "FROM EF2BIENALTERNO "
    strSQLBienAlterno = strSQLBienAlterno & "WHERE "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPRO = '" & strCodigoBien & "' AND "
    strSQLBienAlterno = strSQLBienAlterno & "F5CODPROALTERNO = '" & strCodigoBienAlterno & "'"
    
    cnn_dbbancos.Execute strSQLBienAlterno
    
    eliminarBienAlterno = True
    
    strSQLBienAlterno = vbNullString
    
    Exit Function
errEliminarBienAlterno:
    eliminarBienAlterno = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBienAlterno: EliminarBienAlterno"
    
    Err.Clear
End Function
