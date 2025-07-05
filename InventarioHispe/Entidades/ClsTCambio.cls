VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StrFecha                As String
Private dblCambio               As Double
Private dblTCCompra             As Double
Private dblTCVenta              As Double
Private bolTcOficial            As Boolean

Private strSQLSelectAlter       As String

Private rstTCambio              As ADODB.Recordset
Private strSQLTCambio           As String

Public Property Let Fecha(ByVal value As String)
    StrFecha = value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property

Public Property Let Cambio(ByVal value As Double)
    dblCambio = value
End Property

Public Property Get Cambio() As Double
    Cambio = dblCambio
End Property

Public Property Let TCCompra(ByVal value As Double)
    dblTCCompra = value
End Property

Public Property Get TCCompra() As Double
    TCCompra = dblTCCompra
End Property

Public Property Let TCVenta(ByVal value As Double)
    dblTCVenta = value
End Property

Public Property Get TCVenta() As Double
    TCVenta = dblTCVenta
End Property

Public Property Let TcOficial(ByVal value As Boolean)
    bolTcOficial = value
End Property

Public Property Get TcOficial() As Boolean
    TcOficial = bolTcOficial
End Property




Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property

Public Sub inicializarEntidades()
    StrFecha = vbNullString
    dblCambio = 0
    dblTCCompra = 0
    dblTCVenta = 0
    bolTcOficial = False
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerTCambio() As Boolean
    On Error GoTo errObtenerTCambio
    
    Set rstTCambio = New ADODB.Recordset
    
    strSQLTCambio = vbNullString
    strSQLTCambio = strSQLTCambio & "SELECT "
    strSQLTCambio = strSQLTCambio & "* "
    strSQLTCambio = strSQLTCambio & "FROM "
    strSQLTCambio = strSQLTCambio & "CAMBIOS "
    strSQLTCambio = strSQLTCambio & "WHERE "
    strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    rstTCambio.Open strSQLTCambio, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTCambio.EOF Then
        StrFecha = Trim(rstTCambio!Fecha & "")
        dblCambio = Val(rstTCambio!Cambio & "")
        dblTCCompra = Val(rstTCambio!CAMBIOCOMP & "")
        dblTCVenta = Val(rstTCambio!CAMBIO_VENTA & "")
        bolTcOficial = IIf(Trim(rstTCambio!CodMoneda & "") = "*", True, False)
        
        obtenerTCambio = True
    Else
        obtenerTCambio = False
    End If
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    Set rstTCambio = Nothing
    
    Exit Function
errObtenerTCambio:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTCambio: ObtenerTCambio"
    End Select
    
    obtenerTCambio = False
    
    Err.Clear
End Function

Public Sub obtenerConfigTCambio()
    On Error GoTo errObtenerConfigTCambio
    
    Set rstTCambio = New ADODB.Recordset
    
    strSQLTCambio = vbNullString
    strSQLTCambio = strSQLTCambio & "SELECT "
    strSQLTCambio = strSQLTCambio & "* "
    strSQLTCambio = strSQLTCambio & "FROM "
    strSQLTCambio = strSQLTCambio & "CAMBIOS "
    strSQLTCambio = strSQLTCambio & "WHERE "
    strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    rstTCambio.Open strSQLTCambio, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTCambio.EOF Then
        StrFecha = Trim(rstTCambio!Fecha & "")
        dblCambio = Val(rstTCambio!Cambio & "")
        dblTCCompra = Val(rstTCambio!CAMBIOCOMP & "")
        dblTCVenta = Val(rstTCambio!CAMBIO_VENTA & "")
        
        bolTcOficial = IIf(Trim(rstTCambio!CodMoneda & "") = "*", True, False)
    End If
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    Set rstTCambio = Nothing
    
    Exit Sub
errObtenerConfigTCambio:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTCambio: ObtenerConfigTCambio"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTCambio = New ADODB.Recordset
    
    strSQLTCambio = vbNullString
    strSQLTCambio = strSQLTCambio & "SELECT "
    strSQLTCambio = strSQLTCambio & "CAMBIOS.* "
    strSQLTCambio = strSQLTCambio & "FROM "
    strSQLTCambio = strSQLTCambio & "CAMBIOS "
    strSQLTCambio = strSQLTCambio & "WHERE "
    strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    rstTCambio.Open strSQLTCambio, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTCambio.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    strSQLTCambio = vbNullString
    
    Set rstTCambio = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTCambio: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarTCambio() As Boolean
    On Error GoTo errGuardarTCambio
    
    If Not verificarExistencia Then
        strSQLTCambio = vbNullString
        strSQLTCambio = strSQLTCambio & "INSERT INTO CAMBIOS("
        strSQLTCambio = strSQLTCambio & "FECHA, CAMBIO, CAMBIOCOMP, "
        strSQLTCambio = strSQLTCambio & "CAMBIO_VENTA, CODMONEDA"
        strSQLTCambio = strSQLTCambio & ") "
        strSQLTCambio = strSQLTCambio & "VALUES("
        strSQLTCambio = strSQLTCambio & "CVDATE('" & Format(StrFecha, "Short Date") & "'), "
        strSQLTCambio = strSQLTCambio & dblCambio & ", "
        strSQLTCambio = strSQLTCambio & dblTCCompra & ", "
        strSQLTCambio = strSQLTCambio & dblTCVenta & ", "
        strSQLTCambio = strSQLTCambio & IIf(bolTcOficial, "'*'", "NULL")
        strSQLTCambio = strSQLTCambio & ")"
    Else
        strSQLTCambio = vbNullString
        strSQLTCambio = strSQLTCambio & "UPDATE "
        strSQLTCambio = strSQLTCambio & "CAMBIOS "
        strSQLTCambio = strSQLTCambio & "SET "
        strSQLTCambio = strSQLTCambio & "CAMBIO = " & dblCambio & ", "
        strSQLTCambio = strSQLTCambio & "CAMBIOCOMP = " & dblTCCompra & ", "
        strSQLTCambio = strSQLTCambio & "CAMBIO_VENTA = " & dblTCVenta & " "
        strSQLTCambio = strSQLTCambio & "WHERE "
        strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) = CVDATE('" & StrFecha & "')"
    End If
    
    cnn_dbbancos.Execute strSQLTCambio
    
    guardarTCambio = True
    
    strSQLTCambio = vbNullString
    
    Exit Function
errGuardarTCambio:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTCambio: GuardarTCambio"
    End Select
    
    guardarTCambio = False
    
    Err.Clear
End Function

Public Function eliminarTCambio() As Boolean
    On Error GoTo errEliminarTCambio
    
    strSQLTCambio = vbNullString
    strSQLTCambio = strSQLTCambio & "DELETE FROM "
    strSQLTCambio = strSQLTCambio & "CAMBIOS "
    strSQLTCambio = strSQLTCambio & "WHERE "
    strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) = CVDATE('" & StrFecha & "')"
    
    cnn_dbbancos.Execute strSQLTCambio
    
    eliminarTCambio = True
    
    strSQLTCambio = vbNullString
    
    Exit Function
errEliminarTCambio:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTCambio: EliminarTCambio"
    End Select
    
    eliminarTCambio = False
    
    Err.Clear
End Function

Public Function obtenerFechaUltimoTCambio() As String
    On Error GoTo errObtenerFechaUltimoTCambio
    
    Set rstTCambio = New ADODB.Recordset
    
    strSQLTCambio = vbNullString
    strSQLTCambio = strSQLTCambio & "SELECT "
    strSQLTCambio = strSQLTCambio & "MAX(FECHA) AS ULTIMAFECHA "
    strSQLTCambio = strSQLTCambio & "FROM "
    strSQLTCambio = strSQLTCambio & "CAMBIOS "
    strSQLTCambio = strSQLTCambio & "WHERE "
    strSQLTCambio = strSQLTCambio & "CVDATE(FECHA) <> NULL"
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    rstTCambio.Open strSQLTCambio, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTCambio.EOF Then
        obtenerFechaUltimoTCambio = Trim(rstTCambio!ULTIMAFECHA & "")
    Else
        obtenerFechaUltimoTCambio = vbNullString
    End If
    
    If rstTCambio.State = 1 Then rstTCambio.Close
    
    strSQLTCambio = vbNullString
    
    Set rstTCambio = Nothing
    
    Exit Function
errObtenerFechaUltimoTCambio:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTCambio: ObtenerFechaUltimoTCambio"
    End Select
    
    obtenerFechaUltimoTCambio = vbNullString
    
    Err.Clear
End Function

Public Function validarTipoCambioDelDia(ByVal strFechaConsulta As String, _
                                        Optional ByVal bolMostrarMensaje As Boolean) As Boolean
    On Error GoTo errValidarTipoCambioDelDia
    
    'If Not actualizarTCambioWebSunat(strFechaConsulta, bolMostrarMensaje) Then
    If Not actualizarTCambioWebSunatV2(strFechaConsulta, bolMostrarMensaje) Then
        With objAyudaTCambio
            .Fecha = strFechaConsulta
            
            If Not .verificarExistencia Then
                validarTipoCambioDelDia = False
            Else
                validarTipoCambioDelDia = .obtenerTCambio
            End If
        End With
    Else
        validarTipoCambioDelDia = True
    End If
    
    Exit Function
errValidarTipoCambioDelDia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTCambio: ValidarTipoCambioDelDia"
    End Select
    
    validarTipoCambioDelDia = vbNullString
    
    Err.Clear
End Function

