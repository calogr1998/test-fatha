VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodExterno1              As String
Private strCodExterno2              As String
Private strNombreCompleto           As String
Private strContrasena               As String
Private strMail                     As String
Private strSerieFacturacion         As String
Private strTelefono                 As String
Private strCorreo                   As String
Private strCargo                    As String
Private strUserMail                 As String
Private strPassMail                 As String



Private strSQLSelectAlter           As String

Private rstUsuario            As ADODB.Recordset
Private strSQLUsuario         As String


'Propiedad Codigo de Usuario
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Codigo Externo de Usuario 1
Public Property Let CodigoExterno1(ByVal value As String)
    strCodExterno1 = value
End Property

Public Property Get CodigoExterno1() As String
    CodigoExterno1 = strCodExterno1
End Property

'Propiedad Codigo Externo de Usuario 2
Public Property Let CodigoExterno2(ByVal value As String)
    strCodExterno2 = value
End Property

Public Property Get CodigoExterno2() As String
    CodigoExterno2 = strCodExterno2
End Property

'Propiedad Nombre de Completo de Usuario
Public Property Let NombreCompleto(ByVal value As String)
    strNombreCompleto = value
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = strNombreCompleto
End Property

'Private strContrasena               As String
Public Property Let Contrasena(ByVal value As String)
    strContrasena = value
End Property

Public Property Get Contrasena() As String
    Contrasena = strContrasena
End Property

'Private strMail                     As String
Public Property Let Mail(ByVal value As String)
    strMail = value
End Property

Public Property Get Mail() As String
    Mail = strMail
End Property

'Private strSerieFacturacion                    As String
Public Property Let SerieFacturacion(ByVal value As String)
    strSerieFacturacion = value
End Property

Public Property Get SerieFacturacion() As String
    SerieFacturacion = strSerieFacturacion
End Property

'Private strTelefono                 As String
Public Property Let Telefono(ByVal value As String)
    strTelefono = value
End Property

Public Property Get Telefono() As String
    Telefono = strTelefono
End Property

'Private strCorreo                   As String
Public Property Let Correo(ByVal value As String)
    strCorreo = value
End Property

Public Property Get Correo() As String
    Correo = strCorreo
End Property

'Private strCargo                    As String
Public Property Let Cargo(ByVal value As String)
    strCargo = value
End Property

Public Property Get Cargo() As String
    Cargo = strCargo
End Property

'Private strUserMail                 As String
Public Property Let UserMail(ByVal value As String)
    strUserMail = value
End Property

Public Property Get UserMail() As String
    UserMail = strUserMail
End Property

'Private strPassMail                 As String
Public Property Let PassMail(ByVal value As String)
    strPassMail = value
End Property

Public Property Get PassMail() As String
    PassMail = strPassMail
End Property




'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    strCodExterno1 = vbNullString
    strCodExterno2 = vbNullString
    strNombreCompleto = vbNullString
    strContrasena = vbNullString
    strMail = vbNullString
    strSerieFacturacion = vbNullString
    strTelefono = vbNullString
    strCorreo = vbNullString
    strCargo = vbNullString
    strUserMail = vbNullString
    strPassMail = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerUsuario() As Boolean
    On Error GoTo errObtenerUsuario
    
    Set rstUsuario = New ADODB.Recordset
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "SELECT "
    strSQLUsuario = strSQLUsuario & "* "
    strSQLUsuario = strSQLUsuario & "FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
    strSQLUsuario = strSQLUsuario & "WHERE "
    strSQLUsuario = strSQLUsuario & "F2CODUSER = '" & strCodigo & "'"
    
    If rstUsuario.State = 1 Then rstUsuario.Close
    
    rstUsuario.Open strSQLUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUsuario.EOF Then
        strCodigo = Trim(rstUsuario!F2CODUSER & "")
        strCodExterno1 = Trim(rstUsuario!F2CODUSEREXTERNO & "")
        strCodExterno2 = Trim(rstUsuario!F2CODUSEREXTERNO2 & "")
        strNombreCompleto = Trim(rstUsuario!F2NOMUSER & "")
        
        strContrasena = Trim(rstUsuario!F2PASUSER & "")
        strMail = Trim(rstUsuario!Mail & "")
        strSerieFacturacion = Trim(rstUsuario!NUMSERIE & "")
        strTelefono = Trim(rstUsuario!F2TELUSER & "")
        strCorreo = Trim(rstUsuario!F2CORREO & "")
        strCargo = Trim(rstUsuario!F2CARGO & "")
        strUserMail = Trim(rstUsuario!F2USEMAIL & "")
        strPassMail = Trim(rstUsuario!F2PASWMAIL & "")
        
        obtenerUsuario = True
    Else
        obtenerUsuario = False
    End If
    
    rstUsuario.Close
    
    Set rstUsuario = Nothing
    
    Exit Function
errObtenerUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsUsuario: ObtenerUsuario"
    End Select
    
    obtenerUsuario = False
    
    Err.Clear
End Function

Public Sub obtenerConfigUsuario()
    On Error GoTo errObtenerConfigUsuario
    Set rstUsuario = New ADODB.Recordset
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "SELECT "
    strSQLUsuario = strSQLUsuario & "* "
    strSQLUsuario = strSQLUsuario & "FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
    strSQLUsuario = strSQLUsuario & "WHERE "
    strSQLUsuario = strSQLUsuario & "F2CODUSER = '" & strCodigo & "'"
        
    If rstUsuario.State = 1 Then rstUsuario.Close
    
    rstUsuario.Open strSQLUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUsuario.EOF Then
        strCodigo = Trim(rstUsuario!F2CODUSER & "")
        strCodExterno1 = Trim(rstUsuario!F2CODUSEREXTERNO & "")
        strCodExterno2 = Trim(rstUsuario!F2CODUSEREXTERNO2 & "")
        strNombreCompleto = Trim(rstUsuario!F2NOMUSER & "")
        
        strContrasena = Trim(rstUsuario!F2PASUSER & "")
        strMail = Trim(rstUsuario!Mail & "")
        strSerieFacturacion = Trim(rstUsuario!NUMSERIE & "")
        strTelefono = Trim(rstUsuario!F2TELUSER & "")
        strCorreo = Trim(rstUsuario!F2CORREO & "")
        strCargo = Trim(rstUsuario!F2CARGO & "")
        strUserMail = Trim(rstUsuario!F2USEMAIL & "")
        strPassMail = Trim(rstUsuario!F2PASWMAIL & "")
    End If
    
    rstUsuario.Close
    
    Set rstUsuario = Nothing
    
    Exit Sub
errObtenerConfigUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - SqlClsUsuario: ObtenerConfigUsuario"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstUsuario = New ADODB.Recordset
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "SELECT "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS.* "
    strSQLUsuario = strSQLUsuario & "FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
    strSQLUsuario = strSQLUsuario & "WHERE "
    strSQLUsuario = strSQLUsuario & "F2CODUSER = '" & strCodigo & "'"
    
    If rstUsuario.State = 1 Then rstUsuario.Close
    
    rstUsuario.Open strSQLUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUsuario.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstUsuario.Close
    
    Set rstUsuario = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsUsuario: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

'Public Function generarCodigo() As String
'    On Error GoTo errGenerarCodigo
'
'    Dim rsUsuario As New ADODB.Recordset
'
'    strSQLUsuario = vbNullString
'    strSQLUsuario = strSQLUsuario & "SELECT "
'    strSQLUsuario = strSQLUsuario & "F2CODUSER "
'    strSQLUsuario = strSQLUsuario & "FROM "
'    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
'    strSQLUsuario = strSQLUsuario & "ORDER BY "
'    strSQLUsuario = strSQLUsuario & "F2CODUSER DESC"
'
'    If rsUsuario.State = 1 Then rsUsuario.Close
'
'    rsUsuario.Open strSQLUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
'
'    If Not rsUsuario.EOF Then
'        generarCodigo = Format(Val(strCodigo) + 1, "00")
'    Else
'        generarCodigo = "01"
'    End If
'
'    rsUsuario.Close
'
'    Set rsUsuario = Nothing
'
'    Exit Function
'errGenerarCodigo:
'    Select Case Err.Number
'        Case 3704, 3709
'            cnBdCPlus.Open strCadenaConexioBdCPlus
'
'            Resume
'        Case Else
'            MsgBox "No. Error: " & Err.Number & vbNewLine & _
'                    "TipoCadena: " & Err.Description, _
'                    vbCritical, App.ProductName & " - SqlClsUsuario: GenerarCodigo"
'    End Select
'
'    generarCodigo = vbNullString
'
'    Err.Clear
'End Function

Public Function guardarUsuario() As Boolean
    On Error GoTo errGuardarUsuario
    
    strSQLUsuario = vbNullString
    
    If Not verificarExistencia Then
        strSQLUsuario = vbNullString
        strSQLUsuario = strSQLUsuario & "INSERT INTO MAESTROS.EF2USERS("
        strSQLUsuario = strSQLUsuario & "F2CODUSER, F2CODUSEREXTERNO, F2CODUSEREXTERNO2, "
        strSQLUsuario = strSQLUsuario & "F2NOMUSER, F2PASUSER, [MAIL], NUMSERIE, F2TELUSER, "
        strSQLUsuario = strSQLUsuario & "F2CORREO, F2CARGO, F2USEMAIL, F2PASWMAIL"
        strSQLUsuario = strSQLUsuario & ") "
        
        strSQLUsuario = strSQLUsuario & "VALUES("
        strSQLUsuario = strSQLUsuario & "'" & strCodigo & "', "
        strSQLUsuario = strSQLUsuario & "'" & strCodExterno1 & "', "
        strSQLUsuario = strSQLUsuario & "'" & strCodExterno2 & "', "
        strSQLUsuario = strSQLUsuario & "'" & strNombreCompleto & "', "
        strSQLUsuario = strSQLUsuario & "'" & strContrasena & "', "
        strSQLUsuario = strSQLUsuario & "'" & strMail & "', "
        strSQLUsuario = strSQLUsuario & "'" & strSerieFacturacion & "', "
        strSQLUsuario = strSQLUsuario & "'" & strTelefono & "', "
        strSQLUsuario = strSQLUsuario & "'" & strCorreo & "', "
        strSQLUsuario = strSQLUsuario & "'" & strCargo & "', "
        strSQLUsuario = strSQLUsuario & "'" & strUserMail & "', "
        strSQLUsuario = strSQLUsuario & "'" & strPassMail & "'"
        strSQLUsuario = strSQLUsuario & ")"
    Else
        strSQLUsuario = strSQLUsuario & "UPDATE "
        strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
        
        strSQLUsuario = strSQLUsuario & "SET "
        strSQLUsuario = strSQLUsuario & "F2CODUSEREXTERNO = '" & strCodExterno1 & "', "
        strSQLUsuario = strSQLUsuario & "F2CODUSEREXTERNO2 = '" & strCodExterno2 & "', "
        strSQLUsuario = strSQLUsuario & "F2NOMUSER = '" & strNombreCompleto & "', "
        strSQLUsuario = strSQLUsuario & "F2PASUSER = '" & strContrasena & "', "
        strSQLUsuario = strSQLUsuario & "MAIL = '" & strMail & "', "
        strSQLUsuario = strSQLUsuario & "NUMSERIE = '" & strSerieFacturacion & "', "
        strSQLUsuario = strSQLUsuario & "F2TELUSER = '" & strTelefono & "', "
        strSQLUsuario = strSQLUsuario & "F2CORREO = '" & strCorreo & "', "
        strSQLUsuario = strSQLUsuario & "F2CARGO = '" & strCargo & "', "
        strSQLUsuario = strSQLUsuario & "F2USEMAIL = '" & strUserMail & "', "
        strSQLUsuario = strSQLUsuario & "F2PASWMAIL = '" & strPassMail & "' "
        strSQLUsuario = strSQLUsuario & "WHERE "
        strSQLUsuario = strSQLUsuario & "F2CODUSER = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLUsuario
    
    strSQLSelectAlter = strSQLUsuario
    strSQLUsuario = vbNullString
    
    guardarUsuario = True
    
    Exit Function
errGuardarUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsUsuario: GuardarUsuario"
    End Select
    
    guardarUsuario = False
    
    Err.Clear
End Function

Public Function eliminarUsuario() As Boolean
    On Error GoTo errEliminarUsuario
    
'    If Val(validarUsoRegistro(cnBdCPlus, "F2ZONA", "EF2CLIENTES", strCodigo, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarUsuario = False
'
'        Exit Function
'    End If
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "DELETE FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
    strSQLUsuario = strSQLUsuario & "WHERE "
    strSQLUsuario = strSQLUsuario & "F2CODUSER = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLUsuario
    
    eliminarUsuario = True
    
    Exit Function
errEliminarUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsUsuario: EliminarUsuario"
    End Select
    
    eliminarUsuario = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Usuarios en Grilla (QuamtumGrid)
Public Sub vistaUsuario(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaUsuario
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "SELECT "
    strSQLUsuario = strSQLUsuario & "D.F2CODUSER, "
    strSQLUsuario = strSQLUsuario & "D.F2DESZON "
    
    strSQLUsuario = strSQLUsuario & "FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS AS D "
    
    strSQLUsuario = strSQLUsuario & "ORDER BY "
    strSQLUsuario = strSQLUsuario & "D.F2CODUSER"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODUSER"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = True
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DESZON"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLUsuario
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2CODUSER"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLUsuario = vbNullString
    
    Exit Sub
errVistaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsUsuario: VistaUsuario"
    End Select
    
    Err.Clear
End Sub

Public Sub listarUsuario(ByVal comboList As Object, _
                            Optional ByVal bolSoloUsuariosConCodExterno2 As Boolean)
                            
    On Error GoTo errListarUsuario
    
    Set rstUsuario = New ADODB.Recordset
    
    strSQLUsuario = vbNullString
    strSQLUsuario = strSQLUsuario & "SELECT "
    strSQLUsuario = strSQLUsuario & "F2CODUSER, "
    strSQLUsuario = strSQLUsuario & "F2DESZON "
    strSQLUsuario = strSQLUsuario & "FROM "
    strSQLUsuario = strSQLUsuario & "MAESTROS.EF2USERS "
    strSQLUsuario = strSQLUsuario & "WHERE "
    strSQLUsuario = strSQLUsuario & "F2CODUSER <> '' "
        
        If bolSoloUsuariosConCodExterno2 Then
            strSQLUsuario = strSQLUsuario & "AND F2CODUSEREXTERNO1 <> '' "
        End If
        
    strSQLUsuario = strSQLUsuario & "ORDER BY "
    strSQLUsuario = strSQLUsuario & "F2DESZON"
    
    If rstUsuario.State = 1 Then rstUsuario.Close
    
    rstUsuario.Open strSQLUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstUsuario.EOF Then
        
        rstUsuario.MoveFirst
        
        Do While Not rstUsuario.EOF
            comboList.AddItem Trim(rstUsuario!F2DESZON & "") & Space(100) & Trim(rstUsuario!F2CODUSER & "")
            
            rstUsuario.MoveNext
        Loop
    End If
    
    If rstUsuario.State = 1 Then rstUsuario.Close
    
    strSQLUsuario = vbNullString
    
    Set rstUsuario = Nothing
    
    Exit Sub
errListarUsuario:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsUsuario: ListarUsuario"
    
    Err.Clear
End Sub

