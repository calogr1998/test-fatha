VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strOrigenCliente            As String
Private strClaseCliente             As String
Private strCodTipoDocumento         As String
Private strNroDocumento             As String
Private strNombreCliente            As String
Private strDireccionCliente         As String
Private strCodDistrito              As String
Private strPais                     As String

Private strCodigo                   As String
Private strNombreAbrev              As String
Private strCodSector                As String

Private strDirRecepcion             As String
Private strDirCobranza              As String

Private strTelefono                 As String
Private strFax                      As String
Private strMovil                    As String
Private strEmail                    As String
Private strEmailCotizacion          As String
Private strWeb                      As String
Private strContacto                 As String
Private bolEsAgenteRetencion        As Boolean

Private strCodVendedor              As String
Private dblPorcentajeComision       As Double
Private strCodCobrador              As String

Private strCodFormaPago             As String
Private strObservacion              As String

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter           As String

Private rstCliente                As ADODB.Recordset
Private strSQLCliente             As String


'Propiedad Origen de Cliente
Public Property Let OrigenCliente(ByVal value As String)
    strOrigenCliente = value
End Property

Public Property Get OrigenCliente() As String
    OrigenCliente = strOrigenCliente
End Property

'Propiedad Clase de Cliente
Public Property Let ClaseCliente(ByVal value As String)
    strClaseCliente = value
End Property

Public Property Get ClaseCliente() As String
    ClaseCliente = strClaseCliente
End Property

'Propiedad Codigo de Tipo de Documento
Public Property Let CodigoTipoDocumento(ByVal value As String)
    strCodTipoDocumento = value
End Property

Public Property Get CodigoTipoDocumento() As String
    CodigoTipoDocumento = strCodTipoDocumento
End Property

'Propiedad Numero de Documento de Identidad
Public Property Let NumeroDocumento(ByVal value As String)
    strNroDocumento = value
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = strNroDocumento
End Property

'Propiedad Nombre de Cliente
Public Property Let NombreCliente(ByVal value As String)
    strNombreCliente = value
End Property

Public Property Get NombreCliente() As String
    NombreCliente = strNombreCliente
End Property

'Propiedad Direccion de Cliente
Public Property Let DireccionCliente(ByVal value As String)
    strDireccionCliente = value
End Property

Public Property Get DireccionCliente() As String
    DireccionCliente = strDireccionCliente
End Property

'Propiedad Codigo de Distrito
Public Property Let CodigoDistrito(ByVal value As String)
    strCodDistrito = value
End Property

Public Property Get CodigoDistrito() As String
    CodigoDistrito = strCodDistrito
End Property

'Propiedad Pais de Cliente
Public Property Let Pais(ByVal value As String)
    strPais = value
End Property

Public Property Get Pais() As String
    Pais = strPais
End Property




'Propiedad Codigo de Cliente
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Nombre Abreviado de Cliente
Public Property Let NombreAbreviado(ByVal value As String)
    strNombreAbrev = value
End Property

Public Property Get NombreAbreviado() As String
    NombreAbreviado = strNombreAbrev
End Property

'Propiedad Codigo de Sector de Cliente
Public Property Let CodigoSector(ByVal value As String)
    strCodSector = value
End Property

Public Property Get CodigoSector() As String
    CodigoSector = strCodSector
End Property




'Propiedad Direccion de Recepcion de Cliente
Public Property Let DireccionRecepcion(ByVal value As String)
    strDirRecepcion = value
End Property

Public Property Get DireccionRecepcion() As String
    DireccionRecepcion = strDirRecepcion
End Property

'Propiedad Direccion de Cobranza de Cliente
Public Property Let DireccionCobranza(ByVal value As String)
    strDirCobranza = value
End Property

Public Property Get DireccionCobranza() As String
    DireccionCobranza = strDirCobranza
End Property




'Propiedad Telefono de Cliente
Public Property Let Telefono(ByVal value As String)
    strTelefono = value
End Property

Public Property Get Telefono() As String
    Telefono = strTelefono
End Property

'Propiedad Fax de Cliente
Public Property Let Fax(ByVal value As String)
    strFax = value
End Property

Public Property Get Fax() As String
    Fax = strFax
End Property

'Propiedad Telefono Movil de Cliente
Public Property Let Movil(ByVal value As String)
    strMovil = value
End Property

Public Property Get Movil() As String
    Movil = strMovil
End Property

'Propiedad Email de Cliente
Public Property Let Email(ByVal value As String)
    strEmail = value
End Property

Public Property Get Email() As String
    Email = strEmail
End Property

'Propiedad Email de Cotizacion de Cliente
Public Property Let EmailCotizacion(ByVal value As String)
    strEmailCotizacion = value
End Property

Public Property Get EmailCotizacion() As String
    EmailCotizacion = strEmailCotizacion
End Property

'Propiedad Web de Cliente
Public Property Let Web(ByVal value As String)
    strWeb = value
End Property

Public Property Get Web() As String
    Web = strWeb
End Property

'Propiedad Contacto de Cliente
Public Property Let Contacto(ByVal value As String)
    strContacto = value
End Property

Public Property Get Contacto() As String
    Contacto = strContacto
End Property

'Propiedad Cliente Es Agente de Retencion
Public Property Let EsAgenteRetencion(ByVal value As Boolean)
    bolEsAgenteRetencion = value
End Property

Public Property Get EsAgenteRetencion() As Boolean
    EsAgenteRetencion = bolEsAgenteRetencion
End Property




'Propiedad Codigo de Vendedor de Cliente
Public Property Let CodigoVendedor(ByVal value As String)
    strCodVendedor = value
End Property

Public Property Get CodigoVendedor() As String
    CodigoVendedor = strCodVendedor
End Property

'Propiedad Porcentaje de Comision por Cliente
Public Property Let PorcentajeComision(ByVal value As Double)
    dblPorcentajeComision = value
End Property

Public Property Get PorcentajeComision() As Double
    PorcentajeComision = dblPorcentajeComision
End Property

'Propiedad Codigo de Cobrador de Cliente
Public Property Let CodigoCobrador(ByVal value As String)
    strCodCobrador = value
End Property

Public Property Get CodigoCobrador() As String
    CodigoCobrador = strCodCobrador
End Property


'Propiedad Codigo de Forma de Pago de Cliente
Public Property Let CodigoFormaPago(ByVal value As String)
    strCodFormaPago = value
End Property

Public Property Get CodigoFormaPago() As String
    CodigoFormaPago = strCodFormaPago
End Property

'Propiedad Observacion de Cliente
Public Property Let Observacion(ByVal value As String)
    strObservacion = value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property




'Propiedad Fecha de Registro de Cliente
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Cliente
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Cliente
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Cliente
Public Property Let UsuarioMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property


'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strOrigenCliente = vbNullString
    strClaseCliente = vbNullString
    strCodTipoDocumento = vbNullString
    strNroDocumento = vbNullString
    strNombreCliente = vbNullString
    strDireccionCliente = vbNullString
    strCodDistrito = vbNullString
    strPais = vbNullString
    
    strCodigo = vbNullString
    strNombreAbrev = vbNullString
    strCodSector = vbNullString
    
    strDirRecepcion = vbNullString
    strDirCobranza = vbNullString
    
    strTelefono = vbNullString
    strFax = vbNullString
    strMovil = vbNullString
    strEmail = vbNullString
    strEmailCotizacion = vbNullString
    strWeb = vbNullString
    strContacto = vbNullString
    bolEsAgenteRetencion = False
    
    strCodVendedor = vbNullString
    dblPorcentajeComision = 0
    strCodCobrador = vbNullString
    
    strCodFormaPago = vbNullString
    strObservacion = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerCliente() As Boolean
    On Error GoTo errObtenerCliente
    
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "* "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2CODCLI = '" & strCodigo & "'"
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCliente.EOF Then
        strOrigenCliente = IIf(Trim(rstCliente!F2TIPOCLI & "") <> vbNullString, Trim(rstCliente!F2TIPOCLI & ""), "N")
        strClaseCliente = IIf(Trim(rstCliente!F2TIPDOC & "") <> vbNullString, Trim(rstCliente!F2TIPDOC & ""), "J")
        strCodTipoDocumento = IIf(Trim(rstCliente!F2DOCCLI & "") <> vbNullString, Trim(rstCliente!F2DOCCLI & ""), "6")
        strNroDocumento = Trim(rstCliente!F2NEWRUC & "")
        strNombreCliente = Trim(rstCliente!F2nomcli & "")
        strDireccionCliente = Trim(rstCliente!F2DIRCLI & "")
        strCodDistrito = Trim(rstCliente!f2zona & "")
        strPais = Trim(rstCliente!F2PAIS & "")
        
        strCodigo = Trim(rstCliente!f2codcli & "")
        strNombreAbrev = Trim(rstCliente!F2NOMABREV & "")
        strCodSector = Trim(rstCliente!F2TIPOCLIE & "")
        
        strDirRecepcion = Trim(rstCliente!f2dirent & "")
        strDirCobranza = Trim(rstCliente!f2lugent & "")
        
        strTelefono = Trim(rstCliente!F2TELCLI & "")
        strFax = Trim(rstCliente!F2FAXCLI & "")
        strMovil = Trim(rstCliente!F2CELULAR & "")
        strEmail = Trim(rstCliente!F2EMAIL & "")
        strEmailCotizacion = Trim(rstCliente!EMAILCOTI & "")
        strWeb = Trim(rstCliente!F2WEB & "")
        strContacto = Trim(rstCliente!F2CONTACTO & "")
        bolEsAgenteRetencion = IIf(Trim(rstCliente!F2AGENTE & "") = "*", True, False)
        
        strCodVendedor = Trim(rstCliente!f2codven & "")
        dblPorcentajeComision = Val(rstCliente!F2CLICOM & "")
        strCodCobrador = Trim(rstCliente!F2CODCOB & "")
        
        strCodFormaPago = Trim(rstCliente!f2forpag & "")
        strObservacion = Trim(rstCliente!f2observa & "")
        
        strFecReg = Trim(rstCliente!FECING & "")
        strUsuReg = Trim(rstCliente!F4USEGRA & "")
        
        obtenerCliente = True
    Else
        obtenerCliente = False
    End If
    
    rstCliente.Close
    
    Set rstCliente = Nothing
    
    Exit Function
errObtenerCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: ObtenerCliente"
    End Select
    
    obtenerCliente = False
    
    Err.Clear
End Function

Public Sub obtenerConfigCliente()
    On Error GoTo errObtenerConfigCliente
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "* "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2CODCLI = '" & strCodigo & "'"
        
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCliente.EOF Then
        strOrigenCliente = IIf(Trim(rstCliente!F2TIPOCLI & "") <> vbNullString, Trim(rstCliente!F2TIPOCLI & ""), "N")
        strClaseCliente = IIf(Trim(rstCliente!F2TIPDOC & "") <> vbNullString, Trim(rstCliente!F2TIPDOC & ""), "J")
        strCodTipoDocumento = IIf(Trim(rstCliente!F2DOCCLI & "") <> vbNullString, Trim(rstCliente!F2DOCCLI & ""), "6")
        strNroDocumento = Trim(rstCliente!F2NEWRUC & "")
        strNombreCliente = Trim(rstCliente!F2nomcli & "")
        strDireccionCliente = Trim(rstCliente!F2DIRCLI & "")
        strCodDistrito = Trim(rstCliente!f2zona & "")
        strPais = Trim(rstCliente!F2PAIS & "")
        
        strCodigo = Trim(rstCliente!f2codcli & "")
        strNombreAbrev = Trim(rstCliente!F2NOMABREV & "")
        strCodSector = Trim(rstCliente!F2TIPOCLIE & "")
        
        strDirRecepcion = Trim(rstCliente!f2dirent & "")
        strDirCobranza = Trim(rstCliente!f2lugent & "")
        
        strTelefono = Trim(rstCliente!F2TELCLI & "")
        strFax = Trim(rstCliente!F2FAXCLI & "")
        strMovil = Trim(rstCliente!F2CELULAR & "")
        strEmail = Trim(rstCliente!F2EMAIL & "")
        strEmailCotizacion = Trim(rstCliente!EMAILCOTI & "")
        strWeb = Trim(rstCliente!F2WEB & "")
        strContacto = Trim(rstCliente!F2CONTACTO & "")
        bolEsAgenteRetencion = IIf(Trim(rstCliente!F2AGENTE & "") = "*", True, False)
        
        strCodVendedor = Trim(rstCliente!f2codven & "")
        dblPorcentajeComision = Val(rstCliente!F2CLICOM & "")
        strCodCobrador = Trim(rstCliente!F2CODCOB & "")
        
        strCodFormaPago = Trim(rstCliente!f2forpag & "")
        strObservacion = Trim(rstCliente!f2observa & "")
        
        strFecReg = Trim(rstCliente!FECING & "")
        strUsuReg = Trim(rstCliente!F4USEGRA & "")
    End If
    
    rstCliente.Close
    
    Set rstCliente = Nothing
    
    Exit Sub
errObtenerConfigCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsCliente: ObtenerConfigCliente"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "EF2CLIENTES.* "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2CODCLI = '" & strCodigo & "'"
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCliente.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstCliente.Close
    
    strSQLCliente = vbNullString
    
    Set rstCliente = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function verificarExistenciaPorNroDocumento() As Boolean
    On Error GoTo errVerificarExistenciaPorNroDocumento
    
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "EF2CLIENTES.* "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2NEWRUC = '" & strNroDocumento & "'"
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCliente.EOF Then
        verificarExistenciaPorNroDocumento = True
    Else
        verificarExistenciaPorNroDocumento = False
    End If
    
    rstCliente.Close
    
    Set rstCliente = Nothing
    
    Exit Function
errVerificarExistenciaPorNroDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: VerificarExistenciaPorNroDocumento"
    End Select
    
    verificarExistenciaPorNroDocumento = False
    
    Err.Clear
End Function

Public Function generarCodigoCliente() As String
    On Error GoTo errGenerarCodigoCliente
    
    Dim rsCliente As New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "F2CODCLI "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "ORDER BY "
    strSQLCliente = strSQLCliente & "F2CODCLI DESC"
    
    If rsCliente.State = 1 Then rsCliente.Close
    
    rsCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCliente.EOF Then
        generarCodigoCliente = Format(Val(rsCliente!f2codcli & "") + 1, "0000")
    Else
        generarCodigoCliente = "0001"
    End If
    
    rsCliente.Close
    
    Set rsCliente = Nothing
    
    Exit Function
errGenerarCodigoCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "NombreCliente: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: GenerarCodigoCliente"
    End Select
        
    generarCodigoCliente = vbNullString
    
    Err.Clear
End Function

Public Function guardarCliente() As Boolean
    On Error GoTo errGuardarCliente
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoCliente
        End If
        
        strSQLCliente = vbNullString
        strSQLCliente = strSQLCliente & "INSERT INTO EF2CLIENTES("
        strSQLCliente = strSQLCliente & "F2TIPOCLI, F2TIPDOC, F2DOCCLI, F2NEWRUC, "
        strSQLCliente = strSQLCliente & "F2NOMCLI, F2DIRCLI, F2ZONA, F2PAIS, F2CODCLI, "
        strSQLCliente = strSQLCliente & "F2NOMABREV, F2TIPOCLIE, F2DIRENT, F2LUGENT, "
        strSQLCliente = strSQLCliente & "F2TELCLI, F2FAXCLI, F2CELULAR, F2EMAIL, "
        strSQLCliente = strSQLCliente & "EMAILCOTI, F2WEB, F2CONTACTO, F2AGENTE, "
        strSQLCliente = strSQLCliente & "F2CODVEN, F2CLICOM, F2CODCOB, F2FORPAG, "
        strSQLCliente = strSQLCliente & "F2OBSERVA, FECING, F4USEGRA"
        strSQLCliente = strSQLCliente & ") "
        
        strSQLCliente = strSQLCliente & "VALUES("
        strSQLCliente = strSQLCliente & "'" & strOrigenCliente & "', "
        strSQLCliente = strSQLCliente & "'" & strClaseCliente & "', "
        strSQLCliente = strSQLCliente & "'" & strCodTipoDocumento & "', "
        strSQLCliente = strSQLCliente & "'" & strNroDocumento & "', "
        strSQLCliente = strSQLCliente & "'" & strNombreCliente & "', "
        strSQLCliente = strSQLCliente & "'" & strDireccionCliente & "', "
        strSQLCliente = strSQLCliente & "'" & strCodDistrito & "', "
        strSQLCliente = strSQLCliente & "'" & strPais & "', "
        strSQLCliente = strSQLCliente & "'" & strCodigo & "', "
        strSQLCliente = strSQLCliente & IIf(strNombreAbrev <> vbNullString, "'" & strNombreAbrev & "'", "NULL") & ", "
        strSQLCliente = strSQLCliente & "'" & strCodSector & "', "
        strSQLCliente = strSQLCliente & "'" & strDirRecepcion & "', "
        strSQLCliente = strSQLCliente & "'" & strDirCobranza & "', "
        strSQLCliente = strSQLCliente & "'" & strTelefono & "', "
        strSQLCliente = strSQLCliente & "'" & strFax & "', "
        strSQLCliente = strSQLCliente & "'" & strMovil & "', "
        strSQLCliente = strSQLCliente & "'" & strEmail & "', "
        strSQLCliente = strSQLCliente & "'" & strEmailCotizacion & "', "
        strSQLCliente = strSQLCliente & "'" & strWeb & "', "
        strSQLCliente = strSQLCliente & "'" & strContacto & "', "
        strSQLCliente = strSQLCliente & "'" & IIf(bolEsAgenteRetencion, "*", vbNullString) & "', "
        strSQLCliente = strSQLCliente & "'" & strCodVendedor & "', "
        strSQLCliente = strSQLCliente & dblPorcentajeComision & ", "
        strSQLCliente = strSQLCliente & "'" & strCodCobrador & "', "
        strSQLCliente = strSQLCliente & "'" & strCodFormaPago & "', "
        strSQLCliente = strSQLCliente & "'" & strObservacion & "', "
        strSQLCliente = strSQLCliente & "CVDATE('" & strFecReg & "'), "
        strSQLCliente = strSQLCliente & "'" & strUsuReg & "'"
        strSQLCliente = strSQLCliente & ")"
    Else
        strSQLCliente = vbNullString
        strSQLCliente = strSQLCliente & "UPDATE "
        strSQLCliente = strSQLCliente & "EF2CLIENTES "
        
        strSQLCliente = strSQLCliente & "SET "
        strSQLCliente = strSQLCliente & "F2TIPOCLI = '" & strOrigenCliente & "', "
        strSQLCliente = strSQLCliente & "F2TIPDOC = '" & strClaseCliente & "', "
        strSQLCliente = strSQLCliente & "F2DOCCLI = '" & strCodTipoDocumento & "', "
        strSQLCliente = strSQLCliente & "F2NEWRUC = '" & strNroDocumento & "', "
        strSQLCliente = strSQLCliente & "F2NOMCLI = '" & strNombreCliente & "', "
        strSQLCliente = strSQLCliente & "F2DIRCLI = '" & strDireccionCliente & "', "
        strSQLCliente = strSQLCliente & "F2ZONA = '" & strCodDistrito & "', "
        strSQLCliente = strSQLCliente & "F2PAIS = '" & strPais & "', "
        strSQLCliente = strSQLCliente & "F2NOMABREV = " & IIf(strNombreAbrev <> vbNullString, "'" & strNombreAbrev & "'", "NULL") & ", "
        strSQLCliente = strSQLCliente & "F2TIPOCLIE = '" & strCodSector & "', "
        strSQLCliente = strSQLCliente & "F2DIRENT = '" & strDirRecepcion & "', "
        strSQLCliente = strSQLCliente & "F2LUGENT = '" & strDirCobranza & "', "
        strSQLCliente = strSQLCliente & "F2TELCLI = '" & strTelefono & "', "
        strSQLCliente = strSQLCliente & "F2FAXCLI = '" & strFax & "', "
        strSQLCliente = strSQLCliente & "F2CELULAR = '" & strMovil & "', "
        strSQLCliente = strSQLCliente & "F2EMAIL = '" & strEmail & "', "
        strSQLCliente = strSQLCliente & "EMAILCOTI = '" & strEmailCotizacion & "', "
        strSQLCliente = strSQLCliente & "F2WEB = '" & strWeb & "', "
        strSQLCliente = strSQLCliente & "F2CONTACTO = '" & strContacto & "', "
        strSQLCliente = strSQLCliente & "F2AGENTE = '" & IIf(bolEsAgenteRetencion, "*", vbNullString) & "', "
        strSQLCliente = strSQLCliente & "F2CODVEN = '" & strCodVendedor & "', "
        strSQLCliente = strSQLCliente & "F2CLICOM = " & dblPorcentajeComision & ", "
        strSQLCliente = strSQLCliente & "F2CODCOB = '" & strCodCobrador & "', "
        strSQLCliente = strSQLCliente & "F2FORPAG = '" & strCodFormaPago & "', "
        strSQLCliente = strSQLCliente & "F2OBSERVA = '" & strObservacion & "' "
        
        strSQLCliente = strSQLCliente & "WHERE "
        strSQLCliente = strSQLCliente & "F2CODCLI = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLCliente
    
    strSQLSelectAlter = strSQLCliente
    
    strSQLCliente = vbNullString
    
    guardarCliente = True
    
    Exit Function
errGuardarCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: GuardarCliente"
    End Select
    
    guardarCliente = False
    
    Err.Clear
End Function

Public Function eliminarCliente() As Boolean
    On Error GoTo errEliminarCliente
    
    If Val(validarUsoRegistro(cnn_dbbancos, "CLIENTE", "CTA_DCTO", strCodigo, "T")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "F2CODCLI", "EF2PROD_CLIE", strNroDocumento, "T")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "CLIENTE", "LETRAS", strCodigo, "T")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "F2CODCLI", "TBVENTA_CAB", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName

        eliminarCliente = False
        
        Exit Function
    End If
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "DELETE FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2CODCLI = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLCliente
    
    eliminarCliente = True
    
    strSQLSelectAlter = strSQLCliente
    
    strSQLCliente = vbNullString
    
    Exit Function
errEliminarCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: EliminarCliente"
    End Select
    
    eliminarCliente = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Clientes en Grilla (QuamtumGrid)
Public Sub vistaCliente(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaCliente
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "C.F2CODCLI, "
    strSQLCliente = strSQLCliente & "TD.ABREVIATURA, "
    strSQLCliente = strSQLCliente & "C.F2NEWRUC, "
    strSQLCliente = strSQLCliente & "C.F2NOMCLI, "
    strSQLCliente = strSQLCliente & "C.F2DIRCLI, "
    strSQLCliente = strSQLCliente & "IIF(C.F2TIPOCLI = 'E', 'EXTRANJERO', 'NACIONAL') AS ORIGEN, "
    strSQLCliente = strSQLCliente & "IIF(C.F2TIPDOC = 'N', 'NATURAL', 'JURIDICA') AS PERSONA, "
    strSQLCliente = strSQLCliente & "C.F2TELCLI, "
    strSQLCliente = strSQLCliente & "LCASE(C.F2EMAIL) AS EMAIL, "
    strSQLCliente = strSQLCliente & "C.F2CONTACTO, "
    strSQLCliente = strSQLCliente & "C.F2AGENTE "
    
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES AS C "
    strSQLCliente = strSQLCliente & "LEFT JOIN EF2TIPODOCID AS TD ON TD.CODIGO = C.F2DOCCLI "
    
    strSQLCliente = strSQLCliente & "ORDER BY "
    strSQLCliente = strSQLCliente & "C.F2NOMCLI"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo Interno de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODCLI"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = False
            End With
            
            
            'Columna Abreviatura de Tipo de Doc. Identidad
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo D.I."
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "ABREVIATURA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAbreviatura"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Numero de Documento de Identidad
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Documento"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2NEWRUC"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroDocIdentidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            'Columna Nombre de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Nombres o Razón Social"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2NOMCLI"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNombreCliente"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = True
            End With
            
            'Columna Direccion de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Dirección Fiscal"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DIRCLI"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDireccion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
                .Visible = True
            End With
            
            'Columna Origen de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Origen"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "ORIGEN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOrigen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            'Columna Tipo de Persona
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Persona"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "PERSONA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColPersona"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 55
            End With
            
            'Columna Telefonos
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Telefono(s)"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F2TELCLI"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTelefono"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna E-mail
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "E-mail"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "EMAIL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColEmail"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Contacto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Contacto"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F2CONTACTO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColContacto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCliente
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2NEWRUC"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLCliente = vbNullString
    
    Exit Sub
errVistaCliente:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: VistaCliente"
    End Select
    
    Err.Clear
End Sub

'Vista de Detalle de Clientes Presupuestales en Grilla (QuamtumGrid)
Public Sub vistaClienteContacto(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaClienteContacto
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "C.CODIGO, "
    strSQLCliente = strSQLCliente & "TRIM(C.APEPAT & ' ' & C.APEMAT) & ', ' & TRIM(NOMBRE) AS NOMBRES, "
    strSQLCliente = strSQLCliente & "C.TELEFONO1, "
    strSQLCliente = strSQLCliente & "C.ANEXO, "
    strSQLCliente = strSQLCliente & "C.TELEFONO2, "
    strSQLCliente = strSQLCliente & "C.CELULAR, "
    strSQLCliente = strSQLCliente & "C.EMAIL, "
    strSQLCliente = strSQLCliente & "(C.OBSERVACION & '') AS OBSERVA "
    
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "CONTACTOS AS C "
    
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "C.CODCLI = '" & strCodigo & "'"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna de Codigo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODIGO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = False
            End With
            
            'Columna de Nombres
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Apellidos y Nombres"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMBRES"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNombres"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
                .Visible = True
            End With
            
            'Columna Telefono 1
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Telefono 1"
                .Color = &HC0FFFF
                .BandIndex = 1
                .DisableEditor = True
                .FieldName = "Telefono1"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTelefono1"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Anexo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 1
                .Caption = "Anexo"
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "ANEXO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAnexo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Telefono 2
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Telefono 2"
                .Color = &HC0FFFF
                .BandIndex = 1
                .DisableEditor = True
                .FieldName = "Telefono2"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTelefono2"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Numero Celular
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Telf. Movil"
                .Color = &HC0FFFF
                .BandIndex = 1
                .DisableEditor = True
                .FieldName = "Celular"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCelular"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna E-mail
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 1
                .Caption = "E-mail"
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "EMAIL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColEmail"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = True
            End With
            
            'Columna Observacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 1
                .Caption = "Observaciones"
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "OBSERVA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservacion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = True
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCliente
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODIGO"
        End With
    End If
    
    strSQLCliente = vbNullString
    
    Exit Sub
errVistaClienteContacto:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCliente: VistaClienteContacto"
    End Select
    
    Err.Clear
End Sub

Public Sub listarAnnoCliente(ByVal combo As Object)
    On Error GoTo errListarAnnoCliente
    
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "YEAR(FECHA) AS ANNO "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "GROUP BY "
    strSQLCliente = strSQLCliente & "YEAR(FECHA)"
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    combo.Clear
    combo.AddItem "(*) - Todos"
    
    If Not rstCliente.EOF Then
        
        rstCliente.MoveFirst
        
        Do While Not rstCliente.EOF
            combo.AddItem Trim(rstCliente!Anno & "")
            
            rstCliente.MoveNext
        Loop
    End If
    
    rstCliente.Close
    
    strSQLCliente = vbNullString
    
    Set rstCliente = Nothing
    
    Exit Sub
errListarAnnoCliente:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "NombreCliente: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCliente: ListarAnnoCliente"
    
    Err.Clear
End Sub

Public Sub listarTipoCliente(ByVal comboList As Object)
    On Error GoTo errListarTipoCliente
    
    Set rstCliente = New ADODB.Recordset
    
    strSQLCliente = vbNullString
    strSQLCliente = strSQLCliente & "SELECT "
    strSQLCliente = strSQLCliente & "F2NEWRUC, "
    strSQLCliente = strSQLCliente & "F2DESDOC "
    strSQLCliente = strSQLCliente & "FROM "
    strSQLCliente = strSQLCliente & "EF2CLIENTES "
    strSQLCliente = strSQLCliente & "WHERE "
    strSQLCliente = strSQLCliente & "F2TIPO IN ('C', 'A') "
    strSQLCliente = strSQLCliente & "ORDER BY "
    strSQLCliente = strSQLCliente & "F2DESDOC"
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    rstCliente.Open strSQLCliente, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstCliente.EOF Then
        
        rstCliente.MoveFirst
        
        Do While Not rstCliente.EOF
            comboList.AddItem Trim(rstCliente!F2DESDOC & "") & Space(100) & Trim(rstCliente!F2NEWRUC & "")
            
            rstCliente.MoveNext
        Loop
    End If
    
    If rstCliente.State = 1 Then rstCliente.Close
    
    strSQLCliente = vbNullString
    
    Set rstCliente = Nothing
    
    Exit Sub
errListarTipoCliente:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "NombreCliente: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCliente: ListarTipoCliente"
    
    Err.Clear
End Sub

