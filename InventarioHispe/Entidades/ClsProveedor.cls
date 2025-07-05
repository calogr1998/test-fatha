VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strOrigenProveedor          As String 'F2TIPPROV
Private strClaseProveedor           As String 'F2TIPO_PERSONA
Private strCodTipoDocumento         As String 'TIPPROV
Private strNroDocumento             As String 'F2NEWRUC
Private strNombreProveedor          As String 'F2NOMPROV
Private strDireccionProveedor       As String 'F2DIRPROV
Private strCodDistrito              As String 'F7CODPOS


Private strCodigo                   As String 'F2CODPROV
Private strCodigoExterno            As String 'F2CODPROVEXTERNO
Private strNombreAbrev              As String 'F2NOMABREV
Private lngCodCategoria             As Long 'INTCODCATEGORIA

Private strCodigoMoneda                As String 'F2TIPMON
Private strCuentaContable           As String 'F5CODCTA
Private strCodigoGasto              As String 'F2CODGAS
Private strGrupoFlujo               As String 'F2GRUPO
Private strGrupoResultado           As String 'F2GRUPORESULTADO

Private strTelefono                 As String 'F2TELPROV
Private strFax                      As String 'F2FAXPROV
Private strEmail                    As String 'F2EMAIL
Private strCuentaAbono                      As String 'F2WEB
Private strContacto                 As String 'F2CONTACTO

Private strArticuloProveedor        As String 'F2ARTPROV
Private strCodTipoComprobante       As String 'F2TIPDOC
Private bolEsAptoParaOrden          As Boolean 'F2ORDEN
Private strCodFormaPago             As String 'F2FORPAG
Private strObservacion              As String 'F2OBSPROV


Private strFecReg                   As String 'FECHA_CREACION
Private strUsuReg                   As String 'USUARIO_CREA
Private strFecMod                   As String 'FECHA_MODIFICACION
Private strUsuMod                   As String 'USUARIO_MOD

Private strSQLSelectAlter           As String

Private rstProveedor                As ADODB.Recordset
Private strSQLProveedor             As String


'Propiedad Origen de Proveedor
Public Property Let OrigenProveedor(ByVal Value As String)
    strOrigenProveedor = Value
End Property

Public Property Get OrigenProveedor() As String
    OrigenProveedor = strOrigenProveedor
End Property

'Propiedad Clase de Proveedor
Public Property Let ClaseProveedor(ByVal Value As String)
    strClaseProveedor = Value
End Property

Public Property Get ClaseProveedor() As String
    ClaseProveedor = strClaseProveedor
End Property

'Propiedad Codigo de Tipo de Documento
Public Property Let CodigoTipoDocumento(ByVal Value As String)
    strCodTipoDocumento = Value
End Property

Public Property Get CodigoTipoDocumento() As String
    CodigoTipoDocumento = strCodTipoDocumento
End Property

'Propiedad Numero de Documento de Identidad
Public Property Let NumeroDocumento(ByVal Value As String)
    strNroDocumento = Value
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = strNroDocumento
End Property

'Propiedad Nombre de Proveedor
Public Property Let NombreProveedor(ByVal Value As String)
    strNombreProveedor = Value
End Property

Public Property Get NombreProveedor() As String
    NombreProveedor = strNombreProveedor
End Property

'Propiedad Direccion de Proveedor
Public Property Let DireccionProveedor(ByVal Value As String)
    strDireccionProveedor = Value
End Property

Public Property Get DireccionProveedor() As String
    DireccionProveedor = strDireccionProveedor
End Property

'Propiedad Codigo de Distrito
Public Property Let CodigoDistrito(ByVal Value As String)
    strCodDistrito = Value
End Property

Public Property Get CodigoDistrito() As String
    CodigoDistrito = strCodDistrito
End Property




'Propiedad Codigo de Proveedor
Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Codigo Externo de Proveedor
Public Property Let CodigoExterno(ByVal Value As String)
    strCodigoExterno = Value
End Property

Public Property Get CodigoExterno() As String
    CodigoExterno = strCodigoExterno
End Property

'Propiedad Nombre Abreviado de Proveedor
Public Property Let NombreAbreviado(ByVal Value As String)
    strNombreAbrev = Value
End Property

Public Property Get NombreAbreviado() As String
    NombreAbreviado = strNombreAbrev
End Property

'Propiedad Codigo de Categoria de Proveedor
Public Property Let CodigoCategoria(ByVal Value As Long)
    lngCodCategoria = Value
End Property

Public Property Get CodigoCategoria() As Long
    CodigoCategoria = lngCodCategoria
End Property




'Propiedad Codigo de Moneda de Proveedor
Public Property Let CodigoMoneda(ByVal Value As String)
    strCodigoMoneda = Value
End Property

Public Property Get CodigoMoneda() As String
    CodigoMoneda = strCodigoMoneda
End Property

'Propiedad Cuenta Contable de Proveedor
Public Property Let CuentaContable(ByVal Value As String)
    strCuentaContable = Value
End Property

Public Property Get CuentaContable() As String
    CuentaContable = strCuentaContable
End Property

'Propiedad Codigo de Gasto de Proveedor
Public Property Let CodigoGasto(ByVal Value As String)
    strCodigoGasto = Value
End Property

Public Property Get CodigoGasto() As String
    CodigoGasto = strCodigoGasto
End Property

'Propiedad Grupo de Flujo de Proveedor
Public Property Let GrupoFlujo(ByVal Value As String)
    strGrupoFlujo = Value
End Property

Public Property Get GrupoFlujo() As String
    GrupoFlujo = strGrupoFlujo
End Property

'Propiedad Grupo de Resultados de Proveedor
Public Property Let GrupoResultado(ByVal Value As String)
    strGrupoResultado = Value
End Property

Public Property Get GrupoResultado() As String
    GrupoResultado = strGrupoResultado
End Property


'Propiedad Telefono de Proveedor
Public Property Let Telefono(ByVal Value As String)
    strTelefono = Value
End Property

Public Property Get Telefono() As String
    Telefono = strTelefono
End Property

'Propiedad Fax de Proveedor
Public Property Let Fax(ByVal Value As String)
    strFax = Value
End Property

Public Property Get Fax() As String
    Fax = strFax
End Property

'Propiedad Email de Proveedor
Public Property Let Email(ByVal Value As String)
    strEmail = Value
End Property

Public Property Get Email() As String
    Email = strEmail
End Property

'Propiedad Web de Proveedor
Public Property Let CuentaAbono(ByVal Value As String)
    strCuentaAbono = Value
End Property

Public Property Get CuentaAbono() As String
    CuentaAbono = strCuentaAbono
End Property

'Propiedad Contacto de Proveedor
Public Property Let Contacto(ByVal Value As String)
    strContacto = Value
End Property

Public Property Get Contacto() As String
    Contacto = strContacto
End Property



'Propiedad Articulos de Proveedor
Public Property Let ArticuloProveedor(ByVal Value As String)
    strArticuloProveedor = Value
End Property

Public Property Get ArticuloProveedor() As String
    ArticuloProveedor = strArticuloProveedor
End Property

'Propiedad Codigo de Tipo de Comprobante de Proveedor
Public Property Let CodTipoComprobante(ByVal Value As String)
    strCodTipoComprobante = Value
End Property

Public Property Get CodTipoComprobante() As String
    CodTipoComprobante = strCodTipoComprobante
End Property

'Propiedad Proveedor Es Apto para Orden
Public Property Let EsAptoParaOrden(ByVal Value As Boolean)
    bolEsAptoParaOrden = Value
End Property

Public Property Get EsAptoParaOrden() As Boolean
    EsAptoParaOrden = bolEsAptoParaOrden
End Property

'Propiedad Codigo de Forma de Pago de Proveedor
Public Property Let CodigoFormaPago(ByVal Value As String)
    strCodFormaPago = Value
End Property

Public Property Get CodigoFormaPago() As String
    CodigoFormaPago = strCodFormaPago
End Property

'Propiedad Observacion de Proveedor
Public Property Let Observacion(ByVal Value As String)
    strObservacion = Value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property




'Propiedad Fecha de Registro de Proveedor
Public Property Let FechaReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Proveedor
Public Property Let UsuarioReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Proveedor
Public Property Let FechaMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Proveedor
Public Property Let UsuarioMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property


'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strOrigenProveedor = vbNullString
    strClaseProveedor = vbNullString
    strCodTipoDocumento = vbNullString
    strNroDocumento = vbNullString
    strNombreProveedor = vbNullString
    strDireccionProveedor = vbNullString
    strCodDistrito = vbNullString
    
    strCodigo = vbNullString
    strCodigoExterno = vbNullString
    strNombreAbrev = vbNullString
    lngCodCategoria = 0
    
    strCodigoMoneda = vbNullString
    strCuentaContable = vbNullString
    strCodigoGasto = vbNullString
    strGrupoFlujo = vbNullString
    strGrupoResultado = vbNullString 'F2GRUPORESULTADO
    
    strTelefono = vbNullString
    strFax = vbNullString
    strEmail = vbNullString
    strCuentaAbono = vbNullString
    strContacto = vbNullString
    
    strArticuloProveedor = vbNullString
    strCodTipoComprobante = vbNullString
    bolEsAptoParaOrden = False
    strCodFormaPago = vbNullString
    strObservacion = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerProveedor() As Boolean
    On Error GoTo errObtenerProveedor
    
    Set rstProveedor = New ADODB.Recordset
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        strSQLProveedor = vbNullString
        strSQLProveedor = strSQLProveedor & "SELECT "
        strSQLProveedor = strSQLProveedor & "* "
        strSQLProveedor = strSQLProveedor & "FROM "
        strSQLProveedor = strSQLProveedor & "MAESTROS.EF2PROVEEDORES "
        strSQLProveedor = strSQLProveedor & "WHERE "
        strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
        
        If rstProveedor.State = 1 Then rstProveedor.Close
        
        rstProveedor.Open strSQLProveedor, cnBdCPlus, adOpenForwardOnly, adLockReadOnly

    Else
        strSQLProveedor = vbNullString
        strSQLProveedor = strSQLProveedor & "SELECT "
        strSQLProveedor = strSQLProveedor & "* "
        strSQLProveedor = strSQLProveedor & "FROM "
        strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
        strSQLProveedor = strSQLProveedor & "WHERE "
        strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
        
        If rstProveedor.State = 1 Then rstProveedor.Close
        
        rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not rstProveedor.EOF Then
        strOrigenProveedor = IIf(Trim(rstProveedor!F2TIPPROV & "") <> vbNullString, Trim(rstProveedor!F2TIPPROV & ""), "N")
        strClaseProveedor = IIf(Trim(rstProveedor!F2TIPO_PERSONA & "") <> vbNullString, Trim(rstProveedor!F2TIPO_PERSONA & ""), "J")
        strCodTipoDocumento = IIf(Trim(rstProveedor!TIPPROV & "") <> vbNullString, Trim(rstProveedor!TIPPROV & ""), "6")
        strNroDocumento = Trim(rstProveedor!F2NEWRUC & "")
        strNombreProveedor = Trim(rstProveedor!F2NOMPROV & "")
        strDireccionProveedor = Trim(rstProveedor!F2DIRPROV & "")
        strCodDistrito = Trim(rstProveedor!F7CODPOS & "")
        
        
        strCodigo = Trim(rstProveedor!F2CODPROV & "")
        strCodigoExterno = Trim(rstProveedor!F2CODPROVEXTERNO & "")
        strNombreAbrev = Trim(rstProveedor!F2NOMABREV & "")
        lngCodCategoria = Val(rstProveedor!INTCODCATEGORIA & "")
        
        
        strCodigoMoneda = Trim(rstProveedor!F2TIPMON & "")
        strCuentaContable = Trim(rstProveedor!F5CODCTA & "")
        strCodigoGasto = Trim(rstProveedor!F2CODGAS & "")
        strGrupoFlujo = Trim(rstProveedor!F2GRUPO & "")
        strGrupoResultado = Trim(rstProveedor!F2GRUPORESULTADO & "")
        
        
        strTelefono = Trim(rstProveedor!F2TELPROV & "")
        strFax = Trim(rstProveedor!F2FAXPROV & "")
        strEmail = Trim(rstProveedor!F2EMAIL & "")
        strCuentaAbono = Trim(rstProveedor!F2NUMCTA & "")
        strContacto = Trim(rstProveedor!F2CONTACTO & "")
        
        
        strArticuloProveedor = Trim(rstProveedor!F2ARTPROV & "")
        strCodTipoComprobante = Trim(rstProveedor!F2TIPDOC & "")
        bolEsAptoParaOrden = CBool(rstProveedor!F2ORDEN)
        strCodFormaPago = Trim(rstProveedor!F2FORPAG & "")
        strObservacion = Trim(rstProveedor!F2OBSPROV & "")
        
        strFecReg = Trim(rstProveedor!FECHA_CREACION & "")
        strUsuReg = Trim(rstProveedor!USUARIO_CREA & "")
        strFecMod = Trim(rstProveedor!FECHA_MODIFICACION & "")
        strUsuMod = Trim(rstProveedor!USUARIO_MOD & "")
        
        obtenerProveedor = True
    Else
        obtenerProveedor = False
    End If
    
    rstProveedor.Close
    
    Set rstProveedor = Nothing
    
    Exit Function
errObtenerProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case 3265 'ORDINAL NO ENCONTRADO
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: ObtenerProveedor"
    End Select
    
    obtenerProveedor = False
    
    Err.Clear
End Function

Public Sub obtenerConfigProveedor()
    On Error GoTo errObtenerConfigProveedor
    Set rstProveedor = New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "* "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProveedor.EOF Then
        strOrigenProveedor = IIf(Trim(rstProveedor!F2TIPPROV & "") <> vbNullString, Trim(rstProveedor!F2TIPPROV & ""), "N")
        strClaseProveedor = IIf(Trim(rstProveedor!F2TIPO_PERSONA & "") <> vbNullString, Trim(rstProveedor!F2TIPO_PERSONA & ""), "J")
        strCodTipoDocumento = IIf(Trim(rstProveedor!TIPPROV & "") <> vbNullString, Trim(rstProveedor!TIPPROV & ""), "6")
        strNroDocumento = Trim(rstProveedor!F2NEWRUC & "")
        strNombreProveedor = Trim(rstProveedor!F2NOMPROV & "")
        strDireccionProveedor = Trim(rstProveedor!F2DIRPROV & "")
        strCodDistrito = Trim(rstProveedor!F7CODPOS & "")
        
        
        strCodigo = Trim(rstProveedor!F2CODPROV & "")
        strCodigoExterno = Trim(rstProveedor!F2CODPROVEXTERNO & "")
        strNombreAbrev = Trim(rstProveedor!F2NOMABREV & "")
        lngCodCategoria = Val(rstProveedor!INTCODCATEGORIA & "")
        
        
        strCodigoMoneda = Trim(rstProveedor!F2TIPMON & "")
        strCuentaContable = Trim(rstProveedor!F5CODCTA & "")
        strCodigoGasto = Trim(rstProveedor!F2CODGAS & "")
        strGrupoFlujo = Trim(rstProveedor!F2GRUPO & "")
        strGrupoResultado = Trim(rstProveedor!F2GRUPORESULTADO & "")
        
        
        strTelefono = Trim(rstProveedor!F2TELPROV & "")
        strFax = Trim(rstProveedor!F2FAXPROV & "")
        strEmail = Trim(rstProveedor!F2EMAIL & "")
        strCuentaAbono = Trim(rstProveedor!F2NUMCTA & "")
        strContacto = Trim(rstProveedor!F2CONTACTO & "")
        
        
        strArticuloProveedor = Trim(rstProveedor!F2ARTPROV & "")
        strCodTipoComprobante = Trim(rstProveedor!F2TIPDOC & "")
        bolEsAptoParaOrden = CBool(rstProveedor!F2ORDEN)
        strCodFormaPago = Trim(rstProveedor!F2FORPAG & "")
        strObservacion = Trim(rstProveedor!F2OBSPROV & "")
        
        strFecReg = Trim(rstProveedor!FECHA_CREACION & "")
        strUsuReg = Trim(rstProveedor!USUARIO_CREA & "")
        strFecMod = Trim(rstProveedor!FECHA_MODIFICACION & "")
        strUsuMod = Trim(rstProveedor!USUARIO_MOD & "")
    End If
    
    rstProveedor.Close
    
    Set rstProveedor = Nothing
    
    Exit Sub
errObtenerConfigProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case 3265 'ORDINAL NO ENCONTRADO
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsProveedor: ObtenerConfigProveedor"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstProveedor = New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES.* "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProveedor.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstProveedor.Close
    
    strSQLProveedor = vbNullString
    
    Set rstProveedor = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function verificarExistenciaPorNroDocumento() As Boolean
    On Error GoTo errVerificarExistenciaPorNroDocumento
    
    Set rstProveedor = New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES.* "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "F2NEWRUC = '" & strNroDocumento & "'"
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProveedor.EOF Then
        verificarExistenciaPorNroDocumento = True
    Else
        verificarExistenciaPorNroDocumento = False
    End If
    
    rstProveedor.Close
    
    Set rstProveedor = Nothing
    
    Exit Function
errVerificarExistenciaPorNroDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: VerificarExistenciaPorNroDocumento"
    End Select
    
    verificarExistenciaPorNroDocumento = False
    
    Err.Clear
End Function

Public Function generarCodigoProveedor() As String
    On Error GoTo errGenerarCodigoProveedor
    
    Dim RsProveedor As New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "F2CODPROV "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "ORDER BY "
    strSQLProveedor = strSQLProveedor & "F2CODPROV DESC"
    
    If RsProveedor.State = 1 Then RsProveedor.Close
    
    RsProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not RsProveedor.EOF Then
        generarCodigoProveedor = Format(Val(RsProveedor!F2CODPROV & "") + 1, "0000")
    Else
        generarCodigoProveedor = "0001"
    End If
    
    RsProveedor.Close
    
    Set RsProveedor = Nothing
    
    Exit Function
errGenerarCodigoProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "NombreProveedor: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: GenerarCodigoProveedor"
    End Select
        
    generarCodigoProveedor = vbNullString
    
    Err.Clear
End Function

Public Function guardarProveedor() As Boolean
    On Error GoTo errGuardarProveedor
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoProveedor
        End If
        
        strSQLProveedor = vbNullString
        strSQLProveedor = strSQLProveedor & "INSERT INTO EF2PROVEEDORES("
        strSQLProveedor = strSQLProveedor & "F2TIPPROV, F2TIPO_PERSONA, TIPPROV, F2NEWRUC, F2NOMPROV, F2DIRPROV, F7CODPOS, "
        strSQLProveedor = strSQLProveedor & "F2CODPROV, F2CODPROVEXTERNO, INTCODCATEGORIA, F2TIPMON, F5CODCTA, "
        strSQLProveedor = strSQLProveedor & "F2CODGAS, F2GRUPO, "
        'strSQLProveedor = strSQLProveedor & "F2GRUPORESULTADO, "
        strSQLProveedor = strSQLProveedor & "F2TELPROV, F2FAXPROV, F2EMAIL, F2NUMCTA, F2CONTACTO, F2ARTPROV, "
        strSQLProveedor = strSQLProveedor & "F2TIPDOC, F2ORDEN, F2FORPAG, FECHA_CREACION, USUARIO_CREA"
        strSQLProveedor = strSQLProveedor & ") "
        
        strSQLProveedor = strSQLProveedor & "VALUES("
        strSQLProveedor = strSQLProveedor & "'" & strOrigenProveedor & "', '" & strClaseProveedor & "', '" & strCodTipoDocumento & "', "
        strSQLProveedor = strSQLProveedor & "'" & strNroDocumento & "', '" & strNombreProveedor & "', '" & strDireccionProveedor & "', "
        strSQLProveedor = strSQLProveedor & "'" & strCodDistrito & "', '" & strCodigo & "', " & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "NULL") & ", "
        strSQLProveedor = strSQLProveedor & lngCodCategoria & ", '" & strCodigoMoneda & "', '" & strCuentaContable & "', "
        strSQLProveedor = strSQLProveedor & "'" & strCodigoGasto & "', '" & strGrupoFlujo & "', "
        'strSQLProveedor = strSQLProveedor & IIf(strGrupoResultado <> vbNullString, "'" & strGrupoResultado & "'", "NULL") & ",
        strSQLProveedor = strSQLProveedor & "'" & strTelefono & "', "
        strSQLProveedor = strSQLProveedor & "'" & strFax & "', '" & strEmail & "', '" & strCuentaAbono & "', '" & strContacto & "', "
        strSQLProveedor = strSQLProveedor & "'" & strArticuloProveedor & "', '" & strCodTipoComprobante & "', "
        strSQLProveedor = strSQLProveedor & IIf(bolEsAptoParaOrden, "TRUE", "FALSE") & ", "
        strSQLProveedor = strSQLProveedor & "'" & strCodFormaPago & "', CVDATE('" & strFecReg & "'), "
        strSQLProveedor = strSQLProveedor & "'" & strUsuReg & "'"
        strSQLProveedor = strSQLProveedor & ")"
    Else
        strSQLProveedor = vbNullString
        strSQLProveedor = strSQLProveedor & "UPDATE "
        strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
        
        strSQLProveedor = strSQLProveedor & "SET "
        strSQLProveedor = strSQLProveedor & "F2TIPPROV = '" & strOrigenProveedor & "', "
        strSQLProveedor = strSQLProveedor & "F2TIPO_PERSONA = '" & strClaseProveedor & "', "
        strSQLProveedor = strSQLProveedor & "TIPPROV = '" & strCodTipoDocumento & "', "
        strSQLProveedor = strSQLProveedor & "F2NEWRUC = '" & strNroDocumento & "', "
        strSQLProveedor = strSQLProveedor & "F2NOMPROV = '" & strNombreProveedor & "', "
        strSQLProveedor = strSQLProveedor & "F2DIRPROV = '" & strDireccionProveedor & "', "
        strSQLProveedor = strSQLProveedor & "F7CODPOS = '" & strCodDistrito & "', "
        strSQLProveedor = strSQLProveedor & "F2CODPROVEXTERNO = " & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "NULL") & ", "
        strSQLProveedor = strSQLProveedor & "INTCODCATEGORIA = '" & lngCodCategoria & "', "
        strSQLProveedor = strSQLProveedor & "F2TIPMON = '" & strCodigoMoneda & "', "
        strSQLProveedor = strSQLProveedor & "F5CODCTA = '" & strCuentaContable & "', "
        strSQLProveedor = strSQLProveedor & "F2CODGAS = '" & strCodigoGasto & "', "
        strSQLProveedor = strSQLProveedor & "F2GRUPO = '" & strGrupoFlujo & "', "
        'strSQLProveedor = strSQLProveedor & "F2GRUPORESULTADO = '" & strGrupoResultado & "', "
        strSQLProveedor = strSQLProveedor & "F2TELPROV = '" & strTelefono & "', "
        strSQLProveedor = strSQLProveedor & "F2FAXPROV = '" & strFax & "', "
        strSQLProveedor = strSQLProveedor & "F2EMAIL = '" & strEmail & "', "
        strSQLProveedor = strSQLProveedor & "F2NUMCTA = '" & strCuentaAbono & "', "
        strSQLProveedor = strSQLProveedor & "F2CONTACTO = '" & strContacto & "', "
        strSQLProveedor = strSQLProveedor & "F2ARTPROV = '" & strArticuloProveedor & "', "
        strSQLProveedor = strSQLProveedor & "F2TIPDOC = '" & strCodTipoComprobante & "', "
        strSQLProveedor = strSQLProveedor & "F2ORDEN = " & IIf(bolEsAptoParaOrden, "TRUE", "FALSE") & ", "
        strSQLProveedor = strSQLProveedor & "F2FORPAG = '" & strCodFormaPago & "' "
        
        strSQLProveedor = strSQLProveedor & "WHERE "
        strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLProveedor
    
    strSQLSelectAlter = strSQLProveedor
    
    strSQLProveedor = vbNullString
    
    guardarProveedor = True
    
    Exit Function
errGuardarProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: GuardarProveedor"
    End Select
    
    guardarProveedor = False
    
    Err.Clear
End Function

Public Function eliminarProveedor() As Boolean
    On Error GoTo errEliminarProveedor
    
    If Val(validarUsoRegistro(cnn_dbbancos, "PROVEEDOR", "PAG_DCTO", strCodigo, "T")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "F2CODPROV", "IF4VALES", strCodigo, "T", "AND F1TIPPRV = 'P'")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "F2CODPRV", "EF2PROD_PROV", strNroDocumento, "T")) > 0 Or _
        Val(validarUsoRegistro(cnn_dbbancos, "F4CODPRV", "REGISDOC", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarProveedor = False
        
        Exit Function
    End If
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "DELETE FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "F2CODPROV = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLProveedor
    
    eliminarProveedor = True
    
    strSQLSelectAlter = strSQLProveedor
    
    strSQLProveedor = vbNullString
    
    Exit Function
errEliminarProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: EliminarProveedor"
    End Select
    
    eliminarProveedor = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Proveedores en Grilla (QuamtumGrid)
Public Sub vistaProveedor(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaProveedor
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "P.F2CODPROV, "
    strSQLProveedor = strSQLProveedor & "TD.ABREVIATURA, "
    strSQLProveedor = strSQLProveedor & "P.F2NEWRUC, "
    strSQLProveedor = strSQLProveedor & "P.F2NOMPROV, "
    strSQLProveedor = strSQLProveedor & "P.F2DIRPROV, "
    strSQLProveedor = strSQLProveedor & "IIF(P.F2TIPPROV = 'E', 'EXTRANJERO', 'NACIONAL') AS ORIGEN, "
    strSQLProveedor = strSQLProveedor & "IIF(P.F2TIPO_PERSONA = 'N', 'NATURAL', 'JURIDICA') AS PERSONA, "
    strSQLProveedor = strSQLProveedor & "P.F2TELPROV, "
    strSQLProveedor = strSQLProveedor & "LCASE(P.F2EMAIL) AS EMAIL, "
    strSQLProveedor = strSQLProveedor & "P.F2CONTACTO, "
    strSQLProveedor = strSQLProveedor & "P.F2ORDEN "
    
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES AS P "
    strSQLProveedor = strSQLProveedor & "LEFT JOIN EF2TIPODOCID AS TD ON TD.CODIGO = P.TIPPROV "
    
    strSQLProveedor = strSQLProveedor & "ORDER BY "
    strSQLProveedor = strSQLProveedor & "P.F2NOMPROV"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo Interno de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODPROV"
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
                .Width = 70
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
            
            'Columna Nombre de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Nombres o Razón Social"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2NOMPROV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNombreProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = True
            End With
            
            'Columna Direccion de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Dirección Fiscal"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DIRPROV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDireccion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
                .Visible = True
            End With
            
            'Columna Origen de Proveedor
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
                .FieldName = "F2TELPROV"
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
            
            'Columna Orden
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Req. Orden"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F2ORDEN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLProveedor
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2NEWRUC"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLProveedor = vbNullString
    
    Exit Sub
errVistaProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: VistaProveedor"
    End Select
    
    Err.Clear
End Sub

Public Sub vistaProveedorSQL(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaProveedor
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "P.F2CODPROV, "
    strSQLProveedor = strSQLProveedor & "TD.ABREVIATURA, "
    strSQLProveedor = strSQLProveedor & "P.F2NEWRUC, "
    strSQLProveedor = strSQLProveedor & "P.F2NOMPROV, "
    strSQLProveedor = strSQLProveedor & "P.F2DIRPROV, "
    strSQLProveedor = strSQLProveedor & "IIF(P.F2TIPPROV = 'E', 'EXTRANJERO', 'NACIONAL') AS ORIGEN, "
    strSQLProveedor = strSQLProveedor & "IIF(P.F2TIPO_PERSONA = 'N', 'NATURAL', 'JURIDICA') AS PERSONA, "
    strSQLProveedor = strSQLProveedor & "P.F2TELPROV, "
    strSQLProveedor = strSQLProveedor & "UPPER(P.F2EMAIL) AS EMAIL, "
    strSQLProveedor = strSQLProveedor & "P.F2CONTACTO, "
    strSQLProveedor = strSQLProveedor & "P.F2ORDEN "
    
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "MAESTROS.EF2PROVEEDORES AS P "
    strSQLProveedor = strSQLProveedor & "LEFT JOIN MAESTROS.EF2TIPODOCID AS TD ON TD.CODIGO = P.TIPPROV "
    
    strSQLProveedor = strSQLProveedor & "ORDER BY "
    strSQLProveedor = strSQLProveedor & "P.F2NOMPROV"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo Interno de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODPROV"
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
                .Width = 70
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
            
            'Columna Nombre de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Nombres o Razón Social"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2NOMPROV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNombreProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = True
            End With
            
            'Columna Direccion de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Dirección Fiscal"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DIRPROV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDireccion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
                .Visible = True
            End With
            
            'Columna Origen de Proveedor
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
                .FieldName = "F2TELPROV"
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
            
            'Columna Orden
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Req. Orden"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F2ORDEN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus.ConnectionString
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLProveedor
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2NEWRUC"
            
        End With
    End If
    
    strSQLProveedor = vbNullString
    
    Exit Sub
errVistaProveedor:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: VistaProveedor"
    End Select
    
    Err.Clear
End Sub

'Vista de Detalle de Proveedores Contacto en Grilla (QuamtumGrid)
Public Sub vistaProveedorContacto(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaProveedorContacto
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "C.CODIGO, "
    strSQLProveedor = strSQLProveedor & "TRIM(C.APEPAT & ' ' & C.APEMAT) & ', ' & TRIM(NOMBRE) AS NOMBRES, "
    strSQLProveedor = strSQLProveedor & "C.TELEFONO1, "
    strSQLProveedor = strSQLProveedor & "C.ANEXO, "
    strSQLProveedor = strSQLProveedor & "C.TELEFONO2, "
    strSQLProveedor = strSQLProveedor & "C.CELULAR, "
    strSQLProveedor = strSQLProveedor & "C.EMAIL, "
    strSQLProveedor = strSQLProveedor & "(C.OBSERVACION & '') AS OBSERVA "
    
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "CONTACTOS AS C "
    
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "C.CODCLI = '" & strCodigo & "'"
    
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
            .Dataset.ADODataset.CommandText = strSQLProveedor
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODIGO"
        End With
    End If
    
    strSQLProveedor = vbNullString
    
    Exit Sub
errVistaProveedorContacto:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsProveedor: VistaProveedorContacto"
    End Select
    
    Err.Clear
End Sub

Public Sub listarAnnoProveedor(ByVal combo As Object)
    On Error GoTo errListarAnnoProveedor
    
    Set rstProveedor = New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "YEAR(FECHA) AS ANNO "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "GROUP BY "
    strSQLProveedor = strSQLProveedor & "YEAR(FECHA)"
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    combo.Clear
    combo.AddItem "(*) - Todos"
    
    If Not rstProveedor.EOF Then
        
        rstProveedor.MoveFirst
        
        Do While Not rstProveedor.EOF
            combo.AddItem Trim(rstProveedor!Anno & "")
            
            rstProveedor.MoveNext
        Loop
    End If
    
    rstProveedor.Close
    
    strSQLProveedor = vbNullString
    
    Set rstProveedor = Nothing
    
    Exit Sub
errListarAnnoProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "NombreProveedor: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProveedor: ListarAnnoProveedor"
    
    Err.Clear
End Sub

Public Sub listarTipoProveedor(ByVal comboList As Object)
    On Error GoTo errListarTipoProveedor
    
    Set rstProveedor = New ADODB.Recordset
    
    strSQLProveedor = vbNullString
    strSQLProveedor = strSQLProveedor & "SELECT "
    strSQLProveedor = strSQLProveedor & "F2NEWRUC, "
    strSQLProveedor = strSQLProveedor & "F2DESDOC "
    strSQLProveedor = strSQLProveedor & "FROM "
    strSQLProveedor = strSQLProveedor & "EF2PROVEEDORES "
    strSQLProveedor = strSQLProveedor & "WHERE "
    strSQLProveedor = strSQLProveedor & "F2TIPO IN ('C', 'A') "
    strSQLProveedor = strSQLProveedor & "ORDER BY "
    strSQLProveedor = strSQLProveedor & "F2DESDOC"
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open strSQLProveedor, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstProveedor.EOF Then
        
        rstProveedor.MoveFirst
        
        Do While Not rstProveedor.EOF
            comboList.AddItem Trim(rstProveedor!F2DESDOC & "") & Space(100) & Trim(rstProveedor!F2NEWRUC & "")
            
            rstProveedor.MoveNext
        Loop
    End If
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    strSQLProveedor = vbNullString
    
    Set rstProveedor = Nothing
    
    Exit Sub
errListarTipoProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "NombreProveedor: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProveedor: ListarTipoProveedor"
    
    Err.Clear
End Sub

