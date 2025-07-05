VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strTipoDocumento            As String 'cs_documento
Private strCodigo                   As String 'cod_solicitud
Private StrFecha                    As String 'cs_fecha
Private strCodArea                  As String 'cs_codarea
Private strCodProyecto              As String 'cs_codproy
Private strCodPresupuesto           As String 'cs_codpresupuesto

Private strCodSolicitante           As String 'cs_codsolicitante
Private strCodPreparadoPor1         As String 'cs_codpreparadox
Private strCodPreparadoPor2         As String 'cs_codpreparadoy
Private strCodAprobadoPor1          As String 'cs_aprobadox
Private strCodAprobadoPor2          As String 'cs_aprobadoy

Private strLugarEntrega             As String 'cs_lugEntr
Private strObservaciones            As String 'cs_observaciones
Private strNombreReferencial        As String 'cs_nomref

Private strEstado1                  As String 'cs_estado
Private strEstado2                  As String 'cs_estado2
Private strAsunto                   As String  'cs_asunto
Private strEmpresa                  As String 'empresa

Private strCodCosto                 As String 'cs_codcosto
Private strDesCosto                 As String 'cs_descosto

Private strPrioridad                As String 'cs_prioridad
Private strFechaEntrega             As String 'cs_fentrega
Private strUsuario                  As String 'cs_usuario
Private strCargo                    As String 'cs_cargo

Private dblTotal                    As String 'cs_total

Private strCodMoneda                As String 'cs_moneda
Private dblTCambio                  As String 'f4tipcam

Private bolAnulado                  As Boolean 'anulado

Private strNumeroPorUsuario         As String 'numxusuario
Private strNumeroPorObra            As String 'numxobra

Private strProveedorSugerido        As String 'cs_proveedor
Private strUltimaOrden              As String 'cs_orden
Private bolCerrado                  As Boolean 'f4cerrado
Private strNumeroOrden              As String 'numorden
Private bolVBJefe                   As Boolean 'vbjefecc
Private strVBFecha                  As String 'vbfecha
Private strVBUsuario                As String 'vbuser


Private strFecReg                   As String 'fecha_graba
Private strUsuReg                   As String
Private strFecMod                   As String 'fecha_modifica
Private strUsuMod                   As String

'Atributos de Detalle de Solicitud
Private dblCantidad As Double 'DS_CANTIDAD
Private strCodProducto As String 'COD_PRODUCTO
Private strCodUniMed As String 'DS_UNIDMED
Private strDescripcion As String 'DS_DESCRIPCION
Private dblPrecio As Double 'PRECIO
Private dblCantidad2 As Double 'CANDIS
Private bolAfecto As Boolean 'AFECTO y F5AFECTO
Private dblSubTotal As Double 'CS_SUBTOT
Private dblItem As Double 'ITEM
Private strMarca As String 'F5MARCA
Private strCodMarca As String 'F5CODMARCA
Private strCodFabricante As String 'F5CODFAB


Private dblPrecioConIGV As Double 'F5CONIGV
Private dblPrecioSinIGV As Double 'F5SINIGV

Private strCodPartida As String 'PARTIDA
Private strDesPartida As String 'DES_PARTIDA

Private strNombreProveedor As String 'PROVEEDOR



'Atributos Adicionales
Private dblCompromisoEA As Double
Private dblCompromisoPL As Double
Private dblSaldoPorAtender As Double


Private strSQLSelectAlter           As String

Private rstSolicitud                As ADODB.Recordset
Private strSQLSolicitud             As String


Public Property Let TipoDocumento(ByVal Value As String)
    strTipoDocumento = Value
End Property

Public Property Get TipoDocumento() As String
    TipoDocumento = strTipoDocumento
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Fecha(ByVal Value As String)
    StrFecha = Value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property

Public Property Let CodigoArea(ByVal Value As String)
    strCodArea = Value
End Property

Public Property Get CodigoArea() As String
    CodigoArea = strCodArea
End Property

Public Property Let CodigoProyecto(ByVal Value As String)
    strCodProyecto = Value
End Property

Public Property Get CodigoProyecto() As String
    CodigoProyecto = strCodProyecto
End Property

Public Property Let CodigoPresupuesto(ByVal Value As String)
    strCodPresupuesto = Value
End Property

Public Property Get CodigoPresupuesto() As String
    CodigoPresupuesto = strCodPresupuesto
End Property

Public Property Let CodigoSolicitante(ByVal Value As String)
    strCodSolicitante = Value
End Property

Public Property Get CodigoSolicitante() As String
    CodigoSolicitante = strCodSolicitante
End Property

Public Property Let CodigoPreparadoPor1(ByVal Value As String)
    strCodPreparadoPor1 = Value
End Property

Public Property Get CodigoPreparadoPor1() As String
    CodigoPreparadoPor1 = strCodPreparadoPor1
End Property

Public Property Let CodigoPreparadoPor2(ByVal Value As String)
    strCodPreparadoPor2 = Value
End Property

Public Property Get CodigoPreparadoPor2() As String
    CodigoPreparadoPor2 = strCodPreparadoPor2
End Property

Public Property Let CodigoAprobadoPor1(ByVal Value As String)
    strCodAprobadoPor1 = Value
End Property

Public Property Get CodigoAprobadoPor1() As String
    CodigoAprobadoPor1 = strCodAprobadoPor1
End Property

Public Property Let CodigoAprobadoPor2(ByVal Value As String)
    strCodAprobadoPor2 = Value
End Property

Public Property Get CodigoAprobadoPor2() As String
    CodigoAprobadoPor2 = strCodAprobadoPor2
End Property

Public Property Let LugarEntrega(ByVal Value As String)
    strLugarEntrega = Value
End Property

Public Property Get LugarEntrega() As String
    LugarEntrega = strLugarEntrega
End Property

Public Property Let observaciones(ByVal Value As String)
    strObservaciones = Value
End Property

Public Property Get observaciones() As String
    observaciones = strObservaciones
End Property
'NombreReferencial
Public Property Let NombreReferencial(ByVal Value As String)
    strNombreReferencial = Value
End Property

Public Property Get NombreReferencial() As String
    NombreReferencial = strNombreReferencial
End Property
'Estado1
Public Property Let Estado1(ByVal Value As String)
    strEstado1 = Value
End Property

Public Property Get Estado1() As String
    Estado1 = strEstado1
End Property
'Estado2
Public Property Let Estado2(ByVal Value As String)
    strEstado2 = Value
End Property

Public Property Get Estado2() As String
    Estado2 = strEstado2
End Property
'Asunto
Public Property Let asunto(ByVal Value As String)
    strAsunto = Value
End Property

Public Property Get asunto() As String
    asunto = strAsunto
End Property
'Empresa
Public Property Let Empresa(ByVal Value As String)
    strEmpresa = Value
End Property

Public Property Get Empresa() As String
    Empresa = strEmpresa
End Property
'CodCosto
Public Property Let CodigoCosto(ByVal Value As String)
    strCodCosto = Value
End Property

Public Property Get CodigoCosto() As String
    CodigoCosto = strCodCosto
End Property
'DesCosto
Public Property Let DescripcionCosto(ByVal Value As String)
    strDesCosto = Value
End Property

Public Property Get DescripcionCosto() As String
    DescripcionCosto = strDesCosto
End Property
'Prioridad
Public Property Let Prioridad(ByVal Value As String)
    strPrioridad = Value
End Property

Public Property Get Prioridad() As String
    Prioridad = strPrioridad
End Property
'FechaEntrega
Public Property Let FechaEntrega(ByVal Value As String)
    strFechaEntrega = Value
End Property

Public Property Get FechaEntrega() As String
    FechaEntrega = strFechaEntrega
End Property
'Usuario
Public Property Let Usuario(ByVal Value As String)
    strUsuario = Value
End Property

Public Property Get Usuario() As String
    Usuario = strUsuario
End Property
'Cargo
Public Property Let Cargo(ByVal Value As String)
    strCargo = Value
End Property

Public Property Get Cargo() As String
    Cargo = strCargo
End Property
'Total
Public Property Let TOTAL(ByVal Value As Double)
    dblTotal = Value
End Property

Public Property Get TOTAL() As Double
    TOTAL = dblTotal
End Property
'CodMoneda
Public Property Let CodigoMoneda(ByVal Value As String)
    strCodMoneda = Value
End Property

Public Property Get CodigoMoneda() As String
    CodigoMoneda = strCodMoneda
End Property
'TCambio
Public Property Let TCambio(ByVal Value As Double)
    dblTCambio = Value
End Property

Public Property Get TCambio() As Double
    TCambio = dblTCambio
End Property
'Anulado
Public Property Let Anulado(ByVal Value As Boolean)
    bolAnulado = Value
End Property

Public Property Get Anulado() As Boolean
    Anulado = bolAnulado
End Property
'NumeroPorUsuario
Public Property Let NumeroPorUsuario(ByVal Value As String)
    strNumeroPorUsuario = Value
End Property

Public Property Get NumeroPorUsuario() As String
    NumeroPorUsuario = strNumeroPorUsuario
End Property
'NumeroPorObra
Public Property Let NumeroPorObra(ByVal Value As String)
    strNumeroPorObra = Value
End Property

Public Property Get NumeroPorObra() As String
    NumeroPorObra = strNumeroPorObra
End Property
'ProveedorSugerido
Public Property Let ProveedorSugerido(ByVal Value As String)
    strProveedorSugerido = Value
End Property

Public Property Get ProveedorSugerido() As String
    ProveedorSugerido = strProveedorSugerido
End Property
'UltimaOrden
Public Property Let UltimaOrden(ByVal Value As String)
    strUltimaOrden = Value
End Property

Public Property Get UltimaOrden() As String
    UltimaOrden = strUltimaOrden
End Property
'Cerrado
Public Property Let Cerrado(ByVal Value As Boolean)
    bolCerrado = Value
End Property

Public Property Get Cerrado() As Boolean
    Cerrado = bolCerrado
End Property
'NumeroOrden
Public Property Let NumeroOrden(ByVal Value As String)
    strNumeroOrden = Value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property
'VBJefe
Public Property Let VBJefe(ByVal Value As Boolean)
    bolVBJefe = Value
End Property

Public Property Get VBJefe() As Boolean
    VBJefe = bolVBJefe
End Property
'VBFecha
Public Property Let VBFecha(ByVal Value As String)
    strVBFecha = Value
End Property

Public Property Get VBFecha() As String
    VBFecha = strVBFecha
End Property
'VBUsuario
Public Property Let VBUsuario(ByVal Value As String)
    strVBUsuario = Value
End Property

Public Property Get VBUsuario() As String
    VBUsuario = strVBUsuario
End Property



'Propiedad Fecha de Registro de Color de Bien
Public Property Let FechaReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Solicitud
Public Property Let FechaMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Solicitud
Public Property Let UsuarioMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property




'Cantidad
Public Property Let Cantidad(ByVal Value As Double)
    dblCantidad = Value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property
'CodProducto
Public Property Let CodProducto(ByVal Value As String)
    strCodProducto = Value
End Property

Public Property Get CodProducto() As String
    CodProducto = strCodProducto
End Property
'CodUniMed
Public Property Let CodUniMed(ByVal Value As String)
    strCodUniMed = Value
End Property

Public Property Get CodUniMed() As String
    CodUniMed = strCodUniMed
End Property
'Descripcion
Public Property Let Descripcion(ByVal Value As String)
    strDescripcion = Value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property
'Precio
Public Property Let Precio(ByVal Value As Double)
    dblPrecio = Value
End Property

Public Property Get Precio() As Double
    Precio = dblPrecio
End Property
'Cantidad2
Public Property Let Cantidad2(ByVal Value As Double)
    dblCantidad2 = Value
End Property

Public Property Get Cantidad2() As Double
    Cantidad2 = dblCantidad2
End Property
'Afecto
Public Property Let Afecto(ByVal Value As Boolean)
    bolAfecto = Value
End Property

Public Property Get Afecto() As Boolean
    Afecto = bolAfecto
End Property
'SubTotal
Public Property Let SUBTOTAL(ByVal Value As Double)
    dblSubTotal = Value
End Property

Public Property Get SUBTOTAL() As Double
    SUBTOTAL = dblSubTotal
End Property
'Item
Public Property Let ITEM(ByVal Value As Double)
    dblItem = Value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property
'Marca
Public Property Let Marca(ByVal Value As String)
    strMarca = Value
End Property

Public Property Get Marca() As String
    Marca = strMarca
End Property
'CodMarca
Public Property Let CodMarca(ByVal Value As String)
    strCodMarca = Value
End Property

Public Property Get CodMarca() As String
    CodMarca = strCodMarca
End Property
'CodFabricante
Public Property Let CodFabricante(ByVal Value As String)
    strCodFabricante = Value
End Property

Public Property Get CodFabricante() As String
    CodFabricante = strCodFabricante
End Property
'PrecioConIGV
Public Property Let PrecioConIGV(ByVal Value As Double)
    dblPrecioConIGV = Value
End Property

Public Property Get PrecioConIGV() As Double
    PrecioConIGV = dblPrecioConIGV
End Property
'PrecioSinIGV
Public Property Let PrecioSinIGV(ByVal Value As Double)
    dblPrecioSinIGV = Value
End Property

Public Property Get PrecioSinIGV() As Double
    PrecioSinIGV = dblPrecioSinIGV
End Property
'CodPartida
Public Property Let CodPartida(ByVal Value As String)
    strCodPartida = Value
End Property

Public Property Get CodPartida() As String
    CodPartida = strCodPartida
End Property
'DesPartida
Public Property Let DesPartida(ByVal Value As String)
    strDesPartida = Value
End Property

Public Property Get DesPartida() As String
    DesPartida = strDesPartida
End Property
'NombreProveedor
Public Property Let NombreProveedor(ByVal Value As String)
    strNombreProveedor = Value
End Property

Public Property Get NombreProveedor() As String
    NombreProveedor = strNombreProveedor
End Property



'Atributos Adicionales
'Private dblCompromisoEA As Double
Public Property Let CompromisoEA(ByVal Value As Double)
    dblCompromisoEA = Value
End Property

Public Property Get CompromisoEA() As Double
    CompromisoEA = dblCompromisoEA
End Property
'Private dblCompromisoPL As Double
Public Property Let CompromisoPL(ByVal Value As Double)
    dblCompromisoPL = Value
End Property

Public Property Get CompromisoPL() As Double
    CompromisoPL = dblCompromisoPL
End Property
'Private dblSaldoPorAtender As Double
Public Property Let SaldoPorAtender(ByVal Value As Double)
    dblSaldoPorAtender = Value
End Property

Public Property Get SaldoPorAtender() As Double
    SaldoPorAtender = dblSaldoPorAtender
End Property






Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strTipoDocumento = vbNullString 'cs_documento
    strCodigo = vbNullString 'cod_solicitud
    StrFecha = vbNullString 'cs_fecha
    strCodArea = vbNullString 'cs_codarea
    strCodProyecto = vbNullString 'cs_codproy
    strCodPresupuesto = vbNullString 'cs_codpresupuesto
    
    strCodSolicitante = vbNullString 'cs_codsolicitante
    strCodPreparadoPor1 = vbNullString 'cs_codpreparadox
    strCodPreparadoPor2 = vbNullString 'cs_codpreparadoy
    strCodAprobadoPor1 = vbNullString 'cs_aprobadox
    strCodAprobadoPor2 = vbNullString 'cs_aprobadoy
    
    strLugarEntrega = vbNullString 'cs_lugEntr
    strObservaciones = vbNullString 'cs_observaciones
    strNombreReferencial = vbNullString 'cs_nomref
    
    strEstado1 = vbNullString 'cs_estado
    strEstado2 = vbNullString 'cs_estado2
    strAsunto = vbNullString 'cs_asunto
    strEmpresa = vbNullString 'empresa
    
    strCodCosto = vbNullString 'cs_codcosto
    strDesCosto = vbNullString 'cs_descosto
    
    strPrioridad = vbNullString 'cs_prioridad
    strFechaEntrega = vbNullString 'cs_fentrega
    strUsuario = vbNullString 'cs_usuario
    strCargo = vbNullString 'cs_cargo
    
    dblTotal = 0 'cs_total
    
    strCodMoneda = vbNullString 'cs_moneda
    dblTCambio = 0 'f4tipcam
    
    bolAnulado = False 'anulado
    
    strNumeroPorUsuario = vbNullString 'numxusuario
    strNumeroPorObra = vbNullString 'numxobra
    
    strProveedorSugerido = vbNullString 'cs_proveedor
    strUltimaOrden = vbNullString 'cs_orden
    bolCerrado = False 'f4cerrado
    strNumeroOrden = vbNullString 'numorden
    bolVBJefe = False 'vbjefecc
    strVBFecha = vbNullString 'vbfecha
    strVBUsuario = vbNullString 'vbuser
    
    
    strFecReg = vbNullString 'fecha_graba
    strUsuReg = vbNullString
    strFecMod = vbNullString 'fecha_modifica
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesDetalle()
    dblCantidad = 0 'DS_CANTIDAD
    strCodProducto = vbNullString 'COD_PRODUCTO
    strCodUniMed = vbNullString 'DS_UNIDMED
    strDescripcion = vbNullString 'DS_DESCRIPCION
    dblPrecio = 0 'PRECIO
    dblCantidad2 = 0 'CANDIS
    bolAfecto = False 'AFECTO y F5AFECTO
    dblSubTotal = 0 'CS_SUBTOT
    dblItem = 0 'ITEM
    strMarca = vbNullString 'F5MARCA
    strCodMarca = vbNullString 'F5CODMARCA
    strCodFabricante = vbNullString 'F5CODFAB
    
    
    dblPrecioConIGV = 0 'F5CONIGV
    dblPrecioSinIGV = 0 'F5SINIGV
    
    strObservaciones = vbNullString
    
    strCodPartida = vbNullString 'PARTIDA
    strDesPartida = vbNullString 'DES_PARTIDA
    
    strNombreProveedor = vbNullString 'PROVEEDOR
End Sub

Public Sub inicializarEntidadesAdicionales()
    dblCompromisoEA = 0
    dblCompromisoPL = 0
    dblSaldoPorAtender = 0
End Sub

Public Function obtenerSolicitud() As Boolean
    On Error GoTo errObtenerSolicitud
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "* "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' AND "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD = '" & strCodigo & "'"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        strTipoDocumento = Trim(rstSolicitud!CS_DOCUMENTO & "") 'cs_documento
        strCodigo = Trim(rstSolicitud!COD_SOLICITUD & "") 'cod_solicitud
        StrFecha = Trim(rstSolicitud!CS_FECHA & "") 'cs_fecha
        strCodArea = Trim(rstSolicitud!CS_CODAREA & "") 'cs_codarea
        strCodProyecto = Trim(rstSolicitud!CS_CODPROY & "") 'cs_codproy
        strCodPresupuesto = Trim(rstSolicitud!CS_CODPRESUPUESTO & "") 'cs_codpresupuesto
        
        strCodSolicitante = Trim(rstSolicitud!CS_CODSOLICITANTE & "") 'cs_codsolicitante
        strCodPreparadoPor1 = Trim(rstSolicitud!CS_CODPREPARADOX & "") 'cs_codpreparadox
        strCodPreparadoPor2 = Trim(rstSolicitud!CS_CODPREPARADOY & "") 'cs_codpreparadoy
        strCodAprobadoPor1 = Trim(rstSolicitud!CS_APROBADOX & "") 'cs_aprobadox
        strCodAprobadoPor2 = Trim(rstSolicitud!CS_APROBADOY & "") 'cs_aprobadoy
        
        strLugarEntrega = Trim(rstSolicitud!CS_LUGENTR & "") 'cs_lugEntr
        strObservaciones = Trim(rstSolicitud!CS_OBSERVACIONES & "") 'cs_observaciones
        strNombreReferencial = Trim(rstSolicitud!CS_NOMREF & "") 'cs_nomref
        
        strEstado1 = Trim(rstSolicitud!CS_ESTADO & "") 'cs_estado
        strEstado2 = Trim(rstSolicitud!CS_ESTADO2 & "") 'cs_estado2
        strAsunto = Trim(rstSolicitud!CS_ASUNTO & "") 'cs_asunto
        strEmpresa = Trim(rstSolicitud!Empresa & "") 'empresa
        
        strCodCosto = Trim(rstSolicitud!CS_CODCOSTO & "") 'cs_codcosto
        strDesCosto = Trim(rstSolicitud!CS_DESCOSTO & "") 'cs_descosto
        
        strPrioridad = Trim(rstSolicitud!CS_PRIORIDAD & "") 'cs_prioridad
        strFechaEntrega = Trim(rstSolicitud!cs_fentrega & "") 'cs_fentrega
        strUsuario = Trim(rstSolicitud!CS_USUARIO & "") 'cs_usuario
        strCargo = Trim(rstSolicitud!CS_CARGO & "") 'cs_cargo
        
        dblTotal = Val(rstSolicitud!CS_TOTAL & "") 'cs_total
        
        strCodMoneda = Trim(rstSolicitud!CS_MONEDA & "") 'cs_moneda
        dblTCambio = Val(rstSolicitud!F4TIPCAM & "") 'f4tipcam
        
        bolAnulado = IIf(Trim(rstSolicitud!Anulado & "") = "S", True, False) 'anulado
        
        strNumeroPorUsuario = Trim(rstSolicitud!NUMXUSUARIO & "") 'numxusuario
        strNumeroPorObra = Trim(rstSolicitud!NUMXOBRA & "") 'numxobra
        
        strProveedorSugerido = Trim(rstSolicitud!CS_PROVEEDOR & "") 'cs_proveedor
        strUltimaOrden = Trim(rstSolicitud!CS_ORDEN & "") 'cs_orden
        bolCerrado = IIf(Trim(rstSolicitud!F4CERRADO & "") = "S", True, False) 'f4cerrado
        strNumeroOrden = Trim(rstSolicitud!NUMORDEN & "") 'numorden
        bolVBJefe = CBool(rstSolicitud!VBJEFECC) 'vbjefecc
        strVBFecha = Trim(rstSolicitud!VBFecha & "") 'vbfecha
        strVBUsuario = Trim(rstSolicitud!vbUser & "") 'vbuser
        
        
        strFecReg = Trim(rstSolicitud!FECHA_GRABA & "") 'fecha_graba
        strUsuReg = vbNullString
        strFecMod = Trim(rstSolicitud!Fecha_modifica & "") 'fecha_modifica
        strUsuMod = vbNullString
        
        obtenerSolicitud = True
    Else
        obtenerSolicitud = False
    End If
    
    rstSolicitud.Close
    
    Set rstSolicitud = Nothing
    
    Exit Function
errObtenerSolicitud:
    obtenerSolicitud = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: ObtenerSolicitud"
    
    Err.Clear
End Function

Public Sub obtenerConfigSolicitud()
    On Error GoTo errObtenerConfigSolicitud
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "* "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' AND "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD = '" & strCodigo & "'"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        strTipoDocumento = Trim(rstSolicitud!CS_DOCUMENTO & "") 'cs_documento
        strCodigo = Trim(rstSolicitud!COD_SOLICITUD & "") 'cod_solicitud
        StrFecha = Trim(rstSolicitud!CS_FECHA & "") 'cs_fecha
        strCodArea = Trim(rstSolicitud!CS_CODAREA & "") 'cs_codarea
        strCodProyecto = Trim(rstSolicitud!CS_CODPROY & "") 'cs_codproy
        strCodPresupuesto = Trim(rstSolicitud!CS_CODPRESUPUESTO & "") 'cs_codpresupuesto
        
        strCodSolicitante = Trim(rstSolicitud!CS_CODSOLICITANTE & "") 'cs_codsolicitante
        strCodPreparadoPor1 = Trim(rstSolicitud!CS_CODPREPARADOX & "") 'cs_codpreparadox
        strCodPreparadoPor2 = Trim(rstSolicitud!CS_CODPREPARADOY & "") 'cs_codpreparadoy
        strCodAprobadoPor1 = Trim(rstSolicitud!CS_APROBADOX & "") 'cs_aprobadox
        strCodAprobadoPor2 = Trim(rstSolicitud!CS_APROBADOY & "") 'cs_aprobadoy
        
        strLugarEntrega = Trim(rstSolicitud!CS_LUGENTR & "") 'cs_lugEntr
        strObservaciones = Trim(rstSolicitud!CS_OBSERVACIONES & "") 'cs_observaciones
        strNombreReferencial = Trim(rstSolicitud!CS_NOMREF & "") 'cs_nomref
        
        strEstado1 = Trim(rstSolicitud!CS_ESTADO & "") 'cs_estado
        strEstado2 = Trim(rstSolicitud!CS_ESTADO2 & "") 'cs_estado2
        strAsunto = Trim(rstSolicitud!CS_ASUNTO & "") 'cs_asunto
        strEmpresa = Trim(rstSolicitud!Empresa & "") 'empresa
        
        strCodCosto = Trim(rstSolicitud!CS_CODCOSTO & "") 'cs_codcosto
        strDesCosto = Trim(rstSolicitud!CS_DESCOSTO & "") 'cs_descosto
        
        strPrioridad = Trim(rstSolicitud!CS_PRIORIDAD & "") 'cs_prioridad
        strFechaEntrega = Trim(rstSolicitud!cs_fentrega & "") 'cs_fentrega
        strUsuario = Trim(rstSolicitud!CS_USUARIO & "") 'cs_usuario
        strCargo = Trim(rstSolicitud!CS_CARGO & "") 'cs_cargo
        
        dblTotal = Val(rstSolicitud!CS_TOTAL & "") 'cs_total
        
        strCodMoneda = Trim(rstSolicitud!CS_MONEDA & "") 'cs_moneda
        dblTCambio = Val(rstSolicitud!F4TIPCAM & "") 'f4tipcam
        
        bolAnulado = IIf(Trim(rstSolicitud!Anulado & "") = "S", True, False) 'anulado
        
        strNumeroPorUsuario = Trim(rstSolicitud!NUMXUSUARIO & "") 'numxusuario
        strNumeroPorObra = Trim(rstSolicitud!NUMXOBRA & "") 'numxobra
        
        strProveedorSugerido = Trim(rstSolicitud!CS_PROVEEDOR & "") 'cs_proveedor
        strUltimaOrden = Trim(rstSolicitud!CS_ORDEN & "") 'cs_orden
        bolCerrado = IIf(Trim(rstSolicitud!F4CERRADO & "") = "S", True, False) 'f4cerrado
        strNumeroOrden = Trim(rstSolicitud!NUMORDEN & "") 'numorden
        bolVBJefe = CBool(rstSolicitud!VBJEFECC) 'vbjefecc
        strVBFecha = Trim(rstSolicitud!VBFecha & "") 'vbfecha
        strVBUsuario = Trim(rstSolicitud!vbUser & "") 'vbuser
        
        
        strFecReg = Trim(rstSolicitud!FECHA_GRABA & "") 'fecha_graba
        strUsuReg = vbNullString
        strFecMod = Trim(rstSolicitud!Fecha_modifica & "") 'fecha_modifica
        strUsuMod = vbNullString
    End If
    
    rstSolicitud.Close
    
    Set rstSolicitud = Nothing
    
    Exit Sub
errObtenerConfigSolicitud:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: ObtenerSolicitud"
    
    Err.Clear
End Sub

Public Function generarCodigoSolicitud() As String
    On Error GoTo errGenerarCodigoSolicitud
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "TOP 1 COD_SOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' "
    strSQLSolicitud = strSQLSolicitud & "ORDER BY "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD DESC"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        generarCodigoSolicitud = Trim((Val(Trim(rstSolicitud!Codigo & "")) + 1) & "")
    Else
        generarCodigoSolicitud = "1"
    End If
    
    rstSolicitud.Close
    
    Set rstSolicitud = Nothing
    
    strSQLSolicitud = vbNullString
    
    Exit Function
errGenerarCodigoSolicitud:
    generarCodigoSolicitud = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: GenerarCodigoSolicitud"
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD.* "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' AND "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD = '" & strCodigo & "'"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstSolicitud.Close
    
    strSQLSolicitud = vbNullString
    
    Set rstSolicitud = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarSolicitud(ByVal bolRegistroImportado As Boolean) As Boolean
    On Error GoTo errGuardarSolicitud
    
    strSQLSolicitud = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoSolicitud
        End If
        
        strSQLSolicitud = strSQLSolicitud & "INSERT INTO PROCESOS.TB_CABSOLICITUD("
        strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO, COD_SOLICITUD, CS_FECHA, "
        strSQLSolicitud = strSQLSolicitud & "CS_CODAREA, CS_CODPROY, CS_CODPRESUPUESTO, "
        strSQLSolicitud = strSQLSolicitud & "CS_CODSOLICITANTE, CS_CODPREPARADOX, CS_CODPREPARADOY, "
        strSQLSolicitud = strSQLSolicitud & "CS_APROBADOX, CS_APROBADOY, CS_LUGENTR, "
        strSQLSolicitud = strSQLSolicitud & "CS_OBSERVACIONES, CS_NOMREF, CS_ESTADO, "
        strSQLSolicitud = strSQLSolicitud & "CS_ESTADO2, CS_ASUNTO, EMPRESA, "
        strSQLSolicitud = strSQLSolicitud & "CS_CODCOSTO, CS_DESCOSTO, CS_PRIORIDAD, "
        strSQLSolicitud = strSQLSolicitud & "CS_FENTREGA, CS_USUARIO, CS_CARGO, "
        strSQLSolicitud = strSQLSolicitud & "CS_TOTAL, CS_MONEDA, F4TIPCAM, "
        strSQLSolicitud = strSQLSolicitud & "ANULADO, NUMXUSUARIO, NUMXOBRA, "
        strSQLSolicitud = strSQLSolicitud & "CS_PROVEEDOR, CS_ORDEN, F4CERRADO, "
        strSQLSolicitud = strSQLSolicitud & "NUMORDEN, VBJEFECC, VBFECHA, "
        strSQLSolicitud = strSQLSolicitud & "VBUSER, FECHA_GRABA, FECHA_MODIFICA) "
        
        strSQLSolicitud = strSQLSolicitud & "VALUES("
        strSQLSolicitud = strSQLSolicitud & "'" & strTipoDocumento & "', '" & strCodigo & "', '" & StrFecha & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strCodArea & "', '" & strCodProyecto & "', '" & strCodPresupuesto & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strCodSolicitante & "', '" & strCodPreparadoPor1 & "', '" & strCodPreparadoPor2 & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strCodAprobadoPor1 & "', '" & strCodAprobadoPor2 & "', '" & strLugarEntrega & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strObservaciones & "', '" & strNombreReferencial & "', '" & strEstado1 & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strEstado2 & "', '" & strAsunto & "', '" & strEmpresa & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strCodCosto & "', '" & strDesCosto & "', '" & strPrioridad & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strFechaEntrega & "', '" & strUsuario & "', '" & strCargo & "', "
        strSQLSolicitud = strSQLSolicitud & dblTotal & ", '" & strCodMoneda & "', " & dblTCambio & ", "
        strSQLSolicitud = strSQLSolicitud & "'" & IIf(bolAnulado, "S", "N") & "', '" & strNumeroPorUsuario & "', '" & strNumeroPorObra & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strProveedorSugerido & "', '" & strUltimaOrden & "', '" & IIf(bolCerrado, "S", "N") & "', "
        strSQLSolicitud = strSQLSolicitud & "'" & strNumeroOrden & "', " & IIf(bolVBJefe, "1", "0") & ", "
        strSQLSolicitud = strSQLSolicitud & IIf(strVBFecha <> vbNullString, "'" & strVBFecha & "'", "NULL") & ", "
        strSQLSolicitud = strSQLSolicitud & IIf(strVBUsuario <> vbNullString, "'" & strVBUsuario & "'", "NULL") & ", "
        strSQLSolicitud = strSQLSolicitud & "CONVERT(DATETIME, '" & Format(strFecReg, "dd-mm-yy hh:mm:ss") & "', 5), "
        strSQLSolicitud = strSQLSolicitud & "CONVERT(DATETIME, '" & Format(strFecMod, "dd-mm-yy hh:mm:ss") & "', 5)"
        strSQLSolicitud = strSQLSolicitud & ")"
    Else
        strSQLSolicitud = strSQLSolicitud & "UPDATE "
        strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
        strSQLSolicitud = strSQLSolicitud & "SET "
        
        strSQLSolicitud = strSQLSolicitud & "CS_FECHA = '" & StrFecha & "', "
        
            If Not bolRegistroImportado Then
                strSQLSolicitud = strSQLSolicitud & "CS_ESTADO = '" & strEstado1 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_ESTADO2 = '" & strEstado2 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_PRIORIDAD = '" & strPrioridad & "', "
                
                strSQLSolicitud = strSQLSolicitud & "CS_CODPREPARADOX = '" & strCodPreparadoPor1 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_CODPREPARADOY = '" & strCodPreparadoPor2 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_APROBADOX = '" & strCodAprobadoPor1 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_APROBADOY = '" & strCodAprobadoPor2 & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_LUGENTR = '" & strLugarEntrega & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_OBSERVACIONES = '" & strObservaciones & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_NOMREF = '" & strNombreReferencial & "', "
                
                strSQLSolicitud = strSQLSolicitud & "CS_CODAREA = '" & strCodArea & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_CODPROY = '" & strCodProyecto & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_CODPRESUPUESTO = '" & strCodPresupuesto & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_CODSOLICITANTE = '" & strCodSolicitante & "', "
                
                
                strSQLSolicitud = strSQLSolicitud & "CS_ASUNTO = '" & strAsunto & "', "
                strSQLSolicitud = strSQLSolicitud & "EMPRESA = '" & strEmpresa & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_CODCOSTO = '" & strCodCosto & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_DESCOSTO = '" & strDesCosto & "', "
                
                strSQLSolicitud = strSQLSolicitud & "VBJEFECC = " & IIf(bolVBJefe, 1, 0) & ", "
                strSQLSolicitud = strSQLSolicitud & "VBFECHA = " & IIf(strVBFecha <> vbNullString, "'" & strVBFecha & "'", "NULL") & ", "
                strSQLSolicitud = strSQLSolicitud & "VBUSER = " & IIf(strVBUsuario <> vbNullString, "'" & strVBUsuario & "'", "NULL") & ", "
                
                strSQLSolicitud = strSQLSolicitud & "CS_CARGO = '" & strCargo & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_TOTAL = " & dblTotal & ", "
                strSQLSolicitud = strSQLSolicitud & "CS_MONEDA = '" & strCodMoneda & "', "
                strSQLSolicitud = strSQLSolicitud & "F4TIPCAM = " & dblTCambio & ", "
                
                strSQLSolicitud = strSQLSolicitud & "NUMXUSUARIO = '" & strNumeroPorUsuario & "', "
                strSQLSolicitud = strSQLSolicitud & "NUMXOBRA = '" & strNumeroPorObra & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_PROVEEDOR = '" & strProveedorSugerido & "', "
                strSQLSolicitud = strSQLSolicitud & "CS_ORDEN = '" & strUltimaOrden & "', "
                strSQLSolicitud = strSQLSolicitud & "NUMORDEN = '" & strNumeroOrden & "', "
            End If
                
        strSQLSolicitud = strSQLSolicitud & "CS_FENTREGA = '" & strFechaEntrega & "', "
        strSQLSolicitud = strSQLSolicitud & "CS_USUARIO = '" & strUsuario & "', "
        strSQLSolicitud = strSQLSolicitud & "ANULADO = '" & IIf(bolAnulado, "S", "N") & "', "
        strSQLSolicitud = strSQLSolicitud & "F4CERRADO = '" & IIf(bolCerrado, "S", "N") & "', "
        strSQLSolicitud = strSQLSolicitud & "FECHA_MODIFICA = CONVERT(DATETIME, '" & Format(strFecMod, "dd-mm-yy hh:mm:ss") & "', 5) "
        strSQLSolicitud = strSQLSolicitud & "WHERE "
        strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' AND "
        strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLSolicitud
    
    guardarSolicitud = True
    
    strSQLSelectAlter = strSQLSolicitud
    strSQLSolicitud = vbNullString
    
    Exit Function
errGuardarSolicitud:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: GuardarSolicitud"
    
    guardarSolicitud = False
    
    Err.Clear
End Function

Public Sub guardarSolicitudDetalleOneByOne()
    On Error GoTo errGuardarSolicitudDetalleOneByOne
    
    strSQLSolicitud = vbNullString
        
    strSQLSolicitud = strSQLSolicitud & "INSERT INTO PROCESOS.TB_DETSOLICITUD("
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO, COD_SOLICITUD, CS_FENTREGA, "
    strSQLSolicitud = strSQLSolicitud & "DS_CANTIDAD, COD_PRODUCTO, DS_UNIDMED, "
    strSQLSolicitud = strSQLSolicitud & "DS_DESCRIPCION, PRECIO, CANDIS, "
    strSQLSolicitud = strSQLSolicitud & "CS_AFECTO, F5AFECTO, CS_SUBTOT, "
    strSQLSolicitud = strSQLSolicitud & "ITEM, F5MARCA, F5CODMARCA, "
    strSQLSolicitud = strSQLSolicitud & "F5CODFAB, F5CODCOSTO, CENTRO, "
    strSQLSolicitud = strSQLSolicitud & "F5CONIGV, F5SINIGV, OBSERVA, "
    strSQLSolicitud = strSQLSolicitud & "PARTIDA, DES_PARTIDA, RUC_PROVEEDOR, "
    strSQLSolicitud = strSQLSolicitud & "PROVEEDOR) "
    
    strSQLSolicitud = strSQLSolicitud & "VALUES("
    strSQLSolicitud = strSQLSolicitud & "'" & strTipoDocumento & "', '" & strCodigo & "', '" & strFechaEntrega & "', "
    strSQLSolicitud = strSQLSolicitud & dblCantidad & ", '" & strCodProducto & "', '" & strCodUniMed & "', "
    strSQLSolicitud = strSQLSolicitud & "'" & strDescripcion & "', " & dblPrecio & ", " & dblCantidad2 & ", "
    strSQLSolicitud = strSQLSolicitud & "'" & IIf(bolAfecto, "*", vbNullString) & "', " & IIf(bolAfecto, "1", "0") & ", " & dblSubTotal & ", "
    strSQLSolicitud = strSQLSolicitud & dblItem & ", '" & strMarca & "', '" & strCodMarca & "', "
    strSQLSolicitud = strSQLSolicitud & "'', '', '', "
    strSQLSolicitud = strSQLSolicitud & dblPrecioConIGV & ", " & dblPrecioSinIGV & ", '" & strObservaciones & "', "
    strSQLSolicitud = strSQLSolicitud & "'" & strCodPartida & "' ,'" & strDesPartida & "' , '" & strProveedorSugerido & "', "
    strSQLSolicitud = strSQLSolicitud & "'" & strNombreProveedor & "'"
    strSQLSolicitud = strSQLSolicitud & ")"
    
    cnBdCPlus.Execute strSQLSolicitud
    
    strSQLSelectAlter = strSQLSolicitud
    strSQLSolicitud = vbNullString
    
    Exit Sub
errGuardarSolicitudDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: GuardarSolicitudDetalleOneByOne"
    
    Err.Clear
End Sub

Public Function eliminarSolicitud() As Boolean
    On Error GoTo errEliminarSolicitud
    
    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "COD_SOLICITUD", "PROCESOS.IF3ORDEN", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarSolicitud = False
        
        Exit Function
    End If
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "DELETE "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CS_DOCUMENTO = '" & strTipoDocumento & "' AND "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLSolicitud
    
    eliminarSolicitud = True
    
    strSQLSolicitud = vbNullString
    
    Exit Function
errEliminarSolicitud:
    eliminarSolicitud = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: EliminarSolicitud"
    
    Err.Clear
End Function

Public Sub listarSolicitud(ByVal comboList As Object)
    On Error GoTo errListarSolicitud
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DESCRIPCION "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "ESTADO = TRUE "
    strSQLSolicitud = strSQLSolicitud & "ORDER BY "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        comboList.Clear
        comboList.AddItem "(*) - Todos"
        
        Do While Not rstSolicitud.EOF
            comboList.AddItem Trim(rstSolicitud!observaciones & "") & Space(150) & Trim(rstSolicitud!Codigo & "")
            
            rstSolicitud.MoveNext
        Loop
    End If
    
    comboList.ListIndex = 0
    
    rstSolicitud.Close
    strSQLSolicitud = vbNullString
    
    Set rstSolicitud = Nothing
    
    Exit Sub
errListarSolicitud:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: ListarSolicitud"
    
    Err.Clear
End Sub

Public Sub listarSolicitudSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarSolicitudSoloSeleccion
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DESCRIPCION "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "PROCESOS.TB_CABSOLICITUD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "ESTADO = TRUE "
    strSQLSolicitud = strSQLSolicitud & "ORDER BY "
    strSQLSolicitud = strSQLSolicitud & "COD_SOLICITUD"
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        comboList.Clear
        
        Do While Not rstSolicitud.EOF
            comboList.AddItem Trim(rstSolicitud!observaciones & "") & Space(50) & Trim(rstSolicitud!Codigo & "")
            
            rstSolicitud.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstSolicitud.Close
    strSQLSolicitud = vbNullString
    
    Set rstSolicitud = Nothing
    
    Exit Sub
errListarSolicitudSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: ListarSolicitudSoloSeleccion"
    
    Err.Clear
End Sub







'VISUALIZAR DETALLE DE CONSOLIDADO DE INSUMOS DE PEDIDO
Public Sub visualizarDetalleMovimientoPedidoCompromiso(ByVal grilla As dxDBGrid, _
                                                        ByVal strNroPedido As String, _
                                                        ByVal strCodProducto As String, _
                                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errVisualizarDetalleMovimientoPedidoCompromiso
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_DetalleMovimientoPedidoCompromiso"
        
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 50, strNroPedido)
        .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT PRODUCTO, CONVERT(NVARCHAR(10), FECHA,103) AS FECHA, VALE, CONCEPTO, USUARIO, NROOC, NROPEDIDO, CANTIDAD FROM " & strNombreTablaSQLParaResumen & " ORDER BY FECHA, VALE"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Producto"
                .DisableEditor = True
                .FieldName = "PRODUCTO"
                .Font.Bold = True
                .Font.Charset = 0
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
            End With
            
            'Columna Fecha de Movimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fecha de Mov."
                .DisableEditor = True
                .FieldName = "FECHA"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaMovimiento"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Numero de Vale
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Vale"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "VALE"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoVale"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Concepto de Movimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Concepto de Movimiento"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CONCEPTO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColConceptoMovimiento"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
            End With
            
            'Columna Usuario de Registro Movimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Usuario de Registro"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "USUARIO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColUsuarioRegistro"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROOC"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Nro de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Cantidad en Movimiento
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 80
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLSolicitud
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "VALE"
            
            .Columns.ColumnByFieldName("PRODUCTO").GroupIndex = 0
            
            .m.FullExpand
            
            .Bands.BandByName("bndPrincipal").Caption = "No. Pedido: " & strNroPedido & " / " & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CS_OBSERVACIONES", "PROCESOS.TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
            .Columns.HeaderFont.Bold = True
            .GroupNodeColor = RGB(75, 172, 198)
        End With
    End If
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errVisualizarDetalleMovimientoPedidoCompromiso:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsSolicitud: VisualizarDetalleMovimientoPedidoCompromiso"
    End Select
    
    Err.Clear
End Sub

'VISUALIZAR DETALLE DE CONSOLIDADO DE INSUMOS DE PEDIDO
Public Sub visualizarDetalleMovimientoPedidoPorLlegar(ByVal grilla As dxDBGrid, _
                                                        ByVal strNroPedido As String, _
                                                        ByVal strCodProducto As String, _
                                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errVisualizarDetalleMovimientoPedidoPorLlegar
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_DetalleMovimientoPedidoPorLlegar"
        
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 50, strNroPedido)
        .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT PRODUCTO, CONVERT(NVARCHAR(10), FECHA, 103) AS FECHA, NROOC, USUARIO, NROPEDIDO, CANTIDAD FROM " & strNombreTablaSQLParaResumen & " ORDER BY FECHA, PRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Producto"
                .DisableEditor = True
                .FieldName = "PRODUCTO"
                .Font.Bold = True
                .Font.Charset = 0
                .FontColor = vbWhite
                .HeaderAlignment = taCenter
                .ObjectName = "ColProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
            End With
            
            'Columna Fecha de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fecha de Orden"
                .DisableEditor = True
                .FieldName = "FECHA"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Numero de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Orden de Compra"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROOC"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOc"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Usuario de Registro de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Solicitante"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "USUARIO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColUsuarioOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Nro de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Cantidad en Movimiento
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                '.Font.Bold = True
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 80
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLSolicitud
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "NROOC"
            
            .Columns.ColumnByFieldName("PRODUCTO").GroupIndex = 0
            
            .m.FullExpand
            
            .Bands.BandByName("bndPrincipal").Caption = "No. Pedido: " & strNroPedido & " / " & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CS_OBSERVACIONES", "PROCESOS.TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
            .Columns.HeaderFont.Bold = True
            .GroupNodeColor = RGB(75, 172, 198)
        End With
    End If
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errVisualizarDetalleMovimientoPedidoPorLlegar:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsSolicitud: VisualizarDetalleMovimientoPedidoPorLlegar"
    End Select
    
    Err.Clear
End Sub

'Listar Resumen de Estado de Requerimiento en Grilla (QuamtumGrid)
Public Sub listarGrillaResumenRequerimiento(ByVal grilla As dxDBGrid, _
                                            ByVal listaImagenes As Object, _
                                            ByVal strNroPedido As String, _
                                            Optional ByVal strCodigoProducto As String)
    
    On Error GoTo errListarGrillaResumenRequerimiento
    
    strSQLSolicitud = vbNullString
    
    strSQLSolicitud = strSQLSolicitud & "INSERT INTO TMPUTILRESUMENREQUERIMIENTO("
    strSQLSolicitud = strSQLSolicitud & "LLAVE, NROPEDIDO, ITEM, CODPRODUCTO, "
    strSQLSolicitud = strSQLSolicitud & "NOMPRODUCTO, UM, CANTIDAD, "
    strSQLSolicitud = strSQLSolicitud & "COMPROMISOEA, COMPROMISOPL, SALDOACTUAL, "
    strSQLSolicitud = strSQLSolicitud & "COMPROMISOEAG, COMPROMISOPLG, "
    strSQLSolicitud = strSQLSolicitud & "LIBREEAG, LIBREPLG, "
    strSQLSolicitud = strSQLSolicitud & "STOCKEAG, STOCKPLG, "
    strSQLSolicitud = strSQLSolicitud & "CANTIDADPC, PROCESAR, "
    strSQLSolicitud = strSQLSolicitud & "ATENDIDOPORPROV) "
    
    strSQLSolicitud = strSQLSolicitud & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "TRIM(DS.COD_SOLICITUD & '') & TRIM(DS.COD_PRODUCTO & '') AS LLAVE, "
    strSQLSolicitud = strSQLSolicitud & "DS.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DS.ITEM, "
    strSQLSolicitud = strSQLSolicitud & "DS.COD_PRODUCTO, "
    strSQLSolicitud = strSQLSolicitud & "TRIM(DS.DS_DESCRIPCION & '') AS DS_DESCRIPCION, "
    strSQLSolicitud = strSQLSolicitud & "MED.F7SIGMED AS DS_UNIDMED, "
    strSQLSolicitud = strSQLSolicitud & "VAL(DS.DS_CANTIDAD & '') AS DS_CANTIDAD, "
    strSQLSolicitud = strSQLSolicitud & "0 AS CANTCOMPROMETIDA, "
    strSQLSolicitud = strSQLSolicitud & "0 AS CANTPORLLEGAR, "
    strSQLSolicitud = strSQLSolicitud & "0 AS SALDO, "
    'strSQLSolicitud = strSQLSolicitud & "VAL(COMPROMISO.CANTIDAD & '') AS CANTCOMPROMETIDA, "
    'strSQLSolicitud = strSQLSolicitud & "VAL(PORLLEGAR.CANTIDAD & '') AS CANTPORLLEGAR, "
    'strSQLSolicitud = strSQLSolicitud & "VAL(DS.DS_CANTIDAD & '') - (VAL(COMPROMISO.CANTIDAD & '') + VAL(PORLLEGAR.CANTIDAD & '')) AS SALDO, "
    strSQLSolicitud = strSQLSolicitud & "0 AS COMPROMISOEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS COMPROMISOPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS LIBREEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS LIBREPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS STOCKEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS STOCKPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS CANTIDADPC, "
    strSQLSolicitud = strSQLSolicitud & "FALSE AS PROCESAR, "
    strSQLSolicitud = strSQLSolicitud & "FALSE AS ATENDIDOPORPROV "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    'strSQLSolicitud = strSQLSolicitud & "(((PROCESOS.TB_DETSOLICITUD AS DS "
    strSQLSolicitud = strSQLSolicitud & "(PROCESOS.TB_DETSOLICITUD AS DS "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN PROCESOS.TB_CABSOLICITUD AS CS ON CS.CS_DOCUMENTO = DS.CS_DOCUMENTO AND CS.COD_SOLICITUD = DS.COD_SOLICITUD) "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DS.DS_UNIDMED " ') "
    
    'COMPROMISO
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
'    strSQLSolicitud = strSQLSolicitud & "(SELECT "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, DET.F5CODPRO, SUM(DET.F3CANPRO) AS CANTIDAD "
'    strSQLSolicitud = strSQLSolicitud & "FROM "
'    strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
'    strSQLSolicitud = strSQLSolicitud & "WHERE "
'    strSQLSolicitud = strSQLSolicitud & "CAB.F4TIPOVALE = 'I' AND "
'    strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XCS') "
'
'        If strNroPedido <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strNroPedido & "' "
'        End If
'
'        If strCodigoProducto <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPRO = '" & strCodigoProducto & "' "
'        End If
'
'    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F5CODPRO) AS COMPROMISO "
'    strSQLSolicitud = strSQLSolicitud & "ON COMPROMISO.NROPEDIDO = DS.COD_SOLICITUD AND COMPROMISO.F5CODPRO = DS.COD_PRODUCTO) "
'
'    'POR LLEGAR
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
'    strSQLSolicitud = strSQLSolicitud & "(SELECT "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F3CODPRO, "
'    strSQLSolicitud = strSQLSolicitud & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
'    strSQLSolicitud = strSQLSolicitud & "FROM "
'    strSQLSolicitud = strSQLSolicitud & "(IF3ORDEN AS DET "
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
'
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
'    strSQLSolicitud = strSQLSolicitud & "(SELECT "
'    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
'    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL, "
'    strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
'    strSQLSolicitud = strSQLSolicitud & "FROM "
'    strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
'    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'    strSQLSolicitud = strSQLSolicitud & "WHERE "
'    'strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
'    strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0') AND "
'    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD <> '' AND "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD <> '' " 'AND "
'    'strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
'    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
'    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
'    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL) AS INGRESOS "
'    strSQLSolicitud = strSQLSolicitud & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
'
'
'    strSQLSolicitud = strSQLSolicitud & "WHERE "
'    strSQLSolicitud = strSQLSolicitud & "TRIM(DET.F3CODPRO & '') <> '' "
'
'        If strNroPedido <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strNroPedido & "' "
'        End If
'
'        If strCodigoProducto <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.F3CODPRO = '" & strCodigoProducto & "' "
'        End If
'
'    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
'    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F3CODPRO) AS PORLLEGAR "
'    strSQLSolicitud = strSQLSolicitud & "ON PORLLEGAR.COD_SOLICITUD = DS.COD_SOLICITUD AND PORLLEGAR.F3CODPRO = DS.COD_PRODUCTO "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "DS.CS_DOCUMENTO = 'OC' AND "
        
        If strNroPedido <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "DS.COD_SOLICITUD = '" & strNroPedido & "' AND "
        End If
        
        If strCodigoProducto <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "DS.COD_PRODUCTO = '" & strCodigoProducto & "' AND "
        End If
        
    strSQLSolicitud = strSQLSolicitud & "CS.CS_ESTADO NOT IN ('0', '1', '5') "
    strSQLSolicitud = strSQLSolicitud & "ORDER BY "
    strSQLSolicitud = strSQLSolicitud & "DS.COD_SOLICITUD, DS.ITEM"
    
    If Not grilla Is Nothing Then
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
        
        abrirCnTemporal
        
        cnBdCPlus.Execute strSQLSolicitud
        
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Llave
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Llave"
                .Color = &H80C0FF
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Requerimiento"
                .Color = &H80C0FF
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = IIf(strCodigoProducto = vbNullString, False, True) 'False
            End With
            
            'Columna Item
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Item"
                .DisableEditor = True
                .FieldName = "ITEM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColItem"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Descripcion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción del Producto"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso en Almacen a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEA"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEA"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso por Llegar a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPL"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPL"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Saldo del Producto por Atender
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Saldo"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDOACTUAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldoActual"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            'Columna Compromiso en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            'Columna Libre en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibreEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Libre Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibrePLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPLG"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            
            
            
'            'Columna Estado de Seguimiento
'            Set gColumn = .Columns.Add(gedImageEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .BandIndex = 0
'                .Caption = "Obs."
'                .DisableEditor = True
'                .FieldName = "ESTADOSEGUIMIENTO"
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColEstadoSeguimiento"
'
'                With .ImageColumn
'                    .Images = listaImagenes.hImageList
'
'                    .ImageIndexes.Add ("0") 'Agregar
'                    .Values.Add ("ADD")
'                    .Descriptions.Add ("Add")
'
'                    .ImageIndexes.Add ("1") 'Editar
'                    .Values.Add ("EDIT")
'                    .Descriptions.Add ("Edit")
'
'                    .ImageIndexes.Add ("2") 'Alarma
'                    .Values.Add ("ALARM")
'                    .Descriptions.Add ("Alarm")
'
'                    .ShowDescription = False
'                End With
'
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 40
'            End With
            
            'Columna Cantidad Por Comprar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. Req."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = False
                .FieldName = "CANTIDADPC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadPC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "O.K."
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Atendido por Proveedor
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "Atte. Prov."
                .FieldName = "ATENDIDOPORPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAtendidoPorProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = False
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILRESUMENREQUERIMIENTO ORDER BY NOMPRODUCTO"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            '.Columns.ColumnByFieldName("INFOOC").GroupIndex = 0
            
            '.m.FullExpand
            
            '.Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errListarGrillaResumenRequerimiento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsVale: ListarGrillaResumenRequerimiento"
    End Select
    
    Err.Clear
End Sub

Public Sub verificarAtencionPorProducto()
    On Error GoTo errVerificarAtencionPorProducto
    
    Set rstSolicitud = New ADODB.Recordset
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "DS.COD_PRODUCTO AS CODIGO, "
    strSQLSolicitud = strSQLSolicitud & "VAL(DS.DS_CANTIDAD & '') AS CANTSOLICITADA, "
    strSQLSolicitud = strSQLSolicitud & "VAL(COMPROMISO.CANTIDAD & '') AS CANTCOMPROMETIDA, "
    strSQLSolicitud = strSQLSolicitud & "VAL(PORLLEGAR.CANTIDAD & '') AS CANTPORLLEGAR, "
    strSQLSolicitud = strSQLSolicitud & "VAL(DS.DS_CANTIDAD & '') - (VAL(COMPROMISO.CANTIDAD & '') + VAL(PORLLEGAR.CANTIDAD & '')) AS SALDO "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "(PROCESOS.TB_DETSOLICITUD AS DS "
    
    'COMPROMISO EN ALMACEN
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
    strSQLSolicitud = strSQLSolicitud & "(SELECT "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, "
    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO, "
    strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CAB.F4TIPOVALE = 'I' AND "
    strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XCS') "
    
        If strCodigo <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPRO = '" & strCodProducto & "' "
        End If
        
    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO) AS COMPROMISO "
    strSQLSolicitud = strSQLSolicitud & "ON COMPROMISO.NROPEDIDO = DS.COD_SOLICITUD AND COMPROMISO.F5CODPRO = DS.COD_PRODUCTO) "
    
    'COMPROMISO POR LLEGAR
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
    strSQLSolicitud = strSQLSolicitud & "(SELECT "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO, "
    strSQLSolicitud = strSQLSolicitud & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "(IF3ORDEN AS DET "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
    strSQLSolicitud = strSQLSolicitud & "(SELECT "
    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL, "
    strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    'strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0') AND "
    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD <> '' "
        
        If strCodigo <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
        End If
        
    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
    strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL) AS INGRESOS "
    strSQLSolicitud = strSQLSolicitud & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
    
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD = '" & strCodigo & "' "
        
        If strCodProducto <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DET.F3CODPRO = '" & strCodProducto & "' "
        End If
        
    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
    strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO) AS PORLLEGAR "
    strSQLSolicitud = strSQLSolicitud & "ON PORLLEGAR.COD_SOLICITUD = DS.COD_SOLICITUD AND PORLLEGAR.F3CODPRO = DS.COD_PRODUCTO "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "DS.CS_DOCUMENTO = 'OC' "
        
        If strCodigo <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DS.COD_SOLICITUD = '" & strCodigo & "' "
        End If
            
        If strCodProducto <> vbNullString Then
            strSQLSolicitud = strSQLSolicitud & "AND DS.COD_PRODUCTO = '" & strCodProducto & "' "
        End If
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    rstSolicitud.Open strSQLSolicitud, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitud.EOF Then
        strCodProducto = Trim(rstSolicitud!Codigo & "")
        
        dblCantidad = Val(rstSolicitud!CANTSOLICITADA & "")
        dblCompromisoEA = Val(rstSolicitud!CANTCOMPROMETIDA & "")
        dblCompromisoPL = Val(rstSolicitud!CANTPORLLEGAR & "")
        dblSaldoPorAtender = Val(rstSolicitud!SALDO & "")
    Else
        strCodProducto = vbNullString
        
        dblCantidad = 0
        dblCompromisoEA = 0
        dblCompromisoPL = 0
        dblSaldoPorAtender = 0
    End If
    
    If rstSolicitud.State = 1 Then rstSolicitud.Close
    
    Set rstSolicitud = Nothing
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errVerificarAtencionPorProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: VerificarAtencionPorProducto"
    'Resume
    Err.Clear
End Sub

Public Sub descargarResumenAtencionPorProducto(ByVal strTipoCompromisoForV As String)
    On Error GoTo errDescargarResumenAtencionPorProducto
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "INSERT INTO TMPUTILRESUMENATENCIONREQUERIMIENTO("
    strSQLSolicitud = strSQLSolicitud & "TIPO, NROPEDIDO, CODPRODUCTO, CANTIDAD"
    strSQLSolicitud = strSQLSolicitud & ") "
    
    strSQLSolicitud = strSQLSolicitud & "IN '" & wrutatemp & "Templus.mdb' "
    
    Select Case strTipoCompromisoForV
        Case "F"
            strSQLSolicitud = strSQLSolicitud & "SELECT "
            strSQLSolicitud = strSQLSolicitud & "'" & strTipoCompromisoForV & "' AS TIPO, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO AS CODPRODUCTO, "
            strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "CAB.F4TIPOVALE = 'I' AND "
            strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XCS') "
                
                If strCodigo <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
                End If
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPRO = '" & strCodProducto & "' "
                End If
                
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO"
        Case "V"
            strSQLSolicitud = strSQLSolicitud & "SELECT "
            strSQLSolicitud = strSQLSolicitud & "'" & strTipoCompromisoForV & "' AS TIPO, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, "
            strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO AS CODPRODUCTO, "
            strSQLSolicitud = strSQLSolicitud & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "(IF3ORDEN AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
            strSQLSolicitud = strSQLSolicitud & "(SELECT "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL, "
            strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0') AND "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD <> '' "
                
                If strCodigo <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
                End If
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
                End If
                
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL) AS INGRESOS "
            strSQLSolicitud = strSQLSolicitud & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD = '" & strCodigo & "' "
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F3CODPRO = '" & strCodProducto & "' "
                End If
            
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO"
    End Select
    
    cnBdCPlus.Execute strSQLSolicitud
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errDescargarResumenAtencionPorProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: DescargarResumenAtencionPorProducto"
    
    Err.Clear
End Sub

Public Sub descargarResumenAtencionPorProductoResOP(ByVal strTipoCompromisoForV As String, _
                                                    Optional ByVal strNombreTablaTemporalSQL As String)
    
    On Error GoTo errDescargarResumenAtencionPorProductoResOP
    
    Dim cmdOP As ADODB.Command
    Dim rsAtencion As ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    Set rsAtencion = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarResumenAtencionPorProductoResOP"
        
        .Parameters.Append .CreateParameter("@TIPOCOMPROMISO", adVarChar, adParamInput, 1, strTipoCompromisoForV)
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strCodigo)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaTemporalSQL)
        
        Set rsAtencion = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rsAtencion.EOF Then
        Do While Not rsAtencion.EOF
            abrirCnTemporal
            
            strSQLSolicitud = vbNullString
            strSQLSolicitud = strSQLSolicitud & "INSERT INTO TMPUTILRESUMENATENCIONREQUERIMIENTOOP "
            strSQLSolicitud = strSQLSolicitud & "VALUES("
            strSQLSolicitud = strSQLSolicitud & "'" & Trim(rsAtencion!Tipo & "") & "', "
            strSQLSolicitud = strSQLSolicitud & "'" & Trim(rsAtencion!NroPedido & "") & "', "
            strSQLSolicitud = strSQLSolicitud & "'" & Trim(rsAtencion!CodProducto & "") & "', "
            strSQLSolicitud = strSQLSolicitud & Val(rsAtencion!Cantidad & "")
            strSQLSolicitud = strSQLSolicitud & ")"
            
            cnDBTemp.Execute strSQLSolicitud
            
            rsAtencion.MoveNext
        Loop
    End If
    
    If rsAtencion.State = 1 Then rsAtencion.Close
    
    Set rsAtencion = Nothing
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errDescargarResumenAtencionPorProductoResOP:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: DescargarResumenAtencionPorProductoResOP"
    
    Err.Clear
End Sub

Public Sub descargarResumenAtencionPorProductoResCompromisoAfectado(ByVal strTipoCompromisoForV As String)
    On Error GoTo errDescargarResumenAtencionPorProductoResCompromisoAfectado
    
    strSQLSolicitud = vbNullString
    strSQLSolicitud = strSQLSolicitud & "INSERT INTO TMPUTILRESUMENATENCIONCOMPROMISOAFECTADO("
    strSQLSolicitud = strSQLSolicitud & "TIPO, NROPEDIDO, CODPRODUCTO, CANTIDAD"
    strSQLSolicitud = strSQLSolicitud & ") "
    
    strSQLSolicitud = strSQLSolicitud & "IN '" & wrutatemp & "Templus.mdb' "
    
    Select Case strTipoCompromisoForV
        Case "F"
            strSQLSolicitud = strSQLSolicitud & "SELECT "
            strSQLSolicitud = strSQLSolicitud & "'" & strTipoCompromisoForV & "' AS TIPO, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO AS CODPRODUCTO, "
            strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "CAB.F4TIPOVALE = 'I' AND "
            strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XCS') "
                
                If strCodigo <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
                End If
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPRO = '" & strCodProducto & "' "
                End If
                
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPRO"
        Case "V"
            strSQLSolicitud = strSQLSolicitud & "SELECT "
            strSQLSolicitud = strSQLSolicitud & "'" & strTipoCompromisoForV & "' AS TIPO, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, "
            strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO AS CODPRODUCTO, "
            strSQLSolicitud = strSQLSolicitud & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "(IF3ORDEN AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
            strSQLSolicitud = strSQLSolicitud & "(SELECT "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL, "
            strSQLSolicitud = strSQLSolicitud & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
            strSQLSolicitud = strSQLSolicitud & "FROM "
            strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
            strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0') AND "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD <> '' "
                
                If strCodigo <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strCodigo & "' "
                End If
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
                End If
                
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.F4NUMORD, "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F5CODPROORIGINAL) AS INGRESOS "
            strSQLSolicitud = strSQLSolicitud & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            
            strSQLSolicitud = strSQLSolicitud & "WHERE "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD = '" & strCodigo & "' "
                
                If strCodProducto <> vbNullString Then
                    strSQLSolicitud = strSQLSolicitud & "AND DET.F3CODPRO = '" & strCodProducto & "' "
                End If
            
            strSQLSolicitud = strSQLSolicitud & "GROUP BY "
            strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, "
            strSQLSolicitud = strSQLSolicitud & "DET.F3CODPRO"
    End Select
    
    cnBdCPlus.Execute strSQLSolicitud
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errDescargarResumenAtencionPorProductoResCompromisoAfectado:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsSolicitud: DescargarResumenAtencionPorProductoResCompromisoAfectado"
    
    Err.Clear
End Sub

'Listar Reposicion de Compromiso en Grilla (QuamtumGrid)
Public Sub listarGrillaReposicionCompromiso(ByVal grilla As dxDBGrid, _
                                            ByVal listaImagenes As Object, _
                                            ByVal strNroPedido As String, _
                                            Optional ByVal strCodigoProducto As String)
    
    On Error GoTo errListarGrillaReposicionCompromiso
    
    strSQLSolicitud = vbNullString
    
    strSQLSolicitud = strSQLSolicitud & "INSERT INTO TMPUTILREPOSICIONCOMPROMISO("
    strSQLSolicitud = strSQLSolicitud & "ID, NROPEDIDO, CODPRODUCTO, "
    strSQLSolicitud = strSQLSolicitud & "NOMPRODUCTO, UM, CANTIDAD, "
    strSQLSolicitud = strSQLSolicitud & "COMPROMISOEA, COMPROMISOPL, SALDOACTUAL, "
    strSQLSolicitud = strSQLSolicitud & "COMPROMISOEAG, COMPROMISOPLG, "
    strSQLSolicitud = strSQLSolicitud & "LIBREEAG, LIBREPLG, "
    strSQLSolicitud = strSQLSolicitud & "STOCKEAG, STOCKPLG, "
    strSQLSolicitud = strSQLSolicitud & "CANTIDADPC, DESCARTAR, "
    strSQLSolicitud = strSQLSolicitud & "PROCESAR, ATENDIDOPORPROV) "
    
    strSQLSolicitud = strSQLSolicitud & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLSolicitud = strSQLSolicitud & "SELECT "
    strSQLSolicitud = strSQLSolicitud & "RC.IDREPOSICION, "
    strSQLSolicitud = strSQLSolicitud & "RC.NROPEDIDO, "
    strSQLSolicitud = strSQLSolicitud & "RC.IDINSUMO, "
    strSQLSolicitud = strSQLSolicitud & "PROD.F5NOMPRO, "
    strSQLSolicitud = strSQLSolicitud & "MED.F7SIGMED, "
    strSQLSolicitud = strSQLSolicitud & "RC.CANTIDAD, "
    strSQLSolicitud = strSQLSolicitud & "VAL(COMPROMISO.CANTIDAD & '') AS CANTCOMPROMETIDA, "
    strSQLSolicitud = strSQLSolicitud & "VAL(PORLLEGAR.CANTIDAD & '') AS CANTPORLLEGAR, "
    strSQLSolicitud = strSQLSolicitud & "VAL(RC.CANTIDAD & '') - (VAL(COMPROMISO.CANTIDAD & '') + VAL(PORLLEGAR.CANTIDAD & '')) AS SALDO, "
    strSQLSolicitud = strSQLSolicitud & "0 AS COMPROMISOEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS COMPROMISOPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS LIBREEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS LIBREPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS STOCKEAG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS STOCKPLG, "
    strSQLSolicitud = strSQLSolicitud & "0 AS CANTIDADPC, "
    strSQLSolicitud = strSQLSolicitud & "FALSE AS DESCARTAR, "
    strSQLSolicitud = strSQLSolicitud & "FALSE AS PROCESAR, "
    strSQLSolicitud = strSQLSolicitud & "FALSE AS ATENDIDOPORPROV "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "(((SF1REPOSICIONCOMPROMISO AS RC "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN PROCESOS.TB_DETSOLICITUD AS DS ON DS.COD_SOLICITUD = RC.NROPEDIDO) "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF5PLA AS PROD ON PROD.F5NOMPRO = RC.IDINSUMO) "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
    strSQLSolicitud = strSQLSolicitud & "(SELECT "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD AS NROPEDIDO, DET.F5CODPRO, SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "IF3VALES AS DET "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "CAB.F4TIPOVALE = 'I' AND "
    strSQLSolicitud = strSQLSolicitud & "CAB.F1CODORI IN ('XC0', 'XCS') "
        
'        If strNroPedido <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strNroPedido & "' "
'        End If
'
'        If strCodigoProducto <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.F5CODPRO = '" & strCodigoProducto & "' "
'        End If
        
    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F5CODPRO) AS COMPROMISO "
    strSQLSolicitud = strSQLSolicitud & "ON COMPROMISO.NROPEDIDO = RC.NROPEDIDO AND COMPROMISO.F5CODPRO = RC.IDINSUMO) "
    
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN "
    strSQLSolicitud = strSQLSolicitud & "(SELECT "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F3CODPRO, "
    strSQLSolicitud = strSQLSolicitud & "SUM((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD "
    strSQLSolicitud = strSQLSolicitud & "FROM "
    strSQLSolicitud = strSQLSolicitud & "IF3ORDEN AS DET "
    strSQLSolicitud = strSQLSolicitud & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "TRIM(DET.F3CODPRO & '') <> '' "
        
'        If strNroPedido <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.COD_SOLICITUD = '" & strNroPedido & "' "
'        End If
'
'        If strCodigoProducto <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "AND DET.F3CODPRO = '" & strCodigoProducto & "' "
'        End If
        
    strSQLSolicitud = strSQLSolicitud & "GROUP BY "
    strSQLSolicitud = strSQLSolicitud & "DET.COD_SOLICITUD, DET.F3CODPRO) AS PORLLEGAR "
    strSQLSolicitud = strSQLSolicitud & "ON PORLLEGAR.COD_SOLICITUD = RC.NROPEDIDO AND PORLLEGAR.F3CODPRO = RC.IDINSUMO "
    strSQLSolicitud = strSQLSolicitud & "WHERE "
    strSQLSolicitud = strSQLSolicitud & "RC.DESCARTARREPOSICION = FALSE AND "
    strSQLSolicitud = strSQLSolicitud & "RC.REPOSICIONENCURSO = FALSE "
        
'        If strNroPedido <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "DS.COD_SOLICITUD = '" & strNroPedido & "' AND "
'        End If
'
'        If strCodigoProducto <> vbNullString Then
'            strSQLSolicitud = strSQLSolicitud & "DS.COD_PRODUCTO = '" & strCodigoProducto & "' AND "
'        End If
        
    'strSQLSolicitud = strSQLSolicitud & "CS.CS_ESTADO NOT IN ('0', '1', '5') "
    'strSQLSolicitud = strSQLSolicitud & "ORDER BY "
    'strSQLSolicitud = strSQLSolicitud & "DS.COD_SOLICITUD, DS.ITEM"
    
    If Not grilla Is Nothing Then
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILREPOSICIONCOMPROMISO"
        
        abrirCnTemporal
        
        cnBdCPlus.Execute strSQLSolicitud
        
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna ID
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "ID"
                .DisableEditor = True
                .FieldName = "ID"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColID"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Requerimiento"
                .Color = &H80C0FF
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = True
                '.Visible = IIf(strCodigoProducto = vbNullString, False, True) 'False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Descripcion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción del Producto"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso en Almacen a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEA"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEA"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso por Llegar a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPL"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPL"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Saldo del Producto por Atender
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Saldo"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDOACTUAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldoActual"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            'Columna Compromiso en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            'Columna Libre en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibreEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Libre Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibrePLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPLG"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
'            'Columna Estado de Seguimiento
'            Set gColumn = .Columns.Add(gedImageEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .BandIndex = 0
'                .Caption = "Obs."
'                .DisableEditor = True
'                .FieldName = "ESTADOSEGUIMIENTO"
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColEstadoSeguimiento"
'
'                With .ImageColumn
'                    .Images = listaImagenes.hImageList
'
'                    .ImageIndexes.Add ("0") 'Agregar
'                    .Values.Add ("ADD")
'                    .Descriptions.Add ("Add")
'
'                    .ImageIndexes.Add ("1") 'Editar
'                    .Values.Add ("EDIT")
'                    .Descriptions.Add ("Edit")
'
'                    .ImageIndexes.Add ("2") 'Alarma
'                    .Values.Add ("ALARM")
'                    .Descriptions.Add ("Alarm")
'
'                    .ShowDescription = False
'                End With
'
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 40
'            End With
            
            'Columna Cantidad Por Comprar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. Req."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = False
                .FieldName = "CANTIDADPC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadPC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Descartar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "Descartar"
                .FieldName = "DESCARTAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescartar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "O.K."
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Atendido por Proveedor
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "Atte. Prov."
                .FieldName = "ATENDIDOPORPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAtendidoPorProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = False
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILREPOSICIONCOMPROMISO ORDER BY NOMPRODUCTO"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "ID"
            
            '.Columns.ColumnByFieldName("INFOOC").GroupIndex = 0
            
            '.m.FullExpand
            
            '.Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLSolicitud = vbNullString
    
    Exit Sub
errListarGrillaReposicionCompromiso:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsVale: ListarGrillaReposicionCompromiso"
    End Select
    
    Err.Clear
End Sub

