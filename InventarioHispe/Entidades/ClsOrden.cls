VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strTipoOrden                    As String 'F4LOCAL
Private strNumeroOrden                  As String 'F4NUMORD

Private strFechaEmision                 As String 'F4FECEMI
Private strCodProveedor                 As String 'F4CODCLI
Private strNomProveedor                 As String 'F4NOMPROV
Private strRucProveedor                 As String 'F4CODPRV
Private strContactoProveedor            As String 'F4CONTACTO
Private strCodTipoComprobante           As String 'F4TIPDOC
Private bolOrdenRegularizada            As Boolean 'F4REGULARIZA

Private strFechaEntrega                 As String 'F4FECENT
Private strCodigoSolicitante            As String 'F4CODSOL
Private strCodFormaPago                 As String 'F4FORPAG
Private strCentroCosto                  As String 'F4CENTRO
Private strLugarEntrega                 As String 'F4LUGAR_ENTREGA
Private bolPagoParcial                  As Boolean 'F4PAGOPARCIAL
Private strCodMoneda                    As String 'F4TIPMON
Private dblTipoCambio                   As Double 'F4TIPCAM
Private strFechaPago                    As String 'F4DIAPAGO
Private strNumeroCotizacion             As String 'F4NUMCOTIZA

Private bolVB1                          As Boolean 'F4VB1
Private strVB1Usuario                   As String 'F4VBUSER1
Private strVB1Fecha                     As String 'F4VBFECHA1
Private bolVB2                          As Boolean 'F4VB2
Private strVB2Usuario                   As String 'F4VBUSER2
Private strVB2Fecha                     As String 'F4VBFECHA2
Private bolColocada                     As Boolean 'F4COLOCADA
Private strColocadaUsuario              As String 'F4COLOCADAUSER
Private strColocadaFecha                As String 'F4COLOCADAFECHA
Private bolAtendida                     As Boolean 'F4ATENDIDA
Private strAtendidaUsuario              As String 'F4ATENDIDAUSER
Private strAtendidaFecha                As String 'F4ATENDIDAFECHA

Private strEmpresa                      As String 'F4EMPRESA
Private strObservacion                  As String 'F4OBSERVA

Private dblSubTotal                     As Double 'F4BASIMP
Private dblTotalImpuesto                As Double 'F4IGV
Private dblTotalFacturado               As Double 'F4MONTO
Private dblTotalInafecto                As Double 'F4MONINA
Private dblTotalRedondeo                As Double 'F4RND

Private intEstado                       As Integer 'F4ESTADO
Private strCodigoRequerimiento          As String 'F4CODSOLICITUD
Private bolDocAnulado                   As Boolean 'F4ESTNUL
Private bolDocFalta                     As Boolean 'F4FALTA
Private bolOrdenConVale                 As Boolean 'F4ESTVAL
Private strReferencia                   As String 'F4REFERE
Private dblCorrelativo                  As Double 'F4CORRELA
Private bolAplicaDetraccion             As Boolean 'F4DETRACCIONAPLICA
Private dblPorcentajeDetraccion         As Double 'F4DETRACCIONPORC
Private bolSinProveedorEspecifico       As Boolean 'F4SINPROVEEDORESP



'Atributos de Detalle de Orden
Private dblItem                         As Double 'ITEM
Private strRequerimiento                As String 'COD_SOLICITUD
Private strCodProducto                  As String 'F3CODPRO
Private strCodFabricante                As String 'F3CODFAB
Private strNomProducto                  As String 'F5NOMPRO
Private strNomProductoInterno           As String 'F5NOMPRO_ING
Private strCodUM                        As String 'UNIDAD
Private dblCantidad                     As Double 'F3CANPRO
Private dblCantidadMaxima               As Double 'F3CANPRO2
Private dblPorcentajeDemasia            As Double 'F3PORCDEMASIA
Private dblPrecioSinImpuesto            As Double 'F3PRECOS
Private dblPrecioConImpuesto            As Double 'F3PREUNI
Private dblPrecioNetoSinImpuesto        As Double 'F3PRENETO
Private dblPorcentajeDscto              As Double 'F3PORDCT
Private dblTotalDscto                   As Double 'F3TOTDCT
Private bolAfecto                       As Boolean 'F5AFECTO
Private dblImpuestoPorItem              As Double 'F3IGV
Private dblBasePorItem                  As Double 'F5VALVTA
Private dblTotalPorItem                 As Double 'F3TOTAL
Private strCodColor                     As String 'CODCOLOR
Private strObservacionPorItem           As String 'F3OBSERVA
Private strCodigoGasto                  As String 'F3GASTO
Private strCuentaContable               As String 'F3CUENTA


'Atributos Auxiliares
Private dblPorcentajeImpuesto           As Double
Private dblSignoImpuesto                As Double
Private dblCantidadFinal                As Double
Private dblExoneradoPorItem             As Double


'Atributo Adicionales
Private dblCantidadFaltante             As Double
Private bolItemAjustado                 As Boolean


Private strFecReg                      As String 'F4FECGRA
Private strUsuReg                      As String 'F4USEGRA
Private strFecMod                      As String 'F4FECMOD
Private strUsuMod                      As String 'F4USEMOD

Private strResponsable                  As String 'F4RESPONSABLE

Private strSQLSelectAlter               As String

Private rstOrden             As ADODB.Recordset
Private strSQLOrden          As String

Private strDerivadoDNI                     As String

Public Property Let TipoOrden(ByVal Value As String)
    strTipoOrden = Value
End Property

Public Property Get TipoOrden() As String
    TipoOrden = strTipoOrden
End Property

Public Property Let NumeroOrden(ByVal Value As String)
    strNumeroOrden = Value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property




Public Property Let FechaEmision(ByVal Value As String)
    strFechaEmision = Value
End Property

Public Property Get FechaEmision() As String
    FechaEmision = strFechaEmision
End Property

Public Property Let CodProveedor(ByVal Value As String)
    strCodProveedor = Value
End Property

Public Property Get CodProveedor() As String
    CodProveedor = strCodProveedor
End Property

Public Property Let NomProveedor(ByVal Value As String)
    strNomProveedor = Value
End Property

Public Property Get NomProveedor() As String
    NomProveedor = strNomProveedor
End Property

Public Property Let RucProveedor(ByVal Value As String)
    strRucProveedor = Value
End Property

Public Property Get RucProveedor() As String
    RucProveedor = strRucProveedor
End Property

Public Property Let ContactoProveedor(ByVal Value As String)
    strContactoProveedor = Value
End Property

Public Property Get ContactoProveedor() As String
    ContactoProveedor = strContactoProveedor
End Property

Public Property Let CodTipoComprobante(ByVal Value As String)
    strCodTipoComprobante = Value
End Property

Public Property Get CodTipoComprobante() As String
    CodTipoComprobante = strCodTipoComprobante
End Property

Public Property Let OrdenRegularizada(ByVal Value As Boolean)
    bolOrdenRegularizada = Value
End Property

Public Property Get OrdenRegularizada() As Boolean
    OrdenRegularizada = bolOrdenRegularizada
End Property




Public Property Let FechaEntrega(ByVal Value As String)
    strFechaEntrega = Value
End Property

Public Property Get FechaEntrega() As String
    FechaEntrega = strFechaEntrega
End Property

Public Property Let CodigoSolicitante(ByVal Value As String)
    strCodigoSolicitante = Value
End Property

Public Property Get CodigoSolicitante() As String
    CodigoSolicitante = strCodigoSolicitante
End Property

Public Property Let CodFormaPago(ByVal Value As String)
    strCodFormaPago = Value
End Property

Public Property Get CodFormaPago() As String
    CodFormaPago = strCodFormaPago
End Property

Public Property Let CentroCosto(ByVal Value As String)
    strCentroCosto = Value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property

Public Property Let LugarEntrega(ByVal Value As String)
    strLugarEntrega = Value
End Property

Public Property Get LugarEntrega() As String
    LugarEntrega = strLugarEntrega
End Property

Public Property Let PagoParcial(ByVal Value As Boolean)
    bolPagoParcial = Value
End Property

Public Property Get PagoParcial() As Boolean
    PagoParcial = bolPagoParcial
End Property

Public Property Let CodMoneda(ByVal Value As String)
    strCodMoneda = Value
End Property

Public Property Get CodMoneda() As String
    CodMoneda = strCodMoneda
End Property

Public Property Let DerivadoDNI(ByVal Value As String)
    strDerivadoDNI = Value
End Property

Public Property Get DerivadoDNI() As String
    DerivadoDNI = strDerivadoDNI
End Property

Public Property Let TipoCambio(ByVal Value As Double)
    dblTipoCambio = Value
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dblTipoCambio
End Property

Public Property Let FechaPago(ByVal Value As String)
    strFechaPago = Value
End Property

Public Property Get FechaPago() As String
    FechaPago = strFechaPago
End Property

Public Property Let NumeroCotizacion(ByVal Value As String)
    strNumeroCotizacion = Value
End Property

Public Property Get NumeroCotizacion() As String
    NumeroCotizacion = strNumeroCotizacion
End Property




Public Property Let VB1(ByVal Value As Boolean)
    bolVB1 = Value
End Property

Public Property Get VB1() As Boolean
    VB1 = bolVB1
End Property

Public Property Let VB1Usuario(ByVal Value As String)
    strVB1Usuario = Value
End Property

Public Property Get VB1Usuario() As String
    VB1Usuario = strVB1Usuario
End Property

Public Property Let VB1Fecha(ByVal Value As String)
    strVB1Fecha = Value
End Property

Public Property Get VB1Fecha() As String
    VB1Fecha = strVB1Fecha
End Property

Public Property Let VB2(ByVal Value As Boolean)
    bolVB2 = Value
End Property

Public Property Get VB2() As Boolean
    VB2 = bolVB2
End Property

Public Property Let VB2Usuario(ByVal Value As String)
    strVB2Usuario = Value
End Property

Public Property Get VB2Usuario() As String
    VB2Usuario = strVB2Usuario
End Property

Public Property Let VB2Fecha(ByVal Value As String)
    strVB2Fecha = Value
End Property

Public Property Get VB2Fecha() As String
    VB2Fecha = strVB2Fecha
End Property

Public Property Let Colocada(ByVal Value As Boolean)
    bolColocada = Value
End Property

Public Property Get Colocada() As Boolean
    Colocada = bolColocada
End Property

Public Property Let ColocadaUsuario(ByVal Value As String)
    strColocadaUsuario = Value
End Property

Public Property Get ColocadaUsuario() As String
    ColocadaUsuario = strColocadaUsuario
End Property

Public Property Let ColocadaFecha(ByVal Value As String)
    strColocadaFecha = Value
End Property

Public Property Get ColocadaFecha() As String
    ColocadaFecha = strColocadaFecha
End Property

Public Property Let Atendida(ByVal Value As Boolean)
    bolAtendida = Value
End Property

Public Property Get Atendida() As Boolean
    Atendida = bolAtendida
End Property

Public Property Let AtendidaUsuario(ByVal Value As String)
    strAtendidaUsuario = Value
End Property

Public Property Get AtendidaUsuario() As String
    AtendidaUsuario = strAtendidaUsuario
End Property

Public Property Let AtendidaFecha(ByVal Value As String)
    strAtendidaFecha = Value
End Property

Public Property Get AtendidaFecha() As String
    AtendidaFecha = strAtendidaFecha
End Property




'Empresa
Public Property Let Empresa(ByVal Value As String)
    strEmpresa = Value
End Property

Public Property Get Empresa() As String
    Empresa = strEmpresa
End Property
'Observacion
Public Property Let Observacion(ByVal Value As String)
    strObservacion = Value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property




Public Property Let SUBTOTAL(ByVal Value As Double)
    dblSubTotal = Value
End Property

Public Property Get SUBTOTAL() As Double
    SUBTOTAL = dblSubTotal
End Property

Public Property Let TotalImpuesto(ByVal Value As Double)
    dblTotalImpuesto = Value
End Property

Public Property Get TotalImpuesto() As Double
    TotalImpuesto = dblTotalImpuesto
End Property

Public Property Let TotalFacturado(ByVal Value As Double)
    dblTotalFacturado = Value
End Property

Public Property Get TotalFacturado() As Double
    TotalFacturado = dblTotalFacturado
End Property

Public Property Let TotalInafecto(ByVal Value As Double)
    dblTotalInafecto = Value
End Property

Public Property Get TotalInafecto() As Double
    TotalInafecto = dblTotalInafecto
End Property

Public Property Let TotalRedondeo(ByVal Value As Double)
    dblTotalRedondeo = Value
End Property

Public Property Get TotalRedondeo() As Double
    TotalRedondeo = dblTotalRedondeo
End Property




Public Property Let Estado(ByVal Value As Integer)
    intEstado = Value
End Property

Public Property Get Estado() As Integer
    Estado = intEstado
End Property

Public Property Let CodigoRequerimiento(ByVal Value As String)
    strCodigoRequerimiento = Value
End Property

Public Property Get CodigoRequerimiento() As String
    CodigoRequerimiento = strCodigoRequerimiento
End Property
'DocAnulado
Public Property Let DocAnulado(ByVal Value As Boolean)
    bolDocAnulado = Value
End Property

Public Property Get DocAnulado() As Boolean
    DocAnulado = bolDocAnulado
End Property

Public Property Let DocFalta(ByVal Value As Boolean)
    bolDocFalta = Value
End Property

Public Property Get DocFalta() As Boolean
    DocFalta = bolDocFalta
End Property

Public Property Let OrdenConVale(ByVal Value As Boolean)
    bolOrdenConVale = Value
End Property

Public Property Get OrdenConVale() As Boolean
    OrdenConVale = bolOrdenConVale
End Property

'Referencia
Public Property Let referencia(ByVal Value As String)
    strReferencia = Value
End Property

Public Property Get referencia() As String
    referencia = strReferencia
End Property

Public Property Let Correlativo(ByVal Value As Double)
    dblCorrelativo = Value
End Property

Public Property Get Correlativo() As Double
    Correlativo = dblCorrelativo
End Property

Public Property Let AplicaDetraccion(ByVal Value As Boolean)
    bolAplicaDetraccion = Value
End Property

Public Property Get AplicaDetraccion() As Boolean
    AplicaDetraccion = bolAplicaDetraccion
End Property

Public Property Let PorcentajeDetraccion(ByVal Value As Double)
    dblPorcentajeDetraccion = Value
End Property

Public Property Get PorcentajeDetraccion() As Double
    PorcentajeDetraccion = dblPorcentajeDetraccion
End Property

Public Property Let SinProveedorEspecifico(ByVal Value As Boolean)
    bolSinProveedorEspecifico = Value
End Property

Public Property Get SinProveedorEspecifico() As Boolean
    SinProveedorEspecifico = bolSinProveedorEspecifico
End Property






'ITEM
Public Property Let ITEM(ByVal Value As Double)
    dblItem = Value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property

'COD_SOLICITUD
Public Property Let Requerimiento(ByVal Value As String)
    strRequerimiento = Value
End Property

Public Property Get Requerimiento() As String
    Requerimiento = strRequerimiento
End Property

'F3CODPRO
Public Property Let CodigoProducto(ByVal Value As String)
    strCodProducto = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

'F3CODFAB
Public Property Let CodigoFabricante(ByVal Value As String)
    strCodFabricante = Value
End Property

Public Property Get CodigoFabricante() As String
    CodigoFabricante = strCodFabricante
End Property

'F5NOMPRO
Public Property Let NombreProducto(ByVal Value As String)
    strNomProducto = Value
End Property

Public Property Get NombreProducto() As String
    NombreProducto = strNomProducto
End Property

'F5NOMPRO_ING
Public Property Let NombreProductoInterno(ByVal Value As String)
    strNomProductoInterno = Value
End Property

Public Property Get NombreProductoInterno() As String
    NombreProductoInterno = strNomProductoInterno
End Property

'UNIDAD
Public Property Let CodigoUM(ByVal Value As String)
    strCodUM = Value
End Property

Public Property Get CodigoUM() As String
    CodigoUM = strCodUM
End Property

'F3CANPRO
Public Property Let Cantidad(ByVal Value As Double)
    dblCantidad = Value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

'F3CANPRO2
Public Property Let CantidadMaxima(ByVal Value As Double)
    dblCantidadMaxima = Value
End Property

Public Property Get CantidadMaxima() As Double
    CantidadMaxima = dblCantidadMaxima
End Property

'F3PORCDEMASIA
Public Property Let PorcentajeDemasia(ByVal Value As Double)
    dblPorcentajeDemasia = Value
End Property

Public Property Get PorcentajeDemasia() As Double
    PorcentajeDemasia = dblPorcentajeDemasia
End Property

'F3PRECOS
Public Property Let PrecioSinImpuesto(ByVal Value As Double)
    dblPrecioSinImpuesto = Value
End Property

Public Property Get PrecioSinImpuesto() As Double
    PrecioSinImpuesto = dblPrecioSinImpuesto
End Property

'F3PREUNI
Public Property Let PrecioConImpuesto(ByVal Value As Double)
    dblPrecioConImpuesto = Value
End Property

Public Property Get PrecioConImpuesto() As Double
    PrecioConImpuesto = dblPrecioConImpuesto
End Property

'F3PRENETO
Public Property Let PrecioNetoSinImpuesto(ByVal Value As Double)
    dblPrecioNetoSinImpuesto = Value
End Property

Public Property Get PrecioNetoSinImpuesto() As Double
    PrecioNetoSinImpuesto = dblPrecioNetoSinImpuesto
End Property

'F3PORDCT
Public Property Let PorcentajeDscto(ByVal Value As Double)
    dblPorcentajeDscto = Value
End Property

Public Property Get PorcentajeDscto() As Double
    PorcentajeDscto = dblPorcentajeDscto
End Property

'F3TOTDCT
Public Property Let TotalDscto(ByVal Value As Double)
    dblTotalDscto = Value
End Property

Public Property Get TotalDscto() As Double
    TotalDscto = dblTotalDscto
End Property

'F5AFECTO
Public Property Let Afecto(ByVal Value As Boolean)
    bolAfecto = Value
End Property

Public Property Get Afecto() As Boolean
    Afecto = bolAfecto
End Property

'F3IGV
Public Property Let ImpuestoPorItem(ByVal Value As Double)
    dblImpuestoPorItem = Value
End Property

Public Property Get ImpuestoPorItem() As Double
    ImpuestoPorItem = dblImpuestoPorItem
End Property

'F5VALVTA
Public Property Let BasePorItem(ByVal Value As Double)
    dblBasePorItem = Value
End Property

Public Property Get BasePorItem() As Double
    BasePorItem = dblBasePorItem
End Property

'F3TOTAL
Public Property Let TotalPorItem(ByVal Value As Double)
    dblTotalPorItem = Value
End Property

Public Property Get TotalPorItem() As Double
    TotalPorItem = dblTotalPorItem
End Property

'CODCOLOR
Public Property Let CodigoColor(ByVal Value As String)
    strCodColor = Value
End Property

Public Property Get CodigoColor() As String
    CodigoColor = strCodColor
End Property

'F3OBSERVA
Public Property Let ObservacionPorItem(ByVal Value As String)
    strObservacionPorItem = Value
End Property

Public Property Get ObservacionPorItem() As String
    ObservacionPorItem = strObservacionPorItem
End Property

'F3GASTO
Public Property Let CodigoGasto(ByVal Value As String)
    strCodigoGasto = Value
End Property

Public Property Get CodigoGasto() As String
    CodigoGasto = strCodigoGasto
End Property

'F3CUENTA
Public Property Let CuentaContable(ByVal Value As String)
    strCuentaContable = Value
End Property

Public Property Get CuentaContable() As String
    CuentaContable = strCuentaContable
End Property





'Atributos Auxiliares
Public Property Let PorcentajeImpuesto(ByVal Value As Double)
    dblPorcentajeImpuesto = Value
End Property

Public Property Get PorcentajeImpuesto() As Double
    PorcentajeImpuesto = dblPorcentajeImpuesto
End Property

Public Property Let SignoImpuesto(ByVal Value As Double)
    dblSignoImpuesto = Value
End Property

Public Property Get SignoImpuesto() As Double
    SignoImpuesto = dblSignoImpuesto
End Property

Public Property Let CantidadFinal(ByVal Value As Double)
    dblCantidadFinal = Value
End Property

Public Property Get CantidadFinal() As Double
    CantidadFinal = dblCantidadFinal
End Property

Public Property Let ExoneradoPorItem(ByVal Value As Double)
    dblExoneradoPorItem = Value
End Property

Public Property Get ExoneradoPorItem() As Double
    ExoneradoPorItem = dblExoneradoPorItem
End Property




'Atributos adicionales
'F3CANFAL
Public Property Let CantidadFaltante(ByVal Value As Double)
    dblCantidadFaltante = Value
End Property

Public Property Get CantidadFaltante() As Double
    CantidadFaltante = dblCantidadFaltante
End Property
'F3AJUSTE
Public Property Let ItemAjustado(ByVal Value As Boolean)
    bolItemAjustado = Value
End Property

Public Property Get ItemAjustado() As Boolean
    ItemAjustado = bolItemAjustado
End Property





'Propiedad Fecha de Registro de Orden
Public Property Let FechaReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Orden
Public Property Let UsuarioReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Orden
Public Property Let FechaMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Responsable de Registro de Orden
Public Property Let Responsable(ByVal Value As String)
    strResponsable = Value
End Property

Public Property Get Responsable() As String
    Responsable = strResponsable
End Property


'Propiedad Usuario de Registro de Orden
Public Property Let UsuarioMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property


Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property

Public Sub inicializarEntidades()
    strTipoOrden = vbNullString 'F4LOCAL
    strNumeroOrden = vbNullString 'F4NUMORD
    
    strFechaEmision = vbNullString 'F4FECEMI
    strCodProveedor = vbNullString 'F4CODCLI
    strNomProveedor = vbNullString 'F4NOMPROV
    strRucProveedor = vbNullString 'F4CODPRV
    strContactoProveedor = vbNullString 'F4CONTACTO
    strCodTipoComprobante = vbNullString 'F4TIPDOC
    bolOrdenRegularizada = False 'F4REGULARIZA
    
    strFechaEntrega = vbNullString 'F4FECENT
    strCodigoSolicitante = vbNullString 'F4CODSOL
    strCodFormaPago = vbNullString 'F4FORPAG
    strCentroCosto = vbNullString 'F4CENTRO
    strLugarEntrega = vbNullString 'F4LUGAR_ENTREGA
    bolPagoParcial = False 'F4PAGOPARCIAL
    strCodMoneda = vbNullString 'F4TIPMON
    dblTipoCambio = 0 'F4TIPCAM
    strFechaPago = vbNullString 'F4DIAPAGO
    strNumeroCotizacion = vbNullString 'F4NUMCOTIZA
    
    bolVB1 = False      'F4VB1
    strVB1Usuario = vbNullString 'F4VBUSER1
    strVB1Fecha = vbNullString 'F4VBFECHA1
    bolVB2 = False      'F4VB2
    strVB2Usuario = vbNullString 'F4VBUSER2
    strVB2Fecha = vbNullString 'F4VBFECHA2
    bolColocada = False 'F4COLOCADA
    strColocadaUsuario = vbNullString 'F4COLOCADAUSER
    strColocadaFecha = vbNullString   'F4COLOCADAFECHA
    bolAtendida = False 'F4ATENDIDA
    strAtendidaUsuario = vbNullString 'F4ATENDIDAUSER
    strAtendidaFecha = vbNullString   'F4ATENDIDAFECHA
    
    strEmpresa = vbNullString 'F4EMPRESA
    strObservacion = vbNullString 'F4OBSERVA
    
    dblSubTotal = 0  'F4BASIMP
    dblTotalImpuesto = 0  'F4IGV
    dblTotalFacturado = 0 'F4MONTO
    dblTotalInafecto = 0 'F4MONINA
    dblTotalRedondeo = 0 'F4RND
    
    intEstado = 0 'F4ESTADO
    strCodigoRequerimiento = vbNullString 'F4CODSOLICITUD
    bolDocAnulado = False 'F4ESTNUL
    bolDocFalta = False 'F4FALTA
    bolOrdenConVale = False 'F4ESTVAL
    strReferencia = vbNullString 'F4REFERE
    dblCorrelativo = 0 'F4CORRELA
    bolAplicaDetraccion = False 'F4DETRACCIONAPLICA
    dblPorcentajeDetraccion = 0 'F4DETRACCIONPORC
    bolSinProveedorEspecifico = False 'F4SINPROVEEDORESP
    
    strFecReg = vbNullString 'F4FECGRA
    strUsuReg = vbNullString 'F4USEGRA
    strFecMod = vbNullString 'F4FECMOD
    strUsuMod = vbNullString 'F4USEMOD
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesDetalle()
    dblItem = 0 'ITEM
    strRequerimiento = vbNullString 'COD_SOLICITUD
    strCodProducto = vbNullString 'F3CODPRO
    strCodFabricante = vbNullString 'F3CODFAB
    strNomProducto = vbNullString 'F5NOMPRO
    strNomProductoInterno = vbNullString 'F5NOMPRO_ING
    strCodUM = vbNullString       'UNIDAD
    dblCantidad = 0 'F3CANPRO
    dblCantidadMaxima = 0 'F3CANPRO2
    dblPorcentajeDemasia = 0 'F3PORCDEMASIA
    dblPrecioSinImpuesto = 0 'F3PRECOS
    dblPrecioConImpuesto = 0 'F3PREUNI
    dblPrecioNetoSinImpuesto = 0   'F3PRENETO
    dblPorcentajeDscto = 0 'F3PORDCT
    dblTotalDscto = 0   'F3TOTDCT
    bolAfecto = False 'F5AFECTO
    dblImpuestoPorItem = 0   'F3IGV
    dblBasePorItem = 0  'F5VALVTA
    dblTotalPorItem = 0 'F3TOTAL
    strCodColor = vbNullString    'CODCOLOR
    strObservacionPorItem = vbNullString 'F3OBSERVA
    
    strCodigoGasto = vbNullString 'F3GASTO
    strCuentaContable = vbNullString 'F3CUENTA
    
    dblPorcentajeImpuesto = 0
    dblSignoImpuesto = 0
    dblCantidadFinal = 0
    dblExoneradoPorItem = 0
    
    dblCantidadFaltante = 0
    bolItemAjustado = False
End Sub

Public Function obtenerOrden() As Boolean
    On Error GoTo errObtenerOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        strTipoOrden = Trim(rstOrden!F4LOCAL & "") 'F4LOCAL
        strNumeroOrden = Trim(rstOrden!F4NUMORD & "") 'F4NUMORD
        
        strFechaEmision = Trim(rstOrden!F4FECEMI & "") 'F4FECEMI
        strCodProveedor = Trim(rstOrden!F4CODCLI & "") 'F4CODCLI
        strNomProveedor = Trim(rstOrden!F4NOMPROV & "") 'F4NOMPROV
        strRucProveedor = Trim(rstOrden!F4CODPRV & "") 'F4CODPRV
        strContactoProveedor = Trim(rstOrden!F4CONTACTO & "") 'F4CONTACTO
        strCodTipoComprobante = Trim(rstOrden!F4TIPDOC & "") 'F4TIPDOC
        bolOrdenRegularizada = IIf(Trim(rstOrden!F4REGULARIZA & "") = "1", True, False) 'F4REGULARIZA
        
        strFechaEntrega = Trim(rstOrden!F4FECENT & "") 'F4FECENT
        strCodigoSolicitante = Trim(rstOrden!F4CODSOL & "") 'F4CODSOL
        strCodFormaPago = Trim(rstOrden!F4FORPAG & "") 'F4FORPAG
        strCentroCosto = Trim(rstOrden!F4CENTRO & "") 'F4CENTRO
        strLugarEntrega = Trim(rstOrden!F4LUGAR_ENTREGA & "") 'F4LUGAR_ENTREGA
        bolPagoParcial = IIf(Val(rstOrden!F4PAGOPARCIAL & "") = 1, True, False) 'F4PAGOPARCIAL
        strCodMoneda = Trim(rstOrden!F4TIPMON & "") 'F4TIPMON
        dblTipoCambio = Val(rstOrden!F4TIPCAM & "") 'F4TIPCAM
        strFechaPago = Trim(rstOrden!F4DIAPAGO & "") 'F4DIAPAGO
        strNumeroCotizacion = Trim(rstOrden!F4NUMCOTIZA & "") 'F4NUMCOTIZA
        
        bolVB1 = CBool(rstOrden!F4VB1)      'F4VB1
        strVB1Usuario = Trim(rstOrden!F4VBUSER1 & "") 'F4VBUSER1
        strVB1Fecha = Trim(rstOrden!F4VBFECHA1 & "") 'F4VBFECHA1
        bolVB2 = CBool(rstOrden!F4VB2)      'F4VB2
        strVB2Usuario = Trim(rstOrden!F4VBUSER2 & "") 'F4VBUSER2
        strVB2Fecha = Trim(rstOrden!F4VBFECHA2 & "") 'F4VBFECHA2
        bolColocada = CBool(rstOrden!F4COLOCADA) 'F4COLOCADA
        strColocadaUsuario = Trim(rstOrden!F4COLOCADAUSER & "") 'F4COLOCADAUSER
        strColocadaFecha = Trim(rstOrden!F4COLOCADAFECHA & "")   'F4COLOCADAFECHA
        bolAtendida = CBool(rstOrden!F4ATENDIDA) 'F4ATENDIDA
        strAtendidaUsuario = Trim(rstOrden!F4ATENDIDAUSER & "") 'F4ATENDIDAUSER
        strAtendidaFecha = Trim(rstOrden!F4ATENDIDAFECHA & "")   'F4ATENDIDAFECHA
        
        strEmpresa = Trim(rstOrden!F4EMPRESA & "") 'F4EMPRESA
        strObservacion = Trim(rstOrden!F4OBSERVA & "") 'F4OBSERVA
        
        dblSubTotal = Val(rstOrden!F4BASIMP & "")  'F4BASIMP
        dblTotalImpuesto = Val(rstOrden!F4IGV & "")  'F4IGV
        dblTotalFacturado = Val(rstOrden!F4MONTO & "") 'F4MONTO
        dblTotalInafecto = Val(rstOrden!F4MONINA & "") 'F4MONINA
        dblTotalRedondeo = Val(rstOrden!F4RND & "") 'F4RND
        
        intEstado = Val(rstOrden!F4ESTADO & "") 'F4ESTADO
        strCodigoRequerimiento = Trim(rstOrden!F4CODSOLICITUD & "") 'F4CODSOLICITUD
        bolDocAnulado = IIf(Trim(rstOrden!F4ESTNUL & "") = "S", True, False) 'F4ESTNUL
        bolDocFalta = IIf(Trim(rstOrden!f4falta & "") = "1", True, False)  'F4FALTA
        bolOrdenConVale = IIf(Trim(rstOrden!F4ESTVAL & "") = "1", True, False)  'F4ESTVAL
        strReferencia = Trim(rstOrden!F4REFERE & "") 'F4REFERE
        dblCorrelativo = Val(rstOrden!F4CORRELA & "") 'F4CORRELA
        bolAplicaDetraccion = IIf(CBool(rstOrden!F4DETRACCIONAPLICA), True, False)   'F4DETRACCIONAPLICA
        dblPorcentajeDetraccion = Val(rstOrden!F4DETRACCIONPORC & "")  'F4DETRACCIONPORC
        bolSinProveedorEspecifico = IIf(CBool(rstOrden!F4SINPROVEEDORESP), True, False)    'F4SINPROVEEDORESP
        
        
        strFecReg = Trim(rstOrden!F4FECGRA & "") 'F4FECGRA
        strUsuReg = Trim(rstOrden!F4USEGRA & "") 'F4USEGRA
        strFecMod = Trim(rstOrden!F4FECMOD & "") 'F4FECMOD
        strUsuMod = Trim(rstOrden!F4USEMOD & "") 'F4USEMOD
        
        strResponsable = Trim(rstOrden!F4RESPONSABLE & "") 'F4RESPONSABLE
        
        obtenerOrdenDetalle
        
        obtenerOrden = True
    Else
        obtenerOrden = False
    End If
    
    rstOrden.Close
    
    Set rstOrden = Nothing
    
    Exit Function
errObtenerOrden:
    obtenerOrden = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerOrden"
    'Resume
    Err.Clear
End Function

Public Sub obtenerOrdenDetalle()
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA"
    
    abrirCnTemporal
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "INSERT INTO TMPORDENDECOMPRA("
        strSQLOrden = strSQLOrden & "ITEM, F3CODPRO, F3CODFAB, F5NOMPRO, F5NOMPRO_ING, "
        strSQLOrden = strSQLOrden & "F3CODMEDIDA, F3DESMEDIDA, F3CANPRO, F3CANPROMAX, CANT_ANT, "
        strSQLOrden = strSQLOrden & "F3PORCDEMASIA, F3CANPROFINAL, F3SINIGV, F3CONIGV, "
        strSQLOrden = strSQLOrden & "F3NETO, F5AFECTO, F3BASEIMP, F3MONINA, F3IGV, "
        strSQLOrden = strSQLOrden & "F3REDONDEO, F3TOTAL, F5CODCOSTO, F5DESCOSTO, "
        strSQLOrden = strSQLOrden & "F3VALDESC, F3PORDESC, F3OBSERVA, "
        strSQLOrden = strSQLOrden & "COD_SOLICITUD, CLIENTE, CODCOLOR, DESCOLOR, "
        strSQLOrden = strSQLOrden & "CERRAR, F3GASTO, F3CUENTA"
        strSQLOrden = strSQLOrden & ") "
        
    strSQLOrden = strSQLOrden & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLOrden = strSQLOrden & "SELECT "
        strSQLOrden = strSQLOrden & "DET.ITEM, "
        strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
        strSQLOrden = strSQLOrden & "DET.F3CODFAB, "
        strSQLOrden = strSQLOrden & "DET.F5NOMPRO, "
        strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
        strSQLOrden = strSQLOrden & "DET.UNIDAD, "
        strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
        strSQLOrden = strSQLOrden & "DET.F3CANPRO, "
        strSQLOrden = strSQLOrden & "DET.F3CANPRO2, "
        strSQLOrden = strSQLOrden & "DET.F3CANFAL, "
        strSQLOrden = strSQLOrden & "DET.F3PORCDEMASIA, "
        strSQLOrden = strSQLOrden & "(VAL(DET.F3CANPRO & '') * (1 + VAL(DET.F3PORCDEMASIA & '') / 100)) AS F3CANPROFINAL, "
        strSQLOrden = strSQLOrden & "DET.F3PRECOS, "
        strSQLOrden = strSQLOrden & "DET.F3PREUNI, "
        strSQLOrden = strSQLOrden & "DET.F3PRENETO, "
        strSQLOrden = strSQLOrden & "IIF(TRIM(DET.F5AFECTO & '') = '*', -1, 0) AS AFECTO, "
        strSQLOrden = strSQLOrden & "IIF(TRIM(DET.F5AFECTO & '') = '*', DET.F5VALVTA, 0) AS BASE,"
        strSQLOrden = strSQLOrden & "IIF(TRIM(DET.F5AFECTO & '') = '*', 0, DET.F5VALVTA) AS INAFECTO,"
        strSQLOrden = strSQLOrden & "DET.F3IGV, "
        strSQLOrden = strSQLOrden & "0 AS REDONDEO, "
        strSQLOrden = strSQLOrden & "DET.F3TOTAL, "
        strSQLOrden = strSQLOrden & "DET.F5CODCOSTO, "
        strSQLOrden = strSQLOrden & "DET.F5DESCOSTO, "
        strSQLOrden = strSQLOrden & "DET.F3TOTDCT, "
        strSQLOrden = strSQLOrden & "DET.F3PORDCT, "
        strSQLOrden = strSQLOrden & "DET.F3OBSERVA, "
        strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
        strSQLOrden = strSQLOrden & "REQ.CS_NOMREF, "
        strSQLOrden = strSQLOrden & "DET.CODCOLOR, "
        strSQLOrden = strSQLOrden & "COL.DESCRIPCION, "
        strSQLOrden = strSQLOrden & "DET.F3AJUSTE, "
        strSQLOrden = strSQLOrden & "DET.F3GASTO, "
        strSQLOrden = strSQLOrden & "DET.F3CUENTA "
        
    strSQLOrden = strSQLOrden & "FROM "
        strSQLOrden = strSQLOrden & "(((IF3ORDEN AS DET "
        strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
        strSQLOrden = strSQLOrden & "LEFT JOIN EF2BIENCOLOR AS COL ON COL.CODIGO = DET.CODCOLOR) "
        strSQLOrden = strSQLOrden & "LEFT JOIN TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD "
    strSQLOrden = strSQLOrden & "WHERE "
            strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
            strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    strSQLOrden = vbNullString
End Sub


Public Sub obtenerConfigOrden()
    On Error GoTo errObtenerConfigOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrden.EOF Then
        strTipoOrden = Trim(rstOrden!F4LOCAL & "") 'F4LOCAL
        strNumeroOrden = Trim(rstOrden!F4NUMORD & "") 'F4NUMORD
        
        strFechaEmision = Trim(rstOrden!F4FECEMI & "") 'F4FECEMI
        strCodProveedor = Trim(rstOrden!F4CODCLI & "") 'F4CODCLI
        strNomProveedor = Trim(rstOrden!F4NOMPROV & "") 'F4NOMPROV
        strRucProveedor = Trim(rstOrden!F4CODPRV & "") 'F4CODPRV
        strContactoProveedor = Trim(rstOrden!F4CONTACTO & "") 'F4CONTACTO
        strCodTipoComprobante = Trim(rstOrden!F4TIPDOC & "") 'F4TIPDOC
        bolOrdenRegularizada = IIf(Trim(rstOrden!F4REGULARIZA & "") = "1", True, False) 'F4REGULARIZA
        
        strFechaEntrega = Trim(rstOrden!F4FECENT & "") 'F4FECENT
        strCodigoSolicitante = Trim(rstOrden!F4CODSOL & "") 'F4CODSOL
        strCodFormaPago = Trim(rstOrden!F4FORPAG & "") 'F4FORPAG
        strCentroCosto = Trim(rstOrden!F4CENTRO & "") 'F4CENTRO
        strLugarEntrega = Trim(rstOrden!F4LUGAR_ENTREGA & "") 'F4LUGAR_ENTREGA
        bolPagoParcial = IIf(Val(rstOrden!F4PAGOPARCIAL & "") = 1, True, False) 'F4PAGOPARCIAL
        strCodMoneda = Trim(rstOrden!F4TIPMON & "") 'F4TIPMON
        dblTipoCambio = Val(rstOrden!F4TIPCAM & "") 'F4TIPCAM
        strFechaPago = Trim(rstOrden!F4DIAPAGO & "") 'F4DIAPAGO
        strNumeroCotizacion = Trim(rstOrden!F4NUMCOTIZA & "") 'F4NUMCOTIZA
        
        bolVB1 = CBool(rstOrden!F4VB1)      'F4VB1
        strVB1Usuario = Trim(rstOrden!F4VBUSER1 & "") 'F4VBUSER1
        strVB1Fecha = Trim(rstOrden!F4VBFECHA1 & "") 'F4VBFECHA1
        bolVB2 = CBool(rstOrden!F4VB2)      'F4VB2
        strVB2Usuario = Trim(rstOrden!F4VBUSER2 & "") 'F4VBUSER2
        strVB2Fecha = Trim(rstOrden!F4VBFECHA2 & "") 'F4VBFECHA2
        bolColocada = CBool(rstOrden!F4COLOCADA) 'F4COLOCADA
        strColocadaUsuario = Trim(rstOrden!F4COLOCADAUSER & "") 'F4COLOCADAUSER
        strColocadaFecha = Trim(rstOrden!F4COLOCADAFECHA & "")   'F4COLOCADAFECHA
        bolAtendida = CBool(rstOrden!F4ATENDIDA) 'F4ATENDIDA
        strAtendidaUsuario = Trim(rstOrden!F4ATENDIDAUSER & "") 'F4ATENDIDAUSER
        strAtendidaFecha = Trim(rstOrden!F4ATENDIDAFECHA & "")   'F4ATENDIDAFECHA
        
        strEmpresa = Trim(rstOrden!F4EMPRESA & "") 'F4EMPRESA
        strObservacion = Trim(rstOrden!F4OBSERVA & "") 'F4OBSERVA
        
        dblSubTotal = Val(rstOrden!F4BASIMP & "")  'F4BASIMP
        dblTotalImpuesto = Val(rstOrden!F4IGV & "")  'F4IGV
        dblTotalFacturado = Val(rstOrden!F4MONTO & "") 'F4MONTO
        dblTotalInafecto = Val(rstOrden!F4MONINA & "") 'F4MONINA
        dblTotalRedondeo = Val(rstOrden!F4RND & "") 'F4RND
        
        intEstado = Val(rstOrden!F4ESTADO & "") 'F4ESTADO
        strCodigoRequerimiento = Trim(rstOrden!F4CODSOLICITUD & "") 'F4CODSOLICITUD
        bolDocAnulado = IIf(Trim(rstOrden!F4ESTNUL & "") = "S", True, False) 'F4ESTNUL
        bolDocFalta = IIf(Trim(rstOrden!f4falta & "") = "1", True, False)  'F4FALTA
        bolOrdenConVale = IIf(Trim(rstOrden!F4ESTVAL & "") = "1", True, False)  'F4ESTVAL
        strReferencia = Trim(rstOrden!F4REFERE & "") 'F4REFERE
        dblCorrelativo = Val(rstOrden!F4CORRELA & "") 'F4CORRELA
        bolAplicaDetraccion = IIf(CBool(rstOrden!F4DETRACCIONAPLICA), True, False)   'F4DETRACCIONAPLICA
        dblPorcentajeDetraccion = Val(rstOrden!F4DETRACCIONPORC & "")  'F4DETRACCIONPORC
        bolSinProveedorEspecifico = IIf(CBool(rstOrden!F4SINPROVEEDORESP), True, False) 'F4SINPROVEEDORESP
        
        
        strFecReg = Trim(rstOrden!F4FECGRA & "") 'F4FECGRA
        strUsuReg = Trim(rstOrden!F4USEGRA & "") 'F4USEGRA
        strFecMod = Trim(rstOrden!F4FECMOD & "") 'F4FECMOD
        strUsuMod = Trim(rstOrden!F4USEMOD & "") 'F4USEMOD
    End If
    
    rstOrden.Close
    
    Set rstOrden = Nothing
    
    Exit Sub
errObtenerConfigOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerConfigOrden"
    
    Err.Clear
End Sub

Public Sub obtenerConfigOrdenDetalleOnebyOne()
    On Error GoTo errObtenerConfigOrdenDetalleOnebyOne
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF3ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "F3CODPRO = '" & strCodProducto & "' AND "
    strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & strRequerimiento & "'"
    
        If dblItem <> 0 Then
            strSQLOrden = strSQLOrden & " AND ITEM = '" & dblItem & "'"
        End If
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrden.EOF Then
        strTipoOrden = Trim(rstOrden!F4LOCAL & "") 'F4LOCAL
        strNumeroOrden = Trim(rstOrden!F4NUMORD & "") 'F4NUMORD
        
        dblItem = Val(rstOrden!ITEM & "") 'ITEM
        strRequerimiento = Trim(rstOrden!COD_SOLICITUD & "") 'COD_SOLICITUD
        strCodProducto = Trim(rstOrden!F3CODPRO & "") 'F3CODPRO
        strCodFabricante = Trim(rstOrden!F3CODFAB & "") 'F3CODFAB
        strNomProducto = Trim(rstOrden!F5NOMPRO & "") 'F5NOMPRO
        strNomProductoInterno = Trim(rstOrden!F5NOMPRO_ING & "") 'F5NOMPRO_ING
        strCodUM = Trim(rstOrden!UNIDAD & "")       'UNIDAD
        dblCantidad = Val(rstOrden!F3CANPRO & "") 'F3CANPRO
        dblCantidadMaxima = Val(rstOrden!F3CANPRO2 & "") 'F3CANPRO2
        dblPorcentajeDemasia = Val(rstOrden!F3PORCDEMASIA & "") 'F3PORCDEMASIA
        dblPrecioSinImpuesto = Val(rstOrden!F3PRECOS & "") 'F3PRECOS
        dblPrecioConImpuesto = Val(rstOrden!F3PREUNI & "") 'F3PREUNI
        dblPrecioNetoSinImpuesto = Val(rstOrden!F3PRENETO & "")   'F3PRENETO
        dblPorcentajeDscto = Val(rstOrden!F3PORDCT & "") 'F3PORDCT
        dblTotalDscto = Val(rstOrden!F3TOTDCT & "")   'F3TOTDCT
        bolAfecto = IIf(Trim(rstOrden!F5AFECTO & "") = "*", True, False)  'F5AFECTO
        dblImpuestoPorItem = Val(rstOrden!F3IGV & "")   'F3IGV
        dblBasePorItem = Val(rstOrden!F5VALVTA & "")  'F5VALVTA
        dblTotalPorItem = Val(rstOrden!F3TOTAL & "") 'F3TOTAL
        strCodColor = Trim(rstOrden!CODCOLOR & "")    'CODCOLOR
        strObservacionPorItem = Trim(rstOrden!F3OBSERVA & "") 'F3OBSERVA
        
        dblCantidadFaltante = Val(rstOrden!F3CANFAL & "") 'F3CANFAL
        bolItemAjustado = CBool(rstOrden!F3AJUSTE & "") 'F3AJUSTE
    End If
    
    rstOrden.Close
    
    Set rstOrden = Nothing
    
    Exit Sub
errObtenerConfigOrdenDetalleOnebyOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerConfigOrdenDetalleOnebyOne"
    
    Err.Clear
End Sub

Public Function generarNumeroOrden() As String
    On Error GoTo errGenerarNumeroOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT TOP 1 "
    strSQLOrden = strSQLOrden & "RIGHT(F4NUMORD,9) AS CODIGO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' "
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "F4NUMORD DESC"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        generarNumeroOrden = strTipoOrden & Format(Val(Trim(rstOrden!Codigo & "")) + 1, "000000000")
    Else
        generarNumeroOrden = strTipoOrden & "000000001"
    End If

    rstOrden.Close

    Set rstOrden = Nothing
    strSQLOrden = vbNullString

    Exit Function
errGenerarNumeroOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: GenerarNumeroOrden"
    
    generarNumeroOrden = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = "SELECT IF4ORDEN.* " & _
                    "FROM IF4ORDEN " & _
                    "WHERE F4LOCAL = '" & strTipoOrden & "' AND " & _
                            "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close

    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If

    rstOrden.Close
    strSQLOrden = vbNullString

    Set rstOrden = Nothing

    Exit Function
errVerificarExistencia:
    verificarExistencia = False

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: VerificarExistencia"

    Err.Clear
End Function

Public Function guardarOrden() As Boolean
    
    strSQLOrden = vbNullString
    
    If Not verificarExistencia Then
        If strNumeroOrden = vbNullString Then
            strNumeroOrden = generarNumeroOrden
        End If
        
        strSQLOrden = strSQLOrden & "INSERT INTO IF4ORDEN("
        strSQLOrden = strSQLOrden & "F4LOCAL, F4NUMORD, F4FECEMI, F4CODCLI, F4NOMPROV, F4CODPRV, "
        strSQLOrden = strSQLOrden & "F4CONTACTO, F4TIPDOC, F4REGULARIZA, F4FECENT, F4CODSOL, "
        strSQLOrden = strSQLOrden & "F4FORPAG, F4CENTRO, F4LUGAR_ENTREGA, F4PAGOPARCIAL, F4TIPMON, "
        strSQLOrden = strSQLOrden & "F4TIPCAM, F4DIAPAGO, F4NUMCOTIZA, F4VB1, F4VBUSER1, F4VBFECHA1, "
        strSQLOrden = strSQLOrden & "F4VB2, F4VBUSER2, F4VBFECHA2, F4COLOCADA, F4COLOCADAUSER, "
        strSQLOrden = strSQLOrden & "F4COLOCADAFECHA, F4ATENDIDA, F4ATENDIDAUSER, F4ATENDIDAFECHA, "
        strSQLOrden = strSQLOrden & "F4EMPRESA, F4OBSERVA, F4BASIMP, F4IGV, F4MONTO, F4MONINA, F4RND, "
        strSQLOrden = strSQLOrden & "F4ESTADO, F4CODSOLICITUD, F4ESTNUL, F4FALTA, F4ESTVAL, F4REFERE, "
        strSQLOrden = strSQLOrden & "F4CORRELA, F4DETRACCIONAPLICA, F4DETRACCIONPORC, F4SINPROVEEDORESP, "
        strSQLOrden = strSQLOrden & "F4FECGRA, F4USEGRA, F4RESPONSABLE) "
        
        strSQLOrden = strSQLOrden & "VALUES("
        strSQLOrden = strSQLOrden & "'" & strTipoOrden & "', '" & strNumeroOrden & "', CVDATE('" & strFechaEmision & "'), "
        strSQLOrden = strSQLOrden & "'" & strCodProveedor & "', '" & strNomProveedor & "', '" & strRucProveedor & "', '" & strContactoProveedor & "', "
        strSQLOrden = strSQLOrden & "'" & strCodTipoComprobante & "', '" & IIf(bolOrdenRegularizada, "1", "0") & "', CVDATE('" & strFechaEntrega & "'), "
        strSQLOrden = strSQLOrden & "'" & strCodigoSolicitante & "', '" & strCodFormaPago & "', '" & strCentroCosto & "', '" & strLugarEntrega & "', "
        strSQLOrden = strSQLOrden & IIf(bolPagoParcial, "1", "0") & ", '" & strCodMoneda & "', " & dblTipoCambio & ", '" & strFechaPago & "', "
        strSQLOrden = strSQLOrden & "'" & strNumeroCotizacion & "', " & IIf(bolVB1, "TRUE", "FALSE") & ", " & IIf(strVB1Usuario <> vbNullString, "'" & strVB1Usuario & "'", "NULL") & ", " & IIf(strVB1Fecha <> vbNullString, "CVDATE('" & strVB1Fecha & "')", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolVB2, "TRUE", "FALSE") & ", " & IIf(strVB2Usuario <> vbNullString, "'" & strVB2Usuario & "'", "NULL") & ", " & IIf(strVB2Fecha <> vbNullString, "CVDATE('" & strVB2Fecha & "')", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolColocada, "TRUE", "FALSE") & ", " & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "NULL") & ", " & IIf(strColocadaFecha <> vbNullString, "CVDATE('" & strColocadaFecha & "')", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolAtendida, "TRUE", "FALSE") & ", " & IIf(strAtendidaUsuario <> vbNullString, "'" & strAtendidaUsuario & "'", "NULL") & ", " & IIf(strAtendidaFecha <> vbNullString, "CVDATE('" & strAtendidaFecha & "')", "NULL") & ", "
        strSQLOrden = strSQLOrden & "'" & strEmpresa & "', '" & strObservacion & "', " & dblSubTotal & ", " & dblTotalImpuesto & ", "
        strSQLOrden = strSQLOrden & dblTotalFacturado & ", " & dblTotalInafecto & ", " & dblTotalRedondeo & ", "
        strSQLOrden = strSQLOrden & intEstado & ", '" & strCodigoRequerimiento & "', '" & IIf(bolDocAnulado, "S", "N") & "', "
        strSQLOrden = strSQLOrden & "'" & IIf(bolDocFalta, "1", "0") & "', '" & IIf(bolOrdenConVale, "1", "0") & "', '" & strReferencia & "', "
        strSQLOrden = strSQLOrden & dblCorrelativo & ", " & IIf(bolAplicaDetraccion, "1", "0") & ", " & dblPorcentajeDetraccion & ", " & IIf(bolSinProveedorEspecifico, "1", "0") & ", "
        strSQLOrden = strSQLOrden & "CVDATE('" & strFecReg & "'), '" & strUsuReg & "', '" & strResponsable & "')"
    Else
        strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF4ORDEN "
        strSQLOrden = strSQLOrden & "SET "
        
            strSQLOrden = strSQLOrden & "F4FECEMI = CVDATE('" & strFechaEmision & "'), "
            strSQLOrden = strSQLOrden & "F4CODCLI = '" & strCodProveedor & "', "
            strSQLOrden = strSQLOrden & "F4NOMPROV = '" & strNomProveedor & "', "
            strSQLOrden = strSQLOrden & "F4CODPRV = '" & strRucProveedor & "', "
            strSQLOrden = strSQLOrden & "F4CONTACTO = '" & strContactoProveedor & "', "
            strSQLOrden = strSQLOrden & "F4TIPDOC = '" & strCodTipoComprobante & "', "
            strSQLOrden = strSQLOrden & "F4REGULARIZA = '" & IIf(bolOrdenRegularizada, "1", "0") & "', "
            strSQLOrden = strSQLOrden & "F4FECENT = CVDATE('" & strFechaEntrega & "'), "
            strSQLOrden = strSQLOrden & "F4CODSOL = '" & strCodigoSolicitante & "', "
            strSQLOrden = strSQLOrden & "F4FORPAG = '" & strCodFormaPago & "', "
            strSQLOrden = strSQLOrden & "F4CENTRO = '" & strCentroCosto & "', "
            strSQLOrden = strSQLOrden & "F4LUGAR_ENTREGA = '" & strLugarEntrega & "', "
            strSQLOrden = strSQLOrden & "F4PAGOPARCIAL = " & IIf(bolPagoParcial, "1", "0") & ", "
            strSQLOrden = strSQLOrden & "F4TIPMON = '" & strCodMoneda & "', "
            strSQLOrden = strSQLOrden & "F4TIPCAM = " & dblTipoCambio & ", "
            strSQLOrden = strSQLOrden & "F4DIAPAGO = '" & strFechaPago & "', "
            strSQLOrden = strSQLOrden & "F4NUMCOTIZA = '" & strNumeroCotizacion & "', "
            'strSQLOrden = strSQLOrden & "F4VB1 = " & IIf(bolVB1, "TRUE", "FALSE") & ", "
            'strSQLOrden = strSQLOrden & "F4VBUSER1 = '" & strVB1Usuario & "', "
            'strSQLOrden = strSQLOrden & "F4VBFECHA1 = CVDATE('" & strVB1Fecha & "'), "
            'strSQLOrden = strSQLOrden & "F4VB2 = " & IIf(bolVB2, "TRUE", "FALSE") & ", "
            'strSQLOrden = strSQLOrden & "F4VBUSER2 = '" & strVB2Usuario & "', "
            'strSQLOrden = strSQLOrden & "F4VBFECHA2 = CVDATE('" & strVB2Fecha & "'), "
            strSQLOrden = strSQLOrden & "F4COLOCADA = " & IIf(bolColocada, "TRUE", "FALSE") & ", "
            strSQLOrden = strSQLOrden & "F4COLOCADAUSER = " & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "NULL") & ", "
            strSQLOrden = strSQLOrden & "F4COLOCADAFECHA = " & IIf(strColocadaFecha <> vbNullString, "CVDATE('" & strColocadaFecha & "')", "NULL") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDA = " & IIf(bolAtendida, "TRUE", "FALSE") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDAUSER = " & IIf(strAtendidaUsuario <> vbNullString, "'" & strAtendidaUsuario & "'", "NULL") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDAFECHA = " & IIf(strAtendidaFecha <> vbNullString, "CVDATE('" & strAtendidaFecha & "')", "NULL") & ", "
            strSQLOrden = strSQLOrden & "F4EMPRESA = '" & strEmpresa & "', "
            strSQLOrden = strSQLOrden & "F4OBSERVA = '" & strObservacion & "', "
            strSQLOrden = strSQLOrden & "F4BASIMP = " & dblSubTotal & ", "
            strSQLOrden = strSQLOrden & "F4IGV = " & dblTotalImpuesto & ", "
            strSQLOrden = strSQLOrden & "F4MONTO = " & dblTotalFacturado & ", "
            strSQLOrden = strSQLOrden & "F4MONINA = " & dblTotalInafecto & ", "
            strSQLOrden = strSQLOrden & "F4RND = " & dblTotalRedondeo & ", "
                
                If intEstado > 0 Then
                    strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & ", "
                End If
                
            strSQLOrden = strSQLOrden & "F4CODSOLICITUD = '" & strCodigoRequerimiento & "', "
            strSQLOrden = strSQLOrden & "F4ESTNUL = '" & IIf(bolDocAnulado, "S", "N") & "', "
            strSQLOrden = strSQLOrden & "F4FALTA = '" & IIf(bolDocFalta, "1", "0") & "', "
            strSQLOrden = strSQLOrden & "F4ESTVAL = '" & IIf(bolOrdenConVale, "1", "0") & "', "
            strSQLOrden = strSQLOrden & "F4REFERE = '" & strReferencia & "', "
            strSQLOrden = strSQLOrden & "F4CORRELA = " & dblCorrelativo & ", "
            strSQLOrden = strSQLOrden & "F4DETRACCIONAPLICA = " & IIf(bolAplicaDetraccion, "1", "0") & ", "
            strSQLOrden = strSQLOrden & "F4DETRACCIONPORC = " & dblPorcentajeDetraccion & ", "
            strSQLOrden = strSQLOrden & "F4SINPROVEEDORESP = " & IIf(bolSinProveedorEspecifico, "1", "0") & ", "
            strSQLOrden = strSQLOrden & "F4FECMOD = CVDATE('" & strFecMod & "'), "
            strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "', "
            strSQLOrden = strSQLOrden & "F4RESPONSABLE = '" & strResponsable & "' "
            
        strSQLOrden = strSQLOrden & "WHERE "
            strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
            strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    End If
    
    cnn_dbbancos.Execute strSQLOrden
    
    guardarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    
End Function

Public Sub guardarOrdenDetalleOneByOne()
    On Error GoTo errGuardarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "INSERT INTO IF3ORDEN("
    strSQLOrden = strSQLOrden & "F4LOCAL, F4NUMORD, ITEM, "
    strSQLOrden = strSQLOrden & "COD_SOLICITUD, F3CODPRO, F3CODFAB, "
    strSQLOrden = strSQLOrden & "F5NOMPRO, F5NOMPRO_ING, UNIDAD, "
    strSQLOrden = strSQLOrden & "F3CANPRO, F3CANPRO2, F3CANFAL, F3PORCDEMASIA, "
    strSQLOrden = strSQLOrden & "F3PRECOS, F3PREUNI, F3PRENETO, "
    strSQLOrden = strSQLOrden & "F3PORDCT, F3TOTDCT, F5AFECTO, "
    strSQLOrden = strSQLOrden & "F5VALVTA, F3IGV, F3TOTAL, "
    strSQLOrden = strSQLOrden & "CODCOLOR, F3OBSERVA, F3AJUSTE, "
    strSQLOrden = strSQLOrden & "F3GASTO, F3CUENTA) "
    
    strSQLOrden = strSQLOrden & "VALUES("
    strSQLOrden = strSQLOrden & "'" & strTipoOrden & "', "
    strSQLOrden = strSQLOrden & "'" & strNumeroOrden & "', "
    strSQLOrden = strSQLOrden & dblItem & ", "
    strSQLOrden = strSQLOrden & "'" & strRequerimiento & "', "
    strSQLOrden = strSQLOrden & "'" & strCodProducto & "', "
    strSQLOrden = strSQLOrden & "'" & strCodFabricante & "', "
    strSQLOrden = strSQLOrden & "'" & strNomProducto & "', "
    strSQLOrden = strSQLOrden & "'" & strNomProductoInterno & "', "
    strSQLOrden = strSQLOrden & "'" & strCodUM & "', "
    strSQLOrden = strSQLOrden & dblCantidad & ", "
    strSQLOrden = strSQLOrden & dblCantidadMaxima & ", "
    strSQLOrden = strSQLOrden & dblCantidadFaltante & ", "
    strSQLOrden = strSQLOrden & dblPorcentajeDemasia & ", "
    strSQLOrden = strSQLOrden & dblPrecioSinImpuesto & ", "
    strSQLOrden = strSQLOrden & dblPrecioConImpuesto & ", "
    strSQLOrden = strSQLOrden & dblPrecioNetoSinImpuesto & ", "
    strSQLOrden = strSQLOrden & dblPorcentajeDscto & ", "
    strSQLOrden = strSQLOrden & dblTotalDscto & ", "
    strSQLOrden = strSQLOrden & IIf(bolAfecto, "'*'", "NULL") & ", "
    strSQLOrden = strSQLOrden & dblBasePorItem & ", "
    strSQLOrden = strSQLOrden & dblImpuestoPorItem & ", "
    strSQLOrden = strSQLOrden & dblTotalPorItem & ", "
    strSQLOrden = strSQLOrden & "'" & strCodColor & "', "
    strSQLOrden = strSQLOrden & "'" & strObservacionPorItem & "', "
    strSQLOrden = strSQLOrden & IIf(bolItemAjustado, "1", "0") & ", "
    strSQLOrden = strSQLOrden & IIf(strCodigoGasto <> vbNullString, "'" & strCodigoGasto & "'", "NULL") & ", "
    strSQLOrden = strSQLOrden & IIf(strCuentaContable <> vbNullString, "'" & strCuentaContable & "'", "NULL")
    strSQLOrden = strSQLOrden & ")"
    
    cnn_dbbancos.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errGuardarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: GuardarOrdenDetalleOneByOne"
    
    Err.Clear
End Sub

Public Sub actualizarOrdenDetalleOneByOne()
    On Error GoTo errActualizarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "UPDATE "
    strSQLOrden = strSQLOrden & "IF3ORDEN "
    strSQLOrden = strSQLOrden & "SET "
    strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & strRequerimiento & "', "
    strSQLOrden = strSQLOrden & "F3CODPRO = '" & strCodProducto & "', "
    strSQLOrden = strSQLOrden & "F3CODFAB = '" & strCodFabricante & "', "
    strSQLOrden = strSQLOrden & "F5NOMPRO = '" & strNomProducto & "', "
    strSQLOrden = strSQLOrden & "F5NOMPRO_ING = '" & strNomProductoInterno & "', "
    strSQLOrden = strSQLOrden & "UNIDAD = '" & strCodUM & "', "
    strSQLOrden = strSQLOrden & "F3CANPRO = " & dblCantidad & ", "
    strSQLOrden = strSQLOrden & "F3CANPRO2 = " & dblCantidadMaxima & ", "
    strSQLOrden = strSQLOrden & "F3CANFAL = " & dblCantidadFaltante & ", "
    strSQLOrden = strSQLOrden & "F3PORCDEMASIA = " & dblPorcentajeDemasia & ", "
    strSQLOrden = strSQLOrden & "F3PRECOS = " & dblPrecioSinImpuesto & ", "
    strSQLOrden = strSQLOrden & "F3PREUNI = " & dblPrecioConImpuesto & ", "
    strSQLOrden = strSQLOrden & "F3PRENETO = " & dblPrecioNetoSinImpuesto & ", "
    strSQLOrden = strSQLOrden & "F3PORDCT = " & dblPorcentajeDscto & ", "
    strSQLOrden = strSQLOrden & "F3TOTDCT = " & dblTotalDscto & ", "
    strSQLOrden = strSQLOrden & "F5AFECTO = " & IIf(bolAfecto, "'*'", "NULL") & ", "
    strSQLOrden = strSQLOrden & "F5VALVTA = " & dblBasePorItem & ", "
    strSQLOrden = strSQLOrden & "F3IGV = " & dblImpuestoPorItem & ", "
    strSQLOrden = strSQLOrden & "F3TOTAL = " & dblTotalPorItem & ", "
    strSQLOrden = strSQLOrden & "CODCOLOR = '" & strCodColor & "', "
    strSQLOrden = strSQLOrden & "F3OBSERVA = '" & strObservacionPorItem & "', "
    strSQLOrden = strSQLOrden & "F3AJUSTE = " & IIf(bolItemAjustado, "1", "0") & ", "
    strSQLOrden = strSQLOrden & "F3GASTO = " & IIf(strCodigoGasto <> vbNullString, "'" & strCodigoGasto & "'", "NULL") & ", "
    strSQLOrden = strSQLOrden & "F3CUENTA = " & IIf(strCuentaContable <> vbNullString, "'" & strCuentaContable & "'", "NULL") & " "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "ITEM = '" & dblItem & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errActualizarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ActualizarOrdenDetalleOneByOne"
    
    Err.Clear
End Sub

Public Sub eliminarOrdenDetalleOneByOne()
    On Error GoTo errEliminarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "DELETE FROM "
    strSQLOrden = strSQLOrden & "IF3ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "ITEM = '" & dblItem & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errEliminarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: EliminarOrdenDetalleOneByOne"
    
    Err.Clear
End Sub

Public Sub cerrarOrdenDetalleOneByOne()
    On Error GoTo errCerrarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF3ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F3AJUSTE = 1, "
        strSQLOrden = strSQLOrden & "F3PORCDEMASIA = 0, "
        strSQLOrden = strSQLOrden & "F3CANFAL = " & dblCantidadFaltante & ", "
        strSQLOrden = strSQLOrden & "F3CANPRO2 = " & dblCantidadMaxima & ", "
        strSQLOrden = strSQLOrden & "F3CANPRO = " & dblCantidad & " "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
        strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & strRequerimiento & "' AND "
        strSQLOrden = strSQLOrden & "F3CODPRO = '" & strCodProducto & "'"
            
            If dblItem <> 0 Then
                strSQLOrden = strSQLOrden & " AND ITEM = '" & dblItem & "'"
            End If
    
    cnn_dbbancos.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errCerrarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: CerrarOrdenDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    strSQLOrden = vbNullString
    
    Err.Clear
End Sub

Public Function eliminarOrden() As Boolean
    On Error GoTo errEliminarOrden
    
'    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F2CODALM", "IF3ORDEN", strNumeroVale, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarOrden = False
'
'        Exit Function
'    End If
    
    ''cnn_dbbancos.BeginTrans
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "DELETE "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    ''cnn_dbbancos.CommitTrans
    
    eliminarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errEliminarOrden:
    eliminarOrden = False
    
    ''cnn_dbbancos.RollbackTrans
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: EliminarOrden"
    
    Err.Clear
End Function

Public Function anularOrden() As Boolean
    On Error GoTo errAnularOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = 8, "
        strSQLOrden = strSQLOrden & "F4ESTNUL = 'S', "
        strSQLOrden = strSQLOrden & "F4BASIMP = 0, "
        strSQLOrden = strSQLOrden & "F4MONINA = 0, "
        strSQLOrden = strSQLOrden & "F4IGV = 0, "
        strSQLOrden = strSQLOrden & "F4MONTO = 0, "
        strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "', "
        strSQLOrden = strSQLOrden & "F4FECMOD = CVDATE('" & strFecMod & "') "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    Actualiza_Log strSQLOrden, StrConexDbBancos
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF3ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F3CANPRO = 0, "
        strSQLOrden = strSQLOrden & "F3CANPRO2 = 0, "
        strSQLOrden = strSQLOrden & "F3PORCDEMASIA = 0, "
        strSQLOrden = strSQLOrden & "F3PRECOS = 0, "
        strSQLOrden = strSQLOrden & "F3PREUNI = 0, "
        strSQLOrden = strSQLOrden & "F3PRENETO = 0, "
        strSQLOrden = strSQLOrden & "F3IGV = 0, "
        strSQLOrden = strSQLOrden & "F3TOTAL = 0, "
        strSQLOrden = strSQLOrden & "F3PORDCT = 0, "
        strSQLOrden = strSQLOrden & "F3TOTDCT = 0, "
        strSQLOrden = strSQLOrden & "F5VALVTA = 0 "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    Actualiza_Log strSQLOrden, StrConexDbBancos
    
    anularOrden = True
    
    strSQLOrden = vbNullString
    
    Exit Function
errAnularOrden:
    anularOrden = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: AnularOrden"
    
    Err.Clear
End Function

Public Function aprobarOrden() As Boolean
    On Error GoTo errAprobarOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & ", "
        strSQLOrden = strSQLOrden & "F4VB1 = " & IIf(bolVB1, "TRUE", "FALSE") & ", "
        strSQLOrden = strSQLOrden & "F4VBUSER1 = " & IIf(strVB1Usuario <> vbNullString, "'" & strVB1Usuario & "'", "NULL") & ", "
        strSQLOrden = strSQLOrden & "F4VBFECHA1 = " & IIf(strVB1Fecha <> vbNullString, "CVDATE('" & strVB1Fecha & "')", "NULL") & " "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    Actualiza_Log strSQLOrden, StrConexDbBancos
    
    aprobarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errAprobarOrden:
    aprobarOrden = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: AprobarOrden"
    
    Err.Clear
End Function

Public Function enviarViaMailOrden() As Boolean
    On Error GoTo errEnviarViaMailOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADA = " & IIf(bolColocada, "TRUE", "FALSE") & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADAUSER = " & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "NULL") & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADAFECHA = " & IIf(strColocadaFecha <> vbNullString, "CVDATE('" & strColocadaFecha & "')", "NULL") & " "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOrden
    
    Actualiza_Log strSQLOrden, StrConexDbBancos
    
    enviarViaMailOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errEnviarViaMailOrden:
    enviarViaMailOrden = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: EnviarViaMailOrden"
    
    Err.Clear
End Function

Public Function atencionOrden() As Boolean
    On Error GoTo errAtencionOrden
    
    Set rstOrden = New ADODB.Recordset
    
    Dim dblCantItems As Double
    Dim dblCantItemsNoAtendidos As Double
    Dim dblCantItemsAtendidoParcial As Double
    Dim dblCantItemsAtendidoTotal As Double
    
    atencionOrden = False
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS SALDO "
    
    strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) , '#0.0000')) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) , '#0.0000')) - VAL(FORMAT( (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR))  , '#0.0000')) AS SALDO "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON RESUMEN.F4NUMORD = DET.F4NUMORD AND RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
    
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) <> 0 " 'AND "
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) > 0 "
    
    strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) , '#0.0000')) > 0 "
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        rstOrden.MoveFirst
        
        dblCantItems = ModUtilitario.devuelveCantRegistros(rstOrden)
        dblCantItemsNoAtendidos = 0
        dblCantItemsAtendidoParcial = 0
        dblCantItemsAtendidoTotal = 0
        
        'intEstado = 0
        
        Do While Not rstOrden.EOF
'            If Val(rstOrden!SALDO & "") = Val(rstOrden!Cantidad & "") And (intEstado <> 5 And intEstado <> 6) Then
'                intEstado = 4
'            ElseIf Val(rstOrden!SALDO & "") < Val(rstOrden!Cantidad & "") And Val(rstOrden!SALDO & "") > 0 Then
'                intEstado = 5
'            ElseIf Val(rstOrden!SALDO & "") = 0 And intEstado <> 5 Then
'                intEstado = 6
'            End If
            
            If Val(rstOrden!SALDO & "") = Val(rstOrden!Cantidad & "") Then
                dblCantItemsNoAtendidos = dblCantItemsNoAtendidos + 1
            ElseIf Val(rstOrden!SALDO & "") < Val(rstOrden!Cantidad & "") And Val(rstOrden!SALDO & "") > 0 Then
                dblCantItemsAtendidoParcial = dblCantItemsAtendidoParcial + 1
            ElseIf Val(rstOrden!SALDO & "") <= 0 Then
                dblCantItemsAtendidoTotal = dblCantItemsAtendidoTotal + 1
            End If
            
            rstOrden.MoveNext
        Loop
            If dblCantItemsNoAtendidos = dblCantItems Then
                If intEstado >= 4 Then
                    intEstado = 4
                End If
            ElseIf dblCantItemsAtendidoTotal = dblCantItems Then
                intEstado = 6
            Else
                intEstado = 5
            End If
            
            If dblCantItemsNoAtendidos + dblCantItemsAtendidoParcial + dblCantItemsAtendidoTotal = dblCantItems Then
                strSQLOrden = vbNullString
                strSQLOrden = strSQLOrden & "UPDATE "
                    strSQLOrden = strSQLOrden & "IF4ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & " "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
                
                cnn_dbbancos.Execute strSQLOrden
                
                Actualiza_Log strSQLOrden, StrConexDbBancos
            End If
        'intEstado = 5
    'Else
    '    intEstado = 6
    End If
    
    dblCantItems = 0
    dblCantItemsNoAtendidos = 0
    dblCantItemsAtendidoParcial = 0
    dblCantItemsAtendidoTotal = 0
    
    strSQLOrden = vbNullString
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    Set rstOrden = Nothing
    
    atencionOrden = True
    
    Exit Function
errAtencionOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: AtencionOrden"
    
    atencionOrden = False
    
    Err.Clear
End Function

Public Function cerrarOrden(Optional ByVal bolCierrePorItemDeOrden As Boolean) As Boolean
    On Error GoTo errCerrarOrden
    
    Set rstOrden = New ADODB.Recordset
    
    cerrarOrden = False
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO, "
    strSQLOrden = strSQLOrden & "DET.F3PORCDEMASIA, "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA / 100))) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (DET.F3PORCDEMASIA / 100)) AS DEMASIA, "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA / 100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS SALDO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON RESUMEN.F4NUMORD = DET.F4NUMORD AND RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
    
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) <> 0 AND "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) > 0 "
    
        If bolCierrePorItemDeOrden Then
            strSQLOrden = strSQLOrden & "AND "
            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD = '" & strRequerimiento & "' AND "
            strSQLOrden = strSQLOrden & "DET.F3CODPRO = '" & strCodProducto & "' "
        End If
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        rstOrden.MoveFirst
        
        Do While Not rstOrden.EOF
            If Val(rstOrden!SALDO & "") > 0 Then
                strSQLOrden = vbNullString
                strSQLOrden = strSQLOrden & "UPDATE "
                    strSQLOrden = strSQLOrden & "IF3ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F3AJUSTE = 1, "
                    strSQLOrden = strSQLOrden & "F3PORCDEMASIA = 0, "
                    strSQLOrden = strSQLOrden & "F3CANFAL = " & Val(rstOrden!SALDO & "") - Val(rstOrden!DEMASIA & "") & ", "
                    strSQLOrden = strSQLOrden & "F3CANPRO2 = " & Val(rstOrden!F3CANPRO & "") & ", "
                    strSQLOrden = strSQLOrden & "F3CANPRO = " & Val(rstOrden!F3CANPRO & "") - (Val(rstOrden!SALDO & "") - Val(rstOrden!DEMASIA & "")) & " "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
                    strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & Trim(rstOrden!COD_SOLICITUD & "") & "' AND "
                    strSQLOrden = strSQLOrden & "F3CODPRO = '" & Trim(rstOrden!F3CODPRO & "") & "'"
                
                cnn_dbbancos.Execute strSQLOrden
                
                Actualiza_Log strSQLOrden, StrConexDbBancos
            End If
            
            rstOrden.MoveNext
        Loop
            If Not bolCierrePorItemDeOrden Then
                strSQLOrden = vbNullString
                strSQLOrden = strSQLOrden & "UPDATE "
                    strSQLOrden = strSQLOrden & "IF4ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F4ESTADO = 7, "
                    strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "', "
                    strSQLOrden = strSQLOrden & "F4FECMOD = CVDATE('" & strFecMod & "') "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
                
                cnn_dbbancos.Execute strSQLOrden
                
                Actualiza_Log strSQLOrden, StrConexDbBancos
            End If
    End If
    
    strSQLOrden = vbNullString
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    Set rstOrden = Nothing
    
    cerrarOrden = True
    
    Exit Function
errCerrarOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: CerrarOrden"
    
    cerrarOrden = False
    
    Err.Clear
End Function

''-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
''-----------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub listarOrdenParaSeleccion(ByVal comboList As Object, _
                                    ByVal strEstadoOrden As String, _
                                    Optional ByVal bolAdicionarOpcionNuevaOrden As Boolean)
                                    
    On Error GoTo errListarOrdenParaSeleccion
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "F4NUMORD, "
    strSQLOrden = strSQLOrden & "F4NOMPROV "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4CODPRV = '" & strRucProveedor & "' "
        
        If strEstadoOrden <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND F4ESTADO IN (" & strEstadoOrden & ") "
        End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "F4NUMORD"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If bolAdicionarOpcionNuevaOrden Then
        comboList.AddItem "( Nueva Orden )" & Space(200) & "*" & Space(50)
    End If
    
    If Not rstOrden.EOF Then
        
        rstOrden.MoveFirst
        
        Do While Not rstOrden.EOF
            comboList.AddItem Trim(rstOrden!F4NUMORD & "") & " - " & Trim(rstOrden!F4NOMPROV & "") & Space(200) & "*" & Trim(rstOrden!F4NUMORD & "")
            
            rstOrden.MoveNext
        Loop
    End If
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    strSQLOrden = vbNullString
    
    Set rstOrden = Nothing
    
    Exit Sub
errListarOrdenParaSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ListarOrdenParaSeleccion"
    
    Err.Clear
End Sub

'Listar Ordenes en Grilla (QuamtumGrid)
Public Sub listarGrillaOrden(ByVal grilla As dxDBGrid, _
                                ByVal strTipoOrden As String, _
                                ByVal strFiltroSensitivo As String, _
                                ByVal bolVerFechas As Boolean, _
                                ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "CAB.F4LOCAL, "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "CAB.F4OBSERVA, "
    strSQLOrden = strSQLOrden & "CAB.F4TIPMON, "
    strSQLOrden = strSQLOrden & "CAB.F4TIPCAM, "
    strSQLOrden = strSQLOrden & "IIf(CAB.F4TIPMON='S',CAB.F4MONTO, '') AS F4SOLES, "
    strSQLOrden = strSQLOrden & "IIf(CAB.F4TIPMON='D',CAB.F4MONTO,'') AS F4DOLARES, "
    strSQLOrden = strSQLOrden & "CAB.F4COLOCADA, "
    strSQLOrden = strSQLOrden & "CAB.F4COLOCADAUSER, "
    strSQLOrden = strSQLOrden & "CAB.F4COLOCADAFECHA, "
    strSQLOrden = strSQLOrden & "CAB.F4ATENDIDA, "
    strSQLOrden = strSQLOrden & "CAB.F4ATENDIDAUSER, "
    strSQLOrden = strSQLOrden & "CAB.F4ATENDIDAFECHA, "
    strSQLOrden = strSQLOrden & "CAB.F4VB1, "
    strSQLOrden = strSQLOrden & "CAB.F4VBUSER1, "
    strSQLOrden = strSQLOrden & "CAB.F4VBFECHA1, "
    strSQLOrden = strSQLOrden & "CAB.F4VB2, "
    strSQLOrden = strSQLOrden & "CAB.F4VBUSER2, "
    strSQLOrden = strSQLOrden & "CAB.F4VBFECHA2, "
    strSQLOrden = strSQLOrden & "CAB.F4ESTADO "
    
    strSQLOrden = strSQLOrden & "FROM "
    
    strSQLOrden = strSQLOrden & "IF4ORDEN AS CAB "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV "
    
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "TRIM(CAB.F4NUMORD & '') <> '' "
        
        If strTipoOrden <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND F4LOCAL = '" & strTipoOrden & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND ("
            'strSQLOrden = strSQLOrden & "CAB.F4LOCAL LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLOrden = strSQLOrden & "CAB.F4NUMORD  LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLOrden = strSQLOrden & "PROV.F2NOMPROV  LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLOrden = strSQLOrden & "CAB.F4COLOCADAUSER  LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLOrden = strSQLOrden & "CAB.F4ATENDIDAUSER LIKE '%" & strFiltroSensitivo & "%'"
            strSQLOrden = strSQLOrden & ") "
        End If
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD DESC"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Tipo de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4LOCAL"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 30
            End With
            
            'Columna Numero de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Orden"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4NUMORD"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Fecha de Emisión
            Set gColumn = .Columns.Add(gedDateEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fecha Emisión"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4FECEMI"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaEmision"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Fecha de Entrega
            Set gColumn = .Columns.Add(gedDateEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fecha Entrega"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4FECENT"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaEntrega"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Proveedor"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F2NOMPROV"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 180
            End With
            
            'Columna Observaciones
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Observaciones"
                .Color = &H8000000F
                .DisableEditor = True
                .FieldName = "F4OBSERVA"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservaciones"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 130
            End With
            
            'Columna Soles
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Soles"
                .Color = &HC0FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F4SOLES"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSoles"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 65
            End With
            
            'Columna Dolares
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Dólares"
                .Color = &HC0FFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F4DOLARES"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDolares"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 65
            End With
            
            'Columna Enviada Por
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Enviada Por"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4COLOCADAUSER"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColEnviadaPor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Fecha de Envio
            Set gColumn = .Columns.Add(gedDateEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fec. Envió"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4COLOCADAFECHA"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaEnvio"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = bolVerFechas
            End With
            
            'Columna Recibida Por
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Recibida Por"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4ATENDIDAUSER"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColRecibidaPor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Fecha de Recepcion
            Set gColumn = .Columns.Add(gedDateEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fec. Recepción"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4ATENDIDAFECHA"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaRecepcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = bolVerFechas
            End With
            
            'Columna V°B° de Orden
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "V.B."
                .Color = &HFFFFFF
                .DisableEditor = False
                .FieldName = "F4VB1"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColVBOrden"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 35
            End With
            
            'Columna Aprobado Por
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Aprobada Por"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4VBUSER1"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAprobadaPor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = bolVerFechas
            End With
            
            'Columna Fecha de Aprobación
            Set gColumn = .Columns.Add(gedDateEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fec. Aprobación"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "F4VBFECHA1"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaAprobacion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = bolVerFechas
            End With
            
            'Columna Estado de Orden
            Set gColumn = .Columns.Add(gedImageEdit)
            
            With gColumn
                .Alignment = taCenter
                '.BandIndex = 0
                .Caption = "Estado"
                .DisableEditor = True
                .FieldName = "F4ESTADO"
                .Font.Name = "Arial"
                .Font.Size = 8.25
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                
                '.ButtonColumn.ShowButtonStyle = sbNone
                
                With .ImageColumn
                    .Images = listaImagenes.hImageList

                    .ImageIndexes.Add ("0") 'Orden en Edición
                    .Values.Add ("1")
                    .Descriptions.Add ("Orden en Edición")

                    .ImageIndexes.Add ("1") 'Orden Aprobada
                    .Values.Add ("2")
                    .Descriptions.Add ("Orden Aprobada")

                    .ImageIndexes.Add ("2") 'Orden Enviada
                    .Values.Add ("3")
                    .Descriptions.Add ("Orden Enviada")
                    
                    .ImageIndexes.Add ("3") 'Orden Recepcionada
                    .Values.Add ("4")
                    .Descriptions.Add ("Orden Recepcionada")
                    
                    .ImageIndexes.Add ("4") 'Atención Parcial
                    .Values.Add ("5")
                    .Descriptions.Add ("Atención Parcial")
                    
                    .ImageIndexes.Add ("5") 'Atención Total
                    .Values.Add ("6")
                    .Descriptions.Add ("Atención Total")
                    
                    .ImageIndexes.Add ("6") 'Orden Cerrada
                    .Values.Add ("7")
                    .Descriptions.Add ("Orden Cerrada")
                    
                    .ImageIndexes.Add ("7") 'Orden Anulada
                    .Values.Add ("8")
                    .Descriptions.Add ("Orden Anulada")

                    .ShowDescription = False
                End With
                
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexDbBancos
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            '.Dataset.ADODataset.CursorType = ctStatic
            '.Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLOrden
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F4NUMORD"
            
            .Columns.ColumnByFieldName("F4OBSERVA").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            '.Columns.ColumnByFieldName("F4ESTADO").ButtonColumn.ShowButtonStyle = sbNone
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLOrden
    End If
    
    strSQLOrden = vbNullString
    
    Exit Sub
errListarGrillaOrden:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsOrden: ListarGrillaOrden"
    End Select
    
    Err.Clear
End Sub

Public Function generarCadenaSolicitud() As String
    Dim rstSolicitudOC As New ADODB.Recordset
    
    generarCadenaSolicitud = vbNullString
    
    If rstSolicitudOC.State = 1 Then rstSolicitudOC.Close
    
    rstSolicitudOC.Open "SELECT COD_SOLICITUD FROM IF3ORDEN WHERE F4LOCAL = '" & strTipoOrden & "' AND F4NUMORD = '" & strNumeroOrden & "' GROUP BY COD_SOLICITUD", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitudOC.EOF Then
        rstSolicitudOC.MoveFirst
        
        Do While Not rstSolicitudOC.EOF
            If generarCadenaSolicitud = vbNullString Then
                generarCadenaSolicitud = IIf(Trim(rstSolicitudOC!COD_SOLICITUD & "") = vbNullString, "STOCK LIBRE", Trim(rstSolicitudOC!COD_SOLICITUD & "") & " - " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_NOMREF", "TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & strTipoOrden & "'"))
            Else
                generarCadenaSolicitud = generarCadenaSolicitud & "; " & IIf(Trim(rstSolicitudOC!COD_SOLICITUD & "") = vbNullString, "STOCK LIBRE", Trim(rstSolicitudOC!COD_SOLICITUD & "") & "-" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_NOMREF", "TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & strTipoOrden & "'"))
            End If
            
            rstSolicitudOC.MoveNext
        Loop
    End If
    
    Screen.MousePointer = vbDefault
End Function

Public Function obtenerUltimaDescripcionProductoDeProveedor() As String
    On Error GoTo errObtenerUltimaDescripcionProductoDeProveedor
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT TOP 1 "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "TRIM(DET.F5NOMPRO & '') AS ULTIMADESCRIPCIONPROVEEDOR "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F4CODPRV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO = '" & strCodProducto & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "TRIM(DET.F5NOMPRO & '') "
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI DESC"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        obtenerUltimaDescripcionProductoDeProveedor = Trim(rstOrden!ULTIMADESCRIPCIONPROVEEDOR & "")
    Else
        obtenerUltimaDescripcionProductoDeProveedor = vbNullString
    End If
    
    rstOrden.Close
    
    strSQLOrden = vbNullString

    Set rstOrden = Nothing

    Exit Function
errObtenerUltimaDescripcionProductoDeProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerUltimaDescripcionProductoDeProveedor"
    
    obtenerUltimaDescripcionProductoDeProveedor = vbNullString
    
    Err.Clear
End Function

Public Function obtenerUltimoDescuentoProductoDeProveedor() As Double
    On Error GoTo errObtenerUltimoDescuentoProductoDeProveedor
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT TOP 1 "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F3PORDCT AS ULTIMODESCUENTOPROVEEDOR "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F4CODPRV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO = '" & strCodProducto & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "CAB.F4NUMORD, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F3PORDCT "
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI DESC"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        obtenerUltimoDescuentoProductoDeProveedor = Val(rstOrden!ULTIMODESCUENTOPROVEEDOR & "")
    Else
        obtenerUltimoDescuentoProductoDeProveedor = 0
    End If
    
    rstOrden.Close
    
    strSQLOrden = vbNullString

    Set rstOrden = Nothing

    Exit Function
errObtenerUltimoDescuentoProductoDeProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerUltimoDescuentoProductoDeProveedor"
    
    obtenerUltimoDescuentoProductoDeProveedor = 0
    
    Err.Clear
End Function

Public Sub calculosPorItem()
    On Error GoTo errCalculosPorItem
    
    'LIMPIAR VARIABLES
    dblCantidadFinal = 0
    dblPrecioNetoSinImpuesto = 0
    
    dblBasePorItem = 0
    dblExoneradoPorItem = 0
    dblImpuestoPorItem = 0
    dblTotalPorItem = 0
    
    'CALCULOS
    If dblCantidadMaxima > 0 Then
        If dblCantidad > dblCantidadMaxima Then
            dblCantidad = dblCantidadMaxima
        End If
    End If
    
    If dblPrecioSinImpuesto <> 0 Or dblPrecioConImpuesto <> 0 Then
        If dblPrecioSinImpuesto <> 0 And dblPrecioConImpuesto <> 0 Then
            
        ElseIf dblPrecioSinImpuesto <> 0 Then
            dblPrecioConImpuesto = Val(Format((dblPrecioSinImpuesto * (1 + (dblPorcentajeImpuesto * dblSignoImpuesto))) * IIf(bolAfecto, 1, 0), "#0.0000"))
        Else
            dblPrecioSinImpuesto = Val(Format(dblPrecioConImpuesto / (1 + (dblPorcentajeImpuesto * dblSignoImpuesto)), "#.0000"))
            dblPrecioConImpuesto = dblPrecioConImpuesto * IIf(bolAfecto, 1, 0)
        End If
        
        If dblPorcentajeImpuesto = 0 Then dblPrecioConImpuesto = 0
        
        If dblPorcentajeDscto <> 0 Then
            dblTotalDscto = Val(Format(dblPrecioSinImpuesto * dblPorcentajeDscto, "#0.0000"))
        'ElseIf dblPorcentajeDscto = 0 And dblTotalDscto <> 0 Then
        '    dblPorcentajeDscto = Val(Format((dblTotalDscto) / dblPrecioSinImpuesto, "#0.00"))
        Else
            dblPorcentajeDscto = 0
            dblTotalDscto = 0
        End If
    End If
    
    'RESULTADOS
    dblCantidadFinal = dblCantidad * (1 + dblPorcentajeDemasia)
    dblPrecioNetoSinImpuesto = dblPrecioSinImpuesto - dblTotalDscto
    
    dblBasePorItem = Val(Format((dblCantidadFinal * dblPrecioNetoSinImpuesto) * IIf(bolAfecto, 1, 0), "#.0000"))
    dblExoneradoPorItem = Val(Format((dblCantidadFinal * dblPrecioNetoSinImpuesto) * IIf(bolAfecto, 0, 1), "#.0000"))
    dblImpuestoPorItem = Val(Format((dblBasePorItem * dblPorcentajeImpuesto) * IIf(bolAfecto, 1, 0), "#.0000"))
    dblTotalPorItem = Val(Format(dblBasePorItem + dblExoneradoPorItem + (dblImpuestoPorItem * dblSignoImpuesto), "#.00"))
    
    Exit Sub
errCalculosPorItem:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: CalculosPorItem"
    
    Err.Clear
End Sub

Public Function obtenerMovimientoOrdenEnAlmacen() As Boolean
    On Error GoTo errObtenerMovimientoOrdenEnAlmacen
    
    Set rstOrden = New ADODB.Recordset
    
    obtenerMovimientoOrdenEnAlmacen = False
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "'"
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        rstOrden.MoveFirst
        
        Do While Not rstOrden.EOF
            If Trim(rstOrden!Cantidad & "") > 0 Then
                obtenerMovimientoOrdenEnAlmacen = True
                
                Exit Do
            End If
            
            rstOrden.MoveNext
        Loop
    End If
    
    rstOrden.Close
    
    Set rstOrden = Nothing
    
    Exit Function
errObtenerMovimientoOrdenEnAlmacen:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerMovimientoOrdenEnAlmacen"
    
    obtenerMovimientoOrdenEnAlmacen = False
    
    Err.Clear
End Function

Public Function devuelveSQLOrdenMovimiento() As String
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    strSQLOrden = strSQLOrden & "(CAB.F4FECENT - DATE()) AS DIASVENCIMIENTO, "
    strSQLOrden = strSQLOrden & "(DET.F4NUMORD & DET.COD_SOLICITUD & DET.F3CODPRO) AS LLAVE, "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', DET.COD_SOLICITUD, '--') AS COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', REQ.CS_OBSERVACIONES, '< REQUERIDO PARA STOCK LIBRE >') AS CS_OBSERVACIONES, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS F3CANPRO, "
    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#.00')) ) AS SALDO, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F2NOMALM, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4FECVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F1NOMORI, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPRO, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "SUM(VAL(MOVIMIENTO.CANTIDAD & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "((((((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL"
    strSQLOrden = strSQLOrden & ") AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON "
    strSQLOrden = strSQLOrden & "RESUMEN.F4NUMORD = DET.F4NUMORD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "ALM.F2NOMALM, "
    strSQLOrden = strSQLOrden & "DET.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "CAB.F4FECVAL, "
    strSQLOrden = strSQLOrden & "ORI.F1NOMORI, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL AS F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "INS.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO * IIF(TIPO = 'S', -1, 1) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & ") AS MOVIMIENTO "
    strSQLOrden = strSQLOrden & "ON "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMORD = DET.F4NUMORD AND "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.COD_SOLICITUD = DET.COD_SOLICITUD AND "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPROORIGINAL = DET.F3CODPRO "
    
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    strSQLOrden = strSQLOrden & "(CAB.F4FECENT - DATE()), "
    strSQLOrden = strSQLOrden & "(DET.F4NUMORD & DET.COD_SOLICITUD & DET.F3CODPRO), "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', DET.COD_SOLICITUD, '--'), "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', REQ.CS_OBSERVACIONES, '< REQUERIDO PARA STOCK LIBRE >'), "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    'strSQLOrden = strSQLOrden & "VAL(FORMAT( SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#.00')), "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F2NOMALM, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4FECVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F1NOMORI, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPRO, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5NOMPRO "
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', DET.COD_SOLICITUD, '--'), "
    strSQLOrden = strSQLOrden & "IIF(DET.COD_SOLICITUD <> '', REQ.CS_OBSERVACIONES, '< REQUERIDO PARA STOCK LIBRE >'), "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4FECVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMVAL"
    
    devuelveSQLOrdenMovimiento = strSQLOrden
    
    strSQLOrden = vbNullString
End Function

Public Function devuelveSQLOrdenPendienteAtencionPorProveedor(Optional ByVal bolMostrarNroRequerimiento As Boolean, _
                                                                Optional ByVal bolSoloProductoConSaldo As Boolean) As String
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC AS RUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV AS PROVEEDOR, "
    
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING AS DESCRIPCIONINTERNA, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO AS DESCRIPCIONPROV, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED AS UM, "
    
    strSQLOrden = strSQLOrden & "DET.F4NUMORD AS NROOC, "
        
        If bolMostrarNroRequerimiento Then
            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD AS NROPEDIDO, "
        End If
        
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI AS EMISION, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT AS ENTREGA, "
    strSQLOrden = strSQLOrden & "IIF((DATE() - CAB.F4FECENT) = 0, 'PLAZO DE ENTREGA VENCE HOY', IIF((DATE() - CAB.F4FECENT) < 0, 'PLAZO DE ENTREGA A VENCER EN ' & (DATE() - CAB.F4FECENT) * -1 & ' DIA(S).',  'PLAZO DE ENTREGA VENCIDO HACE ' & (DATE() - CAB.F4FECENT) & ' DIA(S).')) AS DIASVENCIMIENTO, "
    
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "SUM(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000'))) AS SALDO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "((((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "CAB.F2CODPROV = '" & strCodProveedor & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON "
    strSQLOrden = strSQLOrden & "RESUMEN.F4NUMORD = DET.F4NUMORD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F4LOCAL = 'OC' AND "
    strSQLOrden = strSQLOrden & "PROV.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "CAB.F4ESTADO > 1 "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    
        If bolMostrarNroRequerimiento Then
            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
        End If
        
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT "
        
        If bolSoloProductoConSaldo Then
            strSQLOrden = strSQLOrden & "HAVING "
            strSQLOrden = strSQLOrden & "SUM((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR))) > 0 "
        End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD"
    
    devuelveSQLOrdenPendienteAtencionPorProveedor = strSQLOrden
    
    strSQLOrden = vbNullString
End Function

Public Function devuelveSQLOrdenPendienteAtencionPorProveedorResumen() As String
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC AS RUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV AS PROVEEDOR, "
    
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING AS DESCRIPCIONINTERNA, "
    'strSQLOrden = strSQLOrden & "DET.F5NOMPRO AS DESCRIPCIONPROV, "
    strSQLOrden = strSQLOrden & "MED.F7NOMMED AS UM, "
    
    'strSQLOrden = strSQLOrden & "DET.F4NUMORD AS NROOC, "
        
        'If bolMostrarNroRequerimiento Then
        '    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD AS NROPEDIDO, "
        'End If
        
    'strSQLOrden = strSQLOrden & "CAB.F4FECEMI AS EMISION, "
    'strSQLOrden = strSQLOrden & "CAB.F4FECENT AS ENTREGA, "
    'strSQLOrden = strSQLOrden & "IIF((DATE() - CAB.F4FECENT) = 0, 'PLAZO DE ENTREGA VENCE HOY', IIF((DATE() - CAB.F4FECENT) < 0, 'PLAZO DE ENTREGA A VENCER EN ' & (DATE() - CAB.F4FECENT) * -1 & ' DIA(S).',  'PLAZO DE ENTREGA VENCIDO HACE ' & (DATE() - CAB.F4FECENT) & ' DIA(S).')) AS DIASVENCIMIENTO, "
    
    'strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "SUM(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000'))) AS SALDO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "((((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "CAB.F2CODPROV = '" & strCodProveedor & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL "
    'strSQLOrden = strSQLOrden & "HAVING "
    'strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) > 0 "
    strSQLOrden = strSQLOrden & ") AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON "
    strSQLOrden = strSQLOrden & "RESUMEN.F4NUMORD = DET.F4NUMORD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND "
    strSQLOrden = strSQLOrden & "RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F4LOCAL = 'OC' AND "
    strSQLOrden = strSQLOrden & "PROV.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "CAB.F4ESTADO > 1 "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    'strSQLOrden = strSQLOrden & "DET.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "MED.F7NOMMED "
    
    'strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    
        'If bolMostrarNroRequerimiento Then
        '    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
        'End If
        
    'strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    'strSQLOrden = strSQLOrden & "CAB.F4FECENT "
        
        'If bolSoloProductoConSaldo Then
            strSQLOrden = strSQLOrden & "HAVING "
            strSQLOrden = strSQLOrden & "SUM((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR))) > 0 "
        'End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    'strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    'strSQLOrden = strSQLOrden & "DET.F4NUMORD"
    
    devuelveSQLOrdenPendienteAtencionPorProveedorResumen = strSQLOrden
    
    strSQLOrden = vbNullString
End Function

Public Function devuelveSQLOrdenPendienteAtencionPorProveedorV2(ByVal bolVerConsolidado As Boolean, _
                                                                Optional ByVal bolSoloProductoConSaldo As Boolean, _
                                                                Optional ByVal bolMostrarNroRequerimiento As Boolean) As String
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC AS RUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV AS PROVEEDOR, "
    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI) AS ANNO, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255) AS DESCRIPCIONINTERNA, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED AS UM, "
    
    If Not bolVerConsolidado Then
        strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO & ''), 255) AS DESCRIPCIONPROV, "
        strSQLOrden = strSQLOrden & "CAB.F4FECEMI AS EMISION, "
        strSQLOrden = strSQLOrden & "CAB.F4FECENT AS ENTREGA, "
        strSQLOrden = strSQLOrden & "DET.F4NUMORD AS NROOC, "
        
        If bolMostrarNroRequerimiento Then
            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
        End If
    End If
    
    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT(DET.F3CANPRO, '#0.0000') & '') * (1 + (DET.F3PORCDEMASIA/100))  ) AS CANTIDAD, "
    
    'strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
    'strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) ) AS SALDO, "
    
    strSQLOrden = strSQLOrden & "SUM("
    strSQLOrden = strSQLOrden & "IIF(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
    strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) < 0, 0, "
    strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
    strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) )"
    strSQLOrden = strSQLOrden & ") AS SALDO "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "((((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON (CAB.F4NUMORD = DET.F4NUMORD) AND (CAB.F4LOCAL = DET.F4LOCAL)) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON (MEDALTER.F7CODMED = DET.UNIDAD) AND (MEDALTER.F5CODPRO = DET.F3CODPRO)) "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT( DET.F3CANPRO, '#0.0000')) * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') "
    'strSQLOrden = strSQLOrden & "CAB.F2CODPROV = '" & strCodProveedor & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON (RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO) AND (RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD) AND (RESUMEN.F4NUMORD = DET.F4NUMORD) "
    
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F4LOCAL = 'OC' AND "
    strSQLOrden = strSQLOrden & "PROV.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "CAB.F4ESTADO NOT IN (1, 6, 7, 8) "
        
        If strCodProducto <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND DET.F3CODPRO = '" & strCodProducto & "' "
        End If
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI), "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255), "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED "
        
    If Not bolVerConsolidado Then
        strSQLOrden = strSQLOrden & ", LEFT(TRIM(DET.F5NOMPRO & ''), 255), "
        strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
        strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
        strSQLOrden = strSQLOrden & "DET.F4NUMORD "
        
        If bolMostrarNroRequerimiento Then
            strSQLOrden = strSQLOrden & ", DET.COD_SOLICITUD "
        End If
    End If
    
    strSQLOrden = strSQLOrden & "HAVING "
    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT(DET.F3CANPRO, '#0.0000') & '') * (1 + (DET.F3PORCDEMASIA/100))  ) > 0 "
    
    If bolSoloProductoConSaldo Then
        strSQLOrden = strSQLOrden & "AND "
        strSQLOrden = strSQLOrden & "SUM("
        strSQLOrden = strSQLOrden & "IIF(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
        strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) < 0, 0, "
        strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
        strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) )"
        strSQLOrden = strSQLOrden & ") > 0 "
    End If
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI), "
    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255), "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED"
    
    If Not bolVerConsolidado Then
        strSQLOrden = strSQLOrden & ", DET.F4NUMORD, "
        strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
        strSQLOrden = strSQLOrden & "CAB.F4FECENT"
        
        If bolMostrarNroRequerimiento Then
            strSQLOrden = strSQLOrden & ", DET.COD_SOLICITUD"
        End If
    End If
    
    devuelveSQLOrdenPendienteAtencionPorProveedorV2 = strSQLOrden
    
    strSQLOrden = vbNullString
End Function

Public Function devuelveSQLAtencionOrden(Optional ByVal bolSoloProductosConSaldoPorEntregar As Boolean)
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS SALDO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(((IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLOrden = strSQLOrden & "LEFT JOIN SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
    strSQLOrden = strSQLOrden & "LEFT JOIN IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
    strSQLOrden = strSQLOrden & "ON RESUMEN.F4NUMORD = DET.F4NUMORD AND RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
    
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) <> 0 "
        
        If bolSoloProductosConSaldoPorEntregar Then
            strSQLOrden = strSQLOrden & "AND (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) > 0 "
        End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    
    devuelveSQLAtencionOrden = strSQLOrden
    
    strSQLOrden = vbNullString
End Function

