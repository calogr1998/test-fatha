VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsOrden"
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

Private strSQLSelectAlter               As String

Private rstOrden             As ADODB.Recordset
Private strSQLOrden          As String



Public Property Let TipoOrden(ByVal value As String)
    strTipoOrden = value
End Property

Public Property Get TipoOrden() As String
    TipoOrden = strTipoOrden
End Property

Public Property Let NumeroOrden(ByVal value As String)
    strNumeroOrden = value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property




Public Property Let FechaEmision(ByVal value As String)
    strFechaEmision = value
End Property

Public Property Get FechaEmision() As String
    FechaEmision = strFechaEmision
End Property

Public Property Let CodProveedor(ByVal value As String)
    strCodProveedor = value
End Property

Public Property Get CodProveedor() As String
    CodProveedor = strCodProveedor
End Property

Public Property Let NomProveedor(ByVal value As String)
    strNomProveedor = value
End Property

Public Property Get NomProveedor() As String
    NomProveedor = strNomProveedor
End Property

Public Property Let RucProveedor(ByVal value As String)
    strRucProveedor = value
End Property

Public Property Get RucProveedor() As String
    RucProveedor = strRucProveedor
End Property

Public Property Let ContactoProveedor(ByVal value As String)
    strContactoProveedor = value
End Property

Public Property Get ContactoProveedor() As String
    ContactoProveedor = strContactoProveedor
End Property

Public Property Let CodTipoComprobante(ByVal value As String)
    strCodTipoComprobante = value
End Property

Public Property Get CodTipoComprobante() As String
    CodTipoComprobante = strCodTipoComprobante
End Property

Public Property Let OrdenRegularizada(ByVal value As Boolean)
    bolOrdenRegularizada = value
End Property

Public Property Get OrdenRegularizada() As Boolean
    OrdenRegularizada = bolOrdenRegularizada
End Property




Public Property Let FechaEntrega(ByVal value As String)
    strFechaEntrega = value
End Property

Public Property Get FechaEntrega() As String
    FechaEntrega = strFechaEntrega
End Property

Public Property Let CodigoSolicitante(ByVal value As String)
    strCodigoSolicitante = value
End Property

Public Property Get CodigoSolicitante() As String
    CodigoSolicitante = strCodigoSolicitante
End Property

Public Property Let CodFormaPago(ByVal value As String)
    strCodFormaPago = value
End Property

Public Property Get CodFormaPago() As String
    CodFormaPago = strCodFormaPago
End Property

Public Property Let CentroCosto(ByVal value As String)
    strCentroCosto = value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property

Public Property Let LugarEntrega(ByVal value As String)
    strLugarEntrega = value
End Property

Public Property Get LugarEntrega() As String
    LugarEntrega = strLugarEntrega
End Property

Public Property Let PagoParcial(ByVal value As Boolean)
    bolPagoParcial = value
End Property

Public Property Get PagoParcial() As Boolean
    PagoParcial = bolPagoParcial
End Property

Public Property Let CodMoneda(ByVal value As String)
    strCodMoneda = value
End Property

Public Property Get CodMoneda() As String
    CodMoneda = strCodMoneda
End Property

Public Property Let TipoCambio(ByVal value As Double)
    dblTipoCambio = value
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dblTipoCambio
End Property

Public Property Let FechaPago(ByVal value As String)
    strFechaPago = value
End Property

Public Property Get FechaPago() As String
    FechaPago = strFechaPago
End Property

Public Property Let NumeroCotizacion(ByVal value As String)
    strNumeroCotizacion = value
End Property

Public Property Get NumeroCotizacion() As String
    NumeroCotizacion = strNumeroCotizacion
End Property




Public Property Let VB1(ByVal value As Boolean)
    bolVB1 = value
End Property

Public Property Get VB1() As Boolean
    VB1 = bolVB1
End Property

Public Property Let VB1Usuario(ByVal value As String)
    strVB1Usuario = value
End Property

Public Property Get VB1Usuario() As String
    VB1Usuario = strVB1Usuario
End Property

Public Property Let VB1Fecha(ByVal value As String)
    strVB1Fecha = value
End Property

Public Property Get VB1Fecha() As String
    VB1Fecha = strVB1Fecha
End Property

Public Property Let VB2(ByVal value As Boolean)
    bolVB2 = value
End Property

Public Property Get VB2() As Boolean
    VB2 = bolVB2
End Property

Public Property Let VB2Usuario(ByVal value As String)
    strVB2Usuario = value
End Property

Public Property Get VB2Usuario() As String
    VB2Usuario = strVB2Usuario
End Property

Public Property Let VB2Fecha(ByVal value As String)
    strVB2Fecha = value
End Property

Public Property Get VB2Fecha() As String
    VB2Fecha = strVB2Fecha
End Property

Public Property Let Colocada(ByVal value As Boolean)
    bolColocada = value
End Property

Public Property Get Colocada() As Boolean
    Colocada = bolColocada
End Property

Public Property Let ColocadaUsuario(ByVal value As String)
    strColocadaUsuario = value
End Property

Public Property Get ColocadaUsuario() As String
    ColocadaUsuario = strColocadaUsuario
End Property

Public Property Let ColocadaFecha(ByVal value As String)
    strColocadaFecha = value
End Property

Public Property Get ColocadaFecha() As String
    ColocadaFecha = strColocadaFecha
End Property

Public Property Let Atendida(ByVal value As Boolean)
    bolAtendida = value
End Property

Public Property Get Atendida() As Boolean
    Atendida = bolAtendida
End Property

Public Property Let AtendidaUsuario(ByVal value As String)
    strAtendidaUsuario = value
End Property

Public Property Get AtendidaUsuario() As String
    AtendidaUsuario = strAtendidaUsuario
End Property

Public Property Let AtendidaFecha(ByVal value As String)
    strAtendidaFecha = value
End Property

Public Property Get AtendidaFecha() As String
    AtendidaFecha = strAtendidaFecha
End Property




'Empresa
Public Property Let Empresa(ByVal value As String)
    strEmpresa = value
End Property

Public Property Get Empresa() As String
    Empresa = strEmpresa
End Property
'Observacion
Public Property Let Observacion(ByVal value As String)
    strObservacion = value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property




Public Property Let SUBTOTAL(ByVal value As Double)
    dblSubTotal = value
End Property

Public Property Get SUBTOTAL() As Double
    SUBTOTAL = dblSubTotal
End Property

Public Property Let TotalImpuesto(ByVal value As Double)
    dblTotalImpuesto = value
End Property

Public Property Get TotalImpuesto() As Double
    TotalImpuesto = dblTotalImpuesto
End Property

Public Property Let TotalFacturado(ByVal value As Double)
    dblTotalFacturado = value
End Property

Public Property Get TotalFacturado() As Double
    TotalFacturado = dblTotalFacturado
End Property

Public Property Let TotalInafecto(ByVal value As Double)
    dblTotalInafecto = value
End Property

Public Property Get TotalInafecto() As Double
    TotalInafecto = dblTotalInafecto
End Property

Public Property Let TotalRedondeo(ByVal value As Double)
    dblTotalRedondeo = value
End Property

Public Property Get TotalRedondeo() As Double
    TotalRedondeo = dblTotalRedondeo
End Property




Public Property Let Estado(ByVal value As Integer)
    intEstado = value
End Property

Public Property Get Estado() As Integer
    Estado = intEstado
End Property

Public Property Let CodigoRequerimiento(ByVal value As String)
    strCodigoRequerimiento = value
End Property

Public Property Get CodigoRequerimiento() As String
    CodigoRequerimiento = strCodigoRequerimiento
End Property
'DocAnulado
Public Property Let DocAnulado(ByVal value As Boolean)
    bolDocAnulado = value
End Property

Public Property Get DocAnulado() As Boolean
    DocAnulado = bolDocAnulado
End Property

Public Property Let DocFalta(ByVal value As Boolean)
    bolDocFalta = value
End Property

Public Property Get DocFalta() As Boolean
    DocFalta = bolDocFalta
End Property

Public Property Let OrdenConVale(ByVal value As Boolean)
    bolOrdenConVale = value
End Property

Public Property Get OrdenConVale() As Boolean
    OrdenConVale = bolOrdenConVale
End Property

'Referencia
Public Property Let referencia(ByVal value As String)
    strReferencia = value
End Property

Public Property Get referencia() As String
    referencia = strReferencia
End Property

Public Property Let Correlativo(ByVal value As Double)
    dblCorrelativo = value
End Property

Public Property Get Correlativo() As Double
    Correlativo = dblCorrelativo
End Property

Public Property Let AplicaDetraccion(ByVal value As Boolean)
    bolAplicaDetraccion = value
End Property

Public Property Get AplicaDetraccion() As Boolean
    AplicaDetraccion = bolAplicaDetraccion
End Property

Public Property Let PorcentajeDetraccion(ByVal value As Double)
    dblPorcentajeDetraccion = value
End Property

Public Property Get PorcentajeDetraccion() As Double
    PorcentajeDetraccion = dblPorcentajeDetraccion
End Property

Public Property Let SinProveedorEspecifico(ByVal value As Boolean)
    bolSinProveedorEspecifico = value
End Property

Public Property Get SinProveedorEspecifico() As Boolean
    SinProveedorEspecifico = bolSinProveedorEspecifico
End Property






'ITEM
Public Property Let ITEM(ByVal value As Double)
    dblItem = value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property

'COD_SOLICITUD
Public Property Let Requerimiento(ByVal value As String)
    strRequerimiento = value
End Property

Public Property Get Requerimiento() As String
    Requerimiento = strRequerimiento
End Property

'F3CODPRO
Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

'F3CODFAB
Public Property Let CodigoFabricante(ByVal value As String)
    strCodFabricante = value
End Property

Public Property Get CodigoFabricante() As String
    CodigoFabricante = strCodFabricante
End Property

'F5NOMPRO
Public Property Let NombreProducto(ByVal value As String)
    strNomProducto = value
End Property

Public Property Get NombreProducto() As String
    NombreProducto = strNomProducto
End Property

'F5NOMPRO_ING
Public Property Let NombreProductoInterno(ByVal value As String)
    strNomProductoInterno = value
End Property

Public Property Get NombreProductoInterno() As String
    NombreProductoInterno = strNomProductoInterno
End Property

'UNIDAD
Public Property Let CodigoUM(ByVal value As String)
    strCodUM = value
End Property

Public Property Get CodigoUM() As String
    CodigoUM = strCodUM
End Property

'F3CANPRO
Public Property Let Cantidad(ByVal value As Double)
    dblCantidad = value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

'F3CANPRO2
Public Property Let CantidadMaxima(ByVal value As Double)
    dblCantidadMaxima = value
End Property

Public Property Get CantidadMaxima() As Double
    CantidadMaxima = dblCantidadMaxima
End Property

'F3PORCDEMASIA
Public Property Let PorcentajeDemasia(ByVal value As Double)
    dblPorcentajeDemasia = value
End Property

Public Property Get PorcentajeDemasia() As Double
    PorcentajeDemasia = dblPorcentajeDemasia
End Property

'F3PRECOS
Public Property Let PrecioSinImpuesto(ByVal value As Double)
    dblPrecioSinImpuesto = value
End Property

Public Property Get PrecioSinImpuesto() As Double
    PrecioSinImpuesto = dblPrecioSinImpuesto
End Property

'F3PREUNI
Public Property Let PrecioConImpuesto(ByVal value As Double)
    dblPrecioConImpuesto = value
End Property

Public Property Get PrecioConImpuesto() As Double
    PrecioConImpuesto = dblPrecioConImpuesto
End Property

'F3PRENETO
Public Property Let PrecioNetoSinImpuesto(ByVal value As Double)
    dblPrecioNetoSinImpuesto = value
End Property

Public Property Get PrecioNetoSinImpuesto() As Double
    PrecioNetoSinImpuesto = dblPrecioNetoSinImpuesto
End Property

'F3PORDCT
Public Property Let PorcentajeDscto(ByVal value As Double)
    dblPorcentajeDscto = value
End Property

Public Property Get PorcentajeDscto() As Double
    PorcentajeDscto = dblPorcentajeDscto
End Property

'F3TOTDCT
Public Property Let TotalDscto(ByVal value As Double)
    dblTotalDscto = value
End Property

Public Property Get TotalDscto() As Double
    TotalDscto = dblTotalDscto
End Property

'F5AFECTO
Public Property Let Afecto(ByVal value As Boolean)
    bolAfecto = value
End Property

Public Property Get Afecto() As Boolean
    Afecto = bolAfecto
End Property

'F3IGV
Public Property Let ImpuestoPorItem(ByVal value As Double)
    dblImpuestoPorItem = value
End Property

Public Property Get ImpuestoPorItem() As Double
    ImpuestoPorItem = dblImpuestoPorItem
End Property

'F5VALVTA
Public Property Let BasePorItem(ByVal value As Double)
    dblBasePorItem = value
End Property

Public Property Get BasePorItem() As Double
    BasePorItem = dblBasePorItem
End Property

'F3TOTAL
Public Property Let TotalPorItem(ByVal value As Double)
    dblTotalPorItem = value
End Property

Public Property Get TotalPorItem() As Double
    TotalPorItem = dblTotalPorItem
End Property

'CODCOLOR
Public Property Let CodigoColor(ByVal value As String)
    strCodColor = value
End Property

Public Property Get CodigoColor() As String
    CodigoColor = strCodColor
End Property

'F3OBSERVA
Public Property Let ObservacionPorItem(ByVal value As String)
    strObservacionPorItem = value
End Property

Public Property Get ObservacionPorItem() As String
    ObservacionPorItem = strObservacionPorItem
End Property

'F3GASTO
Public Property Let CodigoGasto(ByVal value As String)
    strCodigoGasto = value
End Property

Public Property Get CodigoGasto() As String
    CodigoGasto = strCodigoGasto
End Property

'F3CUENTA
Public Property Let CuentaContable(ByVal value As String)
    strCuentaContable = value
End Property

Public Property Get CuentaContable() As String
    CuentaContable = strCuentaContable
End Property





'Atributos Auxiliares
Public Property Let PorcentajeImpuesto(ByVal value As Double)
    dblPorcentajeImpuesto = value
End Property

Public Property Get PorcentajeImpuesto() As Double
    PorcentajeImpuesto = dblPorcentajeImpuesto
End Property

Public Property Let SignoImpuesto(ByVal value As Double)
    dblSignoImpuesto = value
End Property

Public Property Get SignoImpuesto() As Double
    SignoImpuesto = dblSignoImpuesto
End Property

Public Property Let CantidadFinal(ByVal value As Double)
    dblCantidadFinal = value
End Property

Public Property Get CantidadFinal() As Double
    CantidadFinal = dblCantidadFinal
End Property

Public Property Let ExoneradoPorItem(ByVal value As Double)
    dblExoneradoPorItem = value
End Property

Public Property Get ExoneradoPorItem() As Double
    ExoneradoPorItem = dblExoneradoPorItem
End Property




'Atributos adicionales
'F3CANFAL
Public Property Let CantidadFaltante(ByVal value As Double)
    dblCantidadFaltante = value
End Property

Public Property Get CantidadFaltante() As Double
    CantidadFaltante = dblCantidadFaltante
End Property
'F3AJUSTE
Public Property Let ItemAjustado(ByVal value As Boolean)
    bolItemAjustado = value
End Property

Public Property Get ItemAjustado() As Boolean
    ItemAjustado = bolItemAjustado
End Property





'Propiedad Fecha de Registro de Orden
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Orden
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Orden
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Orden
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
    
    dblCantidadFaltante = 0 'F3CANFAL
    bolItemAjustado = False 'F3AJUSTE
End Sub

Public Function obtenerOrden() As Boolean
    On Error GoTo errObtenerOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
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
        
        'obtenerOrdenDetalle
        
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
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerOrden"
    'Resume
    Err.Clear
End Function

Public Sub obtenerOrdenDetalle()
    Dim rstOrdenDetalle As New ADODB.Recordset
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA"
    
    abrirCnTemporal
    
'    strSQLOrden = vbNullString
'    strSQLOrden = strSQLOrden & "INSERT INTO TMPORDENDECOMPRA("
'        strSQLOrden = strSQLOrden & "ITEM, F3CODPRO, F3CODFAB, F5NOMPRO, F5NOMPRO_ING, "
'        strSQLOrden = strSQLOrden & "F3CODMEDIDA, F3DESMEDIDA, F3CANPRO, F3CANPROMAX, CANT_ANT, "
'        strSQLOrden = strSQLOrden & "F3PORCDEMASIA, F3CANPROFINAL, F3SINIGV, F3CONIGV, "
'        strSQLOrden = strSQLOrden & "F3NETO, F5AFECTO, F3BASEIMP, F3MONINA, F3IGV, "
'        strSQLOrden = strSQLOrden & "F3REDONDEO, F3TOTAL, F5CODCOSTO, F5DESCOSTO, "
'        strSQLOrden = strSQLOrden & "F3VALDESC, F3PORDESC, F3OBSERVA, "
'        strSQLOrden = strSQLOrden & "COD_SOLICITUD, CLIENTE, CODCOLOR, DESCOLOR, "
'        strSQLOrden = strSQLOrden & "CERRAR, F3GASTO, F3CUENTA"
'        strSQLOrden = strSQLOrden & ") "
'
'    strSQLOrden = strSQLOrden & "IN '" & wrutatemp & "Templus.mdb' "
    
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
        strSQLOrden = strSQLOrden & "(CASE WHEN DET.F5AFECTO = '*' THEN 1 ELSE 0 END) AS AFECTO, "
        strSQLOrden = strSQLOrden & "(CASE WHEN DET.F5AFECTO = '*' THEN DET.F5VALVTA ELSE 0 END) AS BASE,"
        strSQLOrden = strSQLOrden & "(CASE WHEN DET.F5AFECTO = '*' THEN 0 ELSE DET.F5VALVTA END) AS INAFECTO,"
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
        strSQLOrden = strSQLOrden & "(((PROCESOS.IF3ORDEN AS DET "
        strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
        strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2BIENCOLOR AS COL ON COL.CODIGO = DET.CODCOLOR) "
        strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD "
    strSQLOrden = strSQLOrden & "WHERE "
            strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
            strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "'"
    
    'cnBdCPlus.Execute strSQLOrden
    
    If rstOrdenDetalle.State = 1 Then rstOrdenDetalle.Close
    
    rstOrdenDetalle.Open strSQLOrden, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrdenDetalle.EOF Then
        Do While Not rstOrdenDetalle.EOF
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
            
            strSQLOrden = strSQLOrden & "VALUES("
                strSQLOrden = strSQLOrden & "'" & Trim(rstOrdenDetalle!ITEM & "") & "', '" & Trim(rstOrdenDetalle!F3CODPRO & "") & "' , '" & Trim(rstOrdenDetalle!F3CODFAB & "") & "', "
                strSQLOrden = strSQLOrden & "'" & Trim(rstOrdenDetalle!F5NOMPRO & "") & "', '" & Trim(rstOrdenDetalle!F5NOMPRO_ING & "") & "', '" & Trim(rstOrdenDetalle!UNIDAD & "") & "', "
                strSQLOrden = strSQLOrden & "'" & Trim(rstOrdenDetalle!F7SIGMED & "") & "', " & Val(rstOrdenDetalle!F3CANPRO & "") & ", " & Val(rstOrdenDetalle!F3CANPRO2 & "") & ", "
                strSQLOrden = strSQLOrden & Val(rstOrdenDetalle!F3CANFAL & "") & ", " & Val(rstOrdenDetalle!F3PORCDEMASIA & "") & ", " & Val(rstOrdenDetalle!F3CANPROFINAL & "") & ", "
                strSQLOrden = strSQLOrden & Val(rstOrdenDetalle!F3PRECOS & "") & ", " & Val(rstOrdenDetalle!F3PREUNI & "") & ", " & Val(rstOrdenDetalle!F3PRENETO & "") & ", "
                strSQLOrden = strSQLOrden & Val(rstOrdenDetalle!Afecto & "") & ", " & Val(rstOrdenDetalle!Base & "") & ", " & Val(rstOrdenDetalle!INAFECTO & "") & ", "
                strSQLOrden = strSQLOrden & Val(rstOrdenDetalle!F3IGV & "") & ", " & Val(rstOrdenDetalle!REDONDEO & "") & ", " & Val(rstOrdenDetalle!F3TOTAL & "") & ", "
                strSQLOrden = strSQLOrden & "'" & Trim(rstOrdenDetalle!F5CODCOSTO & "") & "', '" & Trim(rstOrdenDetalle!F5DESCOSTO & "") & "', " & Val(rstOrdenDetalle!F3TOTDCT & "") & ", "
                strSQLOrden = strSQLOrden & Val(rstOrdenDetalle!F3PORDCT & "") & ", '" & Trim(rstOrdenDetalle!F3OBSERVA & "") & "', '" & Trim(rstOrdenDetalle!COD_SOLICITUD & "") & "', "
                strSQLOrden = strSQLOrden & "'" & Trim(rstOrdenDetalle!CS_NOMREF & "") & "', '" & Trim(rstOrdenDetalle!CODCOLOR & "") & "', " & Val(rstOrdenDetalle!Descripcion & "") & ", "
                strSQLOrden = strSQLOrden & IIf(CBool(rstOrdenDetalle!F3AJUSTE), "TRUE", "FALSE") & ", '" & Trim(rstOrdenDetalle!F3GASTO & "") & "', '" & Trim(rstOrdenDetalle!F3CUENTA & "") & "'"
            strSQLOrden = strSQLOrden & ")"
            
            cnDBTemp.Execute strSQLOrden
            
            rstOrdenDetalle.MoveNext
        Loop
    End If
    
    If rstOrdenDetalle.State = 1 Then rstOrdenDetalle.Close
    
    Set rstOrdenDetalle = Nothing
    
    strSQLOrden = vbNullString
End Sub


Public Sub obtenerConfigOrden()
    On Error GoTo errObtenerConfigOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerConfigOrden"
    
    Err.Clear
End Sub

Public Sub obtenerConfigOrdenDetalleOnebyOne()
    On Error GoTo errObtenerConfigOrdenDetalleOnebyOne
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "F3CODPRO = '" & strCodProducto & "' AND "
    strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & strRequerimiento & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerConfigOrdenDetalleOnebyOne"
    
    Err.Clear
End Sub

Public Function generarNumeroOrden() As String
    On Error GoTo errGenerarNumeroOrden
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT TOP 1 "
    strSQLOrden = strSQLOrden & "RIGHT(F4NUMORD,10) AS CODIGO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' "
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "F4NUMORD DESC"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        generarNumeroOrden = strTipoOrden & Format(Val(Trim(rstOrden!Codigo & "")) + 1, "0000000000")
    Else
        generarNumeroOrden = strTipoOrden & "0000000001"
    End If

    rstOrden.Close

    Set rstOrden = Nothing
    strSQLOrden = vbNullString

    Exit Function
errGenerarNumeroOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: GenerarNumeroOrden"
    
    generarNumeroOrden = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstOrden = New ADODB.Recordset
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN.* "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
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
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarOrden() As Boolean
    On Error GoTo errGuardarOrden
    
    strSQLOrden = vbNullString
    
    If Not verificarExistencia Then
        If strNumeroOrden = vbNullString Then
            strNumeroOrden = generarNumeroOrden
        End If
        
        strSQLOrden = strSQLOrden & "INSERT INTO PROCESOS.IF4ORDEN("
        strSQLOrden = strSQLOrden & "F4LOCAL, F4NUMORD, F4FECEMI, F4CODCLI, F4NOMPROV, F4CODPRV, "
        strSQLOrden = strSQLOrden & "F4CONTACTO, F4TIPDOC, F4REGULARIZA, F4FECENT, F4CODSOL, "
        strSQLOrden = strSQLOrden & "F4FORPAG, F4CENTRO, F4LUGAR_ENTREGA, F4PAGOPARCIAL, F4TIPMON, "
        strSQLOrden = strSQLOrden & "F4TIPCAM, F4DIAPAGO, F4NUMCOTIZA, F4VB1, F4VBUSER1, F4VBFECHA1, "
        strSQLOrden = strSQLOrden & "F4VB2, F4VBUSER2, F4VBFECHA2, F4COLOCADA, F4COLOCADAUSER, "
        strSQLOrden = strSQLOrden & "F4COLOCADAFECHA, F4ATENDIDA, F4ATENDIDAUSER, F4ATENDIDAFECHA, "
        strSQLOrden = strSQLOrden & "F4EMPRESA, F4OBSERVA, F4BASIMP, F4IGV, F4MONTO, F4MONINA, F4RND, "
        strSQLOrden = strSQLOrden & "F4ESTADO, F4CODSOLICITUD, F4ESTNUL, F4FALTA, F4ESTVAL, F4REFERE, "
        strSQLOrden = strSQLOrden & "F4CORRELA, F4DETRACCIONAPLICA, F4DETRACCIONPORC, F4SINPROVEEDORESP, "
        strSQLOrden = strSQLOrden & "F4FECGRA, F4USEGRA) "
        
        strSQLOrden = strSQLOrden & "VALUES("
        strSQLOrden = strSQLOrden & "'" & strTipoOrden & "', '" & strNumeroOrden & "', '" & strFechaEmision & "', "
        strSQLOrden = strSQLOrden & "'" & strCodProveedor & "', '" & strNomProveedor & "', '" & strRucProveedor & "', '" & strContactoProveedor & "', "
        strSQLOrden = strSQLOrden & "'" & strCodTipoComprobante & "', '" & IIf(bolOrdenRegularizada, "1", "0") & "', '" & strFechaEntrega & "', "
        strSQLOrden = strSQLOrden & "'" & strCodigoSolicitante & "', '" & strCodFormaPago & "', '" & strCentroCosto & "', '" & strLugarEntrega & "', "
        strSQLOrden = strSQLOrden & IIf(bolPagoParcial, "1", "0") & ", '" & strCodMoneda & "', " & dblTipoCambio & ", '" & strFechaPago & "', "
        strSQLOrden = strSQLOrden & "'" & strNumeroCotizacion & "', " & IIf(bolVB1, "1", "0") & ", "
        strSQLOrden = strSQLOrden & IIf(strVB1Usuario <> vbNullString, "'" & strVB1Usuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & IIf(strVB1Fecha <> vbNullString, "CONVERT(DATETIME, '" & Format(strVB1Fecha, "'dd-mm-yy hh:mm:ss") & "', 5)", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolVB2, "1", "0") & ", " & IIf(strVB2Usuario <> vbNullString, "'" & strVB2Usuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & IIf(strVB2Fecha <> vbNullString, "CONVERT(DATETIME, '" & Format(strVB2Fecha, "'dd-mm-yy hh:mm:ss") & "', 5)", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolColocada, "1", "0") & ", "
        strSQLOrden = strSQLOrden & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & IIf(strColocadaFecha <> vbNullString, "'" & strColocadaFecha & "'", "NULL") & ", "
        strSQLOrden = strSQLOrden & IIf(bolAtendida, "1", "0") & ", "
        strSQLOrden = strSQLOrden & IIf(strAtendidaUsuario <> vbNullString, "'" & strAtendidaUsuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & IIf(strAtendidaFecha <> vbNullString, "'" & strAtendidaFecha & "'", "NULL") & ", "
        strSQLOrden = strSQLOrden & "'" & strEmpresa & "', '" & strObservacion & "', " & dblSubTotal & ", " & dblTotalImpuesto & ", "
        strSQLOrden = strSQLOrden & dblTotalFacturado & ", " & dblTotalInafecto & ", " & dblTotalRedondeo & ", "
        strSQLOrden = strSQLOrden & intEstado & ", '" & strCodigoRequerimiento & "', '" & IIf(bolDocAnulado, "S", "N") & "', "
        strSQLOrden = strSQLOrden & "'" & IIf(bolDocFalta, "1", "0") & "', '" & IIf(bolOrdenConVale, "1", "0") & "', '" & strReferencia & "', "
        strSQLOrden = strSQLOrden & dblCorrelativo & ", " & IIf(bolAplicaDetraccion, "1", "0") & ", " & dblPorcentajeDetraccion & ", "
        strSQLOrden = strSQLOrden & IIf(bolSinProveedorEspecifico, "1", "0") & ", "
        strSQLOrden = strSQLOrden & "CONVERT(DATETIME, '" & Format(strFecReg, "dd-mm-yy hh:mm:ss") & "', 5), '" & strUsuReg & "')"
    Else
        strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
        strSQLOrden = strSQLOrden & "SET "
            
            strSQLOrden = strSQLOrden & "F4FECEMI = '" & strFechaEmision & "', "
            strSQLOrden = strSQLOrden & "F4CODCLI = '" & strCodProveedor & "', "
            strSQLOrden = strSQLOrden & "F4NOMPROV = '" & strNomProveedor & "', "
            strSQLOrden = strSQLOrden & "F4CODPRV = '" & strRucProveedor & "', "
            strSQLOrden = strSQLOrden & "F4CONTACTO = '" & strContactoProveedor & "', "
            strSQLOrden = strSQLOrden & "F4TIPDOC = '" & strCodTipoComprobante & "', "
            strSQLOrden = strSQLOrden & "F4REGULARIZA = '" & IIf(bolOrdenRegularizada, "1", "0") & "', "
            strSQLOrden = strSQLOrden & "F4FECENT = '" & strFechaEntrega & "', "
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
            strSQLOrden = strSQLOrden & "F4COLOCADA = " & IIf(bolColocada, "1", "0") & ", "
            strSQLOrden = strSQLOrden & "F4COLOCADAUSER = " & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "''") & ", "
            strSQLOrden = strSQLOrden & "F4COLOCADAFECHA = " & IIf(strColocadaFecha <> vbNullString, "'" & strColocadaFecha & "'", "NULL") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDA = " & IIf(bolAtendida, "1", "0") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDAUSER = " & IIf(strAtendidaUsuario <> vbNullString, "'" & strAtendidaUsuario & "'", "''") & ", "
            strSQLOrden = strSQLOrden & "F4ATENDIDAFECHA = " & IIf(strAtendidaFecha <> vbNullString, "'" & strAtendidaFecha & "'", "NULL") & ", "
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
            strSQLOrden = strSQLOrden & "F4FECMOD = CONVERT(DATETIME, '" & Format(strFecMod, "dd-mm-yy hh:mm:ss") & "', 5), "
            strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "' "
            
        strSQLOrden = strSQLOrden & "WHERE "
            strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
            strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    End If
    
    cnBdCPlus.Execute strSQLOrden
    
    guardarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errGuardarOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vuelva a intentarlo nuevamente.", _
            vbInformation + vbOKOnly, App.ProductName & " - SqlClsOrden: GuardarOrden"
    'Resume
    guardarOrden = False
    
    Err.Clear
End Function

Public Sub guardarOrdenDetalleOneByOne()
    On Error GoTo errGuardarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "INSERT INTO PROCESOS.IF3ORDEN("
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
    strSQLOrden = strSQLOrden & "'" & strTipoOrden & "', '" & strNumeroOrden & "', " & dblItem & ", "
    strSQLOrden = strSQLOrden & "'" & strRequerimiento & "', '" & strCodProducto & "', '" & strCodFabricante & "', "
    strSQLOrden = strSQLOrden & "'" & strNomProducto & "', '" & strNomProductoInterno & "', '" & strCodUM & "', "
    strSQLOrden = strSQLOrden & dblCantidad & ", " & dblCantidadMaxima & ", " & dblCantidadFaltante & ", " & dblPorcentajeDemasia & ", "
    strSQLOrden = strSQLOrden & dblPrecioSinImpuesto & ", " & dblPrecioConImpuesto & ", " & dblPrecioNetoSinImpuesto & ", "
    strSQLOrden = strSQLOrden & dblPorcentajeDscto & ", " & dblTotalDscto & ", " & IIf(bolAfecto, "'*'", "''") & ", "
    strSQLOrden = strSQLOrden & dblBasePorItem & ", " & dblImpuestoPorItem & ", " & dblTotalPorItem & ", "
    strSQLOrden = strSQLOrden & "'" & strCodColor & "', '" & strObservacionPorItem & "', "
    strSQLOrden = strSQLOrden & IIf(bolItemAjustado, "1", "0") & ", "
    strSQLOrden = strSQLOrden & IIf(strCodigoGasto <> vbNullString, "'" & strCodigoGasto & "'", "''") & ", "
    strSQLOrden = strSQLOrden & IIf(strCuentaContable <> vbNullString, "'" & strCuentaContable & "'", "''") & ")"
    
    cnBdCPlus.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errGuardarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: GuardarOrdenDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    strSQLOrden = vbNullString
    
    Err.Clear
End Sub

Public Sub actualizarOrdenDetalleOneByOne()
    On Error GoTo errActualizarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "UPDATE "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
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
    strSQLOrden = strSQLOrden & "F5AFECTO = " & IIf(bolAfecto, "'*'", "''") & ", "
    strSQLOrden = strSQLOrden & "F5VALVTA = " & dblBasePorItem & ", "
    strSQLOrden = strSQLOrden & "F3IGV = " & dblImpuestoPorItem & ", "
    strSQLOrden = strSQLOrden & "F3TOTAL = " & dblTotalPorItem & ", "
    strSQLOrden = strSQLOrden & "CODCOLOR = '" & strCodColor & "', "
    strSQLOrden = strSQLOrden & "F3OBSERVA = '" & strObservacionPorItem & "', "
    strSQLOrden = strSQLOrden & "F3AJUSTE = " & IIf(bolItemAjustado, "1", "0") & ", "
    strSQLOrden = strSQLOrden & "F3GASTO = " & IIf(strCodigoGasto <> vbNullString, "'" & strCodigoGasto & "'", "''") & ", "
    strSQLOrden = strSQLOrden & "F3CUENTA = " & IIf(strCuentaContable <> vbNullString, "'" & strCuentaContable & "'", "''") & " "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "ITEM = '" & dblItem & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errActualizarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: ActualizarOrdenDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    strSQLOrden = vbNullString
    
    Err.Clear
End Sub

Public Sub eliminarOrdenDetalleOneByOne()
    On Error GoTo errEliminarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    
    strSQLOrden = strSQLOrden & "DELETE FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "ITEM = '" & dblItem & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errEliminarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: EliminarOrdenDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    strSQLOrden = vbNullString
    
    Err.Clear
End Sub

Public Sub cerrarOrdenDetalleOneByOne()
    On Error GoTo errCerrarOrdenDetalleOneByOne
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
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
    
    cnBdCPlus.Execute strSQLOrden
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Sub
errCerrarOrdenDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: CerrarOrdenDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    strSQLOrden = vbNullString
    
    Err.Clear
End Sub

Public Function eliminarOrden() As Boolean
    On Error GoTo errEliminarOrden
    
'    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F2CODALM", "PROCESOS.IF3ORDEN", strNumeroVale, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarOrden = False
'
'        Exit Function
'    End If
    
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "DELETE "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "DELETE "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    eliminarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errEliminarOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: EliminarOrden"
    
    eliminarOrden = False
    
    Err.Clear
End Function

Public Function anularOrden() As Boolean
    On Error GoTo errAnularOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = 8, "
        strSQLOrden = strSQLOrden & "F4ESTNUL = 'S', "
        strSQLOrden = strSQLOrden & "F4BASIMP = 0, "
        strSQLOrden = strSQLOrden & "F4MONINA = 0, "
        strSQLOrden = strSQLOrden & "F4IGV = 0, "
        strSQLOrden = strSQLOrden & "F4MONTO = 0, "
        strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "', "
        strSQLOrden = strSQLOrden & "F4FECMOD = '" & strFecMod & "' "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    Actualiza_Log " < Replicación BD > " & strSQLOrden, StrConexDbBancos
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
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
    
    cnBdCPlus.Execute strSQLOrden
    
    Actualiza_Log " < Replicación BD > " & strSQLOrden, StrConexDbBancos
    
    anularOrden = True
    
    strSQLOrden = vbNullString
    
    Exit Function
errAnularOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: AnularOrden"
    
    anularOrden = False
    
    Err.Clear
End Function

Public Function aprobarOrden() As Boolean
    On Error GoTo errAprobarOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & ", "
        strSQLOrden = strSQLOrden & "F4VB1 = " & IIf(bolVB1, "1", "0") & ", "
        strSQLOrden = strSQLOrden & "F4VBUSER1 = " & IIf(strVB1Usuario <> vbNullString, "'" & strVB1Usuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & "F4VBFECHA1 = " & IIf(strVB1Fecha <> vbNullString, "CONVERT(DATETIME, '" & Format(strVB1Fecha, "dd-mm-yy hh:mm:ss") & "', 5)", "NULL") & " "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    aprobarOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errAprobarOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: AprobarOrden"
    
    aprobarOrden = False
    
    Err.Clear
End Function

Public Function enviarViaMailOrden() As Boolean
    On Error GoTo errEnviarViaMailOrden
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "UPDATE "
        strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "SET "
        strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADA = " & IIf(bolColocada, "1", "0") & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADAUSER = " & IIf(strColocadaUsuario <> vbNullString, "'" & strColocadaUsuario & "'", "''") & ", "
        strSQLOrden = strSQLOrden & "F4COLOCADAFECHA = " & IIf(strColocadaFecha <> vbNullString, "'" & strColocadaFecha & "'", "NULL") & " "
    strSQLOrden = strSQLOrden & "WHERE "
        strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
        strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
    
    cnBdCPlus.Execute strSQLOrden
    
    enviarViaMailOrden = True
    
    strSQLSelectAlter = strSQLOrden
    strSQLOrden = vbNullString
    
    Exit Function
errEnviarViaMailOrden:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: EnviarViaMailOrden"
    
    enviarViaMailOrden = False
    
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
    
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) - CONVERT(DECIMAL(10, 2), (ISNULL(RESUMEN.CANTIDADTOTAL, 0) / ISNULL(MEDALTER.F5FACTOR, 1) ) ) AS SALDO "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) ) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
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
    
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) > 0 "
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
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
                    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F4ESTADO = " & intEstado & " "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
                
                cnBdCPlus.Execute strSQLOrden
                
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
            vbCritical, App.ProductName & " - SqlClsOrden: AtencionOrden"
    
    atencionOrden = False
    
    Err.Clear
End Function

Public Function cerrarOrden(Optional ByVal bolCierrePorItemDeOrden As Boolean) As Boolean
    On Error GoTo errCerrarOrden
    
    Set rstOrden = New ADODB.Recordset
    
    cerrarOrden = False
    
'    strSQLOrden = vbNullString
'    strSQLOrden = strSQLOrden & "SELECT "
'    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
'    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
'    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
'    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
'    strSQLOrden = strSQLOrden & "DET.F3CANPRO, "
'    strSQLOrden = strSQLOrden & "DET.F3PORCDEMASIA, "
'
'    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA / 100))) AS CANTIDAD, "
'    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (DET.F3PORCDEMASIA / 100)) AS DEMASIA, "
'    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA / 100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS SALDO "
'
'    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) AS CANTIDAD, "
'    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (DET.F3PORCDEMASIA/100) ) AS DEMASIA, "
'    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) - CONVERT(DECIMAL(10, 2), (ISNULL(RESUMEN.CANTIDADTOTAL, 0) / ISNULL(MEDALTER.F5FACTOR, 1) )) AS SALDO "
'
'    strSQLOrden = strSQLOrden & "FROM "
'    strSQLOrden = strSQLOrden & "(((PROCESOS.IF3ORDEN AS DET "
'    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN "
'    strSQLOrden = strSQLOrden & "(SELECT "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
'    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
'    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'S' THEN -1 ELSE 1 END)) AS CANTIDADTOTAL "
'    strSQLOrden = strSQLOrden & "FROM "
'    strSQLOrden = strSQLOrden & "(((PROCESOS.IF3VALES AS DET "
'    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
'    strSQLOrden = strSQLOrden & "WHERE "
'    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
'    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
'    strSQLOrden = strSQLOrden & "GROUP BY "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
'    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
'    strSQLOrden = strSQLOrden & "ON RESUMEN.F4NUMORD = DET.F4NUMORD AND RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD AND RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO "
'    strSQLOrden = strSQLOrden & "WHERE "
'    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
'
'    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) <> 0 AND "
'    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) > 0 "
'
'    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) <> 0 AND "
'    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) - CONVERT(DECIMAL(10, 2), (ISNULL(RESUMEN.CANTIDADTOTAL, 0) / ISNULL(MEDALTER.F5FACTOR, 1) )) > 0 "
'
'        If bolCierrePorItemDeOrden Then
'            strSQLOrden = strSQLOrden & "AND "
'            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD = '" & strRequerimiento & "' AND "
'            strSQLOrden = strSQLOrden & "DET.F3CODPRO = '" & strCodProducto & "' "
'        End If
'
'    strSQLOrden = strSQLOrden & "ORDER BY "
'    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"

    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "OCS.TIPO, "
    strSQLOrden = strSQLOrden & "OCS.OC, "
    strSQLOrden = strSQLOrden & "OCS.NROPEDIDO, "
    strSQLOrden = strSQLOrden & "OCS.CODPRODUCTO, "
    strSQLOrden = strSQLOrden & "OCS.F3CODFAB, "
    strSQLOrden = strSQLOrden & "OCS.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "OCS.F7SIGMED, "
    strSQLOrden = strSQLOrden & "OCS.CANTIDAD, "
    strSQLOrden = strSQLOrden & "OCS.DEMASIA, "
    strSQLOrden = strSQLOrden & "OCS.FACTOR, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) AS INGR, "
    strSQLOrden = strSQLOrden & "(CASE WHEN OCS.CANTIDAD < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.CANTIDAD - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) AS SALDO, "
    strSQLOrden = strSQLOrden & "OCS.F3PRECOS "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL AS TIPO, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD AS OC, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD AS NROPEDIDO, "
    strSQLOrden = strSQLOrden & "(CASE WHEN DET.F4LOCAL = 'OC' THEN (CASE WHEN DET.COD_SOLICITUD <> '' THEN DET.COD_SOLICITUD + ' / ' + REQ.CS_NOMREF ELSE 'STOCK LIBRE' END) ELSE DET.COD_SOLICITUD END) AS COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODPRODUCTO, "
    strSQLOrden = strSQLOrden & "DET.F3CODFAB, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING AS F5NOMPRO, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "SUM(CASE WHEN DET.F4LOCAL = 'OC' THEN "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) "
    strSQLOrden = strSQLOrden & "ELSE "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO "
    strSQLOrden = strSQLOrden & "END) AS CANTIDAD, "
    
    strSQLOrden = strSQLOrden & "SUM(CASE WHEN DET.F4LOCAL = 'OC' THEN "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO * (DET.F3PORCDEMASIA/100) "
    strSQLOrden = strSQLOrden & "ELSE "
    strSQLOrden = strSQLOrden & "0 "
    strSQLOrden = strSQLOrden & "END) AS DEMASIA, "
    
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1) AS FACTOR, "
    strSQLOrden = strSQLOrden & "DET.F3PRENETO AS F3PRECOS "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "(CASE WHEN DET.F4LOCAL = 'OC' THEN "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) "
    strSQLOrden = strSQLOrden & "ELSE "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO "
    strSQLOrden = strSQLOrden & "END) > 0 "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "(CASE WHEN DET.F4LOCAL = 'OC' THEN (CASE WHEN DET.COD_SOLICITUD <> '' THEN DET.COD_SOLICITUD + ' / ' + REQ.CS_NOMREF ELSE 'STOCK LIBRE' END) ELSE DET.COD_SOLICITUD END), "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F3CODFAB, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1), "
    strSQLOrden = strSQLOrden & "DET.F3PRENETO) AS OCS "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'I' THEN 1 ELSE -1 END)) ) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS INGRESOS "
    strSQLOrden = strSQLOrden & "ON INGRESOS.F4NUMORD = OCS.OC AND INGRESOS.COD_SOLICITUD = OCS.NROPEDIDO AND INGRESOS.F5CODPROORIGINAL = OCS.CODPRODUCTO "
    
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "OCS.TIPO = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "OCS.OC = '" & strNumeroOrden & "' AND "
    
    strSQLOrden = strSQLOrden & "OCS.CANTIDAD > 0 AND "
    strSQLOrden = strSQLOrden & "(CASE WHEN OCS.CANTIDAD < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.CANTIDAD - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) > 0 "
    
        If bolCierrePorItemDeOrden Then
            strSQLOrden = strSQLOrden & "AND "
            strSQLOrden = strSQLOrden & "OCS.NROPEDIDO = '" & strRequerimiento & "' AND "
            strSQLOrden = strSQLOrden & "OCS.CODPRODUCTO = '" & strCodProducto & "' "
        End If
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "OCS.F5NOMPRO"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstOrden.EOF Then
        rstOrden.MoveFirst
        
        Do While Not rstOrden.EOF
            If Val(rstOrden!SALDO & "") > 0 Then
                strSQLOrden = vbNullString
                strSQLOrden = strSQLOrden & "UPDATE "
                    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F3AJUSTE = 1, "
                    strSQLOrden = strSQLOrden & "F3PORCDEMASIA = 0, "
                    strSQLOrden = strSQLOrden & "F3CANFAL = " & Val(rstOrden!SALDO & "") - Val(rstOrden!DEMASIA & "") & ", "
                    strSQLOrden = strSQLOrden & "F3CANPRO2 = " & Val(rstOrden!Cantidad & "") & ", "
                    strSQLOrden = strSQLOrden & "F3CANPRO = " & Val(rstOrden!Cantidad & "") - (Val(rstOrden!SALDO & "") - Val(rstOrden!DEMASIA & "")) & " "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "' AND "
                    strSQLOrden = strSQLOrden & "COD_SOLICITUD = '" & Trim(rstOrden!NroPedido & "") & "' AND "
                    strSQLOrden = strSQLOrden & "F3CODPRO = '" & Trim(rstOrden!CodProducto & "") & "'"
                
                cnBdCPlus.Execute strSQLOrden
                
                Actualiza_Log " < Replicación BD > " & strSQLOrden, StrConexDbBancos
            End If
            
            rstOrden.MoveNext
        Loop
            If Not bolCierrePorItemDeOrden Then
                strSQLOrden = vbNullString
                strSQLOrden = strSQLOrden & "UPDATE "
                    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
                strSQLOrden = strSQLOrden & "SET "
                    strSQLOrden = strSQLOrden & "F4ESTADO = 7, "
                    strSQLOrden = strSQLOrden & "F4USEMOD = '" & strUsuMod & "', "
                    strSQLOrden = strSQLOrden & "F4FECMOD = '" & strFecMod & "' "
                strSQLOrden = strSQLOrden & "WHERE "
                    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
                    strSQLOrden = strSQLOrden & "F4NUMORD = '" & strNumeroOrden & "'"
                
                cnBdCPlus.Execute strSQLOrden
                
                Actualiza_Log " < Replicación BD > " & strSQLOrden, StrConexDbBancos
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
            vbCritical, App.ProductName & " - SqlClsOrden: CerrarOrden"
    'Resume
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
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "F4CODPRV = '" & strRucProveedor & "' "
        
        If strEstadoOrden <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND F4ESTADO IN (" & strEstadoOrden & ") "
        End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "F4NUMORD"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: ListarOrdenParaSeleccion"
    
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
    
    strSQLOrden = strSQLOrden & "PROCESOS.IF4ORDEN AS CAB "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV "
    
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
            cnBdCPlus.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsOrden: ListarGrillaOrden"
    End Select
    
    Err.Clear
End Sub

Public Function generarCadenaSolicitud() As String
    Dim rstSolicitudOC As New ADODB.Recordset
    
    generarCadenaSolicitud = vbNullString
    
    If rstSolicitudOC.State = 1 Then rstSolicitudOC.Close
    
    rstSolicitudOC.Open "SELECT COD_SOLICITUD FROM PROCESOS.IF3ORDEN WHERE F4LOCAL = '" & strTipoOrden & "' AND F4NUMORD = '" & strNumeroOrden & "' GROUP BY COD_SOLICITUD", cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSolicitudOC.EOF Then
        rstSolicitudOC.MoveFirst
        
        Do While Not rstSolicitudOC.EOF
            If generarCadenaSolicitud = vbNullString Then
                generarCadenaSolicitud = IIf(Trim(rstSolicitudOC!COD_SOLICITUD & "") = vbNullString, "STOCK LIBRE", Trim(rstSolicitudOC!COD_SOLICITUD & "") & " - " & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CS_NOMREF", "PROCESOS.TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & strTipoOrden & "'"))
            Else
                generarCadenaSolicitud = generarCadenaSolicitud & "; " & IIf(Trim(rstSolicitudOC!COD_SOLICITUD & "") = vbNullString, "STOCK LIBRE", Trim(rstSolicitudOC!COD_SOLICITUD & "") & "-" & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CS_NOMREF", "PROCESOS.TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & strTipoOrden & "'"))
            End If
            
            rstSolicitudOC.MoveNext
        Loop
    End If
    
    Screen.MousePointer = vbDefault
End Function

Public Function obtenerUltimaDescripcionProductoDeProveedor() As String
    On Error GoTo errObtenerUltimaDescripcionProductoDeProveedor
    
    Dim cmdOP As ADODB.Command
    
    Set rstOrden = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ObtenerUltimaDescripcionProductoProveedor"
        
        .Parameters.Append .CreateParameter("@RUCPROVEEDOR", adVarChar, adParamInput, 20, strRucProveedor)
        .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        
        Set rstOrden = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstOrden.EOF Then
        obtenerUltimaDescripcionProductoDeProveedor = Trim(rstOrden!ULTIMADESCRIPCIONPROVEEDOR & "")
    Else
        obtenerUltimaDescripcionProductoDeProveedor = vbNullString
    End If
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    strSQLOrden = vbNullString
    
    Set rstOrden = Nothing
    
    Exit Function
errObtenerUltimaDescripcionProductoDeProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerUltimaDescripcionProductoDeProveedor"
    
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
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
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
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerUltimoDescuentoProductoDeProveedor"
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: CalculosPorItem"
    
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
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'S' THEN -1 ELSE 1)) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    'strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "'"
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL"
    
    If rstOrden.State = 1 Then rstOrden.Close
    
    rstOrden.Open strSQLOrden, cnBdCPlus, adOpenDynamic, adLockBatchOptimistic
    
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
            vbCritical, App.ProductName & " - SqlClsOrden: ObtenerMovimientoOrdenEnAlmacen"
    
    obtenerMovimientoOrdenEnAlmacen = False
    
    Err.Clear
End Function

Public Function devuelveSQLOrdenMovimiento() As String
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "OCS.TIPO AS F4LOCAL, "
    strSQLOrden = strSQLOrden & "OCS.OC AS F4NUMORD, "
    strSQLOrden = strSQLOrden & "OCS.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "OCS.F4FECEMI, "
    strSQLOrden = strSQLOrden & "OCS.F4FECENT, "
    strSQLOrden = strSQLOrden & "OCS.LLAVE, "
    strSQLOrden = strSQLOrden & "OCS.NROPEDIDO AS COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "OCS.CS_OBSERVACIONES, "
    strSQLOrden = strSQLOrden & "OCS.CODPRODUCTO AS F3CODPRO, "
    strSQLOrden = strSQLOrden & "OCS.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "OCS.F7SIGMED, "
    strSQLOrden = strSQLOrden & "SUM(OCS.F3CANPRO) AS F3CANPRO, "
    strSQLOrden = strSQLOrden & "OCS.FACTOR, "
    strSQLOrden = strSQLOrden & "SUM(CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) )) AS INGR, "
    strSQLOrden = strSQLOrden & "SUM(CASE WHEN OCS.F3CANPRO < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.F3CANPRO - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) AS SALDO, "
    
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F2NOMALM, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4FECVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F1NOMORI, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPRO, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "SUM(CONVERT(DECIMAL(10, 2), ISNULL(MOVIMIENTO.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) )) AS CANTIDAD "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL AS TIPO, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD AS OC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "CONVERT(NVARCHAR(10), CAB.F4FECEMI,103) AS F4FECEMI, "
    strSQLOrden = strSQLOrden & "CONVERT(NVARCHAR(10), CAB.F4FECENT,103) AS F4FECENT, "
    
    strSQLOrden = strSQLOrden & "(DET.F4NUMORD + DET.COD_SOLICITUD + DET.F3CODPRO) AS LLAVE,  "
    
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD AS NROPEDIDO, "
    strSQLOrden = strSQLOrden & "(CASE WHEN DET.COD_SOLICITUD <> '' THEN REQ.CS_NOMREF ELSE '< REQUERIDO PARA STOCK LIBRE >' END) AS CS_OBSERVACIONES, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODPRODUCTO,  "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS F3CANPRO, "
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1) AS FACTOR "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) > 0 "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    
    strSQLOrden = strSQLOrden & "(DET.F4NUMORD + DET.COD_SOLICITUD + DET.F3CODPRO), "
    
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "(CASE WHEN DET.COD_SOLICITUD <> '' THEN REQ.CS_NOMREF ELSE '< REQUERIDO PARA STOCK LIBRE >' END), "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1)) AS OCS "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'I' THEN 1 ELSE -1 END)) ) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS INGRESOS "
    strSQLOrden = strSQLOrden & "ON INGRESOS.F4NUMORD = OCS.OC AND INGRESOS.COD_SOLICITUD = OCS.NROPEDIDO AND INGRESOS.F5CODPROORIGINAL = OCS.CODPRODUCTO "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "ALM.F2NOMALM, "
    strSQLOrden = strSQLOrden & "DET.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "CONVERT(NVARCHAR(10), CAB.F4FECVAL,103) AS F4FECVAL, "
    strSQLOrden = strSQLOrden & "ORI.F1NOMORI, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPRO, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "INS.F5NOMPRO, "
    strSQLOrden = strSQLOrden & "DET.F3CANPRO * (CASE WHEN DET.TIPO = 'S' THEN -1 ELSE 1 END) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD = '" & strNumeroOrden & "' "
    strSQLOrden = strSQLOrden & ") AS MOVIMIENTO "
    strSQLOrden = strSQLOrden & "ON MOVIMIENTO.F4NUMORD = OCS.OC AND MOVIMIENTO.COD_SOLICITUD = OCS.NROPEDIDO AND MOVIMIENTO.F5CODPROORIGINAL = OCS.CODPRODUCTO "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "OCS.TIPO, "
    strSQLOrden = strSQLOrden & "OCS.OC, "
    strSQLOrden = strSQLOrden & "OCS.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "OCS.F4FECEMI, "
    strSQLOrden = strSQLOrden & "OCS.F4FECENT, "
    strSQLOrden = strSQLOrden & "OCS.LLAVE, "
    strSQLOrden = strSQLOrden & "OCS.NROPEDIDO, "
    strSQLOrden = strSQLOrden & "OCS.CS_OBSERVACIONES, "
    strSQLOrden = strSQLOrden & "OCS.CODPRODUCTO, "
    strSQLOrden = strSQLOrden & "OCS.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "OCS.F7SIGMED, "
    strSQLOrden = strSQLOrden & "OCS.FACTOR, "
    
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F2NOMALM, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4NUMVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F4FECVAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F1NOMORI, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPRO, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "MOVIMIENTO.F5NOMPRO "
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "OCS.F5NOMPRO_ING, "
    strSQLOrden = strSQLOrden & "OCS.NROPEDIDO, "
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
    strSQLOrden = strSQLOrden & "IIF((GETDATE() - CAB.F4FECENT) = 0, 'PLAZO DE ENTREGA VENCE HOY', IIF((DATE() - CAB.F4FECENT) < 0, 'PLAZO DE ENTREGA A VENCER EN ' & (DATE() - CAB.F4FECENT) * -1 & ' DIA(S).',  'PLAZO DE ENTREGA VENCIDO HACE ' & (DATE() - CAB.F4FECENT) & ' DIA(S).')) AS DIASVENCIMIENTO, "
    
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "SUM(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000'))) AS SALDO "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "((((PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
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
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
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
    strSQLOrden = strSQLOrden & "((((PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
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
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
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
                                                                Optional ByVal bolMostrarNroRequerimiento As Boolean, _
                                                                Optional ByVal bolIncluirOCAtencionTotalYOrdenCerrada As Boolean, _
                                                                Optional ByVal strDesde As String, _
                                                                Optional ByVal strHasta As String) As String
    
    strSQLOrden = vbNullString
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "OCS.RUC, "
    strSQLOrden = strSQLOrden & "OCS.PROVEEDOR, "
    strSQLOrden = strSQLOrden & "OCS.ANNO, "
    strSQLOrden = strSQLOrden & "OCS.CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "OCS.DESCRIPCIONINTERNA, "
    strSQLOrden = strSQLOrden & "OCS.UM, "
    
        If Not bolVerConsolidado Then
            strSQLOrden = strSQLOrden & "OCS.DESCRIPCIONPROV, "
            strSQLOrden = strSQLOrden & "OCS.EMISION, "
            strSQLOrden = strSQLOrden & "OCS.ENTREGA, "
            strSQLOrden = strSQLOrden & "OCS.OC AS NROOC, "
            
            If bolMostrarNroRequerimiento Then
                strSQLOrden = strSQLOrden & "OCS.NROPEDIDO AS COD_SOLICITUD, "
            End If
        End If
    
    strSQLOrden = strSQLOrden & "SUM(OCS.F3CANPRO) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "SUM(CASE WHEN OCS.F3CANPRO < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.F3CANPRO - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) AS SALDO "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC AS RUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV AS PROVEEDOR, "
    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI) AS ANNO, "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "LEFT(RTRIM(LTRIM(DET.F5NOMPRO_ING)), 255) AS DESCRIPCIONINTERNA, "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED AS UM, "
    
    strSQLOrden = strSQLOrden & "LEFT(RTRIM(LTRIM(DET.F5NOMPRO)), 255) AS DESCRIPCIONPROV, "
    strSQLOrden = strSQLOrden & "CONVERT(NVARCHAR(10), CAB.F4FECEMI,103) AS EMISION, "
    strSQLOrden = strSQLOrden & "CONVERT(NVARCHAR(10), CAB.F4FECENT,103) AS ENTREGA, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD AS OC, "
    
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD AS NROPEDIDO, "
    
    strSQLOrden = strSQLOrden & "SUM(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS F3CANPRO, "
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1) AS FACTOR "
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "DET.F4LOCAL = 'OC' AND "
    strSQLOrden = strSQLOrden & "PROV.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLOrden = strSQLOrden & "CAB.F4ESTADO NOT IN (1, " & IIf(bolIncluirOCAtencionTotalYOrdenCerrada, vbNullString, "6, 7, ") & "8) AND "
    strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) > 0 "
        
        If strCodProducto <> vbNullString Then
            strSQLOrden = strSQLOrden & "AND DET.F3CODPRO = '" & strCodProducto & "' "
        End If
        
        If bolIncluirOCAtencionTotalYOrdenCerrada Then
            'strSQLOrden = strSQLOrden & "AND (CONVERT(NVARCHAR(10), CAB.F4FECEMI,103) BETWEEN '" & strDesde & "' AND '" & strHasta & "') "
            strSQLOrden = strSQLOrden & "AND (CAB.F4FECEMI BETWEEN '" & strDesde & "' AND '" & strHasta & "') "
        End If
        
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC, "
    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI), "
    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
    strSQLOrden = strSQLOrden & "LEFT(RTRIM(LTRIM(DET.F5NOMPRO_ING)), 255), "
    strSQLOrden = strSQLOrden & "MED.F7SIGMED, "
    
    strSQLOrden = strSQLOrden & "LEFT(RTRIM(LTRIM(DET.F5NOMPRO)), 255), "
    strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
    strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    
    strSQLOrden = strSQLOrden & "ISNULL(MEDALTER.F5FACTOR, 1)) AS OCS "
    
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "("
    strSQLOrden = strSQLOrden & "SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'I' THEN 1 ELSE -1 END)) ) AS CANTIDAD "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "WHERE "
    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') "
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS INGRESOS "
    strSQLOrden = strSQLOrden & "ON "
    strSQLOrden = strSQLOrden & "INGRESOS.F4NUMORD = OCS.OC AND "
    strSQLOrden = strSQLOrden & "INGRESOS.COD_SOLICITUD = OCS.NROPEDIDO AND "
    strSQLOrden = strSQLOrden & "INGRESOS.F5CODPROORIGINAL = OCS.CODIGOPRODUCTO "
    
    strSQLOrden = strSQLOrden & "GROUP BY "
    strSQLOrden = strSQLOrden & "OCS.RUC, "
    strSQLOrden = strSQLOrden & "OCS.PROVEEDOR, "
    strSQLOrden = strSQLOrden & "OCS.ANNO, "
    strSQLOrden = strSQLOrden & "OCS.CODIGOPRODUCTO, "
    strSQLOrden = strSQLOrden & "OCS.DESCRIPCIONINTERNA, "
    strSQLOrden = strSQLOrden & "OCS.UM"
    
        If Not bolVerConsolidado Then
            strSQLOrden = strSQLOrden & ", OCS.DESCRIPCIONPROV, "
            strSQLOrden = strSQLOrden & "OCS.EMISION, "
            strSQLOrden = strSQLOrden & "OCS.ENTREGA, "
            strSQLOrden = strSQLOrden & "OCS.OC"
            
            If bolMostrarNroRequerimiento Then
                strSQLOrden = strSQLOrden & ", OCS.NROPEDIDO"
            End If
        End If
    
    strSQLOrden = strSQLOrden & " HAVING "
    strSQLOrden = strSQLOrden & "SUM(OCS.F3CANPRO) > 0 "
        
        If bolSoloProductoConSaldo Then
            strSQLOrden = strSQLOrden & "AND SUM(CASE WHEN OCS.F3CANPRO < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.F3CANPRO - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) > 0 "
        End If
    
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "OCS.ANNO, "
    
        If Not bolVerConsolidado Then
            strSQLOrden = strSQLOrden & "OCS.DESCRIPCIONINTERNA, "
            strSQLOrden = strSQLOrden & "OCS.UM"
            
            strSQLOrden = strSQLOrden & ", OCS.OC, "
            strSQLOrden = strSQLOrden & "OCS.EMISION, "
            strSQLOrden = strSQLOrden & "OCS.ENTREGA"
            
            If bolMostrarNroRequerimiento Then
                strSQLOrden = strSQLOrden & ", OCS.NROPEDIDO"
            End If
        Else
            strSQLOrden = strSQLOrden & "OCS.UM, "
            strSQLOrden = strSQLOrden & "OCS.DESCRIPCIONINTERNA"
        End If
    
    
    
    
'    strSQLOrden = strSQLOrden & "SELECT "
'    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC AS RUC, "
'    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV AS PROVEEDOR, "
'    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI) AS ANNO, "
'    strSQLOrden = strSQLOrden & "DET.F3CODPRO AS CODIGOPRODUCTO, "
'    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255) AS DESCRIPCIONINTERNA, "
'    strSQLOrden = strSQLOrden & "MED.F7SIGMED AS UM, "
'
'    If Not bolVerConsolidado Then
'        strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO & ''), 255) AS DESCRIPCIONPROV, "
'        strSQLOrden = strSQLOrden & "CAB.F4FECEMI AS EMISION, "
'        strSQLOrden = strSQLOrden & "CAB.F4FECENT AS ENTREGA, "
'        strSQLOrden = strSQLOrden & "DET.F4NUMORD AS NROOC, "
'
'        If bolMostrarNroRequerimiento Then
'            strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'        End If
'    End If
'
'    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT(DET.F3CANPRO, '#0.0000') & '') * (1 + (DET.F3PORCDEMASIA/100))  ) AS CANTIDAD, "
'
'    'strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
'    'strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) ) AS SALDO, "
'
'    strSQLOrden = strSQLOrden & "SUM("
'    strSQLOrden = strSQLOrden & "IIF(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
'    strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) < 0, 0, "
'    strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
'    strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) )"
'    strSQLOrden = strSQLOrden & ") AS SALDO "
'
'    strSQLOrden = strSQLOrden & "FROM "
'    strSQLOrden = strSQLOrden & "((((PROCESOS.IF3ORDEN AS DET "
'    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON (CAB.F4NUMORD = DET.F4NUMORD) AND (CAB.F4LOCAL = DET.F4LOCAL)) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN EF2PROVEEDORES AS PROV ON PROV.F2NEWRUC = CAB.F4CODPRV) "
'    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
'
'    strSQLOrden = strSQLOrden & "LEFT JOIN MEDIVENTAS AS MEDALTER ON (MEDALTER.F7CODMED = DET.UNIDAD) AND (MEDALTER.F5CODPRO = DET.F3CODPRO)) "
'
'    strSQLOrden = strSQLOrden & "LEFT JOIN "
'    strSQLOrden = strSQLOrden & "(SELECT "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
'    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
'    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT( DET.F3CANPRO, '#0.0000')) * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDADTOTAL "
'    strSQLOrden = strSQLOrden & "FROM "
'    strSQLOrden = strSQLOrden & "IF3VALES AS DET "
'    strSQLOrden = strSQLOrden & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'    strSQLOrden = strSQLOrden & "WHERE "
'    strSQLOrden = strSQLOrden & "CAB.F1CODORI IN ('XC0') "
'    'strSQLOrden = strSQLOrden & "CAB.F2CODPROV = '" & strCodProveedor & "' "
'    strSQLOrden = strSQLOrden & "GROUP BY "
'    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
'    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
'    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL) AS RESUMEN "
'    strSQLOrden = strSQLOrden & "ON (RESUMEN.F5CODPROORIGINAL = DET.F3CODPRO) AND (RESUMEN.COD_SOLICITUD = DET.COD_SOLICITUD) AND (RESUMEN.F4NUMORD = DET.F4NUMORD) "
'
'    strSQLOrden = strSQLOrden & "WHERE "
'    strSQLOrden = strSQLOrden & "CAB.F4LOCAL = 'OC' AND "
'    strSQLOrden = strSQLOrden & "PROV.F2CODPROV = '" & strCodProveedor & "' AND "
'    strSQLOrden = strSQLOrden & "CAB.F4ESTADO > 1 "
'
'    strSQLOrden = strSQLOrden & "GROUP BY "
'    strSQLOrden = strSQLOrden & "PROV.F2NEWRUC, "
'    strSQLOrden = strSQLOrden & "PROV.F2NOMPROV, "
'    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI), "
'    strSQLOrden = strSQLOrden & "DET.F3CODPRO, "
'    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255), "
'    strSQLOrden = strSQLOrden & "MED.F7SIGMED "
'
'    If Not bolVerConsolidado Then
'        strSQLOrden = strSQLOrden & ", LEFT(TRIM(DET.F5NOMPRO & ''), 255), "
'        strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
'        strSQLOrden = strSQLOrden & "CAB.F4FECENT, "
'        strSQLOrden = strSQLOrden & "DET.F4NUMORD "
'
'        If bolMostrarNroRequerimiento Then
'            strSQLOrden = strSQLOrden & ", DET.COD_SOLICITUD "
'        End If
'    End If
'
'    strSQLOrden = strSQLOrden & "HAVING "
'    strSQLOrden = strSQLOrden & "SUM( VAL(FORMAT(DET.F3CANPRO, '#0.0000') & '') * (1 + (DET.F3PORCDEMASIA/100))  ) > 0 "
'
'    If bolSoloProductoConSaldo Then
'        strSQLOrden = strSQLOrden & "AND "
'        strSQLOrden = strSQLOrden & "SUM("
'        strSQLOrden = strSQLOrden & "IIF(VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
'        strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) < 0, 0, "
'        strSQLOrden = strSQLOrden & "VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - "
'        strSQLOrden = strSQLOrden & "(VAL(RESUMEN.CANTIDADTOTAL & '')  / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')) )"
'        strSQLOrden = strSQLOrden & ") > 0 "
'    End If
'
'    strSQLOrden = strSQLOrden & "ORDER BY "
'    strSQLOrden = strSQLOrden & "YEAR(CAB.F4FECEMI), "
'    strSQLOrden = strSQLOrden & "LEFT(TRIM(DET.F5NOMPRO_ING & ''), 255), "
'    strSQLOrden = strSQLOrden & "MED.F7SIGMED"
'
'    If Not bolVerConsolidado Then
'        strSQLOrden = strSQLOrden & ", DET.F4NUMORD, "
'        strSQLOrden = strSQLOrden & "CAB.F4FECEMI, "
'        strSQLOrden = strSQLOrden & "CAB.F4FECENT"
'
'        If bolMostrarNroRequerimiento Then
'            strSQLOrden = strSQLOrden & ", DET.COD_SOLICITUD"
'        End If
'    End If
    
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
    
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS CANTIDAD, "
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(RESUMEN.CANTIDADTOTAL & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS SALDO "
    
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) AS CANTIDAD, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) - CONVERT(DECIMAL(10, 2), (ISNULL(RESUMEN.CANTIDADTOTAL, 0) / ISNULL(MEDALTER.F5FACTOR, 1) ) ) AS SALDO "
    
    
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3ORDEN AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    strSQLOrden = strSQLOrden & "LEFT JOIN "
    strSQLOrden = strSQLOrden & "(SELECT "
    strSQLOrden = strSQLOrden & "DET.F4NUMORD, "
    strSQLOrden = strSQLOrden & "DET.COD_SOLICITUD, "
    strSQLOrden = strSQLOrden & "DET.F5CODPROORIGINAL, "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) ) AS CANTIDADTOTAL "
    strSQLOrden = strSQLOrden & "FROM "
    strSQLOrden = strSQLOrden & "PROCESOS.IF3VALES AS DET "
    strSQLOrden = strSQLOrden & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLOrden = strSQLOrden & "LEFT JOIN MAESTROS.IF5PLA AS INS ON INS.F5CODPRO = DET.F5CODPRO "
    strSQLOrden = strSQLOrden & "WHERE "
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
    
    'strSQLOrden = strSQLOrden & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) <> 0 "
    strSQLOrden = strSQLOrden & "CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) <> 0 "
        
        If bolSoloProductosConSaldoPorEntregar Then
            strSQLOrden = strSQLOrden & "AND (CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) ) - CONVERT(DECIMAL(10, 2), (ISNULL(RESUMEN.CANTIDADTOTAL, 0) / ISNULL(MEDALTER.F5FACTOR, 1) ) )) > 0 "
        End If
        
    strSQLOrden = strSQLOrden & "ORDER BY "
    strSQLOrden = strSQLOrden & "DET.F5NOMPRO_ING"
    
    devuelveSQLAtencionOrden = strSQLOrden
    
    strSQLOrden = vbNullString
End Function
