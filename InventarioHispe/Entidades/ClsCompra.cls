VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strMesMovimiento                As String 'F4MESMOV
Private strNumeroMovimiento             As String 'F4NUMMOV

Private strCodProveedor                 As String 'F4CODPRV
Private strNomProveedor                 As String 'F4NOMPRV
Private strDireccionProveedor           As String 'F4DIRPRV
Private strTelefonoProveedor            As String 'F4TELPRV
Private strRucProveedor                 As String 'F4RUCPRV
Private strTipoDocumento                As String 'F4TIPDOC
Private strSerieDocumento               As String 'F4SERDOC
Private strNumeroDocumento              As String 'F4NUMDOC
Private lngCodCategoria                 As Long 'INTCODCATEGORIA
Private strFechaRegistro                As String 'F4FECHAREC
Private strFechaDocumento               As String 'F4FECHA
Private strCodMoneda                    As String 'F4MONEDA
Private dblTipoCambio                   As Double 'F4TIPCAM
Private strConceptoCompra               As String 'F4REFERE
Private dblBaseImponible                As Double 'F4BASIMP
Private dblMontoInafecto                As Double 'F4MONINA
Private dblTotalIGV                     As Double 'F4IGV
Private dblOtroImpuesto                 As Double 'F4OTRIMP
Private dblTotalFacturado               As Double 'F4TOTAL
Private strCodFormaPago                 As String 'F4FORPAG
Private strFechaVencimiento             As String 'F4FECVEN
Private strCodigoGasto                  As String 'F4GRUPO
Private strCuentaContable               As String 'F4CTACONT
Private strOrdenCompra                  As String 'F4OCOMPRA
Private strCentroCosto                  As String 'F4OBRA
Private dblMontoRetenido                As Double 'F4MONTORET
Private dblMontoFonavi                  As Double 'F4FONAVI
Private bolAsientoContableProcesado     As Boolean 'F4CONTABLE
Private dblCorrelativo                  As Double 'F4CORRELA
Private dblRedSuma                      As Double 'F4REDSUMA
Private dblRedResta                     As Double 'F4REDRESTA
Private dblDescuento                    As Double 'F4DCTO
Private strTipoDocAuxiliar              As String 'TIPODOCAUX
Private strFechaDocReferencia           As String 'FECDOCREF
Private strTipoDocReferencia            As String 'TIPODOCREF
Private strSerieDocReferencia           As String 'SERDOCREF
Private strNumeroDocReferencia          As String 'NUMDOCREF
Private strAnnoDUA                      As String 'ANIODUA

Rem SK ADD:
Private dblDetraccionPorc               As Double 'PORC_DETR
Private strDetraccionNum                As String 'NUMDETRACCION
Private strDetraccionFecha              As String 'FECHADETRACCION

Private strNroComprobante               As String 'F4NROCOMPROBANTE
Private bolVB                           As Boolean 'F4VB
Private strVBUsuario                    As String 'F4VBUSER
Private dblPorcentajeIGV                As Double 'F4PORC_IGV
Private strValeIngreso                  As String 'F4VALESING


Private bolCierre                       As String 'F4CIERRE
Private strCierreUsuario                As String 'F4CIERREUSER
Private strCierreFecha                  As String 'F4CIERREFECHA


Private strTipoRegistro                 As String


Private strFecReg                      As String 'F4FECHING
Private strUsuReg                      As String 'F4USUARIOING
Private strFecMod                      As String 'F4FECHMOD
Private strUsuMod                      As String 'F4USUARIOMOD


Rem: Atributos de Detalle
Private dblItem                         As Double
Private strAnnoOrden                    As String
Private strNumeroOrden                  As String
Private strCodigoGastoDet               As String
Private strCtaContableDet               As String
Private strCentroCostoDet               As String
Private strCodigoProducto               As String
Private strConceptoDet                  As String
Private dblCantidad                     As Double
Private dblPrecioUnitario               As Double
Private dblSubTotalDet                  As Double
Private bolAfecto                       As Boolean
Private strDebHab                       As String
Private strSerieGuiaDet                 As String
Private strNumeroGuiaDet                As String

Private strAuxiliarDet                  As String


Private strSQLSelectAlter               As String

Private rstCompra                       As ADODB.Recordset
Private strSQLCompra                    As String

Public Property Let MesMovimiento(ByVal Value As String)
    strMesMovimiento = Value
End Property

Public Property Get MesMovimiento() As String
    MesMovimiento = strMesMovimiento
End Property

Public Property Let NumeroMovimiento(ByVal Value As String)
    strNumeroMovimiento = Value
End Property

Public Property Get NumeroMovimiento() As String
    NumeroMovimiento = strNumeroMovimiento
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

Public Property Let DireccionProveedor(ByVal Value As String)
    strDireccionProveedor = Value
End Property

Public Property Get DireccionProveedor() As String
    DireccionProveedor = strDireccionProveedor
End Property
'TelefonoProveedor
Public Property Let TelefonoProveedor(ByVal Value As String)
    strTelefonoProveedor = Value
End Property

Public Property Get TelefonoProveedor() As String
    TelefonoProveedor = strTelefonoProveedor
End Property

Public Property Let RucProveedor(ByVal Value As String)
    strRucProveedor = Value
End Property

Public Property Get RucProveedor() As String
    RucProveedor = strRucProveedor
End Property

'TipoDocumento
Public Property Let TipoDocumento(ByVal Value As String)
    strTipoDocumento = Value
End Property

Public Property Get TipoDocumento() As String
    TipoDocumento = strTipoDocumento
End Property

Public Property Let SerieDocumento(ByVal Value As String)
    strSerieDocumento = Value
End Property

Public Property Get SerieDocumento() As String
    SerieDocumento = strSerieDocumento
End Property

Public Property Let NumeroDocumento(ByVal Value As String)
    strNumeroDocumento = Value
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = strNumeroDocumento
End Property

'CodigoCategoria
Public Property Let CodigoCategoria(ByVal Value As Long)
    lngCodCategoria = Value
End Property

Public Property Get CodigoCategoria() As Long
    CodigoCategoria = lngCodCategoria
End Property

'FechaRegistro
Public Property Let FechaRegistro(ByVal Value As String)
    strFechaRegistro = Value
End Property

Public Property Get FechaRegistro() As String
    FechaRegistro = strFechaRegistro
End Property

Public Property Let FechaDocumento(ByVal Value As String)
    strFechaDocumento = Value
End Property

Public Property Get FechaDocumento() As String
    FechaDocumento = strFechaDocumento
End Property

Public Property Let CodMoneda(ByVal Value As String)
    strCodMoneda = Value
End Property

Public Property Get CodMoneda() As String
    CodMoneda = strCodMoneda
End Property

Public Property Let TipoCambio(ByVal Value As Double)
    dblTipoCambio = Value
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dblTipoCambio
End Property

Public Property Let ConceptoCompra(ByVal Value As String)
    strConceptoCompra = Value
End Property

Public Property Get ConceptoCompra() As String
    ConceptoCompra = strConceptoCompra
End Property

Public Property Let BaseImponible(ByVal Value As Double)
    dblBaseImponible = Value
End Property

Public Property Get BaseImponible() As Double
    BaseImponible = dblBaseImponible
End Property

Public Property Let MontoInafecto(ByVal Value As Double)
    dblMontoInafecto = Value
End Property

Public Property Get MontoInafecto() As Double
    MontoInafecto = dblMontoInafecto
End Property

Public Property Let TotalIGV(ByVal Value As Double)
    dblTotalIGV = Value
End Property

Public Property Get TotalIGV() As Double
    TotalIGV = dblTotalIGV
End Property

Public Property Let OtroImpuesto(ByVal Value As Double)
    dblOtroImpuesto = Value
End Property

Public Property Get OtroImpuesto() As Double
    OtroImpuesto = dblOtroImpuesto
End Property

Public Property Let TotalFacturado(ByVal Value As Double)
    dblTotalFacturado = Value
End Property

Public Property Get TotalFacturado() As Double
    TotalFacturado = dblTotalFacturado
End Property

Public Property Let CodFormaPago(ByVal Value As String)
    strCodFormaPago = Value
End Property

Public Property Get CodFormaPago() As String
    CodFormaPago = strCodFormaPago
End Property

'FechaVencimiento
Public Property Let FechaVencimiento(ByVal Value As String)
    strFechaVencimiento = Value
End Property

Public Property Get FechaVencimiento() As String
    FechaVencimiento = strFechaVencimiento
End Property

Public Property Let CodigoGasto(ByVal Value As String)
    strCodigoGasto = Value
End Property

Public Property Get CodigoGasto() As String
    CodigoGasto = strCodigoGasto
End Property

Public Property Let CuentaContable(ByVal Value As String)
    strCuentaContable = Value
End Property

Public Property Get CuentaContable() As String
    CuentaContable = strCuentaContable
End Property

Public Property Let OrdenCompra(ByVal Value As String)
    strOrdenCompra = Value
End Property

Public Property Get OrdenCompra() As String
    OrdenCompra = strOrdenCompra
End Property

Public Property Let CentroCosto(ByVal Value As String)
    strCentroCosto = Value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property
'MontoRetenido
Public Property Let MontoRetenido(ByVal Value As Double)
    dblMontoRetenido = Value
End Property

Public Property Get MontoRetenido() As Double
    MontoRetenido = dblMontoRetenido
End Property
'MontoFonavi
Public Property Let MontoFonavi(ByVal Value As Double)
    dblMontoFonavi = Value
End Property

Public Property Get MontoFonavi() As Double
    MontoFonavi = dblMontoFonavi
End Property

Public Property Let AsientoContableProcesado(ByVal Value As Boolean)
    bolAsientoContableProcesado = Value
End Property

Public Property Get AsientoContableProcesado() As Boolean
    AsientoContableProcesado = bolAsientoContableProcesado
End Property

Public Property Let Correlativo(ByVal Value As Double)
    dblCorrelativo = Value
End Property

Public Property Get Correlativo() As Double
    Correlativo = dblCorrelativo
End Property


Public Property Let RedSuma(ByVal Value As Double)
    dblRedSuma = Value
End Property

Public Property Get RedSuma() As Double
    RedSuma = dblRedSuma
End Property

Public Property Let RedResta(ByVal Value As Double)
    dblRedResta = Value
End Property

Public Property Get RedResta() As Double
    RedResta = dblRedResta
End Property

Public Property Let Descuento(ByVal Value As Double)
    dblDescuento = Value
End Property

Public Property Get Descuento() As Double
    Descuento = dblDescuento
End Property

Public Property Let TipoDocAuxiliar(ByVal Value As String)
    strTipoDocAuxiliar = Value
End Property

Public Property Get TipoDocAuxiliar() As String
    TipoDocAuxiliar = strTipoDocAuxiliar
End Property



Public Property Let FechaDocReferencia(ByVal Value As String)
    strFechaDocReferencia = Value
End Property

Public Property Get FechaDocReferencia() As String
    FechaDocReferencia = strFechaDocReferencia
End Property

Public Property Let TipoDocReferencia(ByVal Value As String)
    strTipoDocReferencia = Value
End Property

Public Property Get TipoDocReferencia() As String
    TipoDocReferencia = strTipoDocReferencia
End Property

Public Property Let SerieDocReferencia(ByVal Value As String)
    strSerieDocReferencia = Value
End Property

Public Property Get SerieDocReferencia() As String
    SerieDocReferencia = strSerieDocReferencia
End Property

Public Property Let NumeroDocReferencia(ByVal Value As String)
    strNumeroDocReferencia = Value
End Property

Public Property Get NumeroDocReferencia() As String
    NumeroDocReferencia = strNumeroDocReferencia
End Property

Public Property Let AnnoDUA(ByVal Value As String)
    strAnnoDUA = Value
End Property

Public Property Get AnnoDUA() As String
    AnnoDUA = strAnnoDUA
End Property

Public Property Let DetraccionPorc(ByVal Value As Double)
    dblDetraccionPorc = Value
End Property

Public Property Get DetraccionPorc() As Double
    DetraccionPorc = dblDetraccionPorc
End Property

Public Property Let DetraccionNum(ByVal Value As String)
    strDetraccionNum = Value
End Property

Public Property Get DetraccionNum() As String
    DetraccionNum = strDetraccionNum
End Property

Public Property Let DetraccionFecha(ByVal Value As String)
    strDetraccionFecha = Value
End Property

Public Property Get DetraccionFecha() As String
    DetraccionFecha = strDetraccionFecha
End Property
'NroComprobante
Public Property Let NroComprobante(ByVal Value As String)
    strNroComprobante = Value
End Property

Public Property Get NroComprobante() As String
    NroComprobante = strNroComprobante
End Property

Public Property Let VB(ByVal Value As Boolean)
    bolVB = Value
End Property

Public Property Get VB() As Boolean
    VB = bolVB
End Property

Public Property Let VBUsuario(ByVal Value As String)
    strVBUsuario = Value
End Property

Public Property Get VBUsuario() As String
    VBUsuario = strVBUsuario
End Property

Public Property Let PorcentajeIGV(ByVal Value As Double)
    dblPorcentajeIGV = Value
End Property

Public Property Get PorcentajeIGV() As Double
    PorcentajeIGV = dblPorcentajeIGV
End Property

Public Property Let ValeIngreso(ByVal Value As String)
    strValeIngreso = Value
End Property

Public Property Get ValeIngreso() As String
    ValeIngreso = strValeIngreso
End Property




'F4CIERRE
Public Property Let Cierre(ByVal Value As Boolean)
    bolCierre = Value
End Property

Public Property Get Cierre() As Boolean
    Cierre = bolCierre
End Property
'F4CIERREUSER
Public Property Let CierreUsuario(ByVal Value As String)
    strCierreUsuario = Value
End Property

Public Property Get CierreUsuario() As String
    CierreUsuario = strCierreUsuario
End Property
'F4CIERREFECHA
Public Property Let CierreFecha(ByVal Value As String)
    strCierreFecha = Value
End Property

Public Property Get CierreFecha() As String
    CierreFecha = strCierreFecha
End Property





'Tipo de Registro de Compra
Public Property Let TipoRegistro(ByVal Value As String)
    strTipoRegistro = Value
End Property

Public Property Get TipoRegistro() As String
    TipoRegistro = strTipoRegistro
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

'Propiedad Usuario de Registro de Orden
Public Property Let UsuarioMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property





Rem: Atributos de Detalle

Public Property Let ITEM(ByVal Value As Double)
    dblItem = Value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property

Public Property Let AnnoOrden(ByVal Value As String)
    strAnnoOrden = Value
End Property

Public Property Get AnnoOrden() As String
    AnnoOrden = strAnnoOrden
End Property

Public Property Let NumeroOrden(ByVal Value As String)
    strNumeroOrden = Value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property

Public Property Let CodigoGastoDet(ByVal Value As String)
    strCodigoGastoDet = Value
End Property

Public Property Get CodigoGastoDet() As String
    CodigoGastoDet = strCodigoGastoDet
End Property

Public Property Let CtaContableDet(ByVal Value As String)
    strCtaContableDet = Value
End Property

Public Property Get CtaContableDet() As String
    CtaContableDet = strCtaContableDet
End Property

Public Property Let CentroCostoDet(ByVal Value As String)
    strCentroCostoDet = Value
End Property

Public Property Get CentroCostoDet() As String
    CentroCostoDet = strCentroCostoDet
End Property

Public Property Let CodigoProducto(ByVal Value As String)
    strCodigoProducto = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodigoProducto
End Property

Public Property Let ConceptoDet(ByVal Value As String)
    strConceptoDet = Value
End Property

Public Property Get ConceptoDet() As String
    ConceptoDet = strConceptoDet
End Property

Public Property Let Cantidad(ByVal Value As Double)
    dblCantidad = Value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

Public Property Let PrecioUnitario(ByVal Value As Double)
    dblPrecioUnitario = Value
End Property

Public Property Get PrecioUnitario() As Double
    PrecioUnitario = dblPrecioUnitario
End Property

Public Property Let SubTotalDet(ByVal Value As Double)
    dblSubTotalDet = Value
End Property

Public Property Get SubTotalDet() As Double
    SubTotalDet = dblSubTotalDet
End Property

Public Property Let Afecto(ByVal Value As Boolean)
    bolAfecto = Value
End Property

Public Property Get Afecto() As Boolean
    Afecto = bolAfecto
End Property

Public Property Let DebHab(ByVal Value As String)
    strDebHab = Value
End Property

Public Property Get DebHab() As String
    DebHab = strDebHab
End Property

Public Property Let SerieGuiaDet(ByVal Value As String)
    strSerieGuiaDet = Value
End Property

Public Property Get SerieGuiaDet() As String
    SerieGuiaDet = strSerieGuiaDet
End Property

Public Property Let NumeroGuiaDet(ByVal Value As String)
    strNumeroGuiaDet = Value
End Property

Public Property Get NumeroGuiaDet() As String
    NumeroGuiaDet = strNumeroGuiaDet
End Property

Public Property Let AuxiliarDet(ByVal Value As String)
    strAuxiliarDet = Value
End Property

Public Property Get AuxiliarDet() As String
    AuxiliarDet = strAuxiliarDet
End Property





Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property

Public Sub inicializarEntidades()
    strMesMovimiento = vbNullString 'F4MESMOV
    strNumeroMovimiento = vbNullString 'F4NUMMOV
    
    strCodProveedor = vbNullString 'F4CODPRV
    strNomProveedor = vbNullString 'F4NOMPRV
    strDireccionProveedor = vbNullString 'F4DIRPRV
    strTelefonoProveedor = vbNullString 'F4TELPRV
    strRucProveedor = vbNullString 'F4RUCPRV
    strTipoDocumento = vbNullString 'F4TIPDOC
    strSerieDocumento = vbNullString 'F4SERDOC
    strNumeroDocumento = vbNullString 'F4NUMDOC
    lngCodCategoria = 0 'INTCODCATEGORIA
    strFechaRegistro = vbNullString 'F4FECHAREC
    strFechaDocumento = vbNullString 'F4FECHA
    strCodMoneda = vbNullString   'F4MONEDA
    dblTipoCambio = 0 'F4TIPCAM
    strConceptoCompra = vbNullString 'F4REFERE
    dblBaseImponible = 0 'F4BASIMP
    dblMontoInafecto = 0 'F4MONINA
    dblTotalIGV = 0 'F4IGV
    dblOtroImpuesto = 0 'F4OTRIMP
    dblTotalFacturado = 0 'F4TOTAL
    strCodFormaPago = vbNullString 'F4FORPAG
    strFechaVencimiento = vbNullString 'F4FECVEN
    strCodigoGasto = vbNullString 'F4GRUPO
    strCuentaContable = vbNullString 'F4CTACONT
    strOrdenCompra = vbNullString 'F4OCOMPRA
    strCentroCosto = vbNullString 'F4OBRA
    dblMontoRetenido = 0 'F4MONTORET
    dblMontoFonavi = 0 'F4FONAVI
    bolAsientoContableProcesado = False 'F4CONTABLE
    dblCorrelativo = 0 'F4CORRELA
    dblRedSuma = 0  'F4REDSUMA
    dblRedResta = 0 'F4REDRESTA
    dblDescuento = 0 'F4DCTO
    strTipoDocAuxiliar = vbNullString 'TIPODOCAUX
    strFechaDocReferencia = vbNullString 'FECDOCREF
    strTipoDocReferencia = vbNullString 'TIPODOCREF
    strSerieDocReferencia = vbNullString 'SERDOCREF
    strNumeroDocReferencia = vbNullString 'NUMDOCREF
    strAnnoDUA = vbNullString     'ANIODUA
    
    Rem SK ADD:
    dblDetraccionPorc = 0 'PORC_DETR
    strDetraccionNum = vbNullString 'NUMDETRACCION
    strDetraccionFecha = vbNullString 'FECHADETRACCION
    
    strNroComprobante = vbNullString 'F4NROCOMPROBANTE
    bolVB = False 'F4VB
    strVBUsuario = vbNullString   'F4VBUSER
    dblPorcentajeIGV = 0 'F4PORC_IGV
    strValeIngreso = vbNullString 'F4VALESING
    
    
    bolCierre = False
    strCierreUsuario = vbNullString
    strCierreFecha = vbNullString
    
    
    strTipoRegistro = vbNullString
    
    
    strFecReg = vbNullString 'F4FECHING
    strUsuReg = vbNullString 'F4USUARIOING
    strFecMod = vbNullString 'F4FECHMOD
    strUsuMod = vbNullString 'F4USUARIOMOD
    
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesDetalle()
    dblItem = 0
    strAnnoOrden = vbNullString
    strNumeroOrden = vbNullString
    strCodigoGastoDet = vbNullString
    strCtaContableDet = vbNullString
    strCentroCostoDet = vbNullString
    strCodigoProducto = vbNullString
    strConceptoDet = vbNullString
    dblCantidad = 0
    dblPrecioUnitario = 0
    dblSubTotalDet = 0
    bolAfecto = False
    strDebHab = vbNullString
    strSerieGuiaDet = vbNullString
    strNumeroGuiaDet = vbNullString
    
    strAuxiliarDet = vbNullString
End Sub

Public Function obtenerCompra() As Boolean
    On Error GoTo errObtenerCompra
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = "SELECT * " & _
                    "FROM REGISDOC " & _
                    "WHERE F4MESMOV = '" & strMesMovimiento & "' AND " & _
                            "F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstCompra.EOF Then
        strMesMovimiento = Trim(rstCompra!F4MESMOV & "")
        strNumeroMovimiento = Trim(rstCompra!F4NUMMOV & "")
        
        strCodProveedor = Trim(rstCompra!F4CODPRV & "")
        strNomProveedor = Trim(rstCompra!F4NOMPRV & "")
        strDireccionProveedor = Trim(rstCompra!F4DIRPRV & "")
        strTelefonoProveedor = Trim(rstCompra!F4TELPRV & "")
        strRucProveedor = Trim(rstCompra!F4RUCPRV & "")
        strTipoDocumento = Trim(rstCompra!F4TIPDOC & "")
        strSerieDocumento = Trim(rstCompra!F4SERDOC & "")
        strNumeroDocumento = Trim(rstCompra!F4NUMDOC & "")
        lngCodCategoria = Val(rstCompra![INTCODCATEGORIA] & "")
        strFechaRegistro = Trim(rstCompra!F4FECHAREC & "")
        strFechaDocumento = Trim(rstCompra!F4FECHA & "")
        strCodMoneda = Trim(rstCompra!F4MONEDA & "")
        dblTipoCambio = Val(rstCompra!F4TIPCAM & "")
        strConceptoCompra = Trim(rstCompra!F4REFERE & "")
        dblBaseImponible = Val(rstCompra!F4BASIMP & "")
        dblMontoInafecto = Val(rstCompra!F4MONINA & "")
        dblTotalIGV = Val(rstCompra!F4IGV & "")
        dblOtroImpuesto = Val(rstCompra!F4OTRIMP & "")
        dblTotalFacturado = Val(rstCompra!F4TOTAL & "")
        strCodFormaPago = Trim(rstCompra!F4FORPAG & "")
        strFechaVencimiento = Trim(rstCompra!F4FECVEN & "")
        strCodigoGasto = Trim(rstCompra!F4GRUPO & "")
        strCuentaContable = Trim(rstCompra!f4ctacont & "")
        strOrdenCompra = Trim(rstCompra!F4OCOMPRA & "")
        strCentroCosto = Trim(rstCompra!F4OBRA & "")
        dblMontoRetenido = Val(rstCompra!F4MONTORET & "")
        dblMontoFonavi = Val(rstCompra!F4FONAVI & "")
        bolAsientoContableProcesado = IIf(Trim(rstCompra!F4CONTABLE & "") = "*", True, False)
        dblCorrelativo = Val(rstCompra!F4CORRELA & "")
        dblRedSuma = Val(rstCompra!F4REDSUMA & "")
        dblRedResta = Val(rstCompra!F4REDRESTA & "")
        dblDescuento = Val(rstCompra!F4DCTO & "")
        strTipoDocAuxiliar = Trim(rstCompra!TIPODOCAUX & "")
        strFechaDocReferencia = Trim(rstCompra!FECDOCREF & "")
        strTipoDocReferencia = Trim(rstCompra!TIPODOCREF & "")
        strSerieDocReferencia = Trim(rstCompra!SERDOCREF & "")
        strNumeroDocReferencia = Trim(rstCompra!NUMDOCREF & "")
        strAnnoDUA = Trim(rstCompra!ANIODUA & "")
        
        dblDetraccionPorc = Val(rstCompra!PORC_DETR & "")
        strDetraccionNum = Trim(rstCompra!NUMDETRACCION & "")
        strDetraccionFecha = Trim(rstCompra!FECHADETRACCION & "")
        
        strNroComprobante = Trim(rstCompra!F4NROCOMPROBANTE & "")
        bolVB = CBool(rstCompra!F4VB)
        strVBUsuario = Trim(rstCompra!F4VBUSER & "")
        dblPorcentajeIGV = Val(rstCompra!F4PORC_IGV & "")
        strValeIngreso = Trim(rstCompra!F4VALESING & "")
        
        strFecReg = Trim(rstCompra!F4FECHING & "") 'F4FECHING
        strUsuReg = Trim(rstCompra!F4USUARIOING & "") 'F4USUARIOING
        strFecMod = Trim(rstCompra!F4FECHMOD & "") 'F4FECHMOD
        strUsuMod = Trim(rstCompra!F4USUARIOMOD & "") 'F4USUARIOMOD
        
        obtenerCompra = True
    Else
        obtenerCompra = False
    End If
    
    rstCompra.Close
    
    Set rstCompra = Nothing
    
    Exit Function
errObtenerCompra:
    Select Case Err.Number
        Case 3704, 3709 'Perdida de Conexion o Objeto de Conexion Cerrada
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ObtenerCompra"
    End Select
    
    obtenerCompra = False
    
    Err.Clear
End Function

Public Sub obtenerOrdenDetalle()
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TEMP_DET"
    
    abrirCnTemporal
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "INSERT INTO TEMP_DET("
        strSQLCompra = strSQLCompra & "F3ITEM, F3ANOORDEN, F3ORDEN, F3GASTO, F3CTACON, F3CENCOS, "
        strSQLCompra = strSQLCompra & "F5CODPRO, F3CONCEPTO, F3CANTIDAD, F3PREUNI, AFECTO, INAFECTO, "
        strSQLCompra = strSQLCompra & "F3IMPORTE, F3AFECTO, F3DEBHAB, F4SERGUI, F4NUMGUI, F3AUXILIAR"
        strSQLCompra = strSQLCompra & ") "
        
    strSQLCompra = strSQLCompra & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLCompra = strSQLCompra & "SELECT "
        strSQLCompra = strSQLCompra & "DET.F3ITEM, "
        strSQLCompra = strSQLCompra & "DET.F3ANOORDEN, "
        strSQLCompra = strSQLCompra & "DET.F3ORDEN, "
        strSQLCompra = strSQLCompra & "DET.F3GASTO, "
        strSQLCompra = strSQLCompra & "DET.F3CTACON, "
        strSQLCompra = strSQLCompra & "DET.F3CENCOS, "
        strSQLCompra = strSQLCompra & "DET.F5CODPRO, "
        strSQLCompra = strSQLCompra & "DET.F3CONCEPTO, "
        strSQLCompra = strSQLCompra & "DET.F3CANTIDAD, "
        strSQLCompra = strSQLCompra & "DET.F3PREUNI, "
        strSQLCompra = strSQLCompra & "IIF(DET.F3AFECTO = '*', (DET.F3CANTIDAD * DET.F3PREUNI) * (CAB.PORC_IGV/100), 0) AS AFECTO, "
        strSQLCompra = strSQLCompra & "IIF(DET.F3AFECTO = '*', 0, (DET.F3CANTIDAD * DET.F3PREUNI)) AS INAFECTO, "
        strSQLCompra = strSQLCompra & "DET.F3IMPORTE, "
        strSQLCompra = strSQLCompra & "IIF(DET.F3AFECTO = '*', -1, 0) AS AFECTO, "
        strSQLCompra = strSQLCompra & "DET.F3DEBHAB, "
        strSQLCompra = strSQLCompra & "DET.F3SERGUI, "
        strSQLCompra = strSQLCompra & "DET.F3NUMGUI, "
        strSQLCompra = strSQLCompra & "DET.F3ACTIVIDAD AS AUXILIAR "
        
    strSQLCompra = strSQLCompra & "FROM "
        strSQLCompra = strSQLCompra & "REGISMOV AS DET "
        strSQLCompra = strSQLCompra & "LEFT JOIN REGISDOC AS CAB ON CAB.F4MESMOV = DET.F4MESMOV AND CAB.F4NUMMOV = DET.F4NUMMOV "
    strSQLCompra = strSQLCompra & "WHERE "
            strSQLCompra = strSQLCompra & "DET.F4MESMOV = '" & strMesMovimiento & "' AND "
            strSQLCompra = strSQLCompra & "DET.F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    cnn_dbbancos.Execute strSQLCompra
    
    strSQLCompra = vbNullString
End Sub

Public Sub obtenerConfigCompra()
    On Error GoTo errObtenerConfigCompra
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = "SELECT * " & _
                    "FROM REGISDOC " & _
                    "WHERE F4MESMOV = '" & strMesMovimiento & "' AND " & _
                            "F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCompra.EOF Then
        strMesMovimiento = Trim(rstCompra!F4MESMOV & "")
        strNumeroMovimiento = Trim(rstCompra!F4NUMMOV & "")
        
        strCodProveedor = Trim(rstCompra!F4CODPRV & "")
        strNomProveedor = Trim(rstCompra!F4NOMPRV & "")
        strDireccionProveedor = Trim(rstCompra!F4DIRPRV & "")
        strTelefonoProveedor = Trim(rstCompra!F4TELPRV & "")
        strRucProveedor = Trim(rstCompra!F4RUCPRV & "")
        strTipoDocumento = Trim(rstCompra!F4TIPDOC & "")
        strSerieDocumento = Trim(rstCompra!F4SERDOC & "")
        strNumeroDocumento = Trim(rstCompra!F4NUMDOC & "")
        lngCodCategoria = Val(rstCompra![INTCODCATEGORIA] & "")
        strFechaRegistro = Trim(rstCompra!F4FECHAREC & "")
        strFechaDocumento = Trim(rstCompra!F4FECHA & "")
        strCodMoneda = Trim(rstCompra!F4MONEDA & "")
        dblTipoCambio = Val(rstCompra!F4TIPCAM & "")
        strConceptoCompra = Trim(rstCompra!F4REFERE & "")
        dblBaseImponible = Val(rstCompra!F4BASIMP & "")
        dblMontoInafecto = Val(rstCompra!F4MONINA & "")
        dblTotalIGV = Val(rstCompra!F4IGV & "")
        dblOtroImpuesto = Val(rstCompra!F4OTRIMP & "")
        dblTotalFacturado = Val(rstCompra!F4TOTAL & "")
        strCodFormaPago = Trim(rstCompra!F4FORPAG & "")
        strFechaVencimiento = Trim(rstCompra!F4FECVEN & "")
        strCodigoGasto = Trim(rstCompra!F4GRUPO & "")
        strCuentaContable = Trim(rstCompra!f4ctacont & "")
        strOrdenCompra = Trim(rstCompra!F4OCOMPRA & "")
        strCentroCosto = Trim(rstCompra!F4OBRA & "")
        dblMontoRetenido = Val(rstCompra!F4MONTORET & "")
        dblMontoFonavi = Val(rstCompra!F4FONAVI & "")
        bolAsientoContableProcesado = IIf(Trim(rstCompra!F4CONTABLE & "") = "*", True, False)
        dblCorrelativo = Val(rstCompra!F4CORRELA & "")
        dblRedSuma = Val(rstCompra!F4REDSUMA & "")
        dblRedResta = Val(rstCompra!F4REDRESTA & "")
        dblDescuento = Val(rstCompra!F4DCTO & "")
        strTipoDocAuxiliar = Trim(rstCompra!TIPODOCAUX & "")
        strFechaDocReferencia = Trim(rstCompra!FECDOCREF & "")
        strTipoDocReferencia = Trim(rstCompra!TIPODOCREF & "")
        strSerieDocReferencia = Trim(rstCompra!SERDOCREF & "")
        strNumeroDocReferencia = Trim(rstCompra!NUMDOCREF & "")
        strAnnoDUA = Trim(rstCompra!ANIODUA & "")
        
        dblDetraccionPorc = Val(rstCompra!PORC_DETR & "")
        strDetraccionNum = Trim(rstCompra!NUMDETRACCION & "")
        strDetraccionFecha = Trim(rstCompra!FECHADETRACCION & "")
        
        strNroComprobante = Trim(rstCompra!F4NROCOMPROBANTE & "")
        bolVB = CBool(rstCompra!F4VB)
        strVBUsuario = Trim(rstCompra!F4VBUSER & "")
        dblPorcentajeIGV = Val(rstCompra!F4PORC_IGV & "")
        strValeIngreso = Trim(rstCompra!F4VALESING & "")
        
        strFecReg = Trim(rstCompra!F4FECHING & "") 'F4FECHING
        strUsuReg = Trim(rstCompra!F4USUARIOING & "") 'F4USUARIOING
        strFecMod = Trim(rstCompra!F4FECHMOD & "") 'F4FECHMOD
        strUsuMod = Trim(rstCompra!F4USUARIOMOD & "") 'F4USUARIOMOD
    End If
    
    rstCompra.Close
    
    Set rstCompra = Nothing
    
    Exit Sub
errObtenerConfigCompra:
    Select Case Err.Number
        Case 3704, 3709 'Perdida de Conexion o Objeto de Conexion Cerrada
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: ObtenerConfigCompra"
    End Select
    
    Err.Clear
End Sub

Public Sub obtenerRegistroCompraProvYdocumento()
    On Error GoTo errObtenerRegistroCompraProvYdocumento
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT "
    strSQLCompra = strSQLCompra & "F4MESMOV, "
    strSQLCompra = strSQLCompra & "F4NUMMOV "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "F4CODPRV = '" & strCodProveedor & "' AND "
    strSQLCompra = strSQLCompra & "F4TIPDOC = '" & strTipoDocumento & "' AND "
    strSQLCompra = strSQLCompra & "F4SERDOC = '" & strSerieDocumento & "' AND "
    strSQLCompra = strSQLCompra & "F4NUMDOC = '" & strNumeroDocumento & "'"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCompra.EOF Then
        strMesMovimiento = Trim(rstCompra!F4MESMOV & "")
        strNumeroMovimiento = Trim(rstCompra!F4NUMMOV & "")
    End If
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    Set rstCompra = Nothing
    
    Exit Sub
errObtenerRegistroCompraProvYdocumento:
    Select Case Err.Number
        Case 3704, 3709 'Perdida de Conexion o Objeto de Conexion Cerrada
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: ObtenerRegistroCompraProvYdocumento"
    End Select
    
    Err.Clear
End Sub

Public Function obtenerRstCompras(ByVal strAnno As String, _
                                    ByVal strCodMes As String, _
                                    ByVal bolObviarDocumentosExportadosAConcar As Boolean, _
                                    ByVal bolSoloDocumentosTransferibles As Boolean, _
                                    ByVal strCodTipoDocIncluidoEnSelect As String, _
                                    ByVal bolIncluirCodTipoDoc As Boolean, _
                                    ByVal bolSoloDocumentosConVBContable As Boolean) As ADODB.Recordset
                                    
    On Error GoTo errObtenerRstCompras
    
    Dim rsCompra As New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT "
    strSQLCompra = strSQLCompra & "REGISDOC.F4MESMOV, "
    strSQLCompra = strSQLCompra & "REGISDOC.F4NUMMOV, "
    strSQLCompra = strSQLCompra & "REGISDOC.F4CODIGV, "
    strSQLCompra = strSQLCompra & "DOC.F2TRANSFER "
    strSQLCompra = strSQLCompra & "FROM REGISDOC "
    strSQLCompra = strSQLCompra & "LEFT JOIN DOCUMENTOS AS DOC ON DOC.F2CODDOC = REGISDOC.F4TIPDOC "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "REGISDOC.F4MESMOV = '" & strAnno & strCodMes & "' AND "
        
        If strCodTipoDocIncluidoEnSelect <> vbNullString Then
            strSQLCompra = strSQLCompra & "REGISDOC.F4TIPDOC " & IIf(bolIncluirCodTipoDoc, vbNullString, "NOT") & " IN ('02') "
        End If
        
'        Select Case strOrigen
'            Case "11"
'                strSQLCompra = strSQLCompra & "REGISDOC.F4TIPDOC NOT IN ('02') "
'            Case "15"
'                strSQLCompra = strSQLCompra & "REGISDOC.F4TIPDOC IN ('02') "
'        End Select
        
        If bolObviarDocumentosExportadosAConcar Then
            strSQLCompra = strSQLCompra & "AND (ISNULL(REGISDOC.F4CONTABLE) OR REGISDOC.F4CONTABLE = '') "
        End If
        
        If bolSoloDocumentosTransferibles Then
            strSQLCompra = strSQLCompra & "AND DOC.F2TRANSFER <> 'N' "
        End If
        
        If bolSoloDocumentosConVBContable Then
            strSQLCompra = strSQLCompra & "AND REGISDOC.F4VB = TRUE "
        End If
        
    'strSQLCompra = strSQLCompra & IIf(bolObviarDocumentosExportadosAConcar, "AND (ISNULL(REGISDOC.F4CONTABLE) OR REGISDOC.F4CONTABLE = '') ", vbNullString)
    'strSQLCompra = strSQLCompra & IIf(bolSoloDocumentosTransferibles, "AND DOC.F2TRANSFER <> 'N' ", vbNullString)
    strSQLCompra = strSQLCompra & "ORDER BY "
    strSQLCompra = strSQLCompra & "REGISDOC.F4MESMOV, REGISDOC.F4NUMMOV"
    
    If rsCompra.State = 1 Then rsCompra.Close
    
    rsCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    Set obtenerRstCompras = rsCompra
    
    Exit Function
errObtenerRstCompras:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ObtenerRstCompras"
    End Select
    
    Set obtenerRstCompras = Nothing
    
    Err.Clear
End Function

Public Function obtenerRstComprasDet() As ADODB.Recordset
    On Error GoTo errObtenerRstComprasDet
    
    Dim rsCompraDet As New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = "SELECT F3CODREF AS REFERENCIA, F3GASTO AS GASTO, F3CTACON AS CUENTA, F3ACTIVIDAD AS AUXILIAR, " & _
                        "F5CODPRO AS CODIGO, F3CONCEPTO AS DESCRIPCION, F3CANTIDAD AS CANTIDAD, " & _
                        "F3IMPORTE AS VALOR, F3AFECTO AS AFECTO, F3DEBHAB AS DH, F3CENCOS AS CC " & _
                    "FROM REGISMOV " & _
                    "WHERE F4MESMOV = '" & strMesMovimiento & "' AND " & _
                            "F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    If rsCompraDet.State = 1 Then rsCompraDet.Close
    
    rsCompraDet.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    Set obtenerRstComprasDet = rsCompraDet
    
    Exit Function
errObtenerRstComprasDet:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: ObtenerRstComprasDet"
    End Select
    
    Set obtenerRstComprasDet = Nothing
    
    Err.Clear
End Function

Public Function generarNumeroMovimiento() As String
    On Error GoTo errGenerarNumeroMovimiento
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT TOP 1 "
    strSQLCompra = strSQLCompra & "F4NUMMOV AS CODIGO "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "F4MESMOV = '" & strMesMovimiento & "' AND "
    strSQLCompra = strSQLCompra & "F4NUMMOV LIKE '" & strTipoRegistro & "%' "
    strSQLCompra = strSQLCompra & "ORDER BY "
    strSQLCompra = strSQLCompra & "F4NUMMOV DESC"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstCompra.EOF Then
        generarNumeroMovimiento = Format(Val(rstCompra!Codigo & "") + 1, "0000000")
    Else
        generarNumeroMovimiento = strTipoRegistro & "00001"
    End If
    
    rstCompra.Close

    Set rstCompra = Nothing
    strSQLCompra = vbNullString

    Exit Function
errGenerarNumeroMovimiento:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: GenerarNumeroMovimiento"
    
    generarNumeroMovimiento = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = "SELECT REGISDOC.* " & _
                    "FROM REGISDOC " & _
                    "WHERE F4MESMOV = '" & strMesMovimiento & "' AND " & _
                            "F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    If rstCompra.State = 1 Then rstCompra.Close

    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstCompra.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If

    rstCompra.Close
    strSQLCompra = vbNullString

    Set rstCompra = Nothing

    Exit Function
errVerificarExistencia:
    verificarExistencia = False

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: VerificarExistencia"

    Err.Clear
End Function

Public Function guardarCompra(Optional ByVal bolRegistroExterno As Boolean) As Boolean
    On Error GoTo errGuardarCompra
    
    ''cnn_dbbancos.BeginTrans
    
    strSQLCompra = vbNullString
    
    If Not verificarExistencia Then
        If strNumeroMovimiento = vbNullString Then
            strNumeroMovimiento = generarNumeroMovimiento
        End If
        
        strSQLCompra = strSQLCompra & "INSERT INTO REGISDOC("
        strSQLCompra = strSQLCompra & "F4MESMOV, F4NUMMOV, F4CODPRV, F4NOMPRV, F4DIRPRV, F4TELPRV, "
        strSQLCompra = strSQLCompra & "F4RUCPRV, F4TIPDOC, F4SERDOC, F4NUMDOC, INTCODCATEGORIA, "
        strSQLCompra = strSQLCompra & "F4FECHAREC, F4FECHA, F4MONEDA, F4TIPCAM, F4REFERE, F4BASIMP, "
        strSQLCompra = strSQLCompra & "F4MONINA, F4IGV, F4OTRIMP, F4TOTAL, F4FORPAG, F4FECVEN, "
        strSQLCompra = strSQLCompra & "F4GRUPO, F4CTACONT, F4OCOMPRA, F4OBRA, F4MONTORET, F4FONAVI, "
        strSQLCompra = strSQLCompra & "F4CONTABLE, F4CORRELA, F4REDSUMA, F4REDRESTA, F4DCTO, TIPODOCAUX, "
        strSQLCompra = strSQLCompra & "FECDOCREF, TIPODOCREF, SERDOCREF, NUMDOCREF, ANIODUA, PORC_DETR, "
        strSQLCompra = strSQLCompra & "NUMDETRACCION, FECHADETRACCION, F4VB, F4VBUSER, "
        strSQLCompra = strSQLCompra & "F4PORC_IGV, F4VALESING, F4NROCOMPROBANTE, F4FECHING, F4USUARIOING"
        strSQLCompra = strSQLCompra & ") "
        
        strSQLCompra = strSQLCompra & "VALUES("
        strSQLCompra = strSQLCompra & "'" & strMesMovimiento & "', '" & strNumeroMovimiento & "', "
        strSQLCompra = strSQLCompra & "'" & strCodProveedor & "', '" & strNomProveedor & "', "
        strSQLCompra = strSQLCompra & "'" & strDireccionProveedor & "', '" & strTelefonoProveedor & "', "
        strSQLCompra = strSQLCompra & "'" & strRucProveedor & "', '" & strTipoDocumento & "', "
        strSQLCompra = strSQLCompra & "'" & strSerieDocumento & "', '" & strNumeroDocumento & "', "
        strSQLCompra = strSQLCompra & lngCodCategoria & ", CVDATE('" & strFechaRegistro & "'), "
        strSQLCompra = strSQLCompra & "CVDATE('" & strFechaDocumento & "'), '" & strCodMoneda & "', "
        strSQLCompra = strSQLCompra & dblTipoCambio & ", '" & strConceptoCompra & "', "
        strSQLCompra = strSQLCompra & dblBaseImponible & ", " & dblMontoInafecto & ", " & dblTotalIGV & ", "
        strSQLCompra = strSQLCompra & dblOtroImpuesto & ", " & dblTotalFacturado & ", '" & strCodFormaPago & "', "
        strSQLCompra = strSQLCompra & "CVDATE('" & strFechaVencimiento & "'), '" & strCodigoGasto & "', "
        strSQLCompra = strSQLCompra & "'" & strCuentaContable & "', '" & strOrdenCompra & "', '" & strCentroCosto & "', "
        strSQLCompra = strSQLCompra & dblMontoRetenido & ", " & dblMontoFonavi & ", "
        strSQLCompra = strSQLCompra & IIf(bolAsientoContableProcesado, "'*'", "NULL") & ", "
        strSQLCompra = strSQLCompra & dblCorrelativo & ", " & dblRedSuma & ", " & dblRedResta & ", "
        strSQLCompra = strSQLCompra & dblDescuento & ", '" & strTipoDocAuxiliar & "', "
        strSQLCompra = strSQLCompra & IIf(strFechaDocReferencia <> vbNullString, "CVDATE('" & strFechaDocReferencia & "')", "NULL") & ", '" & strTipoDocReferencia & "', "
        strSQLCompra = strSQLCompra & "'" & strSerieDocReferencia & "', '" & strNumeroDocReferencia & "', "
        strSQLCompra = strSQLCompra & "'" & strAnnoDUA & "', " & dblDetraccionPorc & ", '" & strDetraccionNum & "', "
        strSQLCompra = strSQLCompra & IIf(strDetraccionFecha <> vbNullString, "CVDATE('" & strDetraccionFecha & "')", "NULL") & ", "
        strSQLCompra = strSQLCompra & IIf(bolVB, "TRUE", "FALSE") & ", '" & strVBUsuario & "', "
        strSQLCompra = strSQLCompra & dblPorcentajeIGV & ", '" & strValeIngreso & "', "
        strSQLCompra = strSQLCompra & IIf(strNroComprobante <> vbNullString, "'" & strNroComprobante & "'", "NULL") & ", CVDATE('" & strFecReg & "'), '" & strUsuReg & "'"
        strSQLCompra = strSQLCompra & ")"
    Else
        strSQLCompra = strSQLCompra & "UPDATE "
        strSQLCompra = strSQLCompra & "REGISDOC "
        strSQLCompra = strSQLCompra & "SET "
            
            strSQLCompra = strSQLCompra & "F4CODPRV = '" & strCodProveedor & "', "
            strSQLCompra = strSQLCompra & "F4NOMPRV = '" & strNomProveedor & "', "
            strSQLCompra = strSQLCompra & "F4DIRPRV = '" & strDireccionProveedor & "', "
            strSQLCompra = strSQLCompra & "F4TELPRV = '" & strTelefonoProveedor & "', "
            strSQLCompra = strSQLCompra & "TIPODOCAUX = '" & strTipoDocAuxiliar & "', "
            strSQLCompra = strSQLCompra & "F4RUCPRV = '" & strRucProveedor & "', "
            strSQLCompra = strSQLCompra & "F4TIPDOC = '" & strTipoDocumento & "', "
            strSQLCompra = strSQLCompra & "F4SERDOC = '" & strSerieDocumento & "', "
            strSQLCompra = strSQLCompra & "F4NUMDOC = '" & strNumeroDocumento & "', "
            strSQLCompra = strSQLCompra & "F4GRUPO = '" & strCodigoGasto & "', "
            strSQLCompra = strSQLCompra & "F4CTACONT = '" & strCuentaContable & "', "
            
            If Not bolRegistroExterno Then
                strSQLCompra = strSQLCompra & "INTCODCATEGORIA = " & lngCodCategoria & ", "
                strSQLCompra = strSQLCompra & "F4FORPAG = '" & strCodFormaPago & "', "
                strSQLCompra = strSQLCompra & "F4FECVEN = CVDATE('" & strFechaVencimiento & "'), "
                
                strSQLCompra = strSQLCompra & "F4OBRA = '" & strCentroCosto & "', "
                
                strSQLCompra = strSQLCompra & "FECDOCREF = " & IIf(strFechaDocReferencia <> vbNullString, "CVDATE('" & strFechaDocReferencia & "')", "NULL") & ", "
                strSQLCompra = strSQLCompra & "TIPODOCREF = '" & strTipoDocReferencia & "', "
                strSQLCompra = strSQLCompra & "SERDOCREF = '" & strSerieDocReferencia & "', "
                strSQLCompra = strSQLCompra & "NUMDOCREF = '" & strNumeroDocReferencia & "', "
                strSQLCompra = strSQLCompra & "ANIODUA = '" & strAnnoDUA & "', "
                strSQLCompra = strSQLCompra & "PORC_DETR = " & dblDetraccionPorc & ", "
                strSQLCompra = strSQLCompra & "NUMDETRACCION = '" & strDetraccionNum & "', "
                strSQLCompra = strSQLCompra & "FECHADETRACCION = " & IIf(strDetraccionFecha <> vbNullString, "CVDATE('" & strDetraccionFecha & "')", "NULL") & ", "
                strSQLCompra = strSQLCompra & "F4VB = " & IIf(bolVB, "TRUE", "FALSE") & ", "
                strSQLCompra = strSQLCompra & "F4VBUSER = '" & strVBUsuario & "', "
                strSQLCompra = strSQLCompra & "F4CONTABLE = " & IIf(bolAsientoContableProcesado, "'*'", "NULL") & ", "
                strSQLCompra = strSQLCompra & "F4CORRELA = " & dblCorrelativo & ", "
                strSQLCompra = strSQLCompra & "F4NROCOMPROBANTE = '" & strNroComprobante & "', "
                strSQLCompra = strSQLCompra & "F4VALESING = " & strValeIngreso & ", "
                
                strSQLCompra = strSQLCompra & "F4BASIMP = " & dblBaseImponible & ", "
                strSQLCompra = strSQLCompra & "F4MONINA = " & dblMontoInafecto & ", "
                strSQLCompra = strSQLCompra & "F4IGV = " & dblTotalIGV & ", "
                strSQLCompra = strSQLCompra & "F4DCTO = " & dblDescuento & ", "
                strSQLCompra = strSQLCompra & "F4OTRIMP = " & dblOtroImpuesto & ", "
                strSQLCompra = strSQLCompra & "F4MONTORET = " & dblMontoRetenido & ", "
                strSQLCompra = strSQLCompra & "F4FONAVI = " & dblMontoFonavi & ", "
                strSQLCompra = strSQLCompra & "F4REDSUMA = " & dblRedSuma & ", "
                strSQLCompra = strSQLCompra & "F4REDRESTA = " & dblRedResta & ", "
                strSQLCompra = strSQLCompra & "F4TOTAL = " & dblTotalFacturado & ", "
            End If
            
            strSQLCompra = strSQLCompra & "F4FECHAREC = CVDATE('" & strFechaRegistro & "'), "
            strSQLCompra = strSQLCompra & "F4FECHA = CVDATE('" & strFechaDocumento & "'), "
            strSQLCompra = strSQLCompra & "F4MONEDA = '" & strCodMoneda & "', "
            strSQLCompra = strSQLCompra & "F4TIPCAM = " & dblTipoCambio & ", "
            strSQLCompra = strSQLCompra & "F4REFERE = '" & strConceptoCompra & "', "
            
            strSQLCompra = strSQLCompra & "F4OCOMPRA = '" & strOrdenCompra & "', "
            
            strSQLCompra = strSQLCompra & "F4PORC_IGV = " & dblPorcentajeIGV & ", "
            
            strSQLCompra = strSQLCompra & "F4FECHMOD = CVDATE('" & strFecMod & "'), "
            strSQLCompra = strSQLCompra & "F4USUARIOMOD = '" & strUsuMod & "' "
            
        strSQLCompra = strSQLCompra & "WHERE "
            strSQLCompra = strSQLCompra & "F4MESMOV = '" & strMesMovimiento & "' AND "
            strSQLCompra = strSQLCompra & "F4NUMMOV = '" & strNumeroMovimiento & "'"
    End If
    
    cnn_dbbancos.Execute strSQLCompra
    
    ''cnn_dbbancos.CommitTrans
    
    guardarCompra = True
    
    strSQLSelectAlter = strSQLCompra
    strSQLCompra = vbNullString
    
    Exit Function
errGuardarCompra:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vuelva a intentarlo nuevamente.", _
            vbInformation + vbOKOnly, App.ProductName & " - ClsCompra: GuardarCompra"
    'Resume
    ''cnn_dbbancos.RollbackTrans
    
    guardarCompra = False
    
    Err.Clear
End Function

Public Sub guardarCompraDetalleOneByOne()
    On Error GoTo errGuardarCompraDetalleOneByOne
    
    strSQLCompra = vbNullString
    
    strSQLCompra = strSQLCompra & "INSERT INTO REGISMOV("
    strSQLCompra = strSQLCompra & "F4MESMOV, F4NUMMOV, F3ITEM, "
    strSQLCompra = strSQLCompra & "F3ANOORDEN, F3ORDEN, F3GASTO, "
    strSQLCompra = strSQLCompra & "F3CTACON, F3CENCOS, F5CODPRO, "
    strSQLCompra = strSQLCompra & "F3CONCEPTO, F3CANTIDAD, F3PREUNI, "
    strSQLCompra = strSQLCompra & "F3IMPORTE, F3AFECTO, F3DEBHAB, "
    strSQLCompra = strSQLCompra & "F3SERGUI, F3NUMGUI, F3ACTIVIDAD) "
    
    strSQLCompra = strSQLCompra & "VALUES("
    strSQLCompra = strSQLCompra & "'" & strMesMovimiento & "', '" & strNumeroMovimiento & "', " & dblItem & ", "
    strSQLCompra = strSQLCompra & "'" & strAnnoOrden & "', '" & strNumeroOrden & "', '" & strCodigoGastoDet & "', "
    strSQLCompra = strSQLCompra & "'" & strCtaContableDet & "', '" & strCentroCostoDet & "', '" & strCodigoProducto & "', "
    strSQLCompra = strSQLCompra & "'" & strConceptoDet & "', " & dblCantidad & ", " & dblPrecioUnitario & ", "
    strSQLCompra = strSQLCompra & dblSubTotalDet & ", " & IIf(bolAfecto, "'*'", "NULL") & ", '" & strDebHab & "', "
    strSQLCompra = strSQLCompra & "'" & strSerieGuiaDet & "', '" & strNumeroGuiaDet & "', '" & strAuxiliarDet & "'"
    strSQLCompra = strSQLCompra & ")"
    
    cnn_dbbancos.Execute strSQLCompra
    
    strSQLSelectAlter = strSQLCompra
    strSQLCompra = vbNullString
    
    Exit Sub
errGuardarCompraDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: GuardarCompraDetalleOneByOne"
    
    strSQLSelectAlter = vbNullString
    
    Err.Clear
End Sub

Public Function eliminarCompra() As Boolean
    On Error GoTo errEliminarCompra
    
'    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "CORRELA", "PAG_DCTO", dblCorrelativo, "N")) > 0 Then
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarCompra = False
'
'        Exit Function
'    End If
    
    ''cnn_dbbancos.BeginTrans
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "DELETE "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "WHERE "
        strSQLCompra = strSQLCompra & "F4MESMOV = '" & strMesMovimiento & "' AND "
        strSQLCompra = strSQLCompra & "F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    cnn_dbbancos.Execute strSQLCompra
    
    ''cnn_dbbancos.CommitTrans
    
    eliminarCompra = True
    
    strSQLSelectAlter = strSQLCompra
    strSQLCompra = vbNullString
    
    Exit Function
errEliminarCompra:
    eliminarCompra = False
    
    ''cnn_dbbancos.RollbackTrans
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: EliminarCompra"
    
    Err.Clear
End Function



Public Sub documentoExportadoAConcar(ByVal strAnno As String, _
                                        ByVal strCodMes As String, _
                                        ByVal bolObviarDocumentosExportadosAConcar As Boolean, _
                                        ByVal bolSoloDocumentosTransferibles As Boolean, _
                                        ByVal strOrigen As String, _
                                        ByVal bolSoloDocumentosConVBContable As Boolean)
    
    On Error GoTo errDocumentoExportadoAConcar
    
    'cnn_dbbancos.BeginTrans
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "UPDATE "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "LEFT JOIN DOCUMENTOS AS DOC ON DOC.F2CODDOC = REGISDOC.F4TIPDOC "
    strSQLCompra = strSQLCompra & "SET "
    strSQLCompra = strSQLCompra & "REGISDOC.F4CONTABLE = '*' "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "REGISDOC.F4MESMOV = '" & strAnno & strCodMes & "' "
    
        Select Case strOrigen
            Case "11"
                strSQLCompra = strSQLCompra & "AND REGISDOC.F4TIPDOC NOT IN ('02') "
            Case "15"
                strSQLCompra = strSQLCompra & "AND REGISDOC.F4TIPDOC IN ('02') "
        End Select
        
        If bolObviarDocumentosExportadosAConcar Then
            strSQLCompra = strSQLCompra & "AND (ISNULL(REGISDOC.F4CONTABLE) OR REGISDOC.F4CONTABLE = '') "
        End If
        
        If bolSoloDocumentosTransferibles Then
            strSQLCompra = strSQLCompra & "AND DOC.F2TRANSFER <> 'N' "
        End If
        
        If bolSoloDocumentosConVBContable Then
            strSQLCompra = strSQLCompra & "AND REGISDOC.F4VB = TRUE "
        End If
        
        'IIf(bolObviarDocumentosExportadosAConcar, "AND (ISNULL(REGISDOC.F4CONTABLE) OR REGISDOC.F4CONTABLE = '') ", vbNullString)
        'IIf(bolSoloDocumentosTransferibles, "AND DOC.F2TRANSFER <> 'N' ", vbNullString)
    
    cnn_dbbancos.Execute strSQLCompra

    'cnn_dbbancos.CommitTrans

    strSQLCompra = vbNullString

    Exit Sub
errDocumentoExportadoAConcar:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: DocumentoExportadoAConcar"
    End Select
    
    'cnn_dbbancos.RollbackTrans

    Err.Clear
End Sub

Public Sub marcarDocumentoComoTransferido()
    On Error GoTo errMarcarDocumentoComoTransferido
    
    'cnn_dbbancos.BeginTrans
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "UPDATE "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "SET "
    strSQLCompra = strSQLCompra & "REGISDOC.F4CONTABLE = '" & IIf(bolAsientoContableProcesado, "*", vbNullString) & "', "
    strSQLCompra = strSQLCompra & "REGISDOC.F4NROCOMPROBANTE = " & IIf(strNroComprobante <> vbNullString, "'" & strNroComprobante & "'", "NULL") & " "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "REGISDOC.F4MESMOV = '" & strMesMovimiento & "' AND "
    strSQLCompra = strSQLCompra & "REGISDOC.F4NUMMOV = '" & strNumeroMovimiento & "'"
    
    cnn_dbbancos.Execute strSQLCompra

    'cnn_dbbancos.CommitTrans
    
    strSQLSelectAlter = strSQLCompra
    
    strSQLCompra = vbNullString

    Exit Sub
errMarcarDocumentoComoTransferido:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: MarcarDocumentoComoTransferido"
    End Select
    
    'cnn_dbbancos.RollbackTrans

    Err.Clear
End Sub

Public Function verificarCierreCompra() As Boolean
    On Error GoTo errVerificarCierreCompra
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT "
    strSQLCompra = strSQLCompra & "F4MESMOV, "
    strSQLCompra = strSQLCompra & "F4CIERRE "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "F4MESMOV = '" & strMesMovimiento & "' "
        
        If strNumeroMovimiento <> vbNullString Then
            strSQLCompra = strSQLCompra & "AND F4NUMMOV = '" & strNumeroMovimiento & "' "
        End If
    
    strSQLCompra = strSQLCompra & "ORDER BY "
    strSQLCompra = strSQLCompra & "F4NUMMOV"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCompra.EOF Then
        rstCompra.MoveFirst
        
        verificarCierreCompra = True
        
        Do While Not rstCompra.EOF
            If Not CBool(rstCompra!F4CIERRE) Then
                verificarCierreCompra = False
                
                Exit Do
            End If
            
            rstCompra.MoveNext
        Loop
    Else
        verificarCierreCompra = False
    End If
    
    rstCompra.Close
    
    strSQLCompra = vbNullString
    
    Set rstCompra = Nothing
    
    Exit Function
errVerificarCierreCompra:
    verificarCierreCompra = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: VerificarCierreCompra"
    
    Err.Clear
End Function

Public Function cerrarCompra() As Boolean
    On Error GoTo errCerrarCompra
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "UPDATE "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "SET "
    strSQLCompra = strSQLCompra & "F4CIERRE = " & IIf(bolCierre, "1", "0") & ", "
    strSQLCompra = strSQLCompra & "F4CIERREUSER = '" & strCierreUsuario & "', "
    strSQLCompra = strSQLCompra & "F4CIERREFECHA = CVDATE('" & strCierreFecha & "') "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "F4MESMOV = '" & strMesMovimiento & "' "
        
        If strNumeroMovimiento <> vbNullString Then
            strSQLCompra = strSQLCompra & "AND F4NUMMOV = '" & strNumeroMovimiento & "' "
        End If
    
    cnn_dbbancos.Execute strSQLCompra
    
    cerrarCompra = True
    
    strSQLSelectAlter = strSQLCompra
    strSQLCompra = vbNullString
    
    Exit Function
errCerrarCompra:
    cerrarCompra = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: CerrarCompra"
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub listarAnno(ByVal comboList As Object)
    
    On Error GoTo errListarAnno
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT "
    strSQLCompra = strSQLCompra & "LEFT(F4MESMOV, LEN(F4MESMOV) -2) AS ANNO "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "GROUP BY "
    strSQLCompra = strSQLCompra & "LEFT(F4MESMOV, LEN(F4MESMOV) -2) "
    strSQLCompra = strSQLCompra & "ORDER BY "
    strSQLCompra = strSQLCompra & "LEFT(F4MESMOV, LEN(F4MESMOV) -2)"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstCompra.EOF Then
        
        rstCompra.MoveFirst
        
        Do While Not rstCompra.EOF
            comboList.AddItem Trim(rstCompra!Anno & "")
            
            rstCompra.MoveNext
        Loop
    Else
        comboList.AddItem Year(Date)
    End If
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    strSQLCompra = vbNullString
    
    Set rstCompra = Nothing
    
    Exit Sub
errListarAnno:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ListarAnno"
    
    Err.Clear
End Sub

Public Sub listarMes(ByVal comboList As Object, _
                        ByVal strAnno As String)
    
    On Error GoTo errListarMes
    
    Set rstCompra = New ADODB.Recordset
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "SELECT "
    strSQLCompra = strSQLCompra & "RIGHT(F4MESMOV, 2) AS MES "
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "REGISDOC "
    strSQLCompra = strSQLCompra & "WHERE "
    strSQLCompra = strSQLCompra & "LEFT(F4MESMOV, LEN(F4MESMOV) -2) = '" & strAnno & "' "
    strSQLCompra = strSQLCompra & "GROUP BY "
    strSQLCompra = strSQLCompra & "RIGHT(F4MESMOV, 2) "
    strSQLCompra = strSQLCompra & "ORDER BY "
    strSQLCompra = strSQLCompra & "RIGHT(F4MESMOV, 2)"
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    rstCompra.Open strSQLCompra, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstCompra.EOF Then
        
        rstCompra.MoveFirst
        
        Do While Not rstCompra.EOF
            comboList.AddItem UCase(Format("01/" & Trim(rstCompra!mes & "") & "/" & strAnno, "MMMM")) & Space(50) & Trim(rstCompra!mes & "")
            
            rstCompra.MoveNext
        Loop
    Else
        comboList.AddItem UCase(Format("01/" & Format(Month(Date), "00") & "/" & strAnno, "MMMM")) & Space(50) & Format(Month(Date), "00")
    End If
    
    If rstCompra.State = 1 Then rstCompra.Close
    
    strSQLCompra = vbNullString
    
    Set rstCompra = Nothing
    
    Exit Sub
errListarMes:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ListarMes"
    
    Err.Clear
End Sub

'Vista de Estado de Ventas Facturadas en Grilla (QuamtumGrid)
Public Sub vistaConsolidadoCompras(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaConsolidadoCompras
    
    Dim rsConsolidado As New ADODB.Recordset
    Dim lngCol As Long
    
    strSQLCompra = vbNullString
    strSQLCompra = strSQLCompra & "TRANSFORM "
    
    strSQLCompra = strSQLCompra & "SUM(VAL(FORMAT(DET.F3IMPORTE * IIF(CAB.F4MONEDA = 'S', 1, CAB.F4TIPCAM), '#0.00')) * IIF(CAB.F4TIPDOC = '07', -1, 1)) "
    
    strSQLCompra = strSQLCompra & "SELECT "
    
    strSQLCompra = strSQLCompra & "IIF(TRIM(GRUP.NOMBRE & '') = '', '< Sin Grupo >', GRUP.NOMBRE) AS GRUPO, "
    strSQLCompra = strSQLCompra & "DET.F3GASTO, "
    strSQLCompra = strSQLCompra & "DET.F3CTACON, "
    strSQLCompra = strSQLCompra & "GAS.NOMBRE AS GASTO "
    
    strSQLCompra = strSQLCompra & "FROM "
    strSQLCompra = strSQLCompra & "((REGISMOV AS DET "
    strSQLCompra = strSQLCompra & "LEFT JOIN REGISDOC AS CAB ON (CAB.F4NUMMOV = DET.F4NUMMOV) AND (CAB.F4MESMOV = DET.F4MESMOV)) "
    strSQLCompra = strSQLCompra & "LEFT JOIN BF9GIN AS GAS ON GAS.CODIGO = DET.F3GASTO) "
    strSQLCompra = strSQLCompra & "LEFT JOIN GRUPOS_FLUJO AS GRUP ON GRUP.CODIGO = GAS.GRUPOFLUJO "
    strSQLCompra = strSQLCompra & "GROUP BY "
    
    strSQLCompra = strSQLCompra & "IIF(TRIM(GRUP.NOMBRE & '') = '', '< Sin Grupo >', GRUP.NOMBRE), "
    strSQLCompra = strSQLCompra & "DET.F3GASTO, "
    strSQLCompra = strSQLCompra & "DET.F3CTACON, "
    'strSQLCompra = strSQLCompra & "GRUP.NOMBRE, "
    strSQLCompra = strSQLCompra & "GAS.NOMBRE "
    
    strSQLCompra = strSQLCompra & "ORDER BY "
    
    strSQLCompra = strSQLCompra & "DET.F3CTACON, "
    strSQLCompra = strSQLCompra & "DET.F3GASTO "
    
    strSQLCompra = strSQLCompra & "PIVOT "
    strSQLCompra = strSQLCompra & "CAB.F4MESMOV"
    
    If rsConsolidado.State = 1 Then rsConsolidado.Close
    
    rsConsolidado.Open strSQLCompra, cnn_dbbancos, adOpenDynamic, adLockReadOnly
    
    If Not rsConsolidado.EOF Then
        rsConsolidado.MoveFirst
        
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        Dim gGroupSum As dxGridSummaryGroup
        Dim gGroupItem As dxGridSummaryItem
        
        With grilla
            'Columna Codigo de Gasto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                '.Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F3GASTO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColF3GASTO"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColF3GASTO"
                .SummaryField = "F3GASTO"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Cuenta Contable
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Cta. Contable"
                '.Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F3CTACON"
                .HeaderAlignment = taCenter
                .ObjectName = "ColF3CTACON"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = True
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColF3CTACON"
                .SummaryField = "F3CTACON"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Descripción de Grupo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Grupo"
                '.Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "GRUPO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColGRUPO"
                '.SummaryType = cstCount
                .SummaryFooterType = cstCount
                '.SummaryFooterFormat = " "
                .Width = 220
                .Visible = True
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColGRUPO"
                .SummaryField = "GRUPO"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Descripción de Gasto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción del Gasto"
                '.Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "GASTO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColGASTO"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 220
                .Visible = True
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColGASTO"
                .SummaryField = "GASTO"
                .SummaryType = cstCount
                '.SummaryFormat = " "
            End With
            
            
            For lngCol = 0 To rsConsolidado.Fields.Count - 1
                If IsNumeric(rsConsolidado.Fields(lngCol).Name) Then
                    'Columna Periodo (Anno & Mes)
                    Set gColumn = .Columns.Add(gedSpinEdit)
                    
                    With gColumn
                        .Alignment = taRightJustify
                        .Caption = Trim(rsConsolidado.Fields(lngCol).Name & "")
                            If lngCol Mod 2 = 0 Then
                                .Color = RGB(176, 196, 222)
                            Else
                                .Color = RGB(220, 220, 220)
                            End If
                        .BandIndex = 1
                        .DecimalPlaces = 2
                        .DisableEditor = True
                        .FieldName = Trim(rsConsolidado.Fields(lngCol).Name & "")
                        .HeaderAlignment = taCenter
                        .ObjectName = "Col" & Trim(rsConsolidado.Fields(lngCol).Name & "")
                        .SummaryFooterType = cstSum
                        '.SummaryFooterFormat = " "
                        .Width = 80
                        .Visible = True
                    End With
                    
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "Col" & Trim(rsConsolidado.Fields(lngCol).Name & "")
                        .SummaryField = Trim(rsConsolidado.Fields(lngCol).Name & "")
                        .SummaryType = cstSum
                    End With
                End If
            Next lngCol
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexDbBancos
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCompra
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F3CTACON"
            
            .Columns.ColumnByFieldName("F3CTACON").Sorted = csUp
            .Columns.ColumnByFieldName("GRUPO").GroupIndex = 0
            
            .m.FullExpand
        End With
    End If
    
    strSQLCompra = vbNullString
    If rsConsolidado.State = 1 Then rsConsolidado.Close
    
    Exit Sub
errVistaConsolidadoCompras:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: VistaConsolidadoCompras"
    End Select
    
    Err.Clear
End Sub

'Obtener Libro Electronico - Registro de Compras
Public Sub obtenerEBookCompras(ByVal bolSoloMesDeCorte As Boolean, _
                                ByVal strAnno As String, _
                                ByVal intCorteMes As Integer, _
                                ByVal strRutaEBook As String, _
                                ByVal barraEstado As Object, _
                                ByVal etiquetaEstado As Object)
    
    On Error GoTo errObtenerEBookCompras
    
    Dim rsEBook As New ADODB.Recordset
    
    Dim strFile As String
    Dim StrLine As String
    Dim intNumSlot As Integer
    Dim intMes As Integer
    
    If Trim(strRutaEBook) = vbNullString Then
        MsgBox "Imposible generar Libros, noo se especifico ninguna ruta de descarga.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
        
        Exit Sub
    End If
    
    For intMes = IIf(bolSoloMesDeCorte, intCorteMes, 1) To intCorteMes
        '*************************************************************************************************************************************************
        'Open strFile For Append As #intNumSlot
        strFile = Trim(strRutaEBook & "\LE" & wrucempresa & strAnno & Format(intMes, "00") & "00080100001111.txt")
        intNumSlot = FreeFile
        
        If Len(Dir$(strFile)) Then Kill strFile
        
        Close #intNumSlot
        
        Open strFile For Append As #intNumSlot
        
        Close #intNumSlot
        
        '*************************************************************************************************************************************************
        
        strSQLCompra = vbNullString
        strSQLCompra = strSQLCompra & "SELECT "
        strSQLCompra = strSQLCompra & "C.F4MESMOV & '00' AS PERIODO, "
        strSQLCompra = strSQLCompra & "C.F4NUMMOV AS CUO, "
        strSQLCompra = strSQLCompra & "TRIM('M' & C.F4NUMMOV) AS CUO2, "
        strSQLCompra = strSQLCompra & "C.F4FECHA, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(D.CODEXT2 & '') = '14', C.F4FECVEN, NULL) AS FECVEN, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(D.CODEXT2 & '') = '', '00', D.CODEXT2) AS F4TIPDOC, "
        strSQLCompra = strSQLCompra & "C.F4SERDOC, "
        strSQLCompra = strSQLCompra & "C.ANIODUA, "
        strSQLCompra = strSQLCompra & "C.F4NUMDOC, "
        strSQLCompra = strSQLCompra & "0 AS IMPORTESINCREDFIS, "
        strSQLCompra = strSQLCompra & "IIF(ISNUMERIC(C.F4RUCPRV), IIF(LEN(C.F4RUCPRV) = 8, '1', IIF(LEN(C.F4RUCPRV) = 11, '6', 'A')), IIF(LEN(C.F4RUCPRV) = 12, '4', '0')) AS TIPODOCAUX, " '"6 AS TIPODOCAUX, "
        strSQLCompra = strSQLCompra & "C.F4RUCPRV, "
        strSQLCompra = strSQLCompra & "P.F2NOMPROV AS F4NOMPRV, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='001', (IIF(D.CODEXT2 <> '02', C.F4BASIMP, 0) * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS F4BASIMP1, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='001', (IIF(D.CODEXT2 <> '02', C.F4IGV, 0) * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS IGVDG, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='002', (C.F4BASIMP * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS F4BASIMP2, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='002', (C.F4IGV * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS IGVDM, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='003', (C.F4BASIMP * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS F4BASIMP3, "
        strSQLCompra = strSQLCompra & "IIF(C.F4CODIGV ='003', (C.F4IGV * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB = 'D', -1, 1)) AS IGVDE, "
        
        strSQLCompra = strSQLCompra & "((IIF(D.CODEXT2 <> '02', 0, C.F4BASIMP) + C.F4MONINA + C.F4REDSUMA - C.F4REDRESTA) * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB='D', -1, 1) AS F4MONINA, "
        
        strSQLCompra = strSQLCompra & "0 AS IMPUESTOSELECTIVO, "
        strSQLCompra = strSQLCompra & "(C.F4OTRIMP * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) * IIF(PD.DEB_HAB='D', -1, 1) AS F4OTRIMP, "
        'strSQLCompra = strSQLCompra & "((IIF(D.CODEXT2 <> '02', C.F4TOTAL - C.F4OTRIMP, C.F4BASIMP)) * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) *  IIF(PD.DEB_HAB = 'D', -1, 1) AS TOTALTC, "
        strSQLCompra = strSQLCompra & "((IIF(D.CODEXT2 <> '02', C.F4TOTAL, C.F4BASIMP + C.F4MONINA) - C.F4OTRIMP) * IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM)) *  IIF(PD.DEB_HAB = 'D', -1, 1) AS TOTALTC, "
        strSQLCompra = strSQLCompra & "IIF(C.F4MONEDA = 'S', 1, C.F4TIPCAM) AS TIPOCAMBIO, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(C.NUMDOCREF & '') <> '', C.FECDOCREF, NULL) AS FECDOCREF, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(C.NUMDOCREF & '') = '', '', IIF(TRIM(DREF.CODEXT2 & '') = '', '00', DREF.CODEXT2)) AS TIPODOCREF, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(C.NUMDOCREF & '') = '', '', C.SERDOCREF) AS SERDOCREF, "
        strSQLCompra = strSQLCompra & "'' AS CODIGODEPENDENCIADUA, "
        strSQLCompra = strSQLCompra & "IIF(TRIM(C.NUMDOCREF & '') = '', '', C.NUMDOCREF) AS NUMDOCREF, "
        strSQLCompra = strSQLCompra & "'' AS NUMCOMPNODOMICI, "
        
        strSQLCompra = strSQLCompra & "IIF(TRIM(C.NUMDETRACCION & '') = '', NULL, C.FECHADETRACCION) AS FECHADETRACCION, "
        strSQLCompra = strSQLCompra & "C.NUMDETRACCION, "
        strSQLCompra = strSQLCompra & "IIF(C.F4MONTORET > 0, 1, 0) AS MARCA, "
        strSQLCompra = strSQLCompra & "1 AS ESTADO "
        strSQLCompra = strSQLCompra & "FROM "
        strSQLCompra = strSQLCompra & "((((REGISDOC AS C "
        strSQLCompra = strSQLCompra & "LEFT JOIN DOCUMENTOS AS D ON D.F2CODDOC = C.F4TIPDOC) "
        strSQLCompra = strSQLCompra & "LEFT JOIN DOCUMENTOS AS DREF ON DREF.F2CODDOC = C.TIPODOCREF) "
        strSQLCompra = strSQLCompra & "LEFT JOIN CENTROS ON C.F4OBRA = CENTROS.F3COSTO) "
        strSQLCompra = strSQLCompra & "LEFT JOIN PAG_DCTO AS PD ON PD.CORRELA = C.F4CORRELA) "
        strSQLCompra = strSQLCompra & "LEFT JOIN EF2PROVEEDORES AS P ON C.F4CODPRV = P.F2CODPROV "
        
        strSQLCompra = strSQLCompra & "WHERE "
        strSQLCompra = strSQLCompra & "C.F4MESMOV = '" & strAnno & Format(intMes, "00") & "' AND "
        strSQLCompra = strSQLCompra & "D.F2OFICIAL = TRUE "
        
        strSQLCompra = strSQLCompra & "ORDER BY "
        strSQLCompra = strSQLCompra & "C.F4MESMOV, "
        strSQLCompra = strSQLCompra & "C.F4NUMMOV"
        
        If rsEBook.State = 1 Then rsEBook.Close
        
        rsEBook.Open strSQLCompra, cnn_dbbancos, adOpenDynamic, adLockReadOnly
        
        If Not rsEBook.EOF Then
            rsEBook.MoveFirst
            
            
            barraEstado.Max = ModUtilitario.devuelveCantRegistros(rsEBook)
            barraEstado.Value = 0
            
            Do While Not rsEBook.EOF
                StrLine = vbNullString
                StrLine = StrLine & IIf(Trim(rsEBook!Periodo & "") <> vbNullString, Trim(Trim(rsEBook!Periodo & "")), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!CUO & "") <> vbNullString, Trim(rsEBook!CUO & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!CUO2 & "") <> vbNullString, Trim(rsEBook!CUO2 & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(IsDate(rsEBook!F4FECHA & ""), Format(Trim(rsEBook!F4FECHA & ""), "Short Date"), "01/01/0001") & "|"
                StrLine = StrLine & IIf(IsDate(rsEBook!FECVEN & ""), Format(Trim(rsEBook!FECVEN & ""), "Short Date"), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!F4TIPDOC & "") <> vbNullString, Trim(rsEBook!F4TIPDOC & ""), vbNullString) & "|"
                
                If Trim(rsEBook!F4SERDOC & "") <> vbNullString Then
                    If IsNumeric(Trim(rsEBook!F4SERDOC & "")) Then
                        StrLine = StrLine & Format(Trim(rsEBook!F4SERDOC & ""), "0000") & "|"
                    Else
                        StrLine = StrLine & UCase(left(Trim(rsEBook!F4SERDOC & ""), 1)) & Format(Mid(Trim(rsEBook!F4SERDOC & ""), 2), "000") & "|"
                    End If
                Else
                    StrLine = StrLine & "|"
                End If
                
                'StrLine = StrLine & IIf(Trim(rsEBook!F4SERDOC & "") <> vbNullString, Trim(rsEBook!F4SERDOC & ""), vbNullString) & "|"
                
                StrLine = StrLine & IIf(Trim(rsEBook!ANIODUA & "") <> vbNullString, Trim(rsEBook!ANIODUA & ""), vbNullString) & "|"
                
                Select Case Val(rsEBook!F4TIPDOC & "")
                    Case 1 To 4, 6 To 8, 23, 25, 34 To 35, 48, 89
                        StrLine = StrLine & IIf(Trim(rsEBook!F4NUMDOC & "") <> vbNullString, Format(right(Trim(rsEBook!F4NUMDOC & ""), 7), "0000000"), vbNullString) & "|"
                    Case 36
                        StrLine = StrLine & IIf(Trim(rsEBook!F4NUMDOC & "") <> vbNullString, Format(right(Trim(rsEBook!F4NUMDOC & ""), 8), "00000000"), vbNullString) & "|"
                    Case 50 To 53
                        StrLine = StrLine & IIf(Trim(rsEBook!F4NUMDOC & "") <> vbNullString, Format(right(Trim(rsEBook!F4NUMDOC & ""), 6), "000000"), vbNullString) & "|"
                    Case Else
                        StrLine = StrLine & IIf(Trim(rsEBook!F4NUMDOC & "") <> vbNullString, Format(Trim(rsEBook!F4NUMDOC & ""), "0000000"), vbNullString) & "|"
                End Select
                
                'StrLine = StrLine & IIf(Trim(rsEBook!F4NUMDOC & "") <> vbNullString, Format(Trim(rsEBook!F4NUMDOC & ""), "0000000"), vbNullString) & "|"
                
                StrLine = StrLine & IIf(Val(rsEBook!IMPORTESINCREDFIS & "") > 0, Val(rsEBook!IMPORTESINCREDFIS & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!TIPODOCAUX & "") <> vbNullString, Trim(rsEBook!TIPODOCAUX & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!F4RUCPRV & "") <> vbNullString, Trim(rsEBook!F4RUCPRV & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!F4NOMPRV & "") <> vbNullString, left(ModUtilitario.limpiarCaracteresEnCadena(Trim(rsEBook!F4NOMPRV & "")), 100), vbNullString) & "|"
                StrLine = StrLine & Format(Val(rsEBook!F4BASIMP1 & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!IGVDG & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!F4BASIMP2 & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!IGVDM & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!F4BASIMP3 & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!IGVDE & ""), "#0.00") & "|"
                
                StrLine = StrLine & Format(Val(rsEBook!F4MONINA & ""), "#0.00") & "|"
                
                StrLine = StrLine & Format(Val(rsEBook!IMPUESTOSELECTIVO & ""), "#0.00") & "|"
                StrLine = StrLine & Format("0", "#0.00") & "|" 'StrLine = StrLine & Format(Val(rsEBook!F4OTRIMP & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!TOTALTC & ""), "#0.00") & "|"
                StrLine = StrLine & Format(Val(rsEBook!TipoCambio & ""), "#0.000") & "|"
                
                StrLine = StrLine & IIf(IsDate(rsEBook!FECDOCREF & ""), Format(Trim(rsEBook!FECDOCREF & ""), "Short Date"), IIf(Trim(rsEBook!NUMDOCREF & "") <> vbNullString, "01/01/0001", vbNullString)) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!TIPODOCREF & "") <> vbNullString, Trim(rsEBook!TIPODOCREF & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!SERDOCREF & "") <> vbNullString, Trim(rsEBook!SERDOCREF & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!CODIGODEPENDENCIADUA & "") <> vbNullString, Trim(rsEBook!CODIGODEPENDENCIADUA & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!NUMDOCREF & "") <> vbNullString, Trim(rsEBook!NUMDOCREF & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!NUMCOMPNODOMICI & "") <> vbNullString, Trim(rsEBook!NUMCOMPNODOMICI & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(IsDate(rsEBook!FECHADETRACCION & ""), Format(Trim(rsEBook!FECHADETRACCION & ""), "Short Date"), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!NUMDETRACCION & "") <> vbNullString, Trim(rsEBook!NUMDETRACCION & ""), vbNullString) & "|"
                StrLine = StrLine & IIf(Trim(rsEBook!Marca & "") <> vbNullString, Trim(rsEBook!Marca & ""), vbNullString) & "|"
                    
'                    Select Case Val(Format(Trim(rsEBook!F4FECHA & ""), "YYYYMM00"))
'                        Case Is = Val(rsEBook!Periodo & "")
'                            Select Case Trim(rsEBook!F4TIPDOC & "")
'                                Case "02"
'                                    StrLine = StrLine & IIf(Trim(rsEBook!Estado & "") <> vbNullString, Trim(rsEBook!Estado & "") & "|", vbNullString)
'                                Case Else
'                                    StrLine = StrLine & "1|"
'                            End Select
'                        Case Is < Val(rsEBook!Periodo & "")
'                            Select Case Trim(rsEBook!F4TIPDOC & "")
'                                Case "02"
'                                    StrLine = StrLine & IIf(Trim(rsEBook!Estado & "") <> vbNullString, Trim(rsEBook!Estado & "") & "|", vbNullString)
'                                Case Else
'                                    Select Case Val(Format(DateAdd("m", 12, CDate(rsEBook!F4FECHA)), "YYYYMM00"))
'                                        Case Is >= Val(rsEBook!Periodo & "")
'                                            StrLine = StrLine & "6|"
'                                        Case Is < Val(rsEBook!Periodo & "")
'                                            StrLine = StrLine & "7|"
'                                    End Select
'                            End Select
'                    End Select


                    
                    Select Case Val(Format(Trim(rsEBook!F4FECHA & ""), "YYYYMM00"))
                        Case Is = Val(rsEBook!Periodo & "")
                            Select Case Trim(rsEBook!F4TIPDOC & "")
                                Case 2 To 3, 10, 15 To 19, 21 To 22, 44 To 46, 49 'Estado = 0 ó 9
                                    StrLine = StrLine & "0|"
                                Case Else
                                    StrLine = StrLine & "1|"
                            End Select
                        Case Is < Val(rsEBook!Periodo & "")
                            Select Case Trim(rsEBook!F4TIPDOC & "")
                                Case 2 To 3, 10, 15 To 19, 21 To 22, 44 To 46, 49 'Estado = 0 ó 9
                                    StrLine = StrLine & "0|"
                                Case Else
                                    Select Case Val(Format(DateAdd("m", 12, CDate(rsEBook!F4FECHA)), "YYYYMM00"))
                                        Case Is >= Val(rsEBook!Periodo & "")
                                            StrLine = StrLine & "6|"
                                        Case Is < Val(rsEBook!Periodo & "")
                                            StrLine = StrLine & "7|"
                                    End Select
                            End Select
                    End Select
                    
                    


                    
'                    Select Case Val(rsEBook!F4TIPDOC & "")
'                        Case 2 To 3, 10, 15 To 19, 21 To 22, 44 To 46, 49
'                            If Val(Year(CDate(rsEBook!F4FECHA & "")) & Month(CDate(rsEBook!F4FECHA & ""))) < Val(Trim(rsEBook!Periodo & "")) Then
'                                StrLine = StrLine & "9|"
'                            Else
'                                StrLine = StrLine & "0|"
'                            End If
'                        Case Else
'                            If Val(Year(CDate(rsEBook!F4FECHA & "")) & Month(CDate(rsEBook!F4FECHA & ""))) < Val(Trim(rsEBook!Periodo & "")) Then
'                                StrLine = StrLine & "6|"
'                            Else
'                                StrLine = StrLine & "1|"
'                            End If
'                    End Select
                    
                    'StrLine = StrLine & IIf(Trim(rsEBook!Estado & "") <> vbNullString, Trim(rsEBook!Estado & ""), vbNullString) & "|"
                
                Open strFile For Append As #intNumSlot
                
                Print #intNumSlot, StrLine
                
                Close #intNumSlot
                
                DoEvents
                
                barraEstado.Value = barraEstado.Value + 1
                etiquetaEstado.Caption = "Procesando " & ModUtilitario.devuelveNombreMes(Format(intMes, "00")) & "... " & FormatPercent(barraEstado.Value / barraEstado.Max, 0)
                
                rsEBook.MoveNext
            Loop
        Else
            Kill strFile
        End If
        
        strSQLCompra = vbNullString
        
        strFile = vbNullString
        StrLine = vbNullString
        intNumSlot = 0
        
        If rsEBook.State = 1 Then rsEBook.Close
    Next intMes
    
    Exit Sub
errObtenerEBookCompras:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCompra: ObtenerEBookCompras"
    End Select
    
    Err.Clear
End Sub


