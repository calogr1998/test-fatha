VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodAlmacen               As String 'F2CODALM
Private strNumeroVale               As String 'F4NUMVAL
Private strTipoVale                 As String 'F4TIPOVALE

Private strTipoPersona              As String 'F1TIPPRV
Private strCodProveedor             As String 'F2CODPROV
Private strNumeroDocumento          As String 'F4NUMDOC
Private StrFecha                    As String 'F4FECVAL
Private strCodOrigen                As String 'F1CODORI
Private strCodDocumento             As String 'F1CODDOC
Private dblTCambio                  As Double 'F4TIPCAM
Private strCodMoneda                As String 'F4MONEDA
Private strTipoProveedor            As String 'F1TIPPRV
Private strCodPar                   As String 'F2CODPAR
Private strNumeroFactura            As String 'F4NUMFAC
Private dblNumeroOrden              As Double 'F4NUMORD
Private strNumeroImportacion        As String 'F4NUMIMP
Private strPartida                  As String 'F4PARTIDA
Private strOrdenTrabajo             As String 'F4ORDTRA
Private strFechaUltima              As String 'F4FECULT
Private strCodUsuario               As String 'F2CODUSE
Private strReferencia               As String 'F4REFERE
Private strCentroCosto              As String 'F4CENTRO
Private strSerieGuia                As String 'F4SERGUIA
Private strNumeroGuia               As String 'F4NUMGUIA
Private strCodTipoComprobante       As String 'F4TIPDOC
Private strFecReg                   As String 'F4FECGRA
Private strUsuReg                   As String 'F4USEGRA
Private strFecMod                   As String 'F4FECMOD
Private strUsuMod                   As String 'F4USEMOD
Private strSerieDocumento           As String 'F4SERDOC
Private strObservaciones            As String 'F4OBSERVA
Private strUUPP                     As String 'F4UUPP
Private strNumeroValeEnvio          As String 'F4NUMVALENVIO
Private strNumeroOrdenCompra        As String 'NUMORDEN
Private strCodAlmacenDestino        As String 'CODALMDES
Private strCodSolicitante           As String 'F2CODUSER
Private strNumeroValeExterno        As String 'NUMENSAM
Private strSeccion                  As String 'SECCION
Private dblLote                     As Double 'LOTE
Private strRegistroCompra           As String 'F4REGCOM
Private strValeIngreso              As String 'VALING
Private bolExportarVale             As String 'EXPORTARVALE
Private bolVB1                      As String 'VB1
Private strVB1Usuario               As String 'VB1USER
Private strVB1Fecha                 As String 'VB1FECHA

'Atributos de Detalle de Vale
Private strCodProducto As String 'F5CODPRO
Private strCodProductoOriginal As String 'F5CODPROORIGINAL
Private dblCantidad As Double 'F3CANPRO
Private dblValorVta As Double 'F3VALVTA
Private dblIgv As Double 'F3IGV
Private dblTotal As Double 'F3TOTITE

Private dblValorVtaDol As Double 'F3VALDOL
Private dblIgvDol As Double 'F3IGVDOL
Private dblTotalDol As Double 'F3TOTDOL

Private strGrupo As String 'F3GRUPO
Private dblItem As Double 'F3ITEM
Private strRequerimiento 'COD_SOLICITUD
Private dblCantidadMaxima 'F3SALPEP
Private dblPorcentajeDscto As Double  'F3PORCENTAJEDSCTO
Private dblMontoDscto As Double 'F3MONTODSCTO
Private strObservacionesPorItem As String 'OBSERVACIONES


'Atributos Adicionales
Private dblCompromisoEAG            As Double
Private dblCompromisoPLG            As Double
Private dblLibreEAG                 As Double
Private dblLibrePLG                 As Double
Private dblStockEAG                 As Double
Private dblStockPLG                 As Double

Private strFechaInicioMes           As String
Private strFechaFinMes              As String



Private strSQLSelectAlter           As String

Private rstVale                     As ADODB.Recordset
Private strSQLVale                  As String


Public Property Let CodigoAlmacen(ByVal Value As String)
    strCodAlmacen = Value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property

Public Property Let NumeroVale(ByVal Value As String)
    strNumeroVale = Value
End Property

Public Property Get NumeroVale() As String
    NumeroVale = strNumeroVale
End Property

Public Property Let TipoVale(ByVal Value As String)
    strTipoVale = Value
End Property

Public Property Get TipoVale() As String
    TipoVale = strTipoVale
End Property

'TipoPersona
Public Property Let TipoPersona(ByVal Value As String)
    strTipoPersona = Value
End Property

Public Property Get TipoPersona() As String
    TipoPersona = strTipoPersona
End Property

Public Property Let CodigoProveedor(ByVal Value As String)
    strCodProveedor = Value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProveedor
End Property

Public Property Let NumeroDocumento(ByVal Value As String)
    strNumeroDocumento = Value
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = strNumeroDocumento
End Property

Public Property Let Fecha(ByVal Value As String)
    StrFecha = Value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property

Public Property Let CodigoOrigen(ByVal Value As String)
    strCodOrigen = Value
End Property

Public Property Get CodigoOrigen() As String
    CodigoOrigen = strCodOrigen
End Property

Public Property Let CodigoDocumento(ByVal Value As String)
    strCodDocumento = Value
End Property

Public Property Get CodigoDocumento() As String
    CodigoDocumento = strCodDocumento
End Property

Public Property Let TipoCambio(ByVal Value As Double)
    dblTCambio = Value
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dblTCambio
End Property

Public Property Let CodigoMoneda(ByVal Value As String)
    strCodMoneda = Value
End Property

Public Property Get CodigoMoneda() As String
    CodigoMoneda = strCodMoneda
End Property

Public Property Let TipoProveedor(ByVal Value As String)
    strTipoProveedor = Value
End Property

Public Property Get TipoProveedor() As String
    TipoProveedor = strTipoProveedor
End Property

Public Property Let CodPar(ByVal Value As String)
    strCodPar = Value
End Property

Public Property Get CodPar() As String
    CodPar = strCodPar
End Property

Public Property Let NumeroFactura(ByVal Value As String)
    strNumeroFactura = Value
End Property

Public Property Get NumeroFactura() As String
    NumeroFactura = strNumeroFactura
End Property

Public Property Let NumeroOrden(ByVal Value As Double)
    dblNumeroOrden = Value
End Property

Public Property Get NumeroOrden() As Double
    NumeroOrden = dblNumeroOrden
End Property

Public Property Let NumeroImportacion(ByVal Value As String)
    strNumeroImportacion = Value
End Property

Public Property Get NumeroImportacion() As String
    NumeroImportacion = strNumeroImportacion
End Property

Public Property Let Partida(ByVal Value As String)
    strPartida = Value
End Property

Public Property Get Partida() As String
    Partida = strPartida
End Property

Public Property Let OrdenTrabajo(ByVal Value As String)
    strOrdenTrabajo = Value
End Property

Public Property Get OrdenTrabajo() As String
    OrdenTrabajo = strOrdenTrabajo
End Property

Public Property Let FechaUltima(ByVal Value As String)
    strFechaUltima = Value
End Property

Public Property Get FechaUltima() As String
    FechaUltima = strFechaUltima
End Property

Public Property Let CodigoUsuario(ByVal Value As String)
    strCodUsuario = Value
End Property

Public Property Get CodigoUsuario() As String
    CodigoUsuario = strCodUsuario
End Property

Public Property Let referencia(ByVal Value As String)
    strReferencia = Value
End Property

Public Property Get referencia() As String
    referencia = strReferencia
End Property

Public Property Let CentroCosto(ByVal Value As String)
    strCentroCosto = Value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property

Public Property Let SerieGuia(ByVal Value As String)
    strSerieGuia = Value
End Property

Public Property Get SerieGuia() As String
    SerieGuia = strSerieGuia
End Property

Public Property Let NumeroGuia(ByVal Value As String)
    strNumeroGuia = Value
End Property

Public Property Get NumeroGuia() As String
    NumeroGuia = strNumeroGuia
End Property

Public Property Let CodTipoComprobante(ByVal Value As String)
    strCodTipoComprobante = Value
End Property

Public Property Get CodTipoComprobante() As String
    CodTipoComprobante = strCodTipoComprobante
End Property

Public Property Let FecReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FecReg() As String
    FecReg = strFecReg
End Property

Public Property Let UsuReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuReg() As String
    UsuReg = strUsuReg
End Property

Public Property Let FecMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FecMod() As String
    FecMod = strFecMod
End Property

Public Property Let UsuMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuMod() As String
    UsuMod = strUsuMod
End Property

Public Property Let SerieDocumento(ByVal Value As String)
    strSerieDocumento = Value
End Property

Public Property Get SerieDocumento() As String
    SerieDocumento = strSerieDocumento
End Property

Public Property Let observaciones(ByVal Value As String)
    strObservaciones = Value
End Property

Public Property Get observaciones() As String
    observaciones = strObservaciones
End Property

Public Property Let UUPP(ByVal Value As String)
    strUUPP = Value
End Property

Public Property Get UUPP() As String
    UUPP = strUUPP
End Property

Public Property Let NumeroValeEnvio(ByVal Value As String)
    strNumeroValeEnvio = Value
End Property

Public Property Get NumeroValeEnvio() As String
    NumeroValeEnvio = strNumeroValeEnvio
End Property

Public Property Let NumeroOrdenCompra(ByVal Value As String)
    strNumeroOrdenCompra = Value
End Property

Public Property Get NumeroOrdenCompra() As String
    NumeroOrdenCompra = strNumeroOrdenCompra
End Property

Public Property Let CodAlmacenDestino(ByVal Value As String)
    strCodAlmacenDestino = Value
End Property

Public Property Get CodAlmacenDestino() As String
    CodAlmacenDestino = strCodAlmacenDestino
End Property

Public Property Let CodSolicitante(ByVal Value As String)
    strCodSolicitante = Value
End Property

Public Property Get CodSolicitante() As String
    CodSolicitante = strCodSolicitante
End Property

Public Property Let NumeroValeExterno(ByVal Value As String)
    strNumeroValeExterno = Value
End Property

Public Property Get NumeroValeExterno() As String
    NumeroValeExterno = strNumeroValeExterno
End Property

Public Property Let Seccion(ByVal Value As String)
    strSeccion = Value
End Property

Public Property Get Seccion() As String
    Seccion = strSeccion
End Property

Public Property Let Lote(ByVal Value As Double)
    dblLote = Value
End Property

Public Property Get Lote() As Double
    Lote = dblLote
End Property

Public Property Let RegistroCompra(ByVal Value As String)
    strRegistroCompra = Value
End Property

Public Property Get RegistroCompra() As String
    RegistroCompra = strRegistroCompra
End Property

Public Property Let ValeIngreso(ByVal Value As String)
    strValeIngreso = Value
End Property

Public Property Get ValeIngreso() As String
    ValeIngreso = strValeIngreso
End Property
'bolExportarVale
Public Property Let ExportarVale(ByVal Value As Boolean)
    bolExportarVale = Value
End Property

Public Property Get ExportarVale() As Boolean
    ExportarVale = bolExportarVale
End Property
'VB1
Public Property Let VB1(ByVal Value As Boolean)
    bolVB1 = Value
End Property

Public Property Get VB1() As Boolean
    VB1 = bolVB1
End Property
'VB1USER
Public Property Let VB1Usuario(ByVal Value As String)
    strVB1Usuario = Value
End Property

Public Property Get VB1Usuario() As String
    VB1Usuario = strVB1Usuario
End Property
'VB1FECHA
Public Property Let VB1Fecha(ByVal Value As String)
    strVB1Fecha = Value
End Property

Public Property Get VB1Fecha() As String
    VB1Fecha = strVB1Fecha
End Property






'Atributos de Detalle de Vale
Public Property Let CodigoProducto(ByVal Value As String)
    strCodProducto = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Property Let CodigoProductoOriginal(ByVal Value As String)
    strCodProductoOriginal = Value
End Property

Public Property Get CodigoProductoOriginal() As String
    CodigoProductoOriginal = strCodProductoOriginal
End Property

Public Property Let Cantidad(ByVal Value As Double)
    dblCantidad = Value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

Public Property Let ValorVenta(ByVal Value As Double)
    dblValorVta = Value
End Property

Public Property Get ValorVenta() As Double
    ValorVenta = dblValorVta
End Property

Public Property Let IGV(ByVal Value As Double)
    dblIgv = Value
End Property

Public Property Get IGV() As Double
    IGV = dblIgv
End Property

Public Property Let TOTAL(ByVal Value As Double)
    dblTotal = Value
End Property

Public Property Get TOTAL() As Double
    TOTAL = dblTotal
End Property




Public Property Let ValorVentaDol(ByVal Value As Double)
    dblValorVtaDol = Value
End Property

Public Property Get ValorVentaDol() As Double
    ValorVentaDol = dblValorVtaDol
End Property

Public Property Let IgvDol(ByVal Value As Double)
    dblIgvDol = Value
End Property

Public Property Get IgvDol() As Double
    IgvDol = dblIgvDol
End Property

Public Property Let TotalDol(ByVal Value As Double)
    dblTotalDol = Value
End Property

Public Property Get TotalDol() As Double
    TotalDol = dblTotalDol
End Property




Public Property Let Grupo(ByVal Value As String)
    strGrupo = Value
End Property

Public Property Get Grupo() As String
    Grupo = strGrupo
End Property

Public Property Let ITEM(ByVal Value As Double)
    dblItem = Value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property

Public Property Let Requerimiento(ByVal Value As String)
    strRequerimiento = Value
End Property

Public Property Get Requerimiento() As String
    Requerimiento = strRequerimiento
End Property
'Private dblCantidadMaxima 'F3SALPEP
Public Property Let CantidadMaxima(ByVal Value As Double)
    dblCantidadMaxima = Value
End Property

Public Property Get CantidadMaxima() As Double
    CantidadMaxima = dblCantidadMaxima
End Property
'Private dblPorcentajeDscto As Double  'F3PORCENTAJEDSCTO
Public Property Let PorcentajeDscto(ByVal Value As Double)
    dblPorcentajeDscto = Value
End Property

Public Property Get PorcentajeDscto() As Double
    PorcentajeDscto = dblPorcentajeDscto
End Property
'Private dblMontoDscto As Double 'F3MONTODSCTO
Public Property Let MontoDscto(ByVal Value As Double)
    dblMontoDscto = Value
End Property

Public Property Get MontoDscto() As Double
    MontoDscto = dblMontoDscto
End Property
'Private strObservacionesPorItem As Double 'OBSERVACIONES
Public Property Let ObservacionesPorItem(ByVal Value As String)
    strObservacionesPorItem = Value
End Property

Public Property Get ObservacionesPorItem() As String
    ObservacionesPorItem = strObservacionesPorItem
End Property




'Atributos Adicionales
'Private dblCompromisoEAG            As Double
Public Property Let CompromisoEAG(ByVal Value As Double)
    dblCompromisoEAG = Value
End Property

Public Property Get CompromisoEAG() As Double
    CompromisoEAG = dblCompromisoEAG
End Property
'Private dblCompromisoPLG            As Double
Public Property Let CompromisoPLG(ByVal Value As Double)
    dblCompromisoPLG = Value
End Property

Public Property Get CompromisoPLG() As Double
    CompromisoPLG = dblCompromisoPLG
End Property
'Private dblLibreEAG                 As Double
Public Property Let LibreEAG(ByVal Value As Double)
    dblLibreEAG = Value
End Property

Public Property Get LibreEAG() As Double
    LibreEAG = dblLibreEAG
End Property
'Private dblLibrePLG                 As Double
Public Property Let LibrePLG(ByVal Value As Double)
    dblLibrePLG = Value
End Property

Public Property Get LibrePLG() As Double
    LibrePLG = dblLibrePLG
End Property
'Private dblStockEAG                 As Double
Public Property Let StockEAG(ByVal Value As Double)
    dblStockEAG = Value
End Property

Public Property Get StockEAG() As Double
    StockEAG = dblStockEAG
End Property
'Private dblStockPLG                 As Double
Public Property Let StockPLG(ByVal Value As Double)
    dblStockPLG = Value
End Property

Public Property Get StockPLG() As Double
    StockPLG = dblStockPLG
End Property




Public Property Let FechaInicioMes(ByVal Value As String)
    strFechaInicioMes = Value
End Property

Public Property Get FechaInicioMes() As String
    FechaInicioMes = strFechaInicioMes
End Property

Public Property Let FechaFinMes(ByVal Value As String)
    strFechaFinMes = Value
End Property

Public Property Get FechaFinMes() As String
    FechaFinMes = strFechaFinMes
End Property






Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodAlmacen = vbNullString 'F2CODALM
    strNumeroVale = vbNullString 'F4NUMVAL
    
    strTipoVale = vbNullString 'F4TIPOVALE
    strTipoPersona = vbNullString 'F1TIPPRV
    strCodProveedor = vbNullString 'F2CODPROV
    strNumeroDocumento = vbNullString 'F4NUMDOC
    StrFecha = vbNullString    'F4FECVAL
    strCodOrigen = vbNullString 'F1CODORI
    strCodDocumento = vbNullString 'F1CODDOC
    dblTCambio = 0 'F4TIPCAM
    strCodMoneda = vbNullString 'F4MONEDA
    strTipoProveedor = vbNullString 'F1TIPPRV
    strCodPar = vbNullString   'F2CODPAR
    strNumeroFactura = vbNullString 'F4NUMFAC
    dblNumeroOrden = 0 'F4NUMORD
    strNumeroImportacion = vbNullString 'F4NUMIMP
    strPartida = vbNullString  'F4PARTIDA
    strOrdenTrabajo = vbNullString 'F4ORDTRA
    strFechaUltima = vbNullString 'F4FECULT
    strCodUsuario = vbNullString 'F2CODUSE
    strReferencia = vbNullString 'F4REFERE
    strCentroCosto = vbNullString 'F4CENTRO
    strSerieGuia = vbNullString 'F4SERGUIA
    strNumeroGuia = vbNullString 'F4NUMGUIA
    strCodTipoComprobante = vbNullString 'F4TIPDOC
    strFecReg = vbNullString   'F4FECGRA
    strUsuReg = vbNullString   'F4USEGRA
    strFecMod = vbNullString   'F4FECMOD
    strUsuMod = vbNullString   'F4USEMOD
    strSerieDocumento = vbNullString 'F4SERDOC
    strObservaciones = vbNullString 'F4OBSERVA
    strUUPP = vbNullString     'F4UUPP
    strNumeroValeEnvio = vbNullString 'F4NUMVALENVIO
    strNumeroOrdenCompra = vbNullString 'NUMORDEN
    strCodAlmacenDestino = vbNullString 'CODALMDES
    strCodSolicitante = vbNullString 'F2CODUSER
    strNumeroValeExterno = vbNullString 'NUMENSAM
    strSeccion = vbNullString  'SECCION
    dblLote = 0 'LOTE
    strRegistroCompra = vbNullString 'F4REGCOM
    strValeIngreso = vbNullString 'VALING
    bolExportarVale = False 'EXPORTARVALE
    bolVB1 = False 'VB1
    strVB1Usuario = vbNullString 'VBUSER
    strVB1Fecha = vbNullString 'VB1FECHA
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesDetalle()
    strCodProducto = vbNullString 'F5CODPRO
    strCodProductoOriginal = vbNullString 'F5CODPROORIGINAL
    dblCantidad = 0 'F3CANPRO
    dblValorVta = 0 'F3VALVTA
    dblIgv = 0 'F3IGV
    dblTotal = 0 'F3TOTITE
    
    dblValorVtaDol = 0 'F3VALDOL
    dblIgvDol = 0 'F3IGVDOL
    dblTotalDol = 0 'F3TOTDOL
    
    strGrupo = vbNullString 'F3GRUPO
    dblItem = 0 'F3ITEM
    strRequerimiento = vbNullString  'COD_SOLICITUD
    dblCantidadMaxima = 0 'F3SALPEP
    dblPorcentajeDscto = 0 'F3PORCENTAJEDSCTO
    dblMontoDscto = 0 'F3MONTODSCTO
    strObservacionesPorItem = vbNullString 'OBSERVACIONES
End Sub

Public Sub inicializarEntidadesAdicionales()
    dblCompromisoEAG = 0
    dblCompromisoPLG = 0
    dblLibreEAG = 0
    dblLibrePLG = 0
    dblStockEAG = 0
    dblStockPLG = 0
    
    strFechaInicioMes = vbNullString
    strFechaFinMes = vbNullString
End Sub

Public Function obtenerVale() As Boolean
    On Error GoTo errObtenerVale
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "* "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        strCodAlmacen = Trim(rstVale!f2codalm & "") 'F2CODALM
        strNumeroVale = Trim(rstVale!F4NUMVAL & "") 'F4NUMVAL
        
        strTipoVale = Trim(rstVale!F4TIPOVALE & "") 'F4TIPOVALE
        strTipoPersona = Trim(rstVale!F1TIPPRV & "") 'F1TIPPRV
        strCodProveedor = Trim(rstVale!F2CODPROV & "") 'F2CODPROV
        strNumeroDocumento = Trim(rstVale!F4NUMDOC & "") 'F4NUMDOC
        StrFecha = Trim(rstVale!F4FECVAL & "")    'F4FECVAL
        strCodOrigen = Trim(rstVale!F1CODORI & "") 'F1CODORI
        strCodDocumento = Trim(rstVale!F1CODDOC & "") 'F1CODDOC
        dblTCambio = Val(rstVale!F4TIPCAM & "") 'F4TIPCAM
        strCodMoneda = Trim(rstVale!F4MONEDA & "") 'F4MONEDA
        strTipoProveedor = Trim(rstVale!F1TIPPRV & "") 'F1TIPPRV
        strCodPar = Trim(rstVale!F2CODPAR & "")   'F2CODPAR
        strNumeroFactura = Trim(rstVale!F4NUMFAC & "") 'F4NUMFAC
        dblNumeroOrden = Val(rstVale!F4NUMORD & "") 'F4NUMORD
        strNumeroImportacion = Trim(rstVale!F4NUMIMP & "") 'F4NUMIMP
        strPartida = Trim(rstVale!F4PARTIDA & "")  'F4PARTIDA
        strOrdenTrabajo = Trim(rstVale!F4ORDTRA & "") 'F4ORDTRA
        strFechaUltima = Trim(rstVale!F4FECULT & "") 'F4FECULT
        strCodUsuario = Trim(rstVale!F2CODUSE & "") 'F2CODUSE
        strReferencia = Trim(rstVale!F4REFERE & "") 'F4REFERE
        strCentroCosto = Trim(rstVale!F4CENTRO & "") 'F4CENTRO
        strSerieGuia = Trim(rstVale!F4SERGUIA & "") 'F4SERGUIA
        strNumeroGuia = Trim(rstVale!F4NUMGUIA & "") 'F4NUMGUIA
        strCodTipoComprobante = Trim(rstVale!F4TIPDOC & "") 'F4TIPDOC
        strFecReg = Trim(rstVale!F4FECGRA & "")   'F4FECGRA
        strUsuReg = Trim(rstVale!F4USEGRA & "")   'F4USEGRA
        strFecMod = Trim(rstVale!F4FECMOD & "")   'F4FECMOD
        strUsuMod = Trim(rstVale!F4USEMOD & "")   'F4USEMOD
        strSerieDocumento = Trim(rstVale!F4SERDOC & "") 'F4SERDOC
        strObservaciones = Trim(rstVale!F4OBSERVA & "") 'F4OBSERVA
        strUUPP = Trim(rstVale!F4UUPP & "")     'F4UUPP
        strNumeroValeEnvio = Trim(rstVale!F4NUMVALENVIO & "") 'F4NUMVALENVIO
        strNumeroOrdenCompra = Trim(rstVale!NUMORDEN & "") 'NUMORDEN
        strCodAlmacenDestino = Trim(rstVale!CODALMDES & "") 'CODALMDES
        strCodSolicitante = Trim(rstVale!F2CODUSER & "") 'F2CODUSER
        strNumeroValeExterno = Trim(rstVale!NUMENSAM & "") 'NUMENSAM
        strSeccion = Trim(rstVale!Seccion & "")  'SECCION
        dblLote = Val(rstVale!Lote & "") 'LOTE
        strRegistroCompra = Trim(rstVale!F4REGCOM & "") 'F4REGCOM
        strValeIngreso = Trim(rstVale!VALING & "") 'VALING
        bolExportarVale = CBool(rstVale!ExportarVale) 'EXPORTARVALE
        bolVB1 = CBool(rstVale!VB1) 'VB1
        strVB1Usuario = Trim(rstVale!VB1USER & "") 'VB1USER
        strVB1Fecha = Trim(rstVale!VB1Fecha & "") 'VB1FECHA
        
        
        'Select Case strTipoVale
        '    Case "I"
        '        obtenerValeDetalleIngreso
        '    Case "S"
        '        obtenerValeDetalleSalida
        'End Select
        
        obtenerVale = True
    Else
        obtenerVale = False
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    Exit Function
errObtenerVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ObtenerVale"
    
    obtenerVale = False
    
    Err.Clear
End Function

Public Sub obtenerConfigVale()
    On Error GoTo errObtenerConfigVale
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "* "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        strCodAlmacen = Trim(rstVale!f2codalm & "") 'F2CODALM
        strNumeroVale = Trim(rstVale!F4NUMVAL & "") 'F4NUMVAL
        
        strTipoVale = Trim(rstVale!F4TIPOVALE & "") 'F4TIPOVALE
        strTipoPersona = Trim(rstVale!F1TIPPRV & "") 'F1TIPPRV
        strCodProveedor = Trim(rstVale!F2CODPROV & "") 'F2CODPROV
        strNumeroDocumento = Trim(rstVale!F4NUMDOC & "") 'F4NUMDOC
        StrFecha = Trim(rstVale!F4FECVAL & "")    'F4FECVAL
        strCodOrigen = Trim(rstVale!F1CODORI & "") 'F1CODORI
        strCodDocumento = Trim(rstVale!F1CODDOC & "") 'F1CODDOC
        dblTCambio = Val(rstVale!F4TIPCAM & "") 'F4TIPCAM
        strCodMoneda = Trim(rstVale!F4MONEDA & "") 'F4MONEDA
        strTipoProveedor = Trim(rstVale!F1TIPPRV & "") 'F1TIPPRV
        strCodPar = Trim(rstVale!F2CODPAR & "")   'F2CODPAR
        strNumeroFactura = Trim(rstVale!F4NUMFAC & "") 'F4NUMFAC
        dblNumeroOrden = Val(rstVale!F4NUMORD & "") 'F4NUMORD
        strNumeroImportacion = Trim(rstVale!F4NUMIMP & "") 'F4NUMIMP
        strPartida = Trim(rstVale!F4PARTIDA & "")  'F4PARTIDA
        strOrdenTrabajo = Trim(rstVale!F4ORDTRA & "") 'F4ORDTRA
        strFechaUltima = Trim(rstVale!F4FECULT & "") 'F4FECULT
        strCodUsuario = Trim(rstVale!F2CODUSE & "") 'F2CODUSE
        strReferencia = Trim(rstVale!F4REFERE & "") 'F4REFERE
        strCentroCosto = Trim(rstVale!F4CENTRO & "") 'F4CENTRO
        strSerieGuia = Trim(rstVale!F4SERGUIA & "") 'F4SERGUIA
        strNumeroGuia = Trim(rstVale!F4NUMGUIA & "") 'F4NUMGUIA
        strCodTipoComprobante = Trim(rstVale!F4TIPDOC & "") 'F4TIPDOC
        strFecReg = Trim(rstVale!F4FECGRA & "")   'F4FECGRA
        strUsuReg = Trim(rstVale!F4USEGRA & "")   'F4USEGRA
        strFecMod = Trim(rstVale!F4FECMOD & "")   'F4FECMOD
        strUsuMod = Trim(rstVale!F4USEMOD & "")   'F4USEMOD
        strSerieDocumento = Trim(rstVale!F4SERDOC & "") 'F4SERDOC
        strObservaciones = Trim(rstVale!F4OBSERVA & "") 'F4OBSERVA
        strUUPP = Trim(rstVale!F4UUPP & "")     'F4UUPP
        strNumeroValeEnvio = Trim(rstVale!F4NUMVALENVIO & "") 'F4NUMVALENVIO
        strNumeroOrdenCompra = Trim(rstVale!NUMORDEN & "") 'NUMORDEN
        strCodAlmacenDestino = Trim(rstVale!CODALMDES & "") 'CODALMDES
        strCodSolicitante = Trim(rstVale!F2CODUSER & "") 'F2CODUSER
        strNumeroValeExterno = Trim(rstVale!NUMENSAM & "") 'NUMENSAM
        strSeccion = Trim(rstVale!Seccion & "")  'SECCION
        dblLote = Val(rstVale!Lote & "") 'LOTE
        strRegistroCompra = Trim(rstVale!F4REGCOM & "") 'F4REGCOM
        strValeIngreso = Trim(rstVale!VALING & "") 'VALING
        bolExportarVale = CBool(rstVale!ExportarVale) 'EXPORTARVALE
        bolVB1 = CBool(rstVale!VB1) 'VB1
        strVB1Usuario = Trim(rstVale!VB1USER & "") 'VB1USER
        strVB1Fecha = Trim(rstVale!VB1Fecha & "") 'VB1FECHA
    End If
    
    rstVale.Close
    
    Set rstVale = Nothing
    
    Exit Sub
errObtenerConfigVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ObtenerConfigVale"
    
    Err.Clear
End Sub

Public Sub obtenerValeDetalleIngreso()
    Dim rstValeDetalle As New ADODB.Recordset
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO"
    
    abrirCnTemporal
    
'    strSQLVale = vbNullString
'    strSQLVale = strSQLVale & "INSERT INTO TMPVALEINGRESO("
'        strSQLVale = strSQLVale & "ITEM, F4NUMORD, COD_SOLICITUD, "
'        strSQLVale = strSQLVale & "CODPROD, CODPRODORIGINAL, "
'        strSQLVale = strSQLVale & "DESCRIPCION, UMEDIDA, CANTIDAD, "
'        'strSQLVale = strSQLVale & "COSTOUNI, PVUNIT, PORDESC, VALDESC, VVTOTAL, IGV, "
'        'strSQLVale = strSQLVale & "TOTAL, AFECTO, "
'        strSQLVale = strSQLVale & "CANTIDADMAX, COSTOUNI, "
'        strSQLVale = strSQLVale & "PORDESC, AFECTO, OBSERVACIONES"
'        strSQLVale = strSQLVale & ") "
'
'    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLVale = strSQLVale & "SELECT "
        strSQLVale = strSQLVale & "DET.F3ITEM, "
        strSQLVale = strSQLVale & "F4NUMORD, "
        strSQLVale = strSQLVale & "COD_SOLICITUD, "
        strSQLVale = strSQLVale & "DET.F5CODPRO, "
        strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
        strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
        strSQLVale = strSQLVale & "MED.F7SIGMED, "
        strSQLVale = strSQLVale & "DET.F3CANPRO, "
        strSQLVale = strSQLVale & "DET.F3SALPEP, "
        
        Select Case strCodMoneda
            Case "S"
                strSQLVale = strSQLVale & "DET.F3VALVTA AS COSTO, "
                'strSQLVale = strSQLVale & "(DET.F3VALVTA * (1 + " & wIgv & "/100)) AS VALVTA, "
                'strSQLVale = strSQLVale & "(DET.F3CANPRO * DET.F3VALVTA) AS SUBTOTAL, "
                'strSQLVale = strSQLVale & "DET.F3IGV AS IGV, "
                'strSQLVale = strSQLVale & "DET.F3TOTITE AS TOTAL, "
            Case Else
                strSQLVale = strSQLVale & "DET.F3VALDOL AS COSTO, "
                'strSQLVale = strSQLVale & "(DET.F3VALDOL * (1 + " & wIgv & "/100)) AS VALVTA, "
                'strSQLVale = strSQLVale & "(DET.F3CANPRO * DET.F3VALDOL) AS SUBTOTAL, "
                'strSQLVale = strSQLVale & "DET.F3IGVDOL AS IGV, "
                'strSQLVale = strSQLVale & "DET.F3TOTDOL AS TOTAL, "
        End Select
        
        strSQLVale = strSQLVale & "DET.F3PORCENTAJEDSCTO AS DSCTO, "
        strSQLVale = strSQLVale & "PROD.F5AFECTO, "
        strSQLVale = strSQLVale & "DET.OBSERVACIONES "
        
    strSQLVale = strSQLVale & "FROM "
        strSQLVale = strSQLVale & "(PROCESOS.IF3VALES AS DET "
        strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
        strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
    strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "DET.F4NUMVAL = '" & strNumeroVale & "'"
    
    'cnBdCPlus.Execute strSQLVale
    
    If rstValeDetalle.State = 1 Then rstValeDetalle.Close
    
    rstValeDetalle.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeDetalle.EOF Then
        Do While Not rstValeDetalle.EOF
            strSQLVale = vbNullString
            strSQLVale = strSQLVale & "INSERT INTO TMPVALEINGRESO("
                strSQLVale = strSQLVale & "ITEM, F4NUMORD, COD_SOLICITUD, "
                strSQLVale = strSQLVale & "CODPROD, CODPRODORIGINAL, "
                strSQLVale = strSQLVale & "DESCRIPCION, UMEDIDA, CANTIDAD, "
                strSQLVale = strSQLVale & "CANTIDADMAX, COSTOUNI, "
                strSQLVale = strSQLVale & "PORDESC, AFECTO, OBSERVACIONES"
                strSQLVale = strSQLVale & ") "
                
            strSQLVale = strSQLVale & "VALUES("
                strSQLVale = strSQLVale & Val(rstValeDetalle!F3ITEM & "") & ", "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F4NUMORD & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!COD_SOLICITUD & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!f5codpro & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F5CODPROORIGINAL & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F5NOMPRO & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F7SIGMED & "") & "', "
                strSQLVale = strSQLVale & Val(rstValeDetalle!F3CANPRO & "") & ", "
                strSQLVale = strSQLVale & Val(rstValeDetalle!F3SALPEP & "") & ", "
                strSQLVale = strSQLVale & Val(rstValeDetalle!costo & "") & ", "
                strSQLVale = strSQLVale & Val(rstValeDetalle!DSCTO & "") & ", "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F5AFECTO & "") & "', "
                strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!observaciones & "") & "'"
                strSQLVale = strSQLVale & ")"
            
            cnDBTemp.Execute strSQLVale
            
            rstValeDetalle.MoveNext
        Loop
    End If
    
    If rstValeDetalle.State = 1 Then rstValeDetalle.Close
    
    Set rstValeDetalle = Nothing
    
    strSQLVale = vbNullString
End Sub

Public Sub obtenerValeDetalleSalida()
    Dim rstValeDetalle As New ADODB.Recordset
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA"
    
    abrirCnTemporal
    
'    strSQLVale = vbNullString
'    strSQLVale = strSQLVale & "INSERT INTO TMPVALESALIDA("
'        strSQLVale = strSQLVale & "ITEM, CODPROD, CODPRODORIGINAL, "
'        strSQLVale = strSQLVale & "DESCRIPCION, UMEDIDA, CANTIDAD, "
'        strSQLVale = strSQLVale & "PUNIT, TOTAL, F4NUMORD, "
'        strSQLVale = strSQLVale & "COD_SOLICITUD) "
'
'    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLVale = strSQLVale & "SELECT "
        strSQLVale = strSQLVale & "DET.F3ITEM, "
        strSQLVale = strSQLVale & "DET.F5CODPRO, "
        strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
        strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
        strSQLVale = strSQLVale & "MED.F7SIGMED, "
        strSQLVale = strSQLVale & "DET.F3CANPRO, "
        
        Select Case strCodMoneda
            Case "S"
                strSQLVale = strSQLVale & "DET.F3VALVTA AS COSTO, "
                strSQLVale = strSQLVale & "(DET.F3CANPRO * DET.F3VALVTA) AS SUBTOTAL, "
            Case Else
                strSQLVale = strSQLVale & "DET.F3VALDOL AS COSTO, "
                strSQLVale = strSQLVale & "(DET.F3CANPRO * DET.F3VALDOL) AS SUBTOTAL, "
        End Select
        
        strSQLVale = strSQLVale & "F4NUMORD, "
        strSQLVale = strSQLVale & "COD_SOLICITUD "
    strSQLVale = strSQLVale & "FROM "
        strSQLVale = strSQLVale & "(PROCESOS.IF3VALES AS DET "
        strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
        strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
    strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
    
    'cnBdCPlus.Execute strSQLVale
    
    If rstValeDetalle.State = 1 Then rstValeDetalle.Close
    
    rstValeDetalle.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeDetalle.EOF Then
        Do While Not rstValeDetalle.EOF
            strSQLVale = vbNullString
            strSQLVale = strSQLVale & "INSERT INTO TMPVALESALIDA("
                strSQLVale = strSQLVale & "ITEM, CODPROD, CODPRODORIGINAL, "
                strSQLVale = strSQLVale & "DESCRIPCION, UMEDIDA, CANTIDAD, "
                strSQLVale = strSQLVale & "PUNIT, TOTAL, F4NUMORD, "
                strSQLVale = strSQLVale & "COD_SOLICITUD) "
                
            strSQLVale = strSQLVale & "VALUES("
            strSQLVale = strSQLVale & Val(rstValeDetalle!F3ITEM & "") & ", "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!f5codpro & "") & "', "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F5CODPROORIGINAL & "") & "', "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F5NOMPRO & "") & "', "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F7SIGMED & "") & "', "
            strSQLVale = strSQLVale & Val(rstValeDetalle!F3CANPRO & "") & ", "
            strSQLVale = strSQLVale & Val(rstValeDetalle!costo & "") & ", "
            strSQLVale = strSQLVale & Val(rstValeDetalle!SUBTOTAL & "") & ", "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!F4NUMORD & "") & "', "
            strSQLVale = strSQLVale & "'" & Trim(rstValeDetalle!COD_SOLICITUD & "") & "'"
            strSQLVale = strSQLVale & ")"
            
            cnDBTemp.Execute strSQLVale
            
            rstValeDetalle.MoveNext
        Loop
    End If
    
    If rstValeDetalle.State = 1 Then rstValeDetalle.Close
    
    Set rstValeDetalle = Nothing
    
    strSQLVale = vbNullString
End Sub

Public Function obtenerRstValeCompraPorProvYdocumento() As ADODB.Recordset
                                    
    On Error GoTo errObtenerRstValeCompraPorProvYdocumento
    
    Dim rsVale As New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "* "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F4TIPOVALE = '" & strTipoVale & "' AND "
    strSQLVale = strSQLVale & "F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLVale = strSQLVale & "F4TIPDOC = '" & strCodTipoComprobante & "' AND "
    strSQLVale = strSQLVale & "F4SERDOC = '" & strSerieDocumento & "' AND "
    strSQLVale = strSQLVale & "F4NUMDOC = '" & strNumeroDocumento & "' "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "F4REGCOM DESC, "
    strSQLVale = strSQLVale & "F4FECVAL DESC"
    
    If rsVale.State = 1 Then rsVale.Close
    
    rsVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    Set obtenerRstValeCompraPorProvYdocumento = rsVale
    
    Exit Function
errObtenerRstValeCompraPorProvYdocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ObtenerRstValeCompraPorProvYdocumento"
    End Select
    
    Set obtenerRstValeCompraPorProvYdocumento = Nothing
    
    Err.Clear
End Function

Public Function obtenerRstValeDetalleCompraPorProvYdocumento() As ADODB.Recordset
                                    
    On Error GoTo errObtenerRstValeDetalleCompraPorProvYdocumento
    
    Dim rsValeDet As New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "(CASE WHEN PROV.F2TIPPROV = 'N' THEN PROD.F5CTACON ELSE PROD.F5CTACONIMP) AS CUENTA, "
    strSQLVale = strSQLVale & "(CASE WHEN PROV.F2TIPPROV = 'N' THEN PROD.F5ANEXO ELSE PROD.F5ANEXOIMP) AS AUXILIAR, "
    'strSQLVale = strSQLVale & "SUM(DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL)) AS SUBTOTAL, "
    'strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3IGV, DET.F3IGVDOL)) AS IGV, "
    'strSQLVale = strSQLVale & "SUM(F3MONTODSCTO) AS DSCTO, "
    'strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'S', F3TOTITE, F3TOTDOL)) AS TOTAL, "
    
    strSQLVale = strSQLVale & "SUM( CONVERT(DECIMAL(10, 2), DET.F3CANPRO * (CASE WHEN CAB.F4MONEDA = 'S' THEN DET.F3VALVTA ELSE DET.F3VALDOL) ) ) AS SUBTOTAL, "
    strSQLVale = strSQLVale & "SUM((CASE WHEN CAB.F4MONEDA = 'S' THEN DET.F3IGV ELSE DET.F3IGVDOL)) AS IGV, "
    'strSQLVale = strSQLVale & "SUM( DET.F3MONTODSCTO ) AS DSCTO, "
    
    strSQLVale = strSQLVale & "SUM( CONVERT(DECIMAL(10, 2), (DET.F3CANPRO * (CASE WHEN CAB.F4MONEDA = 'S' THEN DET.F3VALVTA ELSE DET.F3VALDOL)) * (DET.F3PORCENTAJEDSCTO / 100) ) ) AS DSCTO, "
    
    strSQLVale = strSQLVale & "SUM((CASE WHEN CAB.F4MONEDA = 'S' THEN DET.F3TOTITE ELSE DET.F3TOTDOL)) AS TOTAL, "
    
    strSQLVale = strSQLVale & "PROD.F5AFECTO, "
    strSQLVale = strSQLVale & "DET.F4NUMORD AS NROOC "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2CODPROV = CAB.F2CODPROV "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4TIPOVALE = '" & strTipoVale & "' AND "
    strSQLVale = strSQLVale & "CAB.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLVale = strSQLVale & "CAB.F4TIPDOC = '" & strCodTipoComprobante & "' AND "
    strSQLVale = strSQLVale & "CAB.F4SERDOC = '" & strSerieDocumento & "' AND "
    strSQLVale = strSQLVale & "CAB.F4NUMDOC = '" & strNumeroDocumento & "' "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "(CASE WHEN PROV.F2TIPPROV = 'N' THEN PROD.F5CTACON ELSE PROD.F5CTACONIMP), "
    strSQLVale = strSQLVale & "(CASE WHEN PROV.F2TIPPROV = 'N' THEN PROD.F5ANEXO ELSE PROD.F5ANEXOIMP), "
    strSQLVale = strSQLVale & "PROD.F5AFECTO, "
    strSQLVale = strSQLVale & "DET.F4NUMORD"
    
    If rsValeDet.State = 1 Then rsValeDet.Close
    
    rsValeDet.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    Set obtenerRstValeDetalleCompraPorProvYdocumento = rsValeDet
    
    Exit Function
errObtenerRstValeDetalleCompraPorProvYdocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ObtenerRstValeDetalleCompraPorProvYdocumento"
    End Select
    
    Set obtenerRstValeDetalleCompraPorProvYdocumento = Nothing
    
    Err.Clear
End Function

Public Function obtenerConceptoCompraPorProvYdocumento() As String
    
    On Error GoTo errObtenerConceptoCompraPorProvYdocumento
    
    Dim rsValeDet As New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
    strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
    strSQLVale = strSQLVale & "MED.F7SIGMED AS UM, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "((((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    strSQLVale = strSQLVale & "WHERE "
    
    strSQLVale = strSQLVale & "CAB.F4TIPOVALE = '" & strTipoVale & "' AND "
    strSQLVale = strSQLVale & "CAB.F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLVale = strSQLVale & "CAB.F4TIPDOC = '" & strCodTipoComprobante & "' AND "
    strSQLVale = strSQLVale & "CAB.F4SERDOC = '" & strSerieDocumento & "' AND "
    strSQLVale = strSQLVale & "CAB.F4NUMDOC = '" & strNumeroDocumento & "' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "FAM.F7DESCON, "
    strSQLVale = strSQLVale & "SFAM.F7DESCON, "
    strSQLVale = strSQLVale & "MED.F7SIGMED"
    
    If rsValeDet.State = 1 Then rsValeDet.Close
    
    rsValeDet.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rsValeDet.EOF Then
        
        Do While Not rsValeDet.EOF
            If obtenerConceptoCompraPorProvYdocumento = vbNullString Then
                obtenerConceptoCompraPorProvYdocumento = Trim(rsValeDet!SUBFAMILIA & "") & " ( " & Val(rsValeDet!Cantidad & "") & " " & Trim(rsValeDet!um & "") & " )"
            Else
                obtenerConceptoCompraPorProvYdocumento = obtenerConceptoCompraPorProvYdocumento & ", " & Trim(rsValeDet!SUBFAMILIA & "") & " ( " & Val(rsValeDet!Cantidad & "") & " " & Trim(rsValeDet!um & "") & " )"
            End If
            
            rsValeDet.MoveNext
        Loop
    End If
    
    Exit Function
errObtenerConceptoCompraPorProvYdocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCompra: ObtenerConceptoCompraPorProvYdocumento"
    End Select
    
    obtenerConceptoCompraPorProvYdocumento = vbNullString
    
    Err.Clear
End Function

Public Function generarNumeroVale() As String
    On Error GoTo errGenerarNumeroVale
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "TOP 1 F4NUMVAL "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLVale = strSQLVale & "F4NUMVAL LIKE '" & strTipoVale & "-" & Format(StrFecha, "yyyymm") & "%' "
    strSQLVale = strSQLVale & "ORDER BY "
    'strSQLVale = strSQLVale & "F4NUMVAL DESC"
    strSQLVale = strSQLVale & "SUBSTRING(F4NUMVAL, LEN('" & strTipoVale & "-" & Format(StrFecha, "yyyymm") & "') + 1, LEN(F4NUMVAL)) DESC"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        'generarNumeroVale = strTipoVale & "-" & Format(Val(Mid(Trim(rstVale!F4NUMVAL & ""), InStr(1, Trim(rstVale!F4NUMVAL & ""), "-") + 1)) + 1, "00000000000")
        generarNumeroVale = strTipoVale & "-" & Format(StrFecha, "yyyymm") & Format(Val(Mid(Trim(rstVale!F4NUMVAL & ""), Len(strTipoVale & "-" & Format(StrFecha, "yyyymm")) + 1)) + 1, "000000")
    Else
        generarNumeroVale = strTipoVale & "-" & Format(StrFecha, "yyyymm") & "000001"
    End If
    
    rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errGenerarNumeroVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: GenerarNumeroVale"
    
    generarNumeroVale = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia(Optional ByVal bolRegistroImportado As Boolean) As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES.* "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
        
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstVale.Close
    
    strSQLVale = vbNullString
    
    Set rstVale = Nothing
    
    Exit Function
    
    Resume
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function verificarExistenciaPorNumRef() As Boolean
    On Error GoTo errVerificarExistenciaPorNumRef
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES.* "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F2CODPROV = '" & strCodProveedor & "' AND "
    strSQLVale = strSQLVale & "F4SERGUIA = '" & strSerieGuia & "' AND "
    strSQLVale = strSQLVale & "F4NUMGUIA = '" & strNumeroGuia & "' AND "
    strSQLVale = strSQLVale & "F4TIPDOC = '" & strCodTipoComprobante & "' AND "
    strSQLVale = strSQLVale & "F4SERDOC = '" & strSerieDocumento & "' AND "
    strSQLVale = strSQLVale & "F4NUMDOC = '" & strNumeroDocumento & "'"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        verificarExistenciaPorNumRef = True
    Else
        verificarExistenciaPorNumRef = False
    End If
    
    rstVale.Close
    
    strSQLVale = vbNullString
    
    Set rstVale = Nothing
    
    Exit Function
errVerificarExistenciaPorNumRef:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarExistenciaPorNumRef"
    
    verificarExistenciaPorNumRef = False
    
    Err.Clear
End Function

Public Function guardarVale(Optional ByVal bolRegistroImportado As Boolean) As Boolean
    On Error GoTo errGuardarVale
    
    strSQLVale = vbNullString
    
    If Not verificarExistencia(bolRegistroImportado) Then
        If strNumeroVale = vbNullString Then
            strNumeroVale = generarNumeroVale
        End If
        
        strSQLVale = vbNullString
        strSQLVale = strSQLVale & "INSERT INTO PROCESOS.IF4VALES("
        strSQLVale = strSQLVale & "F2CODALM, F4NUMVAL, F4TIPOVALE, "
        strSQLVale = strSQLVale & "F1TIPPRV, F2CODPROV, F4NUMDOC, F4FECVAL, "
        strSQLVale = strSQLVale & "F1CODORI, F1CODDOC, F4TIPCAM, "
        strSQLVale = strSQLVale & "F4MONEDA, F2CODPAR, "
        strSQLVale = strSQLVale & "F4NUMFAC, F4NUMORD, F4NUMIMP, "
        strSQLVale = strSQLVale & "F4PARTIDA, F4ORDTRA, F4FECULT, "
        strSQLVale = strSQLVale & "F2CODUSE, F4REFERE, F4CENTRO, "
        strSQLVale = strSQLVale & "F4SERGUIA, F4NUMGUIA, F4TIPDOC, "
        strSQLVale = strSQLVale & "F4FECGRA, F4USEGRA, F4SERDOC, "
        strSQLVale = strSQLVale & "F4OBSERVA, F4UUPP, F4NUMVALENVIO, "
        strSQLVale = strSQLVale & "NUMORDEN, CODALMDES, F2CODUSER, "
        strSQLVale = strSQLVale & "NUMENSAM, SECCION, LOTE, "
        strSQLVale = strSQLVale & "F4REGCOM, VALING, EXPORTARVALE) "
        
        strSQLVale = strSQLVale & "VALUES("
        strSQLVale = strSQLVale & "'" & strCodAlmacen & "', '" & strNumeroVale & "', '" & strTipoVale & "', "
        strSQLVale = strSQLVale & "'" & strTipoPersona & "', '" & strCodProveedor & "', '" & strNumeroDocumento & "', '" & StrFecha & "', "
        strSQLVale = strSQLVale & "'" & strCodOrigen & "', '" & strCodDocumento & "', " & dblTCambio & ", "
        strSQLVale = strSQLVale & "'" & strCodMoneda & "', '" & strCodPar & "', "
        strSQLVale = strSQLVale & "'" & strNumeroFactura & "', " & dblNumeroOrden & ", '" & strNumeroImportacion & "', "
        strSQLVale = strSQLVale & "'" & strPartida & "', '" & strOrdenTrabajo & "', "
        strSQLVale = strSQLVale & IIf(strFechaUltima = vbNullString, "NULL", "'" & strFechaUltima & "'") & ", "
        strSQLVale = strSQLVale & "'" & strCodUsuario & "', '" & strReferencia & "', '" & strCentroCosto & "', "
        strSQLVale = strSQLVale & "'" & strSerieGuia & "', '" & strNumeroGuia & "', '" & strCodTipoComprobante & "', "
        strSQLVale = strSQLVale & "'" & strFecReg & "', '" & strUsuReg & "', '" & strSerieDocumento & "', "
        strSQLVale = strSQLVale & "'" & strObservaciones & "', '" & strUUPP & "', '" & strNumeroValeEnvio & "', "
        strSQLVale = strSQLVale & "'" & strNumeroOrdenCompra & "', '" & strCodAlmacenDestino & "', '" & strCodSolicitante & "', "
        strSQLVale = strSQLVale & "'" & strNumeroValeExterno & "', '" & strSeccion & "', " & dblLote & ", "
        strSQLVale = strSQLVale & IIf(strRegistroCompra = vbNullString, "''", "'" & strRegistroCompra & "'") & ", '" & strValeIngreso & "', "
        strSQLVale = strSQLVale & IIf(bolExportarVale, "1", "0") & ")"
    Else
        strSQLVale = vbNullString
        strSQLVale = strSQLVale & "UPDATE "
        strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
        strSQLVale = strSQLVale & "SET "
        
        strSQLVale = strSQLVale & "F4TIPOVALE = '" & strTipoVale & "', "
        strSQLVale = strSQLVale & "F1TIPPRV = '" & strTipoPersona & "', "
        strSQLVale = strSQLVale & "F2CODPROV = '" & strCodProveedor & "', "
        strSQLVale = strSQLVale & "F4NUMDOC = '" & strNumeroDocumento & "', "
        strSQLVale = strSQLVale & "F4FECVAL = '" & StrFecha & "', "
        strSQLVale = strSQLVale & "F1CODORI = '" & strCodOrigen & "', "
        strSQLVale = strSQLVale & "F1CODDOC = '" & strCodDocumento & "', "
        strSQLVale = strSQLVale & "F4TIPCAM = " & dblTCambio & ", "
        strSQLVale = strSQLVale & "F4MONEDA = '" & strCodMoneda & "', "
        strSQLVale = strSQLVale & "F2CODPAR = '" & strCodPar & "', "
        strSQLVale = strSQLVale & "F4NUMFAC = '" & strNumeroFactura & "', "
        strSQLVale = strSQLVale & "F4NUMORD = " & dblNumeroOrden & ", "
        strSQLVale = strSQLVale & "F4NUMIMP = '" & strNumeroImportacion & "', "
        strSQLVale = strSQLVale & "F4PARTIDA = '" & strPartida & "', "
        strSQLVale = strSQLVale & "F4ORDTRA = '" & strOrdenTrabajo & "', "
        strSQLVale = strSQLVale & "F4FECULT = " & IIf(strFechaUltima = vbNullString, "NULL", "'" & strFechaUltima & "'") & ", "
        strSQLVale = strSQLVale & "F2CODUSE = '" & strCodUsuario & "', "
        strSQLVale = strSQLVale & "F4REFERE = '" & strReferencia & "', "
        strSQLVale = strSQLVale & "F4CENTRO = '" & strCentroCosto & "', "
        strSQLVale = strSQLVale & "F4SERGUIA = '" & strSerieGuia & "', "
        strSQLVale = strSQLVale & "F4NUMGUIA = '" & strNumeroGuia & "', "
        strSQLVale = strSQLVale & "F4TIPDOC = '" & strCodTipoComprobante & "', "
        strSQLVale = strSQLVale & "F4FECMOD = '" & strFecMod & "', "
        strSQLVale = strSQLVale & "F4USEMOD = '" & strUsuMod & "', "
        strSQLVale = strSQLVale & "F4SERDOC = '" & strSerieDocumento & "', "
        strSQLVale = strSQLVale & "F4OBSERVA = '" & strObservaciones & "', "
        strSQLVale = strSQLVale & "F4UUPP = '" & strUUPP & "', "
        strSQLVale = strSQLVale & "F4NUMVALENVIO = '" & strNumeroValeEnvio & "', "
        strSQLVale = strSQLVale & "NUMORDEN = '" & strNumeroOrdenCompra & "', "
        strSQLVale = strSQLVale & "CODALMDES = '" & strCodAlmacenDestino & "', "
        strSQLVale = strSQLVale & "F2CODUSER = '" & strCodSolicitante & "', "
        strSQLVale = strSQLVale & "NUMENSAM = '" & strNumeroValeExterno & "', "
        strSQLVale = strSQLVale & "SECCION = '" & strSeccion & "', "
        strSQLVale = strSQLVale & "LOTE = " & dblLote & ", "
        strSQLVale = strSQLVale & "F4REGCOM = " & IIf(strRegistroCompra = vbNullString, "NULL", "'" & strRegistroCompra & "'") & ", "
        strSQLVale = strSQLVale & "VALING = '" & strValeIngreso & "', "
        strSQLVale = strSQLVale & "EXPORTARVALE = " & IIf(bolExportarVale, "1", "0") & " "
        
        strSQLVale = strSQLVale & "WHERE "
        strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
    End If
    
    cnBdCPlus.Execute strSQLVale
    
    guardarVale = True
    
    strSQLSelectAlter = strSQLVale
    strSQLVale = vbNullString
    
    Exit Function
errGuardarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vuelva a intentarlo nuevamente.", _
            vbInformation + vbOKOnly, App.ProductName & " - SqlClsVale: GuardarVale"
    
    guardarVale = False
    
    Err.Clear
End Function

Public Sub guardarValeDetalleOneByOne()
    On Error GoTo errGuardarValeDetalleOneByOne
    
    strSQLVale = vbNullString
        
    strSQLVale = strSQLVale & "INSERT INTO PROCESOS.IF3VALES("
    strSQLVale = strSQLVale & "F2CODALM, F4NUMVAL, TIPO, "
    strSQLVale = strSQLVale & "F4FECVAL, F5CODPRO, F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "F3CANPRO, F3SALPEP, F3VALVTA, F3IGV, "
    strSQLVale = strSQLVale & "F3TOTITE, F3IGVDOL, F3VALDOL, "
    strSQLVale = strSQLVale & "F3TOTDOL, F3GRUPO, F3JCG, "
    strSQLVale = strSQLVale & "F3ITEM, F3PUNIT, SECPROCESO, "
    strSQLVale = strSQLVale & "F4NUMORD, COD_SOLICITUD, "
    strSQLVale = strSQLVale & "F3PORCENTAJEDSCTO, F3MONTODSCTO, "
    strSQLVale = strSQLVale & "OBSERVACIONES) "
    
    strSQLVale = strSQLVale & "VALUES("
    strSQLVale = strSQLVale & "'" & strCodAlmacen & "', '" & strNumeroVale & "', '" & strTipoVale & "', "
    strSQLVale = strSQLVale & "'" & StrFecha & "', '" & strCodProducto & "', '" & strCodProductoOriginal & "', "
    strSQLVale = strSQLVale & dblCantidad & ", " & dblCantidadMaxima & ", " & dblValorVta & ", " & dblIgv & ", "
    strSQLVale = strSQLVale & dblTotal & ", " & dblIgvDol & ", " & dblValorVtaDol & ", "
    strSQLVale = strSQLVale & dblTotalDol & ", '" & strGrupo & "', '*', "
    strSQLVale = strSQLVale & dblItem & ", 0, NULL, "
    strSQLVale = strSQLVale & "'" & strNumeroOrdenCompra & "', "
    strSQLVale = strSQLVale & "'" & strRequerimiento & "', "
    strSQLVale = strSQLVale & dblPorcentajeDscto & ", "
    strSQLVale = strSQLVale & dblMontoDscto & ", "
    strSQLVale = strSQLVale & IIf(strObservacionesPorItem <> vbNullString, "'" & strObservacionesPorItem & "'", "''") & ")"
    
    cnBdCPlus.Execute strSQLVale
    
    strSQLSelectAlter = strSQLVale
    strSQLVale = vbNullString
    
    Exit Sub
errGuardarValeDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: GuardarValeDetalleOneByOne"
    
    Err.Clear
End Sub

Public Function eliminarVale(Optional ByVal bolRegistroImportado As Boolean) As Boolean
    On Error GoTo errEliminarVale
    
'    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F2CODALM", "IF3ORDEN", strNumeroVale, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarVale = False
'
'        Exit Function
'    End If
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "DELETE "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES "
    strSQLVale = strSQLVale & "WHERE "
    
        'If Not bolRegistroImportado Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
        'Else
        '    strSQLVale = strSQLVale & "NUMENSAM = '" & strNumeroValeExterno & "' AND "
        '    strSQLVale = strSQLVale & "F4TIPOVALE = '" & strTipoVale & "'"
        'End If
    
    cnBdCPlus.Execute strSQLVale
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "DELETE "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    
        'If Not bolRegistroImportado Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
        'Else
        '    strSQLVale = strSQLVale & "NUMENSAM = '" & strNumeroValeExterno & "' AND "
        '    strSQLVale = strSQLVale & "F4TIPOVALE = '" & strTipoVale & "'"
        'End If
    
    cnBdCPlus.Execute strSQLVale
    
    eliminarVale = True
    
    strSQLSelectAlter = strSQLVale
    strSQLVale = vbNullString
    
    Exit Function
errEliminarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: EliminarVale"
    
    eliminarVale = False
    
    Err.Clear
End Function

Public Function verificarCierreVale() As Boolean
    On Error GoTo errVerificarCierreVale
    
    Dim cmdOP As ADODB.Command
    
    Set cmdOP = New ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_VerificarCierreVale"
        
        .Parameters.Append .CreateParameter("@CODIGOALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@NUMEROVALE", adVarChar, adParamInput, 20, strNumeroVale)
        .Parameters.Append .CreateParameter("@FECHAINICIAL", adDate, adParamInput, 10, strFechaInicioMes)
        .Parameters.Append .CreateParameter("@FECHAFINAL", adDate, adParamInput, 10, strFechaFinMes)
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        'rstVale.MoveFirst
        
        verificarCierreVale = True
        
        Do While Not rstVale.EOF
            If Not CBool(rstVale!VB1) Then
                verificarCierreVale = False
                
                Exit Do
            End If
            
            rstVale.MoveNext
        Loop
    Else
        verificarCierreVale = False
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    strSQLVale = vbNullString
    
    Set rstVale = Nothing
    
    Exit Function
errVerificarCierreVale:
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - SqlClsVale: VerificarCierreVale"
    
    Actualiza_Log "SqlClsVale: verificarCierreVale. / No. Error: " & Err.Number & " / " & "Descripción: " & Err.Description, StrConexDbBancos
    
    verificarCierreVale = False
    
    Err.Clear
End Function

Public Function cerrarVale() As Boolean
    On Error GoTo errCerrarVale
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "UPDATE "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "SET "
    strSQLVale = strSQLVale & "VB1 = " & IIf(bolVB1, "1", "0") & ", "
    strSQLVale = strSQLVale & "VB1USER = '" & strVB1Usuario & "', "
    strSQLVale = strSQLVale & "VB1FECHA = '" & strVB1Fecha & "' "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F1CODORI <> 'XCS' AND "
    strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        
        If strCodAlmacen <> vbNullString And strNumeroVale <> vbNullString Then
            strSQLVale = strSQLVale & "F4NUMVAL = '" & strNumeroVale & "'"
        Else
            strSQLVale = strSQLVale & "F4FECVAL BETWEEN "
            strSQLVale = strSQLVale & "'" & strFechaInicioMes & "' AND "
            strSQLVale = strSQLVale & "'" & strFechaFinMes & "'"
        End If
    
    cnBdCPlus.Execute strSQLVale
    
    cerrarVale = True
    
    strSQLSelectAlter = strSQLVale
    strSQLVale = vbNullString
    
    Exit Function
errCerrarVale:
    cerrarVale = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: CerrarVale"
    
    Err.Clear
End Function

Public Function generarCierreMensual(ByVal intAnno As Integer, _
                                    ByVal intMes As Integer, _
                                    ByVal strCodAlmacen As String, _
                                    Optional ByVal bolSoloReiniciarCierreMensual As Boolean) As Boolean
    
    On Error GoTo errGenerarCierreMensual
    
    'Limpiar Cierre Mes
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "DELETE FROM IF3CIERREMENSUAL WHERE ANNO = " & intAnno & " AND MES = " & intMes
    
    cnBdCPlus.Execute strSQLVale
    
    If bolSoloReiniciarCierreMensual Then
        generarCierreMensual = True
        
        Exit Function
    End If
    
    
    'Cierre Mes Stock Fisico
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "INSERT INTO IF3CIERREMENSUAL(ANNO, MES, TIPO, ALMACEN, CODPRODUCTO, COMPROMISO, CANTIDAD) "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & intAnno & " AS ANNO, "
    strSQLVale = strSQLVale & intMes & " AS MES, "
    strSQLVale = strSQLVale & "CIERREMES.TIPO, "
    strSQLVale = strSQLVale & "CIERREMES.ALMACEN, "
    strSQLVale = strSQLVale & "CIERREMES.CODPRODUCTO, "
    strSQLVale = strSQLVale & "CIERREMES.COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(CIERREMES.CANTIDAD) ) AS CANTIDAD "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "TIPO, "
    strSQLVale = strSQLVale & "CM.ALMACEN, "
    strSQLVale = strSQLVale & "CM.CODPRODUCTO, "
    strSQLVale = strSQLVale & "CM.COMPROMISO AS COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), CM.CANTIDAD ) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = " & IIf(intMes = 1, intAnno - 1, intAnno) & " AND "
    strSQLVale = strSQLVale & "CM.MES = " & IIf(intMes = 1, 12, intMes - 1) & " AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'F' AND "
    strSQLVale = strSQLVale & "CM.ALMACEN = '" & strCodAlmacen & "' "
    
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "'F' AS TIPO, "
    strSQLVale = strSQLVale & "DET.F2CODALM AS ALMACEN, "
    strSQLVale = strSQLVale & "DET.F5CODPRO AS CODPRODUCTO, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'S' THEN -1 ELSE 1) ) ) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN "
            strSQLVale = strSQLVale & "CVDATE('" & DateSerial(intAnno, intMes + 0, 1) & "') AND "
            strSQLVale = strSQLVale & "CVDATE('" & DateSerial(intAnno, intMes + 1, 0) & "') AND "
    strSQLVale = strSQLVale & "CAB.F2CODALM = '" & strCodAlmacen & "' "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS CIERREMES "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CIERREMES.TIPO, "
    strSQLVale = strSQLVale & "CIERREMES.ALMACEN, "
    strSQLVale = strSQLVale & "CIERREMES.CODPRODUCTO, "
    strSQLVale = strSQLVale & "CIERREMES.COMPROMISO "
    strSQLVale = strSQLVale & "HAVING "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(CIERREMES.CANTIDAD) ) <> 0"
    
    cnBdCPlus.Execute strSQLVale
    
    Actualiza_Log strSQLVale, strCadenaConexioBdCPlus
    
    'Cierre Mes Stock Virtual
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "INSERT INTO IF3CIERREMENSUAL(ANNO, MES, TIPO, ALMACEN, CODPRODUCTO, COMPROMISO, CANTIDAD) "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & intAnno & " AS ANNO, "
    strSQLVale = strSQLVale & intMes & " AS MES, "
    strSQLVale = strSQLVale & "CIERREMES.TIPO, "
    strSQLVale = strSQLVale & "CIERREMES.ALMACEN, "
    strSQLVale = strSQLVale & "CIERREMES.CODPRODUCTO, "
    strSQLVale = strSQLVale & "CIERREMES.COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(CIERREMES.CANTIDAD) ) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "CM.TIPO, "
    strSQLVale = strSQLVale & "CM.ALMACEN, "
    strSQLVale = strSQLVale & "CM.CODPRODUCTO, "
    strSQLVale = strSQLVale & "TRIM(CM.COMPROMISO & '') AS COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), CM.CANTIDAD ) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = " & IIf(intMes = 1, intAnno - 1, intAnno) & " AND "
    strSQLVale = strSQLVale & "CM.MES = " & IIf(intMes = 1, 12, intMes - 1) & " AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'V' "
    
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "'V' AS TIPO, "
    strSQLVale = strSQLVale & "DET.F4NUMORD AS ALMACEN, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL AS CODPRODUCTO, "
    strSQLVale = strSQLVale & "TRIM(DET.COD_SOLICITUD & '') AS COMPROMISO, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'S' THEN -1 ELSE 1)) ) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN "
    strSQLVale = strSQLVale & "CVDATE('" & DateSerial(intAnno, intMes + 0, 1) & "') AND "
    strSQLVale = strSQLVale & "CVDATE('" & DateSerial(intAnno, intMes + 1, 0) & "') AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS CIERREMES "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CIERREMES.TIPO, "
    strSQLVale = strSQLVale & "CIERREMES.ALMACEN, "
    strSQLVale = strSQLVale & "CIERREMES.CODPRODUCTO, "
    strSQLVale = strSQLVale & "CIERREMES.COMPROMISO "
    strSQLVale = strSQLVale & "HAVING "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10,2), SUM(CIERREMES.CANTIDAD) ) <> 0"
    
    cnBdCPlus.Execute strSQLVale
    
    Actualiza_Log strSQLVale, strCadenaConexioBdCPlus
    
    generarCierreMensual = True
    
    strSQLVale = vbNullString
    
    Exit Function
errGenerarCierreMensual:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: GenerarCierreMensual"
    
    generarCierreMensual = False
    
    Err.Clear
End Function



'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub listarAnnoValeSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarAnnoValeSoloSeleccion
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "YEAR(F4FECVAL) AS ANNO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F1CODORI NOT IN ('XCS') "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "YEAR(F4FECVAL) "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "YEAR(F4FECVAL)"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        comboList.Clear
        
        Do While Not rstVale.EOF
            comboList.AddItem Trim(rstVale!Anno & "")
            
            rstVale.MoveNext
        Loop
            'comboList.ListIndex = 0
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    strSQLVale = vbNullString
    
    Set rstVale = Nothing
    
    Exit Sub
    Resume
errListarAnnoValeSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ListarAnnoValeSoloSeleccion"
    
    Err.Clear
End Sub

Public Sub listarMesValeSoloSeleccion(ByVal comboList As ComboBox, ByVal strAnno As String)
    On Error GoTo errListarMesValeSoloSeleccion
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "MONTH(F4FECVAL) AS MES "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF4VALES "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "F1CODORI NOT IN ('XCS') AND "
    strSQLVale = strSQLVale & "YEAR(F4FECVAL) = " & strAnno & " "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "MONTH(F4FECVAL) "
    strSQLVale = strSQLVale & "ORDER BY MONTH(F4FECVAL)"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        comboList.Clear
        
        Do While Not rstVale.EOF
            comboList.AddItem UCase(Format("01/" & Format(Trim(rstVale!mes & ""), "00") & "/" & strAnno, "MMMM")) & Space(50) & Trim(rstVale!mes & "")
            
            rstVale.MoveNext
        Loop
            'comboList.ListIndex = 0
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    strSQLVale = vbNullString
    
    Set rstVale = Nothing
    
    Exit Sub
errListarMesValeSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ListarMesValeSoloSeleccion"
    
    Err.Clear
End Sub

'Listar Movimiento de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaMovimientoProductoResumen(ByVal grilla As dxDBGrid, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strCodFamilia As String, _
                                                    ByVal strCodSubFamilia As String, _
                                                    ByVal strFiltroSensitivo As String, _
                                                    ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaMovimientoProductoResumen
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnn_dbbancos
        .CommandType = adCmdStoredProc
        '.CommandTimeout = "180"
        .CommandText = "Procesos.usp_ListarResumenStock"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@CODFAMILIA", adVarChar, adParamInput, 10, strCodFamilia)
        .Parameters.Append .CreateParameter("@CODSUBFAMILIA", adVarChar, adParamInput, 10, strCodSubFamilia)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strFiltroSensitivo)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY F5NOMPRO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close

            .Columns.DestroyColumns
        End With

        Dim gColumn As dxGridColumn

        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With

            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With

            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With

            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Stock Comprometido en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKCOMPROMETIDO"
                .Font.Charset = 0
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoEnAlmacen"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Comprometido por llegar
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARCOMP"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoPorLlegar"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKLIBRE"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibreEnAlmacen"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARLIBRE"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibrePorLlegar"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Total en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalEnAlmacen"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Total por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalPorLlegar"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Sctock Total
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Stock Total"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCK"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStock"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoResumen:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoResumen"
    End Select
    
    Err.Clear
End Sub

'Listar Movimiento de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaMovimientoProductoResumenV2(ByVal grilla As dxDBGrid, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strFiltroSensitivo As String, _
                                                    Optional ByVal strCodFamilia As String, _
                                                    Optional ByVal strCodSubFamilia As String, _
                                                    Optional ByVal strUltimoPeriodoCierreMes As String, _
                                                    Optional ByVal strFechaCorte As String, _
                                                    Optional ByVal listaImagenes As Object)
        
    On Error GoTo errListarGrillaMovimientoProductoResumenV2
    
    Dim strAnno As String, strMes As String, strFechaInicial As String
    
    strAnno = Mid(strUltimoPeriodoCierreMes, 1, Len(strUltimoPeriodoCierreMes) - 2)
    strMes = Mid(strUltimoPeriodoCierreMes, Len(strUltimoPeriodoCierreMes) - 1)
    
    strFechaInicial = "01/" & IIf(Val(strMes) < 12, Format(Val(strMes + 1), "00"), "01") & "/" & IIf(Val(strMes) < 12, strAnno, Val(strAnno) + 1)
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
    strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
    strSQLVale = strSQLVale & "PROD.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) AS STOCKCOMPROMETIDO, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) AS STOCKPORLLEGARCOMP, "
    strSQLVale = strSQLVale & "VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) AS STOCKLIBRE, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCKPORLLEGARLIBRE, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) + VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) AS STOCKTOTAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCKPORLLEGAR, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) + VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) + VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCK "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "((((((MAESTROS.IF5PLA AS PROD "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01) "
    
    'RESUMEN DE STOCK COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "RESUMEN.F2CODALM,  "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPRO, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(RESUMEN.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "CM.ALMACEN AS F2CODALM, CM.CODPRODUCTO AS F5CODPRO, VAL(FORMAT( SUM(CM.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = '" & strAnno & "' AND "
    strSQLVale = strSQLVale & "CM.MES = '" & strMes & "' AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'F' AND "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "CM.ALMACEN = '" & strCodAlmacen & "' AND "
        End If
        
        strSQLVale = strSQLVale & "TRIM(CM.COMPROMISO & '') <> '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CM.ALMACEN, CM.CODPRODUCTO "
        
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "CAB.F2CODALM, DET.F5CODPRO, VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & strFechaCorte & "') AND "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "CAB.F2CODALM = '" & strCodAlmacen & "' AND "
        End If
        
        strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CAB.F2CODALM, DET.F5CODPRO) AS RESUMEN "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "RESUMEN.F2CODALM, "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPRO"
    
    strSQLVale = strSQLVale & ") AS COMPROMISO ON COMPROMISO.F5CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "RESUMEN.F2CODALM,  "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPRO, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(RESUMEN.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "CM.ALMACEN AS F2CODALM, CM.CODPRODUCTO AS F5CODPRO, VAL(FORMAT( SUM(CM.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = '" & strAnno & "' AND "
    strSQLVale = strSQLVale & "CM.MES = '" & strMes & "' AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'F' AND "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "CM.ALMACEN = '" & strCodAlmacen & "' AND "
        End If
        
        strSQLVale = strSQLVale & "TRIM(CM.COMPROMISO & '') = '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CM.ALMACEN, CM.CODPRODUCTO "
        
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "CAB.F2CODALM, DET.F5CODPRO, VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & strFechaCorte & "') AND  "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "CAB.F2CODALM = '" & strCodAlmacen & "' AND "
        End If
        
        strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CAB.F2CODALM, DET.F5CODPRO) AS RESUMEN "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "RESUMEN.F2CODALM,  "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPRO"
    
    strSQLVale = strSQLVale & ") AS LIBRE ON LIBRE.F5CODPRO = PROD.F5CODPRO) "
    
    
    
    'RESUMEN DE STOCK POR LLEGAR COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "((((IF3ORDEN AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "RESUMEN.F4NUMORD,  "
    strSQLVale = strSQLVale & "RESUMEN.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(RESUMEN.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "CM.ALMACEN AS F4NUMORD, TRIM(CM.COMPROMISO & '') AS COD_SOLICITUD, CM.CODPRODUCTO AS F5CODPROORIGINAL, VAL(FORMAT( SUM(CM.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = '" & strAnno & "' AND "
    strSQLVale = strSQLVale & "CM.MES = '" & strMes & "' AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'V' AND "
        
        strSQLVale = strSQLVale & "TRIM(CM.COMPROMISO & '') <> '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CM.ALMACEN, TRIM(CM.COMPROMISO & ''), CM.CODPRODUCTO "
        
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F4NUMORD, DET.COD_SOLICITUD, DET.F5CODPROORIGINAL, VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & strFechaCorte & "') AND  "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
    
        strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F4NUMORD, DET.COD_SOLICITUD, DET.F5CODPROORIGINAL) AS RESUMEN "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "RESUMEN.F4NUMORD,  "
    strSQLVale = strSQLVale & "RESUMEN.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPROORIGINAL"
    
'    strSQLVale = strSQLVale & "SELECT "
'    strSQLVale = strSQLVale & "DET.F4NUMORD, "
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
'    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
'    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
'    strSQLVale = strSQLVale & "FROM "
'    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
'    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'    strSQLVale = strSQLVale & "WHERE "
'    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
'    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
'    'strSQLVale = strSQLVale & "CVDATE(DET.F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
'
'
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' "
'
'    strSQLVale = strSQLVale & "GROUP BY "
'    strSQLVale = strSQLVale & "DET.F4NUMORD, "
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
'    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
    
    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
        
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARCOMP ON PORLLEGARCOMP.F3CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK POR LLEGAR LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "((((IF3ORDEN AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN IF4ORDEN AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "RESUMEN.F4NUMORD,  "
    strSQLVale = strSQLVale & "RESUMEN.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(RESUMEN.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "CM.ALMACEN AS F4NUMORD, TRIM(CM.COMPROMISO & '') AS COD_SOLICITUD, CM.CODPRODUCTO AS F5CODPROORIGINAL, VAL(FORMAT( SUM(CM.CANTIDAD) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL AS CM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CM.ANNO = '" & strAnno & "' AND "
    strSQLVale = strSQLVale & "CM.MES = '" & strMes & "' AND "
    strSQLVale = strSQLVale & "CM.TIPO = 'V' AND "
        
        strSQLVale = strSQLVale & "TRIM(CM.COMPROMISO & '') = '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CM.ALMACEN, TRIM(CM.COMPROMISO & ''), CM.CODPRODUCTO "
        
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F4NUMORD, DET.COD_SOLICITUD, DET.F5CODPROORIGINAL, VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & strFechaCorte & "') AND  "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
    
        strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F4NUMORD, DET.COD_SOLICITUD, DET.F5CODPROORIGINAL) AS RESUMEN "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "RESUMEN.F4NUMORD,  "
    strSQLVale = strSQLVale & "RESUMEN.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "RESUMEN.F5CODPROORIGINAL"
    
    
    
    
'    strSQLVale = strSQLVale & "SELECT "
'    strSQLVale = strSQLVale & "DET.F4NUMORD, "
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
'    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
'    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
'    strSQLVale = strSQLVale & "FROM "
'    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
'    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'    strSQLVale = strSQLVale & "WHERE "
'    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
'    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
'    'strSQLVale = strSQLVale & "CVDATE(DET.F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
'
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' "
'
'    strSQLVale = strSQLVale & "GROUP BY "
'    strSQLVale = strSQLVale & "DET.F4NUMORD, "
'    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
'    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
    
    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    'strSQLVale = strSQLVale & "CVDATE(CAB.F4FECEMI) <= CVDATE('" & strFechaCorte & "') AND "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARLIBRE ON PORLLEGARLIBRE.F3CODPRO = PROD.F5CODPRO "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "PROD.F5CODPRO <> '' AND "
    strSQLVale = strSQLVale & "PROD.F5DESCONTINUADO = 'N' AND "
    strSQLVale = strSQLVale & "PROD.TIENEMOVENALM = 1 "
        
        If strCodFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND FAM.F7CODCON = '" & strCodFamilia & "' "
        End If
        
        If strCodSubFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND SFAM.F7CODCON = '" & strCodSubFamilia & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND ("
            strSQLVale = strSQLVale & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' "
            strSQLVale = strSQLVale & ") "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
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
            
            'Columna Stock Comprometido en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKCOMPROMETIDO"
                .Font.Charset = 0
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Comprometido por llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARCOMP"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKLIBRE"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibreEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARLIBRE"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibrePorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Sctock Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Stock Total"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCK"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStock"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoResumenV2:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoResumenV2"
    End Select
    
    Err.Clear
End Sub

'Listar Movimiento de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaMovimientoProductoResumenV3(ByVal grilla As dxDBGrid, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strFiltroSensitivo As String, _
                                                    Optional ByVal strCodFamilia As String, _
                                                    Optional ByVal strCodSubFamilia As String, _
                                                    Optional ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaMovimientoProductoResumenV3
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROD.FAMILIA, "
    strSQLVale = strSQLVale & "PROD.SUBFAMILIA, "
    strSQLVale = strSQLVale & "PROD.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "PROD.F7SIGMED, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.STOCK, '#0.00')) AS STOCKCOMPROMETIDO, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.00')) AS STOCKPORLLEGARCOMP, "
    strSQLVale = strSQLVale & "VAL(FORMAT(LIBRE.STOCK, '#0.00')) AS STOCKLIBRE, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.00')) AS STOCKPORLLEGARLIBRE, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.STOCK, '#0.00')) + VAL(FORMAT(LIBRE.STOCK, '#0.00')) AS STOCKTOTAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.00')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.00')) AS STOCKPORLLEGAR, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.STOCK, '#0.00')) + VAL(FORMAT(LIBRE.STOCK, '#0.00')) + VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.00')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.00')) AS STOCK "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(((TMPUTILSTOCKPRODUCTO AS PROD "
    
    'RESUMEN DE STOCK COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F5CODPRO, VAL(FORMAT( SUM(CANTIDAD) , '#.00')) AS STOCK "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKRESUMEN "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TIPO = 'CEA' AND "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        End If
        
    strSQLVale = strSQLVale & "COD_SOLICITUD <> '' "
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F5CODPRO) AS COMPROMISO ON COMPROMISO.F5CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F5CODPRO, VAL(FORMAT( SUM(CANTIDAD) , '#.00')) AS STOCK "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKRESUMEN "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TIPO = 'LEA' AND "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        End If
    
    strSQLVale = strSQLVale & "COD_SOLICITUD = '' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F5CODPRO) AS LIBRE ON LIBRE.F5CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK POR LLEGAR COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKORDENDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F2CODALM, COD_SOLICITUD, F5CODPRO, VAL(FORMAT( SUM(CANTIDAD) , '#.00')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKRESUMEN "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TIPO = 'CPL' AND "
    strSQLVale = strSQLVale & "COD_SOLICITUD <> '' "

    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F2CODALM, COD_SOLICITUD, F5CODPRO"

    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F2CODALM = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPRO = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARCOMP ON PORLLEGARCOMP.F3CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK POR LLEGAR LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKORDENDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F2CODALM, COD_SOLICITUD, F5CODPRO, VAL(FORMAT( SUM(CANTIDAD) , '#.00')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKRESUMEN "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TIPO = 'LPL' AND "
    strSQLVale = strSQLVale & "COD_SOLICITUD = '' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F2CODALM, COD_SOLICITUD, F5CODPRO"

    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F2CODALM = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPRO = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARLIBRE ON PORLLEGARLIBRE.F3CODPRO = PROD.F5CODPRO "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "PROD.F5CODPRO <> '' "
        
        If strCodFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND PROD.CODFAMILIA = '" & strCodFamilia & "' "
        End If
        
        If strCodSubFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND PROD.CODSUBFAMILIA = '" & strCodSubFamilia & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND ("
            strSQLVale = strSQLVale & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' "
            strSQLVale = strSQLVale & ") "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO"
    
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
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
            
            'Columna Stock Comprometido en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKCOMPROMETIDO"
                .Font.Charset = 0
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Comprometido por llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARCOMP"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKLIBRE"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibreEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARLIBRE"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibrePorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Sctock Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Stock Total"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCK"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStock"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoResumenV3:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoResumenV3"
    End Select
    
    Err.Clear
End Sub

'Listar Movimiento de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaMovimientoProductoResumenV4(ByVal grilla As dxDBGrid, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strFiltroSensitivo As String, _
                                                    Optional ByVal strNroPedido As String, _
                                                    Optional ByVal strCodFamilia As String, _
                                                    Optional ByVal strCodSubFamilia As String, _
                                                    Optional ByVal listaImagenes As Object)
        
    On Error GoTo errListarGrillaMovimientoProductoResumenV4
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROD.FAMILIA, "
    strSQLVale = strSQLVale & "PROD.SUBFAMILIA, "
    strSQLVale = strSQLVale & "PROD.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "PROD.F7SIGMED, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) AS STOCKCOMPROMETIDO, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) AS STOCKPORLLEGARCOMP, "
    strSQLVale = strSQLVale & "VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) AS STOCKLIBRE, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCKPORLLEGARLIBRE, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) + VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) AS STOCKTOTAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCKPORLLEGAR, "
    
    strSQLVale = strSQLVale & "VAL(FORMAT(COMPROMISO.CANTIDAD, '#0.0000')) + VAL(FORMAT(LIBRE.CANTIDAD, '#0.0000')) + VAL(FORMAT(PORLLEGARCOMP.SALDO, '#0.0000')) + VAL(FORMAT(PORLLEGARLIBRE.SALDO, '#0.0000')) AS STOCK "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "(((TMPUTILSTOCKPRODUCTO AS PROD "
    
    'RESUMEN DE STOCK COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F5CODPRO, F2CODALM, VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKVALEDET "
    strSQLVale = strSQLVale & "WHERE "
    'strSQLVale = strSQLVale & "CVDATE(F4FECVAL) <= CVDATE('" & strFechaCorte & "') "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        End If
        
        'ADD
        If strNroPedido <> vbNullString Then
            strSQLVale = strSQLVale & "COD_SOLICITUD = '" & strNroPedido & "' "
        Else
            strSQLVale = strSQLVale & "COD_SOLICITUD <> '' "
        End If
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F5CODPRO, F2CODALM) AS COMPROMISO ON COMPROMISO.F5CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "F5CODPRO, F2CODALM, VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKVALEDET "
    strSQLVale = strSQLVale & "WHERE "
    'strSQLVale = strSQLVale & "CVDATE(F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "F2CODALM = '" & strCodAlmacen & "' AND "
        End If
        
    strSQLVale = strSQLVale & "COD_SOLICITUD = '' "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "F5CODPRO, F2CODALM) AS LIBRE ON LIBRE.F5CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK POR LLEGAR COMPROMETIDO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKORDENDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKVALEDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILSTOCKVALECAB AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
    'strSQLVale = strSQLVale & "CVDATE(DET.F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
        
        'ADD
        If strNroPedido <> vbNullString Then
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' "
        Else
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' "
        End If
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL" ') AS INGRESOS "
    
    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    'strSQLVale = strSQLVale & "CVDATE(CAB.F4FECEMI) <= CVDATE('" & strFechaCorte & "') AND "
        
        'ADD
        If strNroPedido <> vbNullString Then
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' AND "
        Else
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
        End If
        
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARCOMP ON PORLLEGARCOMP.F3CODPRO = PROD.F5CODPRO) "
    
    'RESUMEN DE STOCK POR LLEGAR LIBRE
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F3CODPRO, "
    strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKORDENDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) , '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "TMPUTILSTOCKVALEDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILSTOCKVALECAB AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
    strSQLVale = strSQLVale & "WHERE "
    'strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
    strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
    'strSQLVale = strSQLVale & "CVDATE(DET.F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
    
    strSQLVale = strSQLVale & ") AS INGRESOS "
    strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
    'strSQLVale = strSQLVale & "CVDATE(CAB.F4FECEMI) <= CVDATE('" & strFechaCorte & "') AND "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
    strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(DET.FACTORCONVERSION), 1, DET.FACTORCONVERSION)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F3CODPRO) AS PORLLEGARLIBRE ON PORLLEGARLIBRE.F3CODPRO = PROD.F5CODPRO "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "PROD.F5CODPRO <> '' "
        
        If strCodFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND PROD.CODFAMILIA = '" & strCodFamilia & "' "
        End If
        
        If strCodSubFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND PROD.CODSUBFAMILIA = '" & strCodSubFamilia & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND ("
            'strSQLVale = strSQLVale & "PROD.F5CODPRO LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLVale = strSQLVale & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' "
            'strSQLVale = strSQLVale & "MED.F7SIGMED  LIKE '%" & strFiltroSensitivo & "%'"
            strSQLVale = strSQLVale & ") "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
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
            
            'Columna Stock Comprometido en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKCOMPROMETIDO"
                .Font.Charset = 0
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Comprometido por llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARCOMP"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockComprometidoPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKLIBRE"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibreEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARLIBRE"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockLibrePorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotalPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Sctock Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Stock Total"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCK"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStock"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoResumenV4:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoResumenV4"
    End Select
    
    Err.Clear
End Sub

'Listar Detalle de Movimiento de Producto en Grilla (QuamtumGrid)
Public Sub listarGrillaMovimientoProductoDetalle(ByVal grilla As dxDBGrid, _
                                                    ByVal listaImagenes As Object, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strCodProducto As String, _
                                                    ByVal strIngreseForV As String, _
                                                    ByVal strIngreseCorL As String, _
                                                    Optional strNroPedido As String, _
                                                    Optional bolVerOpcionDescarga As Boolean)
    
    On Error GoTo errListarGrillaMovimientoProductoDetalle
    
    strSQLVale = vbNullString
    
    strSQLVale = strSQLVale & "INSERT INTO TMPUTILSTOCKDETALLE("
    strSQLVale = strSQLVale & "CODPRODUCTO, "
    strSQLVale = strSQLVale & "NOMPRODUCTO, "
    strSQLVale = strSQLVale & "CODALMACEN, "
    strSQLVale = strSQLVale & "NOMALMACEN, "
    
    strSQLVale = strSQLVale & "NROOC, "
    
    strSQLVale = strSQLVale & "NROPEDIDO, "
    strSQLVale = strSQLVale & "UM, "
    strSQLVale = strSQLVale & "OBSERVACIONES, "
    strSQLVale = strSQLVale & "FECHAENTREGA, "
    strSQLVale = strSQLVale & "CANTIDAD, "
    strSQLVale = strSQLVale & "NROPEDIDODESTINO, "
    strSQLVale = strSQLVale & "CANTIDADDESTINO, "
    strSQLVale = strSQLVale & "PROCESAR) "
    
    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    Select Case strIngreseForV
        Case "F" 'Stock Fisico (Stock en Almacen: Comprometido o Libre)
            strSQLVale = strSQLVale & "SELECT "
            strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
            strSQLVale = strSQLVale & "(PROD.F5NOMPRO & ' (' & MED.F7NOMMED & ')') AS F5NOMPRO, "
            strSQLVale = strSQLVale & "DISPONIBLE.F2CODALM, "
            strSQLVale = strSQLVale & "ALM.F2NOMALM, "
            
            strSQLVale = strSQLVale & "DISPONIBLE.OC, "
            strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO, "
            strSQLVale = strSQLVale & "MED.F7SIGMED, "
                
                Select Case strIngreseCorL
                    Case "C"
                        strSQLVale = strSQLVale & "CABPED.CS_OBSERVACIONES, "
                        strSQLVale = strSQLVale & "CABPED.CS_FENTREGA, "
                    Case "L"
                        'strSQLVale = strSQLVale & "'--' AS PEDIDO, "
                        strSQLVale = strSQLVale & "'< Stock Libre de Compromiso >' AS CS_OBSERVACIONES, "
                        strSQLVale = strSQLVale & "DATE() AS CS_FENTREGA, "
                End Select
                
            strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) AS CANTIDAD, "
            
            strSQLVale = strSQLVale & "NULL AS NROPEDIDODESTINO, "
            strSQLVale = strSQLVale & "0 AS CANTIDADDESTINO, "
            strSQLVale = strSQLVale & "FALSE AS PROCESAR "
            
            strSQLVale = strSQLVale & "FROM "
            
            strSQLVale = strSQLVale & "("
            
            strSQLVale = strSQLVale & "((("
            
            'SCTOCK DISPONIBLE DEL PRODUCTO
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F5CODPRO, "
            strSQLVale = strSQLVale & "DET.F2CODALM, "
            
            strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
            
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
            strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            strSQLVale = strSQLVale & "WHERE "
            'strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XJ0', 'XC0', 'XCS') AND " 'AJUSTES DE INGRESO, COMPRAS Y COMPROMISOS (REDISTRIBUCION DE STOCK)
            strSQLVale = strSQLVale & "DET.TIPO = 'I' AND "
            
            'strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD " & IIf(strIngreseCorL = "C", "<>", "=") & " '' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F5CODPRO, "
            strSQLVale = strSQLVale & "DET.F2CODALM, "
            
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS DISPONIBLE "
            
            'SCTOCK CONSUMIDO DEL PRODUCTO
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F5CODPRO, "
            strSQLVale = strSQLVale & "DET.F2CODALM, "
            
            strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
            
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
            strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.TIPO = 'S' AND "
            'strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD " & IIf(strIngreseCorL = "C", "<>", "=") & " '' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F5CODPRO, "
            strSQLVale = strSQLVale & "DET.F2CODALM, "
            
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS CONSUMIDO "
            strSQLVale = strSQLVale & "ON "
            strSQLVale = strSQLVale & "CONSUMIDO.F5CODPRO = DISPONIBLE.F5CODPRO AND "
            strSQLVale = strSQLVale & "CONSUMIDO.F2CODALM = DISPONIBLE.F2CODALM AND "
            strSQLVale = strSQLVale & "CONSUMIDO.PEDIDO = DISPONIBLE.PEDIDO AND "
            strSQLVale = strSQLVale & "CONSUMIDO.OC = DISPONIBLE.OC) "
            
            strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DISPONIBLE.F5CODPRO) "
            
            strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
            
            strSQLVale = strSQLVale & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DISPONIBLE.F2CODALM) "
            strSQLVale = strSQLVale & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DISPONIBLE.PEDIDO "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) > 0 "
                
                
                Select Case strIngreseCorL
                    Case "C"
                        If strNroPedido <> vbNullString Then
                            strSQLVale = strSQLVale & "AND DISPONIBLE.PEDIDO = '" & strNroPedido & "' "
                        End If
                End Select
                
            strSQLVale = strSQLVale & "ORDER BY "
            strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
            strSQLVale = strSQLVale & "DISPONIBLE.F2CODALM, "
            strSQLVale = strSQLVale & "DISPONIBLE.OC, "
            strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO"
        Case "V" 'Stock Virtual (Stock por Llegar: Comprometido o Libre)
            strSQLVale = strSQLVale & "SELECT "
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "(PROD.F5NOMPRO & ' (' & MED.F7NOMMED & ')') AS F5NOMPRO, "
            strSQLVale = strSQLVale & "'' AS CODALMACEN, "
            strSQLVale = strSQLVale & "'< Almacen por Definir a su llegada >' AS NOMALMACEN, "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "MED.F7SIGMED, "
            
                Select Case strIngreseCorL
                    Case "C"
                        strSQLVale = strSQLVale & "CABPED.CS_OBSERVACIONES, "
                        strSQLVale = strSQLVale & "CABPED.CS_FENTREGA, "
                    Case "L"
                        'strSQLVale = strSQLVale & "'--' AS PEDIDO, "
                        strSQLVale = strSQLVale & "'< Stock Libre de Compromiso >' AS CS_OBSERVACIONES, "
                        strSQLVale = strSQLVale & "DATE() AS CS_FENTREGA, "
                End Select
            
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO, "
            
            strSQLVale = strSQLVale & "NULL AS NROPEDIDODESTINO, "
            strSQLVale = strSQLVale & "0 AS CANTIDADDESTINO, "
            strSQLVale = strSQLVale & "FALSE AS PROCESAR "
            
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "("
            strSQLVale = strSQLVale & "(((IF3ORDEN AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
            strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
            strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
            strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "CAB.F1CODORI = 'XC0' AND "
            
            strSQLVale = strSQLVale & "DET.TIPO = 'I' AND "
            'strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
            
            strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD " & IIf(strIngreseCorL = "C", "<>", "=") & " '' "
            
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
            
            strSQLVale = strSQLVale & ") AS INGRESOS "
            strSQLVale = strSQLVale & "ON "
            strSQLVale = strSQLVale & "INGRESOS.F4NUMORD = DET.F4NUMORD AND "
            strSQLVale = strSQLVale & "INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND "
            strSQLVale = strSQLVale & "INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO) "
            
            strSQLVale = strSQLVale & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DET.COD_SOLICITUD "
            
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
            strSQLVale = strSQLVale & "DET.F3CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD " & IIf(strIngreseCorL = "C", "<>", "=") & " '' AND "
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
            strSQLVale = strSQLVale & "ORDER BY "
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD"
    End Select
    
    If Not grilla Is Nothing Then
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKDETALLE"
        
        abrirCnTemporal
        
        cnBdCPlus.Execute strSQLVale
        
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Codigo de Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Cod. Almacen"
                .DisableEditor = True
                .FieldName = "CODALMACEN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Visible = False
            End With
            
            'Columna Nombre del Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Almacen"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMALMACEN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Observaciones de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Observaciones"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "OBSERVACIONES"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservaciones"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 280
            End With
            
            'Columna Fecha Entrega de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fec. Entrega"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "FECHAENTREGA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaEntrega"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                
                Select Case strIngreseCorL
                    Case "C"
                        .Visible = True
                    Case "L"
                        .Visible = False
                End Select
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
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                '.Color = &HC0&
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Numero de Pedido de Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "No. Pedido Destino"
                '.Color = &H80FFFF
                .DecimalPlaces = 2
                .FieldName = "NROPEDIDODESTINO"
                .Font.Charset = 0
                '.Font.Bold = False
                '.FontColor = &H80000012
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColNroPedidoDestino"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Cantidad"
                '.Color = &HC0&
                .DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstCount
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 2
                .Caption = "Procesar"
                '.Color = &HC000&
                '.DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            'If bolVerOpcionDescarga Then
                'Columna Descargar
                Set gColumn = .Columns.Add(gedCheckEdit)
                
                With gColumn
                    .Alignment = taCenter
                    .BandIndex = 2
                    .Caption = "Descargar"
                    '.Color = &HC000&
                    '.DecimalPlaces = 2
                    '.DisableEditor = True
                    .FieldName = "DESCARGAR"
                    .Font.Charset = 0
                    '.Font.Bold = True
                    '.FontColor = &HFFFFFF
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColDescargar"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 70
                    .Visible = bolVerOpcionDescarga
                End With
            'End If
            
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILSTOCKDETALLE ORDER BY NROPEDIDO"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "NROPEDIDO"
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").GroupIndex = 0
            .Columns.ColumnByFieldName("NOMALMACEN").GroupIndex = 1
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoDetalle:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoDetalle"
    End Select
    
    Err.Clear
End Sub

'Listar Detalle de Movimiento de Producto en Grilla (QuamtumGrid) Version 2.0
Public Sub listarGrillaMovimientoProductoDetalleV2(ByVal grilla As dxDBGrid, _
                                                    ByVal listaImagenes As Object, _
                                                    ByVal strIngreseForV As String, _
                                                    ByVal strIngreseCorL As String, _
                                                    ByVal strNombreTablaSQLParaResumen As String, _
                                                    Optional strNroPedido As String, _
                                                    Optional bolVerOpcionDescarga As Boolean)
    
    On Error GoTo errListarGrillaMovimientoProductoDetalleV2
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarDetalleStock"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@TIPOSTOCK", adVarChar, adParamInput, 1, strIngreseForV)
        .Parameters.Append .CreateParameter("@ESTADOSTOCK", adVarChar, adParamInput, 1, strIngreseCorL)
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strNroPedido)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "LLAVE, "
    strSQLVale = strSQLVale & "CODPRODUCTO, "
    strSQLVale = strSQLVale & "NOMPRODUCTO, "
    strSQLVale = strSQLVale & "CODALMACEN, "
    strSQLVale = strSQLVale & "NOMALMACEN, "
    strSQLVale = strSQLVale & "NROOC, "
    strSQLVale = strSQLVale & "NROPEDIDO, "
    strSQLVale = strSQLVale & "UM, "
    strSQLVale = strSQLVale & "OBSERVACIONES, "
    strSQLVale = strSQLVale & "CONVERT(NVARCHAR(10), FECHAENTREGA,103) AS FECHAENTREGA, "
    strSQLVale = strSQLVale & "CANTIDAD, "
    strSQLVale = strSQLVale & "NROPEDIDODESTINO, "
    strSQLVale = strSQLVale & "CANTIDADDESTINO, "
    strSQLVale = strSQLVale & "PROCESAR, "
    strSQLVale = strSQLVale & "DESCARGAR "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & strNombreTablaSQLParaResumen & " "
    strSQLVale = strSQLVale & "ORDER BY "
    'strSQLVale = strSQLVale & "NROPEDIDO, NROOC"
    strSQLVale = strSQLVale & "FECHAENTREGA"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Codigo de Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Cod. Almacen"
                .DisableEditor = True
                .FieldName = "CODALMACEN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Visible = False
            End With
            
            'Columna Nombre del Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Almacen"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMALMACEN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Observaciones de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Observaciones"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "OBSERVACIONES"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservaciones"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 280
            End With
            
            'Columna Fecha Entrega de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fec. Entrega"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "FECHAENTREGA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaEntrega"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                
                Select Case strIngreseCorL
                    Case "C"
                        .Visible = True
                    Case "L"
                        .Visible = False
                End Select
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
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = "#,0.0000"
                .Width = 70
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                '.Color = &HC0&
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Numero de Pedido de Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "No. Pedido Destino"
                '.Color = &H80FFFF
                .DecimalPlaces = 2
                .FieldName = "NROPEDIDODESTINO"
                .Font.Charset = 0
                '.Font.Bold = False
                '.FontColor = &H80000012
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColNroPedidoDestino"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Cantidad"
                '.Color = &HC0&
                .DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstCount
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 2
                .Caption = "Procesar"
                '.Color = &HC000&
                '.DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            'If bolVerOpcionDescarga Then
                'Columna Descargar
                Set gColumn = .Columns.Add(gedCheckEdit)
                
                With gColumn
                    .Alignment = taCenter
                    .BandIndex = 2
                    .Caption = "Descargar"
                    '.Color = &HC000&
                    '.DecimalPlaces = 2
                    '.DisableEditor = True
                    .FieldName = "DESCARGAR"
                    .Font.Charset = 0
                    '.Font.Bold = True
                    '.FontColor = &HFFFFFF
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColDescargar"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 70
                    .Visible = bolVerOpcionDescarga
                End With
            'End If
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").GroupIndex = 0
            .Columns.ColumnByFieldName("NOMALMACEN").GroupIndex = 1
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaMovimientoProductoDetalleV2:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaMovimientoProductoDetalleV2"
    End Select
    
    Err.Clear
End Sub

'Listar Ingresos Faltantes por Compra en Stock (Comprometido y Libre) en Grilla (QuamtumGrid)
Public Sub listarGrillaFaltantePorOC(ByVal grilla As dxDBGrid, _
                                        ByVal listaImagenes As Object, _
                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaFaltantePorOC
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarFaltantePorOC"
        
        .Parameters.Append .CreateParameter("@NRORUCPROVEEDOR", adVarChar, adParamInput, 50, strCodProveedor)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY NROOC, NROPEDIDO DESC, NOMPRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Numero de Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .DisableEditor = True
                .FieldName = "NROOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Codigo de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo Proveedor"
                .DisableEditor = True
                .FieldName = "CODPROVEEDOR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Informacion de O/C
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "O/C"
                .DisableEditor = True
                .FieldName = "INFOOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Llave de Presentacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Llave"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Descripcion del Producto Interna
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            'Columna Descripcion del Producto del Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción del Proveedor"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTOPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcionProv"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
                .Visible = False
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

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Saldo"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Recepción"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna PRECIOUNITARIO
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "P. Unit."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PRECIOUNITARIO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColPrecioUnitario"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Seleccionar"
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFOOC").GroupIndex = 0
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaFaltantePorOC:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaFaltantePorOC"
    End Select
    
    Err.Clear
End Sub

'Listar Ingresos por Compra Disponibles en Stock (Comprometido y Libre) en Grilla (QuamtumGrid)
Public Sub listarGrillaIngresoPorOCdisponible(ByVal grilla As dxDBGrid, _
                                                ByVal listaImagenes As Object, _
                                                ByVal strCodAlmacen As String, _
                                                ByVal strCodProducto As String, _
                                                ByVal strIngreseCorL As String, _
                                                Optional ByVal strNroPedido As String)
    
    On Error GoTo errListarGrillaIngresoPorOCdisponible
    
    strSQLVale = vbNullString
    
    strSQLVale = strSQLVale & "INSERT INTO TMPUTILDEVOLUCIONOC("
    strSQLVale = strSQLVale & "NROOC, CODPROVEEDOR, INFOOC, LLAVE, NROPEDIDO, CODPRODUCTO, CODPRODUCTOORIGINAL, "
    strSQLVale = strSQLVale & "NOMPRODUCTO, UM, CANTIDAD, CANTIDADDESTINO, PROCESAR) "
    
    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    'Stock Fisico (Stock en Almacen: Comprometido o Libre)
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DISPONIBLE.OC, "
    strSQLVale = strSQLVale & "OC.F4CODPRV, "
    'strSQLVale = strSQLVale & "('[ No.: ' & DISPONIBLE.OC & ' ]" & Space(50) & "[Fecha: ' & OC.F4FECEMI & ']" & Space(50) & "[ Proveedor: ' & OC.F4NOMPROV & ' ]') AS INFOOC, "
    strSQLVale = strSQLVale & "('[ No.: ' & IIF(TRIM(DISPONIBLE.OC & '') <> '', DISPONIBLE.OC, 'INGRESOS SIN O/C') & ' ]" & Space(50) & "[Fecha: ' & OC.F4FECEMI & ']" & Space(50) & "[ Proveedor: ' & OC.F4NOMPROV & ' ]') AS INFOOC, "
    strSQLVale = strSQLVale & "(DISPONIBLE.OC & DISPONIBLE.PEDIDO & DISPONIBLE.F5CODPRO & DISPONIBLE.F5CODPROORIGINAL) AS LLAVE, "
    strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) AS CANTIDAD, "
    strSQLVale = strSQLVale & "0 AS CANTIDADDESTINO, "
    strSQLVale = strSQLVale & "FALSE AS PROCESAR "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "((((("
    
    'SCTOCK DISPONIBLE DEL PRODUCTO
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.TIPO = 'I' AND "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
    
    'strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
    'strSQLVale = strSQLVale & "DET.COD_SOLICITUD " & IIf(strIngreseCorL = "C", "<>", "=") & " '' "
    
    strSQLVale = strSQLVale & "CAB.F1CODORI = 'XC0' " 'AND "
    'strSQLVale = strSQLVale & "DET.F4NUMORD <> '' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS DISPONIBLE "
    
    'SCTOCK CONSUMIDO DEL PRODUCTO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.TIPO = 'S' AND "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS CONSUMIDO "
    strSQLVale = strSQLVale & "ON CONSUMIDO.F5CODPRO = DISPONIBLE.F5CODPRO AND CONSUMIDO.F2CODALM = DISPONIBLE.F2CODALM AND CONSUMIDO.PEDIDO = DISPONIBLE.PEDIDO AND CONSUMIDO.OC = DISPONIBLE.OC) "
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DISPONIBLE.F5CODPRO) "
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    
    strSQLVale = strSQLVale & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DISPONIBLE.F2CODALM) "
    
    strSQLVale = strSQLVale & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DISPONIBLE.PEDIDO) "
    
    strSQLVale = strSQLVale & "LEFT JOIN IF4ORDEN AS OC ON OC.F4NUMORD = DISPONIBLE.OC "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) > 0 "
        
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F2CODALM, "
    strSQLVale = strSQLVale & "DISPONIBLE.OC, "
    strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO"
    
    If Not grilla Is Nothing Then
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILDEVOLUCIONOC"
        
        abrirCnTemporal
        
        cnBdCPlus.Execute strSQLVale
        
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Numero de Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .DisableEditor = True
                .FieldName = "NROOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Codigo de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo Proveedor"
                .DisableEditor = True
                .FieldName = "CODPROVEEDOR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Informacion de O/C
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "O/C"
                .DisableEditor = True
                .FieldName = "INFOOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Llave de Presentacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Llave"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
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

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Procesar"
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILDEVOLUCIONOC ORDER BY NROOC, NROPEDIDO DESC, NOMPRODUCTO"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFOOC").GroupIndex = 0
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaIngresoPorOCdisponible:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaIngresoPorOCdisponible"
    End Select
    
    Err.Clear
End Sub

'Listar Stock Disponible (Libre) en Grilla (QuamtumGrid)
Public Sub listarGrillaStockDisponible(ByVal grilla As dxDBGrid, _
                                        ByVal listaImagenes As Object, _
                                        ByVal strCodAlmacen As String, _
                                        ByVal strCadenaCorte As String)
    
    On Error GoTo errListarGrillaStockDisponible
    
    strSQLVale = vbNullString
    
    strSQLVale = strSQLVale & "INSERT INTO TMPUTILSTOCKDISPONIBLE("
    strSQLVale = strSQLVale & "NROOC, CODPROVEEDOR, INFOPRODUCTO, LLAVE, NROPEDIDO, CODPRODUCTO, CODPRODUCTOORIGINAL, "
    strSQLVale = strSQLVale & "NOMPRODUCTO, UM, CANTIDAD, CANTIDADDESTINO, PROCESAR) "
    
    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    'Stock Fisico (Stock en Almacen: Libre)
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DISPONIBLE.OC, "
    strSQLVale = strSQLVale & "OC.F4CODPRV, "
    strSQLVale = strSQLVale & "(PROD.F5NOMPRO & '   ( ' & MED.F7NOMMED & ' )') AS INFOPRODUCTO, "
    strSQLVale = strSQLVale & "(DISPONIBLE.OC & DISPONIBLE.PEDIDO & DISPONIBLE.F5CODPRO & DISPONIBLE.F5CODPROORIGINAL) AS LLAVE, "
    strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) AS CANTIDAD, "
    strSQLVale = strSQLVale & "0 AS CANTIDADDESTINO, "
    strSQLVale = strSQLVale & "FALSE AS PROCESAR "
    
    strSQLVale = strSQLVale & "FROM "
    
    strSQLVale = strSQLVale & "((((("
    
    
    
    'SCTOCK DISPONIBLE DEL PRODUCTO
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.TIPO = 'I' AND "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS DISPONIBLE "
    
    'SCTOCK CONSUMIDO DEL PRODUCTO
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.TIPO = 'S' AND "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' "
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD) AS CONSUMIDO "
    strSQLVale = strSQLVale & "ON CONSUMIDO.F5CODPRO = DISPONIBLE.F5CODPRO AND CONSUMIDO.F2CODALM = DISPONIBLE.F2CODALM AND CONSUMIDO.PEDIDO = DISPONIBLE.PEDIDO AND CONSUMIDO.OC = DISPONIBLE.OC) "
    
    
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DISPONIBLE.F5CODPRO) "
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    
    strSQLVale = strSQLVale & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DISPONIBLE.F2CODALM) "
    
    strSQLVale = strSQLVale & "LEFT JOIN TB_CABSOLICITUD AS CABPED ON CABPED.COD_SOLICITUD = DISPONIBLE.PEDIDO) "
    
    strSQLVale = strSQLVale & "LEFT JOIN IF4ORDEN AS OC ON OC.F4NUMORD = DISPONIBLE.OC "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "(VAL(DISPONIBLE.CANTIDAD & '') - VAL(CONSUMIDO.CANTIDAD & '')) > 0 "
        
        If strCadenaCorte <> vbNullString Then
            strSQLVale = strSQLVale & "AND PROD.F5NOMPRO LIKE '%" & strCadenaCorte & "%' "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "DISPONIBLE.F5CODPRO, "
    strSQLVale = strSQLVale & "DISPONIBLE.F2CODALM, "
    strSQLVale = strSQLVale & "DISPONIBLE.OC, "
    strSQLVale = strSQLVale & "DISPONIBLE.PEDIDO"
    
    If Not grilla Is Nothing Then
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKDISPONIBLE"
        
        abrirCnTemporal
        
        cnBdCPlus.Execute strSQLVale
        
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Numero de Orden de Compra
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. O/C"
                .Color = RGB(250, 192, 144)
                .DisableEditor = True
                .FieldName = "NROOC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo Proveedor"
                .DisableEditor = True
                .FieldName = "CODPROVEEDOR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Informacion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Producto"
                .DisableEditor = True
                .FieldName = "INFOPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfoProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Llave de Presentacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Llave"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                '.Visible = False
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

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad Requerida"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Procesar"
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILSTOCKDISPONIBLE ORDER BY NOMPRODUCTO, NROOC"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFOPRODUCTO").GroupIndex = 0
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaStockDisponible:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaStockDisponible"
    End Select
    
    Err.Clear
End Sub

'Listar Stock Disponible (Libre) en Grilla (QuamtumGrid) - Version 2
Public Sub listarGrillaStockDisponibleV2(ByVal grilla As dxDBGrid, _
                                        ByVal listaImagenes As Object, _
                                        ByVal strCadenaCorte As String, _
                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaStockDisponibleV2
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarStockDisponible"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strCadenaCorte)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY NOMPRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
                '.Visible = False
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

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
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
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 50
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad Requerida"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstSum
                .Width = 50
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Procesar"
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPRODUCTO"
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaStockDisponibleV2:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaStockDisponibleV2"
    End Select
    
    Err.Clear
End Sub

'Listar Insumos Descargados de OP para su Devolucion en Grilla (QuamtumGrid)
Public Sub listarGrillaDevolucionPorOP(ByVal grilla As dxDBGrid, _
                                        ByVal listaImagenes As Object, _
                                        ByVal strCodAlmacen As String, _
                                        ByVal strIdOrdenProduccion As String, _
                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaDevolucionPorOP
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_DevolucionPorOP"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@IDORDENPRODUCCION", adVarChar, adParamInput, 50, strIdOrdenProduccion)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY NOMPRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Id de Orden Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Id O/P"
                .DisableEditor = True
                .FieldName = "IDORDENPRODUCCION"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOrdenProduccion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Llave de Presentacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Llave"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
'            'Columna Numero de Orden de Compra
'            Set gColumn = .Columns.Add(gedTextEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .Caption = "No. O/C"
'                .DisableEditor = True
'                .FieldName = "NROOC"
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColNoOC"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "No. Pedido"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNoPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo Origen"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGINAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoOriginal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
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

            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Cantidad Destino
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .FieldName = "CANTIDADDESTINO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadDestino"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Procesar"
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "LLAVE"
            
            '.Columns.ColumnByFieldName("INFOOC").GroupIndex = 0
            
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("CANTIDADDESTINO").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaDevolucionPorOP:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaDevolucionPorOP"
    End Select
    
    Err.Clear
End Sub

'Listar Alerta para Reposicion de Stock de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaAlertaReposicionStockProducto(ByVal grilla As dxDBGrid, _
                                                        ByVal strCodAlmacen As String, _
                                                        ByVal strCodFamilia As String, _
                                                        ByVal strCodSubFamilia As String, _
                                                        ByVal strFiltroSensitivo As String, _
                                                        ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaAlertaReposicionStockProducto
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        '.CommandTimeout = "180"
        .CommandText = "Procesos.usp_AlertaReposicionStockProducto"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@CODFAMILIA", adVarChar, adParamInput, 10, strCodFamilia)
        .Parameters.Append .CreateParameter("@CODSUBFAMILIA", adVarChar, adParamInput, 10, strCodSubFamilia)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strFiltroSensitivo)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "FAMILIA, "
    strSQLVale = strSQLVale & "SUBFAMILIA, "
    strSQLVale = strSQLVale & "F5CODPRO, "
    strSQLVale = strSQLVale & "F5NOMPRO, "
    strSQLVale = strSQLVale & "F7SIGMED, "
    strSQLVale = strSQLVale & "F5STOCKMIN, "
    strSQLVale = strSQLVale & "F5STOCKLOG, "
    strSQLVale = strSQLVale & "F5STOCKMAX, "
    strSQLVale = strSQLVale & "ENALMACENCOMP, "
    strSQLVale = strSQLVale & "PORLLEGARCOMP, "
    strSQLVale = strSQLVale & "ENALMACENLIBRE, "
    strSQLVale = strSQLVale & "PORLLEGARLIBRE, "
    strSQLVale = strSQLVale & "STOCKENALMACENTOTAL, "
    strSQLVale = strSQLVale & "STOCKPORLLEGARTOTAL, "
    strSQLVale = strSQLVale & "STOCKTOTAL, "
    strSQLVale = strSQLVale & "IIF(STOCKENALMACENTOTAL <= F5STOCKMIN, 0, IIF(STOCKENALMACENTOTAL > F5STOCKMIN AND STOCKENALMACENTOTAL <= F5STOCKLOG, 1, IIF(STOCKENALMACENTOTAL > F5STOCKLOG AND STOCKENALMACENTOTAL <= F5STOCKMAX, 2, 3))) ESTADO, "
    strSQLVale = strSQLVale & "IIF(STOCKENALMACENTOTAL < F5STOCKLOG, F5STOCKLOG - STOCKENALMACENTOTAL, 0) STOCKAREPONER "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & strNombreTablaSQLParaResumen & " "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "ESTADO, "
    strSQLVale = strSQLVale & "STOCKENALMACENTOTAL"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Familia del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Familia"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Sub-Familia del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Sub-Familia"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Minimo del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Stk Minimo"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKMIN"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockMinimo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Reposicion del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Stk Reposición"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKLOG"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockReposicion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Maximo del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Stk Maximo"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKMAX"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockMaximo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Comprometido en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "ENALMACENCOMP"
                .Font.Charset = 0
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColEnAlmacenComp"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Comprometido por llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PORLLEGARCOMP"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColPorLlegarComp"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "ENALMACENLIBRE"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColEnAlmacenLibre"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PORLLEGARLIBRE"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColPorLlegarLibre"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKENALMACENTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Total por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPORLLEGARTOTAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockPorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Sctock Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Stock Total"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKTOTAL"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockTotal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Estado del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Estado"
                .BandIndex = 4
                .DisableEditor = True
                .FieldName = "ESTADO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Stock A Reponer
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Stk a Reponer"
                .Color = &HFFFFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKAREPONER"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockAReponer"
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
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            '.m.FullExpand
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            '.Columns.ColumnByFieldName("ESTADO").Sorted = csUp
            
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaAlertaReposicionStockProducto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaAlertaReposicionStockProducto"
    End Select
    
    Err.Clear
End Sub

'Listar Alerta de Stock Minimo de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaAlertaStockMinimoProducto(ByVal grilla As dxDBGrid, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strCodFamilia As String, _
                                                    ByVal strCodSubFamilia As String, _
                                                    ByVal strFiltroSensitivo As String, _
                                                    ByVal strNombreTablaSQLParaResumen As String)
    
    On Error GoTo errListarGrillaAlertaStockMinimoProducto
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        '.CommandTimeout = "180"
        .CommandText = "Procesos.usp_AlertaStockMinimoProducto"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@CODFAMILIA", adVarChar, adParamInput, 10, strCodFamilia)
        .Parameters.Append .CreateParameter("@CODSUBFAMILIA", adVarChar, adParamInput, 10, strCodSubFamilia)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strFiltroSensitivo)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY ESTADO, LEA"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Familia del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Familia"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Sub-Familia del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Sub-Familia"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Minimo del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Stk Minimo"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKMIN"
                .Font.Charset = 0
                '.Font.Bold = True
                '.FontColor = &HFFFFFF
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockMinimo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre en Almacen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LEA"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColLibreEnAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre por Llegar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LPL"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColLibrePorLlegar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Libre Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Total"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LTOTAL"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColLibreStockTotal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Por Reponer
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Reponer"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "REPONER"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColReponer"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Estado del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Estado"
                .BandIndex = 4
                .DisableEditor = True
                .FieldName = "ESTADO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Descripcion de Estado
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Descripcion Estado"
                .Color = &HFFFFFF
                .DisableEditor = True
                .FieldName = "ESTADOLETRA"
                .Font.Bold = True
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstadoDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            '.Columns.ColumnByFieldName("ESTADO").Sorted = csUp
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaAlertaStockMinimoProducto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaAlertaStockMinimoProducto"
    End Select
    
    Err.Clear
End Sub

'Listar Consumo Anual de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaConsumoAnualProducto(ByVal grilla As dxDBGrid, _
                                            ByVal intAnno As Integer, _
                                            ByVal intOpcionResumen As Integer, _
                                            ByVal strFiltroSensitivo As String, _
                                            Optional ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaConsumoAnualProducto
    
    Dim intCol As Integer
    Dim intColMax As Integer
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "TRANSFORM "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) "
    
    strSQLVale = strSQLVale & "SELECT "
    
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION AS TIPO, "
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
            Case Else
                strSQLVale = strSQLVale & "TE.DESCRIPCION AS TIPO, "
                strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
                strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
                strSQLVale = strSQLVale & "PROD.F5NOMPRO AS DESCRIPCION, "
        End Select
    
    strSQLVale = strSQLVale & "MED.F7SIGMED "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(((((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN EF2TIPOEXISTENCIA AS TE ON TE.CODIGO = PROD.F5TIPO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "YEAR(CAB.F4FECVAL) = " & intAnno & " AND "
    strSQLVale = strSQLVale & "CAB.F4TIPOVALE = 'S' "
    strSQLVale = strSQLVale & "GROUP BY "
        
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION, "
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON, "
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON, "
            Case Else
                strSQLVale = strSQLVale & "TE.DESCRIPCION, "
                strSQLVale = strSQLVale & "FAM.F7DESCON, "
                strSQLVale = strSQLVale & "SFAM.F7DESCON, "
                strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
        End Select
    
    strSQLVale = strSQLVale & "MED.F7SIGMED "
    strSQLVale = strSQLVale & "ORDER BY "
    
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION "
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON "
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON "
            Case Else
                strSQLVale = strSQLVale & "PROD.F5NOMPRO "
        End Select
    
    strSQLVale = strSQLVale & "PIVOT "
    strSQLVale = strSQLVale & "UCASE(FORMAT(CAB.F4FECVAL, 'MM'))"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            If intOpcionResumen = 0 Or intOpcionResumen = 3 Then
                'Columna Tipo Existencia (Familia Contable)
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Tipo de Existencia"
                    .DisableEditor = True
                    .FieldName = "TIPO"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColTipoExistencia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 1 Or intOpcionResumen = 3 Then
                'Columna Familia de Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Familia"
                    .DisableEditor = True
                    .FieldName = "FAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 2 Or intOpcionResumen = 3 Then
                'Columna Sub-Familia
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Sub-Familia"
                    .DisableEditor = True
                    .FieldName = "SUBFAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColSubFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 3 Then
                'Columna Descripción del Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Descripción de Producto"
                    .DisableEditor = True
                    .FieldName = "DESCRIPCION"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColDescripcionProducto"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 200
                End With
            End If
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            If intAnno < Year(Date) Then
                intColMax = 12
            Else
                intColMax = Month(Date)
            End If
            
            For intCol = 1 To intColMax
                'Columna Consumo Mensual del Ejercicio
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .Caption = UCase(Format("01/" & Format(intCol, "00") & "/" & intAnno, "MMMM"))
                    
                    If intCol Mod 2 Then
                        .Color = &H80FFFF
                    Else
                        .Color = vbWhite
                    End If
                    
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = Format(intCol, "00")
                    .Font.Bold = False
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "Col" & Format(intCol, "00")
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
            Next intCol
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            
            Select Case intOpcionResumen
                Case 0 'Por Tipo de Existencia
                    .KeyField = "TIPO"
                    
                    .Columns.ColumnByFieldName("TIPO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
                Case 1 'Por Familia
                    .KeyField = "FAMILIA"
                    
                    .Columns.ColumnByFieldName("FAMILIA").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
                Case 2 'Por Sub-Familia
                    .KeyField = "SUBFAMILIA"
                    
                    .Columns.ColumnByFieldName("SUBFAMILIA").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
                Case Else
                    .KeyField = "DESCRIPCION"
                    
                    .Columns.ColumnByFieldName("DESCRIPCION").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            End Select
            
            .Filter.FilterActive = True
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaConsumoAnualProducto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaConsumoAnualProducto"
    End Select
    
    Err.Clear
End Sub

'Listar Consumo Anual de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaConsumoAnualProductoV2(ByVal grilla As dxDBGrid, _
                                                ByVal intAnno As Integer, _
                                                ByVal intOpcionResumen As Integer, _
                                                ByVal intVerDetalle As Integer, _
                                                ByVal strFiltroEspecifico As String, _
                                                ByVal strFiltroSensitivo As String, _
                                                ByVal strNombreTablaSQLParaResumen As String, _
                                                Optional ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaConsumoAnualProductoV2
    
    Dim intCol As Integer
    Dim intColMax As Integer
    Dim strNomMes As String
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ConsumoAnualProducto"
        
        .Parameters.Append .CreateParameter("@ANNO", adVarChar, adParamInput, 12, str(intAnno))
        .Parameters.Append .CreateParameter("@OPCIONRESUMEN", adInteger, adParamInput, , intOpcionResumen)
        .Parameters.Append .CreateParameter("@VERDETALLE", adInteger, adParamInput, , intVerDetalle)
        .Parameters.Append .CreateParameter("@FILTROESPECIFICO", adVarChar, adParamInput, 255, strFiltroEspecifico)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strFiltroSensitivo)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    intColMax = Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "TOP 1 MES", strNombreTablaSQLParaResumen, "", "", "", "MES > 0 GROUP BY MES ORDER BY MES DESC"))
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "RESUMEN.TIPO, "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "RESUMEN.FAMILIA, "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "RESUMEN.SUBFAMILIA, "
        Else
            strSQLVale = strSQLVale & "RESUMEN.TIPO, "
            strSQLVale = strSQLVale & "RESUMEN.FAMILIA, "
            strSQLVale = strSQLVale & "RESUMEN.SUBFAMILIA, "
            strSQLVale = strSQLVale & "RESUMEN.DESCRIPCION, "
        End If
        
    strSQLVale = strSQLVale & "RESUMEN.UM, "
        
        For intCol = 1 To intColMax
            strNomMes = ModUtilitario.devuelveNombreMes(Format(intCol, "00"))
            
            strSQLVale = strSQLVale & "CONVERT(DECIMAL(10, 2), ISNULL(" & strNomMes & ".CANT, 0)) AS " & left(strNomMes, 4) & ", "
        Next intCol
    
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10, 2), ISNULL(TOTAL.CANT, 0)) AS TOT, "
    strSQLVale = strSQLVale & "CONVERT(DECIMAL(10, 2), (ISNULL(TOTAL.CANT, 0) / " & intColMax & ")) AS PROMEDIO "
    
    strSQLVale = strSQLVale & "FROM "
        
    'strSQLVale = strSQLVale & Repetir(intColMax, "(")
        
    strSQLVale = strSQLVale & "(SELECT "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TIPO, "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "FAMILIA, "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "SUBFAMILIA, "
        Else
            strSQLVale = strSQLVale & "TIPO, "
            strSQLVale = strSQLVale & "FAMILIA, "
            strSQLVale = strSQLVale & "SUBFAMILIA, "
            strSQLVale = strSQLVale & "DESCRIPCION, "
        End If
    
    strSQLVale = strSQLVale & "UM "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & UCase(strNombreTablaSQLParaResumen) & " "
    strSQLVale = strSQLVale & "GROUP BY "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TIPO, "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "FAMILIA, "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "SUBFAMILIA, "
        Else
            strSQLVale = strSQLVale & "TIPO, "
            strSQLVale = strSQLVale & "FAMILIA, "
            strSQLVale = strSQLVale & "SUBFAMILIA, "
            strSQLVale = strSQLVale & "DESCRIPCION, "
        End If
        
    strSQLVale = strSQLVale & "UM) AS RESUMEN "
        
        For intCol = 1 To intColMax
            strNomMes = ModUtilitario.devuelveNombreMes(Format(intCol, "00"))
            
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "("
            strSQLVale = strSQLVale & "SELECT "
                
                'Por Tipo de Existencia
                If intOpcionResumen = 0 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "TIPO, "
                'Por Familia
                ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "FAMILIA, "
                'Por Sub-Familia
                ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "SUBFAMILIA, "
                Else
                    strSQLVale = strSQLVale & "TIPO, "
                    strSQLVale = strSQLVale & "FAMILIA, "
                    strSQLVale = strSQLVale & "SUBFAMILIA, "
                    strSQLVale = strSQLVale & "DESCRIPCION, "
                End If
                
            strSQLVale = strSQLVale & "UM, SUM(CONVERT(DECIMAL(10, 2), CANTIDAD)) AS CANT "
            strSQLVale = strSQLVale & "FROM " & UCase(strNombreTablaSQLParaResumen) & " "
            strSQLVale = strSQLVale & "WHERE MES = " & intCol & " "
            strSQLVale = strSQLVale & "GROUP BY "
                
                'Por Tipo de Existencia
                If intOpcionResumen = 0 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "TIPO, "
                'Por Familia
                ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "FAMILIA, "
                'Por Sub-Familia
                ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & "SUBFAMILIA, "
                Else
                    strSQLVale = strSQLVale & "TIPO, "
                    strSQLVale = strSQLVale & "FAMILIA, "
                    strSQLVale = strSQLVale & "SUBFAMILIA, "
                    strSQLVale = strSQLVale & "DESCRIPCION, "
                End If
                
            strSQLVale = strSQLVale & "UM"
            strSQLVale = strSQLVale & ") AS " & strNomMes & " ON "
                
                'Por Tipo de Existencia
                If intOpcionResumen = 0 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & strNomMes & ".TIPO = RESUMEN.TIPO AND " & strNomMes & ".UM = RESUMEN.UM "
                'Por Familia
                ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & strNomMes & ".FAMILIA = RESUMEN.FAMILIA AND " & strNomMes & ".UM = RESUMEN.UM "
                'Por Sub-Familia
                ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
                    strSQLVale = strSQLVale & strNomMes & ".SUBFAMILIA = RESUMEN.SUBFAMILIA AND " & strNomMes & ".UM = RESUMEN.UM "
                Else
                    strSQLVale = strSQLVale & strNomMes & ".TIPO = RESUMEN.TIPO AND "
                    strSQLVale = strSQLVale & strNomMes & ".FAMILIA = RESUMEN.FAMILIA AND "
                    strSQLVale = strSQLVale & strNomMes & ".SUBFAMILIA = RESUMEN.SUBFAMILIA AND "
                    strSQLVale = strSQLVale & strNomMes & ".DESCRIPCION = RESUMEN.DESCRIPCION AND "
                    strSQLVale = strSQLVale & strNomMes & ".UM = RESUMEN.UM "
                End If
        Next intCol
        
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TIPO, "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "FAMILIA, "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "SUBFAMILIA, "
        Else
            strSQLVale = strSQLVale & "TIPO, "
            strSQLVale = strSQLVale & "FAMILIA, "
            strSQLVale = strSQLVale & "SUBFAMILIA, "
            strSQLVale = strSQLVale & "DESCRIPCION, "
        End If
        
    strSQLVale = strSQLVale & "UM, SUM(CONVERT(DECIMAL(10, 2), CANTIDAD)) AS CANT "
    strSQLVale = strSQLVale & "FROM " & UCase(strNombreTablaSQLParaResumen) & " "
    strSQLVale = strSQLVale & "GROUP BY "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TIPO, "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "FAMILIA, "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "SUBFAMILIA, "
        Else
            strSQLVale = strSQLVale & "TIPO, "
            strSQLVale = strSQLVale & "FAMILIA, "
            strSQLVale = strSQLVale & "SUBFAMILIA, "
            strSQLVale = strSQLVale & "DESCRIPCION, "
        End If
        
    strSQLVale = strSQLVale & "UM"
    strSQLVale = strSQLVale & ") AS TOTAL ON "
    'TOTAL.FAMILIA = RESUMEN.FAMILIA AND TOTAL.UM = RESUMEN.UM "
        
        'Por Tipo de Existencia
        If intOpcionResumen = 0 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TOTAL.TIPO = RESUMEN.TIPO AND TOTAL.UM = RESUMEN.UM "
        'Por Familia
        ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TOTAL.FAMILIA = RESUMEN.FAMILIA AND TOTAL.UM = RESUMEN.UM "
        'Por Sub-Familia
        ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
            strSQLVale = strSQLVale & "TOTAL.SUBFAMILIA = RESUMEN.SUBFAMILIA AND TOTAL.UM = RESUMEN.UM "
        Else
            strSQLVale = strSQLVale & "TOTAL.TIPO = RESUMEN.TIPO AND "
            strSQLVale = strSQLVale & "TOTAL.FAMILIA = RESUMEN.FAMILIA AND "
            strSQLVale = strSQLVale & "TOTAL.SUBFAMILIA = RESUMEN.SUBFAMILIA AND "
            strSQLVale = strSQLVale & "TOTAL.DESCRIPCION = RESUMEN.DESCRIPCION AND "
            strSQLVale = strSQLVale & "TOTAL.UM = RESUMEN.UM "
        End If
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            If (intOpcionResumen = 0 Or intOpcionResumen = 3 Or intVerDetalle = 1) Then
                'Columna Tipo Existencia (Familia Contable)
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Tipo de Existencia"
                    .DisableEditor = True
                    .FieldName = "TIPO"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColTipoExistencia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 1 Or intOpcionResumen = 3 Or intVerDetalle = 1 Then
                'Columna Familia de Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Familia"
                    .DisableEditor = True
                    .FieldName = "FAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 2 Or intOpcionResumen = 3 Or intVerDetalle = 1 Then
                'Columna Sub-Familia
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Sub-Familia"
                    .DisableEditor = True
                    .FieldName = "SUBFAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColSubFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 3 Or intVerDetalle = 1 Then
                'Columna Descripción del Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Descripción de Producto"
                    .DisableEditor = True
                    .FieldName = "DESCRIPCION"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColDescripcionProducto"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 200
                End With
            End If
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'If intAnno < Year(Date) Then
            '    intColMax = 12
            'Else
            '    intColMax = Month(Date)
            'End If
            
            For intCol = 1 To intColMax
                strNomMes = ModUtilitario.devuelveNombreMes(Format(intCol, "00"))
                
                'Columna Consumo Mensual del Ejercicio
                Set gColumn = .Columns.Add(gedSpinEdit) '.Columns.Add(gedTextEdit) '
                
                With gColumn
                    .Alignment = taRightJustify
                    .Caption = strNomMes
                    
                    If intCol Mod 2 Then
                        .Color = &H80FFFF
                    Else
                        .Color = vbWhite
                    End If
                    
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = left(strNomMes, 4)
                    .Font.Bold = False
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "Col" & strNomMes
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
            Next intCol
            
            'Columna Consumo Total
            Set gColumn = .Columns.Add(gedSpinEdit) '.Columns.Add(gedTextEdit) '
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "TOTAL"
                .Color = RGB(149, 179, 215) '&H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOT"
                .Font.Bold = True
                .FontColor = RGB(255, 255, 255)
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColTOTAL"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Consumo Promedio
            Set gColumn = .Columns.Add(gedSpinEdit) '.Columns.Add(gedTextEdit) '
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "PROMEDIO"
                .Color = RGB(79, 129, 189) '&H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PROMEDIO"
                .Font.Bold = True
                .FontColor = RGB(255, 255, 255) '&H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColPROMEDIO"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseServer
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            
            'Por Tipo de Existencia
            If intOpcionResumen = 0 And intVerDetalle = 0 Then
                .KeyField = "TIPO"

                .Columns.ColumnByFieldName("TIPO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            'Por Familia
            ElseIf intOpcionResumen = 1 And intVerDetalle = 0 Then
                .KeyField = "FAMILIA"

                .Columns.ColumnByFieldName("FAMILIA").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            'Por Sub-Familia
            ElseIf intOpcionResumen = 2 And intVerDetalle = 0 Then
                .KeyField = "SUBFAMILIA"

                .Columns.ColumnByFieldName("SUBFAMILIA").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            Else
                .KeyField = "DESCRIPCION"

                .Columns.ColumnByFieldName("DESCRIPCION").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            End If
            
            .Filter.FilterActive = True
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaConsumoAnualProductoV2:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaConsumoAnualProductoV2"
    End Select
    
    Err.Clear
End Sub

Public Function devuelveSQLResumenVale() As String
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "CAB.F4NUMVAL, "
    strSQLVale = strSQLVale & "CAB.NUMENSAM, "
    strSQLVale = strSQLVale & "UPPER(ALM.F2NOMALM) AS ALMACEN, "
    strSQLVale = strSQLVale & "UPPER(ORI.F1NOMORI) AS CONCEPTO, "
    strSQLVale = strSQLVale & "PER.NOMBRE, "
    strSQLVale = strSQLVale & "CONVERT(NVARCHAR(10), CAB.F4FECVAL,103) AS F4FECVAL, "
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', 'SOLES', 'DOLARES') AS MONEDA, "
    strSQLVale = strSQLVale & "CAB.F4TIPCAM, "
    strSQLVale = strSQLVale & "IIF(RTRIM(LTRIM(CAB.F4SERGUIA)) <> '', CAB.F4SERGUIA + '-', '') + CAB.F4NUMGUIA AS GUIA, "
    strSQLVale = strSQLVale & "UPPER(DOC.F2ABREV) AS TIPDOC, "
    strSQLVale = strSQLVale & "IIF(RTRIM(LTRIM(CAB.F4SERDOC)) <> '', CAB.F4SERDOC + '-', '') + CAB.F4NUMDOC AS DOCUMENTO, "
    strSQLVale = strSQLVale & "CAB.F4ORDTRA, "
    strSQLVale = strSQLVale & "CAB.F4REGCOM, "
    strSQLVale = strSQLVale & "UPPER(CAB.F4OBSERVA) AS OBSERVACION, "
    strSQLVale = strSQLVale & "TE.DESCRIPCION AS TIPOEXISTENCIA, "
    
    'strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "IIF(CAB.F1CODORI = 'XC0', DET.F4NUMORD, '') AS F4NUMORD, "
    
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "DET.F3CANPRO, "
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL) AS COSTO, "
    
        'Select Case strTipoVale
        '    Case "I"
                strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE - DET.F3IGV, DET.F3TOTDOL - DET.F3IGVDOL) AS SUBTOTAL, "
                strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3IGV, DET.F3IGVDOL) AS IGV, "
                strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE, DET.F3TOTDOL) AS TOTAL "
        '    Case "S"
        '        strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE - DET.F3IGV, DET.F3TOTDOL - DET.F3IGVDOL) AS SUBTOTAL "
        '        strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3IGV, DET.F3IGVDOL) AS IGV, "
        '        strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE + DET.F3IGV, DET.F3TOTDOL + DET.F3IGVDOL) AS TOTAL "
        'End Select
    
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = CAB.F2CODALM "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.DOCUMENTOS AS DOC ON DOC.F2CODDOC = CAB.F4TIPDOC "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2TIPOEXISTENCIA AS TE ON TE.CODIGO = PROD.F5TIPO "
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT 'C' AS TIPO, F2CODCLI AS CODIGO, F2NOMCLI AS NOMBRE FROM MAESTROS.EF2CLIENTES "
    strSQLVale = strSQLVale & "UNION "
    strSQLVale = strSQLVale & "SELECT 'P' AS TIPO, F2CODPROV AS CODIGO, F2NOMPROV AS NOMBRE FROM MAESTROS.EF2PROVEEDORES "
    strSQLVale = strSQLVale & ") AS PER ON PER.TIPO = CAB.F1TIPPRV AND PER.CODIGO = CAB.F2CODPROV "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLVale = strSQLVale & "DET.F4NUMVAL = '" & strNumeroVale & "'"
    
    devuelveSQLResumenVale = strSQLVale
    
    strSQLVale = vbNullString
End Function

Public Sub obtenerSaldoYCostoInicialDeProducto()
    On Error GoTo errObtenerSaldoYCostoInicialDeProducto
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) AS CANTIDAD, "
    strSQLVale = strSQLVale & "SUM((DET." & IIf(strCodMoneda = "S", "F3VALVTA", "F3VALDOL") & " * DET.F3CANPRO) * IIF(DET.TIPO = 'I', 1, -1)) AS VALVTA "
    'strSQLVale = strSQLVale & "(VALVTA/CANTIDAD) AS CP "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, DET.F3CANPRO, " & IIf(strCodMoneda = "S", "DET.F3VALVTA", "DET.F3VALDOL") & " "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM))"
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4FECVAL < '" & StrFecha & "' AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "AND DET.F2CODALM = '" & strCodAlmacen & "' "
        End If
        
    'strSQLVale = strSQLVale & "ORDER BY "
    'strSQLVale = strSQLVale & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL"
    strSQLVale = strSQLVale & ") AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB "
    strSQLVale = strSQLVale & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        dblCantidad = Val(Format(Val(rstVale!Cantidad & ""), "#0.0000"))
        
        If Val(rstVale!Cantidad & "") > 0 Then
            dblValorVta = Val(Format(Val(rstVale!ValVta & "") / Val(rstVale!Cantidad & ""), "#.0000"))
            
            If dblValorVta < 0 Then dblValorVta = 0
        End If
        
        dblTotal = Val(Format(Val(rstVale!ValVta & ""), "#0.0000"))
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Sub
    
    Resume
errObtenerSaldoYCostoInicialDeProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ObtenerSaldoYCostoInicialDeProducto"
    
    Err.Clear
End Sub

Public Function devuelveStockFisicoDeProducto(Optional ByVal strIngreseCorLorVacio As String, _
                                                Optional ByVal bolEvaluarStockAlaFechaHoy As Boolean) As Double
    
    On Error GoTo errDevuelveStockFisicoDeProducto
    
    devuelveStockFisicoDeProducto = 0
    
    Dim cmdOP As ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_DevuelveStockFisicoDeProducto"
        
        .Parameters.Append .CreateParameter("@FECHACORTE", adDate, adParamInput, 10, IIf(StrFecha = vbNullString, Date, StrFecha))
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@CODIGOALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@ESTADOSTOCK", adVarChar, adParamInput, 1, strIngreseCorLorVacio)
        .Parameters.Append .CreateParameter("@EVALUARSTOCKALAFECHAHOY", adBigInt, adParamInput, , IIf(bolEvaluarStockAlaFechaHoy, 1, 0))
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        If Val(rstVale!Cantidad & "") > 0 Then
            devuelveStockFisicoDeProducto = Val(Format(Val(rstVale!Cantidad & ""), "#0.00"))
        End If
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
    Resume
errDevuelveStockFisicoDeProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DevuelveStockFisicoDeProducto"
    
    devuelveStockFisicoDeProducto = 0
    
    Err.Clear
End Function

Public Function devuelveStockFisicoDeProductoV2(ByVal strIngreseCorLorVacio As String, _
                                                Optional ByVal bolEvaluarStockAlaFechaHoy As Boolean) As Double
    
    On Error GoTo errDevuelveStockFisicoDeProductoV2
    
    Dim intAnno As Integer, intMes As Integer, strFechaInicial As String
    
    devuelveStockFisicoDeProductoV2 = 0
    
    'Obtener Periodo de Ultimo Cierre
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT TOP 1 ANNO, MES FROM IF3CIERREMENSUAL GROUP BY ANNO, MES ORDER BY ANNO DESC, MES DESC"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        intAnno = Val(rstVale!Anno & "")
        intMes = Val(rstVale!mes & "")
    Else
        If Not bolEvaluarStockAlaFechaHoy Then
            intAnno = IIf(Month(StrFecha) > 1, Year(StrFecha), Year(StrFecha) - 1)
            intMes = IIf(Month(StrFecha) > 1, Month(StrFecha) - 1, 12)
        Else
            intAnno = IIf(Month(Date) > 1, Year(Date), Year(Date) - 1)
            intMes = IIf(Month(Date) > 1, Month(Date) - 1, 12)
        End If
    End If
    
    strFechaInicial = "01/" & IIf(Val(intMes) < 12, Format(Val(intMes + 1), "00"), "01") & "/" & IIf(Val(intMes) < 12, intAnno, Val(intAnno) + 1)
    
    
    
    
    'Obtener el Stock del Ultimo Cierre
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "CODPRODUCTO, SUM(CANTIDAD) AS STOCK "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "IF3CIERREMENSUAL  "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "ANNO = " & intAnno & " AND "
    strSQLVale = strSQLVale & "MES = " & intMes & " AND "
    strSQLVale = strSQLVale & "TIPO = 'F' AND "
    strSQLVale = strSQLVale & "CODPRODUCTO = '" & strCodProducto & "' "
        
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "AND ALMACEN = '" & strCodAlmacen & "' "
        End If
        
        Select Case strIngreseCorLorVacio
            Case "C"
                strSQLVale = strSQLVale & "AND COMPROMISO <> '' "
            Case "L"
                strSQLVale = strSQLVale & "AND COMPROMISO = '' "
        End Select
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "CODPRODUCTO"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        devuelveStockFisicoDeProductoV2 = Val(rstVale!Stock & "")
    End If
    
    
    
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
        
        If Not bolEvaluarStockAlaFechaHoy Then
            strSQLVale = strSQLVale & "(CAB.F4FECVAL BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & StrFecha & "')) AND "
        Else
            strSQLVale = strSQLVale & "(CAB.F4FECVAL BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & Format(Date, "Short Date") & "')) AND "
        End If
        
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "AND CAB.F2CODALM = '" & strCodAlmacen & "' "
        End If
        
        Select Case strIngreseCorLorVacio
            Case "C"
                strSQLVale = strSQLVale & "AND DET.COD_SOLICITUD <> '' "
            Case "L"
                strSQLVale = strSQLVale & "AND DET.COD_SOLICITUD = '' "
        End Select
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        devuelveStockFisicoDeProductoV2 = devuelveStockFisicoDeProductoV2 + Val(Format(Val(rstVale!Cantidad & ""), "#0.00"))
        
        If devuelveStockFisicoDeProductoV2 < 0 Then
            devuelveStockFisicoDeProductoV2 = 0
        End If
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errDevuelveStockFisicoDeProductoV2:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DevuelveStockFisicoDeProductoV2"
    
    devuelveStockFisicoDeProductoV2 = 0
    
    Err.Clear
End Function

Public Function devuelveStockFisicoDeProductoConADO(ByVal strIngreseCorLorVacio As String, _
                                                    Optional ByVal bolEvaluarStockAlaFechaHoy As Boolean) As Double
    
    On Error GoTo errDevuelveStockFisicoDeProductoConADO
    
    devuelveStockFisicoDeProductoConADO = 0
    
    'Set rstVale = New ADODB.Recordset
    
    Dim cnDAO As DAO.Database
    Dim rstValeDetDAO As DAO.Recordset
    
    Dim dblCantidadTotal As Double
    
    dblCantidadTotal = 0
    
    If bolEvaluarStockAlaFechaHoy Then
        StrFecha = Format(Date, "Short Date")
    End If
    
    Set cnDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    Set rstValeDetDAO = cnDAO.OpenRecordset("PROCESOS.IF3VALES")
    
    rstValeDetDAO.Index = "IndiceValeProductoAlmacen"
    
    rstValeDetDAO.Seek "=", strCodProducto, strCodAlmacen
    
    If Not rstValeDetDAO.NoMatch Then
        
        Do While (CDate(rstValeDetDAO!F4FECVAL) <= CDate(StrFecha) And _
                    Trim(rstValeDetDAO!f5codpro & "") = strCodProducto And _
                    Trim(rstValeDetDAO!f2codalm & "") = strCodAlmacen)
            
            Select Case strIngreseCorLorVacio
                Case "C"
                    If Trim(rstValeDetDAO!COD_SOLICITUD & "") <> vbNullString Then
                        dblCantidadTotal = dblCantidadTotal + (Val(rstValeDetDAO!F3CANPRO & "") * IIf(Trim(rstValeDetDAO!Tipo & "") = "I", 1, -1))
                    End If
                Case "L"
                    If Trim(rstValeDetDAO!COD_SOLICITUD & "") = vbNullString Then
                        dblCantidadTotal = dblCantidadTotal + (Val(rstValeDetDAO!F3CANPRO & "") * IIf(Trim(rstValeDetDAO!Tipo & "") = "I", 1, -1))
                    End If
            End Select
            
            rstValeDetDAO.MoveNext
            
            If rstValeDetDAO.EOF Then Exit Do
            
            If Trim(rstValeDetDAO!f5codpro & "") <> strCodProducto Or _
                Trim(rstValeDetDAO!f2codalm & "") <> strCodAlmacen Then Exit Do
        Loop
    End If
    
    If dblCantidadTotal <> 0 Then
        devuelveStockFisicoDeProductoConADO = Val(Format(dblCantidadTotal, "#0.0000"))
    End If
    
    rstValeDetDAO.Close
    
    Set rstValeDetDAO = Nothing
    
    cnDAO.Close
    
    Set cnDAO = Nothing
    
    dblCantidadTotal = 0
    
    Exit Function
errDevuelveStockFisicoDeProductoConADO:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DevuelveStockFisicoDeProductoConADO"
    
    devuelveStockFisicoDeProductoConADO = 0
    
    Err.Clear
End Function

Public Function calcularCostoPromedio() As Double
    On Error GoTo errCalcularCostoPromedio
    
    calcularCostoPromedio = 0
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) AS CANTIDAD, "
    strSQLVale = strSQLVale & "SUM((DET." & IIf(strCodMoneda = "S", "F3VALVTA", "F3VALDOL") & " * DET.F3CANPRO) * IIF(DET.TIPO = 'I',1,-1)) AS VALVTA "
    'strSQLVale = strSQLVale & "(VALVTA/CANTIDAD) AS CP "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB "
    strSQLVale = strSQLVale & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4FECVAL <= CVDATE('" & StrFecha & "') AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
    
'        If strNumeroVale <> vbNullString Then
'            strSQLVale = strSQLVale & "AND  "
'        End If
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        If Val(rstVale!Cantidad & "") > 0 Then
            calcularCostoPromedio = Val(Format(Val(rstVale!ValVta & "") / Val(rstVale!Cantidad & ""), "#.0000"))
        End If
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errCalcularCostoPromedio:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: CalcularCostoPromedio"
    
    calcularCostoPromedio = 0
    
    Err.Clear
End Function

Public Function calcularCostoPromedioV2() As Double
    On Error GoTo errCalcularCostoPromedioV2
    
    calcularCostoPromedioV2 = 0
    
    Dim cmdOP As ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
'    With cmdOP
'        .ActiveConnection = cnBdCPlus
'        .CommandType = adCmdStoredProc
'        .CommandText = "Procesos.usp_CalcularCostoPromedioProducto"
'
'        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
'        .Parameters.Append .CreateParameter("@CODIGOMONEDA", adVarChar, adParamInput, 1, strCodMoneda)
'        .Parameters.Append .CreateParameter("@FECHA", adDate, adParamInput, 10, StrFecha)
'        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
'
'        Set rstVale = .Execute()
'    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        If Val(rstVale!Cantidad & "") > 0 Then
            calcularCostoPromedioV2 = Val(Format(Val(rstVale!ValVta & "") / Val(rstVale!Cantidad & ""), "#.00"))
            
            If calcularCostoPromedioV2 < 0 Then calcularCostoPromedioV2 = 0
        End If
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
    Resume
errCalcularCostoPromedioV2:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: CalcularCostoPromedioV2"
    
    calcularCostoPromedioV2 = 0
    
    Err.Clear
End Function

Public Function obtenerUltimoCostoPromedioV2() As Double
    On Error GoTo errObtenerUltimoCostoPromedioV2
    
    obtenerUltimoCostoPromedioV2 = 0
    
    Dim cmdOP As ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ObtenerUltimoCostoPromedioProducto"
        
        .Parameters.Append .CreateParameter("@CODIGOMONEDA", adVarChar, adParamInput, 1, strCodMoneda)
        .Parameters.Append .CreateParameter("@FECHA", adDate, adParamInput, 10, StrFecha)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        obtenerUltimoCostoPromedioV2 = Val(Format(Val(rstVale!costo & ""), "#.00"))
            
        If obtenerUltimoCostoPromedioV2 < 0 Then obtenerUltimoCostoPromedioV2 = 0
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errObtenerUltimoCostoPromedioV2:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: ObtenerUltimoCostoPromedioV2"
    
    obtenerUltimoCostoPromedioV2 = 0
    
    Err.Clear
End Function

Public Function calcularCostoPromedioV3() As Double
    On Error GoTo errCalcularCostoPromedioV3
    
    calcularCostoPromedioV3 = 0
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(DET.TIPO = 'I', 1, -1)) , '#0.0000')) AS CANTIDAD, "
    strSQLVale = strSQLVale & "VAL(FORMAT( SUM((DET." & IIf(strCodMoneda = "S", "F3VALVTA", "F3VALDOL") & " * DET.F3CANPRO) * IIF(DET.TIPO = 'I', 1, -1))  , '#0.0000')) AS VALVTA "
    'strSQLVale = strSQLVale & "(VALVTA/CANTIDAD) AS CP "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F3CANPRO, " & IIf(strCodMoneda = "S", "DET.F3VALVTA", "DET.F3VALDOL") & " "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(TMPUTILVALEDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILVALECAB AS CAB ON (CAB.F2CODALM = DET.F2CODALM) AND (CAB.F4NUMVAL = DET.F4NUMVAL)) "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4FECVAL < CVDATE('" & StrFecha & "') AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL "
    
    strSQLVale = strSQLVale & "UNION ALL "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F3CANPRO, " & IIf(strCodMoneda = "S", "DET.F3VALVTA", "DET.F3VALDOL") & " "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(TMPUTILVALEDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILVALECAB AS CAB ON (CAB.F2CODALM = DET.F2CODALM) AND (CAB.F4NUMVAL = DET.F4NUMVAL)) "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TRIM(ORI.F1COSTO & '') = '*' AND "
    strSQLVale = strSQLVale & "DET.TIPO = 'I' AND "
    strSQLVale = strSQLVale & "CAB.F4FECVAL = CVDATE('" & StrFecha & "') AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL) AS DET "
    
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILVALECAB AS CAB "
    strSQLVale = strSQLVale & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO"
    
    abrirCnTemporal
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        If Val(rstVale!Cantidad & "") > 0 Then
            calcularCostoPromedioV3 = Val(Format(Val(rstVale!ValVta & "") / Val(rstVale!Cantidad & ""), "#.0000"))
            
            If calcularCostoPromedioV3 < 0 Then calcularCostoPromedioV3 = 0
        End If
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errCalcularCostoPromedioV3:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: CalcularCostoPromedioV3"
    'Resume
    calcularCostoPromedioV3 = 0
    
    Err.Clear
End Function

Public Function calcularCostoPromedioV3conDAO() As Double
    On Error GoTo errCalcularCostoPromedioV3conDAO
    
    calcularCostoPromedioV3conDAO = 0
    
    Dim strUltimaCompra As String
    Dim dblUltimoCostoPromedio As Double
    
    strUltimaCompra = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "TOP 1 CAB.F4FECVAL", _
                                                    "PROCESOS.IF3VALES AS DET " & _
                                                    "LEFT JOIN PROCESOS.IF4VALES AS CAB ON " & _
                                                    "CAB.F2CODALM = DET.F2CODALM AND " & _
                                                    "CAB.F4NUMVAL = DET.F4NUMVAL", _
                                                    "DET.F5CODPRO", strCodProducto, "T", _
                                                    "AND CAB.F1CODORI = 'XC0' " & _
                                                    "AND CAB.F4TIPOVALE = 'I' " & _
                                                    "ORDER BY " & _
                                                    "CAB.F4FECVAL DESC, CAB.F4NUMVAL DESC")
    
    If strUltimaCompra <> vbNullString Then
        If CDate(strUltimaCompra) < CDate(StrFecha) Then
            dblUltimoCostoPromedio = Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "TOP 1 DET.F3VALVTA", _
                                                                        "PROCESOS.IF3VALES AS DET " & _
                                                                        "LEFT JOIN PROCESOS.IF4VALES AS CAB ON " & _
                                                                        "CAB.F2CODALM = DET.F2CODALM AND " & _
                                                                        "CAB.F4NUMVAL = DET.F4NUMVAL", _
                                                                        "DET.F5CODPRO", strCodProducto, "T", _
                                                                        "AND CAB.F4TIPOVALE = 'S' " & _
                                                                        "ORDER BY " & _
                                                                        "CAB.F4FECVAL DESC, CAB.F4NUMVAL DESC"))
        ElseIf CDate(strUltimaCompra) = CDate(StrFecha) Then
            dblUltimoCostoPromedio = Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "TOP 1 DET.F3VALVTA", _
                                                                        "PROCESOS.IF3VALES AS DET " & _
                                                                        "LEFT JOIN PROCESOS.IF4VALES AS CAB ON " & _
                                                                        "CAB.F2CODALM = DET.F2CODALM AND " & _
                                                                        "CAB.F4NUMVAL = DET.F4NUMVAL", _
                                                                        "DET.F5CODPRO", strCodProducto, "T", _
                                                                        "AND CAB.F4TIPOVALE = 'S' " & _
                                                                        "AND CVDATE(CAB.F4FECVAL) = CVDATE('" & StrFecha & "') " & _
                                                                        "ORDER BY " & _
                                                                        "CAB.F4FECVAL DESC, CAB.F4NUMVAL DESC"))
        End If
    End If
    
    If dblUltimoCostoPromedio > 0 Then
        calcularCostoPromedioV3conDAO = dblUltimoCostoPromedio
        
        Exit Function
    End If
    
    Dim cnDAO As DAO.Database
    Dim rstValeDetDAO As DAO.Recordset
    
    Dim dblCantidadTotal As Double
    Dim dblImporteTotal As Double
    
    dblCantidadTotal = 0
    dblImporteTotal = 0
    
    
    Set cnDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    Set rstValeDetDAO = cnDAO.OpenRecordset("PROCESOS.IF3VALES")
    
    rstValeDetDAO.Index = "IndiceValeProducto"
    
    rstValeDetDAO.Seek "=", strCodProducto
    
    If Not rstValeDetDAO.NoMatch Then
        
        Do While (CDate(rstValeDetDAO!F4FECVAL) < CDate(StrFecha) And _
                    Trim(rstValeDetDAO!f5codpro & "") = strCodProducto)
            
            With objAyudaVale
                .inicializarEntidades
                
                .CodigoAlmacen = Trim(rstValeDetDAO!f2codalm & "")
                .NumeroVale = Trim(rstValeDetDAO!F4NUMVAL & "")
                
                .obtenerConfigVale
            End With
            
            If objAyudaVale.CodigoOrigen <> "XCS" Then
                dblCantidadTotal = dblCantidadTotal + (Val(rstValeDetDAO!F3CANPRO & "") * IIf(Trim(rstValeDetDAO!Tipo & "") = "I", 1, -1))
                    
                dblImporteTotal = dblImporteTotal + ((IIf(strCodMoneda = "S", Val(rstValeDetDAO!F3VALVTA & ""), Val(rstValeDetDAO!F3VALDOL & "")) * Val(rstValeDetDAO!F3CANPRO & "")) * IIf(Trim(rstValeDetDAO!Tipo & "") = "I", 1, -1))
            End If
            
            rstValeDetDAO.MoveNext
                        
            If rstValeDetDAO.EOF Then Exit Do
            
            If Trim(rstValeDetDAO!f5codpro & "") <> strCodProducto Then Exit Do
        Loop
    End If
    
    Set rstValeDetDAO = cnDAO.OpenRecordset("PROCESOS.IF3VALES")
    
    rstValeDetDAO.Index = "IndiceValeTipoFechaProducto"
    
    rstValeDetDAO.Seek "=", "I", StrFecha, strCodProducto
    
    If Not rstValeDetDAO.NoMatch Then
        
        Do While (Trim(rstValeDetDAO!Tipo & "") = "I" And _
                    Trim(rstValeDetDAO!f5codpro & "") = strCodProducto And _
                    CDate(rstValeDetDAO!F4FECVAL) = CDate(StrFecha))
            
            With objAyudaVale
                .inicializarEntidades
                
                .CodigoAlmacen = Trim(rstValeDetDAO!f2codalm & "")
                .NumeroVale = Trim(rstValeDetDAO!F4NUMVAL & "")
                
                .obtenerConfigVale
            End With
            
            If objAyudaVale.CodigoOrigen <> "XCS" Then
                With objAyudaOrigen
                    .inicializarEntidades
                    
                    .Codigo = objAyudaVale.CodigoOrigen
                    
                    .obtenerConfigOrigen
                End With
                
                If objAyudaOrigen.RegistrarCosto Then
                    dblCantidadTotal = dblCantidad + Val(rstValeDetDAO!F3CANPRO & "")
                        
                    dblImporteTotal = dblImporteTotal + (IIf(strCodMoneda = "S", Val(rstValeDetDAO!F3VALVTA & ""), Val(rstValeDetDAO!F3VALDOL & "")))
                End If
            End If
            
            rstValeDetDAO.MoveNext
                        
            If rstValeDetDAO.EOF Then Exit Do
            
            If Trim(rstValeDetDAO!f5codpro & "") = strCodProducto And _
                CDate(rstValeDetDAO!F4FECVAL) = CDate(StrFecha) Then Exit Do
        Loop
    End If
    
    If dblCantidadTotal > 0 Then
        calcularCostoPromedioV3conDAO = Val(Format(dblImporteTotal / dblCantidadTotal, "#.0000"))
        
        If calcularCostoPromedioV3conDAO < 0 Then calcularCostoPromedioV3conDAO = 0
    End If
    
    rstValeDetDAO.Close
    
    Set rstValeDetDAO = Nothing
    
    cnDAO.Close
    
    Set cnDAO = Nothing
    
    dblCantidadTotal = 0
    dblImporteTotal = 0
    
    Exit Function
errCalcularCostoPromedioV3conDAO:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: CalcularCostoPromedioV3conDAO"
    
    calcularCostoPromedioV3conDAO = 0
    
    Err.Clear
End Function

Public Sub obtenerUltimoPrecioSinIgvProductoDeProveedor()
    On Error GoTo errObtenerUltimoPrecioSinIgvProductoDeProveedor
    
    Dim cmdOP As ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ObtenerUltimoPrecioSinIgvProductoProveedor"
        
        .Parameters.Append .CreateParameter("@RUCPROVEEDOR", adVarChar, adParamInput, 20, strCodProveedor)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        strCodMoneda = Trim(rstVale!F4MONEDA & "")
        
        dblValorVta = Val(rstVale!F3VALVTA & "")
        dblValorVtaDol = Val(rstVale!F3VALDOL & "")
        dblPorcentajeDscto = Val(rstVale!F3PORCENTAJEDSCTO & "")
    Else
        strCodMoneda = "S"
        
        dblValorVta = 0
        dblValorVtaDol = 0
        dblPorcentajeDscto = 0
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Sub
errObtenerUltimoPrecioSinIgvProductoDeProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrden: ObtenerUltimoPrecioSinIgvProductoDeProveedor"
    
    strCodMoneda = "S"
    
    dblValorVta = 0
    dblValorVtaDol = 0
    
    Err.Clear
End Sub

'Listar Insumos Descargados de OP para su Devolucion en Grilla (QuamtumGrid)
Public Sub listarGrillaHistorialPrecioPorProducto(ByVal grilla As dxDBGrid, _
                                                    ByVal listaImagenes As Object, _
                                                    ByVal strCodProducto As String, _
                                                    ByVal strFechaDesde As String, _
                                                    ByVal strFechaHasta As String, _
                                                    ByVal dblPorcentajeImpuesto As Double, _
                                                    ByVal strFiltroSensitivo As String)
    
    On Error GoTo errListarGrillaHistorialPrecioPorProducto
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "PROV.F2NOMPROV, "
    strSQLVale = strSQLVale & "ALM.F2NOMALM, "
    strSQLVale = strSQLVale & "DET.F4NUMVAL, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "CAB.F4MONEDA, "
    strSQLVale = strSQLVale & "PROD.F7CODMED, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD, "
    strSQLVale = strSQLVale & "VAL(FORMAT(IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL), '#0.0000')) AS COSTOUNITARIO, "
    strSQLVale = strSQLVale & "VAL(FORMAT(IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL) * (1 + " & dblPorcentajeImpuesto & "/100), '#0.0000')) AS PRECIOUNITARIO, "
    strSQLVale = strSQLVale & "IIF(ISNULL(DET.F3PORCENTAJEDSCTO), 0, DET.F3PORCENTAJEDSCTO) AS PORCENTAJEDSCTO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "((((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2CODPROV = CAB.F2CODPROV) "
    strSQLVale = strSQLVale & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = DET.F2CODALM) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI = 'XC0' AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaDesde & "') AND CVDATE('" & strFechaHasta & "') "
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND (PROV.F2NOMPROV LIKE '%" & strFiltroSensitivo & "%' OR DET.F4NUMVAL LIKE '%" & strFiltroSensitivo & "%') "
        End If
        
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "PROV.F2NOMPROV, "
    strSQLVale = strSQLVale & "ALM.F2NOMALM, "
    strSQLVale = strSQLVale & "DET.F4NUMVAL, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "CAB.F4MONEDA, "
    strSQLVale = strSQLVale & "PROD.F7CODMED, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "VAL(FORMAT(IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL), '#0.0000')), "
    strSQLVale = strSQLVale & "VAL(FORMAT(IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL) * (1 + " & dblPorcentajeImpuesto & "/100), '#0.0000')), "
    strSQLVale = strSQLVale & "IIF(ISNULL(DET.F3PORCENTAJEDSCTO), 0, DET.F3PORCENTAJEDSCTO) "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL DESC, "
    strSQLVale = strSQLVale & "DET.F4NUMVAL DESC"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Nombre de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Producto"
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Proveedor"
                .DisableEditor = True
                .FieldName = "F2NOMPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Almacen"
                .DisableEditor = True
                .FieldName = "F2NOMALM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAlmacen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Numero de Vale
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Numero de Vale"
                .Color = &HC000C0
                .DisableEditor = True
                .FieldName = "F4NUMVAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNumeroVale"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Fecha de Vale
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Fecha"
                .DisableEditor = True
                .FieldName = "F4FECVAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecha"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Moneda
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Moneda"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F4MONEDA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColMoneda"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 45
            End With
            
            'Columna Codigo UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F7CODMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigoUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = False
            End With
            
            'Columna UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
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
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Costo S/Igv
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Costo S/Igv"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COSTOUNITARIO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColPrecioSinIgv"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Costo C/Igv
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Costo C/Igv"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PRECIOUNITARIO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColPrecioConIgv"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Porcentaje
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "% Dscto"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PORCENTAJEDSCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColPorcentajeDscto"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F4NUMVAL"
            
            .Columns.ColumnByFieldName("F5NOMPRO").GroupIndex = 0
            .Columns.ColumnByFieldName("F2NOMPROV").GroupIndex = 1
            .Columns.ColumnByFieldName("F2NOMALM").GroupIndex = 2
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaHistorialPrecioPorProducto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaHistorialPrecioPorProducto"
    End Select
    
    Err.Clear
End Sub

Public Sub verificarStockProducto(Optional ByVal strNroPedido As String)
    On Error GoTo errVerificarStockProducto
    
    Dim cmdOP As ADODB.Command
    
    Set cmdOP = New ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_VerificarStockProducto"
        
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@CODIGOALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strNroPedido)
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        strCodProducto = Trim(rstVale!f5codpro & "")
        
        dblCompromisoEAG = Val(rstVale!CompromisoEAG & "")
        dblCompromisoPLG = Val(rstVale!CompromisoPLG & "")
        dblLibreEAG = Val(rstVale!LibreEAG & "")
        dblLibrePLG = Val(rstVale!LibrePLG & "")
        dblStockEAG = Val(rstVale!StockEAG & "")
        dblStockPLG = Val(rstVale!StockPLG & "")
    Else
        strCodProducto = vbNullString
        
        dblCompromisoEAG = 0
        dblCompromisoPLG = 0
        dblLibreEAG = 0
        dblLibrePLG = 0
        dblStockEAG = 0
        dblStockPLG = 0
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Sub
errVerificarStockProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarStockProducto"
    
    Err.Clear
End Sub

Public Sub descargarStockProductoReqOP(ByVal strTipoStockCEAorCPLorLEAorLPL As String, _
                                        Optional ByVal strNroPedido As String, _
                                        Optional ByVal strNombreTablaTemporalSQL As String)
    
    On Error GoTo errDescargarStockProductoReqOP
    
    Dim cmdOP As ADODB.Command
    Dim rsStock As ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    Set rsStock = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarResumenAtencionPorProductoResOP"
        
        .Parameters.Append .CreateParameter("@TIPOSTOCK", adVarChar, adParamInput, 3, strTipoStockCEAorCPLorLEAorLPL)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, strNroPedido)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaTemporalSQL)
        
        Set rsStock = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rsStock.EOF Then
        Do While Not rsStock.EOF
            abrirCnTemporal
            
            strSQLVale = vbNullString
            strSQLVale = strSQLVale & "INSERT INTO TMPUTILRESUMENSTOCKREQUERIMIENTOOP "
            strSQLVale = strSQLVale & "VALUES("
            strSQLVale = strSQLVale & "'" & Trim(rsStock!Tipo & "") & "', "
            strSQLVale = strSQLVale & "'" & Trim(rsStock!CodProducto & "") & "', "
            strSQLVale = strSQLVale & Val(rsStock!Cantidad & "")
            strSQLVale = strSQLVale & ")"
            
            cnDBTemp.Execute strSQLVale
            
            rsStock.MoveNext
        Loop
    End If
    
    If rsStock.State = 1 Then rsStock.Close
    
    Set rsStock = Nothing
    
    strSQLVale = vbNullString
        
    Exit Sub
errDescargarStockProductoReqOP:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DescargarStockProductoReqOP"
    
    Err.Clear
End Sub

Public Sub descargarStockProductoReqOPv2(ByVal strTipoStockCEAorCPLorLEAorLPL As String, _
                                            Optional ByVal strNroPedido As String)
    
    On Error GoTo errDescargarStockProductoReqOPv2
    
    Dim intAnno As Integer, intMes As Integer, strFechaInicial As String
    
    'Obtener Periodo de Ultimo Cierre
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT TOP 1 ANNO, MES FROM IF3CIERREMENSUAL GROUP BY ANNO, MES ORDER BY ANNO DESC, MES DESC"
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        intAnno = Val(rstVale!Anno & "")
        intMes = Val(rstVale!mes & "")
    Else
        intAnno = IIf(Month(Date) > 1, Year(Date), Year(Date) - 1)
        intMes = IIf(Month(Date) > 1, Month(Date) - 1, 12)
    End If
    
    strFechaInicial = "01/" & IIf(Val(intMes) < 12, Format(Val(intMes + 1), "00"), "01") & "/" & IIf(Val(intMes) < 12, intAnno, Val(intAnno) + 1)
    
    
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "INSERT INTO TMPUTILRESUMENSTOCKREQUERIMIENTOOP("
    strSQLVale = strSQLVale & "TIPO, CODPRODUCTO, CANTIDAD"
    strSQLVale = strSQLVale & ") "
    
    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "'" & strTipoStockCEAorCPLorLEAorLPL & "' AS TIPO, "
    
    Select Case strTipoStockCEAorCPLorLEAorLPL
        Case "CEA"
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES "
            strSQLVale = strSQLVale & "WHERE "
                
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "COD_SOLICITUD = '" & strNroPedido & "'  AND "
                Else
                    strSQLVale = strSQLVale & "COD_SOLICITUD <> ''  AND "
                End If
                
            strSQLVale = strSQLVale & "F5CODPRO = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "F5CODPRO, F2CODALM"
        Case "CPL"
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "(IF3ORDEN AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
            strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
            
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' AND "
                Else
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
                End If
                
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL) AS INGRESOS "
            strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
                
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' AND "
                Else
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
                End If
                
            strSQLVale = strSQLVale & "DET.F3CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F3CODPRO"
        Case "LEA"
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "COD_SOLICITUD = ''  AND "
            strSQLVale = strSQLVale & "F5CODPRO = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "F2CODALM"
        Case "LPL"
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "(IF3ORDEN AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
            strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
            strSQLVale = strSQLVale & ") AS INGRESOS "
            strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
            strSQLVale = strSQLVale & "DET.F3CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F3CODPRO"
    End Select
    
    cnBdCPlus.Execute strSQLVale
    
    strSQLVale = vbNullString
    
    Exit Sub
errDescargarStockProductoReqOPv2:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DescargarStockProductoReqOPv2"
    'Resume
    Err.Clear
End Sub

Public Sub descargarStockProductoCompAfectado(ByVal strTipoStockCEAorCPLorLEAorLPL As String, _
                                                Optional ByVal strNroPedido As String)
    
    On Error GoTo errDescargarStockProductoCompAfectado
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "INSERT INTO TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO("
    strSQLVale = strSQLVale & "TIPO, CODPRODUCTO, CANTIDAD"
    strSQLVale = strSQLVale & ") "
    
    strSQLVale = strSQLVale & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "'" & strTipoStockCEAorCPLorLEAorLPL & "' AS TIPO, "
    
    Select Case strTipoStockCEAorCPLorLEAorLPL
        Case "CEA"
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES "
            strSQLVale = strSQLVale & "WHERE "
                
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "COD_SOLICITUD = '" & strNroPedido & "'  AND "
                Else
                    strSQLVale = strSQLVale & "COD_SOLICITUD <> ''  AND "
                End If
                
            strSQLVale = strSQLVale & "F5CODPRO = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "F5CODPRO, F2CODALM"
        Case "CPL"
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "(IF3ORDEN AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
            strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
            
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' AND "
                Else
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
                End If
                
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL) AS INGRESOS "
            strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
                
                If strNroPedido <> vbNullString Then
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' AND "
                Else
                    strSQLVale = strSQLVale & "DET.COD_SOLICITUD <> '' AND "
                End If
                
            strSQLVale = strSQLVale & "DET.F3CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F3CODPRO"
        Case "LEA"
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "COD_SOLICITUD = ''  AND "
            strSQLVale = strSQLVale & "F5CODPRO = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "F5CODPRO, "
            strSQLVale = strSQLVale & "F2CODALM"
        Case "LPL"
            strSQLVale = strSQLVale & "DET.F3CODPRO, "
            strSQLVale = strSQLVale & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "(IF3ORDEN AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            strSQLVale = strSQLVale & "LEFT JOIN "
            strSQLVale = strSQLVale & "(SELECT "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL, "
            strSQLVale = strSQLVale & "VAL(FORMAT( SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "CAB.F1CODORI IN ('XC0') AND "
            strSQLVale = strSQLVale & "DET.F4NUMORD <> '' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL = '" & strCodProducto & "' "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F4NUMORD, "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
            strSQLVale = strSQLVale & "DET.F5CODPROORIGINAL"
            strSQLVale = strSQLVale & ") AS INGRESOS "
            strSQLVale = strSQLVale & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F4LOCAL = 'OC' AND "
            strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '' AND "
            strSQLVale = strSQLVale & "DET.F3CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
            strSQLVale = strSQLVale & "GROUP BY "
            strSQLVale = strSQLVale & "DET.F3CODPRO"
    End Select
    
    cnBdCPlus.Execute strSQLVale
    
    strSQLVale = vbNullString
    
    Exit Sub
errDescargarStockProductoCompAfectado:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: DescargarStockProductoCompAfectado"
    'Resume
    Err.Clear
End Sub

Public Function verificarProductoPorProveedor() As Boolean
    On Error GoTo errVerificarProductoPorProveedor
    
    Dim cmdOP As ADODB.Command
    
    Set rstVale = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_VerificarProductoPorProveedor"
        
        .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, strCodProducto)
        .Parameters.Append .CreateParameter("@CODPROVEEDOR", adVarChar, adParamInput, 10, strCodProveedor)
        
        Set rstVale = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstVale.EOF Then
        verificarProductoPorProveedor = True
    Else
        verificarProductoPorProveedor = False
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    Exit Function
errVerificarProductoPorProveedor:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarProductoPorProveedor"
    
    verificarProductoPorProveedor = False
    
    Err.Clear
End Function

Public Function verificarStockProductoFisicoCorL(ByVal strCodProducto As String, _
                                                    ByVal strCodAlmacen As String, _
                                                    ByVal strNroOC As String, _
                                                    ByVal strNroPedido As String, _
                                                    ByVal dblCantidadAusar As Double, _
                                                    ByVal strFechaCorte As String) As Boolean
                                                    
    On Error GoTo errVerificarStockProductoFisicoCorL
    
    Set rstVale = New ADODB.Recordset
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM, "
    'strSQLVale = strSQLVale & "DET.F4NUMORD AS OC, "
    'strSQLVale = strSQLVale & "DET.COD_SOLICITUD AS PEDIDO, "
    strSQLVale = strSQLVale & "VAL(FORMAT(SUM(DET.F3CANPRO  * IIF(TIPO = 'S', -1, 1)), '#.0000')) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CVDATE(DET.F4FECVAL) <= CVDATE('" & strFechaCorte & "') AND "
    strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
    strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' "
    'strSQLVale = strSQLVale & "DET.F4NUMORD = '" & strNroOC & "' AND "
    'strSQLVale = strSQLVale & "DET.COD_SOLICITUD = '" & strNroPedido & "' "
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "DET.F2CODALM "
    'strSQLVale = strSQLVale & "DET.F4NUMORD, "
    'strSQLVale = strSQLVale & "DET.COD_SOLICITUD "
    strSQLVale = strSQLVale & "HAVING "
    strSQLVale = strSQLVale & "VAL(FORMAT(SUM(DET.F3CANPRO  * IIF(TIPO = 'S', -1, 1)), '#.0000')) >= " & dblCantidadAusar
    
    If rstVale.State = 1 Then rstVale.Close
    
    rstVale.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVale.EOF Then
        verificarStockProductoFisicoCorL = True
    Else
        verificarStockProductoFisicoCorL = False
    End If
    
    If rstVale.State = 1 Then rstVale.Close
    
    Set rstVale = Nothing
    
    strSQLVale = vbNullString
    
    Exit Function
errVerificarStockProductoFisicoCorL:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarStockProductoFisicoCorL"
    'Resume
    verificarStockProductoFisicoCorL = False
    
    Err.Clear
End Function

Public Function verificarStockProductoFisicoCorLconDAO(ByVal strCodProducto As String, _
                                                        ByVal strCodAlmacen As String, _
                                                        ByVal strNroOC As String, _
                                                        ByVal strNroPedido As String, _
                                                        ByVal dblCantidadAusar As Double, _
                                                        ByVal strFechaCorte As String) As Boolean
    
    On Error GoTo errVerificarStockProductoFisicoCorLconDAO
    
    'Set rstVale = New ADODB.Recordset
    
    Dim cnDAO As DAO.Database
    Dim rstValeDetDAO As DAO.Recordset
    
    Dim dblCantidadTotal As Double
    
    dblCantidadTotal = 0
    
    Set cnDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    Set rstValeDetDAO = cnDAO.OpenRecordset("PROCESOS.IF3VALES")
    
    rstValeDetDAO.Index = "IndiceValeProductoAlmacen"
    
    rstValeDetDAO.Seek "=", strCodProducto, strCodAlmacen
    
    If Not rstValeDetDAO.NoMatch Then
        
        Do While (CDate(rstValeDetDAO!F4FECVAL) <= CDate(strFechaCorte) And _
                    Trim(rstValeDetDAO!f5codpro & "") = strCodProducto And _
                    Trim(rstValeDetDAO!f2codalm & "") = strCodAlmacen)
            
            With objAyudaVale
                .inicializarEntidades
                
                .CodigoAlmacen = Trim(rstValeDetDAO!f2codalm & "")
                .NumeroVale = Trim(rstValeDetDAO!F4NUMVAL & "")
                
                .obtenerConfigVale
            End With
            
            If objAyudaVale.CodigoOrigen <> "XCS" Then
                dblCantidadTotal = dblCantidadTotal + (Val(rstValeDetDAO!F3CANPRO & "") * IIf(Trim(rstValeDetDAO!Tipo & "") = "I", 1, -1))
            End If
            
            rstValeDetDAO.MoveNext
            
            If rstValeDetDAO.EOF Then Exit Do
            
            If Trim(rstValeDetDAO!f5codpro & "") <> strCodProducto Or _
                    Trim(rstValeDetDAO!f2codalm & "") <> strCodAlmacen Then Exit Do
        Loop
    End If
    
    If dblCantidadTotal >= dblCantidadAusar Then
        verificarStockProductoFisicoCorLconDAO = True
    Else
        verificarStockProductoFisicoCorLconDAO = False
    End If
    
    rstValeDetDAO.Close
    
    Set rstValeDetDAO = Nothing
    
    cnDAO.Close
    
    Set cnDAO = Nothing
    
    dblCantidadTotal = 0
    
    Exit Function
errVerificarStockProductoFisicoCorLconDAO:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: VerificarStockProductoFisicoCorLconDAO"
    'Resume
    verificarStockProductoFisicoCorLconDAO = False
    
    Err.Clear
End Function

Public Sub recalcularCostoPromedioProducto(ByVal strCodAlmacenEspecifico As String, _
                                            ByVal strFechaInicial As String, _
                                            ByVal strFechaFin As String, _
                                            ByVal strCodMonedaParaRecalculo As String, _
                                            ByVal strCodProductoEspecifico As String, _
                                            ByVal etiquetaProceso1 As Object, _
                                            ByVal barraProceso1 As Object)
    
    On Error GoTo errRecalcularCostoPromedioProducto
    
    Dim rsProceso1 As ADODB.Recordset
    Dim cmdOP As ADODB.Command
    Dim dblCantidadRegistro As Double
    
    Set rsProceso1 = New ADODB.Recordset
    
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ObtenerRegistrosParaCostoPromedioProducto"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacenEspecifico)
        .Parameters.Append .CreateParameter("@FECHAINICIAL", adDate, adParamInput, 10, strFechaInicial)
        .Parameters.Append .CreateParameter("@FECHAFINAL", adDate, adParamInput, 10, strFechaFin)
        .Parameters.Append .CreateParameter("@CODIGOPRODUCTO", adVarChar, adParamInput, 50, strCodProductoEspecifico)
        .Parameters.Append .CreateParameter("@CANTIDADFILAS", adInteger, adParamOutput, 5)
        
        .Execute
        
        dblCantidadRegistro = Val(.Parameters("@CANTIDADFILAS") & "")
        
        Set rsProceso1 = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rsProceso1.EOF Then
        barraProceso1.Max = dblCantidadRegistro 'ModUtilitario.devuelveCantRegistros(rsProceso1)
        barraProceso1.Value = 0
        etiquetaProceso1.Caption = "Recalculo de Costo Promedio de Producto(s)..."
        
        Do While Not rsProceso1.EOF
            strCodAlmacen = Trim(rsProceso1!f2codalm & "")
            strCodMoneda = strCodMonedaParaRecalculo
            strCodProducto = Trim(rsProceso1!f5codpro & "")
            StrFecha = Format(Trim(rsProceso1!F4FECVAL & ""), "Short Date")
            
            dblValorVta = Me.calcularCostoPromedioV2
            
            If dblValorVta <= 0 Then
                strCodMoneda = strCodMonedaParaRecalculo
                strCodProducto = Trim(rsProceso1!f5codpro & "")
                StrFecha = Format(Trim(rsProceso1!F4FECVAL & ""), "Short Date")
                
                dblValorVta = Me.obtenerUltimoCostoPromedioV2
            End If
            
            dblTCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", StrFecha, "F"))
            
            If dblTCambio = 0 Then dblTCambio = 2.8
            
            strSQLVale = vbNullString
            strSQLVale = strSQLVale & "UPDATE DET "
            
            strSQLVale = strSQLVale & "SET "
                
                Select Case strCodMonedaParaRecalculo
                    Case "S"
                        strSQLVale = strSQLVale & "DET.F3VALVTA = " & dblValorVta & ", "
                        strSQLVale = strSQLVale & "DET.F3IGV = 0, "
                        strSQLVale = strSQLVale & "DET.F3TOTITE = (DET.F3CANPRO * " & dblValorVta & "), "
                        strSQLVale = strSQLVale & "DET.F3VALDOL = " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ", "
                        strSQLVale = strSQLVale & "DET.F3IGVDOL = 0, "
                        strSQLVale = strSQLVale & "DET.F3TOTDOL = (DET.F3CANPRO * " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ") "
                    Case Else
                        strSQLVale = strSQLVale & "DET.F3VALDOL = " & dblValorVta & ", "
                        strSQLVale = strSQLVale & "DET.F3IGVDOL = 0, "
                        strSQLVale = strSQLVale & "DET.F3TOTDOL = (DET.F3CANPRO * " & dblValorVta & "), "
                        strSQLVale = strSQLVale & "DET.F3VALVTA = " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ", "
                        strSQLVale = strSQLVale & "DET.F3IGV = 0, "
                        strSQLVale = strSQLVale & "DET.F3TOTITE = (DET.F3CANPRO * " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ") "
                End Select
            
            strSQLVale = strSQLVale & "FROM "
            strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
            strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
                
            strSQLVale = strSQLVale & "WHERE "
            strSQLVale = strSQLVale & "DET.F2CODALM = '" & strCodAlmacen & "' AND "
            strSQLVale = strSQLVale & "ORI.F1COSTO = '' AND "
            strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
            strSQLVale = strSQLVale & "CAB.F4FECVAL = '" & StrFecha & "' AND "
            strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
            
            cnBdCPlus.Execute strSQLVale, dblCantidadMaxima
            
            Actualiza_Log strSQLVale, StrConexDbBancos
            
            dblItem = Val(rsProceso1!Cantidad & "")
            
            'MsgBox "Cantidad de Registros al '" & StrFecha & "' = " & dblItem & vbNewLine & "Cantidad de Registros Actualizados al '" & StrFecha & "' = " & dblCantidadMaxima, vbInformation + vbOKOnly, App.ProductName
            
            DoEvents
            
            barraProceso1.Value = barraProceso1.Value + 1
            etiquetaProceso1.Caption = "Recalculo de Costo Promedio de: " & left(Trim(rsProceso1!F5NOMPRO & ""), 100) & " (" & Trim(rsProceso1!F7SIGMED & "") & ") - [" & Trim(rsProceso1!F4FECVAL & "") & "] " & "... " & FormatPercent(barraProceso1.Value / barraProceso1.Max, 3)
            
            rsProceso1.MoveNext
        Loop
            MsgBox "Evaluación y recalculo finalizado.", vbInformation + vbOKOnly, App.ProductName
    Else
'        strCodMoneda = strCodMonedaParaRecalculo
'        strCodProducto = strCodProductoEspecifico
'        StrFecha = Format(strFechaInicial, "Short Date")
'
'        dblValorVta = Me.obtenerUltimoCostoPromedioV2
'
'        dblTCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", StrFecha, "F"))
'
'        If dblTCambio = 0 Then dblTCambio = 2.8
'
'        strSQLVale = vbNullString
'        strSQLVale = strSQLVale & "UPDATE DET "
'
'        strSQLVale = strSQLVale & "SET "
'
'            Select Case strCodMonedaParaRecalculo
'                Case "S"
'                    strSQLVale = strSQLVale & "DET.F3VALVTA = " & dblValorVta & ", "
'                    strSQLVale = strSQLVale & "DET.F3IGV = 0, "
'                    strSQLVale = strSQLVale & "DET.F3TOTITE = (DET.F3CANPRO * " & dblValorVta & "), "
'                    strSQLVale = strSQLVale & "DET.F3VALDOL = " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ", "
'                    strSQLVale = strSQLVale & "DET.F3IGVDOL = 0, "
'                    strSQLVale = strSQLVale & "DET.F3TOTDOL = (DET.F3CANPRO * " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ") "
'                Case Else
'                    strSQLVale = strSQLVale & "DET.F3VALDOL = " & dblValorVta & ", "
'                    strSQLVale = strSQLVale & "DET.F3IGVDOL = 0, "
'                    strSQLVale = strSQLVale & "DET.F3TOTDOL = (DET.F3CANPRO * " & dblValorVta & "), "
'                    strSQLVale = strSQLVale & "DET.F3VALVTA = " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ", "
'                    strSQLVale = strSQLVale & "DET.F3IGV = 0, "
'                    strSQLVale = strSQLVale & "DET.F3TOTITE = (DET.F3CANPRO * " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ") "
'            End Select
'
'        strSQLVale = strSQLVale & "FROM "
'        strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
'        strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'        strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
'
'        strSQLVale = strSQLVale & "WHERE "
'        strSQLVale = strSQLVale & "ORI.F1COSTO = '' AND "
'        strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
'        strSQLVale = strSQLVale & "CAB.F4FECVAL = '" & StrFecha & "' AND "
'        strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
'
'        cnBdCPlus.Execute strSQLVale, dblCantidadMaxima
'
'        Actualiza_Log strSQLVale, StrConexDbBancos
        
        MsgBox "No se ubico movimiento de almacen para el recalculo al '" & strFechaInicial & "'.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    If rsProceso1.State = 1 Then rsProceso1.Close
    
    Set rsProceso1 = Nothing
    
    strSQLVale = vbNullString
    
    Exit Sub
errRecalcularCostoPromedioProducto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: RecalcularCostoPromedioProducto"
    
    Actualiza_Log "SqlClsVale: RecalcularCostoPromedioProducto. / No. Error: " & Err.Number & " / " & "Descripción: " & Err.Description, StrConexDbBancos
    
    Err.Clear
End Sub

Public Sub recalcularCostoPromedioProductoV2(ByVal strFechaInicial As String, _
                                                ByVal strCodMonedaParaRecalculo As String, _
                                                ByVal strCodProductoEspecifico As String, _
                                                ByVal etiquetaProceso1 As Object, _
                                                ByVal barraProceso1 As Object)
    
    On Error GoTo errRecalcularCostoPromedioProductoV2
    
    Dim rsProceso1 As New ADODB.Recordset
    Dim strNomProductoUm As String
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "COUNT(DET.F4FECVAL) AS CANTIDAD "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(TMPUTILVALEDET AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILVALECAB AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM)) "
    strSQLVale = strSQLVale & "LEFT JOIN TMPUTILORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TRIM(ORI.F1COSTO & '') = '' AND "
    strSQLVale = strSQLVale & "CAB.F4FECVAL BETWEEN CVDATE('" & strFechaInicial & "') AND CVDATE('" & Format(Date, "Short Date") & "') AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
        
        If strCodProductoEspecifico <> vbNullString Then
            strSQLVale = strSQLVale & "AND DET.F5CODPRO = '" & strCodProductoEspecifico & "' "
        End If
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL"
    
    abrirCnTemporal
    
    If rsProceso1.State = 1 Then rsProceso1.Close
    
    rsProceso1.Open strSQLVale, cnDBTemp, adOpenDynamic, adLockReadOnly
    
    If Not rsProceso1.EOF Then
        rsProceso1.MoveFirst
        
        barraProceso1.Max = ModUtilitario.devuelveCantRegistros(rsProceso1)
        barraProceso1.Value = 0
        etiquetaProceso1.Caption = "Recalculo de Producto(s)..."
        
        Do While Not rsProceso1.EOF
            strCodMoneda = strCodMonedaParaRecalculo
            strCodProducto = Trim(rsProceso1!f5codpro & "")
            StrFecha = Trim(rsProceso1!F4FECVAL & "")
            
            dblValorVta = Me.calcularCostoPromedioV3
            
            strNomProductoUm = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "LEFT(PROD.F5NOMPRO, 50) & ' ( ' & MED.F7SIGMED & ' )'", "MAESTROS.IF5PLA AS PROD LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED", "PROD.F5CODPRO", strCodProducto, "T")
            
            If dblValorVta > 0 Then
                dblTCambio = Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "CAMBIO", "CAMBIOS", "FECHA", StrFecha, "F"))
            
                If dblTCambio = 0 Then dblTCambio = 2.8
                
                strSQLVale = vbNullString
                strSQLVale = strSQLVale & "UPDATE "
                strSQLVale = strSQLVale & "(PROCESOS.IF3VALES AS DET "
                strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM)) "
                strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
                strSQLVale = strSQLVale & "SET "
                    
                    Select Case strCodMonedaParaRecalculo
                        Case "S"
                            strSQLVale = strSQLVale & "F3VALVTA = " & dblValorVta & ", "
                            strSQLVale = strSQLVale & "F3IGV = 0, "
                            strSQLVale = strSQLVale & "F3TOTITE = (F3CANPRO * " & dblValorVta & "), "
                            strSQLVale = strSQLVale & "F3VALDOL = " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ", "
                            strSQLVale = strSQLVale & "F3IGVDOL = 0, "
                            strSQLVale = strSQLVale & "F3TOTDOL = (F3CANPRO * " & Val(Format(dblValorVta / dblTCambio, "#.0000")) & ") "
                        Case Else
                            strSQLVale = strSQLVale & "F3VALDOL = " & dblValorVta & ", "
                            strSQLVale = strSQLVale & "F3IGVDOL = 0, "
                            strSQLVale = strSQLVale & "F3TOTDOL = (F3CANPRO * " & dblValorVta & "), "
                            strSQLVale = strSQLVale & "F3VALVTA = " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ", "
                            strSQLVale = strSQLVale & "F3IGV = 0, "
                            strSQLVale = strSQLVale & "F3TOTITE = (F3CANPRO * " & Val(Format(dblValorVta * dblTCambio, "#.0000")) & ") "
                    End Select
                    
                strSQLVale = strSQLVale & "WHERE "
                strSQLVale = strSQLVale & "TRIM(ORI.F1COSTO & '') = '' AND "
                strSQLVale = strSQLVale & "DET.F5CODPRO = '" & strCodProducto & "' AND "
                strSQLVale = strSQLVale & "CAB.F4FECVAL = CVDATE('" & StrFecha & "') AND "
                strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') "
                
                cnBdCPlus.Execute strSQLVale, dblCantidadMaxima
                
                Actualiza_Log strSQLVale, strCadenaConexioBdCPlus
                
                dblItem = Val(rsProceso1!Cantidad & "")
                
                'MsgBox "Cantidad de Registros al '" & StrFecha & "' = " & dblItem & vbNewLine & "Cantidad de Registros Actualizados al '" & StrFecha & "' = " & dblCantidadMaxima, vbInformation + vbOKOnly, App.ProductName
            End If
            
            DoEvents
            
            barraProceso1.Value = barraProceso1.Value + 1
            etiquetaProceso1.Caption = "Recalculo de: " & strNomProductoUm & " - [" & Trim(rsProceso1!F4FECVAL & "") & "] " & "... " & FormatPercent(barraProceso1.Value / barraProceso1.Max, 3)
            
            rsProceso1.MoveNext
        Loop
            MsgBox "Evaluación y recalculo finalizado.", vbInformation + vbOKOnly, App.ProductName
    Else
        MsgBox "No se ubico movimiento de almacen para el recalculo al '" & strFechaInicial & "'.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    If rsProceso1.State = 1 Then rsProceso1.Close
    
    Set rsProceso1 = Nothing
    
    strSQLVale = vbNullString
    
    Exit Sub
errRecalcularCostoPromedioProductoV2:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsVale: RecalcularCostoPromedioProductoV2"
    
    Err.Clear
End Sub

'Listar Inventario de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaInventarioProducto(ByVal grilla As dxDBGrid, _
                                            ByVal strFiltroSensitivo As String, _
                                            ByVal strFechaCorte As String, _
                                            ByVal bolValorizado As Boolean, _
                                            ByVal strCodMonedaValor As String, _
                                            ByVal bolAgrupado As Boolean, _
                                            Optional ByVal strCodAlmacen As String, _
                                            Optional ByVal strCodFamilia As String, _
                                            Optional ByVal strCodSubFamilia As String, _
                                            Optional ByVal listaImagenes As Object, _
                                            Optional ByVal bolMostrarStockNegativo As Boolean, _
                                            Optional ByVal strFiltrarPorProveedor As String)
    
    On Error GoTo errListarGrillaInventarioProducto
    
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    Dim strFechaFinMesAnterior As String
    
    intAnnoCorte = Year(CDate(strFechaCorte)) - IIf(Month(CDate(strFechaCorte)) > 1, 0, 1)
    intMesCorte = IIf(Month(CDate(strFechaCorte)) > 1, Month(CDate(strFechaCorte)) - 1, 12)
    
    strFechaFinMesAnterior = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.FAMILIA, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.SUBFAMILIA, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.F5CODPRO, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.F5NOMPRO, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.F7SIGMED, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.SALDO "
    
        If bolValorizado Then
            strSQLVale = strSQLVale & ", INVENTARIOACTUAL.VALOR, "
            strSQLVale = strSQLVale & "INVENTARIOACTUAL.COSTOPROMEDIO, "
            strSQLVale = strSQLVale & "INVENTARIOACTUAL.TOTAL, "
            strSQLVale = strSQLVale & "INVENTARIOANTERIOR.TOTAL AS TOTALANTERIOR, "
            strSQLVale = strSQLVale & "IIF((INVENTARIOACTUAL.TOTAL - INVENTARIOANTERIOR.TOTAL) > 0, (INVENTARIOACTUAL.TOTAL - INVENTARIOANTERIOR.TOTAL), 0)  AS DIFERENCIA "
        End If
        
    strSQLVale = strSQLVale & "FROM "
    
    'INVENTARIO ACTUAL
    strSQLVale = strSQLVale & "("
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
    strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "VAL(FORMAT(  SUM(VAL(FORMAT(DET.F3CANPRO, '#.0000')) * IIF(DET.TIPO = 'I', 1, -1))  , '#.0000')) AS SALDO "
    
        If bolValorizado Then
            strSQLVale = strSQLVale & ", SUM((DET." & IIf(strCodMonedaValor = "S", "F3VALVTA", "F3VALDOL") & " * VAL(FORMAT(DET.F3CANPRO, '#.0000'))) * IIF(DET.TIPO = 'I', 1, -1)) AS VALOR, "
            strSQLVale = strSQLVale & "IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 0, (VALOR / IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 1, VAL(FORMAT(SALDO, '#.0000'))) ) ) AS COSTOPROMEDIO, "
            strSQLVale = strSQLVale & "(SALDO * COSTOPROMEDIO) AS TOTAL "
        End If
    
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(((("
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, DET.F3CANPRO, DET.F3VALVTA, DET.F3VALDOL "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM) "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4FECVAL <= CVDATE('" & strFechaCorte & "') "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "AND CAB.F2CODALM = '" & strCodAlmacen & "' "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL) AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB "
    strSQLVale = strSQLVale & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO =  DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') AND "
    strSQLVale = strSQLVale & "PROD.F5DESCONTINUADO = 'N' "
        
        If strCodFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND FAM.F7CODCON = '" & strCodFamilia & "' "
        End If
        
        If strCodSubFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND SFAM.F7CODCON = '" & strCodSubFamilia & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND ("
            'strSQLVale = strSQLVale & "PROD.F5CODPRO LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLVale = strSQLVale & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' "
            'strSQLVale = strSQLVale & "MED.F7SIGMED  LIKE '%" & strFiltroSensitivo & "%'"
            strSQLVale = strSQLVale & ") "
        End If
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "FAM.F7DESCON, "
    strSQLVale = strSQLVale & "SFAM.F7DESCON, "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED"
    strSQLVale = strSQLVale & ") AS INVENTARIOACTUAL "
    
    'INVENTARIO ANTERIOR
    strSQLVale = strSQLVale & "LEFT JOIN "
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "SUM(VAL(FORMAT(DET.F3CANPRO, '#.0000')) * IIF(DET.TIPO = 'I', 1, -1)) AS SALDO "
    
        If bolValorizado Then
            strSQLVale = strSQLVale & ", SUM((DET." & IIf(strCodMonedaValor = "S", "F3VALVTA", "F3VALDOL") & " * VAL(FORMAT(DET.F3CANPRO, '#.0000'))) * IIF(DET.TIPO = 'I', 1, -1)) AS VALOR, "
            strSQLVale = strSQLVale & "IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 0, (VALOR / IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 1, VAL(FORMAT(SALDO, '#.0000'))) ) ) AS COSTOPROMEDIO, "
            strSQLVale = strSQLVale & "(SALDO * COSTOPROMEDIO) AS TOTAL "
        End If
    
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(((("
    strSQLVale = strSQLVale & "(SELECT "
    strSQLVale = strSQLVale & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, DET.F3CANPRO, DET.F3VALVTA, DET.F3VALDOL "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM) "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4FECVAL <= CVDATE('" & strFechaFinMesAnterior & "') "
    
        If strCodAlmacen <> vbNullString Then
            strSQLVale = strSQLVale & "AND CAB.F2CODALM = '" & strCodAlmacen & "' "
        End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL) AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB "
    strSQLVale = strSQLVale & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO =  DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F1CODORI NOT IN ('XCS') AND "
    strSQLVale = strSQLVale & "PROD.F5DESCONTINUADO = 'N' "
        
        If strCodFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND FAM.F7CODCON = '" & strCodFamilia & "' "
        End If
        
        If strCodSubFamilia <> vbNullString Then
            strSQLVale = strSQLVale & "AND SFAM.F7CODCON = '" & strCodSubFamilia & "' "
        End If
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND ("
            'strSQLVale = strSQLVale & "PROD.F5CODPRO LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLVale = strSQLVale & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' "
            'strSQLVale = strSQLVale & "MED.F7SIGMED  LIKE '%" & strFiltroSensitivo & "%'"
            strSQLVale = strSQLVale & ") "
        End If
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "DET.F5CODPRO) AS INVENTARIOANTERIOR "
    strSQLVale = strSQLVale & "ON INVENTARIOANTERIOR.F5CODPRO = INVENTARIOACTUAL.F5CODPRO "
    
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "TRIM(INVENTARIOACTUAL.F5CODPRO & '') <> '' "
    
    If bolMostrarStockNegativo Then
        strSQLVale = strSQLVale & "AND INVENTARIOACTUAL.SALDO < 0 "
    End If
    
    If strFiltrarPorProveedor <> vbNullString Then
        strSQLVale = strSQLVale & "AND INVENTARIOACTUAL.F5CODPRO IN ("
        strSQLVale = strSQLVale & "SELECT DET.F5CODPRO "
        strSQLVale = strSQLVale & "FROM "
        strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
        strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        strSQLVale = strSQLVale & "WHERE "
        strSQLVale = strSQLVale & "DET.TIPO = 'I' AND CAB.F1CODORI IN ('XC0') AND "
        strSQLVale = strSQLVale & "CAB.F2CODPROV = '" & strFiltrarPorProveedor & "' "
        strSQLVale = strSQLVale & "GROUP BY DET.F5CODPRO"
        strSQLVale = strSQLVale & ") "
    End If
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.FAMILIA, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.SUBFAMILIA, "
    strSQLVale = strSQLVale & "INVENTARIOACTUAL.F5NOMPRO"
    
    'strSQLVale = strSQLVale & "FAM.F7DESCON, "
    'strSQLVale = strSQLVale & "SFAM.F7DESCON, "
    'strSQLVale = strSQLVale & "PROD.F5NOMPRO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        Dim gGroupSum As dxGridSummaryGroup
        Dim gGroupItem As dxGridSummaryItem
        
        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColFamilia"
                    .SummaryField = "FAMILIA"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColSubFamilia"
                    .SummaryField = "SUBFAMILIA"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "F5CODPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = False
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColCodigo"
                    .SummaryField = "F5CODPRO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColDescripcion"
                    .SummaryField = "F5NOMPRO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColUM"
                    .SummaryField = "F7SIGMED"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Saldo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Saldo"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColSaldo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColSaldo"
                    .SummaryField = "SALDO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            If bolValorizado Then
                'Columna Costo Promedio
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 3
                    .Caption = "C. Promedio"
                    .Color = &HFFFFFF
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "COSTOPROMEDIO"
                    .Font.Bold = True
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColCostoPromedio"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColCostoPromedio"
                        .SummaryField = "COSTOPROMEDIO"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Total
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 3
                    .Caption = "Total"
                    .Color = &HFFFFFF
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "TOTAL"
                    .Font.Bold = True
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColTotal"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColTotal"
                        .SummaryField = "TOTAL"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Total Anterior
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 4
                    .Caption = "Total Anterior"
                    .Color = RGB(255, 235, 156)
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "TOTALANTERIOR"
                    .Font.Bold = True
                    .FontColor = RGB(156, 101, 0)
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColTotalAnterior"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColTotalAnterior"
                        .SummaryField = "TOTALANTERIOR"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Diferencia
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 4
                    .Caption = "Diferencia"
                    .Color = RGB(230, 185, 184)
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "DIFERENCIA"
                    .Font.Bold = True
                    .FontColor = RGB(0, 0, 0)
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColDiferencia"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColDiferencia"
                        .SummaryField = "DIFERENCIA"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
            End If
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            
            If bolValorizado And bolAgrupado Then
                .Columns.ColumnByFieldName("FAMILIA").GroupIndex = 0
                .Columns.ColumnByFieldName("SUBFAMILIA").GroupIndex = 1
                
                .m.FullExpand
            End If
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            
            .Bands.ITEM(3).Caption = "Periodo: " & UCase(Format(strFechaCorte, "MMMM-yyyy"))
            .Bands.ITEM(4).Caption = "Periodo: " & UCase(Format(strFechaFinMesAnterior, "MMMM-yyyy"))
            .Bands.ITEM(4).Visible = bolValorizado
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaInventarioProducto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaInventarioProducto"
    End Select
    
    Err.Clear
End Sub

'Listar Inventario de Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaInventarioProductoV2(ByVal grilla As dxDBGrid, _
                                            ByVal strFiltroSensitivo As String, _
                                            ByVal strFechaCorte As String, _
                                            ByVal bolValorizado As Boolean, _
                                            ByVal strCodMonedaValor As String, _
                                            ByVal bolAgrupado As Boolean, _
                                            Optional ByVal strCodAlmacen As String, _
                                            Optional ByVal strCodFamilia As String, _
                                            Optional ByVal strCodSubFamilia As String, _
                                            Optional ByVal listaImagenes As Object, _
                                            Optional ByVal bolMostrarStockNegativo As Boolean, _
                                            Optional ByVal strFiltrarPorProveedor As String, _
                                            Optional ByVal bolIncluirProductoDescontinuado As Boolean, _
                                            Optional ByVal strNombreTablaSQLParaResumen As String, _
                                            Optional ByVal bolEjecutarCorreccionCostoEnNegativo As Boolean)
    
    On Error GoTo errListarGrillaInventarioProductoV2
    
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    Dim strFechaFinMesAnterior As String
    
    intAnnoCorte = Year(CDate(strFechaCorte)) - IIf(Month(CDate(strFechaCorte)) > 1, 0, 1)
    intMesCorte = IIf(Month(CDate(strFechaCorte)) > 1, Month(CDate(strFechaCorte)) - 1, 12)
    
    strFechaFinMesAnterior = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
    
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        .CommandText = "Procesos.usp_ListarInventarioProducto"
        
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, strFiltroSensitivo)
        .Parameters.Append .CreateParameter("@FECHACORTE", adDate, adParamInput, 10, strFechaCorte)
        .Parameters.Append .CreateParameter("@VALORIZADO", adBoolean, adParamInput, , bolValorizado)
        .Parameters.Append .CreateParameter("@CODMONEDAVALOR", adVarChar, adParamInput, 1, strCodMonedaValor)
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, strCodAlmacen)
        .Parameters.Append .CreateParameter("@CODFAMILIA", adVarChar, adParamInput, 10, strCodFamilia)
        .Parameters.Append .CreateParameter("@CODSUBFAMILIA", adVarChar, adParamInput, 10, strCodSubFamilia)
        .Parameters.Append .CreateParameter("@MOSTRARSTOCKNEGATIVO", adBoolean, adParamInput, , bolMostrarStockNegativo)
        .Parameters.Append .CreateParameter("@CODPROVEEDOR", adVarChar, adParamInput, 4, strFiltrarPorProveedor)
        .Parameters.Append .CreateParameter("@INCLUIRPRODUCTODESCONTINUADO", adBoolean, adParamInput, , bolIncluirProductoDescontinuado)
        .Parameters.Append .CreateParameter("@FECHAFINMESANTERIOR", adDate, adParamInput, 10, strFechaFinMesAnterior)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, strNombreTablaSQLParaResumen)
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    If bolEjecutarCorreccionCostoEnNegativo Then
        Dim rsCorreccionCosto As New ADODB.Recordset
        
        strSQLVale = vbNullString
        strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " WHERE COSTOPROMEDIO < 0"
        
        If rsCorreccionCosto.State = 1 Then rsCorreccionCosto.Close
        
        rsCorreccionCosto.Open strSQLVale, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
        
        If Not rsCorreccionCosto.EOF Then
            rsCorreccionCosto.MoveFirst
            
            Do While Not rsCorreccionCosto.EOF
                strCodMoneda = strCodMonedaValor
                strCodProducto = Trim(rsCorreccionCosto!CodProducto & "")
                StrFecha = Format(strFechaCorte, "Short Date")
                
                dblValorVta = Me.obtenerUltimoCostoPromedioV2
                
                strSQLVale = vbNullString
                strSQLVale = strSQLVale & "UPDATE "
                strSQLVale = strSQLVale & strNombreTablaSQLParaResumen & " "
                strSQLVale = strSQLVale & "SET "
                strSQLVale = strSQLVale & "COSTOPROMEDIO = " & dblValorVta & ", "
                strSQLVale = strSQLVale & "TOTALACTUAL = (SALDO * " & dblValorVta & "), "
                strSQLVale = strSQLVale & "DIFERENCIA = (SALDO * " & dblValorVta & ") - TOTALANTERIOR "
                strSQLVale = strSQLVale & "WHERE "
                strSQLVale = strSQLVale & "CODPRODUCTO = '" & strCodProducto & "'"
                
                cnBdCPlus.Execute strSQLVale
                
                rsCorreccionCosto.MoveNext
            Loop
        End If
        
        If rsCorreccionCosto.State = 1 Then rsCorreccionCosto.Close
        
        Set rsCorreccionCosto = Nothing
    End If
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT * FROM " & strNombreTablaSQLParaResumen & " ORDER BY FAMILIA, SUBFAMILIA, NOMPRODUCTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        Dim gGroupSum As dxGridSummaryGroup
        Dim gGroupItem As dxGridSummaryItem
        
        With grilla
            'Columna Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Familia"
                .DisableEditor = True
                .FieldName = "FAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColFamilia"
                    .SummaryField = "FAMILIA"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Sub-Familia de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Sub-Familia"
                .DisableEditor = True
                .FieldName = "SUBFAMILIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubFamilia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColSubFamilia"
                    .SummaryField = "SUBFAMILIA"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                '.Visible = False
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColCodigo"
                    .SummaryField = "CODPRODUCTO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 300
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColDescripcion"
                    .SummaryField = "NOMPRODUCTO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColUM"
                    .SummaryField = "UM"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            'Columna Saldo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Saldo"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColSaldo"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            If bolValorizado And bolAgrupado Then
                Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                
                With gGroupItem
                    .ColumnName = "ColSaldo"
                    .SummaryField = "SALDO"
                    .SummaryType = cstCount
                    .SummaryFormat = " "
                End With
            End If
            
            If bolValorizado Then
                'Columna Costo Promedio
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 3
                    .Caption = "C. Promedio"
                    .Color = &HFFFFFF
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "COSTOPROMEDIO"
                    .Font.Bold = True
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColCostoPromedio"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColCostoPromedio"
                        .SummaryField = "COSTOPROMEDIO"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Total
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 3
                    .Caption = "Total"
                    .Color = &HFFFFFF
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "TOTALACTUAL"
                    .Font.Bold = True
                    .FontColor = &H80000012
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColTotal"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColTotal"
                        .SummaryField = "TOTALACTUAL"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Total Anterior
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 4
                    .Caption = "Total Anterior"
                    .Color = RGB(255, 235, 156)
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "TOTALANTERIOR"
                    .Font.Bold = True
                    .FontColor = RGB(156, 101, 0)
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColTotalAnterior"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColTotalAnterior"
                        .SummaryField = "TOTALANTERIOR"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
                
                'Columna Diferencia
                Set gColumn = .Columns.Add(gedSpinEdit)
                
                With gColumn
                    .Alignment = taRightJustify
                    .BandIndex = 4
                    .Caption = "Diferencia"
                    .Color = RGB(230, 185, 184)
                    .DecimalPlaces = 2
                    .DisableEditor = True
                    .FieldName = "DIFERENCIA"
                    .Font.Bold = True
                    .FontColor = RGB(0, 0, 0)
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    '.SummaryFooterType = cstSum
                    .ObjectName = "ColDiferencia"
                    .SummaryFooterType = cstSum
                    '.SummaryFooterFormat = " "
                    .Width = 70
                End With
                
                If bolAgrupado Then
                    Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
                    With gGroupItem
                        .ColumnName = "ColDiferencia"
                        .SummaryField = "DIFERENCIA"
                        .SummaryType = cstSum
                        '.SummaryFormat = " "
                    End With
                End If
            End If
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPRODUCTO"
            
            If bolValorizado And bolAgrupado Then
                .Columns.ColumnByFieldName("FAMILIA").GroupIndex = 0
                .Columns.ColumnByFieldName("SUBFAMILIA").GroupIndex = 1
                
                .m.FullExpand
            End If
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
            
            .Bands.ITEM(3).Caption = "Periodo: " & UCase(Format(strFechaCorte, "MMMM-yyyy"))
            .Bands.ITEM(4).Caption = "Periodo: " & UCase(Format(strFechaFinMesAnterior, "MMMM-yyyy"))
            .Bands.ITEM(4).Visible = bolValorizado
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLVale
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarGrillaInventarioProductoV2:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarGrillaInventarioProductoV2"
    End Select
    
    Err.Clear
End Sub

'Listar Resumen de Compras en Periodo Consultado en Grilla (QuamtumGrid)
Public Sub listarResumenCompras(ByVal grilla As dxDBGrid, _
                                ByVal strFechaDesde As String, _
                                ByVal strFechaHasta As String, _
                                ByVal intOpcionResumen As Integer, _
                                ByVal strFiltroSensitivo As String)
    
    On Error GoTo errListarResumenCompras
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
        
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION AS TIPO, "
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
            Case Else
                strSQLVale = strSQLVale & "TE.DESCRIPCION AS TIPO, "
                strSQLVale = strSQLVale & "FAM.F7DESCON AS FAMILIA, "
                strSQLVale = strSQLVale & "SFAM.F7DESCON AS SUBFAMILIA, "
                strSQLVale = strSQLVale & "PROD.F5NOMPRO AS DESCRIPCION, "
        End Select
    
    strSQLVale = strSQLVale & "MED.F7NOMMED AS UM, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD, "
    strSQLVale = strSQLVale & "SUM(VAL(FORMAT(IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE, 0), '#0.00'))) AS TOTALMN, "
    strSQLVale = strSQLVale & "SUM(VAL(FORMAT(IIF(CAB.F4MONEDA <> 'S', DET.F3TOTITE, 0), '#0.00'))) AS TOTALME "
    
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(((((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN EF2TIPOEXISTENCIA AS TE ON TE.CODIGO = PROD.F5TIPO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "(CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaDesde & "') AND CVDATE('" & strFechaHasta & "')) AND "
    strSQLVale = strSQLVale & "CAB.F4TIPOVALE = 'I' AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI = 'XC0' "
    strSQLVale = strSQLVale & "GROUP BY "
        
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION, "
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON, "
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON, "
            Case Else
                strSQLVale = strSQLVale & "TE.DESCRIPCION, "
                strSQLVale = strSQLVale & "FAM.F7DESCON, "
                strSQLVale = strSQLVale & "SFAM.F7DESCON, "
                strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
        End Select
        
    strSQLVale = strSQLVale & "MED.F7NOMMED "
    strSQLVale = strSQLVale & "ORDER BY "
        
        Select Case intOpcionResumen
            Case 0 'Por Tipo de Existencia
                strSQLVale = strSQLVale & "TE.DESCRIPCION"
            Case 1 'Por Familia
                strSQLVale = strSQLVale & "FAM.F7DESCON"
            Case 2 'Por Sub-Familia
                strSQLVale = strSQLVale & "SFAM.F7DESCON"
            Case Else
                strSQLVale = strSQLVale & "PROD.F5NOMPRO"
        End Select
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            If intOpcionResumen = 0 Or intOpcionResumen = 3 Then
                'Columna Tipo Existencia (Familia Contable)
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Tipo de Existencia"
                    .DisableEditor = True
                    .FieldName = "TIPO"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColTipoExistencia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 1 Or intOpcionResumen = 3 Then
                'Columna Familia de Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Familia"
                    .DisableEditor = True
                    .FieldName = "FAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 2 Or intOpcionResumen = 3 Then
                'Columna Sub-Familia
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Sub-Familia"
                    .DisableEditor = True
                    .FieldName = "SUBFAMILIA"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColSubFamilia"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 100
                End With
            End If
            
            If intOpcionResumen = 3 Then
                'Columna Descripción del Producto
                Set gColumn = .Columns.Add(gedTextEdit)
                
                With gColumn
                    .Alignment = taLeftJustify
                    .Caption = "Descripción de Producto"
                    .DisableEditor = True
                    .FieldName = "DESCRIPCION"
                    .Font.Charset = 0
                    .HeaderAlignment = taCenter
                    .ObjectName = "ColDescripcionProducto"
                    .SummaryFooterType = cstCount
                    .SummaryFooterFormat = " "
                    .Width = 200
                End With
            End If
            
            'Columna UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            'Columna Cantidad Adquirida en el Mes
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Total MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Total MN"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTALMN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColTotalMN"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Total ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Total ME"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTALME"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColTotalME"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            
            Select Case intOpcionResumen
                Case 0 'Por Tipo de Existencia
                    .KeyField = "TIPO"
                Case 1 'Por Familia
                    .KeyField = "FAMILIA"
                Case 2 'Por Sub-Familia
                    .KeyField = "SUBFAMILIA"
                Case Else
                    .KeyField = "DESCRIPCION"
            End Select
            
            .Filter.FilterActive = True
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarResumenCompras:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarResumenCompras"
    End Select
    
    Err.Clear
End Sub


'Listar Resumen de Compras, Ingresos y Salidas por Servicios de Terceros en Grilla (QuamtumGrid)
Public Sub listarResumenComprasPorProveedores(ByVal grilla As dxDBGrid, _
                                                    ByVal listaImagenes As Object, _
                                                    ByVal strCodProveedor As String, _
                                                    ByVal strFechaDesde As String, _
                                                    ByVal strFechaHasta As String, _
                                                    ByVal strConceptosEvaluados As String, _
                                                    ByVal strFiltroSensitivo As String)
    
    On Error GoTo errListarResumenComprasPorProveedores
    
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "PROV.F2NOMPROV, "
    strSQLVale = strSQLVale & "CONCEPTO.F1NOMORI, "
    strSQLVale = strSQLVale & "('[Almacen: ' & ALM.F2NOMALM & ']        [No. Vale: ' & CAB.F4NUMVAL & ']        [ID Integrado: ' & CAB.NUMENSAM & ']        [Fecha: ' & CAB.F4FECVAL & ']        [Moneda: ' & IIF(CAB.F4MONEDA = 'S', 'SOLES', 'DOLARES') & ']        [Guia: ' & CAB.F4SERGUIA & '-' & CAB.F4NUMGUIA & ']        [Doc.: ' & DOC.F2ABREV & CAB.F4SERDOC & '/' & CAB.F4NUMDOC & ']        [Fec. Doc.: ' & CAB.F4FECULT & ']') AS DATOS, "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "OC.F4FECEMI, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "SUM(DET.F3CANPRO) AS CANTIDAD, "
    strSQLVale = strSQLVale & "AVG(VAL(DET.F3PORCENTAJEDSCTO & '')) AS PORCENTAJEDSCTO, "
    
    'strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, NULL)) AS COSTOMN, "
    'strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'D', DET.F3VALDOL, NULL)) AS COSTOME, "
    
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, NULL) AS COSTOMN, "
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'D', DET.F3VALDOL, NULL) AS COSTOME, "
    
    strSQLVale = strSQLVale & "SUM((DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, NULL))) AS SUBTOTALMN, "
    strSQLVale = strSQLVale & "SUM((DET.F3CANPRO * IIF(CAB.F4MONEDA = 'D', DET.F3VALDOL, NULL))) AS SUBTOTALME, "
    strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3IGV, NULL)) AS IMPUESTOMN, "
    strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'D', DET.F3IGVDOL, NULL)) AS IMPUESTOME, "
    strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE, NULL)) AS TOTALMN, "
    strSQLVale = strSQLVale & "SUM(IIF(CAB.F4MONEDA = 'D', DET.F3TOTDOL, NULL)) AS TOTALME "
    
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "(((((((PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL) "
    strSQLVale = strSQLVale & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = CAB.F2CODALM) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS CONCEPTO ON CONCEPTO.F1CODORI = CAB.F1CODORI) "
    strSQLVale = strSQLVale & "LEFT JOIN DOCUMENTOS AS DOC ON DOC.F2CODDOC = CAB.F4TIPDOC) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2PROVEEDORES AS PROV ON PROV.F2CODPROV = CAB.F2CODPROV) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLVale = strSQLVale & "LEFT JOIN IF4ORDEN AS OC ON OC.F4NUMORD = DET.F4NUMORD "
    
    strSQLVale = strSQLVale & "WHERE "
        
        If strCodProveedor <> vbNullString Then
            strSQLVale = strSQLVale & "CAB.F2CODPROV IN ('" & strCodProveedor & "') AND "
        End If
        
    strSQLVale = strSQLVale & "CVDATE(CAB.F4FECVAL) BETWEEN CVDATE('" & strFechaDesde & "') AND CVDATE('" & strFechaHasta & "') AND "
    strSQLVale = strSQLVale & "CAB.F1CODORI IN ('" & strConceptosEvaluados & "') "
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLVale = strSQLVale & "AND (PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%') "
        End If
    
    strSQLVale = strSQLVale & "GROUP BY "
    strSQLVale = strSQLVale & "PROV.F2NOMPROV, "
    strSQLVale = strSQLVale & "CONCEPTO.F1NOMORI, "
    strSQLVale = strSQLVale & "ALM.F2NOMALM, "
    strSQLVale = strSQLVale & "CAB.F4NUMVAL, "
    strSQLVale = strSQLVale & "CAB.NUMENSAM, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', 'SOLES', 'DOLARES'), "
    strSQLVale = strSQLVale & "CAB.F4SERGUIA, "
    strSQLVale = strSQLVale & "CAB.F4NUMGUIA, "
    strSQLVale = strSQLVale & "DOC.F2ABREV, "
    strSQLVale = strSQLVale & "CAB.F4SERDOC, "
    strSQLVale = strSQLVale & "CAB.F4NUMDOC, "
    strSQLVale = strSQLVale & "CAB.F4FECULT, "
    strSQLVale = strSQLVale & "DET.F4NUMORD, "
    strSQLVale = strSQLVale & "OC.F4FECEMI, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, NULL), "
    strSQLVale = strSQLVale & "IIF(CAB.F4MONEDA = 'D', DET.F3VALDOL, NULL) "
    
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "PROV.F2NOMPROV, "
    strSQLVale = strSQLVale & "CONCEPTO.F1NOMORI, "
    strSQLVale = strSQLVale & "ALM.F2NOMALM, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "CAB.F4NUMVAL, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "DET.F4NUMORD"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        Dim gGroupSum As dxGridSummaryGroup
        Dim gGroupItem As dxGridSummaryItem
        
        With grilla
            'Columna Proveedor
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Proveedor"
                .DisableEditor = True
                .FieldName = "F2NOMPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColProveedor"
                .SummaryField = "F2NOMPROV"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Concepto de Movimiento de Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Concepto de Movimiento"
                .DisableEditor = True
                .FieldName = "F1NOMORI"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColConcepto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColConcepto"
                .SummaryField = "F1NOMORI"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Datos de Movimiento de Almacen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Datos de Movimiento"
                .DisableEditor = True
                .FieldName = "DATOS"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDatos"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColDatos"
                .SummaryField = "DATOS"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Numero de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "O/C"
                .DisableEditor = True
                .FieldName = "F4NUMORD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColOC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColOC"
                .SummaryField = "F4NUMORD"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Fecha de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Fec. O/C"
                .DisableEditor = True
                .FieldName = "F4FECEMI"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColOCFecEmi"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColOCFecEmi"
                .SummaryField = "F4FECEMI"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Nombre de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripcion de Producto"
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 200
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColProducto"
                .SummaryField = "F5NOMPRO"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "U.M."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F7SIGMED"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColUM"
                .SummaryField = "F7SIGMED"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColCantidad"
                .SummaryField = "CANTIDAD"
                .SummaryType = cstSum
            End With
            
            'Columna Porcentaje
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "% Dscto"
                .Color = &HC000C0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "PORCENTAJEDSCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColPorcentajeDscto"
                .SummaryFooterType = cstSum
                .Width = 50
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColPorcentajeDscto"
                .SummaryField = "PORCENTAJEDSCTO"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Costo MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Costo MN"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COSTOMN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCostoMN"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColCostoMN"
                .SummaryField = "COSTOMN"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Costo ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Costo ME"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COSTOME"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCostoME"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColCostoME"
                .SummaryField = "COSTOME"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Sub-Total MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Sub-Total MN"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SUBTOTALMN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubTotalMN"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColSubTotalMN"
                .SummaryField = "SUBTOTALMN"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Sub-Total ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Sub-Total ME"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SUBTOTALME"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSubTotalME"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColSubTotalME"
                .SummaryField = "SUBTOTALME"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Impuesto MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Impuesto MN"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "IMPUESTOMN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColImpuestoMN"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColImpuestoMN"
                .SummaryField = "IMPUESTOMN"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Impuesto ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Impuesto ME"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "IMPUESTOME"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColImpuestoME"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColImpuestoME"
                .SummaryField = "IMPUESTOME"
                .SummaryType = cstCount
                .SummaryFormat = " "
            End With
            
            'Columna Total MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Total MN"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTALMN"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColTotalMN"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColTotalMN"
                .SummaryField = "TOTALMN"
                .SummaryType = cstSum
            End With
            
            'Columna Total ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Total ME"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTALME"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColTotalME"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            Set gGroupItem = .SummaryGroups.ITEM(0).SummaryItems.Add
                    
            With gGroupItem
                .ColumnName = "ColTotalME"
                .SummaryField = "TOTALME"
                .SummaryType = cstSum
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
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = strSQLVale
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "DATOS"
            
            .Columns.ColumnByFieldName("F2NOMPROV").GroupIndex = 0
            .Columns.ColumnByFieldName("F1NOMORI").GroupIndex = 1
            .Columns.ColumnByFieldName("DATOS").GroupIndex = 2
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("COSTOMN").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("COSTOME").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("SUBTOTALMN").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("SUBTOTALME").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("IMPUESTOMN").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("IMPUESTOME").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("TOTALMN").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("TOTALME").SummaryFooterType = cstSum
        End With
    End If
    
    strSQLVale = vbNullString
    
    Exit Sub
errListarResumenComprasPorProveedores:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsVale: ListarResumenComprasPorProveedores"
    End Select
    
    Err.Clear
End Sub

Public Function devuelveSQLOrdenProduccionMovimiento() As String
    strSQLVale = vbNullString
    strSQLVale = strSQLVale & "SELECT "
    strSQLVale = strSQLVale & "IIF(DET.COD_SOLICITUD = '', 'POR FUERA', 'A PARTIR DE OP') AS TIPO, "
    strSQLVale = strSQLVale & "CAB.F2CODALM, "
    strSQLVale = strSQLVale & "ALM.F2NOMALM, "
    strSQLVale = strSQLVale & "CAB.F4NUMVAL, "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "CAB.F1CODORI, "
    strSQLVale = strSQLVale & "ORI.F1NOMORI, "
    strSQLVale = strSQLVale & "DET.COD_SOLICITUD, "
    strSQLVale = strSQLVale & "DET.F5CODPRO, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO, "
    strSQLVale = strSQLVale & "MED.F7SIGMED, "
    strSQLVale = strSQLVale & "DET.F3CANPRO "
    strSQLVale = strSQLVale & "FROM "
    strSQLVale = strSQLVale & "PROCESOS.IF3VALES AS DET "
    strSQLVale = strSQLVale & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF2ALMACENES AS ALM ON ALM.F2CODALM = CAB.F2CODALM "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.SF1ORIGENES AS ORI ON ORI.F1CODORI = CAB.F1CODORI "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO "
    strSQLVale = strSQLVale & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
    strSQLVale = strSQLVale & "WHERE "
    strSQLVale = strSQLVale & "CAB.F4ORDTRA = '" & strOrdenTrabajo & "' "
    strSQLVale = strSQLVale & "ORDER BY "
    strSQLVale = strSQLVale & "CAB.F4FECVAL, "
    strSQLVale = strSQLVale & "PROD.F5NOMPRO"
    
    devuelveSQLOrdenProduccionMovimiento = strSQLVale
    
    strSQLVale = vbNullString
End Function

