VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPagDcto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dblCorrelativo                      As Double
Private strTipoIngreso                      As String
Private dblItem                             As Double

Private strNumeroComprobante                As String
Private strFechaComprobante                 As String
Private strFechaVencimiento                 As String
Private strCodProveedor                     As String
Private strRucProveedor                     As String
Private strNomProveedor                     As String
Private strCodMoneda                        As String
Private dblTipoCambio                       As Double
Private dblTotalFacturado                   As Double
Private dblSaldoFacturado                   As Double
Private dblTotalDetraccion                  As Double
Private dblSaldoDetraccion                  As Double
Private strDebe_Haber                       As String
Private strReferencia                       As String
Private strDetalle                          As String
Private strCuentaBancaria                   As String
Private strMovimientoBancario               As String
Private bolDocAnulado                       As Boolean
Private lngAnnoReporte                      As Long
Private strFechaReporte                     As String
Private lngNumeroReporte                    As Long
Private lngBanco                            As Long

Rem SK ADD: Atributos Adicionales
Private strGrupo                            As String
Private strCtaContable                      As String
Private strAnnoRegCompra                    As String
Private strMovRegCompra                     As String
Private dblAsiento                          As Double
Private strTipoDocumento                    As String
Private strSerieDocumento                   As String
Private dblNumeroDocumento                  As Double
Private strNotas                            As String
Private strConcepto                         As String
Private strCentroCosto                      As String


Rem SK: Atributos de Cta_Movimiento
Private strMCodProveedor                    As String
Private dblMCorrelaPago                     As Double
Private dblMCorrelaDocumento                As Double
Private dblMImputadoSoles                   As Double
Private dblMImputadoDolares                 As Double
Private dblMTipoCambio                      As Double
Private lngMAnnoReporte                     As Long
Private lngMNumeroReporte                   As Long
Private strMFechaMovimiento                 As String
Private strMFechaReporte                    As String



Private strSQLSelectAlter                   As String

Private rstPagDcto                            As ADODB.Recordset
Private strSQLPagDcto                         As String

Public Property Let Correlativo(ByVal value As Double)
    dblCorrelativo = value
End Property

Public Property Get Correlativo() As Double
    Correlativo = dblCorrelativo
End Property

Public Property Let TipoIngreso(ByVal value As String)
    strTipoIngreso = value
End Property

Public Property Get TipoIngreso() As String
    TipoIngreso = strTipoIngreso
End Property

Public Property Let ITEM(ByVal value As Double)
    dblItem = value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property


'NumeroComprobante
Public Property Let NumeroComprobante(ByVal value As String)
    strNumeroComprobante = value
End Property

Public Property Get NumeroComprobante() As String
    NumeroComprobante = strNumeroComprobante
End Property



Public Property Let FechaComprobante(ByVal value As String)
    strFechaComprobante = value
End Property

Public Property Get FechaComprobante() As String
    FechaComprobante = strFechaComprobante
End Property

Public Property Let FechaVencimiento(ByVal value As String)
    strFechaVencimiento = value
End Property

Public Property Get FechaVencimiento() As String
    FechaVencimiento = strFechaVencimiento
End Property

Public Property Let CodProveedor(ByVal value As String)
    strCodProveedor = value
End Property

Public Property Get CodProveedor() As String
    CodProveedor = strCodProveedor
End Property

'RUCProveedor
Public Property Let RucProveedor(ByVal value As String)
    strRucProveedor = value
End Property

Public Property Get RucProveedor() As String
    RucProveedor = strRucProveedor
End Property

Public Property Let NomProveedor(ByVal value As String)
    strNomProveedor = value
End Property

Public Property Get NomProveedor() As String
    NomProveedor = strNomProveedor
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

Public Property Let TotalFacturado(ByVal value As Double)
    dblTotalFacturado = value
End Property

Public Property Get TotalFacturado() As Double
    TotalFacturado = dblTotalFacturado
End Property

Public Property Let SaldoFacturado(ByVal value As Double)
    dblSaldoFacturado = value
End Property

Public Property Get SaldoFacturado() As Double
    SaldoFacturado = dblSaldoFacturado
End Property

Public Property Let TotalDetraccion(ByVal value As Double)
    dblTotalDetraccion = value
End Property

Public Property Get TotalDetraccion() As Double
    TotalDetraccion = dblTotalDetraccion
End Property

Public Property Let SaldoDetraccion(ByVal value As Double)
    dblSaldoDetraccion = value
End Property

Public Property Get SaldoDetraccion() As Double
    SaldoDetraccion = dblSaldoDetraccion
End Property

Public Property Let Debe_Haber(ByVal value As String)
    strDebe_Haber = value
End Property

Public Property Get Debe_Haber() As String
    Debe_Haber = strDebe_Haber
End Property

Public Property Let referencia(ByVal value As String)
    strReferencia = value
End Property

Public Property Get referencia() As String
    referencia = strReferencia
End Property

Public Property Let Detalle(ByVal value As String)
    strDetalle = value
End Property

Public Property Get Detalle() As String
    Detalle = strDetalle
End Property

Public Property Let CuentaBancaria(ByVal value As String)
    strCuentaBancaria = value
End Property

Public Property Get CuentaBancaria() As String
    CuentaBancaria = strCuentaBancaria
End Property

Public Property Let MovimientoBancario(ByVal value As String)
    strMovimientoBancario = value
End Property

Public Property Get MovimientoBancario() As String
    MovimientoBancario = strMovimientoBancario
End Property
'DocAnulado
Public Property Let DocAnulado(ByVal value As Boolean)
    bolDocAnulado = value
End Property

Public Property Get DocAnulado() As Boolean
    DocAnulado = bolDocAnulado
End Property

Public Property Let AnnoReporte(ByVal value As Long)
    lngAnnoReporte = value
End Property

Public Property Get AnnoReporte() As Long
    AnnoReporte = lngAnnoReporte
End Property

Public Property Let FechaReporte(ByVal value As String)
    strFechaReporte = value
End Property

Public Property Get FechaReporte() As String
    FechaReporte = strFechaReporte
End Property

Public Property Let NumeroReporte(ByVal value As Long)
    lngNumeroReporte = value
End Property

Public Property Get NumeroReporte() As Long
    ITEM = lngNumeroReporte
End Property

Public Property Let Banco(ByVal value As Long)
    lngBanco = value
End Property

Public Property Get Banco() As Long
    Banco = lngBanco
End Property





Rem SK ADD: Atributos Adicionales
'Grupo
Public Property Let Grupo(ByVal value As String)
    strGrupo = value
End Property

Public Property Get Grupo() As String
    Grupo = strGrupo
End Property
'CtaContable
Public Property Let CtaContable(ByVal value As String)
    strCtaContable = value
End Property

Public Property Get CtaContable() As String
    CtaContable = strCtaContable
End Property
'AnnoRegCompra
Public Property Let AnnoRegCompra(ByVal value As String)
    strAnnoRegCompra = value
End Property

Public Property Get AnnoRegCompra() As String
    AnnoRegCompra = strAnnoRegCompra
End Property
'MovRegCompra
Public Property Let MovRegCompra(ByVal value As String)
    strMovRegCompra = value
End Property

Public Property Get MovRegCompra() As String
    MovRegCompra = strMovRegCompra
End Property
'Asiento
Public Property Let asiento(ByVal value As Double)
    dblAsiento = value
End Property

Public Property Get asiento() As Double
    asiento = dblAsiento
End Property
'TipoDocumento
Public Property Let TipoDocumento(ByVal value As String)
    strTipoDocumento = value
End Property

Public Property Get TipoDocumento() As String
    TipoDocumento = strTipoDocumento
End Property
'SerieDocumento
Public Property Let SerieDocumento(ByVal value As String)
    strSerieDocumento = value
End Property

Public Property Get SerieDocumento() As String
    SerieDocumento = strSerieDocumento
End Property
'NumeroDocumento
Public Property Let NumeroDocumento(ByVal value As Double)
    dblNumeroDocumento = value
End Property

Public Property Get NumeroDocumento() As Double
    NumeroDocumento = dblNumeroDocumento
End Property
'Notas
Public Property Let Notas(ByVal value As String)
    strNotas = value
End Property

Public Property Get Notas() As String
    Notas = strNotas
End Property
'Concepto
Public Property Let Concepto(ByVal value As String)
    strConcepto = value
End Property

Public Property Get Concepto() As String
    Concepto = strConcepto
End Property
'Centro de Costo
Public Property Let CentroCosto(ByVal value As String)
    strCentroCosto = value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property





'Atributos de Cuenta_Movimiento
'MCodCliente
Public Property Let MCodProveedor(ByVal value As String)
    strMCodProveedor = value
End Property

Public Property Get MCodProveedor() As String
    MCodProveedor = strMCodProveedor
End Property
'MCorrelaPago
Public Property Let MCorrelaPago(ByVal value As Double)
    dblMCorrelaPago = value
End Property

Public Property Get MCorrelaPago() As Double
    MCorrelaPago = dblMCorrelaPago
End Property
'MCorrelaDocumento
Public Property Let MCorrelaDocumento(ByVal value As Double)
    dblMCorrelaDocumento = value
End Property

Public Property Get MCorrelaDocumento() As Double
    MCorrelaDocumento = dblMCorrelaDocumento
End Property
'MImputadoSoles
Public Property Let MImputadoSoles(ByVal value As Double)
    dblMImputadoSoles = value
End Property

Public Property Get MImputadoSoles() As Double
    MImputadoSoles = dblMImputadoSoles
End Property
'MImputadoDolares
Public Property Let MImputadoDolares(ByVal value As Double)
    dblMImputadoDolares = value
End Property

Public Property Get MImputadoDolares() As Double
    MImputadoDolares = dblMImputadoDolares
End Property
'MTipoCambio
Public Property Let MTipoCambio(ByVal value As Double)
    dblMTipoCambio = value
End Property

Public Property Get MTipoCambio() As Double
    MTipoCambio = dblMTipoCambio
End Property
'MAnnoReporte
Public Property Let MAnnoReporte(ByVal value As Long)
    lngMAnnoReporte = value
End Property

Public Property Get MAnnoReporte() As Long
    MAnnoReporte = lngMAnnoReporte
End Property
'MNumeroReporte
Public Property Let MNumeroReporte(ByVal value As Long)
    lngMNumeroReporte = value
End Property

Public Property Get MNumeroReporte() As Long
    MNumeroReporte = lngMNumeroReporte
End Property
'MFechaMovimiento
Public Property Let MFechaMovimiento(ByVal value As String)
    strMFechaMovimiento = value
End Property

Public Property Get MFechaMovimiento() As String
    MFechaMovimiento = strMFechaMovimiento
End Property
'MFechaReporte
Public Property Let MFechaReporte(ByVal value As String)
    strMFechaReporte = value
End Property

Public Property Get MFechaReporte() As String
    MFechaReporte = strMFechaReporte
End Property



Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property

Public Sub inicializarEntidades()
    dblCorrelativo = 0
    strTipoIngreso = vbNullString
    dblItem = 0
    
    strNumeroComprobante = vbNullString
    strFechaComprobante = vbNullString
    strFechaVencimiento = vbNullString
    strCodProveedor = vbNullString
    strRucProveedor = vbNullString
    strNomProveedor = vbNullString
    strCodMoneda = vbNullString
    dblTipoCambio = 0
    dblTotalFacturado = 0
    dblSaldoFacturado = 0
    dblTotalDetraccion = 0
    dblSaldoDetraccion = 0
    strDebe_Haber = vbNullString
    strReferencia = vbNullString
    strDetalle = vbNullString
    strCuentaBancaria = vbNullString
    strMovimientoBancario = vbNullString
    bolDocAnulado = False
    lngAnnoReporte = 0
    strFechaReporte = vbNullString
    lngNumeroReporte = 0
    lngBanco = 0
    
    Rem SK ADD: Atributos Adicionales
    strGrupo = vbNullString
    strCtaContable = vbNullString
    strAnnoRegCompra = vbNullString
    strMovRegCompra = vbNullString
    dblAsiento = 0
    strTipoDocumento = vbNullString
    strSerieDocumento = vbNullString
    dblNumeroDocumento = 0
    strNotas = vbNullString
    strConcepto = vbNullString
    strCentroCosto = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesMovimiento()
    strMCodProveedor = vbNullString
    dblMCorrelaPago = 0
    dblMCorrelaDocumento = 0
    dblMImputadoSoles = 0
    dblMImputadoDolares = 0
    dblMTipoCambio = 0
    lngMAnnoReporte = 0
    lngMNumeroReporte = 0
    strMFechaMovimiento = vbNullString
    strMFechaReporte = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerPagDcto() As Boolean
    On Error GoTo errObtenerPagDcto
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT * "
    strSQLPagDcto = strSQLPagDcto & "FROM PAG_DCTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE CORRELA = " & dblCorrelativo
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstPagDcto.EOF Then
        dblCorrelativo = Val(rstPagDcto!Correla & "")
        strTipoIngreso = Trim(rstPagDcto!VIA_INGR & "")
        dblItem = Val(rstPagDcto!ITEM & "")
        
        strNumeroComprobante = Trim(rstPagDcto!NRO_COMP & "")
        strFechaComprobante = Trim(rstPagDcto!fch_comp & "")
        strFechaVencimiento = Trim(rstPagDcto!FCH_VCTO & "")
        strCodProveedor = Trim(rstPagDcto!proveedor & "")
        strRucProveedor = Trim(rstPagDcto!ruc & "")
        strNomProveedor = Trim(rstPagDcto!NOMPROV & "")
        strCodMoneda = Trim(rstPagDcto!Moneda & "")
        dblTipoCambio = Val(rstPagDcto!TCambio & "")
        dblTotalFacturado = Val(rstPagDcto!TOTAL & "")
        dblSaldoFacturado = Val(rstPagDcto!SALDO & "")
        dblTotalDetraccion = 0
        dblSaldoDetraccion = 0
        strDebe_Haber = Trim(rstPagDcto!DEB_HAB & "")
        strReferencia = Trim(rstPagDcto!referencia & "")
        strDetalle = Trim(rstPagDcto!Detalle & "")
        strCuentaBancaria = Trim(rstPagDcto!CTABANC & "")
        strMovimientoBancario = Trim(rstPagDcto!MOVBANC & "")
        bolDocAnulado = False
        lngAnnoReporte = Val(rstPagDcto!ANO_REPO & "")
        strFechaReporte = Trim(rstPagDcto!FCH_REPO & "")
        lngNumeroReporte = Val(rstPagDcto!NRO_REPO & "")
        lngBanco = 0
        
        Rem SK ADD: Atributos Adicionales
        strGrupo = Trim(rstPagDcto!F4GRUPO & "")
        strCtaContable = Trim(rstPagDcto!f4ctacont & "")
        strAnnoRegCompra = Trim(rstPagDcto!F4ANNORC & "")
        strMovRegCompra = Trim(rstPagDcto!REG_COM & "")
        dblAsiento = Val(rstPagDcto!asiento & "")
        strTipoDocumento = Trim(rstPagDcto!TipDoc & "")
        strSerieDocumento = Trim(rstPagDcto!SERIE & "")
        dblNumeroDocumento = Val(rstPagDcto!NUMDOC & "")
        strNotas = Trim(rstPagDcto!Notas & "")
        strConcepto = Trim(rstPagDcto!Concepto & "")
        strCentroCosto = Trim(rstPagDcto!F4CENTRO & "")
        
        strSQLSelectAlter = vbNullString
        
        obtenerPagDcto = True
    Else
        obtenerPagDcto = False
    End If
    
    rstPagDcto.Close
    
    Set rstPagDcto = Nothing
    
    strSQLPagDcto = vbNullString
    
    Exit Function
errObtenerPagDcto:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsPagDcto: ObtenerPagDcto"
    End Select
    
    obtenerPagDcto = False
    
    Err.Clear
End Function

Public Sub obtenerConfigPagDcto()
    On Error GoTo errObtenerConfigPagDcto
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = "SELECT * " & _
                    "FROM PAG_DCTO " & _
                    "WHERE CORRELA = " & dblCorrelativo
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPagDcto.EOF Then
        dblCorrelativo = Val(rstPagDcto!Correla & "")
        strTipoIngreso = Trim(rstPagDcto!VIA_INGR & "")
        dblItem = Val(rstPagDcto!ITEM & "")
        
        strNumeroComprobante = Trim(rstPagDcto!NRO_COMP & "")
        strFechaComprobante = Trim(rstPagDcto!fch_comp & "")
        strFechaVencimiento = Trim(rstPagDcto!FCH_VCTO & "")
        strCodProveedor = Trim(rstPagDcto!proveedor & "")
        strRucProveedor = Trim(rstPagDcto!ruc & "")
        strNomProveedor = Trim(rstPagDcto!NOMPROV & "")
        strCodMoneda = Trim(rstPagDcto!Moneda & "")
        dblTipoCambio = Val(rstPagDcto!TCambio & "")
        dblTotalFacturado = Val(rstPagDcto!TOTAL & "")
        dblSaldoFacturado = Val(rstPagDcto!SALDO & "")
        dblTotalDetraccion = 0
        dblSaldoDetraccion = 0
        strDebe_Haber = Trim(rstPagDcto!DEB_HAB & "")
        strReferencia = Trim(rstPagDcto!referencia & "")
        strDetalle = Trim(rstPagDcto!Detalle & "")
        strCuentaBancaria = Trim(rstPagDcto!CTABANC & "")
        strMovimientoBancario = Trim(rstPagDcto!MOVBANC & "")
        bolDocAnulado = False
        lngAnnoReporte = Val(rstPagDcto!ANO_REPO & "")
        strFechaReporte = Trim(rstPagDcto!FCH_REPO & "")
        lngNumeroReporte = Val(rstPagDcto!NRO_REPO & "")
        lngBanco = 0
        
        Rem SK ADD: Atributos Adicionales
        strGrupo = Trim(rstPagDcto!F4GRUPO & "")
        strCtaContable = Trim(rstPagDcto!f4ctacont & "")
        strAnnoRegCompra = Trim(rstPagDcto!F4ANNORC & "")
        strMovRegCompra = Trim(rstPagDcto!REG_COM & "")
        dblAsiento = Val(rstPagDcto!asiento & "")
        strTipoDocumento = Trim(rstPagDcto!TipDoc & "")
        strSerieDocumento = Trim(rstPagDcto!SERIE & "")
        dblNumeroDocumento = Val(rstPagDcto!NUMDOC & "")
        strNotas = Trim(rstPagDcto!Notas & "")
        strConcepto = Trim(rstPagDcto!Concepto & "")
        strCentroCosto = Trim(rstPagDcto!F4CENTRO & "")
    End If
    
    rstPagDcto.Close
    
    Set rstPagDcto = Nothing
    
    strSQLPagDcto = vbNullString
    
    Exit Sub
errObtenerConfigPagDcto:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsPagDcto: ObtenerConfigPagDcto"
    End Select
    
    Err.Clear
End Sub

Public Function obtenerPagMovimiento() As Boolean
    On Error GoTo errObtenerPagMovimiento
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO.* "
    strSQLPagDcto = strSQLPagDcto & "FROM "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "CORR_COMP = " & dblMCorrelaPago & " AND "
    strSQLPagDcto = strSQLPagDcto & "CORR_DCTO = " & dblMCorrelaDocumento
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstPagDcto.EOF Then
        strMCodProveedor = Trim(rstPagDcto!proveedor & "")
        dblMCorrelaPago = Val(rstPagDcto!CORR_COMP & "")
        dblMCorrelaDocumento = Val(rstPagDcto!CORR_DCTO & "")
        dblMImputadoSoles = Val(rstPagDcto!IMPUTASO & "")
        dblMImputadoDolares = Val(rstPagDcto!IMPUTADO & "")
        dblMTipoCambio = Val(rstPagDcto!TCambio & "")
        lngMAnnoReporte = Val(rstPagDcto!ANO_REPO & "")
        lngMNumeroReporte = Val(rstPagDcto!NRO_REPO & "")
        strMFechaMovimiento = Trim(rstPagDcto!FCH_MVTO & "")
        strMFechaReporte = Trim(rstPagDcto!FCH_REPO & "")
        
        obtenerPagMovimiento = True
    Else
        obtenerPagMovimiento = False
    End If
    
    rstPagDcto.Close
    
    Set rstPagDcto = Nothing
    
    strSQLPagDcto = vbNullString
    
    Exit Function
errObtenerPagMovimiento:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsPagDcto: ObtenerPagMovimiento"
    End Select
    
    obtenerPagMovimiento = False
    
    Err.Clear
End Function

Public Sub obtenerConfigPagMovimiento()
    On Error GoTo errObtenerConfigPagMovimiento
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO.* "
    strSQLPagDcto = strSQLPagDcto & "FROM "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "CORR_COMP = " & dblMCorrelaPago & " AND "
    strSQLPagDcto = strSQLPagDcto & "CORR_DCTO = " & dblMCorrelaDocumento
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstPagDcto.EOF Then
        strMCodProveedor = Trim(rstPagDcto!proveedor & "")
        dblMCorrelaPago = Val(rstPagDcto!CORR_COMP & "")
        dblMCorrelaDocumento = Val(rstPagDcto!CORR_DCTO & "")
        dblMImputadoSoles = Val(rstPagDcto!IMPUTASO & "")
        dblMImputadoDolares = Val(rstPagDcto!IMPUTADO & "")
        dblMTipoCambio = Val(rstPagDcto!TCambio & "")
        lngMAnnoReporte = Val(rstPagDcto!ANO_REPO & "")
        lngMNumeroReporte = Val(rstPagDcto!NRO_REPO & "")
        strMFechaMovimiento = Trim(rstPagDcto!FCH_MVTO & "")
        strMFechaReporte = Trim(rstPagDcto!FCH_REPO & "")
    End If
    
    rstPagDcto.Close
    
    Set rstPagDcto = Nothing
    
    strSQLPagDcto = vbNullString
    
    Exit Sub
errObtenerConfigPagMovimiento:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsPagDcto: ObtenerConfigPagMovimiento"
    End Select
    
    Err.Clear
End Sub

Public Function generarCodigoPagDcto() As Double
    On Error GoTo errGenerarCodigoPagDcto
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT TOP 1 CORRELA "
    strSQLPagDcto = strSQLPagDcto & "FROM PAG_DCTO "
    strSQLPagDcto = strSQLPagDcto & "ORDER BY CORRELA DESC"
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPagDcto.EOF Then
        generarCodigoPagDcto = Val(rstPagDcto!Correla & "") + 1
    Else
        generarCodigoPagDcto = 1
    End If
    
    rstPagDcto.Close
    
    Set rstPagDcto = Nothing
    
    strSQLPagDcto = vbNullString
    
    Exit Function
errGenerarCodigoPagDcto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: GenerarCodigoPagDcto"
    
    generarCodigoPagDcto = 0
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT PAG_DCTO.* "
    strSQLPagDcto = strSQLPagDcto & "FROM PAG_DCTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE CORRELA = " & dblCorrelativo
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPagDcto.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstPagDcto.Close
    strSQLPagDcto = vbNullString
    
    Set rstPagDcto = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function verificarExistenciaMovimiento() As Boolean
    On Error GoTo errVerificarExistenciaMovimiento
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO.* "
    strSQLPagDcto = strSQLPagDcto & "FROM "
    strSQLPagDcto = strSQLPagDcto & "PAG_MVTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "CORR_COMP = " & dblMCorrelaPago & " AND "
    strSQLPagDcto = strSQLPagDcto & "CORR_DCTO = " & dblMCorrelaDocumento
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPagDcto.EOF Then
        verificarExistenciaMovimiento = True
    Else
        verificarExistenciaMovimiento = False
    End If
    
    rstPagDcto.Close
    strSQLPagDcto = vbNullString
    
    Set rstPagDcto = Nothing
    
    Exit Function
errVerificarExistenciaMovimiento:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: VerificarExistenciaMovimiento"
    
    verificarExistenciaMovimiento = False
    
    Err.Clear
End Function

Public Function guardarPagDcto(Optional ByVal bolActualizarSoloSaldo As Boolean) As Boolean
    On Error GoTo errGuardarPagDcto
    
    strSQLPagDcto = vbNullString
    
    ''cnn_dbbancos.BeginTrans
    
    If Not verificarExistencia Then
        dblCorrelativo = generarCodigoPagDcto
        
        strSQLPagDcto = strSQLPagDcto & "INSERT INTO PAG_DCTO(CORRELA, VIA_INGR, ITEM, NRO_COMP, FCH_COMP, FCH_VCTO, PROVEEDOR, PROVEEDORO, " & _
                                                "RUC, NOMPROV, MONEDA, MONEDAO, TOTAL, TOTALO, SALDO, TCAMBIO, TCAMBIOO, DEB_HAB, " & _
                                                "ANO_REPO, NRO_REPO, FCH_REPO, CTABANC, MOVBANC, F4GLOSA, F4GRUPO, F4CTACONT, " & _
                                                "F4ANNORC, REG_COM, ASIENTO, TIPDOC, SERIE, NUMDOC, NOTAS, CONCEPTO, DETALLE, REFERENCIA, F4CENTRO) "
                                                
        strSQLPagDcto = strSQLPagDcto & "VALUES(" & dblCorrelativo & ", '" & strTipoIngreso & "', " & dblItem & ", '" & strNumeroComprobante & "', " & _
                                                "CVDATE('" & strFechaComprobante & "'), " & IIf(strFechaVencimiento = vbNullString, _
                                                "NULL", "CVDATE('" & strFechaVencimiento & "')") & ", '" & _
                                                strCodProveedor & "', '" & strCodProveedor & "', '" & strRucProveedor & "', '" & _
                                                strNomProveedor & "', '" & strCodMoneda & "', '" & strCodMoneda & "', " & _
                                                dblTotalFacturado & ", " & dblTotalFacturado & ", " & dblSaldoFacturado & "," & dblTipoCambio & ", " & _
                                                dblTipoCambio & ", '" & strDebe_Haber & "', " & lngAnnoReporte & ", " & lngNumeroReporte & ", " & _
                                                IIf(strFechaReporte <> vbNullString, "CVDATE('" & strFechaReporte & "')", "NULL") & ", '" & strCuentaBancaria & "', '" & strMovimientoBancario & "', '" & _
                                                strReferencia & "', '" & strGrupo & "', '" & strCtaContable & "', '" & strAnnoRegCompra & "', '" & _
                                                strMovRegCompra & "', " & dblAsiento & ", '" & strTipoDocumento & "', " & IIf(strSerieDocumento = vbNullString, "NULL", "'" & strSerieDocumento & "'") & ", " & _
                                                dblNumeroDocumento & ", '" & strNotas & "', '" & strConcepto & "', '" & strDetalle & "', '" & strReferencia & "', '" & strCentroCosto & "')"
    Else
        strSQLPagDcto = strSQLPagDcto & "UPDATE "
            strSQLPagDcto = strSQLPagDcto & "PAG_DCTO "
        strSQLPagDcto = strSQLPagDcto & "SET "
            
            If Not bolActualizarSoloSaldo Then
                If strTipoIngreso <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "VIA_INGR = '" & strTipoIngreso & "', "
                End If
                
                If dblItem <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "ITEM = " & dblItem & ", "
                End If
                
                If strNumeroComprobante <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "NRO_COMP = '" & strNumeroComprobante & "', "
                End If
                
                If IsDate(strFechaComprobante) Then
                    strSQLPagDcto = strSQLPagDcto & "FCH_COMP = CVDATE('" & strFechaComprobante & "'), "
                End If
                    
                If IsDate(strFechaVencimiento) Then
                    strSQLPagDcto = strSQLPagDcto & "FCH_VCTO = CVDATE('" & strFechaVencimiento & "'), "
                End If
                
                If strCodProveedor <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "PROVEEDOR = '" & strCodProveedor & "', "
                    strSQLPagDcto = strSQLPagDcto & "PROVEEDORO = '" & strCodProveedor & "', "
                End If
                
                If strRucProveedor <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "RUC = '" & strRucProveedor & "', "
                End If
                
                If strNomProveedor <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "NOMPROV = '" & strNomProveedor & "', "
                End If
                
                If strCodMoneda <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "MONEDA = '" & strCodMoneda & "', "
                    strSQLPagDcto = strSQLPagDcto & "MONEDAO = '" & strCodMoneda & "', "
                End If
                
                If dblTipoCambio <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "TCAMBIO = " & dblTipoCambio & ", "
                    strSQLPagDcto = strSQLPagDcto & "TCAMBIOO = " & dblTipoCambio & ", "
                End If
                
                If dblTotalFacturado <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "TOTAL = " & dblTotalFacturado & ", "
                    strSQLPagDcto = strSQLPagDcto & "TOTALO = " & dblTotalFacturado & ", "
                End If
                
                If dblSaldoFacturado <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "SALDO = " & dblSaldoFacturado & ", "
                End If
                
                If strDebe_Haber <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "DEB_HAB = '" & strDebe_Haber & "', "
                End If
                
                If lngAnnoReporte <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "ANO_REPO = " & lngAnnoReporte & ", "
                End If
                
                If lngNumeroReporte <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "NRO_REPO = " & lngNumeroReporte & ", "
                End If
                
                If IsDate(strFechaReporte) Then
                    strSQLPagDcto = strSQLPagDcto & "FCH_REPO = CVDATE('" & strFechaReporte & "'), "
                End If
                
                If strCuentaBancaria <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "CTABANC = '" & strCuentaBancaria & "', "
                End If
                
                If strMovimientoBancario <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "MOVBANC = '" & strMovimientoBancario & "', "
                End If
                
                If strReferencia <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "F4GLOSA = '" & strReferencia & "', "
                End If
                
                If strGrupo <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "F4GRUPO = '" & strGrupo & "', "
                End If
                
                If strCtaContable <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "F4CTACONT = '" & strCtaContable & "', "
                End If
                
                If strAnnoRegCompra <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "F4ANNORC = '" & strAnnoRegCompra & "', "
                End If
                
                If strMovRegCompra <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "REG_COM = '" & strMovRegCompra & "', "
                End If
                
                If strTipoDocumento <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "TIPDOC = '" & strTipoDocumento & "', "
                End If
                
                If strSerieDocumento <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "SERIE = '" & strSerieDocumento & "', "
                End If
                
                If dblNumeroDocumento <> 0 Then
                    strSQLPagDcto = strSQLPagDcto & "NUMDOC = " & dblNumeroDocumento & ",  "
                End If
                
                If strNotas <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "NOTAS = '" & strNotas & "', "
                End If
                
                If strConcepto <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "CONCEPTO = '" & strConcepto & "', "
                End If
                
                If strDetalle <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "DETALLE = '" & strDetalle & "', "
                End If
                
                If strReferencia <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "REFERENCIA = '" & strReferencia & "', "
                End If
                
                If strCentroCosto <> vbNullString Then
                    strSQLPagDcto = strSQLPagDcto & "F4CENTRO = '" & strCentroCosto & "', "
                End If
                
                strSQLPagDcto = strSQLPagDcto & "ASIENTO = " & dblAsiento & " "
            Else
                strSQLPagDcto = strSQLPagDcto & "SALDO = " & dblSaldoFacturado & " "
            End If
            
        strSQLPagDcto = strSQLPagDcto & "WHERE "
            strSQLPagDcto = strSQLPagDcto & "CORRELA = " & dblCorrelativo
    End If
    
    cnn_dbbancos.Execute strSQLPagDcto
        
    ''cnn_dbbancos.CommitTrans
    
    guardarPagDcto = True
    
    strSQLSelectAlter = strSQLPagDcto
    strSQLPagDcto = vbNullString
    
    Exit Function
errGuardarPagDcto:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsPagDcto: GuardarPagDcto"
    End Select
    
    ''cnn_dbbancos.RollbackTrans
    
    guardarPagDcto = False
    
    Err.Clear
End Function

Public Sub guardarPagMovimiento(ByVal strCodProveedorPM As String, _
                                        ByVal dblCorrelaComprobante As Double, _
                                        ByVal dblCorrelaDocumento As Double, _
                                        ByVal dblImputadoPM As Double, _
                                        ByVal dblTipoCambioPM As Double, _
                                        ByVal strCodMonedaPM As String, _
                                        ByVal lngAnnoReportePM As Long, _
                                        ByVal lngNroReportePM As Long, _
                                        ByVal strFechaMovimiento As String, _
                                        ByVal strFechaReportePM As String)
    
    On Error GoTo errGuardarPagMovimiento
    
    strSQLPagDcto = vbNullString
    
    strSQLPagDcto = strSQLPagDcto & "INSERT INTO PAG_MVTO(PROVEEDOR, CORR_COMP, CORR_DCTO, IMPUTADO, TCAMBIO, IMPUTASO, " & _
                                                                "ANO_REPO, NRO_REPO, FCH_MVTO, FCH_REPO) "
                                                
    strSQLPagDcto = strSQLPagDcto & "VALUES('" & strCodProveedorPM & "', " & dblCorrelaComprobante & ", " & dblCorrelaDocumento & ", " & _
                                                dblImputadoPM & ", " & dblTipoCambioPM & ", FORMAT((" & dblImputadoPM & ")" & _
                                                IIf(strCodMonedaPM = "S", "/", "*") & dblTipoCambioPM & ", '#0.00'), " & lngAnnoReportePM & ", " & _
                                                lngNroReportePM & ", CVDATE('" & strFechaMovimiento & "'), CVDATE('" & strFechaReportePM & "'))"
    
    cnn_dbbancos.Execute strSQLPagDcto
    
    strSQLPagDcto = vbNullString
    
    Exit Sub
errGuardarPagMovimiento:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: GuardarPagMovimiento"
    
    Err.Clear
End Sub

Public Function guardarPagMovimientoV2() As Boolean
    On Error GoTo errGuardarPagMovimientoV2
    
    strSQLPagDcto = vbNullString
    
    'cnn_dbbancos.BeginTrans
    
    If Not verificarExistenciaMovimiento Then
        strSQLPagDcto = strSQLPagDcto & "INSERT INTO PAG_MVTO"
        strSQLPagDcto = strSQLPagDcto & "(PROVEEDOR, CORR_COMP, CORR_DCTO, "
        strSQLPagDcto = strSQLPagDcto & "IMPUTASO, IMPUTADO, TCAMBIO, "
        strSQLPagDcto = strSQLPagDcto & "ANO_REPO, NRO_REPO, FCH_MVTO, "
        strSQLPagDcto = strSQLPagDcto & "FCH_REPO) "
        
        strSQLPagDcto = strSQLPagDcto & "VALUES("
        strSQLPagDcto = strSQLPagDcto & "'" & strMCodProveedor & "', " & dblMCorrelaPago & ", " & dblMCorrelaDocumento & ", "
        strSQLPagDcto = strSQLPagDcto & dblMImputadoSoles & ", " & dblMImputadoDolares & ", " & dblMTipoCambio & ", "
        strSQLPagDcto = strSQLPagDcto & IIf(lngMAnnoReporte = 0, "NULL", lngMAnnoReporte) & ", "
        strSQLPagDcto = strSQLPagDcto & IIf(lngMNumeroReporte = 0, "NULL", lngMNumeroReporte) & ", "
        strSQLPagDcto = strSQLPagDcto & IIf(IsDate(strMFechaMovimiento), "CVDATE('" & strMFechaMovimiento & "')", "NULL") & ", "
        strSQLPagDcto = strSQLPagDcto & IIf(IsDate(strMFechaReporte), "CVDATE('" & strMFechaReporte & "')", "NULL") & ")"
    Else
        strSQLPagDcto = strSQLPagDcto & "UPDATE "
            strSQLPagDcto = strSQLPagDcto & "PAG_MVTO "
        strSQLPagDcto = strSQLPagDcto & "SET "
            
            If strMCodProveedor Then
                strSQLPagDcto = strSQLPagDcto & "PROVEEDOR = '" & strMCodProveedor & "', "
            End If
            
            If dblMImputadoSoles <> 0 Then
                strSQLPagDcto = strSQLPagDcto & "IMPUTASO = " & dblMImputadoSoles & ", "
            End If
            
            If dblMImputadoDolares <> 0 Then
                strSQLPagDcto = strSQLPagDcto & "IMPUTADO = " & dblMImputadoDolares & ", "
            End If
            
            If dblMTipoCambio <> 0 Then
                strSQLPagDcto = strSQLPagDcto & "TCAMBIO = " & dblMTipoCambio & ", "
            End If
            
            If lngMAnnoReporte <> 0 Then
                strSQLPagDcto = strSQLPagDcto & "ANO_REPO = " & IIf(lngMAnnoReporte = 0, "NULL", lngMAnnoReporte) & ", "
            End If
            
            If lngMNumeroReporte <> 0 Then
                strSQLPagDcto = strSQLPagDcto & "NRO_REPO = " & IIf(lngMNumeroReporte = 0, "NULL", lngMNumeroReporte) & ", "
            End If
            
            If IsDate(strMFechaMovimiento) Then
                strSQLPagDcto = strSQLPagDcto & "FCH_MVTO = " & IIf(IsDate(strMFechaMovimiento), "CVDATE('" & strMFechaMovimiento & "')", "NULL") & ", "
            End If
            
            If IsDate(strMFechaReporte) Then
                strSQLPagDcto = strSQLPagDcto & "FCH_REPO = " & IIf(IsDate(strMFechaReporte), "CVDATE('" & strMFechaReporte & "')", "NULL") & ", "
            End If
            
            strSQLPagDcto = strSQLPagDcto & "CHE_DIFE = '' "
            
        strSQLPagDcto = strSQLPagDcto & "WHERE "
            strSQLPagDcto = strSQLPagDcto & "CORR_COMP = " & dblMCorrelaPago & " AND "
            strSQLPagDcto = strSQLPagDcto & "CORR_DCTO = " & dblMCorrelaDocumento
    End If
    
    cnn_dbbancos.Execute strSQLPagDcto
    
    'cnn_dbbancos.CommitTrans
    
    guardarPagMovimientoV2 = True
    
    strSQLSelectAlter = strSQLPagDcto
    strSQLPagDcto = vbNullString
    
    Exit Function
errGuardarPagMovimientoV2:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsPagDcto: GuardarPagMovimientoV2"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    guardarPagMovimientoV2 = False
    
    Err.Clear
End Function

Public Function eliminarPagDcto() As Boolean
    On Error GoTo errEliminarPagDcto
    
'    If Val(validarUsoRegistro(cnn_DbBancos, "CORRELA", "PAG_DCTO", dblCorrelativo, "N")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarPagDcto = False
'
'        Exit Function
'    End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "DELETE "
    strSQLPagDcto = strSQLPagDcto & "FROM PAG_DCTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "CORRELA = " & dblCorrelativo
    
    cnn_dbbancos.Execute strSQLPagDcto
    
    'cnn_dbbancos.CommitTrans
    
    eliminarPagDcto = True
    
    strSQLSelectAlter = strSQLPagDcto
    strSQLPagDcto = vbNullString
    
    Exit Function
errEliminarPagDcto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: EliminarPagDcto"

    'cnn_dbbancos.RollbackTrans
    
    eliminarPagDcto = False
    
    Err.Clear
End Function

Public Function eliminarPagMovimiento() As Boolean
    On Error GoTo errEliminarPagMovimiento
    
'    If Val(validarUsoRegistro(cnn_DbBancos, "CORRELA", "CTA_DCTO", dblCorrelativo, "N")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarPagMovimiento = False
'
'        Exit Function
'    End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "DELETE "
    strSQLPagDcto = strSQLPagDcto & "FROM PAG_MVTO "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "CORR_COMP = " & dblMCorrelaPago & " AND "
    strSQLPagDcto = strSQLPagDcto & "CORR_DCTO = " & dblMCorrelaDocumento
    
    cnn_dbbancos.Execute strSQLPagDcto
    
    'cnn_dbbancos.CommitTrans
    
    eliminarPagMovimiento = True
    
    strSQLSelectAlter = strSQLPagDcto
    strSQLPagDcto = vbNullString
    
    Exit Function
errEliminarPagMovimiento:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: EliminarPagMovimiento"

    'cnn_dbbancos.RollbackTrans
    
    eliminarPagMovimiento = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'Listar Cuentas por Pagar en Grilla (QuamtumGrid)
Public Sub listarGrillaCuentasPorPagar(ByVal grilla As dxDBGrid, _
                                        ByVal intTipoConsulta As Integer, _
                                        ByVal dblValorPorcDetraccionActual As Double, _
                                        ByVal strAnno As String, _
                                        ByVal intMes As String, _
                                        ByVal intTiempoVencimiento As Integer)
    
    On Error GoTo errListarGrillaCuentasPorPagar
    
    strSQLPagDcto = vbNullString
    strSQLPagDcto = strSQLPagDcto & "SELECT "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.CORRELA, "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.MONEDA, "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.FCH_COMP, "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.FCH_VCTO, "
    
    strSQLPagDcto = strSQLPagDcto & "CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO) AS DIAS, "
    
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.NRO_COMP, "
    strSQLPagDcto = strSQLPagDcto & "CLI.F2NEWRUC, "
    strSQLPagDcto = strSQLPagDcto & "CLI.F2NOMCLI, "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.DETALLE, "
    
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', VENTA.F4SUBTOT, NULL) AS BASEIMP_ME, "
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', VENTA.F4SUBTOT, NULL) AS BASEIMP_MN, "
    
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', VENTA.F4TOTIGV, NULL) AS IGV_ME, "
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', VENTA.F4TOTIGV, NULL) AS IGV_MN, "
    
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', PORCOBRAR.TOTAL * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS TOTAL_ME, "
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', PORCOBRAR.TOTAL * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS TOTAL_MN, "
    
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', IIF(ISNULL(PORCOBRAR.DETRACCION), (ABS(TOTAL_ME) * " & dblValorPorcDetraccionActual & ") * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), PORCOBRAR.DETRACCION * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1)), NULL) AS DETRAC_ME, "
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', IIF(ISNULL(PORCOBRAR.DETRACCION), (ABS(TOTAL_MN) * " & dblValorPorcDetraccionActual & ") * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), PORCOBRAR.DETRACCION * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1)), NULL) AS DETRAC_MN, "
    
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', (ABS(TOTAL_ME)-ABS(DETRAC_ME)) * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS NETO_ME, "
    strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', (ABS(TOTAL_MN)-ABS(DETRAC_MN)) * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS NETO_MN, "
        
        Select Case intTipoConsulta
            Case 0 'Ver Pendientes (Importes Netos)
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', (PORCOBRAR.SALDO - PORCOBRAR.DETRACCIONSALDO) * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_ME, "
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', (PORCOBRAR.SALDO - PORCOBRAR.DETRACCIONSALDO) * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_MN "
            Case 1 'Ver Pendientes (Detraccion)
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', PORCOBRAR.DETRACCIONSALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_ME, "
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', PORCOBRAR.DETRACCIONSALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_MN "
            Case 2 'Ver Pendientes (Neto + Detraccion)
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_ME, "
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_MN "
            Case Else 'Por Default
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_ME, "
                strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_MN "
        End Select
        
    'strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'D', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_ME, "
    'strSQLPagDcto = strSQLPagDcto & "IIF(PORCOBRAR.MONEDA = 'S', PORCOBRAR.SALDO * IIF(PORCOBRAR.DEB_HAB = 'D', 1, -1), NULL) AS SALDO_MN "
    
    strSQLPagDcto = strSQLPagDcto & "FROM "
    strSQLPagDcto = strSQLPagDcto & "(PAG_DCTO AS PORCOBRAR "
    strSQLPagDcto = strSQLPagDcto & "LEFT JOIN EF2PROVEEDORS AS CLI ON CLI.F2CODCLI = PORCOBRAR.PROVEEDOR) "
    strSQLPagDcto = strSQLPagDcto & "LEFT JOIN TBVENTA_CAB AS VENTA ON VENTA.F4CORRELA = PORCOBRAR.CORRELA "
    strSQLPagDcto = strSQLPagDcto & "WHERE "
    strSQLPagDcto = strSQLPagDcto & "NOT ISNULL(PORCOBRAR.CORRELA) AND "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.VIA_INGR IN ('1', '5') AND "
    strSQLPagDcto = strSQLPagDcto & "LEFT(PORCOBRAR.NRO_COMP, 1) <> 'L' "
    'strSQLPagDcto = strSQLPagDcto & "AND (PORCOBRAR.SALDO > 0 AND LEFT(PORCOBRAR.NRO_COMP, 1) <> 'L') "
    
        Select Case intTipoConsulta
            Case 0 'Ver Pendientes (Importes Netos)
                strSQLPagDcto = strSQLPagDcto & "AND (PORCOBRAR.SALDO - PORCOBRAR.DETRACCIONSALDO) > 0 "
            Case 1 'Ver Pendientes (Detraccion)
                strSQLPagDcto = strSQLPagDcto & "AND PORCOBRAR.DETRACCIONSALDO > 0 "
            Case 2 'Ver Pendientes (Neto + Detraccion)
                strSQLPagDcto = strSQLPagDcto & "AND PORCOBRAR.SALDO > 0 "
            Case 3 'Ver Canceladas
                strSQLPagDcto = strSQLPagDcto & "AND PORCOBRAR.SALDO = 0 "
            Case Else 'Ver Todas
                'Mostrar toda la información
                'No es necesario ningun filtro adicional
        End Select
        
        If IsNumeric(strAnno) Then
            strSQLPagDcto = strSQLPagDcto & "AND YEAR(PORCOBRAR.FCH_COMP) <= '" & strAnno & "' "
        End If
        
        If intMes > 0 Then
            strSQLPagDcto = strSQLPagDcto & "AND CVDATE(PORCOBRAR.FCH_COMP) <= CVDATE('" & DateSerial(Val(strAnno), intMes + 1, 0) & "') "
        End If
        
        Select Case intTiempoVencimiento
            Case 1 'Por Vencer
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) <= 0 "
            Case 2 'Vencidas
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) > 0 "
            Case 3 'Venc. 30 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) BETWEEN 1 AND 30 "
            Case 4 'Venc. 60 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) BETWEEN 31 AND 60 "
            Case 5 'Venc. 90 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) BETWEEN 61 AND 90 "
            Case 6 'Venc. 120 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) BETWEEN 91 AND 120 "
            Case 7 'Venc. 180 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) BETWEEN 121 AND 180 "
            Case 8 'Venc. > 180 días
                strSQLPagDcto = strSQLPagDcto & "AND (CVDATE(DATE())-CVDATE(PORCOBRAR.FCH_VCTO)) > 180 "
        End Select
        
    strSQLPagDcto = strSQLPagDcto & "ORDER BY "
    strSQLPagDcto = strSQLPagDcto & "PORCOBRAR.FCH_COMP, PORCOBRAR.NRO_COMP"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Correlativo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Correlativo"
                .DisableEditor = True
                .FieldName = "CORRELA"
                .DecimalPlaces = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCorrelativo"
                .Width = 70
                .Visible = False
            End With
            
            'Columna Moneda
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Moneda"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "MONEDA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColMoneda"
                .Width = 80
                .Visible = False
            End With
            
            'Columna Fecha Comprobante
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fec. Emisión"
                .DisableEditor = True
                .FieldName = "FCH_COMP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaComprobante"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
        
            'Columna Fecha Vencimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Fec. Vcto"
                .DisableEditor = True
                .FieldName = "FCH_VCTO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColFechaVencimiento"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantida de Días de Vencimiento
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Días"
                .DisableEditor = True
                .FieldName = "DIAS"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDias"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 45
            End With
            
            'Columna Numero Comprobante
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Documento"
                .DisableEditor = True
                .FieldName = "NRO_COMP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNumeroComprobante"
                .SummaryFooterType = cstCount
                '.SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna RUC de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "RUC"
                .DisableEditor = True
                .FieldName = "F2NEWRUC"
                .HeaderAlignment = taCenter
                .ObjectName = "ColRUC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 75
            End With
            
            'Columna Nombre de Cliente
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Cliente"
                .DisableEditor = True
                .FieldName = "F2NOMCLI"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCliente"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
            End With
            
            'Columna Concepto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Concepto"
                .DisableEditor = True
                .FieldName = "DETALLE"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDetalle"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
            End With
            
            'Columna Base Imponible MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "BASEIMP_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColBaseImponibleMN"
                .Width = 70
            End With
            
            'Columna Base Imponible ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "BASEIMP_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColBaseImponibleME"
                .Width = 70
            End With
            
            'Columna IGV MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "IGV_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColIGVMN"
                .Width = 70
            End With
            
            'Columna IGV ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "IGV_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColIGVME"
                .Width = 70
            End With
            
            'Columna Total MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTAL_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColTotalMN"
                .Width = 70
            End With
            
            'Columna Total ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "TOTAL_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColTotalME"
                .Width = 70
            End With
            
            'Columna Detraccion MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "DETRAC_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColDetraccionMN"
                .Width = 70
            End With
            
            'Columna Detraccion ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "DETRAC_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColDetraccionME"
                .Width = 70
            End With
            
            'Columna Importe Neto MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "NETO_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColImporteNetoMN"
                .Width = 70
            End With
            
            'Columna Importe Neto ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "NETO_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColImporteNetoME"
                .Width = 70
            End With
            
            'Columna Saldo MN
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 6
                .Caption = "MN"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO_MN"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColSaldoMN"
                .Width = 70
            End With
            
            'Columna Saldo ME
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 6
                .Caption = "ME"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO_ME"
                .HeaderAlignment = taCenter
                .SummaryFooterType = cstSum
                .ObjectName = "ColSaldoME"
                .Width = 70
            End With
            
            Select Case intTipoConsulta
                Case 0 'Ver Pendientes (Importe Neto)
                    '.Columns.ColumnByFieldName("DETRAC_MN").Visible = False
                    '.Columns.ColumnByFieldName("DETRAC_ME").Visible = False
                    .Bands(4).Visible = False
                    .Bands(5).Visible = True
                    .Bands(6).Caption = "Saldo Neto por Cobrar"
                Case 1 'Ver Pendientes (Detracciones)
                    '.Columns.ColumnByFieldName("NETO_MN").Visible = False
                    '.Columns.ColumnByFieldName("NETO_ME").Visible = False
                    .Bands(4).Visible = True
                    .Bands(5).Visible = False
                    .Bands(6).Caption = "Saldo de Detraccion"
                Case 2 'Ver Pendientes (Neto + Detraccion)
                    '.Columns.ColumnByFieldName("DETRAC_MN").Visible = False
                    '.Columns.ColumnByFieldName("DETRAC_ME").Visible = False
                    '.Columns.ColumnByFieldName("NETO_MN").Visible = False
                    '.Columns.ColumnByFieldName("NETO_ME").Visible = False
                    .Bands(4).Visible = True
                    .Bands(5).Visible = True
                    .Bands(6).Caption = "Saldo Total por Cobrar"
                Case Else 'Ver Canceladas, Ver Todas
                    'Dejar la Configuración Por Default
                    .Bands(4).Visible = True
                    .Bands(5).Visible = True
                    .Bands(6).Caption = "Saldo por Cobrar"
            End Select
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexDbBancos

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLPagDcto
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "NRO_COMP"
            .M.FullExpand
        End With
    End If
    
    strSQLPagDcto = vbNullString
    
    Exit Sub
errListarGrillaCuentasPorPagar:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsPagDcto: ListarGrillaCuentasPorPagar"
    End Select
    
    Err.Clear
End Sub

Public Sub listarAnnoCuentasPorPagar(ByVal combo As Object)
    On Error GoTo errListarAnnoCuentasPorPagar
    
    Set rstPagDcto = New ADODB.Recordset
    
    strSQLPagDcto = "SELECT YEAR(FCH_COMP) AS ANNO " & _
                    "FROM PAG_DCTO " & _
                    "GROUP BY YEAR(FCH_COMP)"
    
    If rstPagDcto.State = 1 Then rstPagDcto.Close
    
    rstPagDcto.Open strSQLPagDcto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstPagDcto.EOF Then
        combo.Clear
        combo.AddItem "(*) - Todos"
        
        rstPagDcto.MoveFirst
        
        Do While Not rstPagDcto.EOF
            combo.AddItem rstPagDcto!Anno
            
            rstPagDcto.MoveNext
        Loop
    End If
    
    rstPagDcto.Close
    strSQLPagDcto = vbNullString
    
    Set rstPagDcto = Nothing
    
    Exit Sub
errListarAnnoCuentasPorPagar:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsPagDcto: ListarAnnoCuentasPorPagar"
    
    Err.Clear
End Sub

