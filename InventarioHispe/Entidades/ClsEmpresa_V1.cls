VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodEmpresa               As String
Private strCtaContaVenta            As String
Private strOrigenMovimiento         As String
Private strCtaFacXCobrar            As String
Private strCtaFacXCobrarDol         As String
Private bolEvaluarCtaFacXCobrar     As Boolean
Private bolAnexarDocReferencia      As Boolean
Private strCtaContaIgv              As String
Private strCtaContaDscto            As String
Private bolUsarTipoCambioOficial    As Boolean

Private dblDetraccionPorcentaje     As Double
Private bolDetraccionImpMsj         As Boolean
Private strDetraccionMensaje        As String
Private strDetraccionCtaSoles       As String
Private strDetraccionCtaDolar       As String

Private strFormatoNumDocVenta       As String
Private strFormatoNumDocCompra      As String
Private lngCodSistemaContable       As Long

Rem SK: Atributos adicionales para Compras
Private strCtaContaFonavi           As String
Private strCtaContaRetencion        As String
Private strCtaContaOtrosImpuestos   As String
Private strCtaContaRedSuma          As String
Private strCtaContaRedResta         As String
Private strNumeroObra               As String
Private strCtaContaDetraccion       As String

Rem SK: Atributos adicionales para Bancos
Private strCtaContaPendiente        As String
Private strTipoReferencia           As String
Private strCtaContaGanancia         As String
Private strCtaContaPerdida          As String
Private strCodTipoGenCompro         As String
Private strOrigenIngreso            As String
Private strOrigenEgreso             As String
Private bolAgruparAsientos          As Boolean
Private strOrigenAcumulado          As String
Private strComproAcumulado          As String

Private bolEstado                As Boolean

Private strSQLSelectAlter   As String

Private rstEmpresa             As ADODB.Recordset
Private strSQLEmpresa          As String

Public Property Let CodigoEmpresa(ByVal Value As String)
    strCodEmpresa = Value
End Property

Public Property Get CodigoEmpresa() As String
    CodigoEmpresa = strCodEmpresa
End Property

Public Property Let CtaContaVenta(ByVal Value As String)
    strCtaContaVenta = Value
End Property

Public Property Get CtaContaVenta() As String
    CtaContaVenta = strCtaContaVenta
End Property

Public Property Let OrigenMovimiento(ByVal Value As String)
    strOrigenMovimiento = Value
End Property

Public Property Get OrigenMovimiento() As String
    OrigenMovimiento = strOrigenMovimiento
End Property

Public Property Let CtaFacXCobrar(ByVal Value As String)
    strCtaFacXCobrar = Value
End Property

Public Property Get CtaFacXCobrar() As String
    CtaFacXCobrar = strCtaFacXCobrar
End Property

Public Property Let CtaFacXCobrarDol(ByVal Value As String)
    strCtaFacXCobrarDol = Value
End Property

Public Property Get CtaFacXCobrarDol() As String
    CtaFacXCobrarDol = strCtaFacXCobrarDol
End Property

Public Property Let EvaluarCtaFacXCobrar(ByVal Value As Boolean)
    bolEvaluarCtaFacXCobrar = Value
End Property

Public Property Get EvaluarCtaFacXCobrar() As Boolean
    EvaluarCtaFacXCobrar = bolEvaluarCtaFacXCobrar
End Property

Public Property Let AnexarDocReferencia(ByVal Value As Boolean)
    bolAnexarDocReferencia = Value
End Property

Public Property Get AnexarDocReferencia() As Boolean
    AnexarDocReferencia = bolAnexarDocReferencia
End Property

Public Property Let CtaContaIgv(ByVal Value As String)
    strCtaContaIgv = Value
End Property

Public Property Get CtaContaIgv() As String
    CtaContaIgv = strCtaContaIgv
End Property

Public Property Let CtaContaDscto(ByVal Value As String)
    strCtaContaDscto = Value
End Property

Public Property Get CtaContaDscto() As String
    CtaContaDscto = strCtaContaDscto
End Property

Public Property Let UsarTipoCambioOficial(ByVal Value As Boolean)
    bolUsarTipoCambioOficial = Value
End Property

Public Property Get UsarTipoCambioOficial() As Boolean
    UsarTipoCambioOficial = bolUsarTipoCambioOficial
End Property

'DetraccionPorcentaje
Public Property Let DetraccionPorcentaje(ByVal Value As Double)
    dblDetraccionPorcentaje = Value
End Property

Public Property Get DetraccionPorcentaje() As Double
    DetraccionPorcentaje = dblDetraccionPorcentaje
End Property

'DetraccionImpMsj
Public Property Let DetraccionImpMsj(ByVal Value As Boolean)
    bolDetraccionImpMsj = Value
End Property

Public Property Get DetraccionImpMsj() As Boolean
    DetraccionImpMsj = bolDetraccionImpMsj
End Property

'DetraccionMensaje
Public Property Let DetraccionMensaje(ByVal Value As String)
    strDetraccionMensaje = Value
End Property

Public Property Get DetraccionMensaje() As String
    DetraccionMensaje = strDetraccionMensaje
End Property

'DetraccionCtaSoles
Public Property Let DetraccionCtaSoles(ByVal Value As String)
    strDetraccionCtaSoles = Value
End Property

Public Property Get DetraccionCtaSoles() As String
    DetraccionCtaSoles = strDetraccionCtaSoles
End Property

'DetraccionCtaDolar
Public Property Let DetraccionCtaDolar(ByVal Value As String)
    strDetraccionCtaDolar = Value
End Property

Public Property Get DetraccionCtaDolar() As String
    DetraccionCtaDolar = strDetraccionCtaDolar
End Property



'FormatoNumDocVenta
Public Property Let FormatoNumDocVenta(ByVal Value As String)
    strFormatoNumDocVenta = Value
End Property

Public Property Get FormatoNumDocVenta() As String
    FormatoNumDocVenta = strFormatoNumDocVenta
End Property
'FormatoNumDocCompra
Public Property Let FormatoNumDocCompra(ByVal Value As String)
    strFormatoNumDocCompra = Value
End Property

Public Property Get FormatoNumDocCompra() As String
    FormatoNumDocCompra = strFormatoNumDocCompra
End Property


'CodSistemaContable
Public Property Let CodSistemaContable(ByVal Value As Long)
    lngCodSistemaContable = Value
End Property

Public Property Get CodSistemaContable() As Long
    CodSistemaContable = lngCodSistemaContable
End Property


Rem SK: Atributos adicionales para Compras

Public Property Let CtaContaFonavi(ByVal Value As String)
    strCtaContaFonavi = Value
End Property

Public Property Get CtaContaFonavi() As String
    CtaContaFonavi = strCtaContaFonavi
End Property

Public Property Let CtaContaRetencion(ByVal Value As String)
    strCtaContaRetencion = Value
End Property

Public Property Get CtaContaRetencion() As String
    CtaContaRetencion = strCtaContaRetencion
End Property

Public Property Let CtaContaOtrosImpuestos(ByVal Value As String)
    strCtaContaOtrosImpuestos = Value
End Property

Public Property Get CtaContaOtrosImpuestos() As String
    CtaContaOtrosImpuestos = strCtaContaOtrosImpuestos
End Property

Public Property Let CtaContaRedSuma(ByVal Value As String)
    strCtaContaRedSuma = Value
End Property

Public Property Get CtaContaRedSuma() As String
    CtaContaRedSuma = strCtaContaRedSuma
End Property

Public Property Let CtaContaRedResta(ByVal Value As String)
    strCtaContaRedResta = Value
End Property

Public Property Get CtaContaRedResta() As String
    CtaContaRedResta = strCtaContaRedResta
End Property
'NumeroObra
Public Property Let NumeroObra(ByVal Value As String)
    strNumeroObra = Value
End Property

Public Property Get NumeroObra() As String
    NumeroObra = strNumeroObra
End Property
'CtaContaDetraccion
Public Property Let CtaContaDetraccion(ByVal Value As String)
    strCtaContaDetraccion = Value
End Property

Public Property Get CtaContaDetraccion() As String
    CtaContaDetraccion = strCtaContaDetraccion
End Property



Rem SK: Atributos adicionales para Bancos

Public Property Let CtaContaPendiente(ByVal Value As String)
    strCtaContaPendiente = Value
End Property

Public Property Get CtaContaPendiente() As String
    CtaContaPendiente = strCtaContaPendiente
End Property

Public Property Let TipoReferencia(ByVal Value As String)
    strTipoReferencia = Value
End Property

Public Property Get TipoReferencia() As String
    TipoReferencia = strTipoReferencia
End Property

Public Property Let CtaContaGanancia(ByVal Value As String)
    strCtaContaGanancia = Value
End Property

Public Property Get CtaContaGanancia() As String
    CtaContaGanancia = strCtaContaGanancia
End Property

Public Property Let CtaContaPerdida(ByVal Value As String)
    strCtaContaPerdida = Value
End Property

Public Property Get CtaContaPerdida() As String
    CtaContaPerdida = strCtaContaPerdida
End Property

Public Property Let CodTipoGenCompro(ByVal Value As String)
    strCodTipoGenCompro = Value
End Property

Public Property Get CodTipoGenCompro() As String
    CodTipoGenCompro = strCodTipoGenCompro
End Property

Public Property Let OrigenIngreso(ByVal Value As String)
    strOrigenIngreso = Value
End Property

Public Property Get OrigenIngreso() As String
    OrigenIngreso = strOrigenIngreso
End Property

Public Property Let OrigenEgreso(ByVal Value As String)
    strOrigenEgreso = Value
End Property

Public Property Get OrigenEgreso() As String
    OrigenEgreso = strOrigenEgreso
End Property

Public Property Let AgruparAsientos(ByVal Value As Boolean)
    bolAgruparAsientos = Value
End Property

Public Property Get AgruparAsientos() As Boolean
    AgruparAsientos = bolAgruparAsientos
End Property

Public Property Let OrigenAcumulado(ByVal Value As String)
    strOrigenAcumulado = Value
End Property

Public Property Get OrigenAcumulado() As String
    OrigenAcumulado = strOrigenAcumulado
End Property

Public Property Let ComproAcumulado(ByVal Value As String)
    strComproAcumulado = Value
End Property

Public Property Get ComproAcumulado() As String
    ComproAcumulado = strComproAcumulado
End Property








Public Property Let estado(ByVal Value As Boolean)
    bolEstado = Value
End Property

Public Property Get estado() As Boolean
    estado = bolEstado
End Property


Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property

Public Sub inicializarEntidades()
    strCodEmpresa = vbNullString
    strCtaContaVenta = vbNullString
    strOrigenMovimiento = vbNullString
    strCtaFacXCobrar = vbNullString
    strCtaFacXCobrarDol = vbNullString
    bolEvaluarCtaFacXCobrar = False
    bolAnexarDocReferencia = False
    strCtaContaIgv = vbNullString
    strCtaContaDscto = vbNullString
    bolUsarTipoCambioOficial = False
    
    dblDetraccionPorcentaje = 0
    bolDetraccionImpMsj = False
    strDetraccionMensaje = vbNullString
    strDetraccionCtaSoles = vbNullString
    strDetraccionCtaDolar = vbNullString
    
    strFormatoNumDocVenta = vbNullString
    strFormatoNumDocCompra = vbNullString
    
    lngCodSistemaContable = 0
    
    Rem SK: Atributos adicionales para Compras
    strCtaContaFonavi = vbNullString
    strCtaContaRetencion = vbNullString
    strCtaContaOtrosImpuestos = vbNullString
    strCtaContaRedSuma = vbNullString
    strCtaContaRedResta = vbNullString
    strNumeroObra = vbNullString
    strCtaContaDetraccion = vbNullString
    
    Rem SK: Atributos adicionales para Bancos
    strCtaContaPendiente = vbNullString
    strTipoReferencia = vbNullString
    strCtaContaGanancia = vbNullString
    strCtaContaPerdida = vbNullString
    strCodTipoGenCompro = vbNullString
    strOrigenIngreso = vbNullString
    strOrigenEgreso = vbNullString
    bolAgruparAsientos = False
    strOrigenAcumulado = vbNullString
    strComproAcumulado = vbNullString
    
    bolEstado = False
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerEmpresa() As Boolean
    On Error GoTo errObtenerEmpresa
    Set rstEmpresa = New ADODB.Recordset
    
    strSQLEmpresa = "SELECT * " & _
                    "FROM SF1PARCONTA " & _
                    "WHERE F1CODEMP = '" & strCodEmpresa & "'"
    
    If rstEmpresa.State = 1 Then rstEmpresa.Close
    
    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenForwardOnly, adLockReadOnly
    
    If Not rstEmpresa.EOF Then
        strCodEmpresa = Trim(rstEmpresa!F1CODEMP & "")
        strCtaContaVenta = Trim(rstEmpresa!CODCTAVDI & "")
        strOrigenMovimiento = Trim(rstEmpresa!ORIFAC & "")
        strCtaFacXCobrar = Trim(rstEmpresa!FACCOB & "")
        strCtaFacXCobrarDol = Trim(rstEmpresa!FACCOBDOL & "")
        bolEvaluarCtaFacXCobrar = CBool(Val(Trim(rstEmpresa!CLIENTE & "")))
        bolAnexarDocReferencia = CBool(Val(Trim(rstEmpresa!REFERE & "")))
        strCtaContaIgv = Trim(rstEmpresa!CODCTAIGV & "")
        strCtaContaDscto = Trim(rstEmpresa!CTADSCTO & "")
        bolUsarTipoCambioOficial = CBool(Val(Trim(rstEmpresa!TCOFICIAL & "")))
        
        dblDetraccionPorcentaje = Val(rstEmpresa!F1DETRACCIONPORC & "")
        bolDetraccionImpMsj = CBool(rstEmpresa!F1DETRACCIONIMPMSJ)
        strDetraccionMensaje = Trim(rstEmpresa!F1DETRACCIONMENSAJE & "")
        strDetraccionCtaSoles = Trim(rstEmpresa!F1DETRACCIONCTASOLES & "")
        strDetraccionCtaDolar = Trim(rstEmpresa!F1DETRACCIONCTADOLAR & "")
        
        strFormatoNumDocVenta = Trim(rstEmpresa!F1FORMATONUMDOCVENTA & "")
        strFormatoNumDocCompra = Trim(rstEmpresa!F1FORMATONUMDOCCOMPRA & "")
        
        lngCodSistemaContable = Trim(rstEmpresa!F1SISTEMACONTABLE & "")
        
        obtenerEmpresa = True
    Else
        obtenerEmpresa = False
    End If
    
    rstEmpresa.Close
    
    Set rstEmpresa = Nothing
    
    Exit Function
errObtenerEmpresa:
    Select Case Err.Number
        Case 3704, 3709
            cnn_control.Open cconex_control  'StrConexControl
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsEmpresa: ObtenerEmpresa"
    End Select
    
    obtenerEmpresa = False
    
    Err.Clear
End Function

'Para Sistema de Ventas
Public Sub obtenerConfigEmpresa()
    On Error GoTo errObtenerConfigEmpresa
    
    Set rstEmpresa = New ADODB.Recordset
    
    strSQLEmpresa = "SELECT * " & _
                    "FROM SF1PARCONTA " & _
                    "WHERE F1CODEMP = '" & strCodEmpresa & "'"
    
    If rstEmpresa.State = 1 Then rstEmpresa.Close
    
    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenForwardOnly, adLockReadOnly
    
    If Not rstEmpresa.EOF Then
        strCodEmpresa = Trim(rstEmpresa!F1CODEMP & "")
        strCtaContaVenta = Trim(rstEmpresa!CODCTAVDI & "")
        strOrigenMovimiento = Trim(rstEmpresa!ORIFAC & "")
        strCtaFacXCobrar = Trim(rstEmpresa!FACCOB & "")
        strCtaFacXCobrarDol = Trim(rstEmpresa!FACCOBDOL & "")
        bolEvaluarCtaFacXCobrar = CBool(Val(Trim(rstEmpresa!CLIENTE & "")))
        bolAnexarDocReferencia = CBool(Val(Trim(rstEmpresa!REFERE & "")))
        strCtaContaIgv = Trim(rstEmpresa!CODCTAIGV & "")
        strCtaContaDscto = Trim(rstEmpresa!CTADSCTO & "")
        bolUsarTipoCambioOficial = CBool(Val(Trim(rstEmpresa!TCOFICIAL & "")))
        
        dblDetraccionPorcentaje = Val(rstEmpresa!F1DETRACCIONPORC & "")
        bolDetraccionImpMsj = CBool(rstEmpresa!F1DETRACCIONIMPMSJ)
        strDetraccionMensaje = Trim(rstEmpresa!F1DETRACCIONMENSAJE & "")
        strDetraccionCtaSoles = Trim(rstEmpresa!F1DETRACCIONCTASOLES & "")
        strDetraccionCtaDolar = Trim(rstEmpresa!F1DETRACCIONCTADOLAR & "")
        
        strFormatoNumDocVenta = Trim(rstEmpresa!F1FORMATONUMDOCVENTA & "")
        strFormatoNumDocCompra = Trim(rstEmpresa!F1FORMATONUMDOCCOMPRA & "")
        
        lngCodSistemaContable = Val(rstEmpresa!F1SISTEMACONTABLE & "")
    End If
    
    rstEmpresa.Close
    
    Set rstEmpresa = Nothing
    
    Exit Sub
errObtenerConfigEmpresa:
    Select Case Err.Number
        Case 3704, 3709 'Perdida de Conexion o Objeto de Conexion Cerrada
            cnn_control.Open cconex_control  'StrConexControl
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsEmpresa: ObtenerConfigEmpresa"
    End Select
        
    Err.Clear
End Sub

'Para Sistema de Compras
Public Sub obtenerConfigEmpresaCompra()
    On Error GoTo errObtenerConfigEmpresaCompra
    
    Set rstEmpresa = New ADODB.Recordset
    
    strSQLEmpresa = "SELECT * " & _
                    "FROM PARAM_COM " & _
                    "WHERE F1CODEMP = '" & strCodEmpresa & "'"
    
    If rstEmpresa.State = 1 Then rstEmpresa.Close
    
    rstEmpresa.Open strSQLEmpresa, cnn_ControlCompra, adOpenForwardOnly, adLockReadOnly
    
    If Not rstEmpresa.EOF Then
        strCodEmpresa = Trim(rstEmpresa!F1CODEMP & "")
        strOrigenMovimiento = Trim(rstEmpresa!F1ORIGEN & "")
        strCtaContaIgv = Trim(rstEmpresa!F1CTAIGV & "")
        strCtaContaFonavi = Trim(rstEmpresa!F1CTAFONAVI & "")
        strCtaContaRetencion = Trim(rstEmpresa!F1CTARET & "")
        strCtaContaOtrosImpuestos = Trim(rstEmpresa!F1CTAOTROS & "")
        strCtaContaRedSuma = Trim(rstEmpresa!F1REDSUMA & "")
        strCtaContaRedResta = Trim(rstEmpresa!F1REDRESTA & "")
        strNumeroObra = Trim(rstEmpresa!F1NUMOBRA & "")
        
        strCtaContaDetraccion = Trim(rstEmpresa!F1CTADETRACCION & "")
    End If
    
    rstEmpresa.Close
    
    Set rstEmpresa = Nothing
    
    Exit Sub
errObtenerConfigEmpresaCompra:
    Select Case Err.Number
        Case 3704, 3709
            cnn_ControlCompra.Open StrConexControlCompra
            
            Resume
        Case 3265 'Ordinal no encontrado en Recordset
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsEmpresa: ObtenerConfigEmpresaCompra"
    End Select
        
    Err.Clear
End Sub

'Para Sistema de Bancos
Public Sub obtenerConfigEmpresaBancos()
    On Error GoTo errObtenerConfigEmpresaBancos
    
    Set rstEmpresa = New ADODB.Recordset
    
    strSQLEmpresa = "SELECT * " & _
                    "FROM BF1CNT " & _
                    "WHERE F1DIR = '" & strCodEmpresa & "'"
    
    If rstEmpresa.State = 1 Then rstEmpresa.Close
    
    rstEmpresa.Open strSQLEmpresa, cnn_ControlBanco, adOpenForwardOnly, adLockReadOnly
    
    If Not rstEmpresa.EOF Then
        strCodEmpresa = Trim(rstEmpresa!F1DIR & "")
        
        strCtaContaPendiente = Trim(rstEmpresa!F1CTAPEND & "")
        strTipoReferencia = Trim(rstEmpresa!F1REFER & "")
        strCtaContaGanancia = Trim(rstEmpresa!F1CTACAMG & "")
        strCtaContaPerdida = Trim(rstEmpresa!F1CTACAMP & "")
        strCodTipoGenCompro = Trim(rstEmpresa!F1COMPROBA & "")
        strOrigenIngreso = Trim(rstEmpresa!F1ORIGENING & "")
        strOrigenEgreso = Trim(rstEmpresa!F1ORIGENEGR & "")
        bolAgruparAsientos = IIf(Trim(rstEmpresa!F1TRANS_ACUM & "") = vbNullString, False, True)
        strOrigenAcumulado = Trim(rstEmpresa!F1ORIGEN_ACUM & "")
        strComproAcumulado = Trim(rstEmpresa!F1COMPRO_ACUM & "")
    End If
    
    rstEmpresa.Close
    
    Set rstEmpresa = Nothing
    
    Exit Sub
errObtenerConfigEmpresaBancos:
    Select Case Err.Number
        Case 3704, 3709
            cnn_ControlBanco.Open StrConexControlBanco
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsEmpresa: ObtenerConfigEmpresaBancos"
    End Select
        
    Err.Clear
End Sub

'Public Function generarCodigoEmpresa() As String
'    On Error GoTo errGenerarCodigoEmpresa
'    Set rstEmpresa = New ADODB.Recordset
'
'    strSQLEmpresa = "SELECT TOP 1 F1CODEMP " & _
'                    "FROM SF1PARCONTA " & _
'                    "ORDER BY F1CODEMP DESC"
'
'    If rstEmpresa.State = 1 Then rstEmpresa.Close
'
'    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenDynamic, adLockBatchOptimistic
'
'    If Not rstEmpresa.EOF Then
'        generarCodigoEmpresa = Format(Val(Trim(rstEmpresa!CodEmpresa & "")) + 1, "0000")
'    Else
'        generarCodigoEmpresa = "0001"
'    End If
'
'    rstEmpresa.Close
'
'    Set rstEmpresa = Nothing
'    strSQLEmpresa = vbNullString
'
'    Exit Function
'errGenerarCodigoEmpresa:
'    generarCodigoEmpresa = vbNullString
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: GenerarCodigoEmpresa"
'
'    Err.Clear
'End Function
'
'Public Function verificarExistencia() As Boolean
'    On Error GoTo errVerificarExistencia
'    Set rstEmpresa = New ADODB.Recordset
'
'    strSQLEmpresa = "SELECT SF1PARCONTA.* " & _
'                    "FROM SF1PARCONTA " & _
'                    "WHERE F1CODEMP = '" & strCodEmpresa & "'"
'
'    If rstEmpresa.State = 1 Then rstEmpresa.Close
'
'    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenDynamic, adLockBatchOptimistic
'
'    If Not rstEmpresa.EOF Then
'        verificarExistencia = True
'    Else
'        verificarExistencia = False
'    End If
'
'    rstEmpresa.Close
'    strSQLEmpresa = vbNullString
'
'    Set rstEmpresa = Nothing
'
'    Exit Function
'errVerificarExistencia:
'    verificarExistencia = False
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: VerificarExistencia"
'
'    Err.Clear
'End Function
'
'Public Function guardarEmpresa() As Boolean
'    On Error GoTo errGuardarEmpresa
'
'    cnn_control.BeginTrans
'
'    If Not verificarExistencia Then
'        strCodEmpresa = generarCodigoEmpresa
'
'        strSQLEmpresa = "INSERT INTO SF1PARCONTA(F1CODEMP, DESBASE, CUENTASUBBASE, ESTADO, FECREG, USUREG) " & _
'                "VALUES('" & strCodEmpresa & "','" & strCtaContaVenta & "'," & _
'                IIf(bolEvaluarCtaFacXCobrar, "TRUE", "FALSE") & ", " & IIf(bolEstado, "TRUE", "FALSE") & _
'                ", CVDATE('" & Format(Now, "Short Date") & "'), '" & datosUser.codUser & "')"
'
'        cnn_control.Execute strSQLEmpresa
'
'        cnn_control.CommitTrans
'    Else
'        strSQLEmpresa = "UPDATE SF1PARCONTA " & _
'                        "SET DESBASE = '" & strCtaContaVenta & "', " & _
'                            " CUENTASUBBASE = " & IIf(bolEvaluarCtaFacXCobrar, "TRUE, ", "FALSE, ") & _
'                            "ESTADO = " & IIf(bolEstado, "TRUE, ", "FALSE, ") & _
'                            "FECMOD = CVDATE('" & Format(Now, "Short Date") & "'), " & _
'                            "USUMOD = '" & datosUser.codUser & "' " & _
'                        "WHERE F1CODEMP = '" & strCodEmpresa & "'"
'
'        cnn_control.Execute strSQLEmpresa
'
'        cnn_control.CommitTrans
'    End If
'
'    If bolEvaluarCtaFacXCobrar Then
'        guardarSubEmpresas
'    End If
'
'    guardarEmpresa = True
'
'    strSQLEmpresa = vbNullString
'
'    Exit Function
'errGuardarEmpresa:
'    cnn_control.RollbackTrans
'
'    guardarEmpresa = False
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: GuardarEmpresa"
'
'    Err.Clear
'End Function
'
'Private Sub guardarSubEmpresas()
'    On Error GoTo errGuardarSubEmpresas
'
'    cnn_control.BeginTrans
'
'    Set rstEmpresa = New ADODB.Recordset
'
'    rstEmpresa.Open strSQLSelectAlter, cnn_controlTemp, adOpenDynamic, adLockBatchOptimistic
'
'    If Not rstEmpresa.EOF Then
'        cnn_control.Execute "DELETE FROM SF1PARCONTA_SUB WHERE F1CODEMP = '" & strCodEmpresa & "'"
'
'        rstEmpresa.MoveFirst
'
'        Do While Not rstEmpresa.EOF
'            cnn_control.Execute "INSERT INTO SF1PARCONTA_SUB(F1CODEMP, CODSUBBASE, DESSUBBASE) " & _
'                            "VALUES('" & strCodEmpresa & "', '" & rstEmpresa!CodSubEmpresa & "', '" & rstEmpresa!DESSUBBASE & "')"
'
'            rstEmpresa.MoveNext
'        Loop
'    End If
'
'    rstEmpresa.Close
'
'    cnn_control.CommitTrans
'
'    Set rstEmpresa = Nothing
'
'    Exit Sub
'errGuardarSubEmpresas:
'    cnn_control.RollbackTrans
'
'    MsgBox "Nro.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbCritical + vbOKOnly, App.ProductName
'
'    Err.Clear
'End Sub
'
'
'Public Function eliminarEmpresa() As Boolean
'On Error GoTo errEliminarEmpresa
'
'    If Val(validarUsoRegistro(cnn_control, "F1CODEMP", "EF2INDICADOR_METAS", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "EF2EMPLEADOS", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "EF2RESPONSABLE", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "IF3ACCIDENTABILIDAD", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "IF4BASEINDICADORES", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "IF4PLANANUAL", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "SF1FSQ", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "SF1PREMIACION", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "SF1RATIOS", strCodEmpresa, "T")) > 0 Or _
'        Val(validarUsoRegistro(cnn_control, "F1CODEMP", "EF2EMP_HEADCOUNT", strCodEmpresa, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarEmpresa = False
'
'        Exit Function
'    End If
'
'    cnn_control.BeginTrans
'
'    strSQLEmpresa = "DELETE FROM SF1PARCONTA " & _
'                "WHERE F1CODEMP = '" & strCodEmpresa & "'"
'
'    cnn_control.Execute strSQLEmpresa
'
'    strSQLEmpresa = "DELETE FROM SF1PARCONTA_SUB " & _
'                "WHERE F1CODEMP = '" & strCodEmpresa & "'"
'
'    cnn_control.Execute strSQLEmpresa
'
'    cnn_control.CommitTrans
'
'    eliminarEmpresa = True
'
'    strSQLEmpresa = vbNullString
'
'    Exit Function
'errEliminarEmpresa:
'    eliminarEmpresa = False
'
'    cnn_control.RollbackTrans
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: EliminarEmpresa"
'
'    Err.Clear
'End Function
'
'Public Sub listarEmpresas(ByVal comboList As Object)
'    On Error GoTo errListarEmpresas
'
'    Set rstEmpresa = New ADODB.Recordset
'
'    strSQLEmpresa = "SELECT F1CODEMP, DESBASE " & _
'                    "FROM SF1PARCONTA " & _
'                    "WHERE ESTADO = TRUE " & _
'                    "ORDER BY F1CODEMP"
'
'    If rstEmpresa.State = 1 Then rstEmpresa.Close
'
'    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstEmpresa.EOF Then
'        comboList.Clear
'        comboList.AddItem "(*) - Todas"
'
'        Do While Not rstEmpresa.EOF
'            comboList.AddItem Trim(rstEmpresa!DESBASE & "") & Space(150) & Trim(rstEmpresa!CodEmpresa & "")
'
'            rstEmpresa.MoveNext
'        Loop
'            comboList.ListIndex = 0
'    End If
'
'    rstEmpresa.Close
'    strSQLEmpresa = vbNullString
'
'    Set rstEmpresa = Nothing
'
'    Exit Sub
'errListarEmpresas:
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: ListarEmpresas"
'
'    Err.Clear
'End Sub
'
'Public Sub listarEmpresasSoloSeleccion(ByVal comboList As ComboBox)
'    On Error GoTo errListarEmpresasSoloSeleccion
'    Set rstEmpresa = New ADODB.Recordset
'
'    strSQLEmpresa = "SELECT F1CODEMP, DESBASE " & _
'                    "FROM SF1PARCONTA " & _
'                    "WHERE ESTADO = TRUE " & _
'                    "ORDER BY F1CODEMP"
'
'    If rstEmpresa.State = 1 Then rstEmpresa.Close
'
'    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstEmpresa.EOF Then
'        comboList.Clear
'
'        Do While Not rstEmpresa.EOF
'            comboList.AddItem Trim(rstEmpresa!DESBASE & "") & Space(50) & Trim(rstEmpresa!CodEmpresa & "")
'
'            rstEmpresa.MoveNext
'        Loop
'    End If
'
'    comboList.ListIndex = -1
'
'    rstEmpresa.Close
'    strSQLEmpresa = vbNullString
'
'    Set rstEmpresa = Nothing
'
'    Exit Sub
'errListarEmpresasSoloSeleccion:
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: ListarEmpresasSoloSeleccion"
'
'    Err.Clear
'End Sub
'
'Public Sub listarSubEmpresasSoloSeleccion(ByVal combo As ComboBox)
'    On Error GoTo errListarSubEmpresasSoloSeleccion
'    Set rstEmpresa = New ADODB.Recordset
'
'    strSQLEmpresa = "SELECT CODSUBBASE, DESSUBBASE " & _
'                    "FROM SF1PARCONTA_SUB " & _
'                    "WHERE F1CODEMP = '" & strCodEmpresa & "' " & _
'                    "ORDER BY CODSUBBASE"
'
'    If rstEmpresa.State = 1 Then rstEmpresa.Close
'
'    rstEmpresa.Open strSQLEmpresa, cnn_control, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstEmpresa.EOF Then
'        combo.Clear
'
'        Do While Not rstEmpresa.EOF
'            combo.AddItem Trim(rstEmpresa!DESSUBBASE & "") & Space(50) & Trim(rstEmpresa!CodSubEmpresa & "")
'
'            rstEmpresa.MoveNext
'        Loop
'
'        combo.Enabled = True
'    Else
'        combo.Clear
'
'        combo.Enabled = False
'    End If
'
'    combo.ListIndex = -1
'
'    rstEmpresa.Close
'    strSQLEmpresa = vbNullString
'
'    Set rstEmpresa = Nothing
'
'    Exit Sub
'errListarSubEmpresasSoloSeleccion:
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: ListarSubEmpresasSoloSeleccion"
'
'    Err.Clear
'End Sub
'
'Public Function devuelveRstEmpresasEnUsoHC(Optional ByVal strAbrLinea As String, _
'                                        Optional ByVal bolOrdenarPorCodigo As Boolean) As ADODB.Recordset
'
'    On Error GoTo errDevuelveRstEmpresasEnUsoHC
'
'    Dim rsEmpresa As New ADODB.Recordset
'
'    strSQLEmpresa = vbNullString
'    strSQLEmpresa = strSQLEmpresa & "SELECT "
'    strSQLEmpresa = strSQLEmpresa & "EMP.F1CODEMP, "
'    strSQLEmpresa = strSQLEmpresa & "BAS.DESBASE "
'    strSQLEmpresa = strSQLEmpresa & "FROM "
'    strSQLEmpresa = strSQLEmpresa & "(EF2EMPLEADOS AS EMP "
'    strSQLEmpresa = strSQLEmpresa & "LEFT JOIN "
'    strSQLEmpresa = strSQLEmpresa & "SF1PARCONTA AS BAS ON BAS.F1CODEMP=EMP.F1CODEMP) "
'    strSQLEmpresa = strSQLEmpresa & "LEFT JOIN "
'    strSQLEmpresa = strSQLEmpresa & "EF2LINEA AS LIN ON LIN.CODLINEA=EMP.CODLINEA "
'    strSQLEmpresa = strSQLEmpresa & "WHERE "
'    strSQLEmpresa = strSQLEmpresa & "EMP.ESTADOEMPLEADO = 1 AND "
'    strSQLEmpresa = strSQLEmpresa & "BAS.ESTADO = TRUE "
'
'        If strAbrLinea <> vbNullString Then
'            strSQLEmpresa = strSQLEmpresa & "AND LIN.ABRLINEA = '" & strAbrLinea & "' "
'        End If
'
'    strSQLEmpresa = strSQLEmpresa & "GROUP BY "
'    strSQLEmpresa = strSQLEmpresa & "EMP.F1CODEMP, BAS.DESBASE "
'    strSQLEmpresa = strSQLEmpresa & "ORDER BY "
'
'        If bolOrdenarPorCodigo Then
'            strSQLEmpresa = strSQLEmpresa & "EMP.F1CODEMP"
'        Else
'            strSQLEmpresa = strSQLEmpresa & "BAS.DESBASE"
'        End If
'
'    If rsEmpresa.State = 1 Then rsEmpresa.Close
'
'    rsEmpresa.Open strSQLEmpresa, cnn_control, adOpenDynamic, adLockReadOnly
'
'    Set devuelveRstEmpresasEnUsoHC = rsEmpresa
'
'    strSQLEmpresa = vbNullString
'
'    Exit Function
'errDevuelveRstEmpresasEnUsoHC:
'    Set devuelveRstEmpresasEnUsoHC = Nothing
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsEmpresa: DevuelveRstEmpresasEnUsoHC"
'
'    Err.Clear
'End Function

