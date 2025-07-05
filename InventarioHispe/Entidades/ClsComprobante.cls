VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodSunat                 As String
Private strCodConcar                As String
Private strCodExterno               As String


Private strDescripcion              As String
Private strAbreviatura              As String


Private strTipoComprobante          As String
Private strDebHab                   As String
Private bolTransFerir               As Boolean
Private bolEsOficial                As Boolean


Private bolParaRegVentas            As Boolean
Private strRVDehHab                 As String
Private strRVEstados                As String
Private bolRVTieneSerie             As Boolean
Private intRVLongitudSerie          As Integer
Private strRVTipoCadenaSerie        As String
Private bolRVTieneNumero            As Boolean
Private intRVLongitudNumero         As Integer
Private strRVTipoCadenaNumero       As String
Private bolParaRegCompras           As Boolean
Private strRCDebHab                 As String
Private strRCEstados                As String
Private bolRCTieneSerie             As Boolean
Private intRCLongitudSerie          As Integer
Private strRCTipoCadenaSerie        As String
Private bolRCTieneNumero            As Boolean
Private intRCLongitudNumero         As Integer
Private strRCTipoCadenaNumero       As String

Private strCodExterno2              As String
Private strCodCompraRegistro        As String

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter           As String

Private rstComprobante              As ADODB.Recordset
Private strSQLComprobante           As String

'Propiedad Codigo de Comprobante
Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Codigo de SUNAT para Comprobante
Public Property Let CodigoSunat(ByVal Value As String)
    strCodSunat = Value
End Property

Public Property Get CodigoSunat() As String
    CodigoSunat = strCodSunat
End Property

'Propiedad Codigo de CONCAR para Comprobante
Public Property Let CodigoConcar(ByVal Value As String)
    strCodConcar = Value
End Property

Public Property Get CodigoConcar() As String
    CodigoConcar = strCodConcar
End Property

'Propiedad Codigo Externo para Comprobante
Public Property Let CodigoExterno(ByVal Value As String)
    strCodExterno = Value
End Property

Public Property Get CodigoExterno() As String
    CodigoExterno = strCodExterno
End Property




'Propiedad Descripcion de Comprobante
Public Property Let Descripcion(ByVal Value As String)
    strDescripcion = Value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

'Propiedad Abreviatura de Comprobante
Public Property Let Abreviatura(ByVal Value As String)
    strAbreviatura = Value
End Property

Public Property Get Abreviatura() As String
    Abreviatura = strAbreviatura
End Property




'Propiedad Tipo de Comprobante
Public Property Let TipoComprobante(ByVal Value As String)
    strTipoComprobante = Value
End Property

Public Property Get TipoComprobante() As String
    TipoComprobante = strTipoComprobante
End Property

'Propiedad Destino Contable de Comprobante
Public Property Let DebHab(ByVal Value As String)
    strDebHab = Value
End Property

Public Property Get DebHab() As String
    DebHab = strDebHab
End Property

'Propiedad Comprobante es Transferible a Contabilidad
Public Property Let TransFerir(ByVal Value As Boolean)
    bolTransFerir = Value
End Property

Public Property Get TransFerir() As Boolean
    TransFerir = bolTransFerir
End Property

'Propiedad Comprobante Es Oficial
Public Property Let EsOficial(ByVal Value As Boolean)
    bolEsOficial = Value
End Property

Public Property Get EsOficial() As Boolean
    EsOficial = bolEsOficial
End Property



'Private bolParaRegVentas            As Boolean
Public Property Let ParaRegVentas(ByVal Value As Boolean)
    bolParaRegVentas = Value
End Property

Public Property Get ParaRegVentas() As Boolean
    ParaRegVentas = bolParaRegVentas
End Property
'Private strRVDehHab                 As String
Public Property Let RVDehHab(ByVal Value As String)
    strRVDehHab = Value
End Property

Public Property Get RVDehHab() As String
    RVDehHab = strRVDehHab
End Property
'Private strRVEstados                As String
Public Property Let RVEstados(ByVal Value As String)
    strRVEstados = Value
End Property

Public Property Get RVEstados() As String
    RVEstados = strRVEstados
End Property
'Private bolRVTieneSerie             As Boolean
Public Property Let RVTieneSerie(ByVal Value As Boolean)
    bolRVTieneSerie = Value
End Property

Public Property Get RVTieneSerie() As Boolean
    RVTieneSerie = bolRVTieneSerie
End Property
'Private intRVLongitudSerie          As Integer
Public Property Let RVLongitudSerie(ByVal Value As Integer)
    intRVLongitudSerie = Value
End Property

Public Property Get RVLongitudSerie() As Integer
    RVLongitudSerie = intRVLongitudSerie
End Property
'Private strRVTipoCadenaSerie        As String
Public Property Let RVTipoCadenaSerie(ByVal Value As String)
    strRVTipoCadenaSerie = Value
End Property

Public Property Get RVTipoCadenaSerie() As String
    RVTipoCadenaSerie = strRVTipoCadenaSerie
End Property
'Private bolRVTieneNumero            As Boolean
Public Property Let RVTieneNumero(ByVal Value As Boolean)
    bolRVTieneNumero = Value
End Property

Public Property Get RVTieneNumero() As Boolean
    RVTieneNumero = bolRVTieneNumero
End Property
'Private intRVLongitudNumero         As Integer
Public Property Let RVLongitudNumero(ByVal Value As Integer)
    intRVLongitudNumero = Value
End Property

Public Property Get RVLongitudNumero() As Integer
    RVLongitudNumero = intRVLongitudNumero
End Property
'Private strRVTipoCadenaNumero       As String
Public Property Let RVTipoCadenaNumero(ByVal Value As String)
    strRVTipoCadenaNumero = Value
End Property

Public Property Get RVTipoCadenaNumero() As String
    RVTipoCadenaNumero = strRVTipoCadenaNumero
End Property
'Private bolParaRegCompras           As Boolean
Public Property Let ParaRegCompras(ByVal Value As Boolean)
    bolParaRegCompras = Value
End Property

Public Property Get ParaRegCompras() As Boolean
    ParaRegCompras = bolParaRegCompras
End Property
'Private strRCDebHab                 As String
Public Property Let RCDebHab(ByVal Value As String)
    strRCDebHab = Value
End Property

Public Property Get RCDebHab() As String
    RCDebHab = strRCDebHab
End Property
'Private strRCEstados                As String
Public Property Let RCEstados(ByVal Value As String)
    strRCEstados = Value
End Property

Public Property Get RCEstados() As String
    RCEstados = strRCEstados
End Property
'Private bolRCTieneSerie             As Boolean
Public Property Let RCTieneSerie(ByVal Value As Boolean)
    bolRCTieneSerie = Value
End Property

Public Property Get RCTieneSerie() As Boolean
    RCTieneSerie = bolRCTieneSerie
End Property
'Private intRCLongitudSerie          As Integer
Public Property Let RCLongitudSerie(ByVal Value As Integer)
    intRCLongitudSerie = Value
End Property

Public Property Get RCLongitudSerie() As Integer
    RCLongitudSerie = intRCLongitudSerie
End Property
'Private strRCTipoCadenaSerie        As String
Public Property Let RCTipoCadenaSerie(ByVal Value As String)
    strRCTipoCadenaSerie = Value
End Property

Public Property Get RCTipoCadenaSerie() As String
    RCTipoCadenaSerie = strRCTipoCadenaSerie
End Property
'Private bolRCTieneNumero            As Boolean
Public Property Let RCTieneNumero(ByVal Value As Boolean)
    bolRCTieneNumero = Value
End Property

Public Property Get RCTieneNumero() As Boolean
    RCTieneNumero = bolRCTieneNumero
End Property
'Private intRCLongitudNumero         As Integer
Public Property Let RCLongitudNumero(ByVal Value As Integer)
    intRCLongitudNumero = Value
End Property

Public Property Get RCLongitudNumero() As Integer
    RCLongitudNumero = intRCLongitudNumero
End Property
'Private strRCTipoCadenaNumero       As String
Public Property Let RCTipoCadenaNumero(ByVal Value As String)
    strRCTipoCadenaNumero = Value
End Property

Public Property Get RCTipoCadenaNumero() As String
    RCTipoCadenaNumero = strRCTipoCadenaNumero
End Property
'Private strCodExterno2              As String
Public Property Let CodigoExterno2(ByVal Value As String)
    strCodExterno2 = Value
End Property

Public Property Get CodigoExterno2() As String
    CodigoExterno2 = strCodExterno2
End Property


'Propiedad Codigo de Tipo de Registro de Compra
Public Property Let CodCompraRegistro(ByVal Value As String)
    strCodCompraRegistro = Value
End Property

Public Property Get CodCompraRegistro() As String
    CodCompraRegistro = strCodCompraRegistro
End Property




'Propiedad Fecha de Registro de Comprobante
Public Property Let FechaReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Comprobante
Public Property Let UsuarioReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Comprobante
Public Property Let FechaMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Comprobante
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
    strCodigo = vbNullString
    strCodSunat = vbNullString
    strCodConcar = vbNullString
    strCodExterno = vbNullString
    
    
    strDescripcion = vbNullString
    strAbreviatura = vbNullString
    
    
    strTipoComprobante = vbNullString
    strDebHab = vbNullString
    bolTransFerir = False
    bolEsOficial = False
    
    
    bolParaRegVentas = False
    strRVDehHab = vbNullString
    strRVEstados = vbNullString
    bolRVTieneSerie = False
    intRVLongitudSerie = 0
    strRVTipoCadenaSerie = vbNullString
    bolRVTieneNumero = False
    intRVLongitudNumero = 0
    strRVTipoCadenaNumero = vbNullString
    bolParaRegCompras = False
    strRCDebHab = vbNullString
    strRCEstados = vbNullString
    bolRCTieneSerie = False
    intRCLongitudSerie = 0
    strRCTipoCadenaSerie = vbNullString
    bolRCTieneNumero = False
    intRCLongitudNumero = 0
    strRCTipoCadenaNumero = vbNullString
    
    
    strCodExterno2 = vbNullString
    strCodCompraRegistro = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    
    strSQLSelectAlter = vbNullString
End Sub


Public Function obtenerComprobante() As Boolean
    On Error GoTo errObtenerComprobante
    
    Set rstComprobante = New ADODB.Recordset
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "* "
    strSQLComprobante = strSQLComprobante & "FROM "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS "
    strSQLComprobante = strSQLComprobante & "WHERE "
    strSQLComprobante = strSQLComprobante & "F2CODDOC = '" & strCodigo & "'"
            
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    rstComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstComprobante.EOF Then
        strCodigo = Trim(rstComprobante!F2CODDOC & "")
        strCodSunat = Trim(rstComprobante!CODEXT2 & "")
        strCodConcar = Trim(rstComprobante!F2ABREV_CONCAR & "")
        strCodExterno = Trim(rstComprobante!CODEXT3 & "")
        
        
        strDescripcion = Trim(rstComprobante!F2DESDOC & "")
        strAbreviatura = Trim(rstComprobante!F2ABREV & "")
        
        
        strTipoComprobante = Trim(rstComprobante!F2TIPO & "")
        strDebHab = Trim(rstComprobante!F2DEBHAB & "")
        bolTransFerir = IIf(Trim(rstComprobante!F2TRANSFER & "") = "S", True, False)
        bolEsOficial = CBool(rstComprobante!F2OFICIAL)
        
        
        bolParaRegVentas = CBool(rstComprobante!ParaRegVentas)
        strRVDehHab = Trim(rstComprobante!RVDehHab & "")
        strRVEstados = Trim(rstComprobante!RVEstados & "")
        bolRVTieneSerie = CBool(rstComprobante!RVTieneSerie)
        intRVLongitudSerie = Val(rstComprobante!RVLongitudSerie & "")
        strRVTipoCadenaSerie = Trim(rstComprobante!RVTipoCadenaSerie & "")
        bolRVTieneNumero = CBool(rstComprobante!RVTieneNumero)
        intRVLongitudNumero = Val(rstComprobante!RVLongitudNumero & "")
        strRVTipoCadenaNumero = Trim(rstComprobante!RVTipoCadenaNumero & "")
        bolParaRegCompras = CBool(rstComprobante!ParaRegCompras)
        strRCDebHab = Trim(rstComprobante!RCDebHab & "")
        strRCEstados = Trim(rstComprobante!RCEstados & "")
        bolRCTieneSerie = CBool(rstComprobante!RCTieneSerie)
        intRCLongitudSerie = Val(rstComprobante!RCLongitudSerie & "")
        strRCTipoCadenaSerie = Trim(rstComprobante!RCTipoCadenaSerie & "")
        bolRCTieneNumero = CBool(rstComprobante!RCTieneNumero)
        intRCLongitudNumero = Val(rstComprobante!RCLongitudNumero & "")
        strRCTipoCadenaNumero = Trim(rstComprobante!RCTipoCadenaNumero & "")
        
        
        strCodExterno2 = Trim(rstComprobante!CODEXT4 & "")
        strCodCompraRegistro = Trim(rstComprobante!F2CODCOMPRAREG & "")
        
        obtenerComprobante = True
    Else
        obtenerComprobante = False
    End If
    
    rstComprobante.Close
    
    Set rstComprobante = Nothing
    
    Exit Function
errObtenerComprobante:
    obtenerComprobante = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: ObtenerComprobante"
    
    Err.Clear
End Function

Public Sub obtenerConfigComprobante()
    On Error GoTo errObtenerConfigComprobante
    Set rstComprobante = New ADODB.Recordset
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "* "
    strSQLComprobante = strSQLComprobante & "FROM "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS "
    strSQLComprobante = strSQLComprobante & "WHERE "
    strSQLComprobante = strSQLComprobante & "F2CODDOC = '" & strCodigo & "'"
    
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    rstComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstComprobante.EOF Then
        strCodigo = Trim(rstComprobante!F2CODDOC & "")
        strCodSunat = Trim(rstComprobante!CODEXT2 & "")
        
        strCodigo = Trim(rstComprobante!F2CODDOC & "")
        strCodSunat = Trim(rstComprobante!CODEXT2 & "")
        strCodConcar = Trim(rstComprobante!F2ABREV_CONCAR & "")
        strCodExterno = Trim(rstComprobante!CODEXT3 & "")
        
        
        strDescripcion = Trim(rstComprobante!F2DESDOC & "")
        strAbreviatura = Trim(rstComprobante!F2ABREV & "")
        
        
        strTipoComprobante = Trim(rstComprobante!F2TIPO & "")
        strDebHab = Trim(rstComprobante!F2DEBHAB & "")
        bolTransFerir = IIf(Trim(rstComprobante!F2TRANSFER & "") = "S", True, False)
        bolEsOficial = CBool(rstComprobante!F2OFICIAL)
        
        
        bolParaRegVentas = CBool(rstComprobante!ParaRegVentas)
        strRVDehHab = Trim(rstComprobante!RVDehHab & "")
        strRVEstados = Trim(rstComprobante!RVEstados & "")
        bolRVTieneSerie = CBool(rstComprobante!RVTieneSerie)
        intRVLongitudSerie = Val(rstComprobante!RVLongitudSerie & "")
        strRVTipoCadenaSerie = Trim(rstComprobante!RVTipoCadenaSerie & "")
        bolRVTieneNumero = CBool(rstComprobante!RVTieneNumero)
        intRVLongitudNumero = Val(rstComprobante!RVLongitudNumero & "")
        strRVTipoCadenaNumero = Trim(rstComprobante!RVTipoCadenaNumero & "")
        bolParaRegCompras = CBool(rstComprobante!ParaRegCompras)
        strRCDebHab = Trim(rstComprobante!RCDebHab & "")
        strRCEstados = Trim(rstComprobante!RCEstados & "")
        bolRCTieneSerie = CBool(rstComprobante!RCTieneSerie)
        intRCLongitudSerie = Val(rstComprobante!RCLongitudSerie & "")
        strRCTipoCadenaSerie = Trim(rstComprobante!RCTipoCadenaSerie & "")
        bolRCTieneNumero = CBool(rstComprobante!RCTieneNumero)
        intRCLongitudNumero = Val(rstComprobante!RCLongitudNumero & "")
        strRCTipoCadenaNumero = Trim(rstComprobante!RCTipoCadenaNumero & "")
        
        
        strCodExterno2 = Trim(rstComprobante!CODEXT4 & "")
        strCodCompraRegistro = Trim(rstComprobante!F2CODCOMPRAREG & "")
    End If
    
    rstComprobante.Close
    
    Set rstComprobante = Nothing
    
    Exit Sub
errObtenerConfigComprobante:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: ObtenerConfigComprobante"
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstComprobante = New ADODB.Recordset
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS.* "
    strSQLComprobante = strSQLComprobante & "FROM "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS "
    strSQLComprobante = strSQLComprobante & "WHERE "
    strSQLComprobante = strSQLComprobante & "F2CODDOC = '" & strCodigo & "'"
    
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    rstComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstComprobante.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstComprobante.Close
    
    Set rstComprobante = Nothing
    
    Exit Function
errVerificarExistencia:
    verificarExistencia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: VerificarExistencia"
    
    Err.Clear
End Function

Public Function generarCodigoComprobante() As String
    On Error GoTo errGenerarCodigoComprobante
    
    Dim rsComprobante As New ADODB.Recordset
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "F2CODDOC "
    strSQLComprobante = strSQLComprobante & "FROM "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS "
    strSQLComprobante = strSQLComprobante & "ORDER BY "
    strSQLComprobante = strSQLComprobante & "F2CODDOC"
    
    If rsComprobante.State = 1 Then rsComprobante.Close
    
    rsComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsComprobante.EOF Then
        rsComprobante.MoveFirst
        
        Do While Not rsComprobante.EOF
            strCodigo = Format(Val(Trim(rsComprobante!F2CODDOC & "")) + 1, "00")
            
            If Not Me.verificarExistencia Then
                Exit Do
            End If
            
            rsComprobante.MoveNext
        Loop
        
        generarCodigoComprobante = strCodigo
    Else
        generarCodigoComprobante = "01"
    End If
    
    rsComprobante.Close
    
    Set rsComprobante = Nothing
    
    Exit Function
errGenerarCodigoComprobante:
    generarCodigoComprobante = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: GenerarCodigoComprobante"
    
    Err.Clear
End Function

Public Function guardarComprobante() As Boolean
    On Error GoTo errGuardarComprobante
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoComprobante
        End If
        
        strSQLComprobante = vbNullString
        strSQLComprobante = strSQLComprobante & "INSERT INTO DOCUMENTOS("
        strSQLComprobante = strSQLComprobante & "F2CODDOC, CODEXT2, CODEXT3, "
        strSQLComprobante = strSQLComprobante & "F2DESDOC, F2ABREV, F2ABREV_CONCAR, "
        strSQLComprobante = strSQLComprobante & "F2TIPO, F2DEBHAB, F2TRANSFER, F2OFICIAL, "
        strSQLComprobante = strSQLComprobante & "PARAREGVENTAS, RVDEHHAB, RVESTADOS, "
        strSQLComprobante = strSQLComprobante & "RVTIENESERIE,  RVLONGITUDSERIE, "
        strSQLComprobante = strSQLComprobante & "RVTIPOCADENASERIE, RVTIENENUMERO, "
        strSQLComprobante = strSQLComprobante & "RVLONGITUDNUMERO, RVTIPOCADENANUMERO, "
        strSQLComprobante = strSQLComprobante & "PARAREGCOMPRAS, RCDEBHAB, RCESTADOS, "
        strSQLComprobante = strSQLComprobante & "RCTIENESERIE, RCLONGITUDSERIE, RCTIPOCADENASERIE, "
        strSQLComprobante = strSQLComprobante & "RCTIENENUMERO, RCLONGITUDNUMERO, "
        strSQLComprobante = strSQLComprobante & "RCTIPOCADENANUMERO, CODEXT4, "
        strSQLComprobante = strSQLComprobante & "F2CODCOMPRAREG"
        strSQLComprobante = strSQLComprobante & ") "
        strSQLComprobante = strSQLComprobante & "VALUES("
        strSQLComprobante = strSQLComprobante & "'" & strCodigo & "', "
        strSQLComprobante = strSQLComprobante & IIf(strCodSunat <> vbNullString, "'" & strCodSunat & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & IIf(strCodExterno <> vbNullString, "'" & strCodExterno & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & "'" & strDescripcion & "', "
        strSQLComprobante = strSQLComprobante & "'" & strAbreviatura & "', "
        strSQLComprobante = strSQLComprobante & IIf(strCodConcar <> vbNullString, "'" & strCodConcar & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & "'" & strTipoComprobante & "', "
        strSQLComprobante = strSQLComprobante & "'" & strDebHab & "', "
        strSQLComprobante = strSQLComprobante & "'" & IIf(bolTransFerir, "S", "N") & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolEsOficial, "TRUE", "FALSE") & "', "
        
        strSQLComprobante = strSQLComprobante & IIf(bolParaRegVentas, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRVDehHab & "', "
        strSQLComprobante = strSQLComprobante & "'" & strRVEstados & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolRVTieneSerie, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & intRVLongitudSerie & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRVTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolRVTieneNumero, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & intRVLongitudNumero & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRVTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolParaRegCompras, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRCDebHab & "', "
        strSQLComprobante = strSQLComprobante & "'" & strRCEstados & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolRCTieneSerie, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & intRCLongitudSerie & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRCTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & IIf(bolRCTieneNumero, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & intRCLongitudNumero & ", "
        strSQLComprobante = strSQLComprobante & "'" & strRCTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & "'" & strCodExterno2 & "', "
        
        strSQLComprobante = strSQLComprobante & "'" & strCodCompraRegistro & "'"
        strSQLComprobante = strSQLComprobante & ")"
    Else
        strSQLComprobante = strSQLComprobante & "UPDATE DOCUMENTOS "
        strSQLComprobante = strSQLComprobante & "SET "
        strSQLComprobante = strSQLComprobante & "CODEXT2 = " & IIf(strCodSunat <> vbNullString, "'" & strCodSunat & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & "CODEXT3 = " & IIf(strCodExterno <> vbNullString, "'" & strCodExterno & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & "F2DESDOC = '" & strDescripcion & "', "
        strSQLComprobante = strSQLComprobante & "F2ABREV = '" & strAbreviatura & "', "
        strSQLComprobante = strSQLComprobante & "F2ABREV_CONCAR = " & IIf(strCodConcar <> vbNullString, "'" & strCodConcar & "'", "NULL") & ", "
        strSQLComprobante = strSQLComprobante & "F2TIPO = '" & strTipoComprobante & "', "
        strSQLComprobante = strSQLComprobante & "F2DEBHAB = '" & strDebHab & "', "
        strSQLComprobante = strSQLComprobante & "F2TRANSFER = '" & IIf(bolTransFerir, "S", "N") & "', "
        strSQLComprobante = strSQLComprobante & "F2OFICIAL = " & IIf(bolEsOficial, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "PARAREGVENTAS = " & IIf(bolParaRegVentas, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RVDEHHAB = '" & strRVDehHab & "', "
        strSQLComprobante = strSQLComprobante & "RVESTADOS = '" & strRVEstados & "', "
        strSQLComprobante = strSQLComprobante & "RVTIENESERIE = " & IIf(bolRVTieneSerie, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RVLONGITUDSERIE = " & intRVLongitudSerie & ", "
        strSQLComprobante = strSQLComprobante & "RVTIPOCADENASERIE = '" & strRVTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & "RVTIENENUMERO = " & IIf(bolRVTieneNumero, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RVLONGITUDNUMERO = " & intRVLongitudNumero & ", "
        strSQLComprobante = strSQLComprobante & "RVTIPOCADENANUMERO = '" & strRVTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & "PARAREGCOMPRAS = " & IIf(bolParaRegCompras, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RCDEBHAB = '" & strRCDebHab & "', "
        strSQLComprobante = strSQLComprobante & "RCESTADOS = '" & strRCEstados & "', "
        strSQLComprobante = strSQLComprobante & "RCTIENESERIE = " & IIf(bolRCTieneSerie, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RCLONGITUDSERIE = " & intRCLongitudSerie & ", "
        strSQLComprobante = strSQLComprobante & "RCTIPOCADENASERIE = '" & strRCTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & "RCTIENENUMERO = " & IIf(bolRCTieneNumero, "TRUE", "FALSE") & ", "
        strSQLComprobante = strSQLComprobante & "RCLONGITUDNUMERO = " & intRCLongitudNumero & ", "
        strSQLComprobante = strSQLComprobante & "RCTIPOCADENANUMERO = '" & strRCTipoCadenaSerie & "', "
        strSQLComprobante = strSQLComprobante & "CODEXT4 = '" & strCodExterno2 & "', "
        strSQLComprobante = strSQLComprobante & "F2CODCOMPRAREG = '" & strCodCompraRegistro & "' "
        strSQLComprobante = strSQLComprobante & "WHERE "
        strSQLComprobante = strSQLComprobante & "F2CODDOC = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLComprobante
    
    guardarComprobante = True
    
    Exit Function
errGuardarComprobante:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: GuardarComprobante"
    
    guardarComprobante = False
    
    Err.Clear
End Function

Public Function eliminarComprobante() As Boolean
    On Error GoTo errEliminarComprobante
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "DELETE FROM "
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS "
    strSQLComprobante = strSQLComprobante & "WHERE "
    strSQLComprobante = strSQLComprobante & "F2CODDOC = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLComprobante
    
    eliminarComprobante = True
    
    Exit Function
errEliminarComprobante:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: EliminarComprobante"
    
    eliminarComprobante = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Comprobantes en Grilla (QuamtumGrid)
Public Sub vistaComprobante(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaComprobante
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "D.F2CODDOC, "
    strSQLComprobante = strSQLComprobante & "D.CODEXT2, "
    strSQLComprobante = strSQLComprobante & "UCASE(D.F2DESDOC) AS F2DESDOC, "
    strSQLComprobante = strSQLComprobante & "D.F2ABREV, "
    strSQLComprobante = strSQLComprobante & "D.F2TIPO, "
    strSQLComprobante = strSQLComprobante & "IIF(TRIM(D.F2TRANSFER & '') = 'S', 'S', 'N') AS TRANSFER, "
    strSQLComprobante = strSQLComprobante & "D.F2OFICIAL, "
    strSQLComprobante = strSQLComprobante & "CR.DESCRIPCION AS TIPOREGISTRO "
    
    strSQLComprobante = strSQLComprobante & "FROM "
    
    strSQLComprobante = strSQLComprobante & "DOCUMENTOS AS D "
    strSQLComprobante = strSQLComprobante & "LEFT JOIN EF2COMPRAREGISTRO AS CR ON CR.CODIGO = D.F2CODCOMPRAREG "
    
    strSQLComprobante = strSQLComprobante & "ORDER BY "
    strSQLComprobante = strSQLComprobante & "D.F2CODDOC"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Comprobante
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Código"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODDOC"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Codigo de SUNAT
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Cod. Sunat"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODEXT2"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodSunat"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DESDOC"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
                .Visible = True
            End With
            
            'Columna Abreviatura
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Abrev."
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2ABREV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAbreviatura"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Tipo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2TIPO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCentro"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Transferible a Contabilidad
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Transf."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "TRANSFER"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTransfer"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Comprobante Oficial
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Oficial"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F2OFICIAL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColOficial"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Transferible a Contabilidad
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Tipo R.C."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "TIPOREGISTRO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipoRegistro"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexDbBancos
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLComprobante
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2CODDOC"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLComprobante = vbNullString
    
    Exit Sub
errVistaComprobante:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsComprobante: VistaComprobante"
    End Select
    
    Err.Clear
End Sub

'Vista de Detalle de Comprobantes Presupuestales en Grilla (QuamtumGrid)
Public Sub vistaComprobanteDetalle(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaComprobanteDetalle
    
    strSQLComprobante = vbNullString
    strSQLComprobante = strSQLComprobante & "SELECT "
    strSQLComprobante = strSQLComprobante & "CD.CODPARTIDA, "
    strSQLComprobante = strSQLComprobante & "CD.CODINSUMO, "
    strSQLComprobante = strSQLComprobante & "I.F5NOMPRO, "
    strSQLComprobante = strSQLComprobante & "M.F7NOMMED, "
    strSQLComprobante = strSQLComprobante & "VAL(CD.CANTIDAD & '') AS CANT "
    
    strSQLComprobante = strSQLComprobante & "FROM "
    
    strSQLComprobante = strSQLComprobante & "(DOCUMENTOSDETALLE AS CD "
    strSQLComprobante = strSQLComprobante & "LEFT JOIN IF5PLA AS I ON I.F5CODPRO = CD.CODINSUMO) "
    strSQLComprobante = strSQLComprobante & "LEFT JOIN EF7MEDIDAS AS M ON M.F7CODMED = I.F7CODMED "
    
    strSQLComprobante = strSQLComprobante & "WHERE "
    strSQLComprobante = strSQLComprobante & "CD.F2CODDOC = '" & strCodigo & "'"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Partida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Partida"
                .Color = &H80000000
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODPARTIDA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodPartida"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Codigo de Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &H80000000
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODINSUMO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodInsumo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Descripcion de Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción de Recurso"
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 160
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Unidad Med."
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "F7NOMMED"
                .HeaderAlignment = taCenter
                .ObjectName = "ColUnidadMedida"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANT"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexDbBancos
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLComprobante
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPARTIDA"
        End With
    End If
    
    strSQLComprobante = vbNullString
    
    Exit Sub
errVistaComprobanteDetalle:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsComprobante: VistaComprobanteDetalle"
    End Select
    
    Err.Clear
End Sub

Public Sub listarAnnoComprobante(ByVal combo As Object)
    On Error GoTo errListarAnnoComprobante
    
    Set rstComprobante = New ADODB.Recordset
    
    strSQLComprobante = "SELECT YEAR(FECHA) AS ANNO " & _
                    "FROM DOCUMENTOS " & _
                    "GROUP BY YEAR(FECHA)"
    
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    rstComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    combo.Clear
    combo.AddItem "(*) - Todos"
    
    If Not rstComprobante.EOF Then
        
        rstComprobante.MoveFirst
        
        Do While Not rstComprobante.EOF
            combo.AddItem Trim(rstComprobante!Anno & "")
            
            rstComprobante.MoveNext
        Loop
    End If
    
    rstComprobante.Close
    
    strSQLComprobante = vbNullString
    
    Set rstComprobante = Nothing
    
    Exit Sub
errListarAnnoComprobante:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: ListarAnnoComprobante"
    
    Err.Clear
End Sub

Public Sub listarTipoComprobante(ByVal comboList As Object, ByVal strIncluirTipoComprobanteCorPorAorAll As String)
    On Error GoTo errListarTipoComprobante
    
    Set rstComprobante = New ADODB.Recordset
    
    strSQLComprobante = "SELECT F2CODDOC, F2DESDOC " & _
                        "FROM DOCUMENTOS " & _
                        "WHERE F2TIPO IN (" & strIncluirTipoComprobanteCorPorAorAll & ") " & _
                        "ORDER BY F2DESDOC"
    
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    rstComprobante.Open strSQLComprobante, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstComprobante.EOF Then
        
        rstComprobante.MoveFirst
        
        Do While Not rstComprobante.EOF
            comboList.AddItem Trim(rstComprobante!F2DESDOC & "") & Space(100) & Trim(rstComprobante!F2CODDOC & "")
            
            rstComprobante.MoveNext
        Loop
    End If
    
    If rstComprobante.State = 1 Then rstComprobante.Close
    
    strSQLComprobante = vbNullString
    
    Set rstComprobante = Nothing
    
    Exit Sub
errListarTipoComprobante:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsComprobante: ListarTipoComprobante"
    
    Err.Clear
End Sub
