VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strNumeroOrden               As String 'NUMORDEN

Private strNumeroPlan               As String 'NUMPLANEAMIENTO
Private StrFecha                    As String 'FECHA
Private strFechaEntrega             As String 'FECHA_ENTREGA
Private strObservacion              As String 'OBSERVACION
Private strEstacion                 As String 'ESTACION
Private bolAprobadoLogistica        As Boolean 'APROBADO_LOGISTICA
Private bolAprobadoCredito          As Boolean 'APROBADO_CREDITOS
Private bolAprobadoDiseno           As Boolean 'APROBADO_DISEÑO
Private strCodMoneda                As String 'MONEDA
Private dblTotal                    As Double 'TOTAL
Private strFacturado                As String 'FACTURADO


'Atributos de Insumos de O/T
Private strNumeroFormula As String 'F4NUMFORMULA
Private strCodProductoComp As String 'F3CODPROCOM
Private strNombreComp As String 'NOMCOMPONENTE
Private dblItem As Double 'F3ITEM
Private strCodProductoIns As String 'F3CODPROINS
Private strDesProductoIns As String 'DESCRIPCION
Private strCodUM As String 'UMEDIDA
Private dblCantidad As Double 'CANTIDAD
Private dblCantidadUnitario As Double 'F3CANTIDADUNITARIO
Private dblCantidadOrdenFormula As Double 'F3CANTIDADORDENFORMULA
Private dblCantidadRequeridaStock As Double 'CANTIDADREQUERIDASTOCK
Private dblSecProceso 'SECPROCESO
Private strNumeroPedido As String 'CARGADO


'Atributos de Detalle de O/T
Private strCodProducto As String 'F3CODPRO
Private strDesProducto As String 'DESCRIPCION
Private dblPrecio As Double 'PRECIO
Private strCodProceso As String  'CODPROCESO


Private strSQLSelectAlter           As String

Private rstOT                     As ADODB.Recordset
Private strSQLOT                  As String


Public Property Let NumeroOrden(ByVal value As String)
    strNumeroOrden = value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property

Public Property Let NumeroPlan(ByVal value As String)
    strNumeroPlan = value
End Property

Public Property Get NumeroPlan() As String
    NumeroPlan = strNumeroPlan
End Property

Public Property Let Fecha(ByVal value As String)
    StrFecha = value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property

Public Property Let FechaEntrega(ByVal value As String)
    strFechaEntrega = value
End Property

Public Property Get FechaEntrega() As String
    FechaEntrega = strFechaEntrega
End Property

Public Property Let Observacion(ByVal value As String)
    strObservacion = value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property

Public Property Let Estacion(ByVal value As String)
    strEstacion = value
End Property

Public Property Get Estacion() As String
    Estacion = strEstacion
End Property

Public Property Let AprobadoLogistica(ByVal value As Boolean)
    bolAprobadoLogistica = value
End Property

Public Property Get AprobadoLogistica() As Boolean
    AprobadoLogistica = bolAprobadoLogistica
End Property

Public Property Let AprobadoCredito(ByVal value As Boolean)
    bolAprobadoCredito = value
End Property

Public Property Get AprobadoCredito() As Boolean
    AprobadoCredito = bolAprobadoCredito
End Property

Public Property Let AprobadoDiseno(ByVal value As Boolean)
    bolAprobadoDiseno = value
End Property

Public Property Get AprobadoDiseno() As Boolean
    AprobadoDiseno = bolAprobadoDiseno
End Property

Public Property Let CodigoMoneda(ByVal value As String)
    strCodMoneda = value
End Property

Public Property Get CodigoMoneda() As String
    CodigoMoneda = strCodMoneda
End Property

Public Property Let TOTAL(ByVal value As Double)
    dblTotal = value
End Property

Public Property Get TOTAL() As Double
    TOTAL = dblTotal
End Property

Public Property Let Facturado(ByVal value As String)
    strFacturado = value
End Property

Public Property Get Facturado() As String
    Facturado = strFacturado
End Property






'Atributos de Insumos de O/T
Public Property Let NumeroFormula(ByVal value As String)
    strNumeroFormula = value
End Property

Public Property Get NumeroFormula() As String
    NumeroFormula = strNumeroFormula
End Property

Public Property Let CodigoProductoComponente(ByVal value As String)
    strCodProductoComp = value
End Property

Public Property Get CodigoProductoComponente() As String
    CodigoProductoComponente = strCodProductoComp
End Property

Public Property Let NombreComponente(ByVal value As String)
    strNombreComp = value
End Property

Public Property Get NombreComponente() As String
    NombreComponente = strNombreComp
End Property

Public Property Let ITEM(ByVal value As Double)
    dblItem = value
End Property

Public Property Get ITEM() As Double
    ITEM = dblItem
End Property

Public Property Let CodigoProductoInsumo(ByVal value As String)
    strCodProductoIns = value
End Property

Public Property Get CodigoProductoInsumo() As String
    CodigoProductoInsumo = strCodProductoIns
End Property

Public Property Let DescripcionProductoInsumo(ByVal value As String)
    strDesProductoIns = value
End Property

Public Property Get DescripcionProductoInsumo() As String
    DescripcionProductoInsumo = strDesProductoIns
End Property

Public Property Let CodigoUM(ByVal value As String)
    strCodUM = value
End Property

Public Property Get CodigoUM() As String
    CodigoUM = strCodUM
End Property

Public Property Let Cantidad(ByVal value As Double)
    dblCantidad = value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

Public Property Let CantidadUnitario(ByVal value As Double)
    dblCantidadUnitario = value
End Property

Public Property Get CantidadUnitario() As Double
    CantidadUnitario = dblCantidadUnitario
End Property

Public Property Let CantidadOrdenFormula(ByVal value As Double)
    dblCantidadOrdenFormula = value
End Property

Public Property Get CantidadOrdenFormula() As Double
    CantidadOrdenFormula = dblCantidadOrdenFormula
End Property

Public Property Let CantidadRequeridaStock(ByVal value As Double)
    dblCantidadRequeridaStock = value
End Property

Public Property Get CantidadRequeridaStock() As Double
    CantidadRequeridaStock = dblCantidadRequeridaStock
End Property

Public Property Let SecProceso(ByVal value As Double)
    dblSecProceso = value
End Property

Public Property Get SecProceso() As Double
    SecProceso = dblSecProceso
End Property

Public Property Let NumeroPedido(ByVal value As String)
    strNumeroPedido = value
End Property

Public Property Get NumeroPedido() As String
    NumeroPedido = strNumeroPedido
End Property






'Atributos de Detalle de O/T
Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Property Let DescripcionProducto(ByVal value As String)
    strDesProducto = value
End Property

Public Property Get DescripcionProducto() As String
    DescripcionProducto = strDesProducto
End Property

Public Property Let Precio(ByVal value As Double)
    dblPrecio = value
End Property

Public Property Get Precio() As Double
    Precio = dblPrecio
End Property

Public Property Let CodigoProceso(ByVal value As String)
    strCodProceso = value
End Property

Public Property Get CodigoProceso() As String
    CodigoProceso = strCodProceso
End Property





Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strNumeroOrden = vbNullString 'NUMORDEN
    strNumeroPlan = vbNullString 'NUMPLANEAMIENTO
    
    StrFecha = vbNullString 'FECHA
    strFechaEntrega = vbNullString 'FECHA_ENTREGA
    strObservacion = vbNullString 'OBSERVACION
    StrFecha = vbNullString    'ESTACION
    bolAprobadoLogistica = False 'APROBADO_LOGISTICA
    bolAprobadoCredito = False 'APROBADO_CREDITOS
    bolAprobadoDiseno = False 'APROBADO_DISEÑO
    strCodMoneda = vbNullString 'MONEDA
    dblTotal = 0 'TOTAL
    strFacturado = vbNullString   'FACTURADO
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesInsumos()
    strNumeroFormula = vbNullString 'F4NUMFORMULA
    strCodProductoComp = vbNullString 'F3CODPROCOM
    strNombreComp = vbNullString 'NOMCOMPONENTE
    dblItem = 0 'CANTIDADREQUERIDASTOCK
    strCodProductoIns = vbNullString 'F3CODPROINS
    strDesProductoIns = vbNullString 'DESCRIPCION
    
    strCodUM = vbNullString 'UMEDIDA
    dblCantidad = 0 'CANTIDAD
    dblCantidadUnitario = 0 'F3CANTIDADUNITARIO
    
    dblCantidadOrdenFormula = 0 'F3CANTIDADORDENFORMULA
    dblCantidadRequeridaStock = 0 'CANTIDADREQUERIDASTOCK
    dblSecProceso = 0  'SECPROCESO
    strNumeroPedido = vbNullString 'CARGADO
End Sub

Public Sub inicializarEntidadesDetalle()
    strCodProducto = vbNullString 'F3CODPRO
    strDesProducto = vbNullString 'DESCRIPCION
    dblPrecio = 0 'PRECIO
    strCodProceso = vbNullString 'CODPROCESO
End Sub

Public Function obtenerOrdenTrabajo() As Boolean
    On Error GoTo errObtenerOrdenTrabajo
    
    Set rstOT = New ADODB.Recordset
    
    strSQLOT = vbNullString
    strSQLOT = strSQLOT & "SELECT "
    strSQLOT = strSQLOT & "* "
    strSQLOT = strSQLOT & "FROM "
    strSQLOT = strSQLOT & "ORDENTRAB_CAB "
    strSQLOT = strSQLOT & "WHERE "
    strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "'"
    
    If rstOT.State = 1 Then rstOT.Close
    
    rstOT.Open strSQLOT, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOT.EOF Then
        strNumeroOrden = Trim(rstOT!NUMORDEN & "") 'NUMORDEN
        
        strNumeroPlan = Trim(rstOT!NUMPLANEAMIENTO & "") 'NUMPLANEAMIENTO
        StrFecha = Trim(rstOT!Fecha & "") 'FECHA
        strFechaEntrega = Trim(rstOT!FECHA_ENTREGA & "") 'FECHA_ENTREGA
        strObservacion = Trim(rstOT!Observacion & "") 'OBSERVACION
        strEstacion = Trim(rstOT!Estacion & "")     'ESTACION
        bolAprobadoLogistica = CBool(rstOT!APROBADO_LOGISTICA) 'APROBADO_LOGISTICA
        bolAprobadoCredito = CBool(rstOT!APROBADO_CREDITOS) 'APROBADO_CREDITOS
        bolAprobadoDiseno = CBool(rstOT!APROBADO_DISEÑO) 'APROBADO_DISEÑO
        strCodMoneda = Trim(rstOT!Moneda & "") 'MONEDA
        dblTotal = Val(rstOT!TOTAL & "") 'TOTAL
        strFacturado = Trim(rstOT!Facturado & "")   'FACTURADO
        
        obtenerOrdenTrabajo = True
    Else
        obtenerOrdenTrabajo = False
    End If
    
    rstOT.Close
    
    Set rstOT = Nothing
    
    Exit Function
errObtenerOrdenTrabajo:
    obtenerOrdenTrabajo = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: ObtenerOrdenTrabajo"
    
    Err.Clear
End Function

Public Sub obtenerConfigOrdenTrabajo()
    On Error GoTo errObtenerConfigOrdenTrabajo
    
    Set rstOT = New ADODB.Recordset
    
    strSQLOT = vbNullString
    strSQLOT = strSQLOT & "SELECT "
    strSQLOT = strSQLOT & "* "
    strSQLOT = strSQLOT & "FROM "
    strSQLOT = strSQLOT & "ORDENTRAB_CAB "
    strSQLOT = strSQLOT & "WHERE "
    strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "'"
    
    If rstOT.State = 1 Then rstOT.Close
    
    rstOT.Open strSQLOT, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOT.EOF Then
        strNumeroOrden = Trim(rstOT!NUMORDEN & "") 'NUMORDEN
        
        strNumeroPlan = Trim(rstOT!NUMPLANEAMIENTO & "") 'NUMPLANEAMIENTO
        StrFecha = Trim(rstOT!Fecha & "") 'FECHA
        strFechaEntrega = Trim(rstOT!FECHA_ENTREGA & "") 'FECHA_ENTREGA
        strObservacion = Trim(rstOT!Observacion & "") 'OBSERVACION
        strEstacion = Trim(rstOT!Estacion & "")     'ESTACION
        bolAprobadoLogistica = CBool(rstOT!APROBADO_LOGISTICA) 'APROBADO_LOGISTICA
        bolAprobadoCredito = CBool(rstOT!APROBADO_CREDITOS) 'APROBADO_CREDITOS
        bolAprobadoDiseno = CBool(rstOT!APROBADO_DISEÑO) 'APROBADO_DISEÑO
        strCodMoneda = Trim(rstOT!Moneda & "") 'MONEDA
        dblTotal = Val(rstOT!TOTAL & "") 'TOTAL
        strFacturado = Trim(rstOT!Facturado & "")   'FACTURADO
    End If
    
    rstOT.Close
    
    Set rstOT = Nothing
    
    Exit Sub
errObtenerConfigOrdenTrabajo:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: ObtenerConfigOrdenTrabajo"
    
    Err.Clear
End Sub

'Public Function generarNumeroVale() As String
'    On Error GoTo errGenerarNumeroVale
'
'    Set rstOT = New ADODB.Recordset
'
'    strSQLOT = vbNullString
'    strSQLOT = strSQLOT & "SELECT "
'    strSQLOT = strSQLOT & "TOP 1 NUMPLANEAMIENTO "
'    strSQLOT = strSQLOT & "FROM "
'    strSQLOT = strSQLOT & "ORDENTRAB_CAB "
'    strSQLOT = strSQLOT & "WHERE "
'    strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "' AND "
'    strSQLOT = strSQLOT & "LEFT(NUMPLANEAMIENTO, 1) = '" & strFecha & "' AND "
'    strSQLOT = strSQLOT & "MONTH(ESTACION) = " & Month(strFecha) & " "
'    strSQLOT = strSQLOT & "ORDER BY "
'    strSQLOT = strSQLOT & "NUMPLANEAMIENTO DESC"
'
'    If rstOT.State = 1 Then rstOT.Close
'
'    rstOT.Open strSQLOT, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstOT.EOF Then
'        generarNumeroVale = strFecha & "-" & Format(Val(Mid(Trim(rstOT!NUMPLANEAMIENTO & ""), InStr(1, Trim(rstOT!NUMPLANEAMIENTO & ""), "-") + 1)) + 1, "000000")
'    Else
'        generarNumeroVale = strFecha & "-" & Format(Month(strFecha), "00") & "0001"
'    End If
'
'    rstOT.Close
'
'    Set rstOT = Nothing
'
'    strSQLOT = vbNullString
'
'    Exit Function
'errGenerarNumeroVale:
'    generarNumeroVale = vbNullString
'
'    MsgBox "No. Error: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, _
'            vbCritical, App.ProductName & " - ClsOrdenTrabajo: GenerarNumeroVale"
'
'    Err.Clear
'End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstOT = New ADODB.Recordset
    
    
    strSQLOT = vbNullString
    strSQLOT = strSQLOT & "SELECT "
    strSQLOT = strSQLOT & "ORDENTRAB_CAB.* "
    strSQLOT = strSQLOT & "FROM "
    strSQLOT = strSQLOT & "ORDENTRAB_CAB "
    strSQLOT = strSQLOT & "WHERE "
    strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "'"
    
    If rstOT.State = 1 Then rstOT.Close
    
    rstOT.Open strSQLOT, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOT.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstOT.Close
    
    strSQLOT = vbNullString
    
    Set rstOT = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarOrdenTrabajo() As Boolean
    On Error GoTo errGuardarOrdenTrabajo
    
    'cnn_dbbancos.BeginTrans
    
    strSQLOT = vbNullString
    
    If Not verificarExistencia() Then
        strSQLOT = strSQLOT & "INSERT INTO ORDENTRAB_CAB("
        strSQLOT = strSQLOT & "NUMORDEN, NUMPLANEAMIENTO, FECHA, "
        strSQLOT = strSQLOT & "FECHA_ENTREGA, OBSERVACION, ESTACION, "
        strSQLOT = strSQLOT & "APROBADO_LOGISTICA, APROBADO_CREDITOS, [APROBADO_DISEÑO], "
        strSQLOT = strSQLOT & "MONEDA, TOTAL, FACTURADO) "
        
        strSQLOT = strSQLOT & "VALUES("
        strSQLOT = strSQLOT & "'" & strNumeroOrden & "', '" & strNumeroPlan & "', CVDATE('" & StrFecha & "'), "
        strSQLOT = strSQLOT & "CVDATE('" & strFechaEntrega & "'), " & IIf(strObservacion <> vbNullString, "'" & strObservacion & "'", "NULL") & ", '" & strEstacion & "', "
        strSQLOT = strSQLOT & IIf(bolAprobadoLogistica, "TRUE", "FALSE") & ", " & IIf(bolAprobadoCredito, "TRUE", "FALSE") & ", "
        strSQLOT = strSQLOT & IIf(bolAprobadoDiseno, "TRUE", "FALSE") & ", "
        strSQLOT = strSQLOT & "'" & strCodMoneda & "', " & dblTotal & ", '" & strFacturado & "')"
    Else
        strSQLOT = strSQLOT & "UPDATE "
        strSQLOT = strSQLOT & "ORDENTRAB_CAB "
        strSQLOT = strSQLOT & "SET "
        strSQLOT = strSQLOT & "NUMPLANEAMIENTO = '" & strNumeroPlan & "', "
        strSQLOT = strSQLOT & "FECHA = CVDATE('" & StrFecha & "'), "
        strSQLOT = strSQLOT & "FECHA_ENTREGA = CVDATE('" & strFechaEntrega & "'), "
        strSQLOT = strSQLOT & "OBSERVACION = " & IIf(strObservacion <> vbNullString, "'" & strObservacion & "'", "NULL") & ", "
        strSQLOT = strSQLOT & "ESTACION = '" & StrFecha & "', "
        strSQLOT = strSQLOT & "APROBADO_LOGISTICA = " & IIf(bolAprobadoLogistica, "TRUE", "FALSE") & ", "
        strSQLOT = strSQLOT & "APROBADO_CREDITOS = " & IIf(bolAprobadoCredito, "TRUE", "FALSE") & ", "
        strSQLOT = strSQLOT & "APROBADO_DISEÑO = " & IIf(bolAprobadoDiseno, "TRUE", "FALSE") & ", "
        strSQLOT = strSQLOT & "MONEDA = '" & strCodMoneda & "', "
        strSQLOT = strSQLOT & "TOTAL = " & dblTotal & ", "
        strSQLOT = strSQLOT & "FACTURADO = '" & strFacturado & "' "
        strSQLOT = strSQLOT & "WHERE "
        strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "'"
    End If
    
    cnn_dbbancos.Execute strSQLOT
    
    'cnn_dbbancos.CommitTrans
    
    guardarOrdenTrabajo = True
    
    strSQLSelectAlter = strSQLOT
    strSQLOT = vbNullString
    
    Exit Function
errGuardarOrdenTrabajo:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: GuardarOrdenTrabajo"
    
    'cnn_dbbancos.RollbackTrans
    
    guardarOrdenTrabajo = False
    
    Err.Clear
End Function

Public Sub guardarOTInsumoOneByOne()
    On Error GoTo errGuardarOTInsumoOneByOne
    
    strSQLOT = vbNullString
        
    strSQLOT = strSQLOT & "INSERT INTO ORDENTRAB_INS("
    strSQLOT = strSQLOT & "NUMORDEN, F4NUMFORMULA, F3CODPROCOM, "
    strSQLOT = strSQLOT & "NOMCOMPONENTE, F3ITEM, F3CODPROINS, "
    strSQLOT = strSQLOT & "DESCRIPCION, UMEDIDA, CANTIDAD, "
    strSQLOT = strSQLOT & "F3CANTIDADUNITARIO, F3CANTIDADORDENFORMULA, CANTIDADREQUERIDASTOCK, "
    strSQLOT = strSQLOT & "SECPROCESO, CARGADO) "
    
    strSQLOT = strSQLOT & "VALUES("
    strSQLOT = strSQLOT & "'" & strNumeroOrden & "', " & IIf(strNumeroFormula <> vbNullString, "'" & strNumeroFormula & "'", "NULL") & ", '" & strCodProductoComp & "', "
    strSQLOT = strSQLOT & "'" & strNombreComp & "', " & dblItem & ", '" & strCodProductoIns & "', "
    strSQLOT = strSQLOT & "'" & strDesProductoIns & "', '" & strCodUM & "', " & dblCantidad & ", "
    strSQLOT = strSQLOT & dblCantidadUnitario & ", " & dblCantidadOrdenFormula & ", "
    strSQLOT = strSQLOT & dblCantidadRequeridaStock & ", " & dblSecProceso & ", "
    strSQLOT = strSQLOT & "'" & strNumeroPedido & "')"
    
    cnn_dbbancos.Execute strSQLOT
    
    strSQLSelectAlter = strSQLOT
    strSQLOT = vbNullString
    
    Exit Sub
errGuardarOTInsumoOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: GuardarOTInsumoOneByOne"
    
    Err.Clear
End Sub

Public Sub guardarOTDetalleOneByOne()
    On Error GoTo errGuardarOTDetalleOneByOne
    
    strSQLOT = vbNullString
        
    strSQLOT = strSQLOT & "INSERT INTO ORDENTRAB_DET("
    strSQLOT = strSQLOT & "NUMORDEN, F4NUMFORMULA, F3CODPRO, "
    strSQLOT = strSQLOT & "DESCRIPCION, UMEDIDA, CANTIDAD, "
    strSQLOT = strSQLOT & "OBSERVACION, PRECIO, NUMPEDIDO, "
    strSQLOT = strSQLOT & "CODPROCESO) "
    
    strSQLOT = strSQLOT & "VALUES("
    strSQLOT = strSQLOT & "'" & strNumeroOrden & "', " & IIf(strNumeroFormula <> vbNullString, "'" & strNumeroFormula & "'", "NULL") & ", '" & strCodProducto & "', "
    strSQLOT = strSQLOT & "'" & strDesProducto & "', '" & strCodUM & "', " & dblCantidad & ", "
    strSQLOT = strSQLOT & "'" & strObservacion & "', " & dblPrecio & ", '" & strNumeroPedido & "', "
    strSQLOT = strSQLOT & IIf(strCodProceso <> vbNullString, "'" & strCodProceso & "'", "NULL") & ")"
    
    cnn_dbbancos.Execute strSQLOT
    
    strSQLSelectAlter = strSQLOT
    strSQLOT = vbNullString
    
    Exit Sub
errGuardarOTDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: GuardarOTDetalleOneByOne"
    
    Err.Clear
End Sub
Public Function eliminarOrdenTrabajo() As Boolean
    On Error GoTo errEliminarOrdenTrabajo
    
    'If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F4ORDTRA", "IF4VALES", strNumeroOrden, "T")) > 0 Then
    '
    '    MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
    '
    '    eliminarOrdenTrabajo = False
    '
    '    Exit Function
    'End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLOT = vbNullString
    strSQLOT = strSQLOT & "DELETE "
    strSQLOT = strSQLOT & "FROM "
    strSQLOT = strSQLOT & "ORDENTRAB_CAB "
    strSQLOT = strSQLOT & "WHERE "
    strSQLOT = strSQLOT & "NUMORDEN = '" & strNumeroOrden & "'"
    
    cnn_dbbancos.Execute strSQLOT
    
    'cnn_dbbancos.CommitTrans
    
    eliminarOrdenTrabajo = True
    
    strSQLSelectAlter = strSQLOT
    strSQLOT = vbNullString
    
    Exit Function
errEliminarOrdenTrabajo:
    eliminarOrdenTrabajo = False

    'cnn_dbbancos.RollbackTrans
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsOrdenTrabajo: EliminarOrdenTrabajo"
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'Listar Insumos Pendientes de Descarga en Grilla (QuamtumGrid)
Public Sub listarGrillaInsumoPendienteDescarga(ByVal grilla As dxDBGrid, _
                                                    ByVal strIDOrden As String, _
                                                    Optional ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaInsumoPendienteDescarga
    
'    strSQLOT = vbNullString
'
'    strSQLOT = strSQLOT & "INSERT INTO TMPUTILDESCARGAOPPENDIENTE("
'    strSQLOT = strSQLOT & "NROOP, "
'    strSQLOT = strSQLOT & "NROPEDIDO, "
'    strSQLOT = strSQLOT & "LLAVEOP, "
'    strSQLOT = strSQLOT & "CODPRODUCTOORIGEN, "
'    strSQLOT = strSQLOT & "CODPRODUCTOFINAL, "
'    strSQLOT = strSQLOT & "NOMPRODUCTO, "
'    strSQLOT = strSQLOT & "UM, "
'    strSQLOT = strSQLOT & "CANTIDADORIGEN, "
'    strSQLOT = strSQLOT & "CANTIDADFINAL, "
'    strSQLOT = strSQLOT & "SALDO) "
'
'    strSQLOT = strSQLOT & "IN '" & wrutatemp & "Templus.mdb' "
'
'    strSQLOT = strSQLOT & "SELECT "
'    strSQLOT = strSQLOT & "DET.NUMORDEN, "
'    strSQLOT = strSQLOT & "DET.CARGADO, "
'    strSQLOT = strSQLOT & "'[ ID Orden de Produccion: ' & DET.NUMORDEN & ' ] " & Space(30)
'    strSQLOT = strSQLOT & " [ No. Pedido: ' & DET.CARGADO & ' ] " & Space(30)
'    strSQLOT = strSQLOT & " [ No. Modelo: ' & CAB.ESTACION & ' ] " & Space(30)
'    strSQLOT = strSQLOT & " [ Color: ' & CAB.OBSERVACION & ' ] " & Space(30)
'    strSQLOT = strSQLOT & " [ Cantidad: ' & CAB.TOTAL & ' ]' "
'    strSQLOT = strSQLOT & "AS LLAVE , "
'    strSQLOT = strSQLOT & "DET.F3CODPROINS AS ORIGEN, "
'    strSQLOT = strSQLOT & "DET.F3CODPROINS AS FINAL, "
'    strSQLOT = strSQLOT & "PROD.F5NOMPRO, "
'    strSQLOT = strSQLOT & "MED.F7SIGMED, "
'    'strSQLOT = strSQLOT & "(DET.CANTIDAD - VAL(MOVOP.CANTIDAD & '')) AS CANTIDAD, "
'    strSQLOT = strSQLOT & "DET.CANTIDAD AS CANTIDADORIGEN, "
'    strSQLOT = strSQLOT & "DET.CANTIDAD AS CANTIDADFINAL, "
'    strSQLOT = strSQLOT & "(DET.CANTIDAD - VAL(MOVOP.CANTIDAD & '')) AS SALDO "
'    strSQLOT = strSQLOT & "FROM "
'    strSQLOT = strSQLOT & "(((ORDENTRAB_INS AS DET "
'    strSQLOT = strSQLOT & "LEFT JOIN ORDENTRAB_CAB AS CAB ON CAB.NUMORDEN = DET.NUMORDEN) "
'    strSQLOT = strSQLOT & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPROINS) "
'    strSQLOT = strSQLOT & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
'    strSQLOT = strSQLOT & "LEFT JOIN "
'    strSQLOT = strSQLOT & "("
'    strSQLOT = strSQLOT & "SELECT "
'    strSQLOT = strSQLOT & "CAB.F4ORDTRA, DET.COD_SOLICITUD, DET.F5CODPROORIGINAL, SUM(DET.F3CANPRO) AS CANTIDAD "
'    strSQLOT = strSQLOT & "FROM "
'    strSQLOT = strSQLOT & "IF3VALES AS DET "
'    strSQLOT = strSQLOT & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
'    strSQLOT = strSQLOT & "WHERE "
'    'strSQLOT = strSQLOT & "CAB.F1CODORI IN ('XOP', 'XDP') AND " 'INGRESOS POR DEVOLUCION DE OP Y SALIDAS POR DESPACHO DE OP
'    strSQLOT = strSQLOT & "CAB.F1CODORI IN ('XDP') AND " 'SALIDAS POR DESPACHO DE OP
'    strSQLOT = strSQLOT & "TRIM(CAB.F4ORDTRA & '') <> '' "
'    strSQLOT = strSQLOT & "GROUP BY "
'    strSQLOT = strSQLOT & "CAB.F4ORDTRA, "
'    strSQLOT = strSQLOT & "DET.COD_SOLICITUD, "
'    strSQLOT = strSQLOT & "DET.F5CODPROORIGINAL"
'    strSQLOT = strSQLOT & ") AS MOVOP ON "
'    strSQLOT = strSQLOT & "MOVOP.F4ORDTRA = DET.NUMORDEN AND MOVOP.COD_SOLICITUD = DET.CARGADO AND MOVOP.F5CODPROORIGINAL = DET.F3CODPROINS "
'    strSQLOT = strSQLOT & "WHERE "
'    strSQLOT = strSQLOT & "DET.NUMORDEN = '" & strIDOrden & "' AND "
'    strSQLOT = strSQLOT & "(DET.CANTIDAD - VAL(MOVOP.CANTIDAD & '')) > 0 "
'    strSQLOT = strSQLOT & "ORDER BY "
'    strSQLOT = strSQLOT & "DET.F3ITEM"
    
    If Not grilla Is Nothing Then
'        abrirCnTemporal
'
'        cnDBTemp.Execute "DELETE FROM TMPUTILDESCARGAOPPENDIENTE"
'
'        abrirCnTemporal
'
'        cnn_dbbancos.Execute strSQLOT
        
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna ID de Orden
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "O.P."
                .DisableEditor = True
                .FieldName = "NROOP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = False
            End With
            
            'Columna Numero de Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Pedido"
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = False
            End With
            
            'Columna Datos Resumen (Llave)
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Datos de O.P"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "LLAVEOP"
                .Font.Name = "Arial"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlaveOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Codigo de Producto de Origen
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Codigo Origen"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODPRODUCTOORIGEN"
                .Font.Name = "Arial"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProductoOrigen"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Codigo de Producto Final
            'Set gColumn = .Columns.Add(gedLookupEdit)
            Set gColumn = .Columns.Add(gedButtonEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Codigo Final"
                .BandIndex = 0
                '.DisableEditor = True
                .FieldName = "CODPRODUCTOFINAL"
                .Font.Name = "Arial"
                .Font.Charset = 0
                .ButtonColumn.EditButtonStyle = ebsEllipsis
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProductoFinal"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                
'                With .LookupColumn
'                    .LookupDataset.ADODataset.ConnectionString = cnn_dbbancos
'                    .LookupDataset.ADODataset.CommandText = "SELECT F5CODPRO, F5NOMPRO FROM IF5PLA ORDER BY F5NOMPRO"
'                    .LookupKeyField = "F5CODPRO"
'                    .LookupResultField = "F5CODPRO"
'                    .LookupDataset.Active = True
'                    .ListFieldIndex = 0
'                    .DisplaySize = 30
'                    .LookupCache = True
'                    .ListFieldName = "F5NOMPRO"
'                    .ListWidth = 500
'                End With
            End With
            
            'Columna Descripcion del Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Name = "Arial"
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
                .FieldName = "UM"
                .Font.Name = "Arial"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
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
            
            'Columna Cantidad de Origen
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad Original"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDADORIGEN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Final
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad Final"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .FieldName = "CANTIDADFINAL"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Saldo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Saldo"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'ADD
            'Columna Stock Total
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Stock Total"
                '.Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCK"
                .HeaderAlignment = taCenter
                .ObjectName = "ColStock"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILDESCARGAOPPENDIENTE WHERE SALDO > 0 ORDER BY NOMPRODUCTO"
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPRODUCTOORIGEN"
            
            .Columns.ColumnByFieldName("LLAVEOP").GroupIndex = 0
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLOT
    End If
    
    strSQLOT = vbNullString
    
    Exit Sub
errListarGrillaInsumoPendienteDescarga:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsOrdenTrabajo: ListarGrillaInsumoPendienteDescarga"
    End Select
    
    Err.Clear
End Sub

