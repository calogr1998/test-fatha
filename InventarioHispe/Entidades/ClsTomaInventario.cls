VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTomaInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strAnno                     As String 'F4ANNO
Private strMes                      As String 'F4MES
Private strCodAlmacen               As String 'F2CODALM

Private StrFecha                    As String 'F4FECHA
Private strObservacion              As String 'F4OBSERVACION
Private strValeIngreso              As String 'F4VALEINGRESO
Private strValeIngresoExterno       As String 'F4VALEINGRESOEXTERNO
Private strValeSalida               As String 'F4VALESALIDA
Private strValeSalidaExterno        As String 'F4VALESALIDAEXTERNO
Private bolCierreInventario         As Boolean 'F4CIERRE

Private strFecReg                   As String 'F4FECGRA
Private strUsuReg                   As String 'F4USEGRA
Private strFecMod                   As String 'F4FECMOD
Private strUsuMod                   As String 'F4USEMOD




'Atributos de Detalle de Vale
Private strCodProducto              As String 'F5CODPRO
Private dblStockSistema             As Double 'F3STOCKSISTEMA
Private dblStockFisico              As Double 'F3STOCKFISICO
Private dblDiferencia               As Double 'F3DIFERENCIA




Private strSQLSelectAlter           As String

Private rstTomaInv           As ADODB.Recordset
Private strSQLTomaInv        As String


Public Property Let AnnoTI(ByVal value As String)
    strAnno = value
End Property

Public Property Get AnnoTI() As String
    AnnoTI = strAnno
End Property

Public Property Let MesTI(ByVal value As String)
    strMes = value
End Property

Public Property Get MesTI() As String
    MesTI = strMes
End Property

Public Property Let CodigoAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property

Public Property Let Fecha(ByVal value As String)
    StrFecha = value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property

Public Property Let Observacion(ByVal value As String)
    strObservacion = value
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property

Public Property Let ValeIngreso(ByVal value As String)
    strValeIngreso = value
End Property

Public Property Get ValeIngreso() As String
    ValeIngreso = strValeIngreso
End Property

Public Property Let ValeIngresoExterno(ByVal value As String)
    strValeIngresoExterno = value
End Property

Public Property Get ValeIngresoExterno() As String
    ValeIngresoExterno = strValeIngresoExterno
End Property

Public Property Let ValeSalida(ByVal value As String)
    strValeSalida = value
End Property

Public Property Get ValeSalida() As String
    ValeSalida = strValeSalida
End Property

Public Property Let ValeSalidaExterno(ByVal value As String)
    strValeSalidaExterno = value
End Property

Public Property Get ValeSalidaExterno() As String
    ValeSalidaExterno = strValeSalidaExterno
End Property

Public Property Let CierreInventario(ByVal value As Boolean)
    bolCierreInventario = value
End Property

Public Property Get CierreInventario() As Boolean
    CierreInventario = bolCierreInventario
End Property




Public Property Let FecReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FecReg() As String
    FecReg = strFecReg
End Property

Public Property Let UsuReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuReg() As String
    UsuReg = strUsuReg
End Property

Public Property Let FecMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FecMod() As String
    FecMod = strFecMod
End Property

Public Property Let UsuMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuMod() As String
    UsuMod = strUsuMod
End Property




'Atributos de Detalle de Inventario
Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Property Let StockSistema(ByVal value As Double)
    dblStockSistema = value
End Property

Public Property Get StockSistema() As Double
    StockSistema = dblStockSistema
End Property

Public Property Let StockFisico(ByVal value As Double)
    dblStockFisico = value
End Property

Public Property Get StockFisico() As Double
    StockFisico = dblStockFisico
End Property

Public Property Let Diferencia(ByVal value As Double)
    dblDiferencia = value
End Property

Public Property Get Diferencia() As Double
    Diferencia = dblDiferencia
End Property





Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strAnno = vbNullString 'F4ANNO
    strMes = vbNullString 'F4MES
    strCodAlmacen = vbNullString 'F2CODALM
    
    StrFecha = vbNullString 'F4FECHA
    strObservacion = vbNullString 'F4OBSERVACION
    strValeIngreso = vbNullString 'F4VALEINGRESO
    strValeIngresoExterno = vbNullString 'F4VALEINGRESOEXTERNO
    strValeSalida = vbNullString 'F4VALESALIDA
    strValeSalidaExterno = vbNullString  'F4VALESALIDAEXTERNO
    bolCierreInventario = False  'F4CIERRE
    
    strFecReg = vbNullString   'F4FECGRA
    strUsuReg = vbNullString   'F4USEGRA
    strFecMod = vbNullString   'F4FECMOD
    strUsuMod = vbNullString   'F4USEMOD
    
    strSQLSelectAlter = vbNullString
End Sub

Public Sub inicializarEntidadesDetalle()
    strCodProducto = vbNullString 'F5CODPRO
    dblStockSistema = 0 'F3STOCKSISTEMA
    dblStockFisico = 0 'F3STOCKFISICO
    dblDiferencia = 0 'F3DIFERENCIA
End Sub

Public Function obtenerTomaInventario() As Boolean
    On Error GoTo errObtenerTomaInventario
    
    Set rstTomaInv = New ADODB.Recordset

    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "* "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close

    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    If Not rstTomaInv.EOF Then
        strAnno = Trim(rstTomaInv!F4ANNO & "") 'F4ANNO
        strMes = Trim(rstTomaInv!F4MES & "") 'F4MES
        strCodAlmacen = Trim(rstTomaInv!f2codalm & "") 'F2CODALM
        
        StrFecha = Trim(rstTomaInv!F4FECHA & "") 'F4FECHA
        strObservacion = Trim(rstTomaInv!F4OBSERVACION & "") 'F4OBSERVACION
        strValeIngreso = Trim(rstTomaInv!F4VALEINGRESO & "") 'F4VALEINGRESO
        strValeIngresoExterno = Trim(rstTomaInv!F4VALEINGRESOEXTERNO & "") 'F4VALEINGRESOEXTERNO
        strValeSalida = Trim(rstTomaInv!F4VALESALIDA & "") 'F4VALESALIDA
        strValeSalidaExterno = Trim(rstTomaInv!F4VALESALIDAEXTERNO & "") 'F4VALESALIDAEXTERNO
        bolCierreInventario = CBool(rstTomaInv!F4CIERRE)  'F4CIERRE
        
        strFecReg = Trim(rstTomaInv!F4FECGRA & "")   'F4FECGRA
        strUsuReg = Trim(rstTomaInv!F4USEGRA & "")   'F4USEGRA
        strFecMod = Trim(rstTomaInv!F4FECMOD & "")   'F4FECMOD
        strUsuMod = Trim(rstTomaInv!F4USEMOD & "")   'F4USEMOD
        
        obtenerTomaInventarioDetalle
        
        obtenerTomaInventario = True
    Else
        obtenerTomaInventario = False
    End If

    rstTomaInv.Close

    Set rstTomaInv = Nothing

    Exit Function
errObtenerTomaInventario:
    obtenerTomaInventario = False

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: ObtenerTomaInventario"
    
    Err.Clear
End Function

Public Sub obtenerConfigTomaInventario()
    On Error GoTo errObtenerConfigTomaInventario

    Set rstTomaInv = New ADODB.Recordset

    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "* "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close

    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    If Not rstTomaInv.EOF Then
        strAnno = Trim(rstTomaInv!F4ANNO & "") 'F4ANNO
        strMes = Trim(rstTomaInv!F4MES & "") 'F4MES
        strCodAlmacen = Trim(rstTomaInv!f2codalm & "") 'F2CODALM
        
        StrFecha = Trim(rstTomaInv!F4FECHA & "") 'F4FECHA
        strObservacion = Trim(rstTomaInv!F4OBSERVACION & "") 'F4OBSERVACION
        strValeIngreso = Trim(rstTomaInv!F4VALEINGRESO & "") 'F4VALEINGRESO
        strValeIngresoExterno = Trim(rstTomaInv!F4VALEINGRESOEXTERNO & "") 'F4VALEINGRESOEXTERNO
        strValeSalida = Trim(rstTomaInv!F4VALESALIDA & "") 'F4VALESALIDA
        strValeSalidaExterno = Val(rstTomaInv!F4VALESALIDAEXTERNO & "") 'F4VALESALIDAEXTERNO
        bolCierreInventario = CBool(rstTomaInv!F4CIERRE)  'F4CIERRE
        
        strFecReg = Trim(rstTomaInv!F4FECGRA & "")   'F4FECGRA
        strUsuReg = Trim(rstTomaInv!F4USEGRA & "")   'F4USEGRA
        strFecMod = Trim(rstTomaInv!F4FECMOD & "")   'F4FECMOD
        strUsuMod = Trim(rstTomaInv!F4USEMOD & "")   'F4USEMOD
    End If
    
    rstTomaInv.Close
    
    Set rstTomaInv = Nothing
    
    Exit Sub
errObtenerConfigTomaInventario:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: ObtenerConfigTomaInventario"
    
    Err.Clear
End Sub

Public Sub obtenerTomaInventarioDetalle()
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILTOMAINVENTARIO"
    
    abrirCnTemporal
    
    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "INSERT INTO TMPUTILTOMAINVENTARIO("
        strSQLTomaInv = strSQLTomaInv & "FAMILIA, SUBFAMILIA, CODPRODUCTO, "
        strSQLTomaInv = strSQLTomaInv & "NOMPRODUCTO, UM, STOCKSISTEMA, "
        strSQLTomaInv = strSQLTomaInv & "STOCKFISICO, DIFERENCIA, "
        strSQLTomaInv = strSQLTomaInv & "COSTOPROMEDIO, STOCKSISTEMAVALOR, "
        strSQLTomaInv = strSQLTomaInv & "STOCKFISICOVALOR"
        strSQLTomaInv = strSQLTomaInv & ") "
    
    strSQLTomaInv = strSQLTomaInv & "IN '" & wrutatemp & "Templus.mdb' "
    
    strSQLTomaInv = strSQLTomaInv & "SELECT "
        strSQLTomaInv = strSQLTomaInv & "FAM.F7DESCON, "
        strSQLTomaInv = strSQLTomaInv & "SFAM.F7DESCON, "
        strSQLTomaInv = strSQLTomaInv & "DET.F5CODPRO, "
        strSQLTomaInv = strSQLTomaInv & "PROD.F5NOMPRO, "
        strSQLTomaInv = strSQLTomaInv & "MED.F7SIGMED, "
        
        If bolCierreInventario Then
            strSQLTomaInv = strSQLTomaInv & "DET.F3STOCKSISTEMA, "
            strSQLTomaInv = strSQLTomaInv & "DET.F3STOCKFISICO, "
            strSQLTomaInv = strSQLTomaInv & "DET.F3DIFERENCIA, "
        Else
            strSQLTomaInv = strSQLTomaInv & "INVENTARIOACTUAL.SALDO AS STOCKSISTEMA, "
            strSQLTomaInv = strSQLTomaInv & "DET.F3STOCKFISICO, "
            strSQLTomaInv = strSQLTomaInv & "(DET.F3STOCKFISICO - INVENTARIOACTUAL.SALDO) AS DIFERENCIA, "
        End If
        
        strSQLTomaInv = strSQLTomaInv & "INVENTARIOACTUAL.COSTOPROMEDIO, "
        strSQLTomaInv = strSQLTomaInv & "(DET.F3STOCKSISTEMA * INVENTARIOACTUAL.COSTOPROMEDIO) AS STOCKSISTEMAVALOR, "
        strSQLTomaInv = strSQLTomaInv & "(DET.F3STOCKFISICO * INVENTARIOACTUAL.COSTOPROMEDIO) AS STOCKFISICOVALOR "
        
    strSQLTomaInv = strSQLTomaInv & "FROM "
        strSQLTomaInv = strSQLTomaInv & "((((H3TOMAINV AS DET "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO) "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01) "
        
        
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN ("
        strSQLTomaInv = strSQLTomaInv & "SELECT "
        strSQLTomaInv = strSQLTomaInv & "DET.F5CODPRO, "
        strSQLTomaInv = strSQLTomaInv & "VAL(FORMAT(  SUM(VAL(FORMAT(DET.F3CANPRO, '#.0000')) * IIF(DET.TIPO = 'I', 1, -1))  , '#.0000')) AS SALDO, "
        'strSQLTomaInv = strSQLTomaInv & "SUM((DET." & IIf(strCodMonedaValor = "S", "F3VALVTA", "F3VALDOL") & " * VAL(FORMAT(DET.F3CANPRO, '#.0000'))) * IIF(DET.TIPO = 'I', 1, -1)) AS VALOR, "
        strSQLTomaInv = strSQLTomaInv & "SUM((DET.F3VALVTA * VAL(FORMAT(DET.F3CANPRO, '#.0000'))) * IIF(DET.TIPO = 'I', 1, -1)) AS VALOR, "
        strSQLTomaInv = strSQLTomaInv & "VAL(FORMAT(  IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 0, (VALOR / IIF(VAL(FORMAT(SALDO, '#.0000')) <= 0, 1, VAL(FORMAT(SALDO, '#.0000'))) ) )  , '#0.0000')) AS COSTOPROMEDIO "
        
        strSQLTomaInv = strSQLTomaInv & "FROM "
        strSQLTomaInv = strSQLTomaInv & "(SELECT "
        strSQLTomaInv = strSQLTomaInv & "DET.F2CODALM, DET.F4NUMVAL, CAB.F4FECVAL, DET.TIPO, DET.F5CODPRO, DET.F3CANPRO, DET.F3VALVTA, DET.F3VALDOL "
        strSQLTomaInv = strSQLTomaInv & "FROM "
        strSQLTomaInv = strSQLTomaInv & "IF3VALES AS DET "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM) "
        strSQLTomaInv = strSQLTomaInv & "WHERE "
        strSQLTomaInv = strSQLTomaInv & "CAB.F4FECVAL <= CVDATE('" & DateSerial(Val(strAnno), Val(strMes) + 1, 0) & "') AND "
        
        strSQLTomaInv = strSQLTomaInv & "CAB.F2CODALM = '" & strCodAlmacen & "' "
        
        If Not bolCierreInventario Then
            If strValeIngreso <> vbNullString Then
                strSQLTomaInv = strSQLTomaInv & "AND CAB.F4NUMVAL NOT IN ('" & strValeIngreso & "') "
            End If
            
            If strValeSalida <> vbNullString Then
                strSQLTomaInv = strSQLTomaInv & "AND CAB.F4NUMVAL NOT IN ('" & strValeSalida & "') "
            End If
        End If
        
        strSQLTomaInv = strSQLTomaInv & "ORDER BY "
        strSQLTomaInv = strSQLTomaInv & "CAB.F4FECVAL, DET.TIPO, DET.F4NUMVAL) AS DET "
        strSQLTomaInv = strSQLTomaInv & "LEFT JOIN IF4VALES AS CAB "
        strSQLTomaInv = strSQLTomaInv & "ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        
        strSQLTomaInv = strSQLTomaInv & "WHERE "
        strSQLTomaInv = strSQLTomaInv & "CAB.F1CODORI NOT IN ('XCS') "
        
        strSQLTomaInv = strSQLTomaInv & "GROUP BY "
        strSQLTomaInv = strSQLTomaInv & "DET.F5CODPRO"
        strSQLTomaInv = strSQLTomaInv & ") AS INVENTARIOACTUAL "
        strSQLTomaInv = strSQLTomaInv & "ON INVENTARIOACTUAL.F5CODPRO = DET.F5CODPRO "
        
        
    strSQLTomaInv = strSQLTomaInv & "WHERE "
            strSQLTomaInv = strSQLTomaInv & "DET.F4ANNO = '" & strAnno & "' AND "
            strSQLTomaInv = strSQLTomaInv & "DET.F4MES = '" & strMes & "' AND "
            strSQLTomaInv = strSQLTomaInv & "DET.F2CODALM = '" & strCodAlmacen & "'"
    
    cnn_dbbancos.Execute strSQLTomaInv
    
    strSQLTomaInv = vbNullString
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTomaInv = New ADODB.Recordset


    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV.* "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close

    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    If Not rstTomaInv.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If

    rstTomaInv.Close

    strSQLTomaInv = vbNullString

    Set rstTomaInv = Nothing

    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: VerificarExistencia"

    verificarExistencia = False

    Err.Clear
End Function

Public Function guardarTomaInventario() As Boolean
    On Error GoTo errGuardarTomaInventario

    strSQLTomaInv = vbNullString
    
    If Not verificarExistencia Then
        strSQLTomaInv = strSQLTomaInv & "INSERT INTO H4TOMAINV("
        strSQLTomaInv = strSQLTomaInv & "F4ANNO, F4MES, F2CODALM, "
        strSQLTomaInv = strSQLTomaInv & "F4FECHA, F4OBSERVACION, F4VALEINGRESO, "
        strSQLTomaInv = strSQLTomaInv & "F4VALEINGRESOEXTERNO, F4VALESALIDA, "
        strSQLTomaInv = strSQLTomaInv & "F4VALESALIDAEXTERNO, "
        strSQLTomaInv = strSQLTomaInv & "F4FECGRA, F4USEGRA) "
        
        strSQLTomaInv = strSQLTomaInv & "VALUES("
        strSQLTomaInv = strSQLTomaInv & "'" & strAnno & "', "
        strSQLTomaInv = strSQLTomaInv & "'" & strMes & "', "
        strSQLTomaInv = strSQLTomaInv & "'" & strCodAlmacen & "', "
        strSQLTomaInv = strSQLTomaInv & "'" & StrFecha & "', "
        strSQLTomaInv = strSQLTomaInv & "'" & strObservacion & "', "
        strSQLTomaInv = strSQLTomaInv & "'" & strValeIngreso & "', "
        strSQLTomaInv = strSQLTomaInv & IIf(strValeIngresoExterno <> vbNullString, "'" & strValeIngresoExterno & "'", "NULL") & ", "
        strSQLTomaInv = strSQLTomaInv & "'" & strValeSalida & "', "
        strSQLTomaInv = strSQLTomaInv & IIf(strValeSalidaExterno <> vbNullString, "'" & strValeSalidaExterno & "'", "NULL") & ", "
        strSQLTomaInv = strSQLTomaInv & "CVDATE('" & strFecReg & "'), "
        strSQLTomaInv = strSQLTomaInv & "'" & strUsuReg & "')"
    Else
        strSQLTomaInv = strSQLTomaInv & "UPDATE "
        strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
        strSQLTomaInv = strSQLTomaInv & "SET "
        strSQLTomaInv = strSQLTomaInv & "F4FECHA = CVDATE('" & StrFecha & "'), "
        strSQLTomaInv = strSQLTomaInv & "F4OBSERVACION = '" & strObservacion & "', "
        strSQLTomaInv = strSQLTomaInv & "F4VALEINGRESO = '" & strValeIngreso & "', "
        strSQLTomaInv = strSQLTomaInv & "F4VALEINGRESOEXTERNO = " & IIf(strValeIngresoExterno <> vbNullString, "'" & strValeIngresoExterno & "'", "NULL") & ", "
        strSQLTomaInv = strSQLTomaInv & "F4VALESALIDA = '" & strValeSalida & "', "
        strSQLTomaInv = strSQLTomaInv & "F4VALESALIDAEXTERNO = " & IIf(strValeSalidaExterno <> vbNullString, "'" & strValeSalidaExterno & "'", "NULL") & ", "
        strSQLTomaInv = strSQLTomaInv & "F4FECMOD = CVDATE('" & strFecMod & "'), "
        strSQLTomaInv = strSQLTomaInv & "F4USEMOD = '" & strUsuMod & "' "
        
        strSQLTomaInv = strSQLTomaInv & "WHERE "
        strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
        strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
        strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"
    End If

    cnn_dbbancos.Execute strSQLTomaInv
    
    guardarTomaInventario = True

    strSQLSelectAlter = strSQLTomaInv
    strSQLTomaInv = vbNullString

    Exit Function
errGuardarTomaInventario:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vuelva a intentarlo nuevamente.", _
            vbInformation + vbOKOnly, App.ProductName & " - ClsTomaInventario: GuardarTomaInventario"
    
    guardarTomaInventario = False
    
    Err.Clear
End Function

Public Sub guardarTomaInvDetalleOneByOne()
    On Error GoTo errGuardarTomaInvDetalleOneByOne
    
    strSQLTomaInv = vbNullString
    
    strSQLTomaInv = strSQLTomaInv & "INSERT INTO H3TOMAINV("
    strSQLTomaInv = strSQLTomaInv & "F4ANNO, F4MES, F2CODALM, "
    strSQLTomaInv = strSQLTomaInv & "F5CODPRO, F3STOCKSISTEMA, "
    strSQLTomaInv = strSQLTomaInv & "F3STOCKFISICO, F3DIFERENCIA) "
    
    strSQLTomaInv = strSQLTomaInv & "VALUES("
    strSQLTomaInv = strSQLTomaInv & "'" & strAnno & "', "
    strSQLTomaInv = strSQLTomaInv & "'" & strMes & "', "
    strSQLTomaInv = strSQLTomaInv & "'" & strCodAlmacen & "', "
    strSQLTomaInv = strSQLTomaInv & "'" & strCodProducto & "', "
    strSQLTomaInv = strSQLTomaInv & dblStockSistema & ", "
    strSQLTomaInv = strSQLTomaInv & dblStockFisico & ", "
    strSQLTomaInv = strSQLTomaInv & dblDiferencia
    strSQLTomaInv = strSQLTomaInv & ")"
    
    cnn_dbbancos.Execute strSQLTomaInv

    strSQLSelectAlter = strSQLTomaInv
    strSQLTomaInv = vbNullString

    Exit Sub
errGuardarTomaInvDetalleOneByOne:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: GuardarTomaInvDetalleOneByOne"

    Err.Clear
End Sub

Public Function eliminarTomaInventario() As Boolean
    On Error GoTo errEliminarTomaInventario
    
    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "DELETE "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"
    
    cnn_dbbancos.Execute strSQLTomaInv
    
    eliminarTomaInventario = True
    
    strSQLSelectAlter = strSQLTomaInv
    strSQLTomaInv = vbNullString
    
    Exit Function
errEliminarTomaInventario:
    eliminarTomaInventario = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: EliminarTomaInventario"

    Err.Clear
End Function

Public Function cerrarTomaInventario() As Boolean
    On Error GoTo errCerrarTomaInventario
    
    strSQLTomaInv = vbNullString
    
    strSQLTomaInv = strSQLTomaInv & "UPDATE "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "SET "
    strSQLTomaInv = strSQLTomaInv & "F4VALEINGRESO = '" & strValeIngreso & "', "
    strSQLTomaInv = strSQLTomaInv & "F4VALEINGRESOEXTERNO = " & IIf(strValeIngresoExterno <> vbNullString, "'" & strValeIngresoExterno & "'", "NULL") & ", "
    strSQLTomaInv = strSQLTomaInv & "F4VALESALIDA = '" & strValeSalida & "', "
    strSQLTomaInv = strSQLTomaInv & "F4VALESALIDAEXTERNO = " & IIf(strValeSalidaExterno <> vbNullString, "'" & strValeSalidaExterno & "'", "NULL") & ", "
    strSQLTomaInv = strSQLTomaInv & "F4CIERRE = " & IIf(bolCierreInventario, "TRUE", "FALSE") & ","
    strSQLTomaInv = strSQLTomaInv & "F4FECMOD = CVDATE('" & strFecMod & "'), "
    strSQLTomaInv = strSQLTomaInv & "F4USEMOD = '" & strUsuMod & "' "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4MES = '" & strMes & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "'"

    cnn_dbbancos.Execute strSQLTomaInv
    
    cerrarTomaInventario = True

    strSQLSelectAlter = strSQLTomaInv
    strSQLTomaInv = vbNullString

    Exit Function
errCerrarTomaInventario:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Vuelva a intentarlo nuevamente.", _
            vbInformation + vbOKOnly, App.ProductName & " - ClsTomaInventario: CerrarTomaInventario"
    
    cerrarTomaInventario = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub listarAlmacenSoloSeleccion(ByVal comboList As Object)
    
    On Error GoTo errListarAlmacenSoloSeleccion
    
    Set rstTomaInv = New ADODB.Recordset
    
    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "CAB.F2CODALM, "
    strSQLTomaInv = strSQLTomaInv & "ALM.F2NOMALM "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV AS CAB "
    strSQLTomaInv = strSQLTomaInv & "LEFT JOIN EF2ALMACENES AS ALM ON ALM.F2CODALM = CAB.F2CODALM "
    strSQLTomaInv = strSQLTomaInv & "GROUP BY "
    strSQLTomaInv = strSQLTomaInv & "CAB.F2CODALM, "
    strSQLTomaInv = strSQLTomaInv & "ALM.F2NOMALM "
    strSQLTomaInv = strSQLTomaInv & "ORDER BY "
    strSQLTomaInv = strSQLTomaInv & "CAB.F2CODALM"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTomaInv.EOF Then
        comboList.Clear
        
        Do While Not rstTomaInv.EOF
            comboList.AddItem Trim(rstTomaInv!F2NOMALM & "") & Space(50) & Trim(rstTomaInv!f2codalm & "")
            
            rstTomaInv.MoveNext
        Loop
    End If
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    strSQLTomaInv = vbNullString
    
    Set rstTomaInv = Nothing
    
    Exit Sub
errListarAlmacenSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: ListarAlmacenSoloSeleccion"
    
    Err.Clear
End Sub

Public Sub listarAnnoSoloSeleccion(ByVal comboList As Object, _
                                    ByVal strCodAlmacen As String)
    
    On Error GoTo errListarAnnoSoloSeleccion
    
    Set rstTomaInv = New ADODB.Recordset
    
    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "' "
    strSQLTomaInv = strSQLTomaInv & "GROUP BY "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO "
    strSQLTomaInv = strSQLTomaInv & "ORDER BY "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTomaInv.EOF Then
        comboList.Clear
        
        Do While Not rstTomaInv.EOF
            comboList.AddItem Trim(rstTomaInv!F4ANNO & "")
            
            rstTomaInv.MoveNext
        Loop
    End If
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    strSQLTomaInv = vbNullString
    
    Set rstTomaInv = Nothing
    
    Exit Sub
errListarAnnoSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: ListarAnnoSoloSeleccion"
    
    Err.Clear
End Sub

Public Sub listarMesSoloSeleccion(ByVal comboList As Object, _
                                    ByVal strCodAlmacen As String, _
                                    ByVal strAnno As String)
    
    On Error GoTo errListarMesSoloSeleccion
    
    Set rstTomaInv = New ADODB.Recordset
    
    strSQLTomaInv = vbNullString
    strSQLTomaInv = strSQLTomaInv & "SELECT "
    strSQLTomaInv = strSQLTomaInv & "F4MES "
    strSQLTomaInv = strSQLTomaInv & "FROM "
    strSQLTomaInv = strSQLTomaInv & "H4TOMAINV "
    strSQLTomaInv = strSQLTomaInv & "WHERE "
    strSQLTomaInv = strSQLTomaInv & "F2CODALM = '" & strCodAlmacen & "' AND "
    strSQLTomaInv = strSQLTomaInv & "F4ANNO = '" & strAnno & "' "
    strSQLTomaInv = strSQLTomaInv & "GROUP BY "
    strSQLTomaInv = strSQLTomaInv & "F4MES "
    strSQLTomaInv = strSQLTomaInv & "ORDER BY "
    strSQLTomaInv = strSQLTomaInv & "F4MES"
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    rstTomaInv.Open strSQLTomaInv, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTomaInv.EOF Then
        comboList.Clear
        
        Do While Not rstTomaInv.EOF
            comboList.AddItem UCase(Format("01/" & Format(Trim(rstTomaInv!F4MES & ""), "00") & "/" & strAnno, "MMMM")) & Space(50) & Trim(rstTomaInv!F4MES & "")
            
            rstTomaInv.MoveNext
        Loop
    End If
    
    If rstTomaInv.State = 1 Then rstTomaInv.Close
    
    strSQLTomaInv = vbNullString
    
    Set rstTomaInv = Nothing
    
    Exit Sub
errListarMesSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTomaInventario: ListarMesSoloSeleccion"
    
    Err.Clear
End Sub

