VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo               As String
Private strDescripcion          As String
Private strCodTipoExistencia    As String
Private strCtaContable          As String
Private strAnexo                As String
Private strCtaContableImp       As String
Private strAnexoImp             As String
Private strCodGasto             As String
Private strCodCentro            As String

Private bolAfecto               As Boolean
Private bolDescontinuado        As Boolean
Private bolParaVenta            As Boolean


Private strCodigoFab            As String
Private strDescripcionFab       As String
Private strCodUM                As String
Private strModelo               As String
Private strCodAlmacen           As String
Private strCodMarca             As String

Private bolEsImportado          As Boolean

Private dblStockMin             As Double
Private dblStockMax             As Double
Private dblStockReposicion      As Double

Private dblPorcentajeDemasia    As Double

Private strCtaContableVenta     As String
Private strCtaContableInventa   As String

Private strFecIngreso           As String
Private strUsuIngreso           As String
Private strFecModificacion      As String
Private strUsuModificacion      As String

'Atributos Adicionales
Private strAbrevTipoExistencia  As String
Private bolEsInsumoParaOP       As Boolean
Private strCodSubFamilia        As String
Private bolTieneMovimientoEnAlmacen As Boolean
Private strColor                As String
Private strTalla                As String
Private strMarcaModelo          As String

Private strCtaContableVta       As String
Private strAnexoVta             As String
Private strCtaContableImpVta    As String
Private strAnexoImpVta          As String


Private strSQLSelectAlter       As String

Private rstBien               As ADODB.Recordset
Private strSQLBien            As String



Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Descripcion(ByVal Value As String)
    strDescripcion = Value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let CodTipoExistencia(ByVal Value As String)
    strCodTipoExistencia = Value
End Property

Public Property Get CodTipoExistencia() As String
    CodTipoExistencia = strCodTipoExistencia
End Property

Public Property Let CtaContable(ByVal Value As String)
    strCtaContable = Value
End Property

Public Property Get CtaContable() As String
    CtaContable = strCtaContable
End Property

Public Property Let Anexo(ByVal Value As String)
    strAnexo = Value
End Property

Public Property Get Anexo() As String
    Anexo = strAnexo
End Property

Public Property Let CtaContableImportacion(ByVal Value As String)
    strCtaContableImp = Value
End Property

Public Property Get CtaContableImportacion() As String
    CtaContableImportacion = strCtaContableImp
End Property

Public Property Let AnexoImportacion(ByVal Value As String)
    strAnexoImp = Value
End Property

Public Property Get AnexoImportacion() As String
    AnexoImportacion = strAnexoImp
End Property

Public Property Let CodGasto(ByVal Value As String)
    strCodGasto = Value
End Property

Public Property Get CodGasto() As String
    CodGasto = strCodGasto
End Property

Public Property Let CodCentroCosto(ByVal Value As String)
    strCodCentro = Value
End Property

Public Property Get CodCentroCosto() As String
    CodCentroCosto = strCodCentro
End Property


Public Property Let Afecto(ByVal Value As Boolean)
    bolAfecto = Value
End Property

Public Property Get Afecto() As Boolean
    Afecto = bolAfecto
End Property

Public Property Let Descontinuado(ByVal Value As Boolean)
    bolDescontinuado = Value
End Property

Public Property Get Descontinuado() As Boolean
    Descontinuado = bolDescontinuado
End Property

Public Property Let ParaVenta(ByVal Value As Boolean)
    bolParaVenta = Value
End Property

Public Property Get ParaVenta() As Boolean
    ParaVenta = bolParaVenta
End Property



Public Property Let CodigoFab(ByVal Value As String)
    strCodigoFab = Value
End Property

Public Property Get CodigoFab() As String
    CodigoFab = strCodigoFab
End Property

Public Property Let DescripcionFab(ByVal Value As String)
    strDescripcionFab = Value
End Property

Public Property Get DescripcionFab() As String
    DescripcionFab = strDescripcionFab
End Property

Public Property Let CodUM(ByVal Value As String)
    strCodUM = Value
End Property

Public Property Get CodUM() As String
    CodUM = strCodUM
End Property

Public Property Let Modelo(ByVal Value As String)
    strModelo = Value
End Property

Public Property Get Modelo() As String
    Modelo = strModelo
End Property

Public Property Let CodAlmacen(ByVal Value As String)
    strCodAlmacen = Value
End Property

Public Property Get CodAlmacen() As String
    CodAlmacen = strCodAlmacen
End Property

Public Property Let CodMarca(ByVal Value As String)
    strCodMarca = Value
End Property

Public Property Get CodMarca() As String
    CodMarca = strCodMarca
End Property



Public Property Let EsImportado(ByVal Value As Boolean)
    bolEsImportado = Value
End Property

Public Property Get EsImportado() As Boolean
    EsImportado = bolEsImportado
End Property



Public Property Let StockMin(ByVal Value As Double)
    dblStockMin = Value
End Property

Public Property Get StockMin() As Double
    StockMin = dblStockMin
End Property

Public Property Let StockMax(ByVal Value As Double)
    dblStockMax = Value
End Property

Public Property Get StockMax() As Double
    StockMax = dblStockMax
End Property
'StockReposicion
Public Property Let StockReposicion(ByVal Value As Double)
    dblStockReposicion = Value
End Property

Public Property Get StockReposicion() As Double
    StockReposicion = dblStockReposicion
End Property


Public Property Let PorcentajeDemasia(ByVal Value As Double)
    dblPorcentajeDemasia = Value
End Property

Public Property Get PorcentajeDemasia() As Double
    PorcentajeDemasia = dblPorcentajeDemasia
End Property




Public Property Let CtaContableVenta(ByVal Value As String)
    strCtaContableVenta = Value
End Property

Public Property Get CtaContableVenta() As String
    CtaContableVenta = strCtaContableVenta
End Property

Public Property Let CtaContableInventa(ByVal Value As String)
    strCtaContableInventa = Value
End Property

Public Property Get CtaContableInventa() As String
    CtaContableInventa = strCtaContableInventa
End Property




Public Property Let FechaIngreso(ByVal Value As String)
    strFecIngreso = Value
End Property

Public Property Get FechaIngreso() As String
    FechaIngreso = strFecIngreso
End Property

Public Property Let UsuarioIngreso(ByVal Value As String)
    strUsuIngreso = Value
End Property

Public Property Get UsuarioIngreso() As String
    UsuarioIngreso = strUsuIngreso
End Property

Public Property Let FechaModificacion(ByVal Value As String)
    strFecModificacion = Value
End Property

Public Property Get FechaModificacion() As String
    FechaModificacion = strFecModificacion
End Property

Public Property Let UsuarioModificacion(ByVal Value As String)
    strUsuModificacion = Value
End Property

Public Property Get UsuarioModificacion() As String
    UsuarioModificacion = strUsuModificacion
End Property




'AbrevTipoExistencia
Public Property Let AbrevTipoExistencia(ByVal Value As String)
    strAbrevTipoExistencia = Value
End Property

Public Property Get AbrevTipoExistencia() As String
    AbrevTipoExistencia = strAbrevTipoExistencia
End Property
'EsInsumoParaOP
Public Property Let EsInsumoParaOP(ByVal Value As Boolean)
    bolEsInsumoParaOP = Value
End Property

Public Property Get EsInsumoParaOP() As Boolean
    EsInsumoParaOP = bolEsInsumoParaOP
End Property
'CodSubFamilia
Public Property Let CodigoSubFamilia(ByVal Value As String)
    strCodSubFamilia = Value
End Property

Public Property Get CodigoSubFamilia() As String
    CodigoSubFamilia = strCodSubFamilia
End Property
'TieneMovimientoEnAlmacen
Public Property Let TieneMovimientoEnAlmacen(ByVal Value As Boolean)
    bolTieneMovimientoEnAlmacen = Value
End Property

Public Property Get TieneMovimientoEnAlmacen() As Boolean
    TieneMovimientoEnAlmacen = bolTieneMovimientoEnAlmacen
End Property
'Color
Public Property Let Color(ByVal Value As String)
    strColor = Value
End Property

Public Property Get Color() As String
    Color = strColor
End Property
'Talla
Public Property Let Talla(ByVal Value As String)
    strTalla = Value
End Property

Public Property Get Talla() As String
    Talla = strTalla
End Property
'Marca de Modelo
Public Property Let MarcaModelo(ByVal Value As String)
    strMarcaModelo = Value
End Property

Public Property Get MarcaModelo() As String
    MarcaModelo = strMarcaModelo
End Property




Public Property Let CtaContableVta(ByVal Value As String)
    strCtaContableVta = Value
End Property

Public Property Get CtaContableVta() As String
    CtaContableVta = strCtaContableVta
End Property

Public Property Let AnexoVta(ByVal Value As String)
    strAnexoVta = Value
End Property

Public Property Get AnexoVta() As String
    AnexoVta = strAnexoVta
End Property

Public Property Let CtaContableImportacionVta(ByVal Value As String)
    strCtaContableImpVta = Value
End Property

Public Property Get CtaContableImportacionVta() As String
    CtaContableImportacionVta = strCtaContableImpVta
End Property

Public Property Let AnexoImportacionVta(ByVal Value As String)
    strAnexoImpVta = Value
End Property

Public Property Get AnexoImportacionVta() As String
    AnexoImportacionVta = strAnexoImpVta
End Property




Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    strDescripcion = vbNullString
    strCodTipoExistencia = vbNullString
    strCtaContable = vbNullString
    strAnexo = vbNullString
    strCtaContableImp = vbNullString
    strAnexoImp = vbNullString
    strCodGasto = vbNullString
    strCodCentro = vbNullString
    
    bolAfecto = False
    bolDescontinuado = False
    bolParaVenta = False
    
    
    strCodigoFab = vbNullString
    strDescripcionFab = vbNullString
    strCodUM = vbNullString
    strModelo = vbNullString
    strCodAlmacen = vbNullString
    strCodMarca = vbNullString
    
    bolEsImportado = False
    
    dblStockMin = 0
    dblStockMax = 0
    dblStockReposicion = 0
    
    dblPorcentajeDemasia = 0
    
    strCtaContableVenta = vbNullString
    strCtaContableInventa = vbNullString
    
    
    strFecIngreso = vbNullString
    strUsuIngreso = vbNullString
    strFecModificacion = vbNullString
    strUsuModificacion = vbNullString
    
    
    strAbrevTipoExistencia = vbNullString
    bolEsInsumoParaOP = False
    strCodSubFamilia = vbNullString
    bolTieneMovimientoEnAlmacen = True
    
    strColor = vbNullString
    strTalla = vbNullString
    strMarcaModelo = vbNullString
    
    strCtaContableVta = vbNullString
    strAnexoVta = vbNullString
    strCtaContableImpVta = vbNullString
    strAnexoImpVta = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerBien() As Boolean
    On Error GoTo errObtenerBien
    
    Set rstBien = New ADODB.Recordset
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "* "
    strSQLBien = strSQLBien & "FROM "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "'"
    
    If rstBien.State = 1 Then rstBien.Close
    
    rstBien.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBien.EOF Then
        strCodigo = Trim(rstBien!f5codpro & "")
        strDescripcion = Trim(rstBien!F5NOMPRO & "")
        strCodTipoExistencia = Trim(rstBien!F5TIPO & "")
        strCodUM = Trim(rstBien!f7codmed & "")
        strCodMarca = Trim(rstBien!F5MARCA & "")
        strCodSubFamilia = Trim(rstBien!F5UBICACIO & "")
        strModelo = Trim(rstBien!F5MODELO & "")
        
        bolAfecto = IIf(Trim(rstBien!F5AFECTO & "") = "*", True, False)
        bolDescontinuado = IIf(Trim(rstBien!f5descontinuado & "") = "S", True, False)
        bolParaVenta = CBool(rstBien!ParaVenta)
        bolEsImportado = IIf(Trim(rstBien!F5ORIGEN & "") = "I", True, False)
        bolEsInsumoParaOP = IIf(Trim(rstBien!F5INSUMO & "") = "S", True, False)
        bolTieneMovimientoEnAlmacen = CBool(rstBien!TIENEMOVENALM)
        
        strCodigoFab = Trim(rstBien!f5codfab & "")
        strDescripcionFab = Trim(rstBien!F5TEXTO & "")
        strCodAlmacen = Trim(rstBien!F2COD_ALM & "")
        
        dblStockMin = Val(rstBien!F5STOCKMIN & "")
        dblStockMax = Val(rstBien!F5STOCKMAX & "")
        dblStockReposicion = Val(rstBien!F5STOCKLOG & "")
        
        dblPorcentajeDemasia = Val(rstBien!PORCDEMASIA & "")
        
        strCtaContable = Trim(rstBien!F5CTACON & "")
        strAnexo = Trim(rstBien!F5ANEXO & "")
        strCtaContableImp = Trim(rstBien!F5CTACONIMP & "")
        strAnexoImp = Trim(rstBien!F5ANEXOIMP & "")
        
        strCodGasto = Trim(rstBien!F3GASTO & "")
        strCodCentro = Trim(rstBien!F5CENTRO & "")
        
        strCtaContableVenta = Trim(rstBien!F5CTACON1 & "")
        strCtaContableInventa = Trim(rstBien!F5CTACON2 & "")
        
        strColor = Trim(rstBien!Color & "")
        strTalla = Trim(rstBien!Talla & "")
        strMarcaModelo = Trim(rstBien!MarcaModelo & "")
        
        strCtaContableVta = Trim(rstBien!F5CTACONVTA & "")
        strAnexoVta = Trim(rstBien!F5ANEXOVTA & "")
        strCtaContableImpVta = Trim(rstBien!F5CTACONIMPVTA & "")
        strAnexoImpVta = Trim(rstBien!F5ANEXOIMPVTA & "")
        
        strFecIngreso = Trim(rstBien!F5FECING & "")
        strUsuIngreso = Trim(rstBien!F5USERING & "")
        strFecModificacion = Trim(rstBien!F5FECMOD & "")
        strUsuModificacion = Trim(rstBien!F5USERMOD & "")
        
        obtenerMedidaAlterna
        
        obtenerBienAlterno
        
        obtenerBien = True
    Else
        obtenerBien = False
    End If
    
    rstBien.Close
    
    Set rstBien = Nothing
    strSQLBien = vbNullString
    
    Exit Function
errObtenerBien:
    obtenerBien = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: ObtenerBien"
    
    Err.Clear
End Function

Public Sub obtenerMedidaAlterna()
    Dim rstBienMA As New ADODB.Recordset
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPMEDVENTA"
    
    abrirCnTemporal
    
    strSQLBien = vbNullString
'    strSQLBien = strSQLBien & "INSERT INTO TMPMEDVENTA("
'        strSQLBien = strSQLBien & "F7CODMED, DESCRIPCION, "
'        strSQLBien = strSQLBien & "FACTOR, F5PREVTA) "
'    strSQLBien = strSQLBien & "IN '" & wrutatemp & "Templus.mdb' "
    strSQLBien = strSQLBien & "SELECT "
        strSQLBien = strSQLBien & "MV.F7CODMED, "
        strSQLBien = strSQLBien & "MED.F7NOMMED, "
        strSQLBien = strSQLBien & "MV.F5FACTOR, "
        strSQLBien = strSQLBien & "MV.F5PREVTA "
    strSQLBien = strSQLBien & "FROM "
        strSQLBien = strSQLBien & "MAESTROS.MEDIVENTAS AS MV "
        strSQLBien = strSQLBien & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = MV.F7CODMED "
    strSQLBien = strSQLBien & "WHERE "
        strSQLBien = strSQLBien & "MV.F5CODPRO = '" & strCodigo & "'"
    
'    cnBdCPlus.Execute strSQLBien

    If rstBienMA.State = 1 Then rstBienMA.Close
    
    rstBienMA.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienMA.EOF Then
        Do While Not rstBienMA.EOF
            strSQLBien = strSQLBien & "INSERT INTO TMPMEDVENTA("
            strSQLBien = strSQLBien & "F7CODMED, DESCRIPCION, "
            strSQLBien = strSQLBien & "FACTOR, F5PREVTA) "
            strSQLBien = strSQLBien & "VALUES("
            strSQLBien = strSQLBien & "'" & Trim(rstBienMA!f7codmed & "") & "', "
            strSQLBien = strSQLBien & "'" & Trim(rstBienMA!F7NOMMED & "") & "', "
            strSQLBien = strSQLBien & Val(rstBienMA!F5FACTOR & "") & ", "
            strSQLBien = strSQLBien & Val(rstBienMA!F5PREVTA & "")
            strSQLBien = strSQLBien & ")"
            
            cnDBTemp.Execute strSQLBien
            
            rstBienMA.MoveNext
        Loop
    End If
    
    If rstBienMA.State = 1 Then rstBienMA.Close
    
    Set rstBienMA = Nothing
    
    strSQLBien = vbNullString
End Sub

Public Sub obtenerBienAlterno()
    Dim rstBienBA As New ADODB.Recordset
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPBIENALTERNO"
    
    abrirCnTemporal
    
    strSQLBien = vbNullString
'    strSQLBien = strSQLBien & "INSERT INTO TMPBIENALTERNO("
'        strSQLBien = strSQLBien & "CODIGO, ESTADO) "
'    strSQLBien = strSQLBien & "IN '" & wrutatemp & "Templus.mdb' "
    strSQLBien = strSQLBien & "SELECT "
        strSQLBien = strSQLBien & "BA.F5CODPROALTERNO, "
        strSQLBien = strSQLBien & "BA.ESTADO "
    strSQLBien = strSQLBien & "FROM "
        strSQLBien = strSQLBien & "EF2BIENALTERNO AS BA "
    strSQLBien = strSQLBien & "WHERE "
        strSQLBien = strSQLBien & "BA.F5CODPRO = '" & strCodigo & "'"
    
'    cnBdCPlus.Execute strSQLBien

    If rstBienBA.State = 1 Then rstBienBA.Close
    
    rstBienBA.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBienBA.EOF Then
        Do While Not rstBienBA.EOF
            strSQLBien = strSQLBien & "INSERT INTO TMPBIENALTERNO("
            strSQLBien = strSQLBien & "CODIGO, ESTADO) "
            strSQLBien = strSQLBien & "VALUES("
            strSQLBien = strSQLBien & "'" & Trim(rstBienBA!F5CODPROALTERNO & "") & "', "
            strSQLBien = strSQLBien & IIf(CBool(rstBienBA!Estado & ""), "TRUE", "FALSE")
            strSQLBien = strSQLBien & ")"
            
            cnDBTemp.Execute strSQLBien
            
            rstBienBA.MoveNext
        Loop
    End If
    
    If rstBienBA.State = 1 Then rstBienBA.Close
    
    Set rstBienBA = Nothing
    
    strSQLBien = vbNullString
End Sub

Public Sub obtenerConfigBien()
    On Error GoTo errObtenerConfigBien
    
    Set rstBien = New ADODB.Recordset
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "* "
    strSQLBien = strSQLBien & "FROM "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "'"
    
    If rstBien.State = 1 Then rstBien.Close
    
    rstBien.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBien.EOF Then
        strCodigo = Trim(rstBien!f5codpro & "")
        strDescripcion = Trim(rstBien!F5NOMPRO & "")
        strCodTipoExistencia = Trim(rstBien!F5TIPO & "")
        strCodUM = Trim(rstBien!f7codmed & "")
        strCodMarca = Trim(rstBien!F5MARCA & "")
        strCodSubFamilia = Trim(rstBien!F5UBICACIO & "")
        strModelo = Trim(rstBien!F5MODELO & "")
        
        bolAfecto = IIf(Trim(rstBien!F5AFECTO & "") = "*", True, False)
        bolDescontinuado = IIf(Trim(rstBien!f5descontinuado & "") = "S", True, False)
        bolParaVenta = CBool(rstBien!ParaVenta)
        bolEsImportado = IIf(Trim(rstBien!F5ORIGEN & "") = "I", True, False)
        bolEsInsumoParaOP = IIf(Trim(rstBien!F5INSUMO & "") = "S", True, False)
        bolTieneMovimientoEnAlmacen = CBool(rstBien!TIENEMOVENALM)
        
        strCodigoFab = Trim(rstBien!f5codfab & "")
        strDescripcionFab = Trim(rstBien!F5TEXTO & "")
        strCodAlmacen = Trim(rstBien!F2COD_ALM & "")
        
        dblStockMin = Val(rstBien!F5STOCKMIN & "")
        dblStockMax = Val(rstBien!F5STOCKMAX & "")
        dblStockReposicion = Val(rstBien!F5STOCKLOG & "")
        
        dblPorcentajeDemasia = Val(rstBien!PORCDEMASIA & "")
        
        strCtaContable = Trim(rstBien!F5CTACON & "")
        strAnexo = Trim(rstBien!F5ANEXO & "")
        strCtaContableImp = Trim(rstBien!F5CTACONIMP & "")
        strAnexoImp = Trim(rstBien!F5ANEXOIMP & "")
        
        strCodGasto = Trim(rstBien!F3GASTO & "")
        strCodCentro = Trim(rstBien!F5CENTRO & "")
        
        strCtaContableVenta = Trim(rstBien!F5CTACON1 & "")
        strCtaContableInventa = Trim(rstBien!F5CTACON2 & "")
        
        strColor = Trim(rstBien!Color & "")
        strTalla = Trim(rstBien!Talla & "")
        strMarcaModelo = Trim(rstBien!MarcaModelo & "")
        
        strCtaContableVta = Trim(rstBien!F5CTACONVTA & "")
        strAnexoVta = Trim(rstBien!F5ANEXOVTA & "")
        strCtaContableImpVta = Trim(rstBien!F5CTACONIMPVTA & "")
        strAnexoImpVta = Trim(rstBien!F5ANEXOIMPVTA & "")
        
        strFecIngreso = Trim(rstBien!F5FECING & "")
        strUsuIngreso = Trim(rstBien!F5USERING & "")
        strFecModificacion = Trim(rstBien!F5FECMOD & "")
        strUsuModificacion = Trim(rstBien!F5USERMOD & "")
    End If
    
    rstBien.Close
    
    Set rstBien = Nothing
    strSQLBien = vbNullString
    
    Exit Sub
errObtenerConfigBien:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: ObtenerConfigBien"
    
    Err.Clear
End Sub

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo

    Set rstBien = New ADODB.Recordset

    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "TOP 1 MID(F5CODPRO, " & Len(strAbrevTipoExistencia) + 1 & ") AS CODIGO "
    strSQLBien = strSQLBien & "FROM "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
    strSQLBien = strSQLBien & "WHERE "
    'strSQLBien = strSQLBien & "ISNUMERIC(F5CODPRO) "
    strSQLBien = strSQLBien & "LEFT(F5CODPRO, " & Len(strAbrevTipoExistencia) & ") = '" & strAbrevTipoExistencia & "' "
    strSQLBien = strSQLBien & "ORDER BY "
    strSQLBien = strSQLBien & "F5CODPRO DESC"

    If rstBien.State = 1 Then rstBien.Close

    rstBien.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly

    If Not rstBien.EOF Then
        generarCodigo = strAbrevTipoExistencia & Format(Val(rstBien!Codigo & "") + 1, "0000000")
    Else
        generarCodigo = strAbrevTipoExistencia & "0000001"
    End If

    rstBien.Close

    Set rstBien = Nothing
    strSQLBien = vbNullString

    Exit Function
errGenerarCodigo:
    generarCodigo = vbNullString

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: GenerarCodigo"

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstBien = New ADODB.Recordset
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA.* "
    strSQLBien = strSQLBien & "FROM "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "'"
    
    If rstBien.State = 1 Then rstBien.Close
    
    rstBien.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBien.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstBien.Close
    
    Set rstBien = Nothing
    
    strSQLBien = vbNullString
    
    Exit Function
errVerificarExistencia:
    verificarExistencia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: VerificarExistencia"
    
    Err.Clear
End Function

Public Function guardarBien(Optional ByVal bolRegistroImportado As Boolean) As Boolean
    On Error GoTo errGuardarBien
    
    strSQLBien = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigo
        End If
        
        strSQLBien = vbNullString
        strSQLBien = strSQLBien & "INSERT INTO MAESTROS.IF5PLA("
        strSQLBien = strSQLBien & "F5CODPRO, F5NOMPRO, F5TIPO, F5CTACON, F5ANEXO, F5CTACONIMP, F5ANEXOIMP, "
        strSQLBien = strSQLBien & "F3GASTO, F5CENTRO, F5AFECTO, F5DESCONTINUADO, "
        strSQLBien = strSQLBien & "PARAVENTA, TIENEMOVENALM, F5CODFAB, F5TEXTO, F7CODMED, "
        strSQLBien = strSQLBien & "F5MODELO, F2COD_ALM, F5MARCA, F5ORIGEN, "
        strSQLBien = strSQLBien & "F5STOCKMIN, F5STOCKMAX, F5STOCKLOG, PORCDEMASIA, F5CTACON1, F5CTACON2, "
        strSQLBien = strSQLBien & "F5INSUMO, F5UBICACIO, COLOR, TALLA, MARCAMODELO, "
        strSQLBien = strSQLBien & "F5CTACONVTA, F5ANEXOVTA, F5CTACONIMPVTA, F5ANEXOIMPVTA, "
        strSQLBien = strSQLBien & "F5FECING, F5USERING) "
        strSQLBien = strSQLBien & "VALUES("
        strSQLBien = strSQLBien & "'" & strCodigo & "', "
        strSQLBien = strSQLBien & "'" & strDescripcion & "', "
        strSQLBien = strSQLBien & "'" & strCodTipoExistencia & "', "
        strSQLBien = strSQLBien & "'" & strCtaContable & "', "
        strSQLBien = strSQLBien & IIf(strAnexo <> vbNullString, "'" & strAnexo & "'", "''") & ", "
        strSQLBien = strSQLBien & IIf(strCtaContableImp <> vbNullString, "'" & strCtaContableImp & "'", "''") & ", "
        strSQLBien = strSQLBien & IIf(strAnexoImp <> vbNullString, "'" & strAnexoImp & "'", "''") & ", "
        strSQLBien = strSQLBien & "'" & strCodGasto & "', "
        strSQLBien = strSQLBien & "'" & strCodCentro & "', "
        strSQLBien = strSQLBien & "'" & IIf(bolAfecto, "*", vbNullString) & "', "
        strSQLBien = strSQLBien & "'" & IIf(bolDescontinuado, "S", "N") & "', "
        strSQLBien = strSQLBien & IIf(bolParaVenta, "1", "0") & ", "
        strSQLBien = strSQLBien & IIf(bolTieneMovimientoEnAlmacen, "1", "0") & ", "
        strSQLBien = strSQLBien & "'" & strCodigoFab & "', "
        strSQLBien = strSQLBien & "'" & strDescripcionFab & "', "
        strSQLBien = strSQLBien & "'" & strCodUM & "', "
        strSQLBien = strSQLBien & "'" & strModelo & "', "
        strSQLBien = strSQLBien & IIf(strCodAlmacen <> vbNullString, "'" & strCodAlmacen & "'", "''") & ", "
        strSQLBien = strSQLBien & "'" & strCodMarca & "', "
        strSQLBien = strSQLBien & "'" & IIf(bolEsImportado, "I", "N") & "', "
        strSQLBien = strSQLBien & dblStockMin & ", "
        strSQLBien = strSQLBien & dblStockMax & ", "
        strSQLBien = strSQLBien & dblStockReposicion & ", "
        strSQLBien = strSQLBien & dblPorcentajeDemasia & ", "
        strSQLBien = strSQLBien & "'" & strCtaContableVenta & "', "
        strSQLBien = strSQLBien & "'" & strCtaContableInventa & "', "
        strSQLBien = strSQLBien & "'" & IIf(bolEsInsumoParaOP, "S", "N") & "', "
        strSQLBien = strSQLBien & "'" & strCodSubFamilia & "', "
        strSQLBien = strSQLBien & "'" & strColor & "', "
        strSQLBien = strSQLBien & "'" & strTalla & "', "
        strSQLBien = strSQLBien & "'" & strMarcaModelo & "', "
        strSQLBien = strSQLBien & "'" & strCtaContableVta & "', "
        strSQLBien = strSQLBien & IIf(strAnexoVta <> vbNullString, "'" & strAnexoVta & "'", "NULL") & ", "
        strSQLBien = strSQLBien & IIf(strCtaContableImpVta <> vbNullString, "'" & strCtaContableImpVta & "'", "NULL") & ", "
        strSQLBien = strSQLBien & IIf(strAnexoImpVta <> vbNullString, "'" & strAnexoImpVta & "'", "NULL") & ", "
        strSQLBien = strSQLBien & "'" & Format(strFecIngreso, "Short Date") & "', "
        strSQLBien = strSQLBien & IIf(strUsuIngreso <> vbNullString, "'" & strUsuIngreso & "'", "''")
        strSQLBien = strSQLBien & ")"
    Else
        strSQLBien = vbNullString
        strSQLBien = strSQLBien & "UPDATE "
            strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
        strSQLBien = strSQLBien & "SET "
            strSQLBien = strSQLBien & "F5NOMPRO = '" & strDescripcion & "', "
            strSQLBien = strSQLBien & "F5TIPO = '" & strCodTipoExistencia & "', "
            strSQLBien = strSQLBien & "F5CTACON = '" & strCtaContable & "', "
            strSQLBien = strSQLBien & "F5ANEXO = " & IIf(strAnexo <> vbNullString, "'" & strAnexo & "'", "''") & ", "
            strSQLBien = strSQLBien & "F5CTACONIMP = " & IIf(strCtaContableImp <> vbNullString, "'" & strCtaContableImp & "'", "''") & ", "
            strSQLBien = strSQLBien & "F5ANEXOIMP = " & IIf(strAnexoImp <> vbNullString, "'" & strAnexoImp & "'", "''") & ", "
            
                If Not bolRegistroImportado Then
                    strSQLBien = strSQLBien & "F3GASTO = '" & strCodGasto & "', "
                    strSQLBien = strSQLBien & "F5CENTRO = '" & strCodCentro & "', "
                    strSQLBien = strSQLBien & "F5AFECTO = '" & IIf(bolAfecto, "*", vbNullString) & "', "
                    strSQLBien = strSQLBien & "PARAVENTA = " & IIf(bolParaVenta, "1", "0") & ", "
                    strSQLBien = strSQLBien & "TIENEMOVENALM = " & IIf(bolTieneMovimientoEnAlmacen, "1", "0") & ", "
                    strSQLBien = strSQLBien & "F5CODFAB = '" & strCodigoFab & "', "
                    strSQLBien = strSQLBien & "F2COD_ALM = " & IIf(strCodAlmacen <> vbNullString, "'" & strCodAlmacen & "'", "''") & ", "
                    strSQLBien = strSQLBien & "PORCDEMASIA = " & dblPorcentajeDemasia & ", "
                    strSQLBien = strSQLBien & "F5CTACON1 = '" & strCtaContableVenta & "', "
                    strSQLBien = strSQLBien & "F5CTACON2 = '" & strCtaContableInventa & "', "
                    
                    strSQLBien = strSQLBien & "F5CTACONVTA = '" & strCtaContableVta & "', "
                    strSQLBien = strSQLBien & "F5ANEXOVTA = " & IIf(strAnexoVta <> vbNullString, "'" & strAnexoVta & "'", "NULL") & ", "
                    strSQLBien = strSQLBien & "F5CTACONIMPVTA = " & IIf(strCtaContableImpVta <> vbNullString, "'" & strCtaContableImpVta & "'", "NULL") & ", "
                    strSQLBien = strSQLBien & "F5ANEXOIMPVTA = " & IIf(strAnexoImpVta <> vbNullString, "'" & strAnexoImpVta & "'", "NULL") & ", "
                End If
            
            strSQLBien = strSQLBien & "F5STOCKMIN = " & dblStockMin & ", "
            strSQLBien = strSQLBien & "F5STOCKMAX = " & dblStockMax & ", "
            strSQLBien = strSQLBien & "F5STOCKLOG = " & dblStockReposicion & ", "
            
            strSQLBien = strSQLBien & "F5DESCONTINUADO = '" & IIf(bolDescontinuado, "S", "N") & "', "
            
            strSQLBien = strSQLBien & "F5TEXTO = '" & strDescripcionFab & "', "
            strSQLBien = strSQLBien & "F7CODMED = '" & strCodUM & "', "
            strSQLBien = strSQLBien & "F5MODELO = '" & strModelo & "', "
            
            strSQLBien = strSQLBien & "F5MARCA = '" & strCodMarca & "', "
            strSQLBien = strSQLBien & "F5ORIGEN = '" & IIf(bolEsImportado, "I", "N") & "', "
            
            strSQLBien = strSQLBien & "F5INSUMO = '" & IIf(bolEsInsumoParaOP, "S", "N") & "', "
            strSQLBien = strSQLBien & "F5UBICACIO = '" & strCodSubFamilia & "', "
            
            strSQLBien = strSQLBien & "COLOR = " & IIf(strColor <> vbNullString, "'" & strColor & "'", "NULL") & ", "
            strSQLBien = strSQLBien & "TALLA = " & IIf(strTalla <> vbNullString, "'" & strTalla & "'", "NULL") & ", "
            strSQLBien = strSQLBien & "MARCAMODELO = " & IIf(strMarcaModelo <> vbNullString, "'" & strMarcaModelo & "'", "NULL") & ", "
            
            strSQLBien = strSQLBien & "F5FECMOD = '" & Format(strFecModificacion, "Short Date") & "', "
            strSQLBien = strSQLBien & "F5USERMOD = '" & strUsuModificacion & "' "
        strSQLBien = strSQLBien & "WHERE "
            strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLBien
    
    guardarBien = True
    
    strSQLSelectAlter = strSQLBien
    strSQLBien = vbNullString
    
    Exit Function
errGuardarBien:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: GuardarBien"
    
    guardarBien = False
    
    strSQLSelectAlter = vbNullString
    strSQLBien = vbNullString
    
    Err.Clear
End Function

Public Sub guardarMedidaAlterna()
    On Error GoTo errGuardarMedidaAlterna
    
    If strSQLSelectAlter = vbNullString Then
        Exit Sub
    End If
    
    Set rstBien = New ADODB.Recordset
    
    If rstBien.State = 1 Then rstBien.Close
    
    rstBien.Open strSQLSelectAlter, cnDBTemp, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstBien.EOF Then
        rstBien.MoveFirst
        
        Do While Not rstBien.EOF
'            If Val(validarUsoRegistro(cnBdCPlus, "F3CODPRO", "IF3ORDEN", strCodigo, "T", "AND UNIDAD = '" & Trim(rstBien!f7codmed & "") & "'")) = 0 Then
'                cnBdCPlus.Execute "DELETE FROM MEDIVENTAS WHERE F5CODPRO = '" & strCodigo & "' AND F7CODMED = '" & Trim(rstBien!f7codmed & "") & "'"
'            End If
            
            If ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F5CODPRO", "MAESTROS.MEDIVENTAS", "F5CODPRO", strCodigo, "T", "AND F7CODMED = '" & Trim(rstBien!f7codmed & "") & "'") = vbNullString Then
                strSQLBien = vbNullString
                strSQLBien = strSQLBien & "INSERT INTO MAESTROS.MEDIVENTAS( "
                strSQLBien = strSQLBien & "F5CODPRO, F7CODMED, "
                strSQLBien = strSQLBien & "F5FACTOR, F5PREVTA) "
                strSQLBien = strSQLBien & "VALUES("
                strSQLBien = strSQLBien & "'" & strCodigo & "', "
                strSQLBien = strSQLBien & "'" & Trim(rstBien!f7codmed & "") & "', "
                strSQLBien = strSQLBien & Val(rstBien!factor & "") & ", " & Val(rstBien!F5PREVTA & "") & ")"
            Else
                strSQLBien = vbNullString
                strSQLBien = strSQLBien & "UPDATE "
                strSQLBien = strSQLBien & "MAESTROS.MEDIVENTAS "
                strSQLBien = strSQLBien & "SET "
                strSQLBien = strSQLBien & "F5FACTOR = " & Val(rstBien!factor & "") & ", "
                strSQLBien = strSQLBien & "F5PREVTA = " & Val(rstBien!F5PREVTA & "") & " "
                strSQLBien = strSQLBien & "WHERE "
                strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "' AND "
                strSQLBien = strSQLBien & "F7CODMED = '" & Trim(rstBien!f7codmed & "") & "'"
            End If
            
            cnBdCPlus.Execute strSQLBien
            
            Actualiza_Log " < Replicación BD > " & strSQLBien, strCadenaConexioBdCPlus
            
            rstBien.MoveNext
        Loop
    Else
        strSQLBien = vbNullString
        strSQLBien = strSQLBien & "INSERT INTO MAESTROS.MEDIVENTAS("
        strSQLBien = strSQLBien & "F5CODPRO, F7CODMED, "
        strSQLBien = strSQLBien & "F5FACTOR, F5PREVTA) "
        strSQLBien = strSQLBien & "VALUES("
        strSQLBien = strSQLBien & "'" & strCodigo & "', "
        strSQLBien = strSQLBien & "'" & strCodUM & "', "
        strSQLBien = strSQLBien & "1, 0)"
        
        cnBdCPlus.Execute strSQLBien
        
        Actualiza_Log " < Replicación BD > " & strSQLBien, strCadenaConexioBdCPlus
    End If
    
    rstBien.Close
    
    Set rstBien = Nothing
    
    Exit Sub
errGuardarMedidaAlterna:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: GuardarMedidaAlterna"
    
    strSQLSelectAlter = vbNullString
    strSQLBien = vbNullString
    
    Err.Clear
End Sub

Public Sub guardarBienAlterno()
    On Error GoTo errGuardarBienAlterno
    
    If strSQLSelectAlter = vbNullString Then
        Exit Sub
    End If
    
    Set rstBien = New ADODB.Recordset
    
    rstBien.Open strSQLSelectAlter, cnDBTemp, adOpenDynamic, adLockBatchOptimistic
    
    If Not rstBien.EOF Then
        cnBdCPlus.Execute "DELETE FROM MAESTROS.EF2BIENALTERNO WHERE F5CODPRO = '" & strCodigo & "'"
        
        rstBien.MoveFirst
        
        Do While Not rstBien.EOF
            If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(*)", "MAESTROS.EF2BIENALTERNO", "F5CODPRO", strCodigo, "T", _
                                                "AND F5CODPROALTERNO = '" & Trim(rstBien!Codigo & "") & "'")) = 0 Then
                
                strSQLBien = vbNullString
                strSQLBien = strSQLBien & "INSERT INTO MAESTROS.EF2BIENALTERNO( "
                strSQLBien = strSQLBien & "F5CODPRO, F5CODPROALTERNO, "
                strSQLBien = strSQLBien & "ESTADO) "
                strSQLBien = strSQLBien & "VALUES("
                strSQLBien = strSQLBien & "'" & strCodigo & "', "
                strSQLBien = strSQLBien & "'" & Trim(rstBien!Codigo & "") & "', "
                strSQLBien = strSQLBien & IIf(CBool(rstBien!Estado), "1", "0") & ")"
                
                cnBdCPlus.Execute strSQLBien
                
                Actualiza_Log " < Replicación BD > " & strSQLBien, strCadenaConexioBdCPlus
            End If
            
            rstBien.MoveNext
        Loop
    End If
    
    rstBien.Close
    
    Set rstBien = Nothing
    
    Exit Sub
errGuardarBienAlterno:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: GuardarBienAlterno"
    
    strSQLSelectAlter = vbNullString
    strSQLBien = vbNullString
    
    Err.Clear
End Sub

Public Function eliminarBien() As Boolean
    On Error GoTo errEliminarBien

    If Val(validarUsoRegistro(cnBdCPlus, "F5CODPRO", "PROCESOS.TBPARAVENTA_DET", strCodigo, "T")) > 0 Or _
       Val(validarUsoRegistro(cnBdCPlus, "F5CODPRO", "PROCESOS.REGISMOV", strCodigo, "T")) > 0 Or _
       Val(validarUsoRegistro(cnBdCPlus, "F5CODPRO", "PROCESOS.IF3VALES", strCodigo, "T")) > 0 Or _
       Val(validarUsoRegistro(cnBdCPlus, "F5CODPRO", "MAESTROS.EF2PROD_CLIE", strCodigo, "T")) > 0 Or _
       Val(validarUsoRegistro(cnBdCPlus, "F5CODPRO", "MAESTROS.EF2PROD_PROV", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarBien = False
        
        Exit Function
    End If
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "DELETE FROM "
    strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "F5CODPRO = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLBien
    
    eliminarBien = True
    
    strSQLSelectAlter = strSQLBien
    strSQLBien = vbNullString
    
    Exit Function
errEliminarBien:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsBien: EliminarBien"
    
    eliminarBien = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'Listar Productos en Grilla (QuamtumGrid)
Public Sub listarGrillaBien(ByVal grilla As dxDBGrid, _
                            ByVal strFiltroSensitivo As String, _
                            ByVal bolAyuda As Boolean, _
                            ByVal bolInsumoOP As Boolean, _
                            ByVal bolParaVenta As Boolean, _
                            ByVal strTipoBienMostrar As String, _
                            Optional bolTienenMovimientoEnAlmacen As Boolean, _
                            Optional strFiltroAdicional As String, _
                            Optional ByVal listaImagenes As Object)
    
    On Error GoTo errListarGrillaBien
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "TE.DESCRIPCION AS TIPO, "
    strSQLBien = strSQLBien & "FAM.F7DESCON AS FAMILIA, "
    strSQLBien = strSQLBien & "SFAM.F7DESCON AS SUBFAMILIA, "
    strSQLBien = strSQLBien & "PROD.F5CODPRO, "
    strSQLBien = strSQLBien & "PROD.F5NOMPRO, "
    strSQLBien = strSQLBien & "MED.F7NOMMED, "
    strSQLBien = strSQLBien & "PROD.F5STOCKMIN, "
    strSQLBien = strSQLBien & "PROD.F5STOCKMAX, "
    strSQLBien = strSQLBien & "PROD.F5STOCKLOG "
    strSQLBien = strSQLBien & "FROM "
    
    strSQLBien = strSQLBien & "(((MAESTROS.IF5PLA AS PROD "
    strSQLBien = strSQLBien & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    strSQLBien = strSQLBien & "LEFT JOIN MAESTROS.EF2TIPOEXISTENCIA AS TE ON TE.CODIGO = PROD.F5TIPO) "
    strSQLBien = strSQLBien & "LEFT JOIN MAESTROS.SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    strSQLBien = strSQLBien & "LEFT JOIN MAESTROS.SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "PROD.F5CODPRO <> '' "
        
        If strFiltroSensitivo <> vbNullString Then
            strSQLBien = strSQLBien & "AND ("
            strSQLBien = strSQLBien & "PROD.F5NOMPRO LIKE '%" & strFiltroSensitivo & "%' OR "
            strSQLBien = strSQLBien & "TE.DESCRIPCION LIKE '%" & strFiltroSensitivo & "%'"
            strSQLBien = strSQLBien & ") "
        End If
        
        If bolAyuda Then
            strSQLBien = strSQLBien & "AND PROD.F5DESCONTINUADO = 'N' "
            
            If bolTienenMovimientoEnAlmacen Then
                strSQLBien = strSQLBien & "AND PROD.TIENEMOVENALM = 1 "
            'Else
            '    strSQLBien = strSQLBien & "AND PROD.TIENEMOVENALM = 0 "
            End If
            
            If bolInsumoOP Then
                strSQLBien = strSQLBien & "AND PROD.F5INSUMO = 'S' "
            End If
            
            If bolParaVenta Then
                strSQLBien = strSQLBien & "AND PROD.PARAVENTA = 1 "
            'Else
                'strSQLBien = strSQLBien & "AND PROD.PARAVENTA = 0 "
            End If
            
            Select Case strTipoBienMostrar
                Case "P"
                    strSQLBien = strSQLBien & "AND PROD.F5TIPO NOT IN ('12') "
                Case "S"
                    strSQLBien = strSQLBien & "AND PROD.F5TIPO IN ('12') "
            End Select
            
            If strFiltroAdicional <> vbNullString Then
                strSQLBien = strSQLBien & strFiltroAdicional
            End If
        Else
            If bolInsumoOP Then
                strSQLBien = strSQLBien & "AND PROD.F5INSUMO = 'S' "
            'Else
            '    strSQLBien = strSQLBien & "AND PROD.F5INSUMO = 'N' "
            End If
            
            If bolParaVenta Then
                strSQLBien = strSQLBien & "AND PROD.PARAVENTA = 1 "
            Else
                strSQLBien = strSQLBien & "AND PROD.PARAVENTA = 0 "
            End If
        End If
    
    strSQLBien = strSQLBien & "ORDER BY "
    strSQLBien = strSQLBien & "FAM.F7DESCON, "
    strSQLBien = strSQLBien & "SFAM.F7DESCON, "
    strSQLBien = strSQLBien & "PROD.F5NOMPRO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
                    
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Tipo de Existencia
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo Existencia"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "TIPO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipoExistencia"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
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
                .Width = 100
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
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F5CODPRO"
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
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
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
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "F7NOMMED"
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Stock Minimo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Stock Min."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKMIN"
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockMinimo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            'Columna Stock Maximo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Stock Max."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "F5STOCKMAX"
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockMaximo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLBien
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODPRO"
            
            .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
        End With
    Else
        strSQLSelectAlter = vbNullString
        strSQLSelectAlter = strSQLBien
    End If
    
    strSQLBien = vbNullString
    
    Exit Sub
errListarGrillaBien:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsBien: ListarGrillaBien"
    End Select
    
    Err.Clear
End Sub

Public Sub actualizarCtasContablesLogistica(ByVal frameProgreso As Object, _
                                                ByVal barraProgreso As Object)
    
    On Error GoTo errActualizarCtasContablesLogistica
    
    Set rstBien = New ADODB.Recordset
    
    strSQLBien = vbNullString
    strSQLBien = strSQLBien & "SELECT "
    strSQLBien = strSQLBien & "CODIGO, "
    strSQLBien = strSQLBien & "CTACONTABLE, "
    strSQLBien = strSQLBien & "ANEXO, "
    strSQLBien = strSQLBien & "CTACONTABLEIMPORTACION, "
    strSQLBien = strSQLBien & "ANEXOIMPORTACION "
    strSQLBien = strSQLBien & "FROM "
    strSQLBien = strSQLBien & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLBien = strSQLBien & "WHERE "
    strSQLBien = strSQLBien & "CODIGOEXTERNO <> ''"
    
    If rstBien.State = 1 Then rstBien.Close
    
    rstBien.Open strSQLBien, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstBien.EOF Then
        frameProgreso.Visible = True
        barraProgreso.Max = ModUtilitario.devuelveCantRegistros(rstBien)
        barraProgreso.Value = 0
        frameProgreso.Caption = "SQL - Actualizando Cuentas Contables..."
        
        Do While Not rstBien.EOF
            strSQLBien = vbNullString
            strSQLBien = strSQLBien & "UPDATE "
            strSQLBien = strSQLBien & "MAESTROS.IF5PLA "
            strSQLBien = strSQLBien & "SET "
            strSQLBien = strSQLBien & "F5CTACON = '" & Trim(rstBien!CtaContable & "") & "', "
            strSQLBien = strSQLBien & "F5ANEXO = '" & Trim(rstBien!Anexo & "") & "', "
            strSQLBien = strSQLBien & "F5CTACONIMP = '" & Trim(rstBien!CtaContableImportacion & "") & "', "
            strSQLBien = strSQLBien & "F5ANEXOIMP = '" & Trim(rstBien!AnexoImportacion & "") & "' "
            strSQLBien = strSQLBien & "WHERE "
            strSQLBien = strSQLBien & "F5TIPO = '" & Trim(rstBien!Codigo & "") & "'"
            
            cnBdCPlus.Execute strSQLBien
            
            Actualiza_Log " < Replicación BD > " & strSQLBien, strCadenaConexioBdCPlus
            
            DoEvents
            
            barraProgreso.Value = barraProgreso.Value + 1
            frameProgreso.Caption = "SQL - Actualizando Cuentas Contables... " & FormatPercent(barraProgreso.Value / barraProgreso.Max, 3)
            
            rstBien.MoveNext
        Loop
    End If
    
    If rstBien.State = 1 Then rstBien.Close
    
    Set rstBien = Nothing
    
    strSQLBien = vbNullString
    
    frameProgreso.Visible = False
    
    Exit Sub
errActualizarCtasContablesLogistica:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsBien: ActualizarCtasContablesLogistica"
    End Select
    
    Err.Clear
End Sub
