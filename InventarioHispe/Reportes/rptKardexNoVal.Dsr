VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptKardexNoVal 
   Caption         =   "Logistica_Suma - rptKardexNoVal (ActiveReport)"
   ClientHeight    =   8490
   ClientLeft      =   480
   ClientTop       =   1785
   ClientWidth     =   13725
   Icon            =   "rptKardexNoVal.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   24209
   _ExtentY        =   14975
   SectionData     =   "rptKardexNoVal.dsx":058A
End
Attribute VB_Name = "rptKardexNoVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double

Dim NumVale As String

Private strCodigoAlmacen            As String
Private strCodigoProducto           As String
Private StrFecha                    As String

Private strSQL                      As String

Dim ITEM                            As Integer
Dim i                               As Integer

Public Property Let CodigoAlmacen(ByVal value As String)
    strCodigoAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodigoAlmacen
End Property

Public Property Let CodigoProducto(ByVal value As String)
    strCodigoProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodigoProducto
End Property

Public Property Let Fecha(ByVal value As String)
    StrFecha = value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property


Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    'NumVale = Link
    
    'IMPRIMIR_VALES '''(rpt, Link, strSQL)
    
    Me.MousePointer = vbHourglass
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = left(Trim(Link), 1)
        .CodigoAlmacen = strCodigoAlmacen
        .NumeroVale = Trim(Link)
        
        .obtenerConfigVale
    End With
    
    With rptValeIngreso
        .TipoVale = objAyudaVale.TipoVale
        .CodAlmacen = objAyudaVale.CodigoAlmacen
        .NumeroVale = objAyudaVale.NumeroVale
        
        'ModMilano.abrirCnDBMilano
        
        .fldCategoria.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CT.NOMBRE", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", Val(objAyudaVale.OrdenTrabajo), "N")
        .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Val(objAyudaVale.OrdenTrabajo), "N")
        
        .Show 1
    End With
    
    Me.MousePointer = vbDefault
End Sub
Private Sub IMPRIMIR_VALES()
'    Dim csql            As String
'    Dim CSQL1           As String
'    Dim Csql2           As String
'    Dim Csql3           As String
'    Dim prov            As String
'    Dim costo           As String
'    Dim RegEmpresa      As New ADODB.Recordset
'    Dim RegCosto        As New ADODB.Recordset
'    Dim RegMoneda       As New ADODB.Recordset
'    Dim ccod_almacen    As String
'    Dim cnum_vale       As String
'    Dim ctipo_vale      As String * 1
'    Dim WMONEDA As String * 1
'
'    Me.MousePointer = vbHourglass
'    ccod_almacen = right(kardex.txtalmacen.Text, 2)
'    cnum_vale = Trim(NumVale)
'    csql = "SELECT CENTROS.F3COSTO, CENTROS.F3DESCRIP, IF4VALES.F4MONEDA, IF4VALES.F4NUMVAL"
'    csql = csql + " FROM CENTROS INNER JOIN IF4VALES ON CENTROS.F3COSTO = IF4VALES.F4CENTRO"
'    csql = csql + " WHERE  F4NUMVAL='" & cnum_vale & "'"
'    RegMoneda.Open csql, cnn_dbbancos, adOpenStatic, adLockReadOnly
'        'costo = Trim(txtccosto.Text)
'        If Not (RegMoneda.EOF Or RegMoneda.Bof) Then
'            WMONEDA = Trim(RegMoneda!F4MONEDA)
'            Costo1 = Trim(RegMoneda!F3DESCRIP)
'        Else
'            WMONEDA = s
'            Costo1 = ""
'        End If
'    RegMoneda.Close
'        If WMONEDA = "D" Then
'            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALDOL AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valdol] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
'                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
'                " WHERE (((IF4VALES.F2CODALM)='" & txtalmacen.Text & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "') AND ((IF4VALES.F1CODORI)='" & txtconcepto.Text & "')) " & _
'                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
'        Else
'            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALVTA AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valvta] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
'                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
'                " WHERE (((IF4VALES.F2CODALM)='" & ccod_almacen & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "'))" & _
'                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
'        End If
'
'        With acr_vales
'        .DataControl1.ConnectionString = cnn_dbbancos
'        If left(cnum_vale, 1) = "I" Then
'            'prov = Trim(txtproveedor.Text)
'            ctipo_vale = "I"
'            .Lbl_vale.Caption = " VALE DE INGRESO "
'            .lblprov.Visible = True
'            .lblpunto.Visible = True
'            .fldprov.Visible = True
'            .lblpie2.Caption = "Entregado por"
'        Else
'            ctipo_vale = "S"
'            .Lbl_vale.Caption = " VALE DE SALIDA "
'            .lblprov.Visible = False
'            .lblpunto.Visible = False
'            .fldprov.Visible = False
'            .lblpie2.Caption = "Hecho por"
'        End If
'
'        .DataControl1.Source = csql
'        '.fldnomprov.Text = pnlproveedor.Caption
'        .fldAlmacen.Text = Mid(kardex.pnlalmacen.Caption, 200)
'        '.fldnomcosto.Text = Costo1
'        .fldempresa.Text = wnomcia
'        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
'        .fldvale.Text = cnum_vale
'        .fldalma.Text = ccod_almacen
'
'        .Show vbModal
'    End With
'    Me.MousePointer = vbDefault
End Sub


Private Sub ActiveReport_ReportStart()
    
    fldcant_ant.Text = "0.00"
   ' fldpeso_ant.Text = "0.00"

End Sub

Private Sub Detail_BeforePrint()
    ant1 = ant1 + CDbl(Val(Format(ENTRADAK.Text & "", "0.00"))) - CDbl(Val(Format(SALIDAK.Text & "", "0.00"))) + CDbl(wcant)
    fldsaldok.Text = ant1
    fldsaldok.Text = Format(fldsaldok.Text, "#,##0.00")
    wcant = 0#
    
fldnumvale.Hyperlink = fldnumvale.Text
End Sub

Private Sub GroupFooter1_BeforePrint()
    
    flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")) + Val(Format(fldcant_ant.Text, "0.00")), "###,###,##0.00")
   ' FLDDIFER.Text = Format(Val(ant2), "###,###,##0.00")
    
    If Val(Format(flddiferk.Text, "0.00")) < 0 Then
        flddiferk.ForeColor = &HC0&
    Else
        flddiferk.ForeColor = &H0&
    End If
    
'    If Val(Format(FLDDIFER.Text, "0.00")) < 0 Then
'        FLDDIFER.ForeColor = &HC0&
'    Else
'        FLDDIFER.ForeColor = &H0&
'    End If
    
    If fldcant_ant.Text > 0 Then
        ENTRADAK_SUB.Text = Format(ENTRADAK_SUB + CDbl(fldcant_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDAK_SUB.Text = Format(SALIDAK_SUB + CDbl(fldcant_ant.Text), "#,##0.00")
    End If
    
'    If fldpeso_ant.Text > 0 Then
'        ENTRADACOS_SUB.Text = Format(ENTRADACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
'    ElseIf fldcant_ant.Text < 0 Then
'        SALIDACOS_SUB.Text = Format(SALIDACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
'    End If
    
    subtotal1(1) = subtotal1(1) + CDbl(ENTRADAK_SUB.Text)
    subtotal1(2) = subtotal1(2) + CDbl(SALIDAK_SUB.Text)
    subtotal1(3) = subtotal1(3) + CDbl(flddiferk.Text)
    
'    subtotal1(4) = subtotal1(4) + CDbl(ENTRADACOS_SUB.Text)
'    subtotal1(5) = subtotal1(5) + CDbl(SALIDACOS_SUB.Text)
'    subtotal1(6) = subtotal1(6) + CDbl(FLDDIFER.Text)

End Sub

Private Sub GroupHeader1_BeforePrint()

    'Saldo_Inicial ccodprod, kardex.aboDesde.value, wcod_alm
    With objSqlAyudaVale
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .CodigoAlmacen = strCodigoAlmacen
        .CodigoProducto = strCodigoProducto
        .Fecha = StrFecha
        .CodigoMoneda = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l")
        
        .obtenerSaldoYCostoInicialDeProducto
        
        wcant = .Cantidad
    End With
    
    
    
'    fldpeso_ant.Text = Format(Costo_Unitario(fldcodprod.Text, CVDate(kardex.abodesde.Text) - 1, IIf(kardex.optmoneda(0).Value = True, "S", "D")), "###,##0.000")
'    rptKardexNoVal.fldcant_ant = Format(wcant, "###,###,##0.00")
'    rptKardexNoVal.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
'    rptKardexNoVal.fldsaldoant = Format(wcant, "###,###,##0.00")
'    rptKardexNoVal.fldsaldoant.OutputFormat = "#,##0.00;(#,##0.00)"
    
    fldcant_ant = Format(wcant, "#,0.00")
    fldcant_ant.OutputFormat = "#,0.00;(#,0.00)"
    fldsaldoant = Format(wcant, "#,0.00")
    fldsaldoant.OutputFormat = "#,0.00;(#,0.00)"
End Sub

Private Sub GroupHeader1_Format()

    ant1 = 0
    ant2 = 0
    
End Sub

Private Sub ReportFooter_Format()
    
    t1.Text = Format(subtotal1(1), "#,##0.00")
    t2.Text = Format(subtotal1(2), "#,##0.00")
    t3.Text = Format(subtotal1(3), "#,##0.00")
    
'    T4.Text = Format(subtotal1(4), "#,##0.00")
    'T5.Text = Format(subtotal1(5), "#,##0.00")
'    T6.Text = Format(subtotal1(6), "#,##0.00")
    
    If Val(Format(t3.Text, "0.00")) < 0 Then
        t3.ForeColor = &HC0&
    Else
        t3.ForeColor = &H0&
    End If
    
'    If Val(Format(T6.Text, "0.00")) < 0 Then
'        T6.ForeColor = &HC0&
'    Else
'        T6.ForeColor = &H0&
'    End If

End Sub

Private Sub ActiveReport_Initialize()
    
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimir"
        .ITEM(2).Tooltip = "Imprimir"
        .ITEM(4).Visible = False
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
        .Insert 6, "&Acrobat"
        .ITEM(6).AddIcon LoadPicture(App.Path & "\Acrobat.ico")
        .ITEM(6).Tooltip = "Exporta el reporte a un archivo *.pdf"
        .ITEM(6).Enabled = True
        .ITEM(7).Tooltip = "Buscar"
        .ITEM(9).Tooltip = "Página única"
        .ITEM(10).Tooltip = "Páginas múltiples"
        .ITEM(12).Tooltip = "Zoom (-)"
        .ITEM(13).Tooltip = "Zoom (+)"
        .ITEM(16).Tooltip = "Página previa"
        .ITEM(17).Tooltip = "Página siguiente"
        .ITEM(20).Caption = "&Anterior"
        .ITEM(21).Caption = "&Siguiente"
        .ITEM(20).Tooltip = ""
        .ITEM(21).Tooltip = ""
        
    End With
    
'lblempresa.Caption = wnomcia

End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim oEXL As ActiveReportsExcelExport.ARExportExcel
    
    Select Case Tool.Id
    Case 4015
        RutaReporte.TipoFile = 1
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        
        If Trim(strFilePath) <> "" Then
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = strFilePath
            oEXL.Export Me.Pages
            MsgBox "Exportación terminada, " & strFilePath, vbInformation, wnomcia
        End If
    Case 4016

        RutaReporte.TipoFile = 0
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        
        If Trim(strFilePath) <> "" Then
            Set oPDF = New ActiveReportsPDFExport.ARExportPDF
            oPDF.FileName = strFilePath
            oPDF.Export Me.Pages
            MsgBox "Exportación terminada, " & strFilePath, vbInformation, wnomcia
        End If
    End Select

End Sub

