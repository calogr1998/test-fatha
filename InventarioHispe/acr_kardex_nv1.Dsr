VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_kardex_nv 
   Caption         =   "Kardex no valorizado"
   ClientHeight    =   8484
   ClientLeft      =   -36
   ClientTop       =   2208
   ClientWidth     =   11892
   WindowState     =   2  'Maximized
   _ExtentX        =   20976
   _ExtentY        =   14965
   SectionData     =   "acr_kardex_nv1.dsx":0000
End
Attribute VB_Name = "acr_kardex_nv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double

Dim NumVale As String
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    
NumVale = Link
 IMPRIMIR_VALES '''(rpt, Link, strSQL)
End Sub
Private Sub IMPRIMIR_VALES()
Dim csql            As String
Dim CSQL1           As String
Dim Csql2           As String
Dim Csql3           As String
Dim prov            As String
Dim costo           As String
Dim RegEmpresa      As New ADODB.Recordset
Dim RegCosto        As New ADODB.Recordset
Dim RegMoneda       As New ADODB.Recordset
Dim ccod_almacen    As String
Dim cnum_vale       As String
Dim ctipo_vale      As String * 1
Dim WMONEDA As String * 1
    Me.MousePointer = 11
    ccod_almacen = Right(kardex.TxtAlmacen.Text, 2)
    cnum_vale = Trim(NumVale)
    csql = "SELECT CENTROS.F3COSTO, CENTROS.F3DESCRIP, IF4VALES.F4MONEDA, IF4VALES.F4NUMVAL"
    csql = csql + " FROM CENTROS INNER JOIN IF4VALES ON CENTROS.F3COSTO = IF4VALES.F4CENTRO"
    csql = csql + " WHERE  F4NUMVAL='" & cnum_vale & "'"
    RegMoneda.Open csql, cnn_dbbancos, adOpenStatic, adLockReadOnly
        'costo = Trim(txtccosto.Text)
        If Not (RegMoneda.EOF Or RegMoneda.Bof) Then
            WMONEDA = Trim(RegMoneda!F4Moneda)
            Costo1 = Trim(RegMoneda!F3DESCRIP)
        Else
            WMONEDA = s
            Costo1 = ""
        End If
    RegMoneda.Close
        If WMONEDA = "D" Then
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALDOL AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valdol] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE (((IF4VALES.F2CODALM)='" & TxtAlmacen.Text & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "') AND ((IF4VALES.F1CODORI)='" & txtconcepto.Text & "')) " & _
                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        Else
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALVTA AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valvta] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE (((IF4VALES.F2CODALM)='" & ccod_almacen & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "'))" & _
                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        End If
    
        With acr_vales
        .DataControl1.ConnectionString = cnn_dbbancos
        If Left(cnum_vale, 1) = "I" Then
            'prov = Trim(txtproveedor.Text)
            ctipo_vale = "I"
            .Lbl_vale.Caption = " VALE DE INGRESO "
            .lblprov.Visible = True
            .lblpunto.Visible = True
            .fldprov.Visible = True
            .lblpie2.Caption = "Entregado por"
        Else
            ctipo_vale = "S"
            .Lbl_vale.Caption = " VALE DE SALIDA "
            .lblprov.Visible = False
            .lblpunto.Visible = False
            .fldprov.Visible = False
            .lblpie2.Caption = "Hecho por"
        End If
        
        .DataControl1.Source = csql
        '.fldnomprov.Text = pnlproveedor.Caption
        .fldalmacen.Text = Mid(kardex.PnlAlmacen.Caption, 200)
        '.fldnomcosto.Text = Costo1
        .fldempresa.Text = wnomcia
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .fldvale.Text = cnum_vale
        .fldalma.Text = ccod_almacen
       
        .Show vbModal
    End With
Me.MousePointer = 1
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

    Saldo_Inicial ccodprod, kardex.abodesde.Text, wcod_alm
'    fldpeso_ant.Text = Format(Costo_Unitario(fldcodprod.Text, CVDate(kardex.abodesde.Text) - 1, IIf(kardex.optmoneda(0).Value = True, "S", "D")), "###,##0.000")
    acr_kardex_nv.fldcant_ant = Format(wcant, "###,###,##0.00")
    acr_kardex_nv.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
    acr_kardex_nv.fldsaldoant = Format(wcant, "###,###,##0.00")
    acr_kardex_nv.fldsaldoant.OutputFormat = "#,##0.00;(#,##0.00)"
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
        .ITEM(4).Tooltip = "Copiar"
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
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
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim oEXL As ActiveReportsExcelExport.ARExportExcel

    If Tool.ID = 4015 Then
        Load frmCommon
        strFilePath = frmCommon.Ruta
        Unload frmCommon
        If Trim(strFilePath) <> "" Then
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = strFilePath
            oEXL.Export Me.Pages
            MsgBox "Exportación terminada", vbInformation, "Reporte"
        End If
    End If

End Sub

