VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_MovDiario_xAlmacen 
   Caption         =   "Proyecto1 - acr_MovDiario_xAlmacen (ActiveReport)"
   ClientHeight    =   5895
   ClientLeft      =   -4845
   ClientTop       =   1800
   ClientWidth     =   12000
   WindowState     =   2  'Maximized
   _ExtentX        =   21167
   _ExtentY        =   10398
   SectionData     =   "acr_MovDiario_xAlmacen2.dsx":0000
End
Attribute VB_Name = "acr_MovDiario_xAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double
Dim CONT As Long
Dim OpcI As String
Dim OpcS As String
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    Dim strSQL      As String
    Dim rpt         As Object
    Dim prProd As New Recordset
    Dim rs1 As New Recordset
    Set rpt = New acr_MovDiario_V_I
        If prProd.State = 1 Then prProd.Close
        If Link = "I" Or Link = "S" Then
            strSQL = "SELECT DISTINCTROW * FROM IF4VALES WHERE IF4VALES.F4FECVAL = " & MovDiario_xAlmacen.TxtFecha.Value & ""
        Else
            strSQL = "SELECT DISTINCTROW MOV_DIARIO_DIA.F5CODPRO FROM MOV_DIARIO_DIA WHERE INGRESO = " & Format(Link, "0.00") & ""
        End If
        prProd.Open strSQL, cnn_form, adOpenStatic, adLockReadOnly
        If Not (prProd.EOF Or prProd.Bof) Then
        strSQL = ""
        Link = prProd!F5CODPRO
        
        If rs1.State = adStateOpen Then rs1.Close
        rs1.Open "DROP TABLE MOV_DIARIO_VALES", cnn_form, adOpenDynamic, adLockOptimistic
        If rs1.State = adStateOpen Then rs1.Close
        SQL = "SELECT IF4VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB,IF3VALES.F3CANPRO, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC,IF4VALES.F4SERGUIA,IF4VALES.F4NUMGUIA INTO MOV_DIARIO_VALES IN '" & wrutatemp & "\TEMPLUS.MDB' "
        SQL = SQL + " FROM  (IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO"
        SQL = SQL + " WHERE (((IF4VALES.F2CODALM)='" & MovDiario_xAlmacen.Txtcodalm.Text & "') and (Left([IF4VALES]![F4NUMVAL],1)='I') AND ((IF4VALES.F4FECVAL)=CVDate('" & Format(MovDiario_xAlmacen.TxtFecha.Text, "DD/MM/YYYY") & "')))"
        If ctipoadm_bd = "M" Then
            cnn_form.Execute SQL
        Else
            cnn_dbbancos.Execute SQL
        End If
            fecha = MovDiario_xAlmacen.TxtFecha.Value
            strSQL = "Select * from MOV_DIARIO_VALES where f5Codpro = '" & Link & "'"
            
            Call ShowReport(rpt, Link, strSQL)
        End If
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

Private Sub ActiveReport_ReportStart()
sw_Report = False
'datconexion.ConnectionString = cconex_dbbancos
    'inicio = True
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


Private Sub Detail_BeforePrint()
    
'    ant1 = ant1 + CDbl(ENTRADAK.Text) - CDbl(SALIDAK.Text) + CDbl(wcant)
'    fldsaldok.Text = ant1
'    fldsaldok.Text = Format(fldsaldok.Text, "#,###0.00")
'    wcant = 0#
'
'    ant2 = ant2 + CDbl(ENTRADACOS.Text) - CDbl(SALIDACOS.Text) + CDbl(fldpeso_ant.Text)
'    fldsaldo.Text = ant2
'    fldsaldo.Text = Format(fldsaldo.Text, "#,###0.00")
OpcI = "I"
INGRESOS.Hyperlink = INGRESOS.Text
fldcodprod.Hyperlink = fldcodprod.Text
OpcS = "S"
SALIDAS.Hyperlink = SALIDAS.Text

LblIngreso.Hyperlink = Left(LblIngreso.Caption, 1)
LblSalida.Hyperlink = Left(LblSalida.Caption, 1)
End Sub

Private Sub GroupFooter1_BeforePrint()
    
    flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")) + Val(Format(fldcant_ant.Text, "0.00")), "###,###,##0.00")
    'FLDDIFER.Text = Format(Val(Format(ENTRADACOS_SUB.Text, "0.00")) - Val(Format(SALIDACOS_SUB.Text, "0.00")) + Val(Format(fldpeso_ant.Text, "0.00")), "###,###,##0.00")
    FLDDIFER.Text = Format(Val(ant2), "###,###,##0.00")
    
    If Val(Format(flddiferk.Text, "0.00")) < 0 Then
        flddiferk.ForeColor = &HC0&
    Else
        flddiferk.ForeColor = &H0&
    End If
    
    If Val(Format(FLDDIFER.Text, "0.00")) < 0 Then
        FLDDIFER.ForeColor = &HC0&
    Else
        FLDDIFER.ForeColor = &H0&
    End If
    
    If fldcant_ant.Text > 0 Then
        ENTRADAK_SUB.Text = Format(ENTRADAK_SUB + CDbl(fldcant_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDAK_SUB.Text = Format(SALIDAK_SUB + Abs(CDbl(fldcant_ant.Text)), "#,##0.00")
    End If
    
    If fldpeso_ant.Text > 0 Then
        ENTRADACOS_SUB.Text = Format(ENTRADACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDACOS_SUB.Text = Format(SALIDACOS_SUB + Abs(CDbl(fldpeso_ant.Text)), "#,##0.00")
    End If
    
    subtotal1(1) = subtotal1(1) + CDbl(ENTRADAK_SUB.Text)
    subtotal1(2) = subtotal1(2) + CDbl(SALIDAK_SUB.Text)
    subtotal1(3) = subtotal1(3) + CDbl(flddiferk.Text)
    
    subtotal1(4) = subtotal1(4) + CDbl(ENTRADACOS_SUB.Text)
    subtotal1(5) = subtotal1(5) + CDbl(SALIDACOS_SUB.Text)
    subtotal1(6) = subtotal1(6) + CDbl(FLDDIFER.Text)

End Sub

Private Sub GroupHeader1_Format()

    ant1 = 0
    ant2 = 0
    Saldo_Inicial fldcodprod, kardex.abodesde.Text
    CONT = CONT + 1
    acr_kardex.fldcant_ant = Format(wcant, "###,###,##0.00")
    acr_kardex.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
    
End Sub


Private Sub ReportFooter_Format()
    
     't1.Text = Format(subtotal1(1), "#,###0.00")
'    t2.Text = Format(subtotal1(2), "#,###0.00")
'    t3.Text = Format(subtotal1(3), "#,###0.00")
'
'    T4.Text = Format(subtotal1(4), "#,###0.00")
'    T5.Text = Format(subtotal1(5), "#,###0.00")
'    T6.Text = Format(subtotal1(6), "#,###0.00")

    'POR INVESTIGAR ANIBAL
''    If Val(Format(t3.Text, "0.00")) < 0 Then
''        t3.ForeColor = &HC0&
''    Else
''        t3.ForeColor = &H0&
''    End If
''
'    If Val(Format(T6.Text, "0.00")) < 0 Then
'        T6.ForeColor = &HC0&
'    Else
'        T6.ForeColor = &H0&
'    End If

End Sub
