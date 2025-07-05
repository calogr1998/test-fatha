VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_movoc 
   Caption         =   "Pago de Ordenes"
   ClientHeight    =   10470
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   18468
   SectionData     =   "acr_movoc.dsx":0000
End
Attribute VB_Name = "acr_movoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim subtotal1 As Double
Dim NumVale As String
Private Sub ActiveReport_ReportStart()
    
'    fldcant_ant.Text = "0.00"

End Sub

Private Sub Detail_BeforePrint()
    'ant1 = ant1 + CDbl(Val(Format(ENTRADAK.Text & "", "0.00"))) - CDbl(Val(Format(SALIDAK.Text & "", "0.00"))) + CDbl(wcant)
    'fldsaldok.Text = ant1
    'fldsaldok.Text = Format(fldsaldok.Text, "#,##0.00")
    'wcant = 0#
    
'fldnumvale.Hyperlink = fldnumvale.Text
End Sub

Private Sub GroupFooter1_BeforePrint()

  ' flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")), "###,###,##0.00")
  '  If Val(Format(flddiferk.Text, "0.00")) < 0 Then
   '     flddiferk.ForeColor = &HC0&
   ' Else
   '     flddiferk.ForeColor = &H0&
   ' End If
    'If fldcant_ant.Text > 0 Then
     '  ENTRADAK_SUB.Text = Format(ENTRADAK_SUB, "#,##0.00")
     '   ENTRADAK_SUB.Text = Format(ENTRADAK_SUB + CDbl(fldcant_ant.Text), "#,##0.00")
    'ElseIf fldcant_ant.Text < 0 Then
    SALIDAK_SUB.ForeColor = &H0&
        SALIDAK_SUB.Text = Format(SALIDAK_SUB * -1, "#,##0.00")
    'End If
    'subtotal1(1) = subtotal1(1) + CDbl(ENTRADAK_SUB.Text)
    flddiferk.Text = 100 + Format(CDbl(Field13.Text), "#,##0.00")
    subtotal1 = subtotal1 + CDbl(SALIDAK_SUB.Text)
    'subtotal1(2) = subtotal1(2) + CDbl(flddiferk.Text)
    Field12.Text = Format(CDbl(Field11.Text) - CDbl(SALIDAK_SUB.Text), "#,##0.00")
End Sub

Private Sub GroupHeader1_BeforePrint()

  '  Saldo_Inicial ccodprod, kardex.aboDesde.Text, wcod_alm
  '  acr_kardex_nv.fldcant_ant = Format(wcant, "###,###,##0.00")
  '  acr_kardex_nv.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
  '  acr_kardex_nv.fldsaldoant = Format(wcant, "###,###,##0.00")
  '  acr_kardex_nv.fldsaldoant.OutputFormat = "#,##0.00;(#,##0.00)"
End Sub

Private Sub ReportFooter_Format()
    
  '  t2.Text = Format(subtotal1, "#,##0.00")
   ' t3.Text = Format(subtotal1(2), "#,##0.00")
    
    'If Val(Format(t3.Text, "0.00")) < 0 Then
     '   t3.ForeColor = &HC0&
    'Else
     '   t3.ForeColor = &H0&
    'End If
    

End Sub

Private Sub ActiveReport_Initialize()
    
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimire"
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

    If Tool.Id = 4015 Then
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

