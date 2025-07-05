VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_kardex_MULTIPLE 
   Caption         =   "Proyecto1 - acr_kardex (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "acr_kardex_MULTIPLE.dsx":0000
End
Attribute VB_Name = "acr_kardex_MULTIPLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double

Private Sub ActiveReport_ReportStart()
    
    fldcant_ant.Text = "0.00"
    fldpeso_ant.Text = "0.00"

End Sub

Private Sub Detail_BeforePrint()
    
    ant1 = ant1 + CDbl(ENTRADAK.Text) - CDbl(SALIDAK.Text) + CDbl(wcant)
    fldsaldok.Text = ant1
    fldsaldok.Text = Format(fldsaldok.Text, "#,##0.00")
    wcant = 0#
    
    ant2 = ant2 + CDbl(ENTRADACOS.Text) - CDbl(SALIDACOS.Text) + CDbl(fldpeso_ant.Text)
    fldsaldo.Text = ant2
    fldsaldo.Text = Format(fldsaldo.Text, "#,##0.00")

End Sub

Private Sub GroupFooter1_BeforePrint()
    
    flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")) + Val(Format(fldcant_ant.Text, "0.00")), "###,###,##0")
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
        ENTRADAK_SUB.Text = Format(ENTRADAK_SUB + CDbl(fldcant_ant.Text), "#,##0")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDAK_SUB.Text = Format(SALIDAK_SUB + CDbl(fldcant_ant.Text), "#,##0")
    End If
    
    If fldpeso_ant.Text > 0 Then
        ENTRADACOS_SUB.Text = Format(ENTRADACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDACOS_SUB.Text = Format(SALIDACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
    End If
    
    subtotal1(1) = subtotal1(1) + CDbl(ENTRADAK_SUB.Text)
    subtotal1(2) = subtotal1(2) + CDbl(SALIDAK_SUB.Text)
    subtotal1(3) = subtotal1(3) + CDbl(flddiferk.Text)
    
    subtotal1(4) = subtotal1(4) + CDbl(ENTRADACOS_SUB.Text)
    subtotal1(5) = subtotal1(5) + CDbl(SALIDACOS_SUB.Text)
    subtotal1(6) = subtotal1(6) + CDbl(FLDDIFER.Text)

End Sub

Private Sub GroupHeader1_BeforePrint()

    Saldo_Inicial fldcodprod.Text, kardex.ABODESDE.Text
    fldpeso_ant.Text = Format(Costo_Unitario(fldcodprod.Text, CVDate(kardex.ABODESDE.Text) - 1, IIf(kardex.optmoneda(0).Value = True, "S", "D")), "###,##0.000")
    acr_kardex.fldcant_ant = Format(wcant, "###,###,##0.00")
    acr_kardex.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
    
End Sub

Private Sub GroupHeader1_Format()

    ant1 = 0
    ant2 = 0
    
End Sub

Private Sub ReportFooter_Format()
    
    t1.Text = Format(subtotal1(1), "#,##0")
    t2.Text = Format(subtotal1(2), "#,##0")
    t3.Text = Format(subtotal1(3), "#,##0")
    
    T4.Text = Format(subtotal1(4), "#,##0.00")
    T5.Text = Format(subtotal1(5), "#,##0.00")
    T6.Text = Format(subtotal1(6), "#,##0.00")
    
    If Val(Format(t3.Text, "0.00")) < 0 Then
        t3.ForeColor = &HC0&
    Else
        t3.ForeColor = &H0&
    End If
    
    If Val(Format(T6.Text, "0.00")) < 0 Then
        T6.ForeColor = &HC0&
    Else
        T6.ForeColor = &H0&
    End If

End Sub

Private Sub ActiveReport_Initialize()
    
    Me.Toolbar.Tools.Insert 6, "&Excel"
    Me.Toolbar.Tools.ITEM(6).AddIcon LoadPicture(App.Path & "\Excel.ico")
    Me.Toolbar.Tools.ITEM(6).ToolTip = "Graba el reporte en un archivo excel(*.xls)"
    Me.Toolbar.Tools.ITEM(6).Enabled = True

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

