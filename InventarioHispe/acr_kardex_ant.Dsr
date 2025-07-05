VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_kardex 
   Caption         =   "Proyecto1 - acr_kardex (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "acr_kardex.dsx":0000
End
Attribute VB_Name = "acr_kardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double
Dim cont As Long


Private Sub ActiveReport_ReportStart()
    
    'inicio = True
    fldcant_ant.Text = "0.00"
    fldpeso_ant.Text = "0.00"
    
End Sub

Private Sub Detail_BeforePrint()
    
    ant1 = ant1 + CDbl(ENTRADAK.Text) - CDbl(SALIDAK.Text) + CDbl(fldcant_ant.Text)
    fldsaldok.Text = ant1
    fldsaldok.Text = Format$(fldsaldok.Text, "#,###0.00")

    ant2 = ant2 + CDbl(ENTRADACOS.Text) - CDbl(SALIDACOS.Text) + CDbl(fldpeso_ant.Text)
    fldsaldo.Text = ant2
    fldsaldo.Text = Format$(fldsaldo.Text, "#,###0.00")

End Sub

Private Sub GroupFooter1_BeforePrint()
    
    flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")) + Val(Format(fldcant_ant.Text, "0.00")), "###,###,##0.00")
    FLDDIFER.Text = Format(Val(Format(ENTRADACOS_SUB.Text, "0.00")) - Val(Format(SALIDACOS_SUB.Text, "0.00")) + Val(Format(fldpeso_ant.Text, "0.00")), "###,###,##0.00")
    
    If Val(Format(flddiferk.Text, "0.00")) < 0 Then
        flddiferk.ForeColor = &HC0&
    End If
    
    If Val(Format(FLDDIFER.Text, "0.00")) < 0 Then
        FLDDIFER.ForeColor = &HC0&
    End If
    
    If fldcant_ant > 0 Then
        ENTRADAK_SUB.Text = Format$(ENTRADAK_SUB + CDbl(fldcant_ant), "#,###0.00")
    ElseIf fldcant_ant < 0 Then
        SALIDAK_SUB.Text = Format$(SALIDAK_SUB + Abs(CDbl(fldcant_ant)), "#,###0.00")
    End If
    
    If fldpeso_ant > 0 Then
        ENTRADACOS_SUB.Text = Format$(ENTRADACOS_SUB + CDbl(fldpeso_ant), "#,###0.00")
    ElseIf fldcant_ant < 0 Then
            SALIDACOS_SUB.Text = Format$(SALIDACOS_SUB + Abs(CDbl(fldpeso_ant)), "#,###0.00")
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
    cont = cont + 1
    
    acr_kardex.fldcant_ant = Format(wcant, "###,###,##0.00")
    acr_kardex.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
    
End Sub

Private Sub ReportFooter_Format()
    
    t1.Text = Format$(subtotal1(1), "#,###0.00")
    t2.Text = Format$(subtotal1(2), "#,###0.00")
    t3.Text = Format$(subtotal1(3), "#,###0.00")
    
    T4.Text = Format$(subtotal1(4), "#,###0.00")
    T5.Text = Format$(subtotal1(5), "#,###0.00")
    T6.Text = Format$(subtotal1(6), "#,###0.00")
    
End Sub
