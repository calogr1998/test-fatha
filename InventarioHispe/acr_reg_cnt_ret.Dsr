VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_reg_cnt_ret 
   Caption         =   "Proyecto1 - acr_reg_cnt_ret (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "acr_reg_cnt_ret.dsx":0000
End
Attribute VB_Name = "acr_reg_cnt_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nsaldo      As Double
Dim sw_i        As Boolean
Dim nretenido   As Double
Dim nbase       As Double
Dim ncompra     As Double

Private Sub ActiveReport_Initialize()

    sw_i = False
    nsaldo = 0#
    nretenido = 0#:  nbase = 0#: ncompra = 0#

End Sub

Private Sub Detail_Format()

    If sw_i = False Then
        sw_i = True
        nsaldo = Val(Format(DEBE.Text, "0.00")) - Val(Format(HABER.Text, "0.00"))
    Else
        nsaldo = nsaldo + (Val(Format(DEBE.Text, "0.00")) - Val(Format(HABER.Text, "0.00")))
    End If
    SALDO.Text = Format(IIf(nsaldo < 0, nsaldo * -1, nsaldo), "###,###,##0.00")
    
    If Val(Format(DEBE.Text, "0.00")) = 0# Then
        DEBE.Text = ""
    End If
    If Val(Format(HABER.Text, "0.00")) = 0# Then
        HABER.Text = ""
    End If
    
    If TIPO.Text = "1" Then
        ncompra = ncompra + Val(Format(DEBE.Text, "0.00")) + Val(Format(HABER.Text, "0.00"))
    End If
    If TIPO.Text = "2" Then
        nbase = nbase + Val(Format(DEBE.Text, "0.00")) + Val(Format(HABER.Text, "0.00"))
    End If
    If TIPO.Text = "3" Then
        nretenido = nretenido + Val(Format(DEBE.Text, "0.00")) + Val(Format(HABER.Text, "0.00"))
    End If
    
End Sub

Private Sub GroupFooter1_Format()
    
    If wf1agente = "*" Then
        lbltit.Caption = "Total  Compra / Base / Retenido"
    Else
        lbltit.Caption = "Total  Venta  / Base / Retenido"
    End If
    
    fldretenido.Text = Format(nretenido, "###,###,##0.00")
    fldbase.Text = Format(nbase, "###,###,##0.00")
    fldcompra.Text = Format(ncompra, "###,###,##0.00")
    
End Sub

Private Sub PageHeader_Format()

    If wf1agente = "*" Then
        lbltipo.Caption = "Proveedor : "
    Else
        lbltipo.Caption = "Cliente   : "
    End If

End Sub
