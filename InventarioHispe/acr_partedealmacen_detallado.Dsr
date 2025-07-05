VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_partedealmacen_detallado 
   Caption         =   "Proyecto1 - acr_partedealmacen_detallado (ActiveReport)"
   ClientHeight    =   9480
   ClientLeft      =   1965
   ClientTop       =   915
   ClientWidth     =   12795
   WindowState     =   2  'Maximized
   _ExtentX        =   22569
   _ExtentY        =   16722
   SectionData     =   "acr_partedealmacen_detallado.dsx":0000
End
Attribute VB_Name = "acr_partedealmacen_detallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GroupFooter1_Format()
lblsub1.Caption = "SUB-TOTAL " & UCase(flddescripcion.Text) & " ==>>"
End Sub

Private Sub GroupFooter2_Format()
lblsub2.Caption = "SUB-TOTAL " & UCase(txtnombre1.Text) & " ==>>"
End Sub

Private Sub GroupHeader1_Format()
If fldtipmov.Text = "I" Then
    flddescripcion.Text = "INGRESOS"
Else
    flddescripcion.Text = "SALIDAS"
End If
End Sub

Private Sub ActiveReport_Initialize()
    
    Me.Toolbar.Tools.Insert 6, "&Excel"
    Me.Toolbar.Tools.item(6).AddIcon LoadPicture(App.Path & "\Excel.ico")
    Me.Toolbar.Tools.item(6).ToolTip = "Graba el reporte en un archivo excel(*.xls)"
    Me.Toolbar.Tools.item(6).Enabled = True

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


