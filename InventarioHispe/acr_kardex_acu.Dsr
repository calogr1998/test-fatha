VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_kardex_acu 
   Caption         =   "Proyecto1 - acr_kardex_acu (ActiveReport)"
   ClientHeight    =   8490
   ClientLeft      =   -30
   ClientTop       =   2205
   ClientWidth     =   11880
   Icon            =   "acr_kardex_acu.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   14975
   SectionData     =   "acr_kardex_acu.dsx":000C
End
Attribute VB_Name = "acr_kardex_acu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double

Dim NumVale As String

Private Sub Detail_AfterPrint()
Field2.Text = Format((Val(Format(fldcant_ant.Text, "0.00")) + Val(Format(fldmedida.Text, "0.00"))) - Val(Format(Field1.Text, "0.00")), "###,##0.00")
End Sub

Private Sub Detail_BeforePrint()
Field2.Text = Format((Val(Format(fldcant_ant.Text, "0.00")) + Val(Format(fldmedida.Text, "0.00"))) - Val(Format(Field1.Text, "0.00")), "###,##0.00")
End Sub

Private Sub ReportFooter_AfterPrint()
Field6.Text = (Val(Format(Field3.Text, "0.00")) + Val(Format(Field4.Text, "0.00"))) - Val(Format(Field5.Text, "0.00"))
End Sub

Private Sub ReportFooter_BeforePrint()
Field6.Text = (Val(Format(Field3.Text, "0.00")) + Val(Format(Field4.Text, "0.00"))) - Val(Format(Field5.Text, "0.00"))
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
