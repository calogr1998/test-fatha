VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_Mov_TipoOrigen_detallado 
   Caption         =   "Reporte de Inventario por Origen"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11880
   Icon            =   "acr_Mov_TipoOrigen_detallado.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   14975
   SectionData     =   "acr_Mov_TipoOrigen_detallado.dsx":000C
End
Attribute VB_Name = "acr_Mov_TipoOrigen_detallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totalsoles As Double
Dim totaldolares As Double

Private Sub ActiveReport_Activate()
Screen.MousePointer = 0
End Sub

Private Sub ActiveReport_PageStart()

    fldpagina.Text = Me.pageNumber
    
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
        .ITEM(9).Tooltip = "P�gina �nica"
        .ITEM(10).Tooltip = "P�ginas m�ltiples"
        .ITEM(12).Tooltip = "Zoom (-)"
        .ITEM(13).Tooltip = "Zoom (+)"
        .ITEM(16).Tooltip = "P�gina previa"
        .ITEM(17).Tooltip = "P�gina siguiente"
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
            MsgBox "Exportaci�n terminada", vbInformation, "Reporte"
        End If
    End If

End Sub

Private Sub Detail_BeforePrint()
'Me.Field14.Text = Format(Val(Format(Me.Field12.Text, "0.00")) * Val(Format(Me.Field20.Text, "0.00")), "###,##0.00")'
'Me.Field15.Text = Format(Val(Format(Me.Field12.Text, "0.00")) * Val(Format(Me.Field22.Text, "0.00")), "###,##0.00"')'
'totalsoles = totalsoles + Val(Format(Me.Field14.Text, "0.00"))
'totaldolares = totaldolares + Val(Format(Me.Field15.Text, "0.00"))
End Sub

Private Sub ReportFooter_BeforePrint()
    'Me.Field13.Text = Format(totaldolares, "###,##0.00")
    'Me.Field11.Text = Format(totalsoles, "###,##0.00")
End Sub

