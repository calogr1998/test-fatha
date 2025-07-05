VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_inventario_novalorizado 
   Caption         =   "Proyecto1 - acr_inventario_novalorizado (ActiveReport)"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   390
   ClientWidth     =   15420
   WindowState     =   2  'Maximized
   _ExtentX        =   27199
   _ExtentY        =   19315
   SectionData     =   "acr_inventario_novalorizado.dsx":0000
End
Attribute VB_Name = "acr_inventario_novalorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_PageStart()

    fldpagina.Text = Me.pageNumber
    
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
    
lblempresa.Caption = wnomcia

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

Private Sub GroupHeader3_Format()

End Sub

