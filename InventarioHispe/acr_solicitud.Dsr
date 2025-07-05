VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_solicitud 
   Caption         =   "Logistica_Suma - acr_solicitud (ActiveReport)"
   ClientHeight    =   10620
   ClientLeft      =   435
   ClientTop       =   1530
   ClientWidth     =   19455
   WindowState     =   2  'Maximized
   _ExtentX        =   34316
   _ExtentY        =   18733
   SectionData     =   "acr_solicitud.dsx":0000
End
Attribute VB_Name = "acr_solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportEnd()
 If sw_nuevo_doc = True Then
        
        Set oPDF = New ActiveReportsPDFExport.ARExportPDF
        wFileName = "D:\Sol" & Format(solicitud.txtsolicitud.Text, "0000") & ".PDF"
        oPDF.FileName = wFileName
        oPDF.Export Me.Pages
        
        sw_nuevo_doc = False
        Me.Hide
        
    End If
End Sub

Private Sub Detail_BeforePrint()

fldsugerido.Text = ObtenerCampo("EF2proveedores", "f2nomprov", "f2newruc", fldrucprov.Text, "T", cnn_dbbancos)
fldcentro.Text = ObtenerCampo("CENTROS", "F3DESCRIP", "F3COSTO", fldcodcentro.Text, "T", cnn_dbbancos)

End Sub

Private Sub PageHeader_BeforePrint()
    Usuario.Text = wusuario
End Sub

Private Sub PageHeader_Format()

    If Trim(wtiposalida) = "*" Then
        'lblxusuario.Visible = True
        'Lbl1.Visible = True
        'fldxusuario.Visible = True
        'lblxobra.Visible = True
        'Lbl2.Visible = True
        'fldxobra.Visible = True
    Else
'        lblxusuario.Visible = False
'        Lbl1.Visible = False
'        fldxusuario.Visible = False
'        lblxobra.Visible = False
'        Lbl2.Visible = False
'        fldxobra.Visible = False
    End If

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
    
    Select Case Tool.ID
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
