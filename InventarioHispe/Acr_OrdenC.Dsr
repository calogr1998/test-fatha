VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Acr_OrdenC_Otros 
   Caption         =   "Logistica_Suma - Acr_OrdenC_Otros (ActiveReport)"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   1395
   ClientWidth     =   15420
   WindowState     =   2  'Maximized
   _ExtentX        =   27199
   _ExtentY        =   19288
   SectionData     =   "Acr_OrdenC.dsx":0000
End
Attribute VB_Name = "Acr_OrdenC_Otros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub ActiveReport_Initialize()
    i = 0
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimir"
        .ITEM(2).Tooltip = "Imprimir"
        .ITEM(4).Visible = False
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
        .Insert 6, "&Correo"
        .ITEM(6).AddIcon LoadPicture(App.Path & "\contactl.ico")
        .ITEM(6).Tooltip = "Envia el documento por correo"
        .ITEM(6).Enabled = True
        .Insert 7, "&Word"
        .ITEM(7).AddIcon LoadPicture(App.Path & "\doc.ico")
        .ITEM(7).Tooltip = "Graba el reporte en un archivo word(*.doc)"
        .ITEM(7).Enabled = True
        .ITEM(9).Tooltip = "Buscar"
        .ITEM(11).Tooltip = "Página única"
        .ITEM(12).Tooltip = "Páginas múltiples"
        .ITEM(14).Tooltip = "Zoom (-)"
        .ITEM(15).Tooltip = "Zoom (+)"
        .ITEM(18).Tooltip = "Página previa"
        .ITEM(19).Tooltip = "Página siguiente"
        .ITEM(22).Caption = "&Anterior"
        .ITEM(23).Caption = "&Siguiente"
        .ITEM(22).Tooltip = ""
        .ITEM(23).Tooltip = ""
    End With

End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
On Error Resume Next
Dim oEXL As ActiveReportsExcelExport.ARExportExcel
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
'Dim oRTF As ActiveReportsRTFExport.ARExportRTF
Select Case Tool.Id
    Case 4015:
        Ind = "0"
        Load frmCommon
        strFilePath = frmCommon.Ruta
        Unload frmCommon
        If Trim(strFilePath) <> "" Then
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = strFilePath
            oEXL.Export Me.Pages
            MsgBox "Exportación terminada", vbInformation, "Reporte"
        End If
        
    Case 4017:
'        Ind = "1"
'        Load frmCommon
'        strFilePath = frmCommon.Ruta
'        Unload frmCommon
'        If Trim(strFilePath) <> "" Then
'            Set oRTF = New ActiveReportsRTFExport.ARExportRTF
'            oRTF.FileName = strFilePath
'            oRTF.Export Me.Pages
'
'        End If
    
    Case 4016:
        strFilePathPDF = wrutatemp & "\OC_" & Mid(LblTitle.Caption, 22, 10) & "_" & Mid(LblTitle.Caption, 33, Len(LblTitle.Caption) - 32) & ".PDF"
        Set oPDF = New ActiveReportsPDFExport.ARExportPDF
        oPDF.FileName = strFilePathPDF
        oPDF.Export Me.Pages
        wasunto = LblTitle.Caption
        'Load correo
        'correo.Show 1
End Select
End Sub



Private Sub Detail_BeforePrint()
L1.Y2 = 0
L1.Y1 = f5nompro.Height
L2.Y2 = 0
L2.Y1 = f5nompro.Height
L3.Y2 = 0
L3.Y1 = f5nompro.Height
L4.Y2 = 0
L4.Y1 = f5nompro.Height
L5.Y2 = 0
L5.Y1 = f5nompro.Height
L6.Y2 = 0
L6.Y1 = f5nompro.Height
L7.Y2 = 0
L7.Y1 = f5nompro.Height
End Sub


Private Sub Detail_Format()
i = i + 1
ITEM.Text = i
End Sub

Private Sub GroupHeader1_Format()
LG1.Y2 = 0
LG1.Y1 = FLDCLI.Height
LG2.Y2 = 0
LG2.Y1 = FLDCLI.Height
LG3.Y2 = FLDCLI.Height
LG3.Y1 = FLDCLI.Height
End Sub
