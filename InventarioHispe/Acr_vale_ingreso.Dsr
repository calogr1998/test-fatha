VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Acr_vale_ingreso 
   Caption         =   "Logistica_Suma - Acr_vale_ingreso (ActiveReport)"
   ClientHeight    =   7080
   ClientLeft      =   435
   ClientTop       =   3000
   ClientWidth     =   11880
   _ExtentX        =   20955
   _ExtentY        =   12488
   SectionData     =   "Acr_vale_ingreso.dsx":0000
End
Attribute VB_Name = "Acr_vale_ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_PageStart()
pag.Text = Me.pageNumber
End Sub

Private Sub ActiveReport_Initialize()
    
    Me.Toolbar.Tools.Insert 6, "&Excel"
    Me.Toolbar.Tools.ITEM(6).AddIcon LoadPicture(App.Path & "\Excel.ico")
    Me.Toolbar.Tools.ITEM(6).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
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


