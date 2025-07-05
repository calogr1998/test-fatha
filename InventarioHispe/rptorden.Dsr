VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptorden 
   Caption         =   "Proyecto1 - rptorden (ActiveReport)"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11880
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   14975
   SectionData     =   "rptorden.dsx":0000
End
Attribute VB_Name = "rptorden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Initialize()
    Me.Toolbar.Tools.Insert 6, "&Excel"
    Me.Toolbar.Tools.item(6).AddIcon LoadPicture(App.Path & "\Excel.ico")
    Me.Toolbar.Tools.item(6).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
    Me.Toolbar.Tools.item(6).Enabled = True
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

Private Sub PageHeader_Format()

End Sub
