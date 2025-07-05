VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RegImporta 
   Caption         =   "Logistica_Suma - RegImporta (ActiveReport)"
   ClientHeight    =   8085
   ClientLeft      =   210
   ClientTop       =   1770
   ClientWidth     =   12795
   WindowState     =   2  'Maximized
   _ExtentX        =   22569
   _ExtentY        =   14261
   SectionData     =   "RegImporta.dsx":0000
End
Attribute VB_Name = "RegImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wopcion As Byte

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


