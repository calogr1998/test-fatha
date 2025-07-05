VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptcosteo 
   Caption         =   "Proyecto1 - rptcosteo (ActiveReport)"
   ClientHeight    =   9120
   ClientLeft      =   1755
   ClientTop       =   1695
   ClientWidth     =   11445
   WindowState     =   2  'Maximized
   _ExtentX        =   20188
   _ExtentY        =   16087
   SectionData     =   "rptcosteo.dsx":0000
End
Attribute VB_Name = "rptcosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Initialize()
    Me.Toolbar.Tools.ITEM(0).Visible = False
    Me.Toolbar.Tools.ITEM(2).Caption = "&Imprimir"
    Me.Toolbar.Tools.ITEM(4).Visible = False
    'Me.ToolBar.Tools.ITEM(5).Visible = False
    Me.Toolbar.Tools.Insert 5, "&Excel"
    Me.Toolbar.Tools.ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
    Me.Toolbar.Tools.ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
    Me.Toolbar.Tools.ITEM(5).Enabled = True
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

