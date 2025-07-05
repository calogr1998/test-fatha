VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_partedealmacen_general 
   Caption         =   "Proyecto1 - acr_partedealmacen_general (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "acr_partedealmacen_general.dsx":0000
End
Attribute VB_Name = "acr_partedealmacen_general"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GroupHeader1_Format()

    If fldtipmov.Text = "I" Then
        flddescripcion.Text = "Ingresos"
    Else
        flddescripcion.Text = "Salidas"
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


