VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_inventario_valorizado_n 
   Caption         =   "Proyecto1 - acr_inventario_valorizado_n (ActiveReport)"
   ClientHeight    =   7560
   ClientLeft      =   1560
   ClientTop       =   4245
   ClientWidth     =   12375
   WindowState     =   2  'Maximized
   _ExtentX        =   21828
   _ExtentY        =   13335
   SectionData     =   "acr_inventario_valorizado_n.dsx":0000
End
Attribute VB_Name = "acr_inventario_valorizado_n"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsM As New ADODB.Recordset
Private Sub ActiveReport_PageStart()

    fldpagina.Text = Me.pageNumber
    
End Sub

Private Sub ActiveReport_Initialize()
    csql = "select * from ef2marcas"
    If RsM.State = 1 Then RsM.Close
    RsM.Open csql, cnn_dbbancos, 3, 1
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimir"
        .ITEM(2).Tooltip = "Imprimir"
        .ITEM(4).Tooltip = "Copiar"
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
        .ITEM(6).Tooltip = "Buscar"
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

Private Sub Detail_Format()

'fldsaldo.Text = Format(Val(Left(fldsaldo.Text, InStr(1, fldsaldo.Text, ".") + 5)), "#,#0.0000")
'fldunitario.Text = Format(Val(Left(fldunitario.Text, InStr(1, fldunitario.Text, ".") + 5)), "#,#0.0000")
'fldtotal.Text = Format(Val(Left(fldtotal.Text, InStr(1, fldtotal.Text, ".") + 5)), "#,#0.0000")

RsM.Filter = ""
RsM.Filter = "f2codmar='" & FldMarca.Text & "'"
If RsM.RecordCount > 0 Then
    FldMarca.Text = RsM!f2desmar
    Me.Refresh
Else
    FldMarca.Text = ""
End If
End Sub

Private Sub GroupHeader1_Format()
RsM.Filter = ""
RsM.Filter = "f2codmar='" & fldquiebre.Text & "'"
If RsM.RecordCount > 0 Then
    fldquiebre.Text = RsM!f2desmar
    Me.Refresh
Else
    'fldquiebre.Text = ""
End If
End Sub
