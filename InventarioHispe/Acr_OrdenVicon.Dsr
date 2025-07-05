VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Acr_OrdenVicon 
   Caption         =   "Proyecto1 - Acr_OrdenVicon (ActiveReport)"
   ClientHeight    =   11160
   ClientLeft      =   345
   ClientTop       =   2235
   ClientWidth     =   19200
   WindowState     =   2  'Maximized
   _ExtentX        =   33867
   _ExtentY        =   19685
   SectionData     =   "Acr_OrdenVicon.dsx":0000
End
Attribute VB_Name = "Acr_OrdenVicon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Private Sub ActiveReport_Initialize()
    I = 0
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
        .ITEM(6).Tooltip = "Graba el reporte en un archivo Acrobat(*.pdf)"
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
Dim oRTF As ActiveReportsRTFExport.ARExportRTF
Dim ABC As New Word.Application
Select Case Tool.Id
    Case 4015:
        
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
        
    Case 4017:
        RutaReporte.TipoFile = 3
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        If Trim(strFilePath) <> "" Then
            Set oRTF = New ActiveReportsRTFExport.ARExportRTF
            oRTF.FileName = strFilePath
            oRTF.Export Me.Pages
            ABC.Visible = True
            ABC.Documents.Open strFilePath
        End If
    
    Case 4016:
        
        RutaReporte.TipoFile = 0
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        
        If Trim(strFilePath) <> "" Then
            'strFilePathPDF = wrutatemp & "\OC_" & cname & ".PDF"
            Set oPDF = New ActiveReportsPDFExport.ARExportPDF
            oPDF.FileName = strFilePath
            oPDF.Export Me.Pages
            MsgBox "Exportación terminada, " & strFilePath, vbInformation, wnomcia
        End If
        
        
        'Load correo
        'correo.Show 1
End Select
End Sub



Private Sub Detail_BeforePrint()

F3CANPRO.Height = F5NOMPRO.Height
F3MEDIDA.Height = F5NOMPRO.Height
F3PREUNI.Height = F5NOMPRO.Height
F5VALVTA1.Height = F5NOMPRO.Height
FldValVta.Height = F5NOMPRO.Height
FldIgv.Height = F5NOMPRO.Height
FldItem.Height = F5NOMPRO.Height
LblSep01.Height = F5NOMPRO.Height
End Sub

Private Sub Detail_Format()
I = I + 1
FldItem.Text = I
FldObserva.Visible = False
F5NOMPRO.Text = F5NOMPRO.Text & IIf(Len(Trim(FldObserva.Text)) > 0, vbCrLf & "(*) " & FldObserva.Text, "")
End Sub

Private Sub GroupFooter1_BeforePrint()
FldSon.top = FldObservaAll.Height + FldObservaAll.top
LblSon1.top = FldObservaAll.Height + FldObservaAll.top
LblSon2.top = FldObservaAll.Height + FldObservaAll.top

End Sub

Private Sub GroupFooter1_Format()
LblBorderDet.Height = FldSon.top + FldSon.Height + 400
GroupFooter1.Height = FldSon.top + FldSon.Height + 400
End Sub

Private Sub GroupHeader1_Format()
'LG1.Y2 = 0
'LG1.Y1 = FLDCLI.Height
'LG2.Y2 = 0
'LG2.Y1 = FLDCLI.Height
'LG3.Y2 = FLDCLI.Height
'LG3.Y1 = FLDCLI.Height
End Sub
