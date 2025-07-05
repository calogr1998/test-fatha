VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form ReporteChildTrue 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   795
   ClientTop       =   1320
   ClientWidth     =   13545
   Icon            =   "ReporteChildTrue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   13545
   WindowState     =   2  'Maximized
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arvPreview 
      Height          =   10515
      Left            =   540
      TabIndex        =   0
      Top             =   420
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   18547
      SectionData     =   "ReporteChildTrue.frx":058A
   End
End
Attribute VB_Name = "ReporteChildTrue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub arvPreview_ToolbarClick(ByVal Tool As DDActiveReportsViewer2Ctl.DDTool)
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
Dim oXLS As ActiveReportsExcelExport.ARExportExcel
Select Case Tool.Id
    Case -25519
        RutaReporte.TipoFile = 1
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        
        If Trim(strFilePath) <> "" Then
            Set oXLS = New ActiveReportsExcelExport.ARExportExcel
            oXLS.FileName = strFilePath
            oXLS.Export Me.arvPreview.Pages
            MsgBox "Exportación terminada, " & strFilePath, vbInformation, wNomCia
        End If
    Case -25520
        RutaReporte.TipoFile = 0
        Load RutaReporte
        strFilePath = RutaReporte.Ruta
        Unload RutaReporte
        
        If Trim(strFilePath) <> "" Then
            Set oPDF = New ActiveReportsPDFExport.ARExportPDF
            oPDF.FileName = strFilePath
            oPDF.Export Me.arvPreview.Pages
            MsgBox "Exportación terminada, " & strFilePath, vbInformation, wNomCia
        End If
    
        
    Case -25517
        Unload Me
End Select
End Sub

Private Sub Form_Load()

arvPreview.Zoom = -1

With arvPreview.Toolbar.Tools
    .Item(0).Visible = False
    .Item(2).Caption = "&Imprimir"
    .Item(2).ToolTip = "Imprimir"
    .Item(4).ToolTip = "Copiar"
    .Item(6).ToolTip = "Buscar"
    .Item(8).ToolTip = "Vista Simple"
    .Item(9).ToolTip = "Vista Múltiple"
    .Item(11).ToolTip = "Zoom (-)"
    .Item(12).ToolTip = "Zoom (+)"
    .Item(15).ToolTip = "Página Anterior"
    .Item(16).ToolTip = "Página Siguiente"
    .Item(19).ToolTip = "Ir Atrás"
    .Item(19).Caption = "Atrás"
    .Item(20).Caption = "Adelante"
    .Item(20).ToolTip = "Ir Adelante"
    .Insert 21, ""
    .Item(21).Type = 2
    .Insert 22, "&Acrobat"
    .Item(22).AddIcon LoadPicture(App.Path & "\acrobat.ico")
    .Item(22).ToolTip = "Exportar a *.pdf"
    .Item(22).Enabled = True
    .Insert 23, "&Excel"
    .Item(23).AddIcon LoadPicture(App.Path & "\Excel.ico")
    .Item(23).ToolTip = "Exportar a *.xls"
    .Item(23).Enabled = True
    .Insert 24, ""
    .Item(24).Type = 2
    
    .Insert 25, "&Salir"
    .Item(25).AddIcon LoadPicture(App.Path & "\Exit.ico")
    .Item(25).ToolTip = "Salir"
    .Item(25).Enabled = True
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
arvPreview.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
