VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptListaBien 
   Caption         =   "rptListaBien (ActiveReport)"
   ClientHeight    =   5670
   ClientLeft      =   1815
   ClientTop       =   1830
   ClientWidth     =   12795
   Icon            =   "rptListaBien.dsx":0000
   _ExtentX        =   22569
   _ExtentY        =   10001
   SectionData     =   "rptListaBien.dsx":058A
End
Attribute VB_Name = "rptListaBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_Initialize()
   On Error GoTo errInicia
    
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimir"
        .ITEM(2).Tooltip = "Imprimir"
        .ITEM(4).Tooltip = "Copiar"
        
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
        
'        .Insert 6, "&Formato 2"
'        .Item(6).AddIcon LoadPicture(App.Path & "\Recursos\Excel.ico")
'        .Item(6).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
'        .Item(6).Enabled = True
        
        .Insert 7, "&PDF"
        .ITEM(7).AddIcon LoadPicture(App.Path & "\Acrobat.ico")
        .ITEM(7).Tooltip = "Graba el reporte en un archivo PDF(*.pdf)"
        .ITEM(7).Enabled = True
        
'        .Insert 8, "&TIFF"
'        .Item(8).AddIcon LoadPicture(App.Path & "\Recursos\TIFF.ico")
'        .Item(8).Tooltip = "Graba el reporte en un archivo TIFF(*.tif)"
'        .Item(8).Enabled = True
        
        .ITEM(9).Tooltip = "Buscar"
        .ITEM(10).Tooltip = "Página única"
        .ITEM(11).Tooltip = "Páginas múltiples"
        .ITEM(12).Tooltip = "Zoom (-)"
        .ITEM(13).Tooltip = "Zoom (+)"
        .ITEM(16).Tooltip = "Página previa"
        .ITEM(17).Tooltip = "Página siguiente"
        .ITEM(20).Caption = "&Anterior"
        .ITEM(21).Caption = "&Siguiente"
        .ITEM(20).Tooltip = ""
        .ITEM(21).Tooltip = ""
    End With
    
    
    
    Exit Sub
errInicia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub


Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    On Error GoTo errToolbar
    
    Select Case Tool.Id
        Case 4015 'Exportar a Excel
            Screen.MousePointer = vbHourglass
            
            With cmdlgProducto
                .DialogTitle = "Guardar como"
                .Filter = "Excel (*.xls)|*.xls"
                .CancelError = False
                .ShowSave
                
                If Trim(.FileName) <> vbNullString Then
                    Dim oEXL As ActiveReportsExcelExport.ARExportExcel
                    
                    Set oEXL = New ActiveReportsExcelExport.ARExportExcel
                    
                    oEXL.FileName = Trim(.FileName)
                    oEXL.Export Me.Pages
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case 4016 'Exportar a Excel - Formato 2
            
        Case 4017 'Exportar a PDF
            With cmdlgProducto
                .DialogTitle = "Guardar como"
                .Filter = "Acrobat (*.pdf)|*.pdf"
                .CancelError = False
                .ShowSave
                
                If Trim(.FileName) <> vbNullString Then
                    Dim oPDF As ActiveReportsPDFExport.ARExportPDF
                    
                    Set oPDF = New ActiveReportsPDFExport.ARExportPDF
                    oPDF.FileName = Trim(.FileName)
                    oPDF.Export Me.Pages
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
        Case 4018 'Exportar a TIFF
        
    End Select

    Exit Sub
errToolbar:
    If Err.Number = 70 Then
        MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description & vbNewLine & _
                vbNewLine & _
                "RECUERDA:" & vbNewLine & _
                "Si esta descargando reportes en excel con algún formato predeterminado" & vbNewLine & _
                "por la aplicación; asegurese de guardar su edición y cerrar el archivo" & vbNewLine & _
                "antes de proceder con la descarga", vbInformation + vbOKOnly, App.ProductName
    Else
        MsgBox "No. Error: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    End If
    
    Screen.MousePointer = vbDefault
    
    Err.Clear
End Sub
