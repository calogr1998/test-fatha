VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptValeIngreso 
   Caption         =   "Logistica - rptValeIngreso (ActiveReport)"
   ClientHeight    =   8895
   ClientLeft      =   375
   ClientTop       =   1650
   ClientWidth     =   12570
   Icon            =   "rptValeIngreso.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   22172
   _ExtentY        =   15690
   SectionData     =   "rptValeIngreso.dsx":058A
End
Attribute VB_Name = "rptValeIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strTipoVale As String
Private strCodAlmacen As String
Private strNumeroVale As String

Private strSQL As String

Dim ITEM As Integer
Dim i As Integer

Public Property Let TipoVale(ByVal value As String)
    strTipoVale = value
End Property

Public Property Get TipoVale() As String
    TipoVale = strTipoVale
End Property

Public Property Let CodAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodAlmacen() As String
    CodAlmacen = strCodAlmacen
End Property

Public Property Let NumeroVale(ByVal value As String)
    strNumeroVale = value
End Property

Public Property Get NumeroVale() As String
    NumeroVale = strNumeroVale
End Property

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

Private Sub ActiveReport_ReportStart()
    imprimirVale
End Sub

Private Sub imprimirVale()
    strSQL = vbNullString
    
    
        With objAyudaVale
            .TipoVale = strTipoVale
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            strSQL = .devuelveSQLResumenVale
            
            Me.lblRegistrado.Caption = "Registrado: " & IIf(.FecReg = vbNullString, "--", .FecReg) & " (" & IIf(.UsuReg = vbNullString, "--", .UsuReg) & ")"
            Me.lblModificado.Caption = "Modificado: " & IIf(.FecMod = vbNullString, "--", .FecMod) & " (" & IIf(.UsuMod = vbNullString, "--", .UsuMod) & ")"
        End With
        
        dtcVale.ConnectionString = StrConexDbBancos
    
    With dtcVale
        '.ConnectionString = cnn_dbbancos
        .CursorLocation = ddADOUseClient
        .CursorType = ddADOOpenStatic
        .LockType = ddADOLockReadOnly
        .Source = strSQL
    End With
    
    Me.Caption = App.ProductName & " - Vale de " & IIf(strTipoVale = "I", "Ingreso", "Salida")
    Me.lblSistema.Caption = App.LegalTrademarks & "-" & App.ProductName
    Me.lblEmpresa.Caption = wnomcia
    Me.lblTitulo.Caption = "Vale de " & IIf(strTipoVale = "I", "Ingreso", "Salida")
    
    
    If objAyudaVale.CodigoOrigen <> "XC0" Then
        lblRegistroCompra.Visible = False
        fldRegistroCompra.Visible = False
        
        lblDescripcion.left = lblRequerimiento.left
        lblDescripcion.Width = lblRequerimiento.Width + lblDescripcion.Width
        
        fldDescripcion.left = fldRequerimiento.left
        fldDescripcion.Width = fldRequerimiento.Width + fldDescripcion.Width
        
        lblRequerimiento.left = lblOC.left
        lblRequerimiento.Width = lblOC.Width
        
        fldRequerimiento.left = fldOrdenCompra.left
        fldRequerimiento.Width = fldOrdenCompra.Width
        
        lblOC.Visible = False
        fldOrdenCompra.Visible = False
    End If
End Sub
                                    
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    On Error GoTo errToolbar
    
    Select Case Tool.ID
        Case 4015 'Exportar a Excel
            Screen.MousePointer = vbHourglass
            
            With cmdlgVale
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
'        Case 4016 'Exportar a Excel - Formato 2
            
        Case 4016 'Exportar a PDF
            With cmdlgVale
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

Private Sub Detail_Format()
    ITEM = ITEM + 1
    
    fldItem.Text = ITEM
    
    If (ITEM Mod 2) = 0 Then
        Detail.BackColor = &HFFFFFF
    Else
        Detail.BackColor = &H80000016
    End If
End Sub
