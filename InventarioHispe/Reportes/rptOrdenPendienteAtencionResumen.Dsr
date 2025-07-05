VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptOrdenPendienteAtencionResumen 
   Caption         =   "Logistica - rptOrdenPendienteAtencionResumen (ActiveReport)"
   ClientHeight    =   5745
   ClientLeft      =   75
   ClientTop       =   7755
   ClientWidth     =   14145
   Icon            =   "rptOrdenPendienteAtencionResumen.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   24950
   _ExtentY        =   10134
   SectionData     =   "rptOrdenPendienteAtencionResumen.dsx":058A
End
Attribute VB_Name = "rptOrdenPendienteAtencionResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCodProveedor As String
Private strCodProducto As String
Private bolMostrarNroRequerimiento As Boolean
Private bolSoloProductoConSaldo As Boolean
Private bolIncluirOCAtencionTotalYOrdenCerrada As Boolean
Private strDesde As String
Private strHasta As String


Private strSQL As String

Dim ITEM As Integer
Dim i As Integer

Public Property Let CodigoProveedor(ByVal value As String)
    strCodProveedor = value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProveedor
End Property

Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Property Let MostrarNroRequerimiento(ByVal value As Boolean)
    bolMostrarNroRequerimiento = value
End Property

Public Property Get MostrarNroRequerimiento() As Boolean
    MostrarNroRequerimiento = bolMostrarNroRequerimiento
End Property

Public Property Let SoloProductoConSaldo(ByVal value As Boolean)
    bolSoloProductoConSaldo = value
End Property

Public Property Get SoloProductoConSaldo() As Boolean
    SoloProductoConSaldo = bolSoloProductoConSaldo
End Property

Public Property Let IncluirOCAtencionTotalYOrdenCerrada(ByVal value As Boolean)
    bolIncluirOCAtencionTotalYOrdenCerrada = value
End Property

Public Property Get IncluirOCAtencionTotalYOrdenCerrada() As Boolean
    IncluirOCAtencionTotalYOrdenCerrada = bolIncluirOCAtencionTotalYOrdenCerrada
End Property

Public Property Let Desde(ByVal value As String)
    strDesde = value
End Property

Public Property Get Desde() As String
    Desde = strDesde
End Property

Public Property Let Hasta(ByVal value As String)
    strHasta = value
End Property

Public Property Get Hasta() As String
    Hasta = strHasta
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
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        imprimirOrdenMovimientoSql
    Else
        imprimirOrdenMovimiento
    End If
    
    Me.Caption = App.ProductName & " - Resumen de Atención de O/C's por Proveedor: " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TRIM(F2NEWRUC & ' - ' & F2NOMPROV) AS PROVEEDOR", "EF2PROVEEDORES", "F2CODPROV", strCodProveedor, "T")
    Me.lblSistema.Caption = App.LegalTrademarks & "-" & App.ProductName
    Me.lblEmpresa.Caption = wnomcia
End Sub

Private Sub imprimirOrdenMovimiento()
    strSQL = vbNullString
    
    With objAyudaOrden
        .CodProveedor = strCodProveedor
        .CodigoProducto = strCodProducto
        
        'strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorResumen
        
        strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorV2(True, bolSoloProductoConSaldo)
    End With
    
    With dtcVale
        .ConnectionString = cnn_dbbancos
        .CursorLocation = ddADOUseClient
        .CursorType = ddADOOpenStatic
        .LockType = ddADOLockReadOnly
        .Source = strSQL
    End With
End Sub

Private Sub imprimirOrdenMovimientoSql()
    strSQL = vbNullString
    
    With objSqlAyudaOrden
        .CodProveedor = strCodProveedor
        .CodigoProducto = strCodProducto
        
        'strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorResumen
        
        strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorV2(True, bolSoloProductoConSaldo, False, bolIncluirOCAtencionTotalYOrdenCerrada, strDesde, strHasta)
    End With
    
    With dtcVale
        .ConnectionString = strCadenaConexioBdCPlus
        .CursorLocation = ddADOUseClient
        .CursorType = ddADOOpenStatic
        .LockType = ddADOLockReadOnly
        .Source = strSQL
    End With
End Sub
                                    
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    On Error GoTo errToolbar
    
    Select Case Tool.Id
        Case 4015 'Exportar a Excel
            Screen.MousePointer = vbHourglass
            
            With cmdlgOrden
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
            With cmdlgOrden
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
    
    If Val(Format(fldCantidad.Text, "#.00")) >= 0 Then
        fldCantidad.ForeColor = vbBlue
    Else
        fldCantidad.ForeColor = vbRed
    End If
End Sub

Private Sub GroupFooter1_AfterPrint()
    'fldPorEntregar.Text = Val(Format(fldCantidadRequerida.Text, "#.00")) - Val(Format(fldCantidadEntregada.Text, "#.00"))
    ITEM = 0
End Sub

