VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTomaInventarioFormato 
   Caption         =   "Logistica - rptTomaInventarioFormato (ActiveReport)"
   ClientHeight    =   5355
   ClientLeft      =   105
   ClientTop       =   1785
   ClientWidth     =   14085
   Icon            =   "rptTomaInventarioFormato.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   24844
   _ExtentY        =   9446
   SectionData     =   "rptTomaInventarioFormato.dsx":058A
End
Attribute VB_Name = "rptTomaInventarioFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intGrupoOpcion As Integer
Private strGrupoCadena As String
Private strFiltroSensitivo As String
Private strGrupoFiltro As String

Private strSQL As String

Dim ITEM As Integer
Dim i As Integer

Public Property Let GrupoOpcion(ByVal value As Integer)
    intGrupoOpcion = value
End Property

Public Property Get GrupoOpcion() As Integer
    GrupoOpcion = intGrupoOpcion
End Property

Public Property Let GrupoCadena(ByVal value As String)
    strGrupoCadena = value
End Property

Public Property Get GrupoCadena() As String
    GrupoCadena = strGrupoCadena
End Property

Public Property Let FiltroSensitivo(ByVal value As String)
    strFiltroSensitivo = value
End Property

Public Property Get FiltroSensitivo() As String
    FiltroSensitivo = strFiltroSensitivo
End Property

Public Property Let GrupoFiltro(ByVal value As String)
    strGrupoFiltro = value
End Property

Public Property Get GrupoFiltro() As String
    GrupoFiltro = strGrupoFiltro
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
        imprimirTomaInventarioSql
    Else
        'imprimirTomaInventario
    End If
    
    Me.Caption = App.ProductName & " - Formato de Toma de Inventario: Por " & strGrupoCadena
    lblGrupo.Caption = UCase(strGrupoCadena)
    Me.lblSistema.Caption = App.LegalTrademarks & "-" & App.ProductName
    Me.lblEmpresa.Caption = wnomcia
End Sub

Private Sub imprimirTomaInventario()
'    strSQL = vbNullString
'
'    With objAyudaOrden
'        .CodProveedor = strCodProveedor
'
'        'strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorResumen
'
'        strSQL = .devuelveSQLOrdenPendienteAtencionPorProveedorV2(True, bolSoloProductoConSaldo)
'    End With
'
'    With dtcVale
'        .ConnectionString = cnn_dbbancos
'        .CursorLocation = ddADOUseClient
'        .CursorType = ddADOOpenStatic
'        .LockType = ddADOLockReadOnly
'        .Source = strSQL
'    End With
End Sub

Private Sub imprimirTomaInventarioSql()
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
        
        Select Case intGrupoOpcion
            Case 0
                strSQL = strSQL & "FAMILIA"
            Case 1
                strSQL = strSQL & "SUBFAMILIA"
        End Select
        
        strSQL = strSQL & " AS GRUPO, "
        strSQL = strSQL & "NOMPRODUCTO AS DESCRIPCION, "
        strSQL = strSQL & "UM, "
        strSQL = strSQL & "STOCKSISTEMA AS SISTEMA, "
        strSQL = strSQL & "(CASE WHEN STOCKFISICO = 0 THEN NULL ELSE STOCKFISICO END) AS FISICO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "tmpCPTomaInventario" & wusuario & " "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "RTRIM(LTRIM(CODPRODUCTO)) <> '' "
        
        If strFiltroSensitivo <> vbNullString Then
            strSQL = strSQL & "AND NOMPRODUCTO LIKE '%" & strFiltroSensitivo & "%' "
        End If
        
        If strGrupoFiltro <> vbNullString And strGrupoFiltro <> "(*) Todos" Then
            Select Case intGrupoOpcion
                Case 0
                    strSQL = strSQL & "AND FAMILIA = '" & strGrupoFiltro & "' "
                Case 1
                    strSQL = strSQL & "AND SUBFAMILIA = '" & strGrupoFiltro & "' "
            End Select
        End If
        
    strSQL = strSQL & "ORDER BY "
        
        Select Case intGrupoOpcion
            Case 0
                strSQL = strSQL & "FAMILIA"
            Case 1
                strSQL = strSQL & "SUBFAMILIA"
        End Select
        
    strSQL = strSQL & ", NOMPRODUCTO"
    
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
        Case 4016 'Exportar a Excel - Formato 2
            
        Case 4017 'Exportar a PDF
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
    
'    If (ITEM Mod 2) = 0 Then
'        Detail.BackColor = &HFFFFFF
'    Else
'        Detail.BackColor = &H80000016
'    End If
'
'    If Val(Format(fldCantidad.Text, "#.00")) >= 0 Then
'        fldCantidad.ForeColor = vbBlue
'    Else
'        fldCantidad.ForeColor = vbRed
'    End If
End Sub

Private Sub GroupFooter1_AfterPrint()
    'fldPorEntregar.Text = Val(Format(fldCantidadRequerida.Text, "#.00")) - Val(Format(fldCantidadEntregada.Text, "#.00"))
    ITEM = 0
End Sub

