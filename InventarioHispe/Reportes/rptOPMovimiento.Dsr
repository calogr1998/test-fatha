VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptOPMovimiento 
   Caption         =   "Logistica - rptOPMovimiento (ActiveReport)"
   ClientHeight    =   8895
   ClientLeft      =   210
   ClientTop       =   1725
   ClientWidth     =   12570
   Icon            =   "rptOPMovimiento.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   22172
   _ExtentY        =   15690
   SectionData     =   "rptOPMovimiento.dsx":058A
End
Attribute VB_Name = "rptOPMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strNumeroOrden As String

Private strSQL As String

Dim ITEM As Integer
Dim i As Integer

Public Property Let NumeroOrden(ByVal value As String)
    strNumeroOrden = value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
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
    obtenerCabeceraReporte
    
    imprimirOrdenMovimiento
End Sub

Private Sub obtenerCabeceraReporte()
    Dim cmdOP As ADODB.Command
    Dim rstOp As ADODB.Recordset
    Dim intCantidadMesesDeValidezCompromiso As Integer
    
    intCantidadMesesDeValidezCompromiso = Val(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "CantidadMesesDeValidezCompromiso", "l"))
    
    Set cmdOP = New ADODB.Command
    Set rstOp = New ADODB.Recordset
    
    With cmdOP
        .ActiveConnection = cnBdStudioModa
        .CommandType = adCmdStoredProc
        .CommandText = "usp_ConsultaOrdenProduccionCP"
        
        .Parameters.Append .CreateParameter("@IDOP", adBigInt, adParamInput, , Val(strNumeroOrden))
        .Parameters.Append .CreateParameter("@CATEGORIA", adVarChar, adParamInput, 10, vbNullString)
        .Parameters.Append .CreateParameter("@OP", adVarChar, adParamInput, 30, vbNullString)
        .Parameters.Append .CreateParameter("@CANTIDADMESESVALIDEZPEDIDO", adBigInt, adParamInput, , intCantidadMesesDeValidezCompromiso)
        
        Set rstOp = .Execute()
    End With
    
    Set cmdOP = Nothing
    
    If Not rstOp.EOF Then
        fldCategoria.Text = Trim(rstOp!CATEGORIA & "")
        fldNroOP.Text = Trim(rstOp!NroOP & "")
        fldFecha.Text = Trim(rstOp!FECOP & "")
        fldEstado.Text = Trim(rstOp!ESTADOOPLETRAS & "")
        
        fldNroPedido.Text = Trim(rstOp!NroPedido & "")
        fldFechaEntrega.Text = Trim(rstOp!FENTREGA & "")
        fldModelo.Text = Trim(rstOp!Modelo & "")
        fldColor.Text = Trim(rstOp!Color & "")
    End If
    
    If rstOp.State = 1 Then rstOp.Close
    
    Set rstOp = Nothing
    
    intCantidadMesesDeValidezCompromiso = 0
End Sub

Private Sub imprimirOrdenMovimiento()
    strSQL = vbNullString
    
    With objAyudaVale
        .OrdenTrabajo = strNumeroOrden
            
        strSQL = .devuelveSQLOrdenProduccionMovimiento
    End With
    
    With dtcVale
        .ConnectionString = cnn_dbbancos
        .CursorLocation = ddADOUseClient
        .CursorType = ddADOOpenStatic
        .LockType = ddADOLockReadOnly
        .Source = strSQL
    End With
    
    Me.Caption = App.ProductName & " - Movimiento de Orden de Producción"
    Me.lblSistema.Caption = App.LegalTrademarks & "-" & App.ProductName
    Me.lblEmpresa.Caption = wnomcia
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

