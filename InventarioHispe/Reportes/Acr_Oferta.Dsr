VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Acr_Oferta 
   Caption         =   "Logistica - Oferta"
   ClientHeight    =   10215
   ClientLeft      =   810
   ClientTop       =   1740
   ClientWidth     =   18960
   Icon            =   "Acr_Oferta.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   33443
   _ExtentY        =   18018
   SectionData     =   "Acr_Oferta.dsx":058A
End
Attribute VB_Name = "Acr_Oferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Rem SK ADD:
Private bolDescargarReporte             As Boolean
Private strTipoOrden                    As String
Private strNumeroOrden                  As String

Private strUManterior                   As String
Private bolUMdiferente                  As Boolean

Public Property Let DescargarReporte(ByVal Value As Boolean)
    bolDescargarReporte = Value
End Property

Public Property Get DescargarReporte() As Boolean
    DescargarReporte = bolDescargarReporte
End Property
'Propiedad Tipo de Orden
Public Property Let TipoOrden(ByVal Value As String)
    strTipoOrden = Value
End Property

Public Property Get TipoOrden() As String
    TipoOrden = strTipoOrden
End Property

'Propiedad Numero de Orden
Public Property Let NumeroOrden(ByVal Value As String)
    strNumeroOrden = Value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property

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
        .ITEM(7).Visible = False
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
    
    strUManterior = vbNullString
    bolUMdiferente = False
    
    Dim cQrCode As ClsQrCode
    Set cQrCode = New ClsQrCode
    
    If objAyudaOrden.DerivadoDNI = "" Then objAyudaOrden.DerivadoDNI = "09992055"
    
'    If left(objAyudaOrden.NumeroOrden, 2) = "OC" Then objAyudaOrden.DerivadoDNI = ""
'    ImageQR.Picture = cQrCode.GetPictureQrCode(objAyudaOrden.NumeroOrden & "|" & objAyudaOrden.RucProveedor & "|" & wrucempresa & "|" & objAyudaOrden.TotalFacturado & "|" & objAyudaOrden.TotalImpuesto & "|" & objAyudaOrden.FechaEntrega & "|" & objAyudaOrden.CodMoneda & "|" & objAyudaOrden.DerivadoDNI & "|", ImageQR.Width, ImageQR.Height)
'    SavePicture ImageQR.Picture, "c:\bancowin\qr.bmp"
    
End Sub

Private Sub ActiveReport_ReportEnd()
    On Error GoTo errReportEnd
    
    If bolDescargarReporte Then
        If Me.Pages.Count > 0 Then
            Dim oPDF As ActiveReportsPDFExport.ARExportPDF
            
            Set oPDF = New ActiveReportsPDFExport.ARExportPDF
                
            oPDF.FileName = wrutatemp & "\ParaAtencionDeOrden.pdf"
            oPDF.Export Me.Pages
            
            Me.Visible = False
            
            Rem SK ADD:
            With EnviarCorreoOcx
                .EmailRemitente = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2USEMAIL", "EF2USERS", "F2CODUSER", wusuario, "T")
                .EmailContrasena = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2PASWMAIL", "EF2USERS", "F2CODUSER", wusuario, "T")
                .NombreRemitente = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", wusuario, "T")
                
                .RutaArchivoAdjunto = wrutatemp & "ParaAtencionDeOrden.pdf"
                
                .EmailDestinatario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "LCASE(F2EMAIL)", "EF2PROVEEDORES", "F2NEWRUC", Trim(F4CODPRV.Text & ""), "T")
                .EmailDestinatarioCC = .EmailRemitente
                
                .EmailAsunto = "Para atención de Orden con " & LblNroOC.Caption
                .EmailCuerpo = "Buen día:" & vbNewLine & _
                                "Se adjunta Orden para su atención." & vbNewLine & vbNewLine & _
                                "Por favor, confirmar la recepción y atención de la Orden adjunta con el remitente." & vbNewLine & vbNewLine & _
                                "Sin mas particulares, quedo de Ud." & vbNewLine & vbNewLine & _
                                "Atte." & vbNewLine & vbNewLine & _
                                ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", wusuario, "T") & vbNewLine & _
                                ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CARGO", "EF2USERS", "F2CODUSER", wusuario, "T") & vbNewLine & _
                                wnomcia
                
                .Show 1
                
                If .EmailEnviado Then
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE EF2PROVEEDORES SET F2EMAIL = '" & Trim(EnviarCorreoOcx.EmailDestinatario) & "' WHERE F2NEWRUC = '" & Trim(F4CODPRV.Text & "") & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                    
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE MAESTROS.EF2PROVEEDORES SET F2EMAIL = '" & Trim(EnviarCorreoOcx.EmailDestinatario) & "' WHERE F2NEWRUC = '" & Trim(F4CODPRV.Text & "") & "'"
                        
                        cnBdCPlus.Execute SqlCad
                        
                        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
                    End If
                    
                    
                    With objAyudaOrden
                        .TipoOrden = strTipoOrden
                        .NumeroOrden = strNumeroOrden
                        .DerivadoDNI = ""
                        
                        .Estado = 3
                        .Colocada = True
                        .ColocadaUsuario = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", wusuario, "T")
                        .ColocadaFecha = Date

                        If .enviarViaMailOrden Then
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                With objSqlAyudaOrden
                                    .TipoOrden = strTipoOrden
                                    .NumeroOrden = strNumeroOrden
                                    
                                    .Estado = 3
                                    .Colocada = True
                                    .ColocadaUsuario = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2NOMUSER", "MAESTROS.EF2USERS", "F2CODUSER", wusuario, "T")
                                    .ColocadaFecha = Date
            
                                    If .enviarViaMailOrden Then
                                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                                        
                                        'ordendecompra.TipoOrden = .TipoOrden
                                        'ordendecompra.NumeroOrden = .NumeroOrden
                                        
                                        'ordendecompra.consultarOrdenSql
                                    End If
                                    
                                    .inicializarEntidades
                                End With
                            End If
                            
                            ordendecompra.TipoOrden = .TipoOrden
                            ordendecompra.NumeroOrden = .NumeroOrden
                            
                            ordendecompra.consultarOrden
                        End If
                        
                        .inicializarEntidades
                    End With
                    
                    SqlCad = vbNullString
                End If
            End With
            
            Unload EnviarCorreoOcx
            
            Unload Me
        End If
    End If
    
    Exit Sub
errReportEnd:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    On Error Resume Next
    
    Select Case Tool.ID
        Case 4015: 'EXCEL
            Me.MousePointer = vbHourglass
            
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
            
            Me.MousePointer = vbDefault
        Case 4017: 'WORD
            
        Case 4016: 'PDF
            Me.MousePointer = vbHourglass
            
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
            
            Me.MousePointer = vbDefault
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
'If ordendecompra.dxCheckBox2.Checked = False Then
    F3PREUNI.Text = Format(F3PREUNI.Text, "#,##0.00")
'End If
'#,##0.00


End Sub

Private Sub Detail_Format()
    I = I + 1
    FldItem.Text = I
    FldObserva.Visible = False
    F5NOMPRO.Text = F5NOMPRO.Text & IIf(Len(Trim(FldObserva.Text)) > 0, vbCrLf & "(*) " & FldObserva.Text, "")
    
    If strUManterior = vbNullString Then
        strUManterior = Trim(F3MEDIDA.Text)
    Else
        If strUManterior <> Trim(F3MEDIDA.Text) Then
            bolUMdiferente = True
        End If
    End If
End Sub

Private Sub GroupFooter1_BeforePrint()
FldSon.top = FldObservaAll.Height + FldObservaAll.top
LblSon1.top = FldObservaAll.Height + FldObservaAll.top
LblSon2.top = FldObservaAll.Height + FldObservaAll.top

End Sub

Private Sub GroupFooter1_Format()
    LblBorderDet.Height = FldSon.top + FldSon.Height + 400
    GroupFooter1.Height = FldSon.top + FldSon.Height + 400
    
    lblCantidadTotal.Visible = Not bolUMdiferente
    fldCantidadTotal.Visible = Not bolUMdiferente
    fldCantidadTotal.Text = Format(Val(Format(fldCantidadTotal.Text, "#0.00")), "#0.00") & " " & strUManterior
End Sub

