VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmUtilStockDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle y Re-distribución de Stock"
   ClientHeight    =   7710
   ClientLeft      =   810
   ClientTop       =   1875
   ClientWidth     =   12945
   Icon            =   "frmUtilStockDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12945
   Begin ActiveToolBars.SSActiveToolBars tlbDetalle 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   2
      ToolsCount      =   6
      Tools           =   "frmUtilStockDetalle.frx":058A
      ToolBars        =   "frmUtilStockDetalle.frx":2057
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDetalle 
      Height          =   7170
      Left            =   120
      OleObjectBlob   =   "frmUtilStockDetalle.frx":22A6
      TabIndex        =   0
      Top             =   120
      Width           =   12705
   End
End
Attribute VB_Name = "frmUtilStockDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTipoNaturaleza As String
Private strTipoDetalle As String
Private strCodProducto As String
Private strCodAlmacen As String

Private bolRedistribucionEjecutada As Boolean
Private bolDeshabilitarRedistribucion As Boolean
Private strNroPedidoSolicitante As String
Private dblCantidadMaximaParaPedido As Double

Private objValeStockDet As ClsVale
Private objOrdenStockDet As ClsOrden


'Propiedad Tipo de Naturaleza [ Stock (F)isico o (V)irtual ]
Public Property Let TipoNaturaleza(ByVal value As String)
    strTipoNaturaleza = value
End Property

Public Property Get TipoNaturaleza() As String
    TipoNaturaleza = strTipoNaturaleza
End Property

'Propiedad Tipo de Detalle [ Stock (C)omprometido o (L)ibre ]
Public Property Let TipoDetalle(ByVal value As String)
    strTipoDetalle = value
End Property

Public Property Get TipoDetalle() As String
    TipoDetalle = strTipoDetalle
End Property

'Propiedad Codigo de Producto
Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

'Propiedad Codigo de Almacen
Public Property Let CodigoAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property



'Propiedad Se activara al Redistribuir Stock
Public Property Let RedistribucionEjecutada(ByVal value As Boolean)
    bolRedistribucionEjecutada = value
End Property

Public Property Get RedistribucionEjecutada() As Boolean
    RedistribucionEjecutada = bolRedistribucionEjecutada
End Property

'Propiedad Se activara para Deshabilitar la Opcion de Re-Distribucion
Public Property Let DeshabilitarRedistribucion(ByVal value As Boolean)
    bolDeshabilitarRedistribucion = value
End Property

Public Property Get DeshabilitarRedistribucion() As Boolean
    DeshabilitarRedistribucion = bolDeshabilitarRedistribucion
End Property
'Propieda de Nro de Pedido Solicitante de Re-distribucion
Public Property Let NroPedidoSolicitante(ByVal value As String)
    strNroPedidoSolicitante = value
End Property

Public Property Get NroPedidoSolicitante() As String
    NroPedidoSolicitante = strNroPedidoSolicitante
End Property
'Propiedad Cantidad Maxima Para Pedido Solicitante
Public Property Let CantidadMaximaParaPedido(ByVal value As Double)
    dblCantidadMaximaParaPedido = value
End Property

Public Property Get CantidadMaximaParaPedido() As Double
    CantidadMaximaParaPedido = dblCantidadMaximaParaPedido
End Property



Private Sub configurarGrilla()
    With dbgDetalle.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        '.Set (egoCanDelete)
        '.Set (egoCanAppend)
        '.Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoExpandOnDblClick)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
End Sub

Private Sub listarDetalle()
    objAyudaVale.listarGrillaMovimientoProductoDetalleV2 dbgDetalle, Nothing, strCodAlmacen, strCodProducto, strTipoNaturaleza, strTipoDetalle
End Sub



Private Sub dbgDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case UCase(Column.FieldName)
        Case "PROCESAR"
            With dbgDetalle
                If .Dataset.State = dsEdit Then
                    .Dataset.Post
                End If
            End With
    End Select
End Sub

Private Sub dbgDetalle_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "NROPEDIDODESTINO"
            If Trim(Text) = vbNullString Then
                Font.Bold = False
                FontColor = vbBlack
            Else
                Font.Bold = True
                FontColor = vbGreen
            End If
        Case "CANTIDAD", "CANTIDADDESTINO"
            If Val(Text) < 0 Then
                FontColor = vbRed
            Else
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgDetalle_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "CANTIDAD", "CANTIDADDESTINO"
            Color = vbWhite
            
            If Val(Text) < 0 Then
                FontColor = vbRed
            Else
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgDetalle.Columns.FocusedColumn.FieldName
        Case "NROPEDIDODESTINO"
            With dbgDetalle
                If .Dataset.State = dsEdit Then
                    If bolDeshabilitarRedistribucion Then
                        MsgBox "Re-distribución deshabilitada, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If strNroPedidoSolicitante <> vbNullString Then
                        If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") <> strNroPedidoSolicitante Then
                            MsgBox "El No. Pedido de re-distribución no puede ser diferente al requerido, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & ""), "N") = vbNullString Then
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    .Dataset.Post
                End If
                
                .Dataset.Edit
                    
                .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & ""), "N")
                
                .Dataset.Post
            End With
        Case "CANTIDADDESTINO"
            With dbgDetalle
                If .Dataset.State = dsEdit Then
                    If bolDeshabilitarRedistribucion Then
                        MsgBox "Re-distribución deshabilitada, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") Then
                        MsgBox "La cantidad de re-distribución no puede ser mayor al stock disponible, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If strNroPedidoSolicitante <> vbNullString Or dblCantidadMaximaParaPedido > 0 Then
                        If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > dblCantidadMaximaParaPedido Then
                            MsgBox "La cantidad de re-distribución no puede ser mayor a la Cantidad Maxima especificada, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    If strTipoDetalle = "L" Then
                        If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") = vbNullString And strNroPedidoSolicitante = vbNullString Then
                            MsgBox "El stock libre no puede ser re-distribuido sin consignar el No. Pedido Destino, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    .Dataset.Post
                End If
                                    
                .Dataset.Edit
                
                If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") = vbNullString Then
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = strNroPedidoSolicitante
                End If
                
                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") <= 0 Then
                    .Columns.ColumnByFieldName("PROCESAR").value = False
                Else
                    .Columns.ColumnByFieldName("PROCESAR").value = True
                End If
                
                .Dataset.Post
            End With
        Case "PROCESAR"
            With dbgDetalle
                If .Dataset.State = dsEdit Then
                     If bolDeshabilitarRedistribucion Then
                        MsgBox "Re-distribución deshabilitada, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If strTipoDetalle = "L" Then
                        If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") = vbNullString And strNroPedidoSolicitante = vbNullString Then
                            MsgBox "El stock libre no puede ser re-distribuido sin consignar el No. Pedido Destino, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    If strNroPedidoSolicitante <> vbNullString Then
                        If CBool(.Columns.ColumnByFieldName("PROCESAR").value) And Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > dblCantidadMaximaParaPedido Then
                            MsgBox "La cantidad de re-distribución no puede ser mayor a la cantidad requerida por el Pedido Actual, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    .Dataset.Post
                End If
                
                .Dataset.Edit
                
                If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") = vbNullString And strNroPedidoSolicitante <> vbNullString Then
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = strNroPedidoSolicitante
                    
                    abrirCnTemporal
                    
                    If Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") > dblCantidadMaximaParaPedido Then
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = dblCantidadMaximaParaPedido ' - Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADDESTINO)", "TMPUTILSTOCKDETALLE", "NROPEDIDODESTINO", Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & ""), "T", "AND PROCESAR = TRUE"))
                    Else
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                    End If
                    
                    abrirCnTemporal
                End If
                
                'If .Columns.ColumnByFieldName("PROCESAR").value <> Null Then
                    If CBool(.Columns.ColumnByFieldName("PROCESAR").value & "") Then
                        If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                            If dblCantidadMaximaParaPedido > 0 Then
                                .Columns.ColumnByFieldName("CANTIDADDESTINO").value = IIf(Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") <= dblCantidadMaximaParaPedido, Val(.Columns.ColumnByFieldName("CANTIDAD").value & ""), dblCantidadMaximaParaPedido)
                            Else
                                .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                            End If
                        End If
                        
                        'If strNroPedidoSolicitante <> vbNullString Then
                        '    dblCantidadMaximaParaPedido = dblCantidadMaximaParaPedido - Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "")
                        'End If
                    Else
                        'If strNroPedidoSolicitante <> vbNullString Then
                        '    dblCantidadMaximaParaPedido = dblCantidadMaximaParaPedido + Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "")
                        'End If
                        
                        .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                    End If
                    
                    .Columns.ColumnByFieldName("PROCESAR").value = CBool(.Columns.ColumnByFieldName("PROCESAR").value)
                
                'End If
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
'    Select Case KeyCode
'        Case vbKeyEscape
'            tlbDetalle_ToolClick tlbDetalle.Tools("Salir")
'    End Select
End Sub

Private Sub Form_Activate()
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
End Sub

Private Sub Form_Load()
    configurarGrilla
    

        listarDetalle
    
    bolRedistribucionEjecutada = False
    
    tlbDetalle.Tools("Redistribuir").Enabled = Not bolDeshabilitarRedistribucion
    tlbDetalle.Tools("ID_NroPedidoT").Edit.Text = strNroPedidoSolicitante
    tlbDetalle.Tools("ID_CantidadReqT").Edit.Text = dblCantidadMaximaParaPedido
    
    'ModUtilitario.deshabilitarBotonCerrarForm frmUtilStockDetalle
End Sub

Private Sub tlbDetalle_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Redistribuir"
            Select Case strTipoNaturaleza
                Case "F"
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        redistribuirStockFisicoSql
                    Else
                        redistribuirStockFisico
                    End If
                Case "V"
                    If strTipoDetalle = "L" Then
                        'redistribuirStockVirtual
                        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                            redistribuirStockVirtualV2Sql
                        Else
                            redistribuirStockVirtualV2
                        End If
                    Else
                        MsgBox "Esta opción no aplica para Stock Comprometido Por Llegar.", vbInformation + vbOKOnly, App.ProductName
                    End If
            End Select
            
            If ModUtilitario.validarFormAbierto("frmUtilStock") Then
                frmUtilStock.listarStock
            End If
            
            If bolRedistribucionEjecutada Then
                dbgDetalle.Dataset.Close
                
                Unload Me
            Else
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    listarDetalleSql
                Else
                    listarDetalle
                End If
            End If
'        Case "Salir"
'            If dbgDetalle.Dataset.State = dsEdit Then
'                dbgDetalle.Dataset.Post
'            Else
'                dbgDetalle.Dataset.Edit
'
'                dbgDetalle.Dataset.Post
'            End If
'
'            bolRedistribucionEjecutada = False
'
'            dbgDetalle.Dataset.Close
'
'            Me.Hide
    End Select
End Sub

