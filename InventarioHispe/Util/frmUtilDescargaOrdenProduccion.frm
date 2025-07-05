VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmUtilDescargaOrdenProduccion 
   Caption         =   "Descarga de Orden de Producción"
   ClientHeight    =   9360
   ClientLeft      =   1470
   ClientTop       =   870
   ClientWidth     =   12975
   Icon            =   "frmUtilDescargaOrdenProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
      Begin MSComctlLib.ProgressBar pgbProceso 
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraResultado 
      Caption         =   " Resultado de Descarga "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   12735
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11040
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dbgResultado 
         Height          =   2130
         Left            =   120
         OleObjectBlob   =   "frmUtilDescargaOrdenProduccion.frx":058A
         TabIndex        =   5
         Top             =   240
         Width           =   12465
      End
   End
   Begin VB.Frame fraDetalleStock 
      Caption         =   " Stock Disponible de Producto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   12735
      Begin MSComCtl2.DTPicker dtpFechaDespacho 
         Height          =   300
         Left            =   10920
         TabIndex        =   0
         Top             =   200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   20905985
         CurrentDate     =   41978
      End
      Begin VB.ComboBox cmbTipoStock 
         Height          =   315
         ItemData        =   "frmUtilDescargaOrdenProduccion.frx":4029
         Left            =   5040
         List            =   "frmUtilDescargaOrdenProduccion.frx":4036
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         ItemData        =   "frmUtilDescargaOrdenProduccion.frx":411C
         Left            =   840
         List            =   "frmUtilDescargaOrdenProduccion.frx":4126
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dbgDetalleStock 
         Height          =   2250
         Left            =   120
         OleObjectBlob   =   "frmUtilDescargaOrdenProduccion.frx":4155
         TabIndex        =   4
         Top             =   600
         Width           =   12465
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8640
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Stock"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraPendiente 
      Caption         =   " Pendiente de Descarga "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   12735
      Begin DXDBGRIDLibCtl.dxDBGrid dbgPendiente 
         Height          =   2610
         Left            =   120
         OleObjectBlob   =   "frmUtilDescargaOrdenProduccion.frx":807A
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   12465
      End
   End
End
Attribute VB_Name = "frmUtilDescargaOrdenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strIdOrdenProduccion As String

Private strFichero          As String

Private objValeDescargaOP   As ClsVale

Private bolOPCargada        As Boolean

'Propiedad ID Orden de Produccion
Public Property Let IdOrdenProduccion(ByVal value As String)
    strIdOrdenProduccion = value
End Property

Public Property Get IdOrdenProduccion() As String
    IdOrdenProduccion = strIdOrdenProduccion
End Property

Private Sub configurarGrilla()
    With dbgPendiente.Options
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

Private Sub listarAlmacenEnCombo()
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        objSqlAyudaAlmacen.listarAlmacenSoloSeleccion cmbAlmacen
    Else
        objAyudaAlmacen.listarAlmacenSoloSeleccion cmbAlmacen
    End If
    
'    Dim rstAlmacen As New ADODB.Recordset
'
'    If rstAlmacen.State = 1 Then rstAlmacen.Close
'
'    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    cmbalmacen.Clear
'
'    If Not rstAlmacen.EOF Then
'        rstAlmacen.MoveFirst
'
'        Do While Not rstAlmacen.EOF
'            cmbalmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
'
'            rstAlmacen.MoveNext
'        Loop
            If cmbAlmacen.ListCount > 0 Then
                cmbAlmacen.ListIndex = 0
            End If
'    End If
End Sub

Private Sub obtenerStockProductoPendiente()
    Dim rstStockPP As New ADODB.Recordset
    
    dbgPendiente.Dataset.Close
    
    fraProceso.Visible = True
    
    abrirCnTemporal
    
    If rstStockPP.State = 1 Then rstStockPP.Close
    
    rstStockPP.Open "SELECT * FROM TMPUTILDESCARGAOPPENDIENTE WHERE SALDO > 0 ORDER BY NOMPRODUCTO", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStockPP.EOF Then
        DoEvents
        
        fraProceso.Visible = True
        pgbProceso.value = 0
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStockPP)
        fraProceso.Caption = "Descargando Stock..."
        
        Do While Not rstStockPP.EOF
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaVale
                    .inicializarEntidades
                    .inicializarEntidadesDetalle

                    .CodigoProducto = Trim(rstStockPP!CODPRODUCTOFINAL & "")
                    .CodigoAlmacen = right(cmbAlmacen.Text, 2)

                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "STOCK = " & .devuelveStockFisicoDeProducto(vbNullString, True) & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
                    
                    'If .CodigoProducto = "9T753-8" Then MsgBox "DETENTE!"
                    
                    cnDBTemp.Execute SqlCad

                    .inicializarEntidades
                    .inicializarEntidadesDetalle
                End With
            Else
                With objAyudaVale
                    .inicializarEntidades
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstStockPP!CODPRODUCTOFINAL & "")
                    .CodigoAlmacen = right(cmbAlmacen.Text, 2)
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "STOCK = " & .devuelveStockFisicoDeProductoV2(vbNullString, True) & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
                    
                    cnDBTemp.Execute SqlCad
                    
                    .inicializarEntidades
                    .inicializarEntidadesDetalle
                End With
            End If
            
            DoEvents

            pgbProceso.value = pgbProceso.value + 1
            fraProceso.Caption = "Descargando Stock... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 0)
            
            rstStockPP.MoveNext
        Loop
    End If
    
    fraProceso.Visible = False
    
    abrirCnTemporal
End Sub

Private Sub obtenerPendiente()
    objAyudaOrdenTrabajo.listarGrillaInsumoPendienteDescarga dbgPendiente, strIdOrdenProduccion
End Sub

Private Sub listarDetalleStock()
    If cmbAlmacen.ListIndex <> -1 And cmbTipoStock.ListIndex <> -1 Then
        Me.MousePointer = vbHourglass
        
        dbgDetalleStock.Dataset.Close
        
        'objAyudaVale.listarGrillaMovimientoProductoDetalle dbgDetalleStock, _
                                                            Nothing, _
                                                            right(Trim(cmbalmacen.Text), 2), _
                                                            Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), _
                                                            "F", _
                                                            right(Trim(cmbTipoStock.Text), 1), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & ""), vbNullString), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, IIf(right(Trim(cmbTipoStock.Text), 1) = "C", True, False), False)
        
        With objAyudaVale
'            .inicializarEntidades
'
'            .CodigoAlmacen = right(cmbAlmacen.Text, 2)
'            .CodigoProducto = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
'
'            If .devuelveStockFisicoDeProducto(vbNullString, True) > 0 Then
                .listarGrillaMovimientoProductoDetalleV2 dbgDetalleStock, _
                                                            Nothing, _
                                                            Trim(right(Trim(cmbAlmacen.Text), 2)), _
                                                            Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), _
                                                            "F", _
                                                            right(Trim(cmbTipoStock.Text), 1), _
                                                            IIf(cmbTipoStock.ListIndex = 0, Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & ""), vbNullString), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, IIf(right(Trim(cmbTipoStock.Text), 1) = "C", True, False), False)
'            Else
'                abrirCnTemporal
'
'                cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKDETALLE"
'            End If
'
'            .inicializarEntidades
        End With
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub listarDetalleStockSql()
    If cmbAlmacen.ListIndex <> -1 And cmbTipoStock.ListIndex <> -1 Then
        Me.MousePointer = vbHourglass
        
        dbgDetalleStock.Dataset.Close
        
        'objAyudaVale.listarGrillaMovimientoProductoDetalle dbgDetalleStock, _
                                                            Nothing, _
                                                            right(Trim(cmbalmacen.Text), 2), _
                                                            Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), _
                                                            "F", _
                                                            right(Trim(cmbTipoStock.Text), 1), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & ""), vbNullString), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, IIf(right(Trim(cmbTipoStock.Text), 1) = "C", True, False), False)
        
        With objSqlAyudaVale
            .inicializarEntidades
'
            .CodigoAlmacen = right(cmbAlmacen.Text, 2)
            .CodigoProducto = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
'
'            If .devuelveStockFisicoDeProducto(vbNullString, True) > 0 Then
                .listarGrillaMovimientoProductoDetalleV2 dbgDetalleStock, _
                                                            Nothing, _
                                                            "F", _
                                                            right(Trim(cmbTipoStock.Text), 1), _
                                                            "tmpCPStockDetalle" & UCase(wusuario), _
                                                            IIf(cmbTipoStock.ListIndex = 0, Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & ""), vbNullString), _
                                                            IIf(cmbTipoStock.ListIndex <> 2, IIf(right(Trim(cmbTipoStock.Text), 1) = "C", True, False), False)
                                                            
                
                
'            Else
'                abrirCnTemporal
'
'                cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKDETALLE"
'            End If
'
'            .inicializarEntidades
        End With
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub listarPendiente()
    With dbgPendiente
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctKeyset
        .Dataset.ADODataset.LockType = ltOptimistic
        .Dataset.ADODataset.CommandText = "SELECT * FROM TMPUTILDESCARGAOPPENDIENTE WHERE SALDO > 0 ORDER BY NOMPRODUCTO"
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODPRODUCTOORIGEN"
        
        .m.FullExpand
        .m.ResetFullRefresh
        .m.FullRefresh
    End With
End Sub

Private Sub listarDetalleVale()
    With dbgResultado
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = "SELECT * FROM TMPVALESALIDA ORDER BY DESCRIPCION"
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "ITEM"
        
        .Columns.ColumnByFieldName("DESCRIPCION").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) descargado(s)."
    End With
End Sub

Private Sub redistribuirStockFisico()
    On Error GoTo errRedistribuirStockFisico
    
    Dim strCodigoOriginal As String
    
    If dbgDetalleStock.Dataset.State = dsEdit Then
        dbgDetalleStock.Dataset.Post
    End If
    
    strCodigoOriginal = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOORIGEN").value & "")
    
    dbgDetalleStock.Dataset.Close
    dbgDetalleStock.Dataset.Open
    
    dbgDetalleStock.m.FullExpand
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(CODPRODUCTO) AS CANTIDAD", "TMPUTILSTOCKDETALLE", "PROCESAR", "TRUE", "N")) = 0 Then
        MsgBox "No se registran Items con re-distribución por Procesar.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    If cmbTipoStock.ListIndex = 2 Then
        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(NROPEDIDO) AS CANTIDAD", "TMPUTILSTOCKDETALLE", "PROCESAR", "TRUE", "N", "AND CODPRODUCTO = '" & Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "") & "' AND NROPEDIDO = '" & Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "'")) = 1 Then
            MsgBox "Imposible re-distribuir Stock ya comprometido para atencion de OP.", vbInformation + vbOKOnly, App.ProductName
            
            dbgPendiente_OnClick
            
            Exit Sub
        End If
    End If
    
    Select Case cmbTipoStock.ListIndex
        Case 1, 2
            If MsgBox("¿Desea re-distribuir y descargar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                With dbgDetalleStock
                    .Dataset.Edit
                    
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                    .Columns.ColumnByFieldName("PROCESAR").value = False
                    
                    .Dataset.Post
                End With
                
                Exit Sub
            End If
        Case Else
            If MsgBox("¿Desea re-distribuir Stock Disponible?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                With dbgDetalleStock
                    .Dataset.Edit
                    
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                    .Columns.ColumnByFieldName("PROCESAR").value = False
                    
                    .Dataset.Post
                End With
                
                Exit Sub
            End If
    End Select
    
    Dim rstValeDet As New ADODB.Recordset
    Dim dblItem As Double
    
    Set objValeDescargaOP = New ClsVale
    
    With objValeDescargaOP
        .inicializarEntidades
        
        .CodigoAlmacen = right(Trim(cmbAlmacen.Text), 2)
        .NumeroVale = vbNullString
        .TipoVale = "I"
        
        .Fecha = Format(dtpFechaDespacho.value, "Short Date")
        .CodigoOrigen = "XCS"
        .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
        
            If .TipoCambio = 0 Then
                .TipoCambio = "2.8"
            End If
        
        .CodigoMoneda = "S"
        
        .referencia = wnomcia
        .observaciones = Me.Caption
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If .guardarVale Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            'Borrar Detalle de Vale
            SqlCad = vbNullString
            SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnn_dbbancos.Execute SqlCad
            Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT * FROM TMPUTILSTOCKDETALLE WHERE PROCESAR = TRUE ORDER BY NROPEDIDO"
            
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            abrirCnTemporal
            
            rstValeDet.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                dblItem = 0
                
'                If .verificarStockProductoFisicoCorL(Trim(rstValeDet!CodProducto & ""), _
                                                        .CodigoAlmacen, _
                                                        Trim(rstValeDet!NROOC & ""), _
                                                        Trim(rstValeDet!NroPedido & ""), _
                                                        Val(rstValeDet!CANTIDADDESTINO & ""), _
                                                        dtpFechaDespacho.value) Then
                
                    Do While Not rstValeDet.EOF
                        .inicializarEntidadesDetalle
                        
                        .NumeroOrdenCompra = Trim(rstValeDet!NROOC & "")
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = Trim(rstValeDet!NroPedido & "")
                        
                        .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                        .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                        .Cantidad = Val(rstValeDet!CANTIDADDESTINO & "") * -1
                        
                        .ObservacionesPorItem = "SE RETIRA DE " & IIf(.Requerimiento <> vbNullString, "STOCK COMPROMETIDO DEL PEDIDO N° " & .Requerimiento, "STOCK LIBRE.")
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        'SI SE RE-DISTRIBUYE EL STOCK COMPROMETIDO; GENERAR REGISTRO DE REPOSICION DE STOCK
                        If cmbTipoStock.ListIndex = 2 Then
                            'Llenar Historial de Reposición
                            Dim lngIdReposicion As Long
                            
                            lngIdReposicion = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 IDREPOSICION", "SF1REPOSICIONCOMPROMISO", vbNullString, vbNullString, vbNullString, "IDREPOSICION > 0 ORDER BY IDREPOSICION DESC") & "") + 1
                            
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "INSERT INTO SF1REPOSICIONCOMPROMISO("
                            SqlCad = SqlCad & "IDREPOSICION, NROPEDIDO, IDINSUMO, "
                            SqlCad = SqlCad & "CANTIDAD, OBSERVACION, USUDESCOMPROMISO, "
                            SqlCad = SqlCad & "FECDESCOMPROMISO) "
                            SqlCad = SqlCad & "VALUES("
                            SqlCad = SqlCad & lngIdReposicion & ", "
                            SqlCad = SqlCad & "'" & .Requerimiento & "', "
                            SqlCad = SqlCad & "'" & .CodigoProducto & "', "
                            SqlCad = SqlCad & Abs(.Cantidad) & ", "
                            SqlCad = SqlCad & "'DESCOMPROMISO EJECUTADO CON VALE Nº " & .CodigoAlmacen & "/" & .NumeroVale & " POR " & wusuario & " PARA ATENCIÓN DE: " & .observaciones & "', "
                            SqlCad = SqlCad & "'" & wusuario & "',"
                            SqlCad = SqlCad & "CVDATE('" & Now & "'))"
                            
                            cnn_dbbancos.Execute SqlCad
                            
                            Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                        End If
                        
                        .inicializarEntidadesDetalle
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = Trim(rstValeDet!NROPEDIDODESTINO & "")
                        
                        .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                        .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                        .Cantidad = Val(rstValeDet!CANTIDADDESTINO & "")
                        
                        .ObservacionesPorItem = "SE CARGA EL " & IIf(.Requerimiento <> vbNullString, "STOCK COMPROMETIDO DEL PEDIDO N° " & .Requerimiento, "STOCK LIBRE.")
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        rstValeDet.MoveNext
                    Loop
                        If dblItem > 0 Then
                            replicarRedistribucionFisica .TipoVale, .NumeroVale, .CodigoAlmacen
                        End If
'                Else
'                    MsgBox "Stock seleccionado no disponible al '" & dtpFechaDespacho.value & "' para atender su descarga." & vbNewLine & vbNewLine & _
'                            "RECOMENDACIÓN: Vuelva a seleccionar el Producto a descargar para" & vbNewLine & _
'                            "actualizar el Stock Disponible.", vbInformation + vbOKOnly, App.ProductName
'
'                    If .eliminarVale Then
'                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'
'                        listarDetalleStock
'                    End If
'
'                    Exit Sub
'                End If
            End If
        End If
        
        Select Case cmbTipoStock.ListIndex
            Case 1, 2
                dbgResultado.Dataset.Close
                
                'INSERTAMOS LA DESCARGA DEL ITEM SELECCIONADO
                .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", _
                                                        "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, _
                                                        "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TMPVALESALIDA("
                SqlCad = SqlCad & "ITEM, "
                SqlCad = SqlCad & "CODPROD, "
                SqlCad = SqlCad & "CODPRODORIGINAL, "
                SqlCad = SqlCad & "DESCRIPCION, "
                SqlCad = SqlCad & "UMEDIDA, "
                SqlCad = SqlCad & "CANTIDAD, "
                SqlCad = SqlCad & "CANTIDADMAX, "
                SqlCad = SqlCad & "F4NUMORD, "
                SqlCad = SqlCad & "COD_SOLICITUD"
                SqlCad = SqlCad & ") "
                
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & .ITEM & " AS ITEM, "
                SqlCad = SqlCad & "'" & .CodigoProducto & "' AS FINAL, "
                SqlCad = SqlCad & "'" & .CodigoProductoOriginal & "' AS ORIGEN, "
                SqlCad = SqlCad & "SD.NOMPRODUCTO, "
                SqlCad = SqlCad & "SD.UM, "
                SqlCad = SqlCad & "SD.CANTIDADDESTINO AS CANTIDAD, "
                SqlCad = SqlCad & "SD.CANTIDADDESTINO AS CANTIDADMAX, "
                SqlCad = SqlCad & "SD.NROOC, "
                SqlCad = SqlCad & "SD.NROPEDIDODESTINO "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "TMPUTILSTOCKDETALLE AS SD "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "SD.CODPRODUCTO = '" & .CodigoProducto & "' AND "
                SqlCad = SqlCad & "TRIM(SD.NROOC & '') = '" & .NumeroOrdenCompra & "' AND "
                SqlCad = SqlCad & "TRIM(SD.NROPEDIDODESTINO & '') = '" & .Requerimiento & "'"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                'ACTUALIZAMOS EL SALDO EN EL ITEM PENDIENTE SELECCIONADO PARA DESCARGA
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "SALDO = SALDO - " & .Cantidad & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CODPRODUCTOORIGEN = '" & strCodigoOriginal & "' AND "
                SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                'MsgBox "Proceso de Re-distribución y Descarga finalizada.", vbInformation + vbOKOnly, App.ProductName
                
                cmbTipoStock.ListIndex = -1
                
                listarDetalleVale
                
                listarPendiente
            Case Else
                'MsgBox "Proceso de Re-distribución finalizado.", vbInformation + vbOKOnly, App.ProductName
                
                cmbTipoStock.ListIndex = -1
                
                listarDetalleVale
                
                listarPendiente
        End Select
    End With
    
    Set objValeDescargaOP = Nothing
    
    Exit Sub
errRedistribuirStockFisico:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - frmUtilDescargaOP: redistribuirStockFisico"
    End Select
    
    Err.Clear
End Sub

Private Sub redistribuirStockFisicoSql()
    On Error GoTo errRedistribuirStockFisicoSql
    
    Dim strCodigoOriginal As String
    
    If dbgDetalleStock.Dataset.State = dsEdit Then
        dbgDetalleStock.Dataset.Post
    End If
    
    strCodigoOriginal = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOORIGEN").value & "")
    
    dbgDetalleStock.Dataset.Close
    dbgDetalleStock.Dataset.Open
    
    dbgDetalleStock.m.FullExpand
    
    If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(CODPRODUCTO) AS CANTIDAD", "TMPCPSTOCKDETALLE" & UCase(wusuario), "PROCESAR", "1", "N")) = 0 Then
        MsgBox "No se registran Items con re-distribución por Procesar.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    If cmbTipoStock.ListIndex = 2 Then
        If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(NROPEDIDO) AS CANTIDAD", "TMPCPSTOCKDETALLE" & UCase(wusuario), "PROCESAR", "1", "N", "AND CODPRODUCTO = '" & Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "") & "' AND NROPEDIDO = '" & Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "'")) = 1 Then
            MsgBox "Imposible re-distribuir Stock ya comprometido para atencion de OP.", vbInformation + vbOKOnly, App.ProductName
            
            dbgPendiente_OnClick
            
            Exit Sub
        End If
    End If
    
    Select Case cmbTipoStock.ListIndex
        Case 1, 2
            If MsgBox("¿Desea re-distribuir y descargar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                With dbgDetalleStock
                    .Dataset.Edit
                    
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                    .Columns.ColumnByFieldName("PROCESAR").value = False
                    
                    .Dataset.Post
                End With
                
                Exit Sub
            End If
        Case Else
            If MsgBox("¿Desea re-distribuir Stock Disponible?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                With dbgDetalleStock
                    .Dataset.Edit
                    
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                    .Columns.ColumnByFieldName("PROCESAR").value = False
                    
                    .Dataset.Post
                End With
                
                Exit Sub
            End If
    End Select
    
    Dim rstValeDet As New ADODB.Recordset
    Dim dblItem As Double
    
    Set objValeDescargaOP = New ClsVale
    
    With objValeDescargaOP
        .inicializarEntidades
        
        .CodigoAlmacen = right(Trim(cmbAlmacen.Text), 2)
        .NumeroVale = vbNullString
        .TipoVale = "I"
        
        .Fecha = Format(dtpFechaDespacho.value, "Short Date")
        .CodigoOrigen = "XCS"
        .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
        
            If .TipoCambio = 0 Then
                .TipoCambio = "2.8"
            End If
        
        .CodigoMoneda = "S"
        
        .referencia = wnomcia
        .observaciones = Me.Caption
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If .guardarVale Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            'Borrar Detalle de Vale
            SqlCad = vbNullString
            SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnn_dbbancos.Execute SqlCad
            Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT * FROM TMPCPSTOCKDETALLE" & UCase(wusuario) & " WHERE PROCESAR = 1 ORDER BY NROPEDIDO"
            
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            abrirCnTemporal
            
            rstValeDet.Open SqlCad, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                dblItem = 0
                
'                If .verificarStockProductoFisicoCorL(Trim(rstValeDet!CodProducto & ""), _
                                                        .CodigoAlmacen, _
                                                        Trim(rstValeDet!NROOC & ""), _
                                                        Trim(rstValeDet!NroPedido & ""), _
                                                        Val(rstValeDet!CANTIDADDESTINO & ""), _
                                                        dtpFechaDespacho.value) Then
                
                With objSqlAyudaVale
                    .inicializarEntidades
                    .inicializarEntidadesAdicionales
                    
                    .Fecha = objValeDescargaOP.Fecha
                    .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                    .CodigoAlmacen = objValeDescargaOP.CodigoAlmacen
                End With
                
                If objSqlAyudaVale.devuelveStockFisicoDeProducto >= Val(rstValeDet!CANTIDADDESTINO & "") Then
                    
                    Do While Not rstValeDet.EOF
                        .inicializarEntidadesDetalle
                        
                        .NumeroOrdenCompra = Trim(rstValeDet!NROOC & "")
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = Trim(rstValeDet!NroPedido & "")
                        
                        .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                        .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                        .Cantidad = Val(rstValeDet!CANTIDADDESTINO & "") * -1
                        
                        .ObservacionesPorItem = "SE RETIRA DE " & IIf(.Requerimiento <> vbNullString, "STOCK COMPROMETIDO DEL PEDIDO N° " & .Requerimiento, "STOCK LIBRE.")
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        'SI SE RE-DISTRIBUYE EL STOCK COMPROMETIDO; GENERAR REGISTRO DE REPOSICION DE STOCK
                        If cmbTipoStock.ListIndex = 2 Then
                            'Llenar Historial de Reposición
                            Dim lngIdReposicion As Long
                            
                            lngIdReposicion = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 IDREPOSICION", "SF1REPOSICIONCOMPROMISO", vbNullString, vbNullString, vbNullString, "IDREPOSICION > 0 ORDER BY IDREPOSICION DESC") & "") + 1
                            
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "INSERT INTO SF1REPOSICIONCOMPROMISO("
                            SqlCad = SqlCad & "IDREPOSICION, NROPEDIDO, IDINSUMO, "
                            SqlCad = SqlCad & "CANTIDAD, OBSERVACION, USUDESCOMPROMISO, "
                            SqlCad = SqlCad & "FECDESCOMPROMISO) "
                            SqlCad = SqlCad & "VALUES("
                            SqlCad = SqlCad & lngIdReposicion & ", "
                            SqlCad = SqlCad & "'" & .Requerimiento & "', "
                            SqlCad = SqlCad & "'" & .CodigoProducto & "', "
                            SqlCad = SqlCad & Abs(.Cantidad) & ", "
                            SqlCad = SqlCad & "'DESCOMPROMISO EJECUTADO CON VALE Nº " & .CodigoAlmacen & "/" & .NumeroVale & " POR " & wusuario & " PARA ATENCIÓN DE: " & .observaciones & "', "
                            SqlCad = SqlCad & "'" & wusuario & "',"
                            SqlCad = SqlCad & "CVDATE('" & Now & "'))"
                            
                            cnn_dbbancos.Execute SqlCad
                            
                            Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                        End If
                        
                        .inicializarEntidadesDetalle
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = Trim(rstValeDet!NROPEDIDODESTINO & "")
                        
                        .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                        .CodigoProductoOriginal = Trim(rstValeDet!CodProducto & "")
                        .Cantidad = Val(rstValeDet!CANTIDADDESTINO & "")
                        
                        .ObservacionesPorItem = "SE CARGA EL " & IIf(.Requerimiento <> vbNullString, "STOCK COMPROMETIDO DEL PEDIDO N° " & .Requerimiento, "STOCK LIBRE.")
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        rstValeDet.MoveNext
                    Loop
                        If dblItem > 0 Then
                            replicarRedistribucionFisica .TipoVale, .NumeroVale, .CodigoAlmacen
                        End If
                Else
                    MsgBox "Stock seleccionado no disponible al '" & dtpFechaDespacho.value & "' para atender su descarga." & vbNewLine & vbNewLine & _
                            "RECOMENDACIÓN: Vuelva a seleccionar el Producto a descargar para" & vbNewLine & _
                            "actualizar el Stock Disponible.", vbInformation + vbOKOnly, App.ProductName

                    If .eliminarVale Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        listarDetalleStockSql
                    End If
                    
                    Exit Sub
                End If
            End If
        End If
        
        Select Case cmbTipoStock.ListIndex
            Case 1, 2
                dbgResultado.Dataset.Close
                
                With objAyudaBien
                    .inicializarEntidades
                    
                    .Codigo = objValeDescargaOP.CodigoProducto
                    
                    .obtenerConfigBien
                End With
                
                'INSERTAMOS LA DESCARGA DEL ITEM SELECCIONADO
                .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", _
                                                        "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, _
                                                        "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TMPVALESALIDA("
                SqlCad = SqlCad & "ITEM, "
                SqlCad = SqlCad & "CODPROD, "
                SqlCad = SqlCad & "CODPRODORIGINAL, "
                SqlCad = SqlCad & "DESCRIPCION, "
                SqlCad = SqlCad & "UMEDIDA, "
                SqlCad = SqlCad & "CANTIDAD, "
                SqlCad = SqlCad & "CANTIDADMAX, "
                SqlCad = SqlCad & "F4NUMORD, "
                SqlCad = SqlCad & "COD_SOLICITUD"
                SqlCad = SqlCad & ") "
                
                'SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "VALUES("
                SqlCad = SqlCad & .ITEM & ", "
                SqlCad = SqlCad & "'" & .CodigoProducto & "', "
                SqlCad = SqlCad & "'" & .CodigoProductoOriginal & "', "
                SqlCad = SqlCad & "'" & objAyudaBien.Descripcion & "', "
                SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F7SIGMED", "MAESTROS.EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T") & "', "
                SqlCad = SqlCad & .Cantidad & ", "
                SqlCad = SqlCad & .Cantidad & ", "
                SqlCad = SqlCad & "'" & .NumeroOrdenCompra & "', "
                SqlCad = SqlCad & "'" & .Requerimiento & "' "
                SqlCad = SqlCad & ")"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                'ACTUALIZAMOS EL SALDO EN EL ITEM PENDIENTE SELECCIONADO PARA DESCARGA
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "SALDO = SALDO - " & .Cantidad & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CODPRODUCTOORIGEN = '" & strCodigoOriginal & "' AND "
                SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                'MsgBox "Proceso de Re-distribución y Descarga finalizada.", vbInformation + vbOKOnly, App.ProductName
                
                cmbTipoStock.ListIndex = -1
                
                listarDetalleVale
                
                listarPendiente
            Case Else
                'MsgBox "Proceso de Re-distribución finalizado.", vbInformation + vbOKOnly, App.ProductName
                
                cmbTipoStock.ListIndex = -1
                
                listarDetalleVale
                
                listarPendiente
        End Select
    End With
    
    Set objValeDescargaOP = Nothing
    
    Exit Sub
errRedistribuirStockFisicoSql:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - frmUtilDescargaOP: redistribuirStockFisicoSql"
    End Select
    
    Err.Clear
End Sub

Private Sub replicarRedistribucionFisica(ByVal strTipoVale As String, _
                                            ByVal strNumeroVale As String, _
                                            ByVal strCodigoAlmacen As String)
                                            
    On Error GoTo errReplicarRedistribucionFisica
    
    Dim rstExportarSql As New ADODB.Recordset
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = strTipoVale
        .NumeroVale = strNumeroVale
        .CodigoAlmacen = strCodigoAlmacen
        
        .obtenerConfigVale
    End With
    
    With objSqlAyudaVale
        .inicializarEntidades
        
        .TipoVale = objAyudaVale.TipoVale
        .NumeroVale = objAyudaVale.NumeroVale
        .NumeroValeExterno = objAyudaVale.NumeroValeExterno
        
        .CodigoAlmacen = objAyudaVale.CodigoAlmacen
        .CodigoOrigen = objAyudaVale.CodigoOrigen
        
        .TipoPersona = objAyudaVale.TipoPersona
            .CodigoProveedor = objAyudaVale.CodigoProveedor
            
        .CentroCosto = objAyudaVale.CentroCosto
        .SerieGuia = objAyudaVale.SerieGuia
        .NumeroGuia = objAyudaVale.NumeroGuia
        
        .CodTipoComprobante = objAyudaVale.CodTipoComprobante
        .SerieDocumento = objAyudaVale.SerieDocumento
        .NumeroDocumento = objAyudaVale.NumeroDocumento
        
        If .CodTipoComprobante <> vbNullString And .NumeroDocumento <> vbNullString Then
            .FechaUltima = Format(objAyudaVale.FechaUltima, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = objAyudaVale.OrdenTrabajo
        
        .Fecha = Format(objAyudaVale.Fecha, "Short Date")
        .CodigoMoneda = objAyudaVale.CodigoMoneda
        .TipoCambio = objAyudaVale.TipoCambio
        
        .NumeroOrdenCompra = objAyudaVale.NumeroOrdenCompra
        .observaciones = objAyudaVale.observaciones
        
        .RegistroCompra = objAyudaVale.RegistroCompra
        
        .ExportarVale = objAyudaVale.ExportarVale
        
        .FecReg = Format(objAyudaVale.FecReg, "Short Date")
        .UsuReg = objAyudaVale.UsuReg
        .FecMod = Format(objAyudaVale.FecMod, "Short Date")
        .UsuMod = objAyudaVale.UsuMod
        
        If .guardarVale Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            If rstExportarSql.State = 1 Then rstExportarSql.Close
            
            rstExportarSql.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstExportarSql.EOF Then
                rstExportarSql.MoveFirst
                
                Do While Not rstExportarSql.EOF
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstExportarSql!f5codpro & "")
                    .CodigoProductoOriginal = Trim(rstExportarSql!F5CODPROORIGINAL & "")
                    .Cantidad = Val(rstExportarSql!F3CANPRO & "")
                    .CantidadMaxima = Val(rstExportarSql!F3SALPEP & "")
                    
                    .ValorVenta = Val(rstExportarSql!F3VALVTA & "")
                    .IGV = Val(rstExportarSql!F3IGV & "")
                    .TOTAL = Val(rstExportarSql!F3TOTITE & "")
                    .ValorVentaDol = Val(rstExportarSql!F3VALDOL & "")
                    .IgvDol = Val(rstExportarSql!F3IGVDOL & "")
                    .TotalDol = Val(rstExportarSql!F3TOTDOL & "")
                    
                    .Grupo = Trim(rstExportarSql!F3GRUPO & "")
                    .ITEM = Val(rstExportarSql!F3ITEM & "")
                    .NumeroOrdenCompra = Trim(rstExportarSql!F4NUMORD & "")
                    .Requerimiento = Trim(rstExportarSql!COD_SOLICITUD & "")
                    .PorcentajeDscto = Val(rstExportarSql!F3PORCENTAJEDSCTO & "")
                    .MontoDscto = Val(rstExportarSql!F3MONTODSCTO & "")
                    
                    .ObservacionesPorItem = Trim(rstExportarSql!observaciones & "")
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    rstExportarSql.MoveNext
                Loop
            End If
        Else
            Actualiza_Log " < Replicación BD > Importación de Vale No. " & .CodigoAlmacen & " / " & .NumeroVale & " fallido.", StrConexDbBancos
        End If
    End With
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    Exit Sub
errReplicarRedistribucionFisica:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Replicacion Redistribución Fisica"
    
    Err.Clear
End Sub

Private Sub descargarItem()
    On Error GoTo errDescargarItem
    
    Dim strCodigoOriginal As String
    
    With objAyudaVale
        .inicializarEntidadesDetalle
        
        .CodigoProducto = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
        '.CodigoProductoOriginal = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOORIGEN").Value & "")
        .CodigoProductoOriginal = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
        
        strCodigoOriginal = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOORIGEN").value & "")
        
        .SQLSelectAlter = Trim(dbgDetalleStock.Columns.ColumnByFieldName("NOMPRODUCTO").value & "")
        .NumeroOrdenCompra = Trim(dbgDetalleStock.Columns.ColumnByFieldName("NROOC").value & "")
        .Requerimiento = Trim(dbgDetalleStock.Columns.ColumnByFieldName("NROPEDIDO").value & "")
        
        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALESALIDA", "CODPROD", .CodigoProductoOriginal, "T", "AND TRIM(F4NUMORD & '') = '" & .NumeroOrdenCompra & "' AND TRIM(COD_SOLICITUD & '') = '" & .Requerimiento & "'")) > 0 Then
            
            MsgBox "Item ya descargado al Vale, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            .inicializarEntidadesDetalle
            
            With dbgDetalleStock
                .Dataset.Edit
                
                .Columns.ColumnByFieldName("DESCARGAR").value = False
                
                .Dataset.Post
            End With
            
            Exit Sub
        End If
        
'        If MsgBox("¿Desea descargar el Item seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'            .inicializarEntidadesDetalle
'
'            With dbgDetalleStock
'                .Dataset.Edit
'
'                .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
'                .Columns.ColumnByFieldName("DESCARGAR").value = False
'
'                .Dataset.Post
'            End With
'
'            Exit Sub
'        End If
        
        'VARIABLES DE CANTIDADES RESULTANTES
        Dim dblCantidadDescarga As Double
        Dim dblSaldoPendiente As Double
        Dim dblSaldoDisponible As Double
        
        'CALCULAR CANTIDADES RESULTANTES --------------------------------------------------------------
        dblCantidadDescarga = Val(dbgDetalleStock.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "")
        
        If Val(dbgDetalleStock.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") <= _
            Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "") Then
            
            dblSaldoPendiente = Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "") - _
                                Val(dbgDetalleStock.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "")
        Else
            dblSaldoPendiente = 0
        End If
        
        dblSaldoDisponible = Val(dbgDetalleStock.Columns.ColumnByFieldName("CANTIDAD").value & "") - dblCantidadDescarga
        '---------------------------------------------------------------------------------------------
        
        dbgResultado.Dataset.Close
        
        With objAyudaBien
            .inicializarEntidades
            
            .Codigo = objAyudaVale.CodigoProducto
            
            .obtenerConfigBien
        End With
        
        'INSERTAMOS LA DESCARGA DEL ITEM SELECCIONADO
        .ITEM = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", _
                                                "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, _
                                                "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "INSERT INTO TMPVALESALIDA("
        SqlCad = SqlCad & "ITEM, "
        SqlCad = SqlCad & "CODPROD, "
        SqlCad = SqlCad & "CODPRODORIGINAL, "
        SqlCad = SqlCad & "DESCRIPCION, "
        SqlCad = SqlCad & "UMEDIDA, "
        SqlCad = SqlCad & "CANTIDAD, "
        SqlCad = SqlCad & "CANTIDADMAX, "
        SqlCad = SqlCad & "F4NUMORD, "
        SqlCad = SqlCad & "COD_SOLICITUD"
        SqlCad = SqlCad & ") "
        
        SqlCad = SqlCad & "VALUES("
        SqlCad = SqlCad & .ITEM & ", "
        SqlCad = SqlCad & "'" & .CodigoProducto & "', "
        SqlCad = SqlCad & "'" & .CodigoProductoOriginal & "', "
        SqlCad = SqlCad & "'" & objAyudaBien.Descripcion & "', "
        SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F7SIGMED", "MAESTROS.EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T") & "', "
        SqlCad = SqlCad & dblCantidadDescarga & ", "
        SqlCad = SqlCad & dblCantidadDescarga & ", "
        SqlCad = SqlCad & "'" & .NumeroOrdenCompra & "', "
        SqlCad = SqlCad & "'" & .Requerimiento & "'"
        SqlCad = SqlCad & ")"
        
        abrirCnTemporal
        
        cnDBTemp.Execute SqlCad
        
        'ACTUALIZAMOS EL SALDO EN EL ITEM PENDIENTE SELECCIONADO PARA DESCARGA
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "SALDO = " & dblSaldoPendiente & " "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CODPRODUCTOORIGEN = '" & strCodigoOriginal & "' AND "
        SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
        
        abrirCnTemporal
        
        cnDBTemp.Execute SqlCad
        
        Rem SK: SE DESHABILITA OPCION DE LIBERAR STOCK COMPROMETIDO, YA QUE ESTE SE REALIZARA MASIVAMENTE A TRAVES DE LA OPCION "REGULARIZACION DE STOCK CEA"
'        'EVALUAMOS EL SALDO DISPONIBLE DEL STOCK COMPROMETIDO, CON EL FIN DE DARLE OPCION AL USUARIO DE LIBERARLA PARA SU POSTERIOR USO
'        If dblSaldoDisponible > 0 Then
'            If MsgBox("Resta un saldo de " & Format(dblSaldoDisponible, "#0.00") & " " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "M.F7SIGMED", "IF5PLA AS P LEFT JOIN EF7MEDIDAS AS M ON M.F7CODMED = P.F7CODMED", "F5CODPRO", .CodigoProducto, "T") & " del insumo seleccionado." & vbNewLine & _
'                        "¿Desea liberar este saldo (a stock libre) para su posterior uso?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'
'                .inicializarEntidades
'
'                .CodigoAlmacen = right(Trim(cmbAlmacen.Text), 2)
'                .NumeroVale = vbNullString
'                .TipoVale = "I"
'
'                .Fecha = Format(dtpFechaDespacho.value, "Short Date")
'                .CodigoOrigen = "XCS"
'                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
'
'                    If .TipoCambio = 0 Then
'                        .TipoCambio = "2.8"
'                    End If
'
'                .CodigoMoneda = "S"
'
'                .referencia = wnomcia
'                .observaciones = "REDISTRIBUCION DE SCTOCK AUTOMATICO."
'
'                .FecReg = Format(Date, "Short Date")
'                .UsuReg = wusuario
'                .FecMod = Format(Date, "Short Date")
'                .UsuMod = wusuario
'
'                ''cnn_dbbancos.BeginTrans
'
'                'bolTransaccionEnCurso = True
'
'                If .guardarVale Then
'                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'
'                    'Borrar Detalle de Vale
'                    SqlCad = vbNullString
'                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
'
'                    cnn_dbbancos.Execute SqlCad
'                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
'
'                    '.inicializarEntidadesDetalle
'
'                    '.NumeroOrdenCompra = Trim(rstValeDet!NROOC & "")
'                    '.Requerimiento = Trim(rstValeDet!NROPEDIDO & "")
'
'                    '.CodigoProducto = Trim(rstValeDet!CodProducto & "")
'                    .CodigoProductoOriginal = .CodigoProducto  'Trim(rstValeDet!CodProducto & "")
'                    .Cantidad = dblSaldoDisponible * -1  'Val(rstValeDet!CANTIDADDESTINO & "") * -1
'                    .ITEM = 1
'
'                    If .verificarStockProductoFisicoCorL(.CodigoProducto, _
'                                                        .CodigoAlmacen, _
'                                                        .NumeroOrdenCompra, _
'                                                        .Requerimiento, _
'                                                        dblSaldoDisponible, _
'                                                        dtpFechaDespacho.value) Then
'                        .guardarValeDetalleOneByOne
'
'                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'
'                        '.inicializarEntidadesDetalle
'
'                        .Requerimiento = vbNullString  'Trim(rstValeDet!NROPEDIDODESTINO & "")
'
'                        '.CodigoProducto = Trim(rstValeDet!CodProducto & "")
'                        .CodigoProductoOriginal = .CodigoProducto 'Trim(rstValeDet!CodProducto & "")
'                        .Cantidad = dblSaldoDisponible 'Val(rstValeDet!CANTIDADDESTINO & "")
'                        .ITEM = 2
'
'                        .guardarValeDetalleOneByOne
'
'                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'                    Else
'                        MsgBox "Saldo de Compromiso no disponible al '" & dtpFechaDespacho.value & "'.", vbInformation + vbOKOnly, App.ProductName
'
'                        If .eliminarVale Then
'                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'                        End If
'
'                        Exit Sub
'                    End If
'                End If
'
'                ''cnn_dbbancos.CommitTrans
'            End If
'        End If
    End With
    
    strCodigoOriginal = vbNullString
    
    listarDetalleVale
    
    listarPendiente
    
    cmbAlmacen.Enabled = False
    
    dbgPendiente.SetFocus
    
    Exit Sub
errDescargarItem:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - frmUtilDescargaOP: descargarItem"
    End Select
    'Resume
    Err.Clear
End Sub

Private Sub devolverItemDescargado()
    On Error GoTo errDevolverItemDescargado
    
    With objAyudaVale
        .inicializarEntidadesDetalle
        
        .CodigoProducto = Trim(dbgResultado.Columns.ColumnByFieldName("CODPROD").value & "")
        .Cantidad = Val(dbgResultado.Columns.ColumnByFieldName("CANTIDAD").value & "")
        
        If MsgBox("¿Desea retornar el Item descargado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            dbgResultado.Dataset.Close
            
            'ACTUALIZAMOS EL SALDO EN EL ITEM PENDIENTE SELECCIONADO PARA DESCARGA
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILDESCARGAOPPENDIENTE "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "SALDO = SALDO + " & .Cantidad & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CODPRODUCTOFINAL = '" & .CodigoProducto & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "DELETE "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "TMPVALESALIDA "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CODPROD = '" & .CodigoProducto & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            listarDetalleVale
            
            listarPendiente
        End If
        
        .inicializarEntidadesDetalle
    End With
    
    dbgPendiente.SetFocus
    
    Exit Sub
errDevolverItemDescargado:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - frmUtilDescargaOP: devolverItemDescargado"
    End Select
    'Resume
    Err.Clear
End Sub

Private Sub cmbAlmacen_Click()
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        listarDetalleStockSql
    Else
        listarDetalleStock
    End If
End Sub

Private Sub cmbalmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbTipoStock_Click()
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        listarDetalleStockSql
    Else
        listarDetalleStock
    End If
End Sub

Private Sub cmbTipoStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub CmdCancelar_Click()
    cmbAlmacen.Enabled = True
    
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    cmbAlmacen.Enabled = False
    
    Me.Hide
End Sub

Private Sub dbgDetalleStock_GotFocus()
    'dbgDetalleStock.Columns.FocusedIndex = dbgDetalleStock.Columns.ColumnByFieldName("OBSERVACIONES").ColIndex
End Sub

Private Sub dbgDetalleStock_OnDblClick()
'    Select Case dbgDetalleStock.Columns.FocusedColumn.FieldName
'        Case "NROOC", "NROPEDIDO", "OBSERVACIONES", "CANTIDAD", "UM"
'            If left(cmbTipoStock.Text, 1) = "C" And dbgDetalleStock.Dataset.RecordCount > 0 Then
'                With dbgDetalleStock
'                    .Dataset.Edit
'
'                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") = 0 Then
'                        If Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "") >= Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").Value & "") Then
'                            .Columns.ColumnByFieldName("CANTIDADDESTINO").Value = Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").Value & "")
'                        Else
'                            .Columns.ColumnByFieldName("CANTIDADDESTINO").Value = Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "")
'                        End If
'
'                        .Columns.ColumnByFieldName("DESCARGAR").Value = CBool(.Columns.ColumnByFieldName("DESCARGAR").Value)
'                    End If
'
'                    .Dataset.Post
'                End With
'
'                descargarItem
'            Else
'                With dbgDetalleStock
'                    .Dataset.Edit
'
'                    .Columns.ColumnByFieldName("PROCESAR").Value = True
'
'                    .Dataset.Post
'
'                    'dbgDetalleStock.Columns.FocusedColumn.FieldName = dbgDetalleStock.Columns.ColumnByFieldName("PROCESAR").FieldName
'
'                    'dbgDetalleStock_OnEdited Nothing
'
'                    .Dataset.Edit
'
'                    dbgDetalleStock_OnCheckEditToggleClick dbgDetalleStock.Columns.ColumnByFieldName("PROCESAR"), Nothing, vbNullString, cbsChecked
'                End With
'            End If
'    End Select
End Sub

Private Sub dbgDetalleStock_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
'    Select Case KeyCode
'        Case vbKeyReturn
'            dbgDetalleStock_OnDblClick
'    End Select
End Sub

Private Sub dbgPendiente_GotFocus()
    'dbgPendiente.Columns.FocusedIndex = dbgPendiente.Columns.ColumnByFieldName("NOMPRODUCTO").ColIndex
End Sub

Private Sub dbgPendiente_OnChangeNodeEx()
'    cmbTipoStock.ListIndex = 0
'
'    listarDetalleStock
    On Error Resume Next
    
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    'Dim i As Integer
    
    Select Case UCase(dbgPendiente.Columns.FocusedColumn.FieldName)
        Case "NOMPRODUCTO"
            cmbTipoStock.ListIndex = -1
            
'            If Val(dbgPendiente.Columns.ColumnByFieldName("STOCK").value & "") <= 0 Then
'                cmbTipoStock.Enabled = False
'
'                dbgDetalleStock.Dataset.Close
'
'                Exit Sub
'            Else
'                cmbTipoStock.Enabled = True
'            End If
            
'
'            i = 0
'
'            For i = 0 To 2
'                DoEvents
'
'                cmbTipoStock.ListIndex = i
'
'
'                'cmbTipoStock_Click
'                'listarDetalleStock
'
'                If dbgDetalleStock.Dataset.RecordCount > 0 Then
'                    dbgDetalleStock.SetFocus
'
'                    'dbgDetalleStock.Columns.FocusedIndex = 9
'
'                    Exit For
'                End If
'            Next i
            
'''            With objAyudaVale
'''                .inicializarEntidades
'''
'''                .CodigoAlmacen = right(cmbAlmacen.Text, 2)
'''                .CodigoProducto = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
'''
'''                'abrirCnnDbBancos
'''
'''                Me.MousePointer = vbHourglass
'''
'''                If .devuelveStockFisicoDeProducto(vbNullString, True) <= 0 Then
'''                    Me.MousePointer = vbDefault
'''
'''                    cmbTipoStock.Enabled = False
'''
'''                    dbgDetalleStock.Dataset.Close
'''
'''                    Exit Sub
'''                Else
'''                    cmbTipoStock.Enabled = True
'''                End If
'''
'''                Me.MousePointer = vbDefault
'''
'''                .inicializarEntidades
'''            End With
            
            cmbTipoStock.ListIndex = 0
            
            If dbgDetalleStock.Dataset.RecordCount > 0 Then
                dbgDetalleStock.SetFocus
                
                Exit Sub
            End If
            
            cmbTipoStock.ListIndex = 1
            
            If dbgDetalleStock.Dataset.RecordCount > 0 Then
                dbgDetalleStock.SetFocus
                
                Exit Sub
            End If
            
            cmbTipoStock.ListIndex = 2
            
            If dbgDetalleStock.Dataset.RecordCount > 0 Then
                dbgDetalleStock.SetFocus
                
                Exit Sub
            End If
    End Select
End Sub

Private Sub dbgDetalleStock_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case UCase(Column.FieldName)
        Case "PROCESAR"
            If dbgDetalleStock.Dataset.State = dsEdit Then
                dbgDetalleStock.Dataset.Post

                'If left(cmbTipoStock.Text, 1) = "C" And dbgDetalleStock.Dataset.RecordCount > 0 Then
                '    descargarItem
                'Else
                    
                    'VALIDAR SI STOCK SELECCIONADO ESTA DISPONIBLE
                    
                        
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        redistribuirStockFisicoSql
                    Else
                        redistribuirStockFisico
                    End If
                    
''                    If left(cmbTipoStock.Text, 1) = "C" Then
''                        listarDetalleStock
''                    Else
''                        cmbTipoStock.ListIndex = 0
''                    End If
''
''                    If dbgDetalleStock.Dataset.RecordCount = 0 Then cmbTipoStock.SetFocus
                'End If
            End If
'            If dbgDetalle.Dataset.State = dsEdit Then
'                dbgDetalle.Dataset.Post
'            End If
        Case "DESCARGAR"
            If dbgDetalleStock.Dataset.State = dsEdit Then
                dbgDetalleStock.Dataset.Post
                
                If left(cmbTipoStock.Text, 1) = "C" And dbgDetalleStock.Dataset.RecordCount > 0 Then
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        descargarItem
                    Else
                        descargarItem
                    End If
                End If
            End If
    End Select
End Sub

Private Sub dbgDetalleStock_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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
    
    'dbgDetalleStock.Columns.ColumnByFieldName("PROCESAR").Visible = True
    'dbgDetalleStock.Columns.ColumnByFieldName("DESCARGAR").Visible = IIf(Mid(Trim(cmbTipoStock.Text), 1, 1) = "C", True, False)
End Sub

Private Sub dbgDetalleStock_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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
    
    'dbgDetalleStock.Columns.ColumnByFieldName("PROCESAR").Visible = True
    'dbgDetalleStock.Columns.ColumnByFieldName("DESCARGAR").Visible = IIf(Mid(Trim(cmbTipoStock.Text), 1, 1) = "C", True, False)
End Sub

Private Sub dbgDetalleStock_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case dbgDetalleStock.Columns.FocusedColumn.FieldName
        Case "NROPEDIDODESTINO"
            With dbgDetalleStock
                If .Dataset.State = dsEdit Then
                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & ""), "N") = vbNullString Then
                        '.Dataset.Cancel
                        
                        .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                        
                        .Dataset.Post
                        
                        Exit Sub
                    End If

                    .Dataset.Post
                End If

                .Dataset.Edit

                .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & ""), "N")

                .Dataset.Post
            End With
        Case "CANTIDADDESTINO"
            With dbgDetalleStock
                If .Dataset.State = dsEdit Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") Then
                        MsgBox "La cantidad de re-distribución no puede ser mayor al stock disponible, verifique.", vbInformation + vbOKOnly, App.ProductName

                        .Dataset.Cancel

                        Exit Sub
                    End If

'                    If left(cmbTipoStock.Text, 1) = "L" Then
'                        If Trim(.Columns.ColumnByFieldName("NROPEDIDODESTINO").value & "") = vbNullString Then
'                            MsgBox "El stock libre no puede ser re-distribuido sin consignar el No. Pedido Destino, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                            .Dataset.Cancel
'
'                            Exit Sub
'                        End If
'                    End If
                    
                    .Dataset.Post
                End If

                
                
                'If left(cmbTipoStock.Text, 1) = "L" Then
                    .Dataset.Edit
                    
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                    
                    .Dataset.Post
                'End If
                
'                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") <= 0 Then
'                    .Columns.ColumnByFieldName("PROCESAR").Value = False
'                Else
'                    .Columns.ColumnByFieldName("PROCESAR").Value = True
'                End IF
            End With
        Case "PROCESAR"
            With dbgDetalleStock
                If .Dataset.State = dsEdit Then
                    .Dataset.Post
                End If
                
                .Dataset.Edit
                
                If left(cmbTipoStock.Text, 1) = "L" Or left(cmbTipoStock.Text, 1) = "O" Then
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                Else
                    .Columns.ColumnByFieldName("NROPEDIDODESTINO").value = vbNullString
                End If
                
                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                    If Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") >= Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "") Then
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "")
                    Else
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                    End If
                End If
                
                .Dataset.Post
            End With
        Case "DESCARGAR"
            With dbgDetalleStock
                If .Dataset.State = dsEdit Then
                    .Dataset.Edit
                    
                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                        If Val(.Columns.ColumnByFieldName("CANTIDAD").value & "") >= Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "") Then
                            .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "")
                        Else
                            .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                        End If
        
                        .Columns.ColumnByFieldName("DESCARGAR").value = CBool(.Columns.ColumnByFieldName("DESCARGAR").value)
                    End If
                    
                    .Dataset.Post
                End If
            End With
    End Select
End Sub

Private Sub dbgPendiente_OnClick()
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    If dbgPendiente.Dataset.RecordCount > 0 Then
        dbgPendiente_OnChangeNodeEx
    End If
End Sub

Private Sub dbgPendiente_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "CODPRODUCTOFINAL"
            If Trim(Text) <> Node.Values(3) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
        Case "CANTIDADFINAL"
            If Val(Text) <> Val(Node.Values(7)) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "CANTIDADORIGEN", "SALDO", "STOCK"
            If Val(Text) < 0 Then
                FontColor = vbRed
            ElseIf Val(Text) = 0 Then
                FontColor = vbGreen
            Else
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgPendiente_OnDblClick()
'    If dbgDetalleStock.Dataset.RecordCount > 0 Then
'        dbgDetalleStock.SetFocus
'    Else
'        cmbTipoStock.SetFocus
'    End If
End Sub

Private Sub dbgPendiente_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case Column.FieldName
        Case "CODPRODUCTOFINAL"
            If Val(dbgPendiente.Columns.ColumnByFieldName("CANTIDADFINAL").value & "") <> Val(dbgPendiente.Columns.ColumnByFieldName("SALDO").value & "") Then
                MsgBox "Imposible reemplazar el Producto, ya fue descargado.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Rem SK: SE DESHABILITA OPCION DE LIBERAR STOCK COMPROMETIDO, YA QUE ESTE SE REALIZARA MASIVAMENTE A TRAVES DE LA OPCION "REGULARIZACION DE STOCK CEA"
'            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILSTOCKDETALLE", "CODPRODUCTO", _
'                                                Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), "T", _
'                                                "AND CODALMACEN = '" & right(cmbAlmacen.Text, 2) & "' AND NROPEDIDO = '" & _
'                                                Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "'")) > 0 Then
'
'                If MsgBox("El Producto cuenta actualmente con Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
'                            "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'
'                    Exit Sub
'                End If
'            End If
            
            
            
            
'            If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                Unload ayuda_productos
'            End If
'
'            With ayuda_productos
'                .CodigoAuxiliar = vbNullString
'                .CodigoRequerimiento = vbNullString
'                .CodigoProducto = vbNullString
'
'                '.txtBusqueda.Text = Trim(dbgPendiente.Columns.ColumnByFieldName("NOMPRODUCTO").value & "")
'                .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
'
'                .listarProductos
'
'                .Show 1
'            End With
'
'            abrirCnTemporal
'
'            objAyudaBien.Codigo = ModUtilitario.ObtenerCampoV2(cnDBTemp, "F5CODPRO", "TMPPRODUCTOS", "F4PERINT", "-1", "N")
            
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                objAyudaBien.Codigo = Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & "")
                
                objAyudaBien.obtenerConfigBien
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True
                '.CadenaCorte = objAyudaBien.Modelo
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = objAyudaBien.Modelo 'vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPRODUCTOFINAL", "TMPUTILDESCARGAOPPENDIENTE", "CODPRODUCTOFINAL", objAyudaBien.Codigo, "T") <> vbNullString Then
                        MsgBox "Imposible realizar el cambio; el producto:" & vbNewLine & _
                                objAyudaBien.Descripcion & ", " & vbNewLine & _
                                "Se encuentra consignado en la actual OP. Se sugiere:" & vbNewLine & _
                                "1) Sumar la Cantidad del producto a cambiar, al que ya existe." & vbNewLine & _
                                "2) En caso que el producto a cambiar cuente con Stock Comprometido, proceder a Liberarlo." & vbNewLine & _
                                "3) Anular (Desestimar) el producto a cambiar, llevando su Cantidad Final a cero (0)." & vbNewLine & _
                                "4) Proceder con la Descarga del Producto ya existente, que cuenta con la suma de ambos.", vbInformation + vbOKOnly, App.ProductName
                        
                        objAyudaBien.inicializarEntidades
                        
                        Exit Sub
                    End If
                    
                    If Trim(dbgPendiente.Columns.ColumnByFieldName("UM").value & "") <> Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")) Then
                        If MsgBox("El Producto seleccionado para el cambio cuenta con diferente Unidad de Medida (U.M.)." & vbNewLine & _
                                    "¿Desea continuar con la acción?", vbInformation + vbYesNo, App.ProductName) = vbNo Then
                            
                            objAyudaBien.inicializarEntidades
                            
                            Exit Sub
                        End If
                    End If
                    
                    If MsgBox("¿Desea efectuar el reemplazo del Producto?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        Exit Sub
                    End If
                    
                    With dbgPendiente
                        If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("NROOP").value), _
                                                            Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value), _
                                                            objAyudaBien.Codigo, _
                                                            Val(.Columns.ColumnByFieldName("CANTIDADORIGEN").value), _
                                                            Val(.Columns.ColumnByFieldName("CANTIDADFINAL").value), "DESCARGA DE OP - CAMBIO DE PRODUCTO") Then
                               
                            MsgBox "Producto reemplazado en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").value), vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Edit
                            
                            .Columns.ColumnByFieldName("CODPRODUCTOFINAL").value = objAyudaBien.Codigo
                            .Columns.ColumnByFieldName("NOMPRODUCTO").value = objAyudaBien.Descripcion
                            .Columns.ColumnByFieldName("UM").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                            
                            .Dataset.Post
                        End If
                        
                        dbgPendiente_OnChangeNodeEx
                    End With
                End If
            End With
    End Select
End Sub

Private Sub dbgPendiente_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
'     With .Dataset
'        If .Columns.FocusedColumn.ColumnType = gedLookupEdit Then
'            If .State = dsEdit Then
'                .m.HideEditor
'                .Post
'                .DisableControls
'                .Close
'                .Open
'                .EnableControls
'            End If
'        End If
'    End With
    
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case dbgPendiente.Columns.FocusedColumn.FieldName
        Case "CANTIDADFINAL"
            With dbgPendiente
                If .Dataset.State = dsEdit Then
                    Dim dblCantidadOrigen As Double
                    
                    dblCantidadOrigen = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDADFINAL", "TMPUTILDESCARGAOPPENDIENTE", "CODPRODUCTOFINAL", Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), "T"))
                    
                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILSTOCKDETALLE", "CODPRODUCTO", _
                                                        Trim(dbgPendiente.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), "T", _
                                                        "AND CODALMACEN = '" & right(cmbAlmacen.Text, 2) & "' AND NROPEDIDO = '" & _
                                                        Trim(dbgPendiente.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "'")) > 0 Then
                                                        
                        If MsgBox("El Producto cuenta actualmente con Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                    "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    'If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDADFINAL", "TMPUTILDESCARGAOPPENDIENTE", "CODPRODUCTOFINAL", Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value & ""), "T"))  Then
                    If MsgBox("¿Desea aplicar el Ajuste de Cantidad?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        .Dataset.Cancel
                        
                        Exit Sub
                    Else
                        .Dataset.Post
                    End If
                    
                    If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("NROOP").value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTOFINAL").value), _
                                                        Val(.Columns.ColumnByFieldName("SALDO").value), _
                                                        Val(.Columns.ColumnByFieldName("CANTIDADFINAL").value), "DESCARGA DE OP - AJUSTE DE CANTIDAD DE PRODUCTO") Then
                        
                        MsgBox "Efectuado ajuste de cantidad de Producto en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").value), vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("SALDO").value = Val(.Columns.ColumnByFieldName("CANTIDADFINAL").value & "") - (dblCantidadOrigen - Val(.Columns.ColumnByFieldName("SALDO").value & ""))
                        
                        .Dataset.Post
                    Else
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("CANTIDADFINAL").value = Val(.Columns.ColumnByFieldName("SALDO").value & "")
                        
                        .Dataset.Post
                    End If
                    
                    dblCantidadOrigen = 0
                    
                    dbgPendiente_OnChangeNodeEx
                End If
            End With
    End Select
End Sub

Private Sub dbgPendiente_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
'    Select Case KeyCode
'        Case vbKeyReturn
'            dbgPendiente_OnDblClick
'    End Select
End Sub

Private Sub dbgResultado_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "CODPROD"
            If Trim(Text) <> Node.Values(4) Then
                Font.Bold = True
                FontColor = RGB(255, 255, 255)
                Color = RGB(217, 151, 149)
            Else
                Font.Bold = True
                FontColor = RGB(156, 101, 0)
                Color = RGB(255, 235, 156)
            End If
        Case "CANTIDAD", "SALDO"
            If Val(Text) < 0 Then
                FontColor = vbRed
            ElseIf Val(Text) = 0 Then
                FontColor = vbGreen
            Else
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgResultado_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Not IsDate(dtpFechaDespacho.value) Then
        MsgBox "Ingrese la Fecha de Despacho.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFechaDespacho.SetFocus
        
        Exit Sub
    End If
    
    Select Case dbgResultado.Columns.FocusedColumn.Caption
        Case "Eliminar"
            'MsgBox "Aqui! :)"
            devolverItemDescargado
    End Select
End Sub

Private Sub dtpFechaDespacho_GotFocus()
    'MsgBox "Sigueme"
End Sub

Private Sub Form_Activate()
    If Not bolOPCargada Then
        bolOPCargada = True
        
        obtenerStockProductoPendiente
        
        obtenerPendiente
        
        Screen.MousePointer = vbDefault
        
        cmbTipoStock.ListIndex = 0
        
        listarDetalleVale
        
        dtpFechaDespacho.value = Empty
    End If
End Sub

Private Sub Form_Load()
    bolOPCargada = False
    
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    dtpFechaDespacho.value = Date
    
    If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaUsarFechaPredeterminada", "l") = "1" Then
        dtpFechaDespacho.value = ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaFechaPredeterminada", "l")
    End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA"
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilDescargaOrdenProduccion
    
    configurarGrilla
    
    cmbAlmacen.Enabled = True
    
    listarAlmacenEnCombo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgPendiente.Dataset.Close
    dbgDetalleStock.Dataset.Close
    dbgResultado.Dataset.Close
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    Dim dblPromH As Double
    Dim dblPromW As Double
    
    dblPromH = Me.ScaleHeight / 10
    dblPromW = Me.ScaleWidth / 2
    
    fraPendiente.Move 0, 0, Me.ScaleWidth, dblPromH * 4
        dbgPendiente.Move 100, 200, fraPendiente.Width - 200, fraPendiente.Height - 300
        
    fraDetalleStock.Move 0, fraPendiente.top + fraPendiente.Height, Me.ScaleWidth, dblPromH * 3
        dbgDetalleStock.Move 100, 600, fraDetalleStock.Width - 200, fraDetalleStock.Height - 700
        
    fraResultado.Move 0, fraDetalleStock.top + fraDetalleStock.Height, Me.ScaleWidth, dblPromH * 3
        dbgResultado.Move 100, 200, fraResultado.Width - 200, fraResultado.Height - 600
        cmdCancelar.Move dbgResultado.Width - cmdCancelar.Width, (dbgResultado.top + dbgResultado.Height) + 50, cmdCancelar.Width, 290
        cmdOk.Move (cmdCancelar.left - cmdOk.Width) - 200, cmdCancelar.top, cmdCancelar.Width, cmdCancelar.Height
End Sub


