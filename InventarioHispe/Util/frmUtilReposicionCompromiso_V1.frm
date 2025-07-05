VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilReposicionCompromiso 
   Caption         =   "Reposición de Compromisos Afectados"
   ClientHeight    =   8310
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   13440
   Icon            =   "frmUtilReposicionCompromiso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   13440
   Begin VB.Frame fraBusqueda 
      Caption         =   " Ingresar cadena a buscar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   9375
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   25
         Scrolling       =   1
      End
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin VB.Frame fraOpciones 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox chkProductoSeleccionado 
         Caption         =   "Mostrar productos seleccionados."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chkProductoProveedor 
         Caption         =   "Mostrar productos de proveedor."
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbReposicion 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilReposicionCompromiso.frx":058A
      ToolBars        =   "frmUtilReposicionCompromiso.frx":B6A1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReposicion 
      Height          =   6825
      Left            =   120
      OleObjectBlob   =   "frmUtilReposicionCompromiso.frx":B838
      TabIndex        =   6
      Top             =   1080
      Width           =   13185
   End
   Begin VB.Frame fraAlmacen 
      Caption         =   " Almacen "
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
      Left            =   6840
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmUtilReposicionCompromiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNroPedido            As String
Private strCodProveedor         As String
Private strCodProducto         As String

Public Property Let NroPedido(ByVal value As String)
    strNroPedido = value
End Property

Public Property Get NroPedido() As String
    NroPedido = strNroPedido
End Property

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

Private Sub listarAlmacenEnCombo()
    Dim rstAlmacen As New ADODB.Recordset

    If rstAlmacen.State = 1 Then rstAlmacen.Close

    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    cmbAlmacen.Clear

    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst

        Do While Not rstAlmacen.EOF
            cmbAlmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")

            rstAlmacen.MoveNext
        Loop
            If cmbAlmacen.ListCount > 0 Then
                cmbAlmacen.ListIndex = 0
            End If
    End If
End Sub

Public Sub cargarResumenRequerimiento()
    Screen.MousePointer = vbHourglass
    
    dbgReposicion.Dataset.Close
    
    txtbusqueda.Text = vbNullString
    
    objAyudaSolicitud.listarGrillaResumenRequerimiento dbgReposicion, Nothing, strNroPedido, strCodProducto
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cargarStockDeProducto(Optional ByVal strNroPedido As String, _
                                    Optional ByVal strCodProducto As String)
    On Error GoTo errCargarStockProducto
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgReposicion.Dataset.Close
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "NROPEDIDO = '" & strNroPedido & "' AND "
            SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        pgbProgresoBusqueda.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProgresoBusqueda.value = 0
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            With objAyudaVale
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                .verificarStockProducto
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstTemporal!NroPedido & "") & "' AND "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                
                cnDBTemp.Execute SqlCad
                
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
            End With
            
            pgbProgresoBusqueda.value = pgbProgresoBusqueda.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    Exit Sub
errCargarStockProducto:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub cargarProductoAtendidoPorProveedor()
    On Error GoTo errCargarProductoAtendidoPorProveedor
    
    Dim rstTemporal As New ADODB.Recordset
    Dim bolAtendidoPorProveedor As Boolean
    
    dbgReposicion.Dataset.Close
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        pgbProgresoBusqueda.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProgresoBusqueda.value = 0
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            With objAyudaVale
                .CodigoProveedor = strCodProveedor
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                bolAtendidoPorProveedor = .verificarProductoPorProveedor
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "ATENDIDOPORPROV = " & IIf(bolAtendidoPorProveedor, "TRUE", "FALSE") & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstTemporal!NroPedido & "") & "' AND "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                
                cnDBTemp.Execute SqlCad
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            pgbProgresoBusqueda.value = pgbProgresoBusqueda.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    Exit Sub
errCargarProductoAtendidoPorProveedor:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub listarResumenRequerimiento()
    dbgReposicion.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPUTILRESUMENREQUERIMIENTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
        If CBool(chkProductoSeleccionado.value) Then
            SqlCad = SqlCad & "AND PROCESAR = TRUE "
        End If
        
        If CBool(chkProductoProveedor.value) Then
            SqlCad = SqlCad & "AND ATENDIDOPORPROV = TRUE "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    With dbgReposicion
        abrirCnTemporal

        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString

        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "LLAVE"
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
    End With
End Sub

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    dbgReposicion.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "CANTIDADPC = " & IIf(bolEstado, "SALDOACTUAL", "0") & ", "
    SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "TRUE", "FALSE") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' AND "
    SqlCad = SqlCad & "VAL(SALDOACTUAL & '') > 0 "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarResumenRequerimiento
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub chkProductoProveedor_Click()
    listarResumenRequerimiento
End Sub

Private Sub chkProductoSeleccionado_Click()
    listarResumenRequerimiento
End Sub

Private Sub cmbAlmacen_Click()
    'cargarResumenRequerimiento
End Sub

Private Sub dbgReposicion_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            With dbgReposicion
                If Val(.Columns.ColumnByFieldName("SALDOACTUAL").value & "") <= 0 Then
                    MsgBox "Producto sin saldo pendiente, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                If .Dataset.State = dsEdit Then
                    .Dataset.Post
                End If
                
                .Dataset.Edit
                
                If CBool(.Columns.ColumnByFieldName("PROCESAR").value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADPC").value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADPC").value = Val(.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                    End If
                Else
                    .Columns.ColumnByFieldName("CANTIDADPC").value = 0
                End If
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgReposicion_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
'    Select Case Column.FieldName
'        Case "NROPEDIDO"
'            If Trim(Text) = vbNullString Then
'                Text = "Stock Libre"
'            End If
'
'            Font.Bold = True
'            FontColor = vbWhite
'        Case "SALDOACTUAL", "CANTIDADPC"
'            Text = Format(Text, "#,0.00;(#,0.00)")
'    End Select
    Select Case UCase(Column.FieldName)
        Case "SALDOACTUAL"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbRed
                Color = vbYellow
            ElseIf Val(Text) < 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = vbRed
            Else
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(128, 128, 128)
            End If
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
        Case "COMPROMISOEAG", "COMPROMISOPLG", "LIBREEAG", "LIBREPLG", "STOCKEAG", "STOCKPLG"
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
    End Select
End Sub

Private Sub dbgReposicion_OnDblClick()
    Select Case dbgReposicion.Columns.FocusedColumn.FieldName
        Case "CODPRODUCTO", "NOMPRODUCTO"
            If Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "") <> vbNullString Then
                With frmUtilDetalleConsolidadoPedido
                    .NroPedido = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                    .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                    
                    .Show 1
                End With
            End If
        Case "COMPROMISOEA"
            If Val(dbgReposicion.Columns.ColumnByFieldName("COMPROMISOEA").value & "") = 0 Then
                MsgBox "Cantidad Comprometida en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                dbgReposicion.M.FullRefresh
                
                Exit Sub
            End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "F"
                .NroPedido = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                
                .Show vbModal
            End With
        Case "COMPROMISOPL"
            If Val(dbgReposicion.Columns.ColumnByFieldName("COMPROMISOPL").value & "") = 0 Then
                MsgBox "Cantidad por Llegar en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "V"
                .NroPedido = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                
                .Show vbModal
            End With
        Case "UM", "SALDOACTUAL"
            If Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") <= 0 Then
                MsgBox "Imposible seleccionar Item, ya se encuentra completamente atendido, verifique 'Saldo'.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With dbgReposicion
                .Dataset.Edit
                
                If Not CBool(dbgReposicion.Columns.ColumnByFieldName("PROCESAR").value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADPC").value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADPC").value = Val(.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                    End If
                Else
                    .Columns.ColumnByFieldName("CANTIDADPC").value = 0
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = IIf(Not CBool(dbgReposicion.Columns.ColumnByFieldName("PROCESAR").value), True, False)
                
                .Dataset.Post
            End With
        Case "COMPROMISOEAG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
        Case "COMPROMISOPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
        Case "LIBREEAG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
        Case "LIBREPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
    End Select
    
    Select Case UCase(dbgReposicion.Columns.FocusedColumn.FieldName)
        Case "COMPROMISOEAG", "COMPROMISOPLG", "LIBREEAG", "LIBREPLG"
            If frmUtilStockDetalle.RedistribucionEjecutada Then
                With objAyudaSolicitud
                    .Codigo = Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                    .CodProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                    
                    .verificarAtencionPorProducto
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "COMPROMISOEA = " & .CompromisoEA & ", "
                    SqlCad = SqlCad & "COMPROMISOPL = " & .CompromisoPL & ", "
                    SqlCad = SqlCad & "SALDOACTUAL = " & .SaldoPorAtender & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"
                    
                    cnDBTemp.Execute SqlCad
                    
                    .inicializarEntidades
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
                
                With objAyudaVale
                    .CodigoProducto = Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                    
                    .verificarStockProducto
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                    SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                    SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                    SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                    SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                    SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgReposicion.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgReposicion.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"
                    
                    cnDBTemp.Execute SqlCad
                    
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
                
                listarResumenRequerimiento
            End If
    End Select
End Sub

Private Sub dbgReposicion_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgReposicion.Columns.FocusedColumn.FieldName
        Case "CANTIDADPC"
            With dbgReposicion
                .Dataset.Edit
                
                If Val(dbgReposicion.Columns.ColumnByFieldName("CANTIDADPC").value & "") > Val(dbgReposicion.Columns.ColumnByFieldName("SALDOACTUAL").value & "") Then
                    MsgBox "La cantidad no puede exceder al Saldo por Atender, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = True
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgReposicion_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            If dbgReposicion.Dataset.State = dsEdit Or dbgReposicion.Dataset.State = dsInsert Then
                dbgReposicion.Dataset.Post
            End If
            
            dbgReposicion_OnDblClick
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    txtbusqueda.Text = vbNullString
    
    If strCodProducto = vbNullString Then
        Me.Caption = "Productos del Requerimiento N° " & strNroPedido & " - " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
    Else
        Me.Caption = "Producto Pendiente de Atención en otro(s) Requerimiento(s)"
    End If
    
    'listarAlmacenEnCombo
    
    cargarResumenRequerimiento
    
    cargarStockDeProducto
    
    cargarProductoAtendidoPorProveedor
    
    listarResumenRequerimiento
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgReposicion.Dataset.Close
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    dbgReposicion.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub tlbReposicion_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Filtrar"
            dbgReposicion.Filter.FilterActive = CBool(Tool.State)
        Case "ID_Salir"
            If dbgReposicion.Dataset.State = dsEdit Then
                dbgReposicion.Dataset.Post
            End If
            
            dbgReposicion.Dataset.Close
            
            Me.Hide
        Case "SeleccionarTodo"
            estadoSeleccion True
        Case "QuitarSeleccion"
            estadoSeleccion False
    End Select
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtbusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarResumenRequerimiento
    End Select
End Sub
