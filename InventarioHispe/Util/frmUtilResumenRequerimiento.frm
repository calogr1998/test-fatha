VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmUtilResumenRequerimiento 
   Caption         =   "Resumen de Requerimiento"
   ClientHeight    =   8040
   ClientLeft      =   225
   ClientTop       =   1785
   ClientWidth     =   13440
   Icon            =   "frmUtilResumenRequerimiento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13440
   WindowState     =   2  'Maximized
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9855
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
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
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox chkProductoProveedor 
         Caption         =   "Mostrar productos de proveedor."
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkProductoSeleccionado 
         Caption         =   "Mostrar productos seleccionados."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
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
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   9375
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   2
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
   Begin ActiveToolBars.SSActiveToolBars tlbResumen 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilResumenRequerimiento.frx":058A
      ToolBars        =   "frmUtilResumenRequerimiento.frx":B6A1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   6825
      Left            =   120
      OleObjectBlob   =   "frmUtilResumenRequerimiento.frx":B838
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmUtilResumenRequerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNroPedido            As String
Private strCodProveedor         As String
Private strCodProducto          As String

Private bolResumenCargado       As Boolean

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
    
    cmbalmacen.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            cmbalmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If cmbalmacen.ListCount > 0 Then
                cmbalmacen.ListIndex = 0
            End If
    End If
End Sub

Public Sub cargarResumenRequerimiento()
    Screen.MousePointer = vbHourglass
    
    dbgResumen.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (1/5)..."
    
    abrirCnTemporal
           
    objAyudaSolicitud.listarGrillaResumenRequerimiento dbgResumen, Nothing, strNroPedido, strCodProducto
    
    dbgResumen.Dataset.Close
    
    Screen.MousePointer = vbDefault
    
    fraProceso.Visible = False
    fraProceso.Caption = vbNullString
    pgbProceso.value = 0
End Sub

Private Sub cargarIndicadoresAtencion(Optional ByVal strNroPedido As String, _
                                        Optional ByVal strCodProducto As String)
    On Error GoTo errCargarIndicadoresAtencion
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen.Dataset.Close
    
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
    
    DoEvents
    
    fraProceso.Visible = True
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.value = 0
        
        DoEvents
        
        fraProceso.Caption = "Actualizando Indicadores de Atencion (3/5)..."
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            With objAyudaSolicitud
                .Codigo = Trim(rstTemporal!NroPedido & "")
                .CodProducto = Trim(rstTemporal!CodProducto & "")

                '.verificarAtencionPorProducto
                ''SqlCad = SqlCad & "(VAL(RESU.CANTIDAD & '') - (VAL(COMPR.CANTIDAD & '') + VAL(PORL.CANTIDAD & ''))) AS SALDOACTUAL, "
                .CompromisoEA = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTO", "TIPO", "F", "T", "AND NROPEDIDO = '" & .Codigo & "' AND CODPRODUCTO = '" & .CodProducto & "'"))
                .CompromisoPL = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTO", "TIPO", "V", "T", "AND NROPEDIDO = '" & .Codigo & "' AND CODPRODUCTO = '" & .CodProducto & "'"))
                .SaldoPorAtender = Val(rstTemporal!Cantidad & "") - (.CompromisoEA + .CompromisoPL)
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEA = " & .CompromisoEA & ", "
                SqlCad = SqlCad & "COMPROMISOPL = " & .CompromisoPL & ", "
                SqlCad = SqlCad & "SALDOACTUAL = " & .SaldoPorAtender & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstTemporal!NroPedido & "") & "' AND "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                
                cnDBTemp.Execute SqlCad
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
            End With
            
            DoEvents
            
            pgbProceso.value = pgbProceso.value + 1
            fraProceso.Caption = "Actualizando Indicadores de Atencion (3/5)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    fraProceso.Caption = vbNullString
    pgbProceso.value = 0
    
    Exit Sub
errCargarIndicadoresAtencion:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub descargarAtencionRequerimiento()
    On Error GoTo errDescargarAtencionRequerimiento
    
    With objAyudaSolicitud
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .Codigo = strNroPedido
        
        DoEvents
    
        fraProceso.Visible = True
        fraProceso.Caption = "Ejecutando consulta (2/5)..."
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONREQUERIMIENTO"
        
        .descargarResumenAtencionPorProducto "F"
        
        abrirCnTemporal
        
        .descargarResumenAtencionPorProducto "V"
        
        .inicializarEntidades
        .inicializarEntidadesDetalle
    End With
    
    Exit Sub
errDescargarAtencionRequerimiento:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub cargarStockDeProducto(Optional ByVal strNroPedido As String, _
                                    Optional ByVal strCodProducto As String)
    On Error GoTo errCargarStockProducto
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen.Dataset.Close
    
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
    
    DoEvents
    
    fraProceso.Visible = True
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.value = 0
        
        DoEvents
        
        fraProceso.Caption = "Actualizando Stock de Productos (4/5)..."
        
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
            
            DoEvents
            
            pgbProceso.value = pgbProceso.value + 1
            fraProceso.Caption = "Actualizando Stock de Productos (4/5)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    fraProceso.Caption = vbNullString
    pgbProceso.value = 0
    
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
    
    dbgResumen.Dataset.Close
    
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
    
    DoEvents
    
    fraProceso.Visible = True
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.value = 0
        
        DoEvents
        
        fraProceso.Caption = "Verificando Productos Atendidos por Proveedor (5/5)..."
        
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
            
            DoEvents
            
            pgbProceso.value = pgbProceso.value + 1
            fraProceso.Caption = "Verificando Productos Atendidos por Proveedor (5/5)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    fraProceso.Caption = vbNullString
    pgbProceso.value = 0
    
    Exit Sub
errCargarProductoAtendidoPorProveedor:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub listarResumenRequerimiento()
    dbgResumen.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "RESU.LLAVE, "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.ITEM, "
    SqlCad = SqlCad & "RESU.CODPRODUCTO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
    SqlCad = SqlCad & "RESU.UM, "
    SqlCad = SqlCad & "RESU.CANTIDAD, "
    SqlCad = SqlCad & "RESU.COMPROMISOEA, "
    SqlCad = SqlCad & "RESU.COMPROMISOPL, "
    SqlCad = SqlCad & "RESU.SALDOACTUAL, "
    'SqlCad = SqlCad & "VAL(COMPR.CANTIDAD & '') AS COMPROMISOEA, "
    'SqlCad = SqlCad & "VAL(PORL.CANTIDAD & '') AS COMPROMISOPL, "
    'SqlCad = SqlCad & "(VAL(RESU.CANTIDAD & '') - (VAL(COMPR.CANTIDAD & '') + VAL(PORL.CANTIDAD & ''))) AS SALDOACTUAL, "
    SqlCad = SqlCad & "RESU.COMPROMISOEAG, "
    SqlCad = SqlCad & "RESU.COMPROMISOPLG, "
    SqlCad = SqlCad & "RESU.LIBREEAG, "
    SqlCad = SqlCad & "RESU.LIBREPLG, "
    SqlCad = SqlCad & "RESU.STOCKEAG, "
    SqlCad = SqlCad & "RESU.STOCKPLG, "
    SqlCad = SqlCad & "RESU.CANTIDADPC, "
    SqlCad = SqlCad & "RESU.PROCESAR, "
    SqlCad = SqlCad & "RESU.ATENDIDOPORPROV "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO AS RESU "
    'SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, CANTIDAD FROM TMPUTILRESUMENATENCIONREQUERIMIENTO WHERE TIPO = 'F') AS COMPR "
    'SqlCad = SqlCad & "ON COMPR.NROPEDIDO = RESU.NROPEDIDO AND COMPR.CODPRODUCTO = RESU.CODPRODUCTO) "
    'SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, CANTIDAD FROM TMPUTILRESUMENATENCIONREQUERIMIENTO WHERE TIPO = 'V') AS PORL "
    'SqlCad = SqlCad & "ON PORL.NROPEDIDO = RESU.NROPEDIDO AND PORL.CODPRODUCTO = RESU.CODPRODUCTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(RESU.CODPRODUCTO & '') <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "RESU.NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
        If CBool(chkProductoSeleccionado.value) Then
            SqlCad = SqlCad & "AND RESU.PROCESAR = TRUE "
        End If
        
        If CBool(chkProductoProveedor.value) Then
            SqlCad = SqlCad & "AND RESU.ATENDIDOPORPROV = TRUE "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO"
    
    With dbgResumen
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
    
    dbgResumen.Dataset.Close
    
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
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
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

Private Sub dbgResumen_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            With dbgResumen
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

Private Sub dbgResumen_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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

Private Sub dbgResumen_OnDblClick()
    Select Case dbgResumen.Columns.FocusedColumn.FieldName
        Case "CODPRODUCTO", "NOMPRODUCTO"
            If Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "") <> vbNullString Then
                With frmUtilDetalleConsolidadoPedido
                    .NroPedido = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                    .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                    
                    .Show 1
                End With
            End If
        Case "COMPROMISOEA"
            'If Val(dbgResumen.Columns.ColumnByFieldName("COMPROMISOEA").value & "") = 0 Then
            '    MsgBox "Cantidad Comprometida en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName
            '
            '    dbgResumen.M.FullRefresh
            '
            '    Exit Sub
            'End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "F"
                .NroPedido = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                
                .Show vbModal
            End With
        Case "COMPROMISOPL"
            If Val(dbgResumen.Columns.ColumnByFieldName("COMPROMISOPL").value & "") = 0 Then
                MsgBox "Cantidad por Llegar en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "V"
                .NroPedido = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                
                .Show vbModal
            End With
        Case "UM", "SALDOACTUAL"
            If Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") <= 0 Then
                MsgBox "Imposible seleccionar Item, ya se encuentra completamente atendido, verifique 'Saldo'.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With dbgResumen
                .Dataset.Edit
                
                If Not CBool(dbgResumen.Columns.ColumnByFieldName("PROCESAR").value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADPC").value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADPC").value = Val(.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                    End If
                Else
                    .Columns.ColumnByFieldName("CANTIDADPC").value = 0
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = IIf(Not CBool(dbgResumen.Columns.ColumnByFieldName("PROCESAR").value), True, False)
                
                .Dataset.Post
            End With
        Case "COMPROMISOEAG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = True 'IIf(Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
        Case "COMPROMISOPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
                .Show 1
            End With
        Case "LIBREEAG"
            With objAyudaBien
                .Codigo = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                
                .obtenerConfigBien
                
                If Val(dbgResumen.Columns.ColumnByFieldName("LIBREEAG").value & "") <= .StockMin Then
                    MsgBox "Imposible hacer uso de Stock actual, se ha configurado como Stock Minimo o de Seguridad." & vbNewLine & _
                            "Se recomienda realizar una Compra para contar con Stock Disponible.", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            End With
        
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") + (Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") * (objAyudaBien.PorcentajeDemasia / 100))
                
                .Show 1
            End With
        Case "LIBREPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                .CantidadMaximaParaPedido = Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "")
                
5                .Show 1
            End With
    End Select
    
    Select Case UCase(dbgResumen.Columns.FocusedColumn.FieldName)
        Case "COMPROMISOEAG", "COMPROMISOPLG", "LIBREEAG", "LIBREPLG"
            If frmUtilStockDetalle.RedistribucionEjecutada Then
                With objAyudaSolicitud
                    .Codigo = Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "")
                    .CodProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")

                    .verificarAtencionPorProducto

                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTO "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "COMPROMISOEA = " & .CompromisoEA & ", "
                    SqlCad = SqlCad & "COMPROMISOPL = " & .CompromisoPL & ", "
                    SqlCad = SqlCad & "SALDOACTUAL = " & .SaldoPorAtender & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"

                    cnDBTemp.Execute SqlCad

'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "UPDATE "
'                    SqlCad = SqlCad & "TMPUTILRESUMENATENCIONREQUERIMIENTO "
'                    SqlCad = SqlCad & "SET "
'                    SqlCad = SqlCad & "CANTIDAD = " & .CompromisoEA & " "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "TIPO = 'F' AND "
'                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
'                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"
'
'                    cnDBTemp.Execute SqlCad
'
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "UPDATE "
'                    SqlCad = SqlCad & "TMPUTILRESUMENATENCIONREQUERIMIENTO "
'                    SqlCad = SqlCad & "SET "
'                    SqlCad = SqlCad & "CANTIDAD = " & .CompromisoPL & " "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "TIPO = 'V' AND "
'                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
'                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"
'
'                    cnDBTemp.Execute SqlCad
                    
                    .inicializarEntidades
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
                
                With objAyudaVale
                    .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "")
                    
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
                    SqlCad = SqlCad & "NROPEDIDO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("NROPEDIDO").value & "") & "' AND "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(dbgResumen.Columns.ColumnByFieldName("CODPRODUCTO").value & "") & "'"
                    
                    cnDBTemp.Execute SqlCad
                    
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
                
                listarResumenRequerimiento
            End If
    End Select
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgResumen_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgResumen.Columns.FocusedColumn.FieldName
        Case "CANTIDADPC"
            With dbgResumen
                .Dataset.Edit
                
                If Val(dbgResumen.Columns.ColumnByFieldName("CANTIDADPC").value & "") > Val(dbgResumen.Columns.ColumnByFieldName("SALDOACTUAL").value & "") Then
                    MsgBox "La cantidad no puede exceder al Saldo por Atender, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = True
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgResumen_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            If dbgResumen.Dataset.State = dsEdit Or dbgResumen.Dataset.State = dsInsert Then
                dbgResumen.Dataset.Post
            End If
            
            dbgResumen_OnDblClick
    End Select
End Sub

Private Sub Form_Activate()
    fraOpciones.Enabled = bolResumenCargado
    tlbResumen.Enabled = bolResumenCargado
    dbgResumen.Enabled = bolResumenCargado
        
    If bolResumenCargado = False Then
        If strCodProducto = vbNullString Then
            Me.Caption = "Productos del Requerimiento N° " & strNroPedido & " - " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
        Else
            Me.Caption = "Producto Pendiente de Atención en otro(s) Requerimiento(s)"
        End If
        
        'PASO 1)
        cargarResumenRequerimiento
        
        'PASO 2)
        descargarAtencionRequerimiento
        
        'PASO 3)
        cargarIndicadoresAtencion
        
        'PASO 4)
        cargarStockDeProducto
        
        'PASO 5)
        cargarProductoAtendidoPorProveedor
        
        listarResumenRequerimiento
        
        bolResumenCargado = True
        
        fraOpciones.Enabled = bolResumenCargado
        tlbResumen.Enabled = bolResumenCargado
        dbgResumen.Enabled = bolResumenCargado
    End If
End Sub

Private Sub Form_Load()
    bolResumenCargado = False
    
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    txtBusqueda.Text = vbNullString
    
    abrirCnTemporal
        
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONREQUERIMIENTO"
    
'    If strCodProducto = vbNullString Then
'        Me.Caption = "Productos del Requerimiento N° " & strNroPedido & " - " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strNroPedido, "T")
'    Else
'        Me.Caption = "Producto Pendiente de Atención en otro(s) Requerimiento(s)"
'    End If
    
    'listarAlmacenEnCombo
    
'    cargarResumenRequerimiento
'
'    cargarStockDeProducto
'
'    cargarProductoAtendidoPorProveedor
'
'    listarResumenRequerimiento
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgResumen.Dataset.Close
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    dbgResumen.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub tlbResumen_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Filtrar"
            dbgResumen.Filter.FilterActive = CBool(Tool.State)
        Case "SeleccionarTodo"
            estadoSeleccion True
        Case "QuitarSeleccion"
            estadoSeleccion False
        Case "ID_Salir"
            If dbgResumen.Dataset.State = dsEdit Then
                dbgResumen.Dataset.Post
            Else
                If dbgResumen.Dataset.RecordCount > 0 Then
                    dbgResumen.Dataset.Edit
                    
                    dbgResumen.Dataset.Post
                End If
            End If
            
            dbgResumen.Dataset.Close
            
            Me.Hide
    End Select
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarResumenRequerimiento
    End Select
End Sub

