VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilDevolucionOC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Mercaderia de O/C"
   ClientHeight    =   8370
   ClientLeft      =   2145
   ClientTop       =   2010
   ClientWidth     =   13275
   Icon            =   "frmUtilDevolucionOC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   13275
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
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   6135
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   8
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   25
         Scrolling       =   1
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
      Height          =   735
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkSinProveedorEsp 
         Caption         =   "Orden sin Proveedor Especifico."
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkVerDescripcionProv 
         Caption         =   "Ver Descripción del Proveedor."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbDevolucion 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilDevolucionOC.frx":038A
      ToolBars        =   "frmUtilDevolucionOC.frx":B4A1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDevolucion 
      Height          =   6945
      Left            =   120
      OleObjectBlob   =   "frmUtilDevolucionOC.frx":B638
      TabIndex        =   3
      Top             =   960
      Width           =   13065
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDevolucion2 
      Height          =   1530
      Left            =   11160
      OleObjectBlob   =   "frmUtilDevolucionOC.frx":C2CE
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fraAlmacen 
      Caption         =   " Almacen "
      Enabled         =   0   'False
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
      Left            =   6720
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmUtilDevolucionOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTipoVale             As String
Private strCodProveedor         As String


Public Property Let TipoVale(ByVal Value As String)
    strTipoVale = Value
End Property

Public Property Get TipoVale() As String
    TipoVale = strTipoVale
End Property

Public Property Let CodigoProveedor(ByVal Value As String)
    strCodProveedor = Value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProveedor
End Property

Private Sub listarAlmacenEnCombo()
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
'            If cmbalmacen.ListCount > 0 Then
'                cmbalmacen.ListIndex = 0
'            End If
'    End If
End Sub

Public Sub cargarOC()
    Dim strCodProveedorFinal As String
    
    Screen.MousePointer = vbHourglass
    
    dbgDevolucion.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    Select Case strTipoVale
        Case "I"
            Me.Caption = "Recepción de Mercaderia de O/C Aprobada"
            
            If CBool(chkSinProveedorEsp.Value) Then
                strCodProveedorFinal = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "CodigoProveedorComprasVarias", "l")
                
                strCodProveedorFinal = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", strCodProveedorFinal, "T")
            Else
                strCodProveedorFinal = strCodProveedor
            End If
            
            With objAyudaVale
                .inicializarEntidades

                .CodigoProveedor = strCodProveedorFinal

                .listarGrillaFaltantePorOC dbgDevolucion, Nothing ', strCodAlmacen, strCodProveedorFinal  'strCodProveedor

                .inicializarEntidades
            End With
        Case "S"
            'Me.Caption = "Devolución de Mercaderia de O/C Ingresada"
            
            'objAyudaVale.listarGrillaIngresoPorOCdisponible dbgDevolucion, Nothing, strCodAlmacen, vbNullString, vbNullString
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub cargarOCSql()
    Dim strCodProveedorFinal As String
    
    Screen.MousePointer = vbHourglass
    
    dbgDevolucion.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    Select Case strTipoVale
        Case "I"
            Me.Caption = "Recepción de Mercaderia de O/C Aprobada"
            
            If CBool(chkSinProveedorEsp.Value) Then
                strCodProveedorFinal = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "CodigoProveedorComprasVarias", "l")
                
                strCodProveedorFinal = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2NEWRUC", "MAESTROS.EF2PROVEEDORES", "F2CODPROV", strCodProveedorFinal, "T")
            Else
                strCodProveedorFinal = strCodProveedor
            End If
            
            With objSqlAyudaVale
                .inicializarEntidades
                
                .CodigoProveedor = strCodProveedorFinal
                
                .listarGrillaFaltantePorOC dbgDevolucion, Nothing, "tmpCPRecepcionOC" & wusuario   ', strCodAlmacen, strCodProveedorFinal  'strCodProveedor
                
                .inicializarEntidades
            End With
        Case "S"
            'Me.Caption = "Devolución de Mercaderia de O/C Ingresada"
            
            'objAyudaVale.listarGrillaIngresoPorOCdisponible dbgDevolucion, Nothing, strCodAlmacen, vbNullString, vbNullString
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub listarOC()
    dbgDevolucion.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPUTILDEVOLUCIONOC "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROOC LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "LLAVE LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "CODPRODUCTO LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & IIf(Not CBool(chkVerDescripcionProv.Value), "NOMPRODUCTO", "NOMPRODUCTOPROV") & " LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NROOC, NROPEDIDO DESC, NOMPRODUCTO"
    
    With dbgDevolucion
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "LLAVE"
        
        .m.FullExpand
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
    End With
End Sub

Private Sub listarOCSql()
    dbgDevolucion.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPCPRECEPCIONOC" & UCase(wusuario) & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CODPRODUCTO <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROOC LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & IIf(Not CBool(chkVerDescripcionProv.Value), "NOMPRODUCTO", "NOMPRODUCTOPROV") & " LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NROOC, NROPEDIDO DESC, NOMPRODUCTO"
    
    With dbgDevolucion
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "LLAVE"
        
        .m.FullExpand
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
    End With
End Sub

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    dbgDevolucion.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPUTILDEVOLUCIONOC "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "CANTIDADDESTINO = " & IIf(bolEstado, "CANTIDAD", "0") & ", "
    SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "TRUE", "FALSE") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROOC LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "LLAVE LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "CODPRODUCTO LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "CODPRODUCTOORIGINAL LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarOC
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub estadoSeleccionSql(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccionSql
    
    dbgDevolucion.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPCPRECEPCIONOC" & UCase(wusuario) & " "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "CANTIDADDESTINO = " & IIf(bolEstado, "CANTIDAD", "0") & ", "
    SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "1", "0") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CODPRODUCTO <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NROOC LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    cnBdCPlus.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarOCSql
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccionSql:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub chkSinProveedorEsp_Click()
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        cargarOCSql
        
        listarOCSql
    Else
        cargarOC
        
        listarOC
    End If
End Sub

Private Sub chkVerDescripcionProv_Click()
    dbgDevolucion.Columns.ColumnByFieldName("NOMPRODUCTO").Visible = Not CBool(chkVerDescripcionProv.Value)
    dbgDevolucion.Columns.ColumnByFieldName("NOMPRODUCTOPROV").Visible = CBool(chkVerDescripcionProv.Value)
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        listarOCSql
    Else
        listarOC
    End If
End Sub

'Private Sub chkActivarFiltro_Click()
'    dbgDevolucion.Filter.FilterActive = CBool(chkActivarFiltro.Value)
'End Sub

Private Sub cmbAlmacen_Click()
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        cargarOCSql
    Else
        cargarOC
    End If
End Sub

Private Sub dbgDevolucion_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            If dbgDevolucion.Dataset.State = dsEdit Then
                dbgDevolucion.Dataset.Post
            End If
            
            With dbgDevolucion
                .Dataset.Edit
                
                If CBool(.Columns.ColumnByFieldName("PROCESAR").Value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").Value = Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "")
                    End If
                Else
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").Value = 0
                End If
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDevolucion_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "NROPEDIDO"
            If Trim(Text) = vbNullString Then
                Text = "Stock Libre"
            End If
            
            Font.Bold = True
            FontColor = vbWhite
        Case "CANTIDAD", "CANTIDADDESTINO"
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgDevolucion_OnDblClick()
    Select Case dbgDevolucion.Columns.FocusedColumn.FieldName
        Case "CODPRODUCTO", "NOMPRODUCTO", "UM", "CANTIDAD"
            With dbgDevolucion
                .Dataset.Edit
                
                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") = 0 Then
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").Value = Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "")
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").Value = IIf(Not CBool(dbgDevolucion.Columns.ColumnByFieldName("PROCESAR").Value), True, False)
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDevolucion_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgDevolucion.Columns.FocusedColumn.FieldName
        Case "CANTIDADDESTINO"
            With dbgDevolucion
                If .Dataset.State = dsEdit Then
                    .Dataset.Edit
                    
                    If Val(dbgDevolucion.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") > Val(dbgDevolucion.Columns.ColumnByFieldName("CANTIDAD").Value & "") Then
                        MsgBox "La cantidad no puede exceder al stock disponible, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    
                    .Columns.ColumnByFieldName("PROCESAR").Value = IIf(Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").Value & "") > 0, True, False)
                    
                    .Dataset.Post
                End If
            End With
    End Select
End Sub

Private Sub dbgDevolucion_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            If dbgDevolucion.Dataset.State = dsEdit Or dbgDevolucion.Dataset.State = dsInsert Then
                dbgDevolucion.Dataset.Post
            End If
            
            dbgDevolucion_OnDblClick
        Case vbKeyUp
            txtBusqueda.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    txtBusqueda.Text = vbNullString
    chkVerDescripcionProv.Value = vbUnchecked
    chkSinProveedorEsp.Value = vbUnchecked
    
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
    
    'listarAlmacenEnCombo
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        cargarOCSql
    Else
        cargarOC
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgDevolucion.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub tlbDevolucion_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Filtrar"
            dbgDevolucion.Filter.FilterActive = CBool(Tool.State)
        Case "ID_Salir"
            If dbgDevolucion.Dataset.State = dsEdit Then
                dbgDevolucion.Dataset.Post
            End If
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(RESUMEN.CODPROVEEDOR)", "(SELECT CODPROVEEDOR FROM TMPCPRECEPCIONOC" & UCase(wusuario) & " WHERE PROCESAR = 1 GROUP BY CODPROVEEDOR) AS RESUMEN", vbNullString, vbNullString, vbNullString, "RTRIM(LTRIM(RESUMEN.CODPROVEEDOR)) <> ''")) > 1 Then
                    MsgBox "Imposible descargar devolución de mas de un proveedor, verifique", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            Else
                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(RESUMEN.CODPROVEEDOR)", "(SELECT CODPROVEEDOR FROM TMPUTILDEVOLUCIONOC WHERE PROCESAR = TRUE GROUP BY CODPROVEEDOR) AS RESUMEN", vbNullString, vbNullString, vbNullString, "TRIM(RESUMEN.CODPROVEEDOR & '') <> ''")) > 1 Then
                    MsgBox "Imposible descargar devolución de mas de un proveedor, verifique", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            End If
            
            dbgDevolucion.Dataset.Close
            
            Me.Hide
        Case "SeleccionarTodo"
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                estadoSeleccionSql True
            Else
                estadoSeleccion True
            End If
        Case "QuitarSeleccion"
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                estadoSeleccion True
            Else
                estadoSeleccion False
            End If
    End Select
End Sub



Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                listarOCSql
            Else
                listarOC
            End If
        Case vbKeyDown
            dbgDevolucion.SetFocus
    End Select
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 3, 8, 22, 24, 26, 32, 37, 45, 46, 58, 65 To 90, 97 To 122, 209, 40, 41, 46, 48 To 57, 241 - 32
'            timTemporizador.Interval = 0
'            timTemporizador.Enabled = True
'            pgbProgresoBusqueda.value = 0
'            pgbProgresoBusqueda.Visible = True
'
'            timTemporizador_Timer
'        Case Else
'            timTemporizador.Interval = 0
'            timTemporizador.Enabled = False
'            pgbProgresoBusqueda.value = 0
'            pgbProgresoBusqueda.Visible = False
'    End Select
End Sub

Private Sub txtNroPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            cargarOC
'    End Select
End Sub
