VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmUtilStockDisponible 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Libre Disponible"
   ClientHeight    =   8430
   ClientLeft      =   2595
   ClientTop       =   1800
   ClientWidth     =   13275
   Icon            =   "frmUtilStockDisponible.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
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
      TabIndex        =   6
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
         TabIndex        =   7
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
      Height          =   735
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox chkActivarFiltro 
         Caption         =   "Activar Auto-filtros"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDisponible2 
      Height          =   4290
      Left            =   12120
      OleObjectBlob   =   "frmUtilStockDisponible.frx":038A
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   5265
   End
   Begin ActiveToolBars.SSActiveToolBars tlbDisponible 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilStockDisponible.frx":0E7E
      ToolBars        =   "frmUtilStockDisponible.frx":BF95
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDisponible 
      Height          =   6945
      Left            =   120
      OleObjectBlob   =   "frmUtilStockDisponible.frx":C12C
      TabIndex        =   8
      Top             =   960
      Width           =   13065
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
      TabIndex        =   4
      Top             =   1200
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
Attribute VB_Name = "frmUtilStockDisponible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strCodAlmacen As String
Private strCadenaCorte As String


'Propiedad Codigo de Almacen
Public Property Let CodigoAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property

'Propiedad Cadena de Corte de Informacion
Public Property Let CadenaCorte(ByVal value As String)
    strCadenaCorte = value
End Property

Public Property Get CadenaCorte() As String
    CadenaCorte = strCadenaCorte
End Property


Private Sub listarAlmacenEnCombo()
    objSqlAyudaAlmacen.listarAlmacen cmbalmacen, False
    
'    Dim rstAlmacen As New ADODB.Recordset
'
'    If rstAlmacen.State = 1 Then rstAlmacen.Close
'
'    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    cmbAlmacen.Clear
'
'    If Not rstAlmacen.EOF Then
'        rstAlmacen.MoveFirst
'
'        Do While Not rstAlmacen.EOF
'            cmbAlmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
'
'            rstAlmacen.MoveNext
'        Loop
'            If cmbAlmacen.ListCount > 0 Then
'                cmbAlmacen.ListIndex = 0
'            End If
'    End If
End Sub

Public Sub cargarStockDisponible()
    Screen.MousePointer = vbHourglass
    
    dbgDisponible.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    With objAyudaVale
        .inicializarEntidades
        
        .CodigoAlmacen = strCodAlmacen
        
        .listarGrillaStockDisponibleV2 dbgDisponible, Nothing, strCadenaCorte
        
        .inicializarEntidades
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub cargarStockDisponibleSql()
    Screen.MousePointer = vbHourglass
    
    dbgDisponible.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    With objSqlAyudaVale
        .inicializarEntidades
        
        .CodigoAlmacen = strCodAlmacen
        
        .listarGrillaStockDisponibleV2 dbgDisponible, Nothing, strCadenaCorte, "tmpCPStockDisponible" & wusuario
        
        .inicializarEntidades
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub listarStockDisponible()
    dbgDisponible.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPUTILSTOCKDISPONIBLE "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    With dbgDisponible
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "CODPRODUCTO"
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
    End With
End Sub

Private Sub listarStockDisponibleSql()
    dbgDisponible.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPCPSTOCKDISPONIBLE" & UCase(wusuario) & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CODPRODUCTO <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    With dbgDisponible
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "CODPRODUCTO"
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
    End With
End Sub

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    dbgDisponible.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPUTILSTOCKDISPONIBLE "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "CANTIDADDESTINO = " & IIf(bolEstado, "CANTIDAD", "0") & ", "
    SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "TRUE", "FALSE") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarStockDisponible
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub estadoSeleccionSql(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccionSql
    
    dbgDisponible.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPCPSTOCKDISPONIBLE" & UCase(wusuario) & " "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "CANTIDADDESTINO = " & IIf(bolEstado, "CANTIDAD", "0") & ", "
    SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "1", "0") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CODPRODUCTO <> '' "
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "NOMPRODUCTO LIKE '%" & txtBusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
    
    cnBdCPlus.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarStockDisponible
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccionSql:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cmbAlmacen_Click()
    'cargarStockDisponible
End Sub

Private Sub cmbalmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyLeft
            txtBusqueda.SetFocus
        Case vbKeyRight
            chkActivarFiltro.SetFocus
        Case vbKeyDown
            dbgDisponible.SetFocus
    End Select
End Sub

Private Sub dbgDisponible_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            If dbgDisponible.Dataset.State = dsEdit Then
                dbgDisponible.Dataset.Post
            End If
            
            With dbgDisponible
                .Dataset.Edit
                
                If CBool(.Columns.ColumnByFieldName("PROCESAR").value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                    End If
                Else
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = 0
                End If
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDisponible_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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

Private Sub dbgDisponible_OnDblClick()
    Select Case dbgDisponible.Columns.FocusedColumn.FieldName
        Case "CODPRODUCTO", "NOMPRODUCTO", "UM", "CANTIDAD"
            With dbgDisponible
                .Dataset.Edit
                
                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = IIf(Not CBool(dbgDisponible.Columns.ColumnByFieldName("PROCESAR").value), True, False)
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDisponible_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgDisponible.Columns.FocusedColumn.FieldName
        Case "CANTIDADDESTINO"
            With dbgDisponible
                If dbgDisponible.Dataset.State = dsEdit Or dbgDisponible.Dataset.State = dsInsert Then
                    dbgDisponible.Dataset.Post
                End If
                
                .Dataset.Edit
                
                If Val(dbgDisponible.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > Val(dbgDisponible.Columns.ColumnByFieldName("CANTIDAD").value & "") Then
                    MsgBox "La cantidad no puede exceder al stock disponible, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    '.Dataset.Cancel
                    
                    'Exit Sub
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(dbgDisponible.Columns.ColumnByFieldName("CANTIDAD").value & "")
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = IIf(Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > 0, True, False)
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgDisponible_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            If dbgDisponible.Dataset.State = dsEdit Or dbgDisponible.Dataset.State = dsInsert Then
                dbgDisponible.Dataset.Post
            End If
            
            dbgDisponible_OnDblClick
        Case vbKeyUp
            txtBusqueda.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    txtBusqueda.Text = vbNullString
    
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
    
    'listarAlmacenEnCombo
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        cargarStockDisponibleSql
    Else
        cargarStockDisponible
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgDisponible.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub tlbDisponible_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Filtrar"
            dbgDisponible.Filter.FilterActive = CBool(Tool.State)
        Case "ID_Salir"
            If dbgDisponible.Dataset.State = dsEdit Then
                dbgDisponible.Dataset.Post
            Else
                dbgDisponible.Dataset.Edit
                
                dbgDisponible.Dataset.Post
            End If
            
'            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(RESUMEN.CODPROVEEDOR)", "(SELECT CODPROVEEDOR FROM TMPUTILSTOCKDISPONIBLE WHERE PROCESAR = TRUE GROUP BY CODPROVEEDOR) AS RESUMEN", vbNullString, vbNullString, vbNullString, "TRIM(RESUMEN.CODPROVEEDOR & '') <> ''")) > 1 Then
'                MsgBox "Imposible descargar devolución de mas de un proveedor, verifique", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            dbgDisponible.Dataset.Close
            
            Me.Hide
        Case "SeleccionarTodo"
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                estadoSeleccionSql True
            Else
                estadoSeleccion True
            End If
        Case "QuitarSeleccion"
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                estadoSeleccionSql False
            Else
                estadoSeleccion False
            End If
    End Select
End Sub

Private Sub timTemporizador_Timer()
'    If timTemporizador.Interval = 25 Then
'        listarStockDisponible
'
'        timTemporizador.Enabled = False
'        pgbProgresoBusqueda.value = 0
'        pgbProgresoBusqueda.Visible = False
'    Else
'        timTemporizador.Interval = timTemporizador.Interval + 1
'        pgbProgresoBusqueda.value = timTemporizador.Interval
'    End If
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                listarStockDisponibleSql
            Else
                listarStockDisponible
            End If
        Case vbKeyDown
            dbgDisponible.SetFocus
        Case vbKeyRight
            cmbalmacen.SetFocus
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
'            cargarStockDisponible
'    End Select
End Sub
