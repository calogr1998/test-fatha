VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmUtilSalidaOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insumos Descargados de OP"
   ClientHeight    =   8295
   ClientLeft      =   4095
   ClientTop       =   2160
   ClientWidth     =   13245
   Icon            =   "frmUtilSalidaOP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13245
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
      Top             =   120
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
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox chkActivarFiltro 
         Caption         =   "Activar Auto-filtros"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   6135
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   2
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
   Begin ActiveToolBars.SSActiveToolBars tlbDevolucion 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilSalidaOP.frx":038A
      ToolBars        =   "frmUtilSalidaOP.frx":B4A1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDevolucion 
      Height          =   6945
      Left            =   120
      OleObjectBlob   =   "frmUtilSalidaOP.frx":B638
      TabIndex        =   7
      Top             =   960
      Width           =   13065
   End
End
Attribute VB_Name = "frmUtilSalidaOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTipoVale                 As String
Private strIdOrdenProduccion        As String
Private strNroOP                    As String

Public Property Let TipoVale(ByVal value As String)
    strTipoVale = value
End Property

Public Property Get TipoVale() As String
    TipoVale = strTipoVale
End Property

Public Property Let IdOrdenProduccion(ByVal value As String)
    strIdOrdenProduccion = value
End Property

Public Property Get IdOrdenProduccion() As String
    IdOrdenProduccion = strIdOrdenProduccion
End Property

Public Property Let NroOP(ByVal value As String)
    strNroOP = value
End Property

Public Property Get NroOP() As String
    NroOP = strNroOP
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

Public Sub cargarOP()
    Screen.MousePointer = vbHourglass
    
    dbgDevolucion.Dataset.Close
    
    txtBusqueda.Text = vbNullString
    
    Select Case strTipoVale
        Case "I"
            Me.Caption = "Devolución de O/P: " & strNroOP
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                objSqlAyudaVale.listarGrillaDevolucionPorOP dbgDevolucion, _
                                                            Nothing, _
                                                            right(cmbalmacen.Text, 2), _
                                                            strIdOrdenProduccion, _
                                                            "tmpCPDevolucionOp" & wusuario
            Else
                objAyudaVale.listarGrillaDevolucionPorOP dbgDevolucion, Nothing, right(cmbalmacen.Text, 2), strIdOrdenProduccion
            End If
        Case "S"
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub listarOC()
    dbgDevolucion.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPUTILDEVOLUCIONOP "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' AND "
    SqlCad = SqlCad & "VAL(CANTIDAD & '') > 0 "
        
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
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
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

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    dbgDevolucion.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPUTILDEVOLUCIONOP "
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

'Private Sub chkActivarFiltro_Click()
'    dbgDevolucion.Filter.FilterActive = CBool(chkActivarFiltro.Value)
'End Sub

Private Sub cmbAlmacen_Click()
    cargarOP
End Sub

Private Sub dbgDevolucion_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            If dbgDevolucion.Dataset.State = dsEdit Then
                dbgDevolucion.Dataset.Post
            End If
            
            With dbgDevolucion
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
                
                If Val(.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") = 0 Then
                    .Columns.ColumnByFieldName("CANTIDADDESTINO").value = Val(.Columns.ColumnByFieldName("CANTIDAD").value & "")
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").value = IIf(Not CBool(dbgDevolucion.Columns.ColumnByFieldName("PROCESAR").value), True, False)
                
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
                    
                    If Val(dbgDevolucion.Columns.ColumnByFieldName("CANTIDADDESTINO").value & "") > Val(dbgDevolucion.Columns.ColumnByFieldName("CANTIDAD").value & "") Then
                        MsgBox "La cantidad no puede exceder a la cantidad descargada, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    .Columns.ColumnByFieldName("PROCESAR").value = True
                    
                    .Dataset.Post
                End If
            End With
    End Select
End Sub

Private Sub dbgDevolucion_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            If dbgDevolucion.Dataset.State = dsEdit Or dbgDevolucion.Dataset.State = dsInsert Then
                dbgDevolucion.Dataset.Post
            End If
            
            dbgDevolucion_OnDblClick
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
    
    listarAlmacenEnCombo
    
    'cargarOP
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgDevolucion.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub tlbDevolucion_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Filtrar"
            dbgDevolucion.Filter.FilterActive = CBool(Tool.State)
        Case "ID_Salir"
            If dbgDevolucion.Dataset.State = dsEdit Then
                dbgDevolucion.Dataset.Post
            End If
            
'            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(RESUMEN.CODPROVEEDOR)", "(SELECT CODPROVEEDOR FROM TMPUTILDEVOLUCIONOP WHERE PROCESAR = TRUE GROUP BY CODPROVEEDOR) AS RESUMEN", vbNullString, vbNullString, vbNullString, "TRIM(RESUMEN.CODPROVEEDOR & '') <> ''")) > 1 Then
'                MsgBox "Imposible descargar devolución de mas de un proveedor, verifique", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            dbgDevolucion.Dataset.Close
            
            Me.Hide
        Case "SeleccionarTodo"
            estadoSeleccion True
        Case "QuitarSeleccion"
            estadoSeleccion False
    End Select
End Sub

Private Sub timTemporizador_Timer()
'    If timTemporizador.Interval = 25 Then
'        listarOC
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
    Select Case KeyCode
        Case vbKeyReturn
            listarOC
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
'            cargarOP
'    End Select
End Sub
