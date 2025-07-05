VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilAlertaStockReponer 
   Caption         =   "Alerta de Stock de Productos a Reponer"
   ClientHeight    =   8775
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   14190
   Icon            =   "frmUtilAlertaStockReponer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
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
      TabIndex        =   2
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   9135
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   3
         Top             =   540
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   40
         Scrolling       =   1
      End
      Begin VB.TextBox txtNroPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   1080
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
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox chkActivarFiltro 
         Caption         =   "Activar Auto-filtros"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   7770
      Left            =   120
      OleObjectBlob   =   "frmUtilAlertaStockReponer.frx":058A
      TabIndex        =   6
      Top             =   960
      Width           =   14025
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
   Begin ActiveToolBars.SSActiveToolBars tlbAlerta 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   16
      ShowShortcutsInToolTips=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "frmUtilAlertaStockReponer.frx":1696
      ToolBars        =   "frmUtilAlertaStockReponer.frx":E14E
   End
   Begin MSComDlg.CommonDialog cmdlgAlerta 
      Left            =   0
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmUtilAlertaStockReponer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double

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

Public Sub listarStock()
    Screen.MousePointer = vbHourglass
    
    dbgResumen.Dataset.Close
    
        objAyudaVale.listarGrillaAlertaReposicionStockProducto dbgResumen, vbNullString, txtBusqueda.Text
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkActivarFiltro_Click()
    dbgResumen.Filter.FilterActive = CBool(chkActivarFiltro.value)
End Sub

Private Sub cmbAlmacen_Click()
    listarStock
End Sub

Private Sub dbgResumen_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "ENALMACENCOMP", "PORLLEGARCOMP", "ENALMACENLIBRE", "PORLLEGARLIBRE", "STOCKENALMACENTOTAL", "STOCKPORLLEGARTOTAL", "STOCKTOTAL"
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "ESTADO"
            Select Case Val(Text) 'Val(Node.Values(16))
                Case 0
                    Color = vbRed
                Case 1
                    Color = RGB(247, 150, 70)
                Case 2
                    Color = RGB(255, 255, 0)
                Case 3
                    Color = RGB(146, 208, 80)
                Case 4
                    Color = RGB(0, 176, 240)
            End Select
    End Select
End Sub

Private Sub dbgResumen_OnDblClick()
    For d = 0 To 25
        nSaveRecNo = dbgResumen.Dataset.RecNo
    Next
    
    With objAyudaBien
        .inicializarEntidades
        
        .Codigo = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").value & "")
        
        .obtenerConfigBien
    End With
    
    Select Case dbgResumen.Columns.FocusedColumn.FieldName
        Case "ENALMACENCOMP"
            If Val(dbgResumen.Columns.ColumnByFieldName("ENALMACENCOMP").value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            
        Case "PORLLEGARCOMP"
            If Val(dbgResumen.Columns.ColumnByFieldName("PORLLEGARCOMP").value & "") <= 0 Then
                'MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                'Exit Sub
            End If

            
        Case "ENALMACENLIBRE"
            
        Case "PORLLEGARLIBRE"
            If Val(dbgResumen.Columns.ColumnByFieldName("PORLLEGARLIBRE").value & "") <= 0 Then
                'MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
                
                'Exit Sub
            End If
            
            
    End Select
    
    If dbgResumen.Dataset.RecordCount >= nSaveRecNo Then
        dbgResumen.Dataset.RecNo = nSaveRecNo
    End If
End Sub

Private Sub dbgResumen_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgResumen_OnDblClick
    End Select
End Sub

Private Sub Form_Load()
    txtBusqueda.Text = vbNullString
        'ADD
        'txtNroPedido.Text = vbNullString
        
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
    
    'listarAlmacenEnCombo
    
    listarStock
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgResumen.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + 300)
End Sub

Private Sub timTemporizador_Timer()
'    If timTemporizador.Interval = 40 Then
'        listarStock
'
'        timTemporizador.Enabled = False
'        pgbProgresoBusqueda.value = 0
'        pgbProgresoBusqueda.Visible = False
'    Else
'        timTemporizador.Interval = timTemporizador.Interval + 1
'        pgbProgresoBusqueda.value = timTemporizador.Interval
'    End If
End Sub

Private Sub tlbAlerta_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Filtrar"
            If Tool.State = ssChecked Then
                dbgResumen.Filter.FilterActive = True
            Else
                dbgResumen.Filter.FilterActive = False
            End If
        Case "Agrupar"
            If Tool.State = ssChecked Then
                dbgResumen.Options.Set (egoShowGroupPanel)
            Else
                dbgResumen.Options.Unset (egoShowGroupPanel)
            End If
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgAlerta
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dbgResumen.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarStock
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
'            listarStock
'    End Select
End Sub

