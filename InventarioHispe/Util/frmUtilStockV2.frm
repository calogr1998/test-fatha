VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilStockV2 
   Caption         =   "Stock  de Productos Version 2.0"
   ClientHeight    =   8880
   ClientLeft      =   3645
   ClientTop       =   1935
   ClientWidth     =   14205
   Icon            =   "frmUtilStockV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlgStock 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveToolBars.SSActiveToolBars tlbStock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   2
      ToolsCount      =   9
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilStockV2.frx":058A
      ToolBars        =   "frmUtilStockV2.frx":39E6
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   1080
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
      Width           =   7215
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   40
         Scrolling       =   1
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   6975
      End
      Begin VB.TextBox txtNroPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   7770
      Left            =   120
      OleObjectBlob   =   "frmUtilStockV2.frx":3B95
      TabIndex        =   2
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
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
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
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   5760
      Width           =   975
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
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmUtilStockV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double


Private Sub listarFamilia()
    objAyudaFamilia.listarFamilia tlbStock.Tools.ITEM("Familia").ComboBox
    
    If tlbStock.Tools.ITEM("Familia").ComboBox.ListCount > 1 Then
        tlbStock.Tools.ITEM("Familia").ComboBox.ListIndex = 1
    End If
    
    If tlbStock.Tools.ITEM("Familia").ComboBox.ListIndex = 0 Then
        tlbStock.Tools.ITEM("SubFamilia").Enabled = False
    Else
        tlbStock.Tools.ITEM("SubFamilia").Enabled = True
    End If
End Sub

Private Sub listarSubFamilia()
    With objAyudaSubFamilia
        .CodigoFamilia = Trim(right(tlbStock.Tools.ITEM("Familia").ComboBox.Text, 4))
        
        .listarSubFamilia tlbStock.Tools.ITEM("SubFamilia").ComboBox
        
        If .CodigoFamilia = vbNullString Then
            tlbStock.Tools.ITEM("SubFamilia").Enabled = False
        Else
            tlbStock.Tools.ITEM("SubFamilia").Enabled = True
        End If
    End With
End Sub

Private Sub listarAlmacenEnCombo()
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM MAESTROS.EF2ALMACENES ORDER BY F2CODALM", cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
        rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    cmbAlmacen.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        cmbAlmacen.AddItem "(*) - Todos" & Space(100)
        
        Do While Not rstAlmacen.EOF
            cmbAlmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If cmbAlmacen.ListCount > 0 Then
                cmbAlmacen.ListIndex = 1
            End If
    End If
End Sub

Private Sub listarStock()
    Screen.MousePointer = vbHourglass
    
    dbgResumen.Dataset.Close
    
        Screen.MousePointer = vbDefault
End Sub

Private Sub cmbAlmacen_Click()
    listarStock
End Sub

Private Sub dbgResumen_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "STOCKCOMPROMETIDO", "STOCKPORLLEGARCOMP", "STOCKLIBRE", "STOCKPORLLEGARLIBRE", "STOCKTOTAL", "STOCKPORLLEGAR", "STOCK"
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgResumen_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "STOCKCOMPROMETIDO", "STOCKPORLLEGARCOMP", "STOCKLIBRE", "STOCKPORLLEGARLIBRE", "STOCKTOTAL", "STOCKPORLLEGAR", "STOCK"
            Text = Format(Text, "#,0.00;(#,0.00)")
            Color = vbWhite
            FontColor = vbBlue
            Font.Bold = True
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
        Case "STOCKCOMPROMETIDO"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKCOMPROMETIDO").value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            
        Case "STOCKPORLLEGARCOMP"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").value & "") <= 0 Then
                'MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                'Exit Sub
            End If

            
        Case "STOCKLIBRE"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").value & "") <= 0 Then
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

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarStock
    End Select
End Sub

Private Sub Form_Load()
    txtBusqueda.Text = vbNullString
    
    listarFamilia
    
    listarSubFamilia
    
    listarAlmacenEnCombo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgResumen.Dataset.Close
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgResumen.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, (Me.ScaleHeight - (fraBusqueda.Height + 300))
End Sub

Private Sub tlbStock_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Familia"
            listarSubFamilia
            
            listarStock
        Case "SubFamilia"
            listarStock
    End Select
End Sub

Private Sub tlbStock_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Filtrar"
            dbgResumen.Filter.FilterActive = CBool(Tool.State)
        Case "ExportarExcel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgStock
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
