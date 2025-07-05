VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilTomaInventario 
   Caption         =   "Toma de Inventario"
   ClientHeight    =   9015
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   15210
   Icon            =   "frmUtilTomaInventario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDatos 
      Caption         =   " Datos de Toma de Inventario "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   14895
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   10440
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117112833
         CurrentDate     =   42019
      End
      Begin VB.TextBox txtValeSalida 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtValeIngreso 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtObservacion 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "frmUtilTomaInventario.frx":058A
         Top             =   480
         Width           =   7335
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   9000
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblValeSalidaExterno 
         Alignment       =   2  'Center
         Caption         =   "< ID Externo >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   12720
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblValeIngresoExterno 
         Alignment       =   2  'Center
         Caption         =   "< ID Externo >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   12720
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Vale de Salida"
         Height          =   255
         Left            =   9000
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Vale de Ingreso"
         Height          =   255
         Left            =   9000
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   7920
         X2              =   7920
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Observación"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Proceso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   12015
      Begin MSComctlLib.ProgressBar pgbProceso1 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblProceso1 
         Caption         =   "Proceso 1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   11535
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgInventario 
      Height          =   6360
      Left            =   120
      OleObjectBlob   =   "frmUtilTomaInventario.frx":0590
      TabIndex        =   11
      Top             =   2280
      Width           =   14970
   End
   Begin MSComDlg.CommonDialog cmdlgInventario 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFiltro 
      Caption         =   " Filtro "
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
      Left            =   8160
      TabIndex        =   9
      Top             =   1440
      Width           =   6855
      Begin VB.OptionButton optFiltro 
         Caption         =   "Familia"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Sub-Familia"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
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
      Top             =   1440
      Width           =   7935
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   7455
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbInventario 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   2
      ToolsCount      =   12
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilTomaInventario.frx":38E9
      ToolBars        =   "frmUtilTomaInventario.frx":C46D
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
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmUtilTomaInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda            As Boolean
Private strAnno             As String
Private strMes              As String
Private strCodAlmacen       As String

Private strFichero          As String
Private bolObviarCierre     As Boolean

Private objTomaInventario   As ClsTomaInventario


Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let Anno(ByVal value As String)
    strAnno = value
End Property

Public Property Get Anno() As String
    Anno = strAnno
End Property

Public Property Let mes(ByVal value As String)
    strMes = value
End Property

Public Property Get mes() As String
    mes = strMes
End Property

Public Property Let CodigoAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property


Private Sub listarAlmacen()
    On Error GoTo errListarAlmacen
    
    objAyudaTomaInventario.listarAlmacenSoloSeleccion tlbInventario.Tools.ITEM("Almacen").ComboBox
    
    If tlbInventario.Tools.ITEM("Almacen").ComboBox.ListCount > 0 Then
        tlbInventario.Tools.ITEM("Almacen").ComboBox.ListIndex = tlbInventario.Tools.ITEM("Almacen").ComboBox.ListCount - 1
    End If
    
    Exit Sub
errListarAlmacen:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarAlmacen"
    
    Err.Clear
End Sub

Private Sub listarAnno()
    On Error GoTo errListarAnno
    
    objAyudaTomaInventario.listarAnnoSoloSeleccion tlbInventario.Tools.ITEM("Anno").ComboBox, _
                                                    Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
    
    If tlbInventario.Tools.ITEM("Anno").ComboBox.ListCount > 0 Then
        tlbInventario.Tools.ITEM("Anno").ComboBox.ListIndex = tlbInventario.Tools.ITEM("Anno").ComboBox.ListCount - 1
    End If
    
    Exit Sub
errListarAnno:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarAnno"
    
    Err.Clear
End Sub

Private Sub listarMes()
    On Error GoTo errListarMes
    
    objAyudaTomaInventario.listarMesSoloSeleccion tlbInventario.Tools.ITEM("Mes").ComboBox, _
                                                    Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), _
                                                    Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
    
    If tlbInventario.Tools.ITEM("Mes").ComboBox.ListCount > 0 Then
        tlbInventario.Tools.ITEM("Mes").ComboBox.ListIndex = tlbInventario.Tools.ITEM("Mes").ComboBox.ListCount - 1
    End If
    
    Exit Sub
errListarMes:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarMes"
    
    Err.Clear
End Sub

Private Sub listarGrilla(Optional ByVal strFiltroSensitivo As String, _
                            Optional ByVal StrFiltro As String, _
                            Optional ByVal intOpcion As Integer)
    
    On Error GoTo errListarGrilla
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILTOMAINVENTARIO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If strFiltroSensitivo <> vbNullString Then
            SqlCad = SqlCad & "AND NOMPRODUCTO LIKE '%" & strFiltroSensitivo & "%' "
        End If
        
        If StrFiltro <> vbNullString And StrFiltro <> "(*) Todos" Then
            Select Case intOpcion
                Case 0
                    SqlCad = SqlCad & "AND FAMILIA = '" & StrFiltro & "' "
                Case 1
                    SqlCad = SqlCad & "AND SUBFAMILIA = '" & StrFiltro & "' "
            End Select
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "FAMILIA, SUBFAMILIA, NOMPRODUCTO"
    
    With dbgInventario
        .Dataset.Close
                
        .Columns.DestroyColumns
    End With
    
    Dim gColumn As dxGridColumn
    
    With dbgInventario
        'Columna Familia de Producto
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .BandIndex = 0
            .Caption = "Familia"
            .DisableEditor = True
            .FieldName = "FAMILIA"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColFamilia"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 80
        End With
        
        'Columna Sub-Familia de Producto
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .BandIndex = 0
            .Caption = "Sub-Familia"
            .DisableEditor = True
            .FieldName = "SUBFAMILIA"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColSubFamilia"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 100
        End With
        
        'Columna Codigo de Producto
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taCenter
            .BandIndex = 0
            .Caption = "Codigo"
            .DisableEditor = True
            .FieldName = "CODPRODUCTO"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColCodigo"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 150
            .Visible = False
        End With
        
        'Columna Descripcion del Producto
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taLeftJustify
            .Caption = "Descripción"
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "NOMPRODUCTO"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColDescripcion"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 300
        End With
        
        'Columna Unidad de Medida
        Set gColumn = .Columns.Add(gedTextEdit)
        
        With gColumn
            .Alignment = taCenter
            .Caption = "U.M."
            .BandIndex = 0
            .DisableEditor = True
            .FieldName = "UM"
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            .ObjectName = "ColUM"
            .SummaryFooterType = cstCount
            .SummaryFooterFormat = " "
            .Width = 50
        End With
        
        'Columna Stock Sistema
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 1
            .Caption = "Stock Sistema"
            .Color = &HC00000
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "STOCKSISTEMA"
            .Font.Bold = True
            .FontColor = &HFFFFFF
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            '.SummaryFooterType = cstSum
            .ObjectName = "ColStockSistema"
            .SummaryFooterType = cstAvg
            '.SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Stock Fisico
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 1
            .Caption = "Stock Fisico"
            .Color = RGB(255, 235, 156)
            .DecimalPlaces = 2
            .DisableEditor = False
            .FieldName = "STOCKFISICO"
            .Font.Bold = True
            .FontColor = RGB(156, 101, 0)
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            '.SummaryFooterType = cstSum
            .ObjectName = "ColStockFisico"
            .SummaryFooterType = cstAvg
            '.SummaryFooterFormat = " "
            .Width = 70
        End With
        
        'Columna Diferencia
        Set gColumn = .Columns.Add(gedSpinEdit)
        
        With gColumn
            .Alignment = taRightJustify
            .BandIndex = 4
            .Caption = "Diferencia"
            .Color = RGB(230, 185, 184)
            .DecimalPlaces = 2
            .DisableEditor = True
            .FieldName = "DIFERENCIA"
            .Font.Bold = True
            .FontColor = RGB(0, 0, 0)
            .Font.Charset = 0
            .HeaderAlignment = taCenter
            '.SummaryFooterType = cstSum
            .ObjectName = "ColDiferencia"
            .SummaryFooterType = cstAvg
            '.SummaryFooterFormat = " "
            .Width = 70
        End With
        
        If CBool(tlbInventario.Tools("Cerrar").State) Then
            'Columna Costo Promedio
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "C. Promedio"
                .Color = vbWhite 'RGB(230, 185, 184)
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COSTOPROMEDIO"
                .Font.Bold = True
                .FontColor = vbBlack 'RGB(0, 0, 0)
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColCostoPromedio"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Sistema Valorizado
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Valorizado Sist."
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKSISTEMAVALOR"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockSistemaValor"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Stock Fisico Valorizado
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Valorizado Fisico"
                .Color = RGB(255, 235, 156)
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKFISICOVALOR"
                .Font.Bold = True
                .FontColor = RGB(156, 101, 0)
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                '.SummaryFooterType = cstSum
                .ObjectName = "ColStockFisicoValor"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
        End If
        
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp

        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctKeyset
        .Dataset.ADODataset.LockType = ltOptimistic
        
        .Dataset.ADODataset.CommandText = SqlCad
        
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODPRODUCTO"
        
        .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = .Dataset.RecordCount & " registro(s) encontrado(s)."
    End With
    
    Exit Sub
errListarGrilla:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarGrilla"
    
    Err.Clear
End Sub

Private Sub listarFiltro(ByVal intOpcion As Integer)
    On Error GoTo errListarFiltro
    
    Dim rstFiltro As New ADODB.Recordset
    
    If rstFiltro.State = 1 Then rstFiltro.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
            
    Select Case intOpcion
        Case 0
            SqlCad = SqlCad & "FAMILIA "
        Case 1
            SqlCad = SqlCad & "SUBFAMILIA "
    End Select
    
    SqlCad = SqlCad & "AS FILTRO "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILTOMAINVENTARIO "
    SqlCad = SqlCad & "GROUP BY "
    
    Select Case intOpcion
        Case 0
            SqlCad = SqlCad & "FAMILIA "
        Case 1
            SqlCad = SqlCad & "SUBFAMILIA "
    End Select
    
    SqlCad = SqlCad & "ORDER BY "
    
    Select Case intOpcion
        Case 0
            SqlCad = SqlCad & "FAMILIA"
        Case 1
            SqlCad = SqlCad & "SUBFAMILIA"
    End Select
    
    rstFiltro.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    cmbFiltro.Clear
    
    If Not rstFiltro.EOF Then
        rstFiltro.MoveFirst
        
        cmbFiltro.AddItem "(*) Todos"
        
        Do While Not rstFiltro.EOF
            cmbFiltro.AddItem Trim(rstFiltro!FILTRO & "")
            
            rstFiltro.MoveNext
        Loop
            cmbFiltro.ListIndex = 0
    End If
    
    Exit Sub
errListarFiltro:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarFiltro"
    
    Err.Clear
End Sub


Private Sub limpiarCajas()
    Me.Caption = "Toma de Inventario"
    
    txtObservacion.Text = vbNullString
    dtpFecha.value = Format(Date, "Short Date")
    txtValeIngreso.Text = vbNullString: lblValeIngresoExterno.Caption = "< ID Externo >"
    txtValeSalida.Text = vbNullString: lblValeSalidaExterno.Caption = "< ID Externo >"
    
    txtBusqueda.Text = vbNullString
    optFiltro(0).value = True
    
    dbgInventario.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILTOMAINVENTARIO"
    
    tlbInventario.Tools("Guardar").Enabled = True
    tlbInventario.Tools("Eliminar").Enabled = True
    tlbInventario.Tools("Cerrar").Enabled = True
    tlbInventario.Tools("Imprimir").Enabled = True
    
    fraDatos.Enabled = True
    fraBusqueda.Enabled = True
    fraFiltro.Enabled = True
    tlbInventario.Enabled = True
End Sub

Private Sub consultarTomaInventario()
    Set objTomaInventario = New ClsTomaInventario
    
    Me.MousePointer = vbHourglass
    
    limpiarCajas
    
    dbgInventario.Dataset.Close
    
    With objTomaInventario
        .inicializarEntidades
        
        .AnnoTI = strAnno
        .MesTI = strMes
        .CodigoAlmacen = strCodAlmacen
        
        If .obtenerTomaInventario Then
            txtObservacion.Text = .Observacion
            dtpFecha.value = Format(.Fecha, "Short Date")
            
            txtValeIngreso.Text = .ValeIngreso: lblValeIngresoExterno.Caption = .ValeIngresoExterno
            txtValeSalida.Text = .ValeSalida: lblValeSalidaExterno.Caption = .ValeSalidaExterno
            
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("Cerrar").State = IIf(.CierreInventario, ssChecked, ssUnchecked)
            
            If .CierreInventario Then
                Me.Caption = Me.Caption & " ( CERRADO )"
                
                tlbInventario.Tools("Guardar").Enabled = Not .CierreInventario
                tlbInventario.Tools("Eliminar").Enabled = Not .CierreInventario
            End If
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0015'") = "0015" Then
                tlbInventario.Tools.ITEM("Cerrar").Enabled = True
            End If
            
            If Not CBool(tlbInventario.Tools.ITEM("Cerrar").State) Then
                tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ce&rrar"
            Else
                tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ab&rir"
            End If
            
            bolObviarCierre = False
            
            listarGrilla
            
            optFiltro_Click 0
        Else
            listarGrilla
            
'            adicionarItemVale
        End If
    End With
    
    Me.MousePointer = vbDefault
    
    Set objTomaInventario = Nothing
End Sub

Private Sub validarCajas()
    On Error Resume Next
    
    If dbgInventario.Dataset.State = dsEdit Or dbgInventario.Dataset.State = dsInsert Then
        dbgInventario.Dataset.Post
    Else
        If dbgInventario.Dataset.RecordCount > 0 Then
            dbgInventario.Dataset.Edit
            
            dbgInventario.Dataset.Post
        End If
    End If
    
    If Not IsDate(dtpFecha.value) Then
        MsgBox "Fecha invalida, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        dtpFecha.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar la Toma de Inventario?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            guardarTomaInventario
    End If
End Sub

Private Sub guardarTomaInventario()
    Dim rstTemporal As New ADODB.Recordset
    
    Set objTomaInventario = New ClsTomaInventario
    
    dbgInventario.Dataset.Close
    
    With objTomaInventario
        .inicializarEntidades
        
        
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        .AnnoTI = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
        .MesTI = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
        
        .Fecha = Format(dtpFecha.value, "Short Date")
        .Observacion = Trim(txtObservacion.Text)
        
        .ValeIngreso = Trim(txtValeIngreso.Text)
        .ValeIngresoExterno = Trim(lblValeIngresoExterno.Caption)
        .ValeSalida = Trim(txtValeSalida.Text)
        .ValeSalidaExterno = Trim(lblValeSalidaExterno.Caption)
        
        .CierreInventario = IIf(CBool(tlbInventario.Tools.ITEM("Cerrar").State), True, False)
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If .guardarTomaInventario Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
'            .SQLSelectAlter = vbNullString
'            .SQLSelectAlter = .SQLSelectAlter & "DELETE "
'            .SQLSelectAlter = .SQLSelectAlter & "FROM "
'            .SQLSelectAlter = .SQLSelectAlter & "(("
'            .SQLSelectAlter = .SQLSelectAlter & "H3TOMAINV AS TI "
'            .SQLSelectAlter = .SQLSelectAlter & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = TI.F5CODPRO) "
'            .SQLSelectAlter = .SQLSelectAlter & "LEFT JOIN SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
'            .SQLSelectAlter = .SQLSelectAlter & "LEFT JOIN SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
'            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
'            .SQLSelectAlter = .SQLSelectAlter & "TI.F2CODALM = '" & .CodigoAlmacen & "' AND "
'            .SQLSelectAlter = .SQLSelectAlter & "TI.F4ANNO = '" & .AnnoTI & "' AND "
'            .SQLSelectAlter = .SQLSelectAlter & "TI.F4MES = '" & .MesTI & "'"
'
'                If left(Trim(cmbFiltro.Text), 3) <> "(*)" Then
'                    If CBool(optFiltro(0).value) Then
'                        .SQLSelectAlter = .SQLSelectAlter & " AND FAM.F7DESCON = '" & Trim(cmbFiltro.Text) & "'"
'                    Else
'                        .SQLSelectAlter = .SQLSelectAlter & " AND SFAM.F7DESCON = '" & Trim(cmbFiltro.Text) & "'"
'                    End If
'                End If
'
'            cnn_dbbancos.Execute .SQLSelectAlter
'
'            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            fraProceso.Visible = True
            pgbProceso1.value = 0
            lblProceso1.Caption = "Guardando Toma de Inventarios..."
            
            If rstTemporal.State = 1 Then rstTemporal.Close
            
            .SQLSelectAlter = vbNullString
            .SQLSelectAlter = .SQLSelectAlter & "SELECT "
            .SQLSelectAlter = .SQLSelectAlter & "* "
            .SQLSelectAlter = .SQLSelectAlter & "FROM "
            .SQLSelectAlter = .SQLSelectAlter & "TMPUTILTOMAINVENTARIO "
            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
            .SQLSelectAlter = .SQLSelectAlter & "TRIM(CODPRODUCTO & '') <> '' "
                
                If left(Trim(cmbFiltro.Text), 3) <> "(*)" Then
                    If CBool(optFiltro(0).value) Then
                        .SQLSelectAlter = .SQLSelectAlter & "AND FAMILIA = '" & Trim(cmbFiltro.Text) & "'"
                    Else
                        .SQLSelectAlter = .SQLSelectAlter & "AND SUBFAMILIA = '" & Trim(cmbFiltro.Text) & "'"
                    End If
                End If
            
            .SQLSelectAlter = .SQLSelectAlter & "ORDER BY "
            .SQLSelectAlter = .SQLSelectAlter & "FAMILIA, SUBFAMILIA, NOMPRODUCTO"
            
            rstTemporal.Open .SQLSelectAlter, cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTemporal.EOF Then
                rstTemporal.MoveFirst
                
                pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
                pgbProceso1.value = 0
                lblProceso1.Caption = "Registrando Toma de Inventario..."
                
                Do While Not rstTemporal.EOF
                    .SQLSelectAlter = vbNullString
                    .SQLSelectAlter = .SQLSelectAlter & "DELETE "
                    .SQLSelectAlter = .SQLSelectAlter & "FROM "
                    .SQLSelectAlter = .SQLSelectAlter & "H3TOMAINV "
                    .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                    .SQLSelectAlter = .SQLSelectAlter & "F2CODALM = '" & .CodigoAlmacen & "' AND "
                    .SQLSelectAlter = .SQLSelectAlter & "F4ANNO = '" & .AnnoTI & "' AND "
                    .SQLSelectAlter = .SQLSelectAlter & "F4MES = '" & .MesTI & "' AND "
                    .SQLSelectAlter = .SQLSelectAlter & "F5CODPRO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                    
        
                    cnn_dbbancos.Execute .SQLSelectAlter
        
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                    .StockSistema = Val(rstTemporal!StockSistema & "")
                    .StockFisico = Val(rstTemporal!StockFisico & "")
                    .Diferencia = Val(rstTemporal!Diferencia & "")
                    
                    .guardarTomaInvDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    DoEvents
                    
                    fraProceso.Visible = True
                    pgbProceso1.value = pgbProceso1.value + 1
                    lblProceso1.Caption = "Registrando Toma de Inventario [" & FormatPercent(pgbProceso1.value / pgbProceso1.Max, 3) & "]: " & Trim(rstTemporal!FAMILIA & "") & " / " & Trim(rstTemporal!SUBFAMILIA & "") & " / " & left(Trim(rstTemporal!NOMPRODUCTO & ""), 100) & " (" & Trim(rstTemporal!um & "") & ") "
                    
                    rstTemporal.MoveNext
                Loop
            End If
            
            MsgBox "Toma de Inventario actualizada." & vbNewLine & _
                    "Productos actualizados: " & pgbProceso1.Max, vbInformation + vbOKOnly, App.ProductName
            
            fraProceso.Visible = False
            pgbProceso1.value = 0
            lblProceso1.Caption = vbNullString
            
            strCodAlmacen = .CodigoAlmacen
            strAnno = .AnnoTI
            strMes = .MesTI
            
            consultarTomaInventario
        Else
            listarGrilla
        End If
    End With
    
    Set objTomaInventario = Nothing
End Sub

Private Sub eliminarTomaInventario()
    Set objTomaInventario = New ClsTomaInventario
    
    With objTomaInventario
        .inicializarEntidades
        
        
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        .AnnoTI = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
        .MesTI = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
        
        If Not objTomaInventario.obtenerTomaInventario Then
            MsgBox "Toma de Inventario no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea eliminar la Toma de Inventario?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarTomaInventario Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodAlmacen = .CodigoAlmacen
                strAnno = .AnnoTI
                strMes = .MesTI
                
                consultarTomaInventario
                
                MsgBox "Toma de Inventario eliminado.", vbInformation + vbOKOnly, App.ProductName
            End If
        End If
    End With
    
    Set objTomaInventario = Nothing
End Sub

Private Sub cerrarTomaInventario()
    Set objTomaInventario = New ClsTomaInventario
    
    With objTomaInventario
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        .AnnoTI = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
        .MesTI = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
        
        If Not .verificarExistencia Then
            MsgBox "Toma de Inventario no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        .inicializarEntidades
    End With
    
    
    
    If MsgBox("Asegurese de haber guardado la Toma de Inventario." & vbNewLine & _
                "¿Desea cerrar la Toma de Inventario?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
              
        bolObviarCierre = True
        
        tlbInventario.Tools.ITEM("Cerrar").State = ssUnchecked
        
        If Not CBool(tlbInventario.Tools.ITEM("Cerrar").State) Then
            tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ce&rrar"
        Else
            tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ab&rir"
        End If
        
        bolObviarCierre = False
        
        Exit Sub
    End If
    
    If dbgInventario.Dataset.State = dsEdit Or dbgInventario.Dataset.State = dsInsert Then
        dbgInventario.Dataset.Post
    Else
        If dbgInventario.Dataset.RecordCount > 0 Then
            dbgInventario.Dataset.Edit
            
            dbgInventario.Dataset.Post
        End If
    End If
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblItem As Double
    
    Set objTomaInventario = New ClsTomaInventario
    
    dbgInventario.Dataset.Close
    
    Screen.MousePointer = vbHourglass
    
    fraDatos.Enabled = False
    fraBusqueda.Enabled = False
    fraFiltro.Enabled = False
    tlbInventario.Enabled = False
    
    'VALE DE AJUSTE DE INGRESO
    fraProceso.Visible = True
    pgbProceso1.value = 0
    lblProceso1.Caption = "Cerrando Toma de Inventarios..."
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open "SELECT * FROM TMPUTILTOMAINVENTARIO WHERE DIFERENCIA > 0 ORDER BY NOMPRODUCTO", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        With objAyudaVale
            .inicializarEntidades
            
            .TipoVale = "I"
            .NumeroVale = Trim(txtValeIngreso.Text)
            .NumeroValeExterno = Trim(lblValeIngresoExterno.Caption)
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            .CodigoOrigen = "XJ0"
            
            .Fecha = Format(dtpFecha.value, "Short Date")
            .CodigoMoneda = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l")
            .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", dtpFecha.value, "F"))
            
            If .TipoCambio = 0 Then
                .TipoCambio = 4.05
            End If
            
            .observaciones = "Toma de Inventario del Periodo " & ModUtilitario.devuelveNombreMes(Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))) & "-" & Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
            
            .ExportarVale = True
            
            .FecReg = Format(Date, "Short Date")
            .UsuReg = wusuario
            .FecMod = Format(Date, "Short Date")
            .UsuMod = wusuario
            
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
            pgbProceso1.value = 0
            lblProceso1.Caption = "Registrando Ajuste de Ingreso..."
            
            If .guardarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                .SQLSelectAlter = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                cnn_dbbancos.Execute .SQLSelectAlter
                
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                Do While Not rstTemporal.EOF
                    .inicializarEntidadesDetalle
                    
                    dblItem = dblItem + 1
                    
                    .ITEM = dblItem
                    
                    .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                    .CodigoProductoOriginal = Trim(rstTemporal!CodProducto & "")
                    .Cantidad = Abs(Val(rstTemporal!Diferencia & ""))
                    
                    .ValorVenta = .calcularCostoPromedioV2
                    .IGV = 0
                    .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                    
                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                    .IgvDol = 0
                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    DoEvents
                    
                    fraProceso.Visible = True
                    pgbProceso1.value = pgbProceso1.value + 1
                    lblProceso1.Caption = "Registrando Ajuste de Ingreso [" & FormatPercent(pgbProceso1.value / pgbProceso1.Max, 3) & "]: " & Trim(rstTemporal!FAMILIA & "") & " / " & Trim(rstTemporal!SUBFAMILIA & "") & " / " & left(Trim(rstTemporal!NOMPRODUCTO & ""), 100) & " (" & Trim(rstTemporal!um & "") & ") "
                    
                    rstTemporal.MoveNext
                Loop
                    txtValeIngreso.Text = .NumeroVale
                    
                    'Exportar el Vale
'                    If ModMilano.exportarValeAserverSQLv2(.CodigoAlmacen, .NumeroVale, lblValeIngresoExterno, lblProceso1, pgbProceso1) Then
'                        MsgBox "ID Ingreso en Sistema Externo: " & lblValeIngresoExterno.Caption & ".", vbInformation + vbOKOnly, App.ProductName
'                    End If
                    
                    MsgBox "Vale de Ingreso por Ajuste, registrado.", vbInformation + vbOKOnly, App.ProductName
            End If
        End With
    End If
    
    'VALE DE AJUSTE DE SALIDA
    fraProceso.Visible = True
    pgbProceso1.value = 0
    lblProceso1.Caption = "Cerrando Toma de Inventarios..."
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open "SELECT * FROM TMPUTILTOMAINVENTARIO WHERE DIFERENCIA < 0 ORDER BY NOMPRODUCTO", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        With objAyudaVale
            .inicializarEntidades
            
            .TipoVale = "S"
            .NumeroVale = Trim(txtValeSalida.Text)
            .NumeroValeExterno = Trim(lblValeSalidaExterno.Caption)
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            .CodigoOrigen = "XJ1"
            
            .Fecha = Format(dtpFecha.value, "Short Date")
            .CodigoMoneda = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l")
            .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", dtpFecha.value, "F"))
            
            If .TipoCambio = 0 Then
                .TipoCambio = 4.05
            End If
            
            .observaciones = "Toma de Inventario del Periodo " & ModUtilitario.devuelveNombreMes(Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))) & "-" & Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
            
            .ExportarVale = True
            
            .FecReg = Format(Date, "Short Date")
            .UsuReg = wusuario
            .FecMod = Format(Date, "Short Date")
            .UsuMod = wusuario
            
            pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
            pgbProceso1.value = 0
            lblProceso1.Caption = "Registrando Ajuste de Salida..."
            
            If .guardarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                .SQLSelectAlter = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                cnn_dbbancos.Execute .SQLSelectAlter
                
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                Do While Not rstTemporal.EOF
                    .inicializarEntidadesDetalle
                    
                    dblItem = dblItem + 1
                    
                    .ITEM = dblItem
                    
                    .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                    .CodigoProductoOriginal = Trim(rstTemporal!CodProducto & "")
                    .Cantidad = Abs(Val(rstTemporal!Diferencia & ""))
                    
                    .ValorVenta = .calcularCostoPromedioV2
                    .IGV = 0
                    .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                    
                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                    .IgvDol = 0
                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    DoEvents
                    
                    fraProceso.Visible = True
                    pgbProceso1.value = pgbProceso1.value + 1
                    lblProceso1.Caption = "Registrando Ajuste de Salida [" & FormatPercent(pgbProceso1.value / pgbProceso1.Max, 3) & "]: " & Trim(rstTemporal!FAMILIA & "") & " / " & Trim(rstTemporal!SUBFAMILIA & "") & " / " & left(Trim(rstTemporal!NOMPRODUCTO & ""), 100) & " (" & Trim(rstTemporal!um & "") & ") "
                    
                    rstTemporal.MoveNext
                Loop
                    txtValeSalida.Text = .NumeroVale
                    
                    'Exportar el Vale
'                    If ModMilano.exportarValeAserverSQLv2(.CodigoAlmacen, .NumeroVale, lblValeSalidaExterno, lblProceso1, pgbProceso1) Then
'                        MsgBox "ID Salida en Sistema Externo: " & lblValeSalidaExterno.Caption & ".", vbInformation + vbOKOnly, App.ProductName
'                    End If
                    
                    MsgBox "Vale de Salida por Ajuste, registrado.", vbInformation + vbOKOnly, App.ProductName
            End If
        End With
    End If
    
    With objTomaInventario
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        .AnnoTI = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
        .MesTI = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
        
        .ValeIngreso = Trim(txtValeIngreso.Text)
        .ValeIngresoExterno = Trim(lblValeIngresoExterno.Caption)
        .ValeSalida = Trim(txtValeSalida.Text)
        .ValeSalidaExterno = Trim(lblValeSalidaExterno.Caption)
        
        .CierreInventario = True
        
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If Not .verificarExistencia Then
            MsgBox "Toma de Inventario no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If .cerrarTomaInventario Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodAlmacen = .CodigoAlmacen
            strAnno = .AnnoTI
            strMes = .MesTI
            
            consultarTomaInventario
            
            MsgBox "Toma de Inventario CERRADO correctamente.", vbInformation + vbOKOnly, App.ProductName
        Else
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
            
            bolObviarCierre = False
        End If
    End With
    
    Set objTomaInventario = Nothing
    
    fraProceso.Visible = False
    pgbProceso1.value = 0
    lblProceso1.Caption = vbNullString
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub abrirTomaInventario()
    
    
    If MsgBox("¿Desea abrir la Toma de Inventario?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        
        bolObviarCierre = True
        
        tlbInventario.Tools.ITEM("Cerrar").State = ssChecked
        
        If Not CBool(tlbInventario.Tools.ITEM("Cerrar").State) Then
            tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ce&rrar"
        Else
            tlbInventario.Tools("Cerrar").ChangeAll ssChangeAllName, "Ab&rir"
        End If
        
        bolObviarCierre = False
        
        Exit Sub
    End If
    
    If dbgInventario.Dataset.State = dsEdit Or dbgInventario.Dataset.State = dsInsert Then
        dbgInventario.Dataset.Post
    Else
        If dbgInventario.Dataset.RecordCount > 0 Then
            dbgInventario.Dataset.Edit
            
            dbgInventario.Dataset.Post
        End If
    End If
    
    If MsgBox("¿Desea conservar los Vales generados anteriormente?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        With objAyudaVale
            'ELIMINAR LOS VALES DE INGRESO
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            .NumeroVale = Trim(txtValeIngreso.Text)
            
            If objAyudaVale.verificarExistencia Then
                'If MsgBox("¿Desea eliminar el Vale con No. " & .NumeroVale & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                    If Val(lblValeIngresoExterno.Caption) > 0 Then
'                        If ModMilano.anularValeExterno("I", lblValeIngresoExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), "T"), lblProceso1, pgbProceso1) Then
'                            Me.MousePointer = vbDefault
'
'                            Exit Sub
                        End If
                    End If
                    
                    If .eliminarVale Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                        SqlCad = SqlCad & "H4TOMAINV "
                        SqlCad = SqlCad & "SET "
                        SqlCad = SqlCad & "F4VALEINGRESO = NULL, "
                        SqlCad = SqlCad & "F4VALEINGRESOEXTERNO = NULL "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "F4ANNO = '" & strAnno & "' AND "
                        SqlCad = SqlCad & "F4MES = '" & strMes & "' AND "
                        SqlCad = SqlCad & "F2CODALM = '" & strCodAlmacen & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        
                        Actualiza_Log SqlCad, StrConexDbBancos
                        'MsgBox "Registro eliminado.", vbInformation + vbOKOnly, App.ProductName
                    End If
                'End If
'            End If
            
            
            'ELIMINAR LOS VALES DE SALIDA
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            .NumeroVale = Trim(txtValeSalida.Text)
            
            If Not objAyudaVale.verificarExistencia Then
                'If MsgBox("¿Desea eliminar el Vale con No. " & .NumeroVale & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                    If Val(lblValeSalidaExterno.Caption) > 0 Then
'                        If Not ModMilano.anularValeExterno("S", lblValeSalidaExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), "T"), lblProceso1, pgbProceso1) Then
'                            Me.MousePointer = vbDefault
'
'                            Exit Sub
'                        End If
                    End If
                    
                    If .eliminarVale Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                        SqlCad = SqlCad & "H4TOMAINV "
                        SqlCad = SqlCad & "SET "
                        SqlCad = SqlCad & "F4VALESALIDA = NULL, "
                        SqlCad = SqlCad & "F4VALESALIDAEXTERNO = NULL "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "F4ANNO = '" & strAnno & "' AND "
                        SqlCad = SqlCad & "F4MES = '" & strMes & "' AND "
                        SqlCad = SqlCad & "F2CODALM = '" & strCodAlmacen & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        
                        Actualiza_Log SqlCad, StrConexDbBancos
                        
                        'MsgBox "Registro eliminado.", vbInformation + vbOKOnly, App.ProductName
                    End If
                'End If
            End If
            
            
        End With
    End If
    
    'Set objTomaInventario = New ClsTomaInventario
    
    dbgInventario.Dataset.Close
    
    Screen.MousePointer = vbHourglass
    
    Set objTomaInventario = New ClsTomaInventario
    
    With objTomaInventario
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        .AnnoTI = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
        .MesTI = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
        
        .ValeIngreso = Trim(txtValeIngreso.Text)
        .ValeIngresoExterno = Trim(lblValeIngresoExterno.Caption)
        .ValeSalida = Trim(txtValeSalida.Text)
        .ValeSalidaExterno = Trim(lblValeSalidaExterno.Caption)
        
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        .CierreInventario = False
        
        If Not .verificarExistencia Then
            MsgBox "Toma de Inventario no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If .cerrarTomaInventario Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodAlmacen = .CodigoAlmacen
            strAnno = .AnnoTI
            strMes = .MesTI
            
            consultarTomaInventario
            
            MsgBox "Toma de Inventario ABIERTO correctamente.", vbInformation + vbOKOnly, App.ProductName
        Else
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("TomarInventario").State = ssChecked
            
            bolObviarCierre = False
        End If
    End With
    
    Set objTomaInventario = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbFiltro_Click()
        listarGrilla txtBusqueda, cmbFiltro.Text, IIf(CBool(optFiltro(0).value), 0, 1)
End Sub

Private Sub dbgInventario_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "STOCKSISTEMA", "STOCKFISICO", "STOCKSISTEMAVALOR", "STOCKFISICOVALOR"
            If Val(Text) < 0 Then
                FontColor = vbRed
            End If
            
            Text = Format(Text, "#,0.0000;(#,0.0000)")
        Case "DIFERENCIA"
            If Val(Text) = 0 Then
                FontColor = vbWhite 'RGB(230, 185, 184)
            ElseIf Val(Text) < 0 Then
                FontColor = vbRed
            ElseIf Val(Text) > 0 Then
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.0000;(#,0.0000)")
    End Select
End Sub

Private Sub dbgInventario_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "STOCKSISTEMA", "STOCKFISICO", "DIFERENCIA", "COSTOPROMEDIO", "STOCKSISTEMAVALOR", "STOCKFISICOVALOR"
            If Val(Text) < 0 Then
                FontColor = vbRed
            Else
                FontColor = vbBlue
            End If
            
            Color = vbWhite
            Text = Format(Text, "#,0.0000;(#,0.0000)")
'        Case "DIFERENCIA"
'            If Val(Text) = 0 Then
'                FontColor = vbWhite 'RGB(230, 185, 184)
'            ElseIf Val(Text) < 0 Then
'                FontColor = vbRed
'            ElseIf Val(Text) > 0 Then
'                FontColor = vbBlue
'            End If
'
'            Text = Format(Text, "#,0.0000;(#,0.0000)")
    End Select
End Sub

Private Sub dbgInventario_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgInventario.Columns.FocusedColumn.FieldName
        Case "STOCKFISICO"
            With dbgInventario
                If IsNull(.Columns.ColumnByFieldName("STOCKFISICO").value) Then
                    Exit Sub
                End If
                
                If .Dataset.State = dsEdit Or .Dataset.State = dsInsert Then
                    .Dataset.Edit
                End If
                
                .Dataset.Edit
                
                .Columns.ColumnByFieldName("DIFERENCIA").value = Val(.Columns.ColumnByFieldName("STOCKFISICO").value & "") - Val(.Columns.ColumnByFieldName("STOCKSISTEMA").value & "")
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgInventario_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case dbgInventario.Columns.FocusedColumn.FieldName
        Case "STOCKFISICO"
            Select Case KeyCode
                Case vbKeyEscape
                    With dbgInventario
                        If .Dataset.State = dsEdit Or .Dataset.State = dsInsert Then
                            .Dataset.Edit
                        End If
                        
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("STOCKFISICO").value = Null
                        .Columns.ColumnByFieldName("DIFERENCIA").value = Null
                        
                        .Dataset.Post
                    End With
            End Select
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    
        listarAlmacen
        
        listarAnno
        
        listarMes
    
    strCodAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
    strAnno = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
    strMes = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
    

        consultarTomaInventario
    
    Exit Sub
errLoad:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation, App.ProductName & " - ListarAlmacen"
    
    Err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgInventario.Move 0, 300 + fraDatos.Height + fraBusqueda.Height, Me.ScaleWidth, (Me.ScaleHeight - (300 + fraDatos.Height + fraBusqueda.Height))
End Sub

Private Sub optFiltro_Click(Index As Integer)
    
        listarFiltro Index
        
        listarGrilla txtBusqueda, cmbFiltro.Text, IIf(CBool(optFiltro(0).value), 0, 1)
End Sub

Private Sub tlbInventario_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Almacen"
            
                listarAnno
            
        Case "Anno"
            
                listarMes
            
        Case "Mes"
            strCodAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            strAnno = Trim(tlbInventario.Tools.ITEM("Anno").ComboBox.Text)
            strMes = Trim(right(tlbInventario.Tools.ITEM("Mes").ComboBox.Text, 2))
            
            
                consultarTomaInventario
            
    End Select
End Sub

Private Sub tlbInventario_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Guardar"
            Me.MousePointer = vbHourglass
            
            validarCajas
            
            Me.MousePointer = vbDefault
        Case "Eliminar"
            Me.MousePointer = vbHourglass
            
            
                eliminarTomaInventario
            
            Me.MousePointer = vbDefault
        Case "Cerrar"
            If bolObviarCierre Then
                Exit Sub
            End If
            

                If CBool(Tool.State) Then
                    cerrarTomaInventario
                Else
                    abrirTomaInventario
                End If
        Case "Imprimir"
            
            With rptTomaInventarioFormato
                .GrupoOpcion = IIf(CBool(optFiltro(0).value), 0, 1)
                .GrupoCadena = IIf(CBool(optFiltro(0).value), optFiltro(0).Caption, optFiltro(1).Caption)
                .FiltroSensitivo = txtBusqueda.Text
                .GrupoFiltro = cmbFiltro.Text
                
                .Show 1
            End With
        Case "ExportarExcel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgInventario
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave

                If .FileName <> vbNullString Then
                    dbgInventario.m.ExportToXLS .FileName

                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Salir"
            If dbgInventario.Dataset.State = dsEdit Then
                dbgInventario.Dataset.Post
            Else
                If dbgInventario.Dataset.RecordCount > 0 Then
                    dbgInventario.Dataset.Edit
                    
                    dbgInventario.Dataset.Post
                End If
            End If
            
            Unload Me
    End Select
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn

                listarGrilla txtBusqueda, cmbFiltro.Text, IIf(CBool(optFiltro(0).value), 0, 1)
    End Select
End Sub





