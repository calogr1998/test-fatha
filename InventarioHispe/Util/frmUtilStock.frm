VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilStock 
   Caption         =   "Stock  de Productos"
   ClientHeight    =   8895
   ClientLeft      =   555
   ClientTop       =   1755
   ClientWidth     =   14205
   Icon            =   "frmUtilStock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin VB.Frame fraStockAl 
      Caption         =   " Stock al..."
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
      Left            =   9720
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   129826817
         CurrentDate     =   41939
      End
   End
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
      Tools           =   "frmUtilStock.frx":058A
      ToolBars        =   "frmUtilStock.frx":39E6
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
      Left            =   11640
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox chkActivarFiltro 
         Caption         =   "Activar Auto-filtros"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
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
      Width           =   6255
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   6015
         _ExtentX        =   10610
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
         Width           =   6015
      End
      Begin VB.TextBox txtNroPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   7770
      Left            =   120
      OleObjectBlob   =   "frmUtilStock.frx":3B95
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
      Left            =   6480
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
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
Attribute VB_Name = "frmUtilStock"
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
    
    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
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

Private Sub actualizarStock()
    Screen.MousePointer = vbHourglass
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKPRODUCTO"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKVALECAB"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKVALEDET"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILSTOCKORDENDET"
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKPRODUCTO IN '" & wrutatemp & "Templus.mdb' "
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "FAM.F7CODCON AS CODFAMILIA, "
    SqlCad = SqlCad & "FAM.F7DESCON AS FAMILIA, "
    SqlCad = SqlCad & "SFAM.F7CODCON AS CODSUBFAMILIA, "
    SqlCad = SqlCad & "SFAM.F7DESCON AS SUBFAMILIA, "
    SqlCad = SqlCad & "PROD.F5CODPRO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO, "
    SqlCad = SqlCad & "MED.F7SIGMED "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "((IF5PLA AS PROD "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN SF7NIVEL02 AS SFAM ON SFAM.F7CODCON = PROD.F5UBICACIO) "
    SqlCad = SqlCad & "LEFT JOIN SF7NIVEL01 AS FAM ON FAM.F7CODCON = SFAM.F7NIVEL01 "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "PROD.F5CODPRO <> '' AND "
    SqlCad = SqlCad & "PROD.F5DESCONTINUADO = 'N' AND "
    SqlCad = SqlCad & "PROD.TIENEMOVENALM = 1"
    
    abrirCnnDbBancos
    
    abrirCnTemporal
    
    cnn_dbbancos.Execute SqlCad
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKVALECAB IN '" & wrutatemp & "Templus.mdb' "
    SqlCad = SqlCad & "SELECT * FROM IF4VALES"
    
    abrirCnnDbBancos
    
    abrirCnTemporal
    
    cnn_dbbancos.Execute SqlCad
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKVALEDET IN '" & wrutatemp & "Templus.mdb' "
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES"
    
    abrirCnnDbBancos
    
    abrirCnTemporal
    
    cnn_dbbancos.Execute SqlCad
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKORDENDET IN '" & wrutatemp & "Templus.mdb' "
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "DET.*, "
    SqlCad = SqlCad & "IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR) AS FACTORCONVERSION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "((IF3ORDEN AS DET "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "DET.F4LOCAL = 'OC'"
    
    abrirCnnDbBancos
    
    abrirCnTemporal
    
    cnn_dbbancos.Execute SqlCad
    
    ModUtilitario.sWrtIni wrutatemp & strNombreFicheroConfigCPusuario, "ConfigCP", "UltimaActualizacionConsultaStock", Now
    
    tlbStock.Tools("ID_UltimaFechaActualizacion").Edit.Text = ModUtilitario.sGetINI(wrutatemp & strNombreFicheroConfigCPusuario, "ConfigCP", "UltimaActualizacionConsultaStock", "l")
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub listarStock()
    Dim strUltimoPeriodoCierre As String
    
    Screen.MousePointer = vbHourglass
    
    dbgResumen.Dataset.Close
    
    'strUltimoPeriodoCierre = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 ANNO & MES AS PERIODO", "IF3CIERREMENSUAL", vbNullString, vbNullString, vbNullString, "ANNO <> '' GROUP BY ANNO, MES ORDER BY ANNO DESC, MES DESC")
    
    'objAyudaVale.listarGrillaMovimientoProductoResumen dbgResumen, right(cmbalmacen.Text, 2), txtBusqueda.Text, Trim(txtNroPedido.Text)
    'objAyudaVale.listarGrillaMovimientoProductoResumen dbgResumen, vbNullString, txtBusqueda.Text, vbNullString, Trim(right(tlbStock.Tools.ITEM("Familia").ComboBox.Text, 4)), Trim(right(tlbStock.Tools.ITEM("SubFamilia").ComboBox.Text, 4))
    
    objAyudaVale.listarGrillaMovimientoProductoResumen dbgResumen, Trim(right(cmbAlmacen.Text, 2)), txtbusqueda.Text, vbNullString, Trim(right(tlbStock.Tools.ITEM("Familia").ComboBox.Text, 4)), Trim(right(tlbStock.Tools.ITEM("SubFamilia").ComboBox.Text, 4))
    
    'objAyudaVale.listarGrillaMovimientoProductoResumenV4 dbgResumen, _
                                                        Trim(right(cmbAlmacen.Text, 2)), _
                                                        txtBusqueda.Text, _
                                                        vbNullString, _
                                                        Trim(right(tlbStock.Tools.ITEM("Familia").ComboBox.Text, 4)), _
                                                        Trim(right(tlbStock.Tools.ITEM("SubFamilia").ComboBox.Text, 4))

    Screen.MousePointer = vbDefault
End Sub

Private Sub chkActivarFiltro_Click()
    dbgResumen.Filter.FilterActive = CBool(chkActivarFiltro.Value)
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

Private Sub dbgResumen_OnDblClick()
    For d = 0 To 25
        nSaveRecNo = dbgResumen.Dataset.RecNo
    Next
    
    With objAyudaBien
        .inicializarEntidades
        
        .Codigo = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").Value & "")
        
        .obtenerConfigBien
    End With
    
    Select Case dbgResumen.Columns.FocusedColumn.FieldName
        Case "STOCKCOMPROMETIDO"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKCOMPROMETIDO").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = Trim(right(cmbAlmacen.Text, 2))
                
                .DeshabilitarRedistribucion = False
                .NroPedidoSolicitante = vbNullString
                .CantidadMaximaParaPedido = 0
                
                .Show 1
            End With
        Case "STOCKPORLLEGARCOMP"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If

            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = Trim(right(cmbAlmacen.Text, 2))

                .DeshabilitarRedistribucion = False
                .NroPedidoSolicitante = vbNullString
                .CantidadMaximaParaPedido = 0
                
                .Show 1
            End With
        Case "STOCKLIBRE"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = Trim(right(cmbAlmacen.Text, 2))
                
                .DeshabilitarRedistribucion = False
                
                If objAyudaBien.StockMin > 0 Then
                    If Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "") <= objAyudaBien.StockMin Then
                        MsgBox "Imposible aplicar re-distribución, se ha configurado Stock Minimo = " & objAyudaBien.StockMin & " como seguridad en Almacen.", vbInformation + vbOKOnly, App.ProductName
                        
                        .DeshabilitarRedistribucion = True
                    End If
                    
                    .CantidadMaximaParaPedido = IIf(Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "") > objAyudaBien.StockMin, Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "") - objAyudaBien.StockMin, 0)
                Else
                    .CantidadMaximaParaPedido = 0
                End If
                
                .NroPedidoSolicitante = vbNullString
                
                .Show 1
            End With
        Case "STOCKPORLLEGARLIBRE"
            If Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = Trim(right(cmbAlmacen.Text, 2))
                
                .DeshabilitarRedistribucion = False
                
                If objAyudaBien.StockMin > 0 Then
                    If Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "") < objAyudaBien.StockMin Then
                        If Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Value & "") <= _
                            (objAyudaBien.StockMin - Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "")) Then
                            
                            MsgBox "Imposible aplicar re-distribución, se aguarda el Stock para completar el Stock Minimo en Almacen.", vbInformation + vbOKOnly, App.ProductName
                            
                            .DeshabilitarRedistribucion = True
                        End If
                    End If
                    
                    .CantidadMaximaParaPedido = IIf(Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Value & "") > (objAyudaBien.StockMin - Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "")), _
                                                        Val(dbgResumen.Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Value & "") - (objAyudaBien.StockMin - Val(dbgResumen.Columns.ColumnByFieldName("STOCKLIBRE").Value & "")), _
                                                        0)
                Else
                    .CantidadMaximaParaPedido = 0
                End If
                
                .NroPedidoSolicitante = vbNullString
                '.CantidadMaximaParaPedido = 0
                
                .Show 1
            End With
    End Select
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
    txtbusqueda.Text = vbNullString
    dtpHasta.Value = Format(Date, "Short Date")
        'ADD
        'txtNroPedido.Text = vbNullString
        
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
    
    listarFamilia
    
    listarSubFamilia
    
    listarAlmacenEnCombo
    
    'listarStock
    
    tlbStock.Tools("ID_UltimaFechaActualizacion").Edit.Text = ModUtilitario.sGetINI(wrutatemp & strNombreFicheroConfigCPusuario, "ConfigCP", "UltimaActualizacionConsultaStock", "l")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgResumen.Move 0, FraBusqueda.Height + 300, Me.ScaleWidth, (Me.ScaleHeight - (FraBusqueda.Height + 300))
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
        Case "Actualizar"
            dbgResumen.Dataset.Close
            
            'actualizarStock
            
            listarStock
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
    ModUtilitario.seleccionarTextoCaja txtbusqueda
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
