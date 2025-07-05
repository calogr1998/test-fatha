VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilResumenProduccion 
   Caption         =   "Resumen de Requerimientos de Producción"
   ClientHeight    =   8490
   ClientLeft      =   525
   ClientTop       =   1965
   ClientWidth     =   13515
   Icon            =   "frmUtilResumenProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13515
   WindowState     =   2  'Maximized
   Begin VB.Frame fraProductoAdd 
      Caption         =   " Producto(s) Adicional(es) en OP: "
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
      Left            =   10920
      TabIndex        =   25
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdProductoOPAdd 
         Caption         =   "Adicionar a OP"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCodProductoAdd 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtProductoAdd 
         Height          =   645
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "frmUtilResumenProduccion.frx":058A
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   " Datos de Consulta "
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
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtProducto 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNroPedido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "No. Pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblNroPedido 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.Frame fraProveedor 
      Caption         =   " Proveedor "
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
      Left            =   10920
      TabIndex        =   14
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Orden"
         Height          =   315
         Left            =   5040
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   550
         Width           =   1215
      End
      Begin VB.ComboBox cmbColocarEnOrden 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   550
         Width           =   3975
      End
      Begin VB.TextBox txtCodProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Colocar en"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblProveedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Timer tmrTemporizador 
      Left            =   0
      Top             =   7800
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
      Height          =   975
      Left            =   11040
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkProductoProveedor 
         Caption         =   "Mostrar productos de proveedor."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkProductoSeleccionado 
         Caption         =   "Mostrar productos seleccionados."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin ActiveToolBars.SSActiveToolBars tlbResumen 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   2
      ToolsCount      =   21
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilResumenProduccion.frx":0590
      ToolBars        =   "frmUtilResumenProduccion.frx":EAA6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   6705
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "frmUtilResumenProduccion.frx":ED58
      TabIndex        =   9
      Top             =   1200
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
      Height          =   855
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   6825
      Index           =   1
      Left            =   240
      OleObjectBlob   =   "frmUtilResumenProduccion.frx":101BB
      TabIndex        =   13
      Top             =   1320
      Width           =   13185
   End
   Begin MSComctlLib.ImageList imgLstEstado 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtilResumenProduccion.frx":1161E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   975
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
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
      Height          =   975
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   4095
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   120
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   212
         _Version        =   393216
         Appearance      =   0
         Max             =   25
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "frmUtilResumenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNroPedido            As String
Private strCodProveedor         As String
Private strCodProducto          As String
Private strNomProducto          As String

Private strIdOrdenProduccion    As String

Private strFichero              As String

Private bolResumenCargado       As Boolean
Private bolObviarConsulta       As Boolean

Rem Variables Adicionales para Control de Anterior/Siguiente
Dim dblFactorAncho As Double
Dim intIndiceGrilla As Integer
Dim intIndiceVisible As Integer
Dim intIndiceOculto As Integer
Dim bolRetroceso As Boolean

Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double

Public Property Let NroPedido(ByVal Value As String)
    strNroPedido = Value
End Property

Public Property Get NroPedido() As String
    NroPedido = strNroPedido
End Property

Public Property Let CodigoProveedor(ByVal Value As String)
    strCodProveedor = Value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProveedor
End Property

Public Property Let CodigoProducto(ByVal Value As String)
    strCodProducto = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Property Let NombreProducto(ByVal Value As String)
    strNomProducto = Value
End Property

Public Property Get NombreProducto() As String
    NombreProducto = strNomProducto
End Property


Private Sub listarOrdenesEnCombo()
    With objAyudaOrden
        .TipoOrden = "OC"
        .RucProveedor = Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", strCodProveedor, "T"))
        
        .listarOrdenParaSeleccion cmbColocarEnOrden, "1", True
        
        cmbColocarEnOrden.ListIndex = -1
    End With
End Sub

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

Public Function cargarResumenRequerimiento() As Boolean
    Screen.MousePointer = vbHourglass
    
    dbgResumen(0).Dataset.Close
    
    txtbusqueda.Text = vbNullString
    
'    If Not ModMilano.importarResumenRequerimientoProduccion(fraProceso, pgbProceso, strNroPedido) Then
'        bolResumenCargado = False
'
'        Unload Me
'    End If
    
    'cargarResumenRequerimiento = ModMilano.importarResumenRequerimientoProduccion(fraProceso, pgbProceso, strNroPedido, _
                                                                                    IIf(strCodProducto <> vbNullString, strCodProducto, vbNullString), _
                                                                                    IIf(strCodProducto = vbNullString, IIf(strNomProducto <> "Todos los Productos (*)", strNomProducto, vbNullString), vbNullString))
    
    cargarResumenRequerimiento = ModMilano.importarResumenRequerimientoProduccionV2(fraProceso, pgbProceso, "tmpCPResumenProduccionCompra" & wusuario, strNroPedido, _
                                                                                    IIf(strCodProducto <> vbNullString, strCodProducto, vbNullString), _
                                                                                    IIf(strCodProducto = vbNullString, IIf(strNomProducto <> "Todos los Productos (*)", strNomProducto, vbNullString), vbNullString))
    
    Screen.MousePointer = vbDefault
End Function

Public Sub verificarExistenciaCodigoInsumo()
    On Error GoTo errVerificarExistenciaCodigoInsumo
    
    Dim rstTemporal As New ADODB.Recordset
    
    Dim bolProductoNoRegistradoEnCP As Boolean
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CODPRODUCTO"

    abrirCnTemporal

    If rstTemporal.State = 1 Then rstTemporal.Close

    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Verificación de Codigos de Insumos..."

    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst

        DoEvents

        fraProceso.Caption = "Contabilizando registros..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0

        Do While Not rstTemporal.EOF
            abrirCnnDbBancos

            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(rstTemporal!CodProducto & ""), "T") = vbNullString Then
                bolProductoNoRegistradoEnCP = True

                Exit Do
            End If

            DoEvents

            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Verificación de Codigos de Insumos... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1

            rstTemporal.MoveNext
        Loop
            If bolProductoNoRegistradoEnCP Then
                importarInsumoServidorExterno fraProceso, pgbProceso
            End If
    End If

    If rstTemporal.State = 1 Then rstTemporal.Close

    Set rstTemporal = Nothing

    fraProceso.Visible = False

    Exit Sub
errVerificarExistenciaCodigoInsumo:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName

    Err.Clear
End Sub

Private Sub descargarAtencionRequerimiento()
    On Error GoTo errDescargarAtencionRequerimiento
    
    Dim rstTemporal As New ADODB.Recordset
    
    'dbgResumen(0).Dataset.Close
    
    abrirCnTemporal
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "DELETE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENATENCIONREQUERIMIENTOOP "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
    
    cnDBTemp.Execute SqlCad
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (2/4)..."
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados (2/4)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0
        
        DoEvents
        
        fraProceso.Caption = "Descargando Indicadores de Atención (2/4)..."
        
        Do While Not rstTemporal.EOF
            With objAyudaSolicitud
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .Codigo = Trim(rstTemporal!NroPedido & "")
                .CodProducto = Trim(rstTemporal!CodProducto & "")
                
                'abrirCnTemporal
                
                'cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONREQUERIMIENTOOP"
                
                abrirCnnDbBancos
                
                abrirCnTemporal
                
                .descargarResumenAtencionPorProductoResOP "F"
                
                abrirCnnDbBancos
                
                abrirCnTemporal
                
                .descargarResumenAtencionPorProductoResOP "V"
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Indicadores de Atención (2/4)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
'    DoEvents
'
'    fraProceso.Visible = True
'    fraProceso.Caption = "Ejecutando consulta (2/4)..."
'
'    With objAyudaSolicitud
'        .inicializarEntidades
'        .inicializarEntidadesDetalle
'
'        .Codigo = strNroPedido
'
'        abrirCnTemporal
'
'        cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONREQUERIMIENTOOP"
'
'        abrirCnnDbBancos
'
'        abrirCnTemporal
'
'        .descargarResumenAtencionPorProductoResOP "F"
'
'        abrirCnnDbBancos
'
'        abrirCnTemporal
'
'        .descargarResumenAtencionPorProductoResOP "V"
'
'        .inicializarEntidades
'        .inicializarEntidadesDetalle
'    End With
    
    fraProceso.Visible = False
    
    Exit Sub
errDescargarAtencionRequerimiento:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub descargarAtencionRequerimientoSQL()
    On Error GoTo errDescargarAtencionRequerimientoSQL
    
    Dim rstTemporal As New ADODB.Recordset
    
    abrirCnTemporal
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "DELETE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENATENCIONREQUERIMIENTOOP "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
    
    cnDBTemp.Execute SqlCad
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (2/4)..."
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados (2/4)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0
        
        DoEvents
        
        fraProceso.Caption = "Descargando Indicadores de Atención (2/4)..."
        
        Do While Not rstTemporal.EOF
            With objSqlAyudaSolicitud
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .Codigo = Trim(rstTemporal!NroPedido & "")
                .CodProducto = Trim(rstTemporal!CodProducto & "")
                
                .descargarResumenAtencionPorProductoResOP "F"
                
                .descargarResumenAtencionPorProductoResOP "V"
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Indicadores de Atención (2/4)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    
    Exit Sub
errDescargarAtencionRequerimientoSQL:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub descargarAtencionRequerimientoPorItem()
'    On Error GoTo errDescargarAtencionRequerimientoPorItem
'
'    Dim rstTemporal As New ADODB.Recordset
'
'    'dbgResumen(0).Dataset.Close
'
'    abrirCnTemporal
'
'    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONREQUERIMIENTOOP"
'
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "NROPEDIDO, "
'    SqlCad = SqlCad & "CODPRODUCTO "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
'
'        If strNroPedido <> vbNullString Then
'            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
'        End If
'
'        If strCodProducto <> vbNullString Then
'            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
'        End If
'
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "NROPEDIDO, "
'    SqlCad = SqlCad & "CODPRODUCTO "
'    SqlCad = SqlCad & "ORDER BY "
'    SqlCad = SqlCad & "NROPEDIDO"
'
'    abrirCnTemporal
'
'    If rstTemporal.State = 1 Then rstTemporal.Close
'
'    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
'
'    DoEvents
'
'    fraProceso.Visible = True
'    fraProceso.Caption = "Ejecutando consulta (2/4)..."
'
'    If Not rstTemporal.EOF Then
'        rstTemporal.MoveFirst
'
'        DoEvents
'
'        fraProceso.Caption = "Contabilizando registros consultados (2/4)..."
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
'        pgbProceso.value = 0
'
'        DoEvents
'
'        fraProceso.Caption = "Descargando Indicadores de Atención (2/4)..."
'
'        Do While Not rstTemporal.EOF
'            With objAyudaSolicitud
'                .inicializarEntidades
'                .inicializarEntidadesDetalle
'
'                .Codigo = Trim(rstTemporal!NroPedido & "")
'                .CodProducto = Trim(rstTemporal!CodProducto & "")
'
'                abrirCnnDbBancos
'
'                abrirCnTemporal
'
'                .descargarResumenAtencionPorProductoResOP "F"
'
'                abrirCnnDbBancos
'
'                abrirCnTemporal
'
'                .descargarResumenAtencionPorProductoResOP "V"
'
'                .inicializarEntidades
'                .inicializarEntidadesDetalle
'            End With
'
'            DoEvents
'
'            pgbProceso.value = pgbProceso.value + 1
'            fraProceso.Caption = "Descargando Indicadores de Atención (2/4)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
'
'            rstTemporal.MoveNext
'        Loop
'    End If
'
'    If rstTemporal.State = 1 Then rstTemporal.Close
'
'    Set rstTemporal = Nothing
'
'    fraProceso.Visible = False
'
'    Exit Sub
'errDescargarAtencionRequerimientoPorItem:
'    MsgBox "Nro.: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
'    'Resume
'    Err.Clear
End Sub

Private Sub cargarStockDeProducto()
    On Error GoTo errCargarStockProducto
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen(0).Dataset.Close
    
    abrirCnTemporal
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "DELETE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENSTOCKREQUERIMIENTOOP "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(NROPEDIDO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
    
    cnDBTemp.Execute SqlCad
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
    
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (3/4)..."
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0
        
        DoEvents
        
        fraProceso.Caption = "Descargando Stock de Productos (3/4)..."
        
        Do While Not rstTemporal.EOF
            With objAyudaVale
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                abrirCnnDbBancos
                
                .descargarStockProductoReqOP "CEA"
                
                abrirCnnDbBancos
                
                .descargarStockProductoReqOP "CPL"
                
                abrirCnnDbBancos
                
                .descargarStockProductoReqOP "LEA"
                
                abrirCnnDbBancos
                
                .descargarStockProductoReqOP "LPL"
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (3/4)..."
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0
        
        DoEvents
        
        fraProceso.Caption = "Actualizando Stock de Productos (3/4)..."
        
        Do While Not rstTemporal.EOF
            With objAyudaVale
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
                
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                .CompromisoEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CEA'"))
                
                .CompromisoPLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CPL'"))
                
                .LibreEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LEA'"))
                
                .LibrePLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LPL'"))
                
                .StockEAG = .CompromisoEAG + .LibreEAG
                
                .StockPLG = .CompromisoPLG + .LibrePLG
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                
                cnDBTemp.Execute SqlCad
                
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Actualizando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    pgbProceso.Value = 0
    
    Exit Sub
errCargarStockProducto:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cargarStockProductoSQL()
    On Error GoTo errCargarStockProductoSQL
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen(0).Dataset.Close
    
    abrirCnTemporal
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "DELETE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENSTOCKREQUERIMIENTOOP "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
    
    cnDBTemp.Execute SqlCad
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
        
        If strNroPedido <> vbNullString Then
            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
        End If
    
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
        End If
        
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    abrirCnTemporal
    
'    If rstTemporal.State = 1 Then rstTemporal.Close
'
'    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
'
'    DoEvents
'
'    fraProceso.Visible = True
'    fraProceso.Caption = "Ejecutando consulta (3/4)..."
'
'    If Not rstTemporal.EOF Then
'        rstTemporal.MoveFirst
'
'        DoEvents
'
'        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
'        pgbProceso.value = 0
'
'        DoEvents
'
'        fraProceso.Caption = "Descargando Stock de Productos (3/4)..."
'
'        Do While Not rstTemporal.EOF
'            With objSqlAyudaVale
'                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
'
'                .descargarStockProductoReqOP "CEA"
'
'                .descargarStockProductoReqOP "CPL"
'
'                .descargarStockProductoReqOP "LEA"
'
'                .descargarStockProductoReqOP "LPL"
'            End With
'
'            DoEvents
'
'            pgbProceso.value = pgbProceso.value + 1
'            fraProceso.Caption = "Descargando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
'
'            rstTemporal.MoveNext
'        Loop
'    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (3/4)..."
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        DoEvents
        
        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0
        
        DoEvents
        
        fraProceso.Caption = "Actualizando Stock de Productos (3/4)..."
        
        Do While Not rstTemporal.EOF
            With objSqlAyudaVale
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
                
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                .verificarStockProducto
                
'                .CompromisoEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CEA'"))
'
'                .CompromisoPLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CPL'"))
'
'                .LibreEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LEA'"))
'
'                .LibrePLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LPL'"))
'
'                .StockEAG = .CompromisoEAG + .LibreEAG
'
'                .StockPLG = .CompromisoPLG + .LibrePLG
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CODPRODUCTO = '" & .CodigoProducto & "'"
                
                cnDBTemp.Execute SqlCad
                
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Actualizando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    Set rstTemporal = Nothing
    
    fraProceso.Visible = False
    pgbProceso.Value = 0
    
    Exit Sub
errCargarStockProductoSQL:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cargarStockProductoPorItem()
'    On Error GoTo errCargarStockProductoPorItem
'
'    Dim rstTemporal As New ADODB.Recordset
'    Dim dblCantidad As Double
'
'    'dbgResumen(0).Dataset.Close
'
'    abrirCnTemporal
'
'    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENSTOCKREQUERIMIENTOOP"
'
'    abrirCnTemporal
'
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "UPDATE "
'    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
'    SqlCad = SqlCad & "SET "
'    SqlCad = SqlCad & "COMPROMISOEAG = 0, "
'    SqlCad = SqlCad & "COMPROMISOPLG = 0, "
'    SqlCad = SqlCad & "LIBREEAG = 0, "
'    SqlCad = SqlCad & "LIBREPLG = 0, "
'    SqlCad = SqlCad & "STOCKEAG = 0, "
'    SqlCad = SqlCad & "STOCKPLG = 0"
'
'    cnDBTemp.Execute SqlCad
'
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "CODPRODUCTO, "
'    SqlCad = SqlCad & "NOMPRODUCTO "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
'
'        If strNroPedido <> vbNullString Then
'            SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
'        End If
'
'        If strCodProducto <> vbNullString Then
'            SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
'        End If
'
'    SqlCad = SqlCad & "GROUP BY "
'    SqlCad = SqlCad & "CODPRODUCTO, "
'    SqlCad = SqlCad & "NOMPRODUCTO "
'    SqlCad = SqlCad & "ORDER BY "
'    SqlCad = SqlCad & "NOMPRODUCTO"
'
'    abrirCnTemporal
'
'    If rstTemporal.State = 1 Then rstTemporal.Close
'
'    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
'
'    DoEvents
'
'    fraProceso.Visible = True
'    fraProceso.Caption = "Ejecutando consulta (3/4)..."
'
'    If Not rstTemporal.EOF Then
'        rstTemporal.MoveFirst
'
'        DoEvents
'
'        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
'        pgbProceso.value = 0
'
'        DoEvents
'
'        fraProceso.Caption = "Descargando Stock de Productos (3/4)..."
'
'        Do While Not rstTemporal.EOF
'            With objAyudaVale
'                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
'
'                abrirCnnDbBancos
'
'                .descargarStockProductoReqOP "CEA"
'
'                abrirCnnDbBancos
'
'                .descargarStockProductoReqOP "CPL"
'
'                abrirCnnDbBancos
'
'                .descargarStockProductoReqOP "LEA"
'
'                abrirCnnDbBancos
'
'                .descargarStockProductoReqOP "LPL"
'            End With
'
'            DoEvents
'
'            pgbProceso.value = pgbProceso.value + 1
'            fraProceso.Caption = "Descargando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
'
'            rstTemporal.MoveNext
'        Loop
'    End If
'
'    If rstTemporal.State = 1 Then rstTemporal.Close
'
'    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
'
'    DoEvents
'
'    fraProceso.Visible = True
'    fraProceso.Caption = "Ejecutando consulta (3/4)..."
'
'    If Not rstTemporal.EOF Then
'        rstTemporal.MoveFirst
'
'        DoEvents
'
'        fraProceso.Caption = "Contabilizando registros consultados (3/4)..."
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
'        pgbProceso.value = 0
'
'        DoEvents
'
'        fraProceso.Caption = "Actualizando Stock de Productos (3/4)..."
'
'        Do While Not rstTemporal.EOF
'            With objAyudaVale
'                .inicializarEntidadesDetalle
'                .inicializarEntidadesAdicionales
'
'                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
'
'                .CompromisoEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CEA'"))
'
'                .CompromisoPLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CPL'"))
'
'                .LibreEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LEA'"))
'
'                .LibrePLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKREQUERIMIENTOOP", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LPL'"))
'
'                .StockEAG = .CompromisoEAG + .LibreEAG
'
'                .StockPLG = .CompromisoPLG + .LibrePLG
'
'                SqlCad = vbNullString
'                SqlCad = SqlCad & "UPDATE "
'                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
'                SqlCad = SqlCad & "SET "
'                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
'                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
'                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
'                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
'                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
'                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
'                SqlCad = SqlCad & "WHERE "
'                SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
'
'                    If strNroPedido <> vbNullString Then
'                        SqlCad = SqlCad & "AND NROPEDIDO = '" & strNroPedido & "' "
'                    End If
'
'                    If strCodProducto <> vbNullString Then
'                        SqlCad = SqlCad & "AND CODPRODUCTO = '" & strCodProducto & "' "
'                    End If
'
'                cnDBTemp.Execute SqlCad
'
'                .inicializarEntidadesDetalle
'                .inicializarEntidadesAdicionales
'            End With
'
'            DoEvents
'
'            pgbProceso.value = pgbProceso.value + 1
'            fraProceso.Caption = "Actualizando Stock de Productos (3/4)... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3) 'pgbProceso.value + 1
'
'            rstTemporal.MoveNext
'        Loop
'    End If
'
'    Set rstTemporal = Nothing
'
'    fraProceso.Visible = False
'    pgbProceso.value = 0
'
'    'dbgResumen(0).Dataset.Open
'
'    'dbgResumen(0).Dataset.ADODataset.Requery
'    'dbgResumen(0).Dataset.Refresh
'
'    Exit Sub
'errCargarStockProductoPorItem:
'    MsgBox "Nro.: " & Err.Number & vbNewLine & _
'            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
'    'Resume
'    Err.Clear
End Sub

Private Sub cargarProductoAtendidoPorProveedor()
    On Error GoTo errCargarProductoAtendidoPorProveedor
    
    Dim rstTemporal As New ADODB.Recordset
    Dim bolAtendidoPorProveedor As Boolean
    
    dbgResumen(0).Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        pgbProgresoBusqueda.Visible = True
        pgbProgresoBusqueda.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProgresoBusqueda.Value = 0
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
        SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"
        
        cnDBTemp.Execute SqlCad
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            With objAyudaVale
                .CodigoProveedor = strCodProveedor
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                bolAtendidoPorProveedor = .verificarProductoPorProveedor
                
                If bolAtendidoPorProveedor Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "ATENDIDOPORPROV = " & IIf(bolAtendidoPorProveedor, "TRUE", "FALSE") & ", "
                    SqlCad = SqlCad & "ATENDIDOPORPROV2 = " & IIf(bolAtendidoPorProveedor, "1", "0") & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                    
                    cnDBTemp.Execute SqlCad
                End If
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            pgbProgresoBusqueda.Value = pgbProgresoBusqueda.Value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    pgbProgresoBusqueda.Visible = False
    
    Exit Sub
errCargarProductoAtendidoPorProveedor:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cargarProductoAtendidoPorProveedorSQL()
    On Error GoTo errCargarProductoAtendidoPorProveedorSQL
    
    Dim rstTemporal As New ADODB.Recordset
    Dim bolAtendidoPorProveedor As Boolean
    
    dbgResumen(0).Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "NOMPRODUCTO"
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        pgbProgresoBusqueda.Visible = True
        pgbProgresoBusqueda.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProgresoBusqueda.Value = 0
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
        SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"
        
        cnDBTemp.Execute SqlCad
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            With objSqlAyudaVale
                .CodigoProveedor = strCodProveedor
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                
                If .verificarProductoPorProveedor Then
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "ATENDIDOPORPROV = " & IIf(bolAtendidoPorProveedor, "TRUE", "FALSE") & ", "
                    SqlCad = SqlCad & "ATENDIDOPORPROV2 = " & IIf(bolAtendidoPorProveedor, "1", "0") & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
                    
                    cnDBTemp.Execute SqlCad
                End If
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            pgbProgresoBusqueda.Value = pgbProgresoBusqueda.Value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    pgbProgresoBusqueda.Visible = False
    
    Exit Sub
errCargarProductoAtendidoPorProveedorSQL:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

'Listar Vista 1 de Requerimientos de Produccion Pendientes en Grilla (QuamtumGrid)
Private Sub cargarResumenRequerimientoVista1()
    On Error GoTo errCargarResumenRequerimientoVista1
    
    SqlCad = vbNullString
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "(RESU.NROPEDIDO & RESU.CODPRODUCTO) AS LLAVE, "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.CODPRODUCTO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
    SqlCad = SqlCad & "RESU.UM, "
    SqlCad = SqlCad & "SUM(RESU.CANTIDAD) AS CANTIDAD, "
    
    SqlCad = SqlCad & "VAL(COMPR.CANT & '') AS COMPROMISOEA, "
    SqlCad = SqlCad & "VAL(PORL.CANT & '') AS COMPROMISOPL, "
    
    SqlCad = SqlCad & "SUM(RESU.SALDO) AS SALDOACTUAL, "
    
    SqlCad = SqlCad & "RESU.COMPROMISOEAG, "
    SqlCad = SqlCad & "RESU.COMPROMISOPLG, "
    SqlCad = SqlCad & "RESU.LIBREEAG, "
    SqlCad = SqlCad & "RESU.LIBREPLG, "
    SqlCad = SqlCad & "RESU.STOCKEAG, "
    SqlCad = SqlCad & "RESU.STOCKPLG, "
    
    SqlCad = SqlCad & "IIF(SUM( VAL(FORMAT(RESU.SALDO, '#0.0000')) ) > 0, VAL(FORMAT( IIF(( SUM(RESU.CANTIDAD) - ((VAL(COMPR.CANT & '') + VAL(PORL.CANT & '') + SUM(RESU.CANTIDADPC))) ) < 0, 0, ( SUM(RESU.CANTIDAD) - ((VAL(COMPR.CANT & '') + VAL(PORL.CANT & '') + SUM(RESU.CANTIDADPC))) ) ) , '#0.0000')), 0) AS COMPRAR, "
    
    SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA, "
    SqlCad = SqlCad & "RESU.ATENDIDOPORPROV2 "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "(TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RESU "
    
    SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, VAL(FORMAT(CANTIDAD, '#0.00')) AS CANT FROM TMPUTILRESUMENATENCIONREQUERIMIENTOOP WHERE TIPO = 'F') AS COMPR "
    SqlCad = SqlCad & "ON COMPR.NROPEDIDO = RESU.NROPEDIDO AND COMPR.CODPRODUCTO = RESU.CODPRODUCTO) "
    SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, VAL(FORMAT(CANTIDAD, '#0.00')) AS CANT FROM TMPUTILRESUMENATENCIONREQUERIMIENTOOP WHERE TIPO = 'V') AS PORL "
    SqlCad = SqlCad & "ON PORL.NROPEDIDO = RESU.NROPEDIDO AND PORL.CODPRODUCTO = RESU.CODPRODUCTO "
        
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(RESU.CODPRODUCTO & '') <> '' "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            'SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
'        If CBool(chkProductoProveedor.value) Then
'            SqlCad = SqlCad & "AND RESU.ATENDIDOPORPROV = TRUE "
'        End If
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.CODPRODUCTO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
    SqlCad = SqlCad & "RESU.UM, "
    SqlCad = SqlCad & "VAL(COMPR.CANT & ''), "
    SqlCad = SqlCad & "VAL(PORL.CANT & ''), "
    
    SqlCad = SqlCad & "RESU.COMPROMISOEAG, "
    SqlCad = SqlCad & "RESU.COMPROMISOPLG, "
    SqlCad = SqlCad & "RESU.LIBREEAG, "
    SqlCad = SqlCad & "RESU.LIBREPLG, "
    SqlCad = SqlCad & "RESU.STOCKEAG, "
    SqlCad = SqlCad & "RESU.STOCKPLG, "
    
    SqlCad = SqlCad & "RESU.ATENDIDOPORPROV2 "
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO"
    
    If Not dbgResumen Is Nothing Then
        With dbgResumen(0)
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With dbgResumen(0)
            'Columna Nro Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Pedido"
                .Color = RGB(250, 192, 144)
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = IIf(Trim(txtNroPedido.Text) = vbNullString, True, False)
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Descripcion de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción del Producto"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 250
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "U.M."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Atendido por Proveedor
            Set gColumn = .Columns.Add(gedImageEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Atte."
                .DisableEditor = True
                .FieldName = "ATENDIDOPORPROV2"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAtendidoPorProv2"
                
                With .ImageColumn
                    .Images = imgLstEstado.hImageList

                    .ImageIndexes.Add ("0") 'Atendido por Proveedor
                    .Values.Add ("1")
                    .Descriptions.Add ("Atendido por Proveedor")
                    
                    .ShowDescription = False
                End With

                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 20
            End With
            
            'Columna Compromiso en Almacen a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEA"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEA"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Compromiso por Llegar a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPL"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPL"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Saldo del Producto por Atender
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Saldo"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDOACTUAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldoActual"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Compromiso Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With


            'Columna Libre en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibreEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Libre Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibrePLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPLG"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Atendida
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Recomendacion"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPRAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColComprar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Por Comprar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Comprar"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPRA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompra"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errCargarResumenRequerimientoVista1:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnTemporal
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - CargarResumenRequerimientoVista1"
    End Select
    
    Err.Clear
End Sub

'Listar Vista 2 de Requerimientos de Produccion Pendientes en Grilla (QuamtumGrid)
Private Sub cargarResumenRequerimientoVista2()
    
    On Error GoTo errCargarResumenRequerimientoVista2
    
    SqlCad = vbNullString
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "LLAVE, "
    SqlCad = SqlCad & "('"
    SqlCad = SqlCad & "[Cliente: ' & RES.CLIENTE & ']" & Space(20)
    SqlCad = SqlCad & "[No Pedido: ' & RES.NROPEDIDO & ']" & Space(20)
    SqlCad = SqlCad & "[Fec. Emision: ' & RES.FEMISION & ']" & Space(20)
    SqlCad = SqlCad & "[Fec. Entrega: ' & RES.FENTREGA & ']" & Space(20)
    SqlCad = SqlCad & "[Vendedor: ' & RES.VENDEDOR & ']"
    SqlCad = SqlCad & "') AS INFO, "
    SqlCad = SqlCad & "RES.CODPRODUCTO, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "RES.IDOP, "
    SqlCad = SqlCad & "RES.NOMPRODUCTO, "
    SqlCad = SqlCad & "RES.UM, "
    SqlCad = SqlCad & "RES.NOMPRODUCTOUM, "
    SqlCad = SqlCad & "RES.CATEGORIA, "
    SqlCad = SqlCad & "RES.NROOP, "
    SqlCad = SqlCad & "RES.MODELO, "
    SqlCad = SqlCad & "RES.COLOR, "
    SqlCad = SqlCad & "RES.CANTIDADPEDIDO, "
    SqlCad = SqlCad & "RES.DESCRIPCIONOP, "
    SqlCad = SqlCad & "RES.OBSERVACIONOP, "
    SqlCad = SqlCad & "RES.CANTIDAD, "
    SqlCad = SqlCad & "RES.SALDO AS SALDOACTUAL, "
    
    SqlCad = SqlCad & "RES.COMPROMISOEAG, "
    SqlCad = SqlCad & "RES.COMPROMISOPLG, "
    SqlCad = SqlCad & "RES.LIBREEAG, "
    SqlCad = SqlCad & "RES.LIBREPLG, "
    SqlCad = SqlCad & "RES.STOCKEAG, "
    SqlCad = SqlCad & "RES.STOCKPLG, "
    
    SqlCad = SqlCad & "RES.CANTIDADPC, "
    SqlCad = SqlCad & "RES.PROCESAR, "
    SqlCad = SqlCad & "RES.ATENDIDOPORPROV "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "RES.NROPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "RES.CODPRODUCTO = '" & strCodProducto & "' "
        
'        If txtBusqueda.Text <> vbNullString Then
'            SqlCad = SqlCad & "AND ("
'            'SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
'            SqlCad = SqlCad & "RES.NROOP LIKE '%" & txtBusqueda.Text & "%'"
'            SqlCad = SqlCad & ") "
'        End If
        
'        If CBool(chkProductoSeleccionado.value) Then
'            SqlCad = SqlCad & "AND RES.PROCESAR = TRUE "
'        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RES.NROOP"
    
    If Not dbgResumen Is Nothing Then
        
        With dbgResumen(1)
            .Dataset.Close
             
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With dbgResumen(1)
            'Columna Informacion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Información del Requerimiento"
                .DisableEditor = True
                .FieldName = "INFO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                '.Width = 80
            End With
            
            'Columna Nro Pedido
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "No. Req."
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Codigo de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Codigo"
                .DisableEditor = True
                .FieldName = "CODPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Descripción de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción de Producto"
                .DisableEditor = True
                .FieldName = "NOMPRODUCTO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProducto"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna UM de Producto
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "UM"
                .DisableEditor = True
                .FieldName = "UM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Descripción de Producto + UM
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción + UM"
                .DisableEditor = True
                .FieldName = "NOMPRODUCTOUM"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNomProductoUM"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna ID de Orden de Produccion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "ID OP"
                .DisableEditor = True
                .FieldName = "IDOP"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColIdOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = False
            End With
            
            'Columna Categoria de OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Categoria"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "CATEGORIA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCategoria"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Nro de OP
            Set gColumn = .Columns.Add(gedButtonEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .ButtonColumn.EditButtonStyle = ebsEllipsis
                '.ButtonColumn.ButtonOnly = True
                .Caption = "Nro. OP"
                .Color = vbWhite
                '.DisableEditor = True
                .FieldName = "NROOP"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
            End With
            
'            'Columna Nro de OP
'            Set gColumn = .Columns.Add(gedTextEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .BandIndex = 0
'                '.ButtonColumn.EditButtonStyle = ebsEllipsis
'                .Caption = "Nro. OP"
'                .Color = vbWhite
'                .DisableEditor = True
'                .FieldName = "NROOP"
'                .Font.Charset = 0
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColNroOP"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 70
'            End With
'
'            'Columna Boton para Cambios
'            Set gColumn = .Columns.Add(gedButtonEdit)
'
'            With gColumn
'                .Alignment = taCenter
'                .BandIndex = 0
'                .ButtonColumn.EditButtonStyle = ebsEllipsis
'                .ButtonColumn.ButtonOnly = True
'                .Caption = "Cambio"
'                .Color = vbWhite
'                .HeaderAlignment = taCenter
'                .ObjectName = "ColCambio"
'                .SummaryFooterType = cstCount
'                .SummaryFooterFormat = " "
'                .Width = 40
'            End With
            
            'Columna Modelo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Modelo"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "MODELO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColModelo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Color
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Color"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "COLOR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColColor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
            End With
            
            'Columna Cantidad de Pedido
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cant. Pedido"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDADPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Descripcion de OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Descripción OP"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "DESCRIPCIONOP"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcionOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Observacion de OP
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Observación OP"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "OBSERVACIONOP"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservacionOP"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 100
                .Visible = False
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Saldo
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Saldo"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDOACTUAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldo"
                .SummaryFooterType = cstSum
                '.SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Compromiso en Almacen a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEA"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEA"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Compromiso por Llegar a favor del Producto
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPL"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPL"
                .SummaryFooterType = cstSum
                .Width = 70
            End With

            'Columna Saldo del Producto por Atender
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Saldo"
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDOACTUAL"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldoActual"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Compromiso en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "En Almacen"
                .Color = &HC0&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Compromiso Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 2
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPROMISOPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompromisoPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With


            'Columna Libre en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "En Almacen"
                .Color = &HC000&
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibreEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Libre Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 3
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "LIBREPLG"
                .Font.Bold = False
                .FontColor = &H80000012
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLibrePLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock en Almacen General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "En Almacen"
                .Color = &HC00000
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKEAG"
                .Font.Bold = True
                .FontColor = &HFFFFFF
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockEAG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With

            'Columna Stock Por Llegar General
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 4
                .Caption = "Por Llegar"
                .Color = &H80FFFF
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "STOCKPLG"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColStockPLG"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
            End With
            
            'Columna Cantidad Por Comprar
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 5
                .Caption = "Cant. Req."
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                '.DisableEditor = True
                .FieldName = "CANTIDADPC"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidadPC"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            'Columna Procesar
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "O.K."
                .FieldName = "PROCESAR"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColProcesar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            'Columna Atendido por Proveedor
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "Atte. Prov."
                .FieldName = "ATENDIDOPORPROV"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColAtendidoPorProveedor"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = False
            End With
            
            abrirCnTemporal
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctKeyset
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh
            
            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFO").GroupIndex = 0
            
            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            .Columns.ColumnByFieldName("SALDOACTUAL").SummaryFooterType = cstSum
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
            
            .m.FullExpand
        End With
    End If
    
    SqlCad = vbNullString
    
    Exit Sub
errCargarResumenRequerimientoVista2:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnTemporal
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - CargarResumenRequerimientoVista2"
    End Select
    
    Err.Clear
End Sub

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    If bolEstado Then
        If Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0 Then
            Exit Sub
        End If
    End If
    
'
'    If CBool(frmUtilStockDetalle.RedistribucionEjecutada) Then Exit Sub
    
    If dbgResumen(0).Visible Then
        dbgResumen(0).Dataset.Close
    ElseIf dbgResumen(1).Visible Then
        dbgResumen(1).Dataset.Close
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "LLAVE, "
    SqlCad = SqlCad & "RES.CODPRODUCTO, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "RES.IDOP, "
    SqlCad = SqlCad & "RES.CANTIDAD, "
    SqlCad = SqlCad & "RES.SALDO, "
    SqlCad = SqlCad & "RES.CANTIDADPC, "
    SqlCad = SqlCad & "RES.PROCESAR "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "RES.NROPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "RES.CODPRODUCTO = '" & strCodProducto & "' AND "
    SqlCad = SqlCad & "RES.SALDO > 0 "
        
        If dbgResumen(0).Visible Then
            If txtbusqueda.Text <> vbNullString Then
                SqlCad = SqlCad & "AND ("
                'SqlCad = SqlCad & "NROPEDIDO LIKE '%" & txtBusqueda.Text & "%' OR "
                'SqlCad = SqlCad & "RES.NROOP LIKE '%" & txtBusqueda.Text & "%'"
                
                SqlCad = SqlCad & "RES.NOMPRODUCTO LIKE '%" & txtbusqueda.Text & "%'"
                
                SqlCad = SqlCad & ") "
            End If
        End If
        
'        If CBool(chkProductoSeleccionado.value) Then
'            SqlCad = SqlCad & "AND RES.PROCESAR = TRUE "
'        End If
    
    'SqlCad = SqlCad & "ORDER BY "
    'SqlCad = SqlCad & "RES.NROOP"
    
    abrirCnTemporal
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst
        
        pgbProgresoBusqueda.Visible = True
        pgbProgresoBusqueda.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProgresoBusqueda.Value = 0
        
        Do While Not rstTemporal.EOF
            DoEvents
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "CANTIDADPC = " & IIf(bolEstado, IIf(Val(rstTemporal!SALDO & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(rstTemporal!SALDO & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)), 0) & ", "
            SqlCad = SqlCad & "PROCESAR = " & IIf(bolEstado, "TRUE", "FALSE") & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "LLAVE = '" & Trim(rstTemporal!LLAVE & "") & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            If bolEstado Then
                tlbResumen.Tools("ID_Cantidad").Edit.Text = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) - IIf(Val(rstTemporal!SALDO & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(rstTemporal!SALDO & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))
                
                If Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0 Then
                    Exit Do
                End If
            Else
                tlbResumen.Tools("ID_Cantidad").Edit.Text = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) + Val(rstTemporal!CANTIDADPC & "")
            End If
            
            pgbProgresoBusqueda.Value = pgbProgresoBusqueda.Value + 1
            
            rstTemporal.MoveNext
        Loop
    End If
    
    If rstTemporal.State = 1 Then rstTemporal.Close
    
    Set rstTemporal = Nothing
    
    pgbProgresoBusqueda.Visible = False
    
    SqlCad = vbNullString
    
    tlbResumen.Tools("Seleccionar").Enabled = Not bolEstado
    tlbResumen.Tools("QuitarSeleccion").Enabled = bolEstado
    
    Me.MousePointer = vbDefault
    
    
    If dbgResumen(0).Visible Then
        cargarResumenRequerimientoVista1
    ElseIf dbgResumen(1).Visible Then
        cargarResumenRequerimientoVista2
    End If
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub evaluarCantidadDeCompra(ByVal dblCantidadActualDeCompra As Double, _
                                    ByVal dblCantidadIngresadaAComprar As Double, _
                                    ByVal bolOperacionDeCompra As Boolean)
    
    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(dblCantidadActualDeCompra + (dblCantidadIngresadaAComprar * IIf(bolOperacionDeCompra, -1, 1)), "#0.00")
        tlbResumen.Tools("ID_EstadoAtte").Edit.Text = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) < 0, "EXCEDENTE DE " & Abs(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)), IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0, "TOTAL", IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > dblCantidadActualDeCompra, "PARCIAL", "PENDIENTE")))
    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) >= 0, Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), 0), "#0.00")
End Sub

Private Sub chkProductoProveedor_Click()
    cargarResumenRequerimientoVista1
End Sub

Private Sub chkProductoSeleccionado_Click()
    'cargarResumenRequerimientoVista2
End Sub

Private Sub cmbAlmacen_Click()
    'cargarResumenRequerimiento
End Sub

'Procedimiento Declarado para Selección y Vista de Detalle de Registro
Private Sub dbgResumen_RowColChange(Index As Integer)
    If dbgResumen(Index).Dataset.RecordCount > 0 Then
        Select Case Index
            Case 0
                'evaluarCantidadDeCompra Val(dbgResumen(0).Columns.ColumnByFieldName("CANTIDAD").value & ""), _
                                        Val(dbgResumen(0).Columns.ColumnByFieldName("COMPROMISOEA").value & "") + _
                                        Val(dbgResumen(0).Columns.ColumnByFieldName("COMPROMISOPL").value & "") + _
                                        Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRA").value & ""), True
                
                strNroPedido = Trim(dbgResumen(0).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                strCodProducto = Trim(dbgResumen(0).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                
                evaluarCantidadDeCompra Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDAD)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'F'")) + _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'V'")) + _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), True
                                        
                If Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") = 0 And Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRA").Value & "") = 0 Then
                    tlbResumen.Tools("Seleccionar").Enabled = False
                    tlbResumen.Tools("QuitarSeleccion").Enabled = False
                Else
                    tlbResumen.Tools("Seleccionar").Enabled = IIf(Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0, True, False)
                    tlbResumen.Tools("QuitarSeleccion").Enabled = IIf(Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") = 0, True, False)
                End If
            Case 1
                If Val(dbgResumen(1).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") < Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDAD").Value & "") Then
                    txtCodProductoAdd.Locked = True: txtCodProductoAdd.BackColor = DH
                    cmdProductoOPAdd.Enabled = False
                Else
                    txtCodProductoAdd.Locked = False: txtCodProductoAdd.BackColor = HA
                    cmdProductoOPAdd.Enabled = True
                End If
                
                strNroPedido = Trim(dbgResumen(1).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                strCodProducto = Trim(dbgResumen(1).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                
                strIdOrdenProduccion = Trim(dbgResumen(1).Columns.ColumnByFieldName("IDOP").Value & "")
                
                fraProductoAdd.Caption = "Producto(s) Adicional(es) en OP: " & Trim(dbgResumen(1).Columns.ColumnByFieldName("NROOP").Value & "")
        End Select
    End If
End Sub

Private Sub cmdGenerar_Click()
    With objAyudaProveedor
        .Codigo = Trim(txtCodProveedor.Text)
        
        If Not .obtenerProveedor Then
            MsgBox "Proveedor ingresado no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            txtCodProveedor.SetFocus
            
            Exit Sub
        End If
    End With
    
    If cmbColocarEnOrden.ListIndex = -1 Then
        MsgBox "Seleccione la Orden donde desea colocar los Items a Comprar.", vbInformation + vbOKOnly, App.ProductName
        
        cmbColocarEnOrden.SetFocus
        
        Exit Sub
    End If
    
    If Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRA").SummaryFooterValue) = 0 Then
        MsgBox "No se ha consignado ningún Item para Compra, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        dbgResumen(0).SetFocus
        
        Exit Sub
    End If
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblItem As Double
    Dim strUltimaDescripcion As String
    Dim dblUltimoPrecioSinIGv As Double
    'Dim dblUltimoDescuento As Double
    
    Dim strCuentaContable As String
    
    dbgResumen(0).Dataset.Close
    dbgResumen(1).Dataset.Close
    
    cmdGenerar.Enabled = False
    
    FraBusqueda.Enabled = False
    fraOpciones.Enabled = False
    tlbResumen.Enabled = False
    fraProveedor.Enabled = False
    
    dbgResumen(0).Enabled = False
    dbgResumen(1).Enabled = False
    
    strNroPedido = vbNullString
    strCodProducto = vbNullString
    
    Me.MousePointer = vbHourglass
    
    With objAyudaOrden
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .TipoOrden = "OC"
        .NumeroOrden = Trim(Mid(cmbColocarEnOrden.Text, InStr(1, cmbColocarEnOrden.Text, "*") + 1))
        
        If Not .obtenerOrden Then
            .FechaEmision = Format(Date, "Short Date")
            .SinProveedorEspecifico = False
            .CodProveedor = Trim(txtCodProveedor.Text)
            .NomProveedor = objAyudaProveedor.NombreProveedor
            .RucProveedor = objAyudaProveedor.NumeroDocumento
            .ContactoProveedor = objAyudaProveedor.Contacto
            
            .CodTipoComprobante = objAyudaProveedor.CodTipoComprobante
            .OrdenRegularizada = False
            
            .FechaEntrega = Format(.FechaEmision, "Short Date")
            .CodigoSolicitante = wusuario
            .CodFormaPago = objAyudaProveedor.CodigoFormaPago
            .CentroCosto = vbNullString
            .LugarEntrega = wdireccion
            .PagoParcial = False
            
            .CodMoneda = objAyudaProveedor.CodigoMoneda
            .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaEmision, "F"))
            .NumeroCotizacion = vbNullString
            
            .Colocada = False
                .ColocadaUsuario = vbNullString
                .ColocadaFecha = vbNullString
                
            .Atendida = False
                .AtendidaUsuario = vbNullString
                .AtendidaFecha = vbNullString
            
            .Empresa = UCase(wnomcia)
            .Observacion = vbNullString
            
            .SUBTOTAL = 0 'Val(Format(Grid.Columns.ColumnByFieldName("F3BASEIMP").SummaryFooterValue, "0.00"))
            .TotalInafecto = 0 'Val(Format(Grid.Columns.ColumnByFieldName("F3MONINA").SummaryFooterValue, "0.00"))
            .TotalImpuesto = 0 'Val(Format(Grid.Columns.ColumnByFieldName("F3IGV").SummaryFooterValue, "0.00"))
            .TotalFacturado = 0 'Val(Format(Grid.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.00"))
            
            .FechaReg = Format(Date, "Short Date")
            .UsuarioReg = wusuario
            .FechaMod = Format(Date, "Short Date")
            .UsuarioMod = wusuario
            
            .Estado = 1
            
            If .guardarOrden Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                .SQLSelectAlter = "DELETE FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"
                
                cnn_dbbancos.Execute .SQLSelectAlter
                
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                If rstTemporal.State = 1 Then rstTemporal.Close
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "(RESU.NROPEDIDO & RESU.CODPRODUCTO) AS LLAVE, "
                SqlCad = SqlCad & "RESU.NROPEDIDO, "
                SqlCad = SqlCad & "RESU.CODPRODUCTO, "
                SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
                SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RESU "
                SqlCad = SqlCad & "GROUP BY "
                SqlCad = SqlCad & "RESU.NROPEDIDO, "
                SqlCad = SqlCad & "RESU.CODPRODUCTO, "
                SqlCad = SqlCad & "RESU.NOMPRODUCTO "
                SqlCad = SqlCad & "HAVING "
                SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) > 0 "
                SqlCad = SqlCad & "ORDER BY "
                SqlCad = SqlCad & "RESU.NROPEDIDO, "
                SqlCad = SqlCad & "RESU.NOMPRODUCTO"
                
                rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
                
                If Not rstTemporal.EOF Then
                    rstTemporal.MoveFirst
                    
                    dblItem = 0
                    
                    Do While Not rstTemporal.EOF
                        .inicializarEntidadesDetalle
                        
                        'Obtener configuracion de Producto
                        With objAyudaBien
                            .inicializarEntidades
                            
                            .Codigo = Trim(rstTemporal!CodProducto & "")
                            
                            .obtenerConfigBien
                        End With
                        
                        strCuentaContable = vbNullString
                        
                        Select Case .TipoOrden
                            Case "OC"
                                Select Case objAyudaProveedor.OrigenProveedor
                                    Case "N"
                                        If objAyudaBien.CtaContable = vbNullString Then
'                                            MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                                    objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                                    "ya que no tiene configurado su Cuenta Contable para Proveedores Nacionales." & vbNewLine & vbNewLine & _
'                                                    "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                            
                                        Else
                                            strCuentaContable = objAyudaBien.CtaContable
                                        End If
                                    Case "E"
                                        If objAyudaBien.CtaContableImportacion = vbNullString Then
'                                            MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                                    objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                                    "ya que no tiene configurado su Cuenta Contable para Proveedores Extranjeros." & vbNewLine & vbNewLine & _
'                                                    "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
                                        Else
                                            strCuentaContable = objAyudaBien.CtaContableImportacion
                                        End If
                                End Select
                            Case "OS"
                                If objAyudaBien.CtaContable = vbNullString Then
'                                    MsgBox "Imposible adicionar el servicio: " & vbNewLine & _
'                                            objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                            "ya que no tiene configurado su Cuenta Contable." & vbNewLine & vbNewLine & _
'                                            "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
                                Else
                                    strCuentaContable = objAyudaBien.CtaContable
                                End If
                        End Select
                        
                        If strCuentaContable <> vbNullString Then
                            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T") = vbNullString Then
                                With objAyudaGasto
                                    .inicializarEntidades
                                    
                                    .Codigo = vbNullString
                                    .Base = "G"
                                    .CuentaContable = strCuentaContable
                                    .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", strCuentaContable, "T")
                                    .TipoGasto = "P"
                                    .Moneda = objAyudaOrden.CodMoneda
                                    .GrupoFlujo = vbNullString
                                    
                                    If .guardarGasto Then
                                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    End If
                                    
                                    .inicializarEntidades
                                End With
                            End If
                            
                            'Obtener la Descripción del Producto en la Ultima Compra (Ordenes de Compra)
                            With objAyudaOrden
                                .CodProveedor = objAyudaProveedor.NumeroDocumento
                                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                                
                                strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
                            End With
                            'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                            With objAyudaVale
                                .CodigoProveedor = objAyudaProveedor.NumeroDocumento
                                .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                                
                                .obtenerUltimoPrecioSinIgvProductoDeProveedor
                                
                                Select Case .CodigoMoneda
                                    Case "S"
                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(objAyudaOrden.CodMoneda = "S", 1, objAyudaOrden.TipoCambio)), "#0.0000")
                                    Case Else
                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(objAyudaOrden.CodMoneda = "D", 1, objAyudaOrden.TipoCambio)), "#0.0000")
                                End Select
                            End With
                        
                            dblItem = dblItem + 1
                            
                            .ITEM = dblItem
                            .Requerimiento = Trim(rstTemporal!NroPedido & "")
                            .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                            .CodigoFabricante = vbNullString
                            .NombreProducto = IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, objAyudaBien.Descripcion)
                            .NombreProductoInterno = Trim(rstTemporal!NOMPRODUCTO & "")
                            .CodigoUM = objAyudaBien.CodUM
                            
                            'Para Calculo por Item
                            .PorcentajeImpuesto = IIf(.CodTipoComprobante = "02", gretenc, wwigv) / 100
                            .SignoImpuesto = IIf(.CodTipoComprobante = "02", -1, 1)
                            
                            .Cantidad = Val(rstTemporal!compra & "")
                            .CantidadMaxima = Val(rstTemporal!compra & "")
                            
                            .PorcentajeDemasia = objAyudaBien.PorcentajeDemasia
                            
                            .PrecioSinImpuesto = dblUltimoPrecioSinIGv
                            .PrecioConImpuesto = 0
                            '.PrecioNetoSinImpuesto = 0
                            
                            .PorcentajeDscto = 0
                            .TotalDscto = 0
                            
                            .Afecto = CBool(objAyudaBien.Afecto)
                            
                            .calculosPorItem
                            
                            .PrecioSinImpuesto = .PrecioSinImpuesto
                            .PrecioConImpuesto = .PrecioConImpuesto
                            .PorcentajeDscto = Val(Format(.PorcentajeDscto * 100, "#0.00"))
                            .CantidadFinal = .CantidadFinal
                            .TotalDscto = .TotalDscto
                            
                            .PrecioNetoSinImpuesto = .PrecioNetoSinImpuesto
                            
                            .BasePorItem = .BasePorItem
                            .ImpuestoPorItem = .ImpuestoPorItem
                            .TotalPorItem = .TotalPorItem
                            
                            'Acumulacion de Totales para Cabecera
                            .SUBTOTAL = .SUBTOTAL + .BasePorItem
                            .TotalInafecto = .TotalInafecto + .ExoneradoPorItem
                            .TotalImpuesto = .TotalImpuesto + .ImpuestoPorItem
                            .TotalFacturado = .TotalFacturado + .TotalPorItem
                            
                            .CodigoColor = vbNullString
                            .ObservacionPorItem = vbNullString
                            
                            .CodigoGasto = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T")
                            .CuentaContable = strCuentaContable
                            
                            .guardarOrdenDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        End If
    
                        rstTemporal.MoveNext
                    Loop
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                        SqlCad = SqlCad & "IF4ORDEN "
                        SqlCad = SqlCad & "SET "
                        SqlCad = SqlCad & "F4BASIMP = " & .SUBTOTAL & ", "
                        SqlCad = SqlCad & "F4MONINA = " & .TotalInafecto & ", "
                        SqlCad = SqlCad & "F4IGV = " & .TotalImpuesto & ", "
                        SqlCad = SqlCad & "F4MONTO = " & .TotalFacturado & " "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "F4LOCAL = '" & .TipoOrden & "' AND "
                        SqlCad = SqlCad & "F4NUMORD = '" & .NumeroOrden & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        
                        Actualiza_Log SqlCad, StrConexDbBancos
                End If
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    exportarOrdenSQL .TipoOrden, .NumeroOrden
                End If
                
                MsgBox "Orden Registrada con No." & .NumeroOrden & ".", vbInformation + vbOKOnly, App.ProductName
                
                txtCodProveedor.Text = vbNullString: lblProveedor.Caption = vbNullString
                
                listarOrdenesEnCombo
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "CANTIDADPC = 0, "
                SqlCad = SqlCad & "PROCESAR = FALSE "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CANTIDADPC > 0"

                abrirCnTemporal

                cnDBTemp.Execute SqlCad
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
                SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"
                
                abrirCnTemporal
    
                cnDBTemp.Execute SqlCad
                
                'dbgResumen(0).Dataset.Close
                
                descargarAtencionRequerimiento
            End If
        Else
            
            If rstTemporal.State = 1 Then rstTemporal.Close
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "(RESU.NROPEDIDO & RESU.CODPRODUCTO) AS LLAVE, "
            SqlCad = SqlCad & "RESU.NROPEDIDO, "
            SqlCad = SqlCad & "RESU.CODPRODUCTO, "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
            SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RESU "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "RESU.NROPEDIDO, "
            SqlCad = SqlCad & "RESU.CODPRODUCTO, "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO "
            SqlCad = SqlCad & "HAVING "
            SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) > 0 "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "RESU.NROPEDIDO, "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO"
            
            rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTemporal.EOF Then
                rstTemporal.MoveFirst
                
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 VAL(ITEM & '') AS NRO", "IF3ORDEN", "F4LOCAL", .TipoOrden, "T", "AND F4NUMORD = '" & .NumeroOrden & "' ORDER BY VAL(ITEM & '') DESC"))
                
                Do While Not rstTemporal.EOF
                    .inicializarEntidadesDetalle
                    
                    'Obtener configuracion de Producto
                    With objAyudaBien
                        .inicializarEntidades
                        
                        .Codigo = Trim(rstTemporal!CodProducto & "")
                        
                        .obtenerConfigBien
                    End With
                    
                    strCuentaContable = vbNullString
                    
                    Select Case .TipoOrden
                        Case "OC"
                            Select Case objAyudaProveedor.OrigenProveedor
                                Case "N"
                                    If objAyudaBien.CtaContable = vbNullString Then
'                                        MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                                objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                                "ya que no tiene configurado su Cuenta Contable para Proveedores Nacionales." & vbNewLine & vbNewLine & _
'                                                "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                        
                                    Else
                                        strCuentaContable = objAyudaBien.CtaContable
                                    End If
                                Case "E"
                                    If objAyudaBien.CtaContableImportacion = vbNullString Then
'                                        MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                                objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                                "ya que no tiene configurado su Cuenta Contable para Proveedores Extranjeros." & vbNewLine & vbNewLine & _
'                                                "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                        
                                    Else
                                        strCuentaContable = objAyudaBien.CtaContableImportacion
                                    End If
                            End Select
                        Case "OS"
                            If objAyudaBien.CtaContable = vbNullString Then
'                                MsgBox "Imposible adicionar el servicio: " & vbNewLine & _
'                                        objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                        "ya que no tiene configurado su Cuenta Contable." & vbNewLine & vbNewLine & _
'                                        "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
                            Else
                                strCuentaContable = objAyudaBien.CtaContable
                            End If
                    End Select
                    
                    If strCuentaContable <> vbNullString Then
                        If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T") = vbNullString Then
                            With objAyudaGasto
                                .inicializarEntidades
                                
                                .Codigo = vbNullString
                                .Base = "G"
                                .CuentaContable = strCuentaContable
                                .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", strCuentaContable, "T")
                                .TipoGasto = "P"
                                .Moneda = objAyudaOrden.CodMoneda
                                .GrupoFlujo = vbNullString
                                
                                If .guardarGasto Then
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                End If
                                
                                .inicializarEntidades
                            End With
                        End If
                        
                        'Obtener la Descripción del Producto en la Ultima Compra (Ordenes de Compra)
                        With objAyudaOrden
                            .CodProveedor = objAyudaProveedor.NumeroDocumento
                            .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                            
                            strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
                        End With
                        'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                        With objAyudaVale
                            .CodigoProveedor = objAyudaProveedor.NumeroDocumento
                            .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                            
                            .obtenerUltimoPrecioSinIgvProductoDeProveedor
                            
                            Select Case .CodigoMoneda
                                Case "S"
                                    dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(objAyudaOrden.CodMoneda = "S", 1, objAyudaOrden.TipoCambio)), "#0.0000")
                                Case Else
                                    dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(objAyudaOrden.CodMoneda = "D", 1, objAyudaOrden.TipoCambio)), "#0.0000")
                            End Select
                        End With
                    
                        dblItem = dblItem + 1
                        
                        .ITEM = dblItem
                        .Requerimiento = Trim(rstTemporal!NroPedido & "")
                        .CodigoProducto = Trim(rstTemporal!CodProducto & "")
                        .CodigoFabricante = vbNullString
                        .NombreProducto = IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, objAyudaBien.Descripcion)
                        .NombreProductoInterno = Trim(rstTemporal!NOMPRODUCTO & "")
                        .CodigoUM = objAyudaBien.CodUM
                        
                        'Para Calculo por Item
                        .PorcentajeImpuesto = IIf(.CodTipoComprobante = "02", gretenc, wwigv) / 100
                        .SignoImpuesto = IIf(.CodTipoComprobante = "02", -1, 1)
                        
                        .Cantidad = Val(rstTemporal!compra & "")
                        .CantidadMaxima = Val(rstTemporal!compra & "")
                        
                        .PorcentajeDemasia = objAyudaBien.PorcentajeDemasia
                        
                        .PrecioSinImpuesto = dblUltimoPrecioSinIGv
                        .PrecioConImpuesto = 0
                        '.PrecioNetoSinImpuesto = 0
                        
                        .PorcentajeDscto = 0
                        .TotalDscto = 0
                        
                        .Afecto = CBool(objAyudaBien.Afecto)
                        
                        .calculosPorItem
                        
                        .PrecioSinImpuesto = .PrecioSinImpuesto
                        .PrecioConImpuesto = .PrecioConImpuesto
                        .PorcentajeDscto = Val(Format(.PorcentajeDscto * 100, "#0.00"))
                        .CantidadFinal = .CantidadFinal
                        .TotalDscto = .TotalDscto
                        
                        .PrecioNetoSinImpuesto = .PrecioNetoSinImpuesto
                        
                        .BasePorItem = .BasePorItem
                        .ImpuestoPorItem = .ImpuestoPorItem
                        .TotalPorItem = .TotalPorItem
                        
                        'Acumulacion de Totales para Cabecera
                        .SUBTOTAL = .SUBTOTAL + .BasePorItem
                        .TotalInafecto = .TotalInafecto + .ExoneradoPorItem
                        .TotalImpuesto = .TotalImpuesto + .ImpuestoPorItem
                        .TotalFacturado = .TotalFacturado + .TotalPorItem
                        
                        .CodigoColor = vbNullString
                        .ObservacionPorItem = vbNullString
                        
                        .CodigoGasto = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T")
                        .CuentaContable = strCuentaContable
                        
                        .guardarOrdenDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If

                    rstTemporal.MoveNext
                Loop
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "IF4ORDEN "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "F4BASIMP = " & .SUBTOTAL & ", "
                    SqlCad = SqlCad & "F4MONINA = " & .TotalInafecto & ", "
                    SqlCad = SqlCad & "F4IGV = " & .TotalImpuesto & ", "
                    SqlCad = SqlCad & "F4MONTO = " & .TotalFacturado & " "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F4LOCAL = '" & .TipoOrden & "' AND "
                    SqlCad = SqlCad & "F4NUMORD = '" & .NumeroOrden & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
            End If
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                exportarOrdenSQL .TipoOrden, .NumeroOrden
            End If
            
            MsgBox "Orden Actualizada con No." & .NumeroOrden & ".", vbInformation + vbOKOnly, App.ProductName
            
            txtCodProveedor.Text = vbNullString: lblProveedor.Caption = vbNullString
            
            listarOrdenesEnCombo
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "CANTIDADPC = 0, "
            SqlCad = SqlCad & "PROCESAR = FALSE "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CANTIDADPC > 0"

            abrirCnTemporal

            cnDBTemp.Execute SqlCad
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
            SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"
            
            abrirCnTemporal

            cnDBTemp.Execute SqlCad
            
            'dbgResumen(0).Dataset.Close
            
            abrirCnTemporal
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                descargarAtencionRequerimientoSQL
            Else
                descargarAtencionRequerimiento
            End If
        End If
        
        .inicializarEntidades
        .inicializarEntidadesDetalle
    End With
    
    cmdGenerar.Enabled = True
    
    FraBusqueda.Enabled = True
    fraOpciones.Enabled = True
    tlbResumen.Enabled = True
    fraProveedor.Enabled = True
    
    dbgResumen(0).Enabled = True
    dbgResumen(1).Enabled = True
    
    cargarResumenRequerimientoVista1
            
    dbgResumen_OnClick 0
    
    Me.MousePointer = vbDefault
End Sub

Private Sub exportarOrdenSQL(ByVal strTipoOrden As String, _
                                ByVal strNumeroOrden As String)
    
    Dim rstExportarDetSql As New ADODB.Recordset
    
    With objAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
        
        .obtenerConfigOrden
    End With
    
    With objSqlAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = objAyudaOrden.TipoOrden
        .NumeroOrden = objAyudaOrden.NumeroOrden
        
        .FechaEmision = Format(objAyudaOrden.FechaEmision, "Short Date")
        .SinProveedorEspecifico = objAyudaOrden.SinProveedorEspecifico
        .NomProveedor = Replace(objAyudaOrden.NomProveedor, "'", "´", 1)
        .RucProveedor = objAyudaOrden.RucProveedor
        .CodProveedor = objAyudaOrden.CodProveedor
        .ContactoProveedor = Replace(objAyudaOrden.ContactoProveedor, "'", "´", 1)
        
        .CodTipoComprobante = objAyudaOrden.CodTipoComprobante
        .OrdenRegularizada = objAyudaOrden.OrdenRegularizada
        
        .FechaEntrega = Format(objAyudaOrden.FechaEntrega, "Short Date")
        .CodigoSolicitante = objAyudaOrden.CodigoSolicitante
        .CodFormaPago = objAyudaOrden.CodFormaPago
        .CentroCosto = objAyudaOrden.CentroCosto
        .LugarEntrega = objAyudaOrden.LugarEntrega
        .PagoParcial = objAyudaOrden.PagoParcial

        .CodMoneda = objAyudaOrden.CodMoneda
        .TipoCambio = Format(objAyudaOrden.TipoCambio, "#.000")
        .NumeroCotizacion = objAyudaOrden.NumeroCotizacion
        
        .Colocada = objAyudaOrden.Colocada
            .ColocadaUsuario = objAyudaOrden.ColocadaUsuario
            .ColocadaFecha = IIf(.Colocada, Format(objAyudaOrden.ColocadaFecha, "Short Date"), vbNullString)
        
        .Atendida = objAyudaOrden.Atendida
            .AtendidaUsuario = objAyudaOrden.AtendidaUsuario
            .AtendidaFecha = IIf(.Atendida, Format(objAyudaOrden.AtendidaFecha, "Short Date"), vbNullString)
        
        .Empresa = objAyudaOrden.Empresa
        .Observacion = objAyudaOrden.Observacion

        .SUBTOTAL = Val(Format(objAyudaOrden.SUBTOTAL, "0.00"))
        .TotalInafecto = Val(Format(objAyudaOrden.TotalInafecto, "0.00"))
        .TotalImpuesto = Val(Format(objAyudaOrden.TotalImpuesto, "0.00"))
        .TotalFacturado = Val(Format(objAyudaOrden.TotalFacturado, "0.00"))
        
        .FechaReg = Format(objAyudaOrden.FechaReg, "Short Date")
        .UsuarioReg = objAyudaOrden.UsuarioReg
        .FechaMod = Format(objAyudaOrden.FechaMod, "Short Date")
        .UsuarioMod = objAyudaOrden.UsuarioMod
        
        .Estado = objAyudaOrden.Estado
        
        If .guardarOrden Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"

            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos

            If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
            
            rstExportarDetSql.Open "SELECT * FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly 'AND VAL(F3CANPRO & '') > 0
            
            If Not rstExportarDetSql.EOF Then
                rstExportarDetSql.MoveFirst

                Do While Not rstExportarDetSql.EOF
                    .inicializarEntidadesDetalle

                    .ITEM = Val(rstExportarDetSql!ITEM & "")
                    .Requerimiento = Trim(rstExportarDetSql!COD_SOLICITUD & "")
                    .CodigoProducto = Trim(rstExportarDetSql!F3CODPRO & "")
                    .CodigoFabricante = Trim(rstExportarDetSql!F3CODFAB & "")
                    .NombreProducto = Trim(rstExportarDetSql!F5NOMPRO & "")
                    .NombreProductoInterno = Trim(rstExportarDetSql!F5NOMPRO_ING & "")
                    .CodigoUM = Trim(rstExportarDetSql!UNIDAD & "")
                    .Cantidad = Val(rstExportarDetSql!F3CANPRO & "")
                    .CantidadMaxima = Val(rstExportarDetSql!F3CANPRO2 & "")
                    .CantidadFaltante = Val(rstExportarDetSql!F3CANFAL & "")

                    .PorcentajeDemasia = Val(rstExportarDetSql!F3PORCDEMASIA & "")

                    .PrecioSinImpuesto = Val(rstExportarDetSql!F3PRECOS & "")
                    .PrecioConImpuesto = Val(rstExportarDetSql!F3PREUNI & "")
                    .PrecioNetoSinImpuesto = Val(rstExportarDetSql!F3PRENETO & "")

                    .PorcentajeDscto = Val(rstExportarDetSql!F3PORDCT & "")
                    .TotalDscto = Val(rstExportarDetSql!F3TOTDCT & "")

                    .Afecto = IIf(Trim(rstExportarDetSql!F5AFECTO) = "*", True, False)

                    .BasePorItem = Val(rstExportarDetSql!F5VALVTA & "")
                    .ImpuestoPorItem = Val(rstExportarDetSql!F3IGV & "")
                    .TotalPorItem = Val(rstExportarDetSql!F3TOTAL & "")
                    
                    .CodigoColor = Trim(rstExportarDetSql!CODCOLOR & "")
                    .ObservacionPorItem = Trim(rstExportarDetSql!F3OBSERVA & "")

                    .ItemAjustado = CBool(rstExportarDetSql!F3AJUSTE)

                    .CodigoGasto = Trim(rstExportarDetSql!F3GASTO & "")
                    .CuentaContable = Trim(rstExportarDetSql!F3CUENTA & "")

                    .guardarOrdenDetalleOneByOne

                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos

                    rstExportarDetSql.MoveNext
                Loop
            End If
        Else
            Actualiza_Log " < Replicación BD > Importación de Orden de Compra No. " & .NumeroOrden & " fallido.", StrConexDbBancos
        End If
    End With
    
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarDetSql = Nothing
End Sub

Private Sub cmdProductoOPAdd_Click()
    If Trim(txtCodProductoAdd.Text) = vbNullString Then
        MsgBox "Seleccione el Producto.", vbInformation + vbOKOnly, App.ProductName
        
        txtCodProductoAdd.SetFocus
        
        Exit Sub
    End If
    
    If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPRODUCTO", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", Trim(txtCodProductoAdd.Text), "T", "AND IDOP = '" & strIdOrdenProduccion & "'") <> vbNullString Then
        MsgBox "Imposible adicionar el Producto, ya se encuentra registrado en la OP.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "(CANTIDAD - SALDO) AS DIFERENCIA", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "AND IDOP = '" & strIdOrdenProduccion & "'")) <> 0 Then
        MsgBox "Imposible adicionar el Producto, la OP ya fue descargada.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea adicionar el producto a la OP?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    Dim rstProductoAdd As ADODB.Recordset
    
    Set rstProductoAdd = New ADODB.Recordset
    
    dbgResumen(1).Dataset.Close
    
    Me.MousePointer = vbHourglass
    
    If ModMilano.insertarProductoEnOPServidorExterno(strIdOrdenProduccion, Trim(txtCodProductoAdd.Text)) Then
        With objAyudaBien
            .inicializarEntidades
            
            .Codigo = Trim(txtCodProductoAdd.Text)
            
            .obtenerConfigBien
        End With
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "'" & strNroPedido & objAyudaBien.Codigo & strIdOrdenProduccion & "' AS LLAVEADD, "
        SqlCad = SqlCad & "NROPEDIDO, "
        SqlCad = SqlCad & "CLIENTE, "
        SqlCad = SqlCad & "FEMISION, "
        SqlCad = SqlCad & "FENTREGA, "
        SqlCad = SqlCad & "VENDEDOR, "
        SqlCad = SqlCad & "'" & objAyudaBien.Codigo & "' AS CODPRODUCTOADD, "
        SqlCad = SqlCad & "'" & objAyudaBien.Descripcion & "' AS NOMPRODUCTOADD, "
        SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T") & "' AS UMADD, "
        SqlCad = SqlCad & "'" & objAyudaBien.Descripcion & " ( " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T") & " )' AS NOMPRODUCTOUMADD, "
        SqlCad = SqlCad & "IDOP, "
        SqlCad = SqlCad & "CATEGORIA, "
        SqlCad = SqlCad & "NROOP, "
        SqlCad = SqlCad & "MODELO, "
        SqlCad = SqlCad & "COLOR, "
        SqlCad = SqlCad & "CANTIDADPEDIDO, "
        SqlCad = SqlCad & "DESCRIPCIONOP, "
        SqlCad = SqlCad & "OBSERVACIONOP "
        
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "IDOP = '" & strIdOrdenProduccion & "' AND "
        SqlCad = SqlCad & "NROPEDIDO = '" & strNroPedido & "' AND "
        SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
        
        abrirCnTemporal
        
        If rstProductoAdd.State = 1 Then rstProductoAdd.Close
        
        rstProductoAdd.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
        
        If Not rstProductoAdd.EOF Then
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!LLAVEADD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!NroPedido & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!CLIENTE & "") & "', "
            SqlCad = SqlCad & "CVDATE('" & Trim(rstProductoAdd!FEMISION & "") & "'), "
            SqlCad = SqlCad & "CVDATE('" & Trim(rstProductoAdd!FENTREGA & "") & "'), "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!VENDEDOR & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!CODPRODUCTOADD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!NOMPRODUCTOADD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!UMADD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!NOMPRODUCTOUMADD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!IDOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!CATEGORIA & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!NroOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!Modelo & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!Color & "") & "', "
            SqlCad = SqlCad & Val(rstProductoAdd!CANTIDADPEDIDO & "") & ", "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!DESCRIPCIONOP & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstProductoAdd!OBSERVACIONOP & "") & "', "
            SqlCad = SqlCad & "1, "
            SqlCad = SqlCad & "1, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "0, "
            SqlCad = SqlCad & "FALSE, "
            SqlCad = SqlCad & "FALSE, "
            SqlCad = SqlCad & "0)"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            With objAyudaVale
                .CodigoProducto = Trim(rstProductoAdd!CODPRODUCTOADD & "")
                
                .verificarStockProducto

                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstProductoAdd!CODPRODUCTOADD & "") & "'"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad

                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales
            End With
            
            MsgBox "Producto adicionado a OP: " & Trim(rstProductoAdd!CATEGORIA & "") & " - " & Trim(rstProductoAdd!NroOP & ""), vbInformation + vbOKOnly, App.ProductName
            
            txtCodProductoAdd.Text = vbNullString
            txtProductoAdd.Text = vbNullString
        End If
    End If
    
    cargarResumenRequerimientoVista2
    
    Set rstProductoAdd = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgResumen_OnChangeNode(Index As Integer, ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    'dbgResumen_RowColChange Index
End Sub

Private Sub dbgResumen_OnClick(Index As Integer)
    If dbgResumen(Index).Dataset.RecordCount = 0 Then
        Exit Sub
    End If
    
    'strNroPedido = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").value & "")
    'strCodProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").value & "")
    
    'Select Case Index
        'Case 0
            'Select Case UCase(dbgResumen(Index).Columns.FocusedColumn.FieldName)
            '    Case "NROPEDIDO", "NOMPRODUCTO", "UM", "COMPRAR" '"CANTIDAD" ', "COMPRAR", "COMPRA"
'                    For d = 0 To 25
'                        nSaveRecNo = dbgResumen(Index).Dataset.RecNo
'                    Next
                    
                    dbgResumen_RowColChange Index
                    
                    Rem SK: DESHABILITADO TEMPORALMENTE
                    
                    'dbgResumen(Index).Columns.FocusedIndex = 2
                    
                    'dbgResumen(Index).Dataset.Close
                    
                    'descargarAtencionRequerimientoPorItem
                    
                    'cargarStockProductoPorItem
                    
                    'cargarResumenRequerimientoVista1
                    
'                    If dbgResumen(Index).Dataset.RecordCount >= nSaveRecNo Then
'                        dbgResumen(Index).Dataset.RecNo = nSaveRecNo
'                    End If
            'End Select
    'End Select
End Sub

Private Sub dbgResumen_OnCheckEditToggleClick(Index As Integer, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "PROCESAR"
            With dbgResumen(1)
                If Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") = 0 And Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) <= 0 Then
                    MsgBox "Imposible seleccionar Item, requerimiento con Stock Comprometido para su atención.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                If Val(.Columns.ColumnByFieldName("SALDOACTUAL").Value & "") <= 0 Then
                    MsgBox "Producto sin saldo pendiente, verifique.", vbInformation + vbOKOnly, App.ProductName

                    .Dataset.Cancel

                    Exit Sub
                End If
                
                If .Dataset.State = dsEdit Then
                    .Dataset.Post
                End If
                
                .Dataset.Edit

                If CBool(.Columns.ColumnByFieldName("PROCESAR").Value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADPC").Value = IIf(Val(.Columns.ColumnByFieldName("SALDOACTUAL").Value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(.Columns.ColumnByFieldName("SALDOACTUAL").Value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))
                                         
                        evaluarCantidadDeCompra Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), True
                    End If
                Else
                    evaluarCantidadDeCompra Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), False
                    
                    .Columns.ColumnByFieldName("CANTIDADPC").Value = 0
                End If

                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgResumen_OnCustomDrawCell(Index As Integer, ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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
        Case "COMPROMISOEA", "COMPROMISOPL"
'            If Val(Text) = 0 Then
'                FontColor = vbWhite
'                Color = vbWhite
'            End If
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
        Case "COMPROMISOEAG", "COMPROMISOPLG", "LIBREEAG", "LIBREPLG", "STOCKEAG", "STOCKPLG"
            If Val(Text) = 0 Then
                FontColor = vbWhite
                Color = vbWhite
            End If
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
        Case "COMPRA", "CANTIDADPC"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(75, 172, 198)
            ElseIf Val(Text) < 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = vbRed
            Else
                Font.Bold = False
                FontColor = vbBlack
                Color = &HFFFFC0
            End If
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
        Case "NOMPRODUCTO", "UM", "CANTIDAD", "ATENDIDOPORPROV2", "COMPRAR"
            If Val(Node.Values(15) & "") = 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(75, 172, 198)
            End If
    End Select
End Sub

Private Sub dbgResumen_OnCustomDrawFooter(Index As Integer, ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "SALDOACTUAL", "COMPRA", "CANTIDADPC"
            Font.Bold = True
            FontColor = vbBlue
            Color = vbWhite
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
    End Select
End Sub

Private Sub dbgResumen_OnDblClick(Index As Integer)
    For d = 0 To 25
        nSaveRecNo = dbgResumen(Index).Dataset.RecNo
    Next
    
    With objAyudaBien
        .inicializarEntidades
        
        .Codigo = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
        
        .obtenerConfigBien
    End With
    
    Select Case dbgResumen(Index).Columns.FocusedColumn.FieldName
        Case "COMPROMISOEA"
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "F"
                .NroPedido = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")

                .Show vbModal
            End With
        Case "COMPROMISOPL"
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "V"
                .NroPedido = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")

                .Show vbModal
            End With
        Case "COMPROMISOEAG" 'Stock Comprometido En Almacen (CEA)
            If Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPROMISOEAG").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("frmUtilStockDetalle") Then
                Unload frmUtilStockDetalle
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0 And Val(dbgResumen(Index).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") > 0, False, True)
                
                If .DeshabilitarRedistribucion Then
                    Exit Sub
                End If
                
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                
                .Show 1
            End With
        Case "COMPROMISOPLG" 'Stock Comprometido Por Llegar (CPL)
            If Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPROMISOPLG").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("frmUtilStockDetalle") Then
                Unload frmUtilStockDetalle
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0 And Val(dbgResumen(Index).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") > 0, False, True)
                
                If .DeshabilitarRedistribucion Then
                    Exit Sub
                End If
                
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                
                .Show 1
            End With
        Case "LIBREEAG" 'Stock Libre En Almacen (LEA)
            If Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("frmUtilStockDetalle") Then
                Unload frmUtilStockDetalle
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0 And Val(dbgResumen(Index).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") > 0, False, True)
                
                If objAyudaBien.StockMin > 0 Then
                    If Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "") <= objAyudaBien.StockMin Then
                        MsgBox "Imposible aplicar re-distribución, se ha alcanzado el Stock Minimo del producto.", vbInformation + vbOKOnly, App.ProductName
                        
                        .DeshabilitarRedistribucion = True
                    End If
                    
                    .CantidadMaximaParaPedido = IIf((Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "") - objAyudaBien.StockMin) < Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), (Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "") - objAyudaBien.StockMin), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))
                Else
                    .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                End If
                
                If .DeshabilitarRedistribucion Then
                    Exit Sub
                End If
                
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                
                .Show 1
            End With
        Case "LIBREPLG" 'Stock Libre Por Llegar (LPL)
            If Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREPLG").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("frmUtilStockDetalle") Then
                Unload frmUtilStockDetalle
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString
                
                .DeshabilitarRedistribucion = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0 And Val(dbgResumen(Index).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") > 0, False, True)
                
                If objAyudaBien.StockMin > 0 Then
                    If Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "") < objAyudaBien.StockMin Then
                        If Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREPLG").Value & "") <= _
                            (objAyudaBien.StockMin - Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "")) Then
                            
                            MsgBox "Imposible aplicar re-distribución, se aguarda el Stock para completar el Stock Minimo en Almacen.", vbInformation + vbOKOnly, App.ProductName
                            
                            .DeshabilitarRedistribucion = True
                        End If
                    End If
                    
                    .CantidadMaximaParaPedido = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREPLG").Value & "") > (objAyudaBien.StockMin - Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "")), _
                                                        IIf((Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREPLG").Value & "") - (objAyudaBien.StockMin - Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & ""))) <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREPLG").Value & "") - (objAyudaBien.StockMin - Val(dbgResumen(Index).Columns.ColumnByFieldName("LIBREEAG").Value & "")), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)), _
                                                        Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))
                Else
                    .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                End If
                
                If .DeshabilitarRedistribucion Then
                    Exit Sub
                End If
                
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                '.CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                
                .Show 1
            End With
        Case "COMPRAR"
            If dbgResumen(0).Visible Then
                If Val(dbgResumen(0).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") > 0 Then
                    If CBool(tlbResumen.Tools("Seleccionar").Enabled) Then
                        tlbResumen_ToolClick tlbResumen.Tools("Seleccionar")
                    ElseIf CBool(tlbResumen.Tools("QuitarSeleccion").Enabled) Then
                        tlbResumen_ToolClick tlbResumen.Tools("QuitarSeleccion")
                    End If
                    
                    dbgResumen(0).Columns.FocusedIndex = 2
                Else
                    MsgBox "El Producto no cuenta con Saldo en Producción, verifique.", vbInformation + vbOKOnly, App.ProductName
                End If
            End If
        Case Else '"CANTIDAD"
            If dbgResumen(0).Visible Then
                tlbResumen_ToolClick tlbResumen.Tools("Siguiente")
            End If
    End Select
    
    Select Case UCase(dbgResumen(Index).Columns.FocusedColumn.FieldName)
        Case "COMPROMISOEAG", "COMPROMISOPLG", "LIBREEAG", "LIBREPLG"
            If frmUtilStockDetalle.RedistribucionEjecutada Then
                Me.MousePointer = vbHourglass
                
                tlbResumen.Tools("Anterior").Enabled = False
                
                dbgResumen(Index).Dataset.Close
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    descargarAtencionRequerimientoSQL
                Else
                    descargarAtencionRequerimiento
                End If
                
                evaluarCantidadDeCompra Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDAD)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'F'")) + _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'V'")) + _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), True
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    With objSqlAyudaVale
                        .CodigoProducto = strCodProducto
                        
                        .verificarStockProducto
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                        SqlCad = SqlCad & "SET "
                        SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                        SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                        SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                        SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                        SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                        SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
    
                        .inicializarEntidadesDetalle
                        .inicializarEntidadesAdicionales
                    End With
                Else
                    Rem SK: HABILITADO TEMPORALMENTE
                    With objAyudaVale
                        .CodigoProducto = strCodProducto
                        
                        .verificarStockProducto
    
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                        SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                        SqlCad = SqlCad & "SET "
                        SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                        SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                        SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                        SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                        SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                        SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                        SqlCad = SqlCad & "WHERE "
                        SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
    
                        .inicializarEntidadesDetalle
                        .inicializarEntidadesAdicionales
                    End With
                End If
                
                Select Case Index
                    Case 0
                        cargarResumenRequerimientoVista1
                    Case 1
                        cargarResumenRequerimientoVista2
                        
                        tlbResumen.Tools("Anterior").Enabled = True
                End Select
                
                Me.MousePointer = vbDefault
            End If
    End Select
    
    If dbgResumen(Index).Dataset.RecordCount >= nSaveRecNo Then
        dbgResumen(Index).Dataset.RecNo = nSaveRecNo
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgResumen_OnEditButtonClick(Index As Integer, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case UCase(Column.FieldName)  'Column.ObjectName
        Case "NROOP" '"ColCambio"
            If Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDAD").Value & "") <> Val(dbgResumen(1).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") Then
                MsgBox "Imposible reemplazar el Producto, ya fue descargado.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDAD").Value & "") = 0 Then
                MsgBox "Imposible modificar la Cantidad, el Item se encuentra anulado.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With objSqlAyudaVale
                    .CodigoProducto = strCodProducto
                    
                    .verificarStockProducto strNroPedido
                    
                    If .CompromisoEAG > 0 Then
                        If MsgBox("El Producto cuenta actualmente con " & .CompromisoEAG & " en Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                            
                            Exit Sub
                        End If
                    End If
                                    
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
            Else
                With objAyudaVale
                    .CodigoProducto = strCodProducto
                    
                    .verificarStockProducto strNroPedido
                    
                    If .CompromisoEAG > 0 Then
                        If MsgBox("El Producto cuenta actualmente con " & .CompromisoEAG & " en Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                            
                            Exit Sub
                        End If
                    End If
                                    
                    .inicializarEntidadesDetalle
                    .inicializarEntidadesAdicionales
                End With
            End If
            
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                objAyudaBien.Codigo = strCodProducto
                
                objAyudaBien.obtenerConfigBien
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True
                '.CadenaCorte = objAyudaBien.Modelo
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = objAyudaBien.Modelo
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    If Trim(dbgResumen(1).Columns.ColumnByFieldName("CODPRODUCTO").Value) = objAyudaBien.Codigo Then
                        MsgBox "Imposible reemplazar el Item por el mismo Codigo de Producto, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        Exit Sub
                    End If
                    
                    If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPRODUCTO", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", objAyudaBien.Codigo, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND IDOP = '" & Trim(dbgResumen(1).Columns.ColumnByFieldName("IDOP").Value & "") & "'") <> vbNullString Then
                        MsgBox "Imposible realizar el cambio; el producto:" & vbNewLine & _
                                objAyudaBien.Descripcion & ", " & vbNewLine & _
                                "Se encuentra consignado en la actual OP. Se sugiere:" & vbNewLine & _
                                "1) Sumar la Cantidad del producto a cambiar, al que ya existe." & vbNewLine & _
                                "2) En caso que el producto a cambiar cuente con Stock Comprometido, proceder a Liberarlo." & vbNewLine & _
                                "3) Anular (Desestimar) el producto a cambiar, llevando su Cantidad Final a cero (0)." & vbNewLine & _
                                "4) Proceder con la Compra del Producto ya existente, que cuenta con la suma de ambos.", vbInformation + vbOKOnly, App.ProductName
                        
                        objAyudaBien.inicializarEntidades
                        
                        Exit Sub
                    End If
                    
                    If Trim(dbgResumen(1).Columns.ColumnByFieldName("UM").Value & "") <> Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")) Then
                        If MsgBox("El Producto seleccionado para el cambio cuenta con diferente Unidad de Medida (U.M.)." & vbNewLine & _
                                    "¿Desea continuar con la acción?", vbInformation + vbYesNo, App.ProductName) = vbNo Then
                            
                            objAyudaBien.inicializarEntidades
                            
                            Exit Sub
                        End If
                    End If
                    
                    If MsgBox("¿Desea efectuar el reemplazo del Producto?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        Exit Sub
                    End If
                    
                    With dbgResumen(1)
                        If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("IDOP").Value), _
                                                            Trim(.Columns.ColumnByFieldName("CODPRODUCTO").Value), _
                                                            objAyudaBien.Codigo, _
                                                            Val(.Columns.ColumnByFieldName("CANTIDAD").Value), _
                                                            Val(.Columns.ColumnByFieldName("CANTIDAD").Value), "PROCESAMIENTO LOGISTICO DE OP - CAMBIO DE PRODUCTO") Then
                            
                            MsgBox "Producto reemplazado en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").Value), vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Edit
                            
                            .Columns.ColumnByFieldName("CODPRODUCTO").Value = objAyudaBien.Codigo
                            .Columns.ColumnByFieldName("NOMPRODUCTO").Value = objAyudaBien.Descripcion
                            .Columns.ColumnByFieldName("UM").Value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                            .Columns.ColumnByFieldName("NOMPRODUCTOUM").Value = objAyudaBien.Descripcion & " ( " & .Columns.ColumnByFieldName("UM").Value & " )"
                            
                            .Dataset.Post
                            
                            dbgResumen(1).Dataset.Close
                            
                            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                descargarAtencionRequerimientoSQL
                            Else
                                descargarAtencionRequerimiento
                            End If
                            
                            evaluarCantidadDeCompra Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDAD)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), _
                                                    Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'F'")) + _
                                                    Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'V'")) + _
                                                    Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), True
                            
                            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                With objSqlAyudaVale
                                    .CodigoProducto = strCodProducto
                                    
                                    .verificarStockProducto
                                    
                                    SqlCad = vbNullString
                                    SqlCad = SqlCad & "UPDATE "
                                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                    SqlCad = SqlCad & "SET "
                                    SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                    SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                    SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                    SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                    SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                    SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                    SqlCad = SqlCad & "WHERE "
                                    SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                                    
                                    cnDBTemp.Execute SqlCad
                                    
                                    .inicializarEntidadesDetalle
                                    .inicializarEntidadesAdicionales
                                    
                                    .CodigoProducto = objAyudaBien.Codigo
                                    
                                    .verificarStockProducto
                                    
                                    SqlCad = vbNullString
                                    SqlCad = SqlCad & "UPDATE "
                                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                    SqlCad = SqlCad & "SET "
                                    SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                    SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                    SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                    SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                    SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                    SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                    SqlCad = SqlCad & "WHERE "
                                    SqlCad = SqlCad & "CODPRODUCTO = '" & objAyudaBien.Codigo & "'"
                                    
                                    cnDBTemp.Execute SqlCad
                                    
                                    .inicializarEntidadesDetalle
                                    .inicializarEntidadesAdicionales
                                End With
                            Else
                                With objAyudaVale
                                    .CodigoProducto = strCodProducto
                                    
                                    .verificarStockProducto
                                    
                                    SqlCad = vbNullString
                                    SqlCad = SqlCad & "UPDATE "
                                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                    SqlCad = SqlCad & "SET "
                                    SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                    SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                    SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                    SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                    SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                    SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                    SqlCad = SqlCad & "WHERE "
                                    SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                                    
                                    cnDBTemp.Execute SqlCad
                                    
                                    .inicializarEntidadesDetalle
                                    .inicializarEntidadesAdicionales
                                    
                                    .CodigoProducto = objAyudaBien.Codigo
                                    
                                    .verificarStockProducto
                                    
                                    SqlCad = vbNullString
                                    SqlCad = SqlCad & "UPDATE "
                                    SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                    SqlCad = SqlCad & "SET "
                                    SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                    SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                    SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                    SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                    SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                    SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                    SqlCad = SqlCad & "WHERE "
                                    SqlCad = SqlCad & "CODPRODUCTO = '" & objAyudaBien.Codigo & "'"
                                    
                                    cnDBTemp.Execute SqlCad
                                    
                                    .inicializarEntidadesDetalle
                                    .inicializarEntidadesAdicionales
                                End With
                            End If
                            
                            cargarResumenRequerimientoVista2
                        End If
                        
                    End With
                End If
            End With
    End Select
End Sub

Private Sub dbgResumen_OnEdited(Index As Integer, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgResumen(1).Columns.FocusedColumn.FieldName
        Case "CANTIDAD"
            With dbgResumen(1)
                If .Dataset.State = dsEdit Then
                    Dim dblCantidadOrigen As Double
                    
                    dblCantidadOrigen = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND IDOP = '" & Trim(dbgResumen(1).Columns.ColumnByFieldName("IDOP").Value & "") & "'"))
                    
                    If dblCantidadOrigen = 0 Then
                        If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0017'") = vbNullString Then
                            MsgBox "Imposible modificar la Cantidad, el Item se encuentra anulado." & vbNewLine & "Se requieren el permiso correspondiente, comuniquese con su administrador de Sistemas.", vbInformation + vbOKOnly, App.ProductName
                            
                            .Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    If dblCantidadOrigen <> Val(dbgResumen(1).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") Then
                        MsgBox "Imposible modificar la Cantidad, el Item se encuentra descargado.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If ModUtilitario.sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        With objSqlAyudaVale
                            .CodigoProducto = strCodProducto
                            
                            .verificarStockProducto strNroPedido
                            
                            If .CompromisoEAG > 0 Then
                                If MsgBox("El Producto cuenta actualmente con " & .CompromisoEAG & " en Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                        "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                                    
                                    dbgResumen(1).Dataset.Cancel
                                    
                                    Exit Sub
                                End If
                            End If
                            
                            .inicializarEntidadesDetalle
                            .inicializarEntidadesAdicionales
                        End With
                    Else
                        With objAyudaVale
                            .CodigoProducto = strCodProducto
                            
                            .verificarStockProducto strNroPedido
                            
                            If .CompromisoEAG > 0 Then
                                If MsgBox("El Producto cuenta actualmente con " & .CompromisoEAG & " en Stock Comprometido Disponible, ¿Desea continuar con el cambio?" & vbNewLine & vbNewLine & _
                                        "RECOMENDACIÓN: Asegurese de liberar el Stock Comprometido del Producto, antes de proceder con el Cambio.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                                    
                                    dbgResumen(1).Dataset.Cancel
                                    
                                    Exit Sub
                                End If
                            End If
                            
                            .inicializarEntidadesDetalle
                            .inicializarEntidadesAdicionales
                        End With
                    End If
                    
                    If MsgBox("¿Desea aplicar el Ajuste de Cantidad?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        .Dataset.Cancel
                        
                        Exit Sub
                    Else
                        .Dataset.Post
                    End If
                    
                    If ModMilano.modificarProductoEnOP(Trim(.Columns.ColumnByFieldName("IDOP").Value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTO").Value), _
                                                        Trim(.Columns.ColumnByFieldName("CODPRODUCTO").Value), _
                                                        dblCantidadOrigen, _
                                                        Val(.Columns.ColumnByFieldName("CANTIDAD").Value), "PROCESAMIENTO LOGISTICO DE OP - AJUSTE DE CANTIDAD DE PRODUCTO") Then
                        
                        MsgBox "Efectuado ajuste de cantidad de Producto en OP: " & Trim(.Columns.ColumnByFieldName("NROOP").Value), vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("SALDOACTUAL").Value = Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "") - (dblCantidadOrigen - Val(.Columns.ColumnByFieldName("SALDOACTUAL").Value & ""))
                        
                        .Dataset.Post
                        
                        dbgResumen(1).Dataset.Close
                        
                        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                            descargarAtencionRequerimientoSQL
                        Else
                            descargarAtencionRequerimiento
                        End If
                        
                        evaluarCantidadDeCompra Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDAD)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), _
                                                Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'F'")) + _
                                                Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENATENCIONREQUERIMIENTOOP", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' AND TIPO = 'V'")) + _
                                                Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM")), True
                        
                        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                            With objSqlAyudaVale
                                .CodigoProducto = strCodProducto
                                
                                .verificarStockProducto
                                
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "UPDATE "
                                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                SqlCad = SqlCad & "SET "
                                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                                
                                cnDBTemp.Execute SqlCad
                                
                                .inicializarEntidadesDetalle
                                .inicializarEntidadesAdicionales
                            End With
                        Else
                            With objAyudaVale
                                .CodigoProducto = strCodProducto
                                
                                .verificarStockProducto
                                
                                SqlCad = vbNullString
                                SqlCad = SqlCad & "UPDATE "
                                SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION "
                                SqlCad = SqlCad & "SET "
                                SqlCad = SqlCad & "COMPROMISOEAG = " & .CompromisoEAG & ", "
                                SqlCad = SqlCad & "COMPROMISOPLG = " & .CompromisoPLG & ", "
                                SqlCad = SqlCad & "LIBREEAG = " & .LibreEAG & ", "
                                SqlCad = SqlCad & "LIBREPLG = " & .LibrePLG & ", "
                                SqlCad = SqlCad & "STOCKEAG = " & .StockEAG & ", "
                                SqlCad = SqlCad & "STOCKPLG = " & .StockPLG & " "
                                SqlCad = SqlCad & "WHERE "
                                SqlCad = SqlCad & "CODPRODUCTO = '" & strCodProducto & "'"
                                
                                cnDBTemp.Execute SqlCad
                                
                                .inicializarEntidadesDetalle
                                .inicializarEntidadesAdicionales
                            End With
                        End If
                    Else
                        .Dataset.Edit
                        
                        .Columns.ColumnByFieldName("CANTIDAD").Value = Val(.Columns.ColumnByFieldName("SALDOACTUAL").Value & "")
                        
                        .Dataset.Post
                    End If
                    
                    dblCantidadOrigen = 0
                    
                    cargarResumenRequerimientoVista2
                End If
            End With
        
        Case "CANTIDADPC"
            With dbgResumen(1)
                .Dataset.Edit
                
                If CBool(.Columns.ColumnByFieldName("PROCESAR").Value) Then
                    Exit Sub
                End If
                
                If Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") > 0 And Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) <= 0 Then
                    MsgBox "Imposible seleccionar Item, requerimiento con Stock Comprometido para su atención.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                If Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") > Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) Then
                    MsgBox "La cantidad no puede exceder al Saldo por Comprar, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    .Dataset.Cancel
                    
                    Exit Sub
                End If
                
                If Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").Value & "") > Val(dbgResumen(1).Columns.ColumnByFieldName("SALDOACTUAL").Value & "") Then
                    MsgBox "La cantidad no puede exceder al Saldo por Atender, verifique.", vbInformation + vbOKOnly, App.ProductName

                    .Dataset.Cancel

                    Exit Sub
                End If
                
                .Columns.ColumnByFieldName("PROCESAR").Value = IIf(Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") <= 0, False, True)
                
                evaluarCantidadDeCompra Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), _
                                        Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), _
                                        CBool(.Columns.ColumnByFieldName("PROCESAR").Value)
                
                .Dataset.Post
            End With
    End Select
End Sub

Private Sub dbgResumen_OnKeyDown(Index As Integer, KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            If dbgResumen(Index).Dataset.State = dsEdit Or dbgResumen(Index).Dataset.State = dsInsert Then
                dbgResumen(Index).Dataset.Post
            End If
            
            dbgResumen_OnDblClick Index
        Case vbKeyEscape
            If dbgResumen(Index).Dataset.State = dsEdit Or dbgResumen(Index).Dataset.State = dsInsert Then
                dbgResumen(Index).Dataset.Post
            End If
            
            Select Case Index
                Case 1
                    tlbResumen_ToolClick tlbResumen.Tools("Anterior")
            End Select
    End Select
End Sub

Private Sub cargarResumenProduccion()
    On Error GoTo errCargarResumenProduccion
    
    If bolResumenCargado = False Then
        Me.MousePointer = vbHourglass
        
        tlbResumen.Tools("Consultar").Enabled = bolResumenCargado
        tlbResumen.Tools("ID_Salir").Enabled = bolResumenCargado
        
        fraDatos.Enabled = bolResumenCargado
        FraBusqueda.Enabled = bolResumenCargado
        fraOpciones.Enabled = bolResumenCargado
        'tlbResumen.Enabled = bolResumenCargado
        tlbResumen.Tools("Anterior").Enabled = bolResumenCargado
        tlbResumen.Tools("Siguiente").Enabled = bolResumenCargado
        tlbResumen.Tools("Seleccionar").Enabled = bolResumenCargado
        tlbResumen.Tools("QuitarSeleccion").Enabled = bolResumenCargado
        
        fraProveedor.Enabled = bolResumenCargado

        dbgResumen(0).Enabled = bolResumenCargado

        bolResumenCargado = True
        
        tlbResumen.Tools("ID_Inicio").Edit.Text = "Inicio: " & Format(Now, "hh:mm:ss AM/PM")

        If strNroPedido <> vbNullString Then
            'abrirCnDBMilano
            
            Me.Caption = "Productos del Requerimiento N° " & strNroPedido & " - " & ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", strNroPedido, "T")
        Else
            Me.Caption = "Producto(s) Pendiente(s) de Atención en Requerimiento(s)"
        End If
        
        If Not cargarResumenRequerimiento Then
            tlbResumen.Tools("Consultar").Enabled = bolResumenCargado
            tlbResumen.Tools("ID_Salir").Enabled = bolResumenCargado
            
            fraDatos.Enabled = bolResumenCargado
            
            bolObviarConsulta = bolResumenCargado
            
            tlbResumen.Tools.ITEM("Consultar").State = ssUnchecked
            
            tlbResumen.Tools("Consultar").ChangeAll ssChangeAllName, "Consul&tar"
            
            bolObviarConsulta = Not bolResumenCargado
            
            bolResumenCargado = False
            
            Me.MousePointer = vbDefault
            
            Exit Sub
        Else
            'Verificar los Codigos de Productos extraidos del Resumen de Requerimientos
            verificarExistenciaCodigoInsumo
        End If
        
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            descargarAtencionRequerimientoSQL
        Else
            descargarAtencionRequerimiento
        End If
        
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            cargarStockProductoSQL
        Else
            Rem SK: HABILITADO TEMPORALMENTE
            cargarStockDeProducto
        End If
        
        cargarResumenRequerimientoVista1
        
        tlbResumen.Tools("Consultar").Enabled = bolResumenCargado
        tlbResumen.Tools("ID_Salir").Enabled = bolResumenCargado
        
        fraDatos.Enabled = Not bolResumenCargado
        FraBusqueda.Enabled = bolResumenCargado
        fraOpciones.Enabled = bolResumenCargado
        'tlbResumen.Enabled = bolResumenCargado
        tlbResumen.Tools("Anterior").Enabled = bolResumenCargado
        tlbResumen.Tools("Siguiente").Enabled = bolResumenCargado
        tlbResumen.Tools("Seleccionar").Enabled = bolResumenCargado
        tlbResumen.Tools("QuitarSeleccion").Enabled = bolResumenCargado
        
        fraProveedor.Enabled = IIf(VerificaAutorizaciones("OCN", wusuario) <> "''", bolResumenCargado, False)
        
        dbgResumen(0).Enabled = bolResumenCargado

        tlbResumen.Tools("ID_Fin").Edit.Text = "Fin: " & Format(Now, "hh:mm:ss AM/PM")

        dbgResumen_OnClick 0
        
        tlbResumen.Tools("Siguiente").Enabled = True
        tlbResumen.Tools("Anterior").Enabled = False
        
        Me.MousePointer = vbDefault
        
        fraProveedor.Enabled = IIf(VerificaAutorizaciones("OCN", wusuario) <> "''", True, False)
        
        dbgResumen(0).SetFocus

        'Activar Control de Apertura de Formulario
        '(Para evitar abrir mas de una vez, el mismo formulario en diferentes Instancias del Programa)
        strFichero = wrutatemp & strNombreFicheroConfigCPusuario

        If bolResumenCargado Then
            ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", "1"
        End If

        frmUtilResumenProduccion.SetFocus
    End If

    Exit Sub
errCargarResumenProduccion:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    End Select

    Err.Clear
End Sub

Private Sub inicializarControles()
    txtCodProducto.Text = vbNullString: txtProducto.Text = vbNullString
    txtNroPedido.Text = vbNullString: lblNroPedido.Caption = vbNullString
    
    txtbusqueda.Text = vbNullString
    
    chkProductoSeleccionado.Enabled = False
    chkProductoProveedor.Enabled = True
    
    tlbResumen.Tools("ID_Cantidad").Edit.Text = "0.00"
    tlbResumen.Tools("ID_EstadoAtte").Edit.Text = "--"
    
    'tlbResumen.Tools("Seleccionar").Enabled = False
    'tlbResumen.Tools("QuitarSeleccion").Enabled = False
    
    txtCodProveedor.Text = vbNullString
        lblProveedor.Caption = vbNullString
        
    txtCodProductoAdd.Text = vbNullString
        txtProductoAdd.Text = vbNullString: txtProductoAdd.Locked = True: txtProductoAdd.BackColor = DF
    
    cmbColocarEnOrden.ListIndex = -1
        
    fraDatos.Enabled = True
    FraBusqueda.Enabled = False
    fraOpciones.Enabled = False
    tlbResumen.Tools("Anterior").Enabled = False
    tlbResumen.Tools("Siguiente").Enabled = False
    tlbResumen.Tools("Seleccionar").Enabled = False
    tlbResumen.Tools("QuitarSeleccion").Enabled = False
    'tlbResumen.Tools("ID_Salir").Enabled = False
    fraProveedor.Enabled = False
    
    fraProductoAdd.Visible = False
    fraProductoAdd.Enabled = False
    
    fraProceso.Visible = False
    
    dbgResumen(0).Dataset.Close
    
    dbgResumen(0).Enabled = False
    
    bolResumenCargado = False
End Sub

Private Sub Form_Load()
    bolResumenCargado = False
    
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    abrirCnTemporal
    
    bolObviarConsulta = False
    
    inicializarControles
    
    'Para Control de Anterior/Siguiente
'    tmrTemporizador.Enabled = False
'    tmrTemporizador.Interval = 0
'    tlbResumen.Tools("Anterior").Enabled = False
'    tlbResumen.Tools("Siguiente").Enabled = True
'
'    intIndiceGrilla = 0
'    intIndiceVisible = 0
'    intIndiceOculto = 0
    
'    If strNroPedido = vbNullString Then
'        MsgBox "No se especifico el No. de Requerimiento a consultar.", vbInformation + vbOKOnly, App.ProductName
'
'        bolResumenCargado = True
'    End If
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilResumenProduccion
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bolResumenCargado Then
    'If strNroPedido <> vbNullString Then
        If MsgBox("¿Desea salir de la Consulta actual?" & vbNewLine & _
                    "RECUERDA: La actual consulta se provee de datos Externos, por lo cual tiende a tardar algunos minutos en cargar.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            
            Cancel = 1
            
            If dbgResumen(0).Visible Then
                dbgResumen(0).Dataset.ADODataset.Requery
                
                dbgResumen(0).m.ResetFullRefresh
            End If
            
            Exit Sub
        End If
    End If
    
    If dbgResumen(1).Dataset.State = dsEdit Or dbgResumen(1).Dataset.State = dsInsert Then
        dbgResumen(1).Dataset.Post
    End If
    
    dbgResumen(1).Dataset.Close
    
    If dbgResumen(0).Dataset.State = dsEdit Or dbgResumen(0).Dataset.State = dsInsert Then
        dbgResumen(0).Dataset.Post
    End If
    
    dbgResumen(0).Dataset.Close
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", "0"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgResumen(0).Move 0, FraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + 300)
    dbgResumen(1).Move 0, FraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + 300)
    
'    dbgResumen(1).Move dbgResumen(0).Width, dbgResumen(0).top, dbgResumen(0).Width, dbgResumen(0).Height
'    dbgResumen(2).Move dbgResumen(0).Width, dbgResumen(0).top, dbgResumen(0).Width, dbgResumen(0).Height
'
'    dblFactorAncho = dbgResumen(intIndiceGrilla).Width / 10
End Sub

Private Sub tlbResumen_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Consultar"
            If Not bolObviarConsulta Then
                If Not CBool(tlbResumen.Tools.ITEM("Consultar").State) Then
                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", vbNullString, vbNullString, vbNullString, "TRIM(CODPRODUCTO & '') <> ''")) > 0 Then
                        If MsgBox("Se cuenta con Items marcados para compra, " & vbNewLine & _
                                    "¿Desea salir sin generar Orden de Compra?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                            
                            If dbgResumen(0).Visible Then
                                dbgResumen(0).Dataset.ADODataset.Requery
                                
                                dbgResumen(0).m.ResetFullRefresh
                            End If
                            
                            bolObviarConsulta = True
                            
                            tlbResumen.Tools("Consultar").State = ssChecked
                            
                            bolObviarConsulta = False
                            
                            Exit Sub
                        End If
                    End If
                    
                    If MsgBox("¿Desea salir de la consulta actual?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        bolObviarConsulta = True
                        
                        tlbResumen.Tools("Consultar").State = ssChecked
                        
                        bolObviarConsulta = False
                        
                        Exit Sub
                    End If
                    
                    tlbResumen.Tools("Consultar").ChangeAll ssChangeAllName, "Consul&tar"
                    
                    strCodProducto = vbNullString
                    strNomProducto = vbNullString
                    strNroPedido = vbNullString
                    
                    inicializarControles
                Else
                    If Trim(txtNroPedido.Text) = vbNullString Then
                        MsgBox "Ingrese el No. de Pedido.", vbInformation + vbOKOnly, App.ProductName
            
                        txtNroPedido.SetFocus
                        
                        Exit Sub
                    End If
                    
                    If lblNroPedido.Caption = vbNullString Then
                        'abrirCnDBMilano
                                
                        lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
                        lblNroPedido.ToolTipText = lblNroPedido.Caption
                        
                        If Trim(txtNroPedido.Text) <> vbNullString And lblNroPedido.Caption = vbNullString Then
                            MsgBox "No. de Pedido no encontrado o inválido.", vbInformation + vbOKOnly, App.ProductName
                
                            txtNroPedido.SetFocus
                            
                            Exit Sub
                        End If
                    End If
                    
                    If MsgBox("¿Desea ejecutar la consulta?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        bolObviarConsulta = True
                        
                        tlbResumen.Tools("Consultar").State = ssUnchecked
                        
                        bolObviarConsulta = False
                        
                        Exit Sub
                    End If
                    
                    tlbResumen.Tools("Consultar").ChangeAll ssChangeAllName, "Salir Consul&ta"
                    
                    strCodProducto = Trim(txtCodProducto.Text)
                    strNomProducto = Trim(txtProducto.Text)
                    strNroPedido = Trim(txtNroPedido.Text)
                    
                    cargarResumenProduccion
                End If
            End If
        Case "ID_Filtrar"
            dbgResumen(0).Filter.FilterActive = CBool(Tool.State)
        Case "Seleccionar"
            estadoSeleccion True
            
            If dbgResumen(0).Visible Then
                dbgResumen(0).Dataset.ADODataset.Requery
                
                dbgResumen(0).m.ResetFullRefresh
            End If
        Case "QuitarSeleccion"
            estadoSeleccion False
            
            If dbgResumen(0).Visible Then
                dbgResumen(0).Dataset.ADODataset.Requery
                
                dbgResumen(0).m.ResetFullRefresh
            End If
        Case "Anterior"
            Screen.MousePointer = vbHourglass
            
            tlbResumen.Tools("Anterior").Enabled = False
            tlbResumen.Tools("Siguiente").Enabled = True
            
            If dbgResumen(1).Dataset.State = dsEdit Then
                dbgResumen(1).Dataset.Post
            Else
                If dbgResumen(1).Dataset.RecordCount > 0 Then
                    dbgResumen(1).Dataset.Edit
                    
                    dbgResumen(1).Dataset.Post
                End If
            End If
            
            dbgResumen(1).Dataset.Close
            
            dbgResumen(1).Visible = False
            dbgResumen(0).Visible = True
            
            strCodProducto = vbNullString
            
            txtbusqueda.Text = vbNullString
            
            chkProductoSeleccionado.Enabled = False
            chkProductoProveedor.Enabled = True
            
            tlbResumen.Tools("ID_Cantidad").Edit.Text = "0.00"
            tlbResumen.Tools("ID_EstadoAtte").Edit.Text = "--"
            
            tlbResumen.Tools("Seleccionar").Enabled = False
            tlbResumen.Tools("QuitarSeleccion").Enabled = False
            
            tlbResumen.Tools("Consultar").Enabled = True
            tlbResumen.Tools("ID_Salir").Enabled = True
            
            If VerificaAutorizaciones("OCN", wusuario) <> "''" Then
                fraProveedor.Enabled = True
            End If
            
            fraProductoAdd.Visible = False
            fraProductoAdd.Enabled = False
            
            cargarResumenRequerimientoVista1
            
            If dbgResumen(0).Dataset.RecordCount >= nSaveRecNo Then
                dbgResumen(0).Dataset.RecNo = nSaveRecNo
            End If
            
            dbgResumen_OnClick 0
            
            Screen.MousePointer = vbDefault
        Case "Siguiente"
            Screen.MousePointer = vbHourglass
            
            tlbResumen.Tools("Siguiente").Enabled = False
            tlbResumen.Tools("Anterior").Enabled = True
            
            If VerificaAutorizaciones("OCN", wusuario) <> "''" Then
                fraProveedor.Enabled = False
            End If
            
            fraProductoAdd.Visible = True
            fraProductoAdd.Enabled = True
            
            dbgResumen(1).Bands(0).Caption = Trim(dbgResumen(0).Columns.ColumnByFieldName("NOMPRODUCTO").Value & "") & " ( " & Trim(dbgResumen(0).Columns.ColumnByFieldName("UM").Value & "") & " )"
            
            strCodProducto = Trim(dbgResumen(0).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
            
            dbgResumen(0).Dataset.Close
            
            dbgResumen(0).Visible = False
            dbgResumen(1).Visible = True
            
            chkProductoSeleccionado.Enabled = True
            chkProductoProveedor.Enabled = False
                    
            tlbResumen.Tools("Seleccionar").Enabled = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0, True, False)
            tlbResumen.Tools("QuitarSeleccion").Enabled = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0, True, False)
            
            tlbResumen.Tools("Consultar").Enabled = False
            tlbResumen.Tools("ID_Salir").Visible = False
            tlbResumen.Tools("ID_Salir").Enabled = False
            
            cargarResumenRequerimientoVista2
            
            Screen.MousePointer = vbDefault
        Case "ID_Salir"
            If dbgResumen(0).Dataset.State = dsEdit Then
                dbgResumen(0).Dataset.Post
            Else
                If dbgResumen(0).Dataset.RecordCount > 0 Then
                    dbgResumen(0).Dataset.Edit
                    
                    dbgResumen(0).Dataset.Post
                End If
            End If
            
            If dbgResumen(1).Dataset.State = dsEdit Then
                dbgResumen(1).Dataset.Post
            Else
                If dbgResumen(1).Dataset.RecordCount > 0 Then
                    dbgResumen(1).Dataset.Edit
                    
                    dbgResumen(1).Dataset.Post
                End If
            End If
            
            abrirCnTemporal
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENREQUERIMIENTOPRODUCCION", vbNullString, vbNullString, vbNullString, "TRIM(CODPRODUCTO & '') <> ''")) > 0 Then
                If MsgBox("Se cuenta con Items marcados para compra, " & vbNewLine & _
                            "¿Desea salir sin generar Orden de Compra?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                    
                    If dbgResumen(0).Visible Then
                        dbgResumen(0).Dataset.ADODataset.Requery
                        
                        dbgResumen(0).m.ResetFullRefresh
                    End If
                    
                    Exit Sub
                End If
            End If
            
            Unload Me
    End Select
End Sub

'Private Sub tmrTemporizador_Timer()
'    On Error GoTo errTmrTemporizador
'
'    If tmrTemporizador.Interval = 10 Then
'        tmrTemporizador.Enabled = False
'
'        Select Case intIndiceOculto
'            Case Is = 0
'                tlbResumen.Tools("Anterior").Enabled = False
'                tlbResumen.Tools("Siguiente").Enabled = True
'            Case Is = (dbgResumen.Count - 1)
'                tlbResumen.Tools("Anterior").Enabled = True
'                tlbResumen.Tools("Siguiente").Enabled = False
'            Case Else
'                tlbResumen.Tools("Anterior").Enabled = True
'                tlbResumen.Tools("Siguiente").Enabled = True
'        End Select
'
'        dbgResumen(intIndiceOculto).SetFocus
'    Else
'        tmrTemporizador.Interval = tmrTemporizador.Interval + 1
'
'        dbgResumen(intIndiceVisible).left = dbgResumen(intIndiceVisible).left + (dblFactorAncho * IIf(bolRetroceso, 1, -1))
'        dbgResumen(intIndiceOculto).left = dbgResumen(intIndiceOculto).left + (dblFactorAncho * IIf(bolRetroceso, 1, -1))
'    End If
'
'    Exit Sub
'errTmrTemporizador:
'    Select Case Err.Number
'        Case 5
'            Resume Next
'        Case Else
'            MsgBox "No.: " & Err.Number & vbNewLine & _
'                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
'    End Select
'
'    Err.Clear
'End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtbusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CBool(dbgResumen(0).Visible) Then
                cargarResumenRequerimientoVista1
            ElseIf CBool(dbgResumen(1).Visible) Then
                cargarResumenRequerimientoVista2
            End If
        Case vbKeyEscape
            If CBool(dbgResumen(1).Visible) Then
                tlbResumen_ToolClick tlbResumen.Tools("Anterior")
            End If
    End Select
End Sub

Private Sub txtCodProducto_DblClick()
    txtCodProducto_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProducto_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtCodProducto
End Sub

Private Sub txtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True 'False
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    txtCodProducto.Text = objAyudaBien.Codigo
                    txtProducto.Text = objAyudaBien.Descripcion
                    txtProducto.ToolTipText = txtProducto.Text
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.ToolTipText = txtProducto.Text
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            txtProducto.Text = "Todos los Productos (*)"
            txtProducto.ToolTipText = vbNullString
        End If
    Else
        txtProducto.Text = "Todos los Productos (*)"
        txtProducto.ToolTipText = vbNullString
    End If
End Sub

Private Sub txtCodProductoAdd_DblClick()
    txtCodProductoAdd_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProductoAdd_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtCodProductoAdd
End Sub

Private Sub txtCodProductoAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtCodProductoAdd.Locked Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                '.InsumoOP = True
                
                .Ayuda = True
                .InsumoOP = True
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    txtCodProductoAdd.Text = objAyudaBien.Codigo
                    txtProductoAdd.Text = objAyudaBien.Descripcion
                    txtProductoAdd.ToolTipText = txtProductoAdd.Text
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProductoAdd_LostFocus()
    If Trim(txtCodProductoAdd.Text) <> vbNullString Then
        txtCodProductoAdd.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProductoAdd.Text), "T")
        txtProductoAdd.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProductoAdd.Text), "T")
        txtProductoAdd.ToolTipText = txtProductoAdd.Text
        
        If Trim(txtCodProductoAdd.Text) = vbNullString Then
            txtProductoAdd.Text = vbNullString
            txtProductoAdd.ToolTipText = vbNullString
        End If
    Else
        txtProductoAdd.Text = vbNullString
        txtProductoAdd.ToolTipText = vbNullString
    End If
End Sub

Private Sub txtCodProveedor_DblClick()
    txtCodProveedor_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errTxtCodProveedor
    
    If Not txtCodProveedor.Enabled Then
        Exit Sub
    End If
    
    Select Case KeyCode
        Case vbKeyF2
            Me.MousePointer = vbHourglass
            
            wcodcliprov = vbNullString
            
            With Ayuda_Proveedores
                .Show 1
            End With
            
            If wcodcliprov <> vbNullString Then
                txtCodProveedor.Text = wcodcliprov
                lblProveedor.Caption = wnomcliprov
                
                strCodProveedor = Trim(txtCodProveedor.Text)
                
                listarOrdenesEnCombo
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    cargarProductoAtendidoPorProveedorSQL
                Else
                    cargarProductoAtendidoPorProveedor
                End If
                
                cargarResumenRequerimientoVista1
                
                ModUtilitario.pulsarTecla vbKeyTab
            Else
            '    MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
            '
            '    ModUtilitario.seleccionarTextoCaja txtCodProveedor
            '
            '    Exit Sub
                txtCodProveedor.SetFocus
            End If
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            If Trim(txtCodProveedor.Text) <> vbNullString Then
                lblProveedor.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtCodProveedor.Text), "T")
                
                If Trim(lblProveedor.Caption) = vbNullString Then
                    MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    ModUtilitario.seleccionarTextoCaja txtCodProveedor
                    
                    Exit Sub
                Else
                    strCodProveedor = Trim(txtCodProveedor.Text)
                
                    listarOrdenesEnCombo
                    
                    cargarProductoAtendidoPorProveedor
                    
                    cargarResumenRequerimientoVista1
                End If
            End If
            
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
    
    Exit Sub
errTxtCodProveedor:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    End Select
    
    Err.Clear
End Sub

Private Sub txtNroPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'abrirCnDBMilano
            
            lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
            lblNroPedido.ToolTipText = lblNroPedido.Caption
            
            If lblNroPedido.Caption = vbNullString Then
                MsgBox "No. de Pedido no encontrado o inválido.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroPedido.SetFocus
            Else
                ModUtilitario.pulsarTecla vbKeyTab
            End If
    End Select
End Sub

Private Sub txtNroPedido_LostFocus()
    If lblNroPedido.Caption = vbNullString Then
        'abrirCnDBMilano
                
        lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
        lblNroPedido.ToolTipText = lblNroPedido.Caption
        
        If lblNroPedido.Caption = vbNullString Then
            txtNroPedido.Text = vbNullString
        End If
    End If
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub
