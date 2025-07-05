VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilReposicionCompromiso 
   Caption         =   "Reposición de Compromisos Afectados"
   ClientHeight    =   8475
   ClientLeft      =   135
   ClientTop       =   1770
   ClientWidth     =   15870
   Icon            =   "frmUtilReposicionCompromiso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   15870
   WindowState     =   2  'Maximized
   Begin VB.Frame fraRango 
      Caption         =   " Rango "
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
      Left            =   6960
      TabIndex        =   18
      Top             =   120
      Width           =   2295
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   129302529
         CurrentDate     =   41939
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   129302529
         CurrentDate     =   41939
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   615
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
      Left            =   9360
      TabIndex        =   12
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Orden"
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   550
         Width           =   1215
      End
      Begin VB.ComboBox cmbColocarEnOrden 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   550
         Width           =   3975
      End
      Begin VB.TextBox txtCodProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Colocar en"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblProveedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Timer tmrTemporizador 
      Left            =   0
      Top             =   7800
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   975
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   2295
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
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
      ToolsCount      =   20
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilReposicionCompromiso.frx":058A
      ToolBars        =   "frmUtilReposicionCompromiso.frx":D199
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   6705
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "frmUtilReposicionCompromiso.frx":D3E4
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgResumen 
      Height          =   6825
      Index           =   1
      Left            =   240
      OleObjectBlob   =   "frmUtilReposicionCompromiso.frx":E847
      TabIndex        =   11
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
            Picture         =   "frmUtilReposicionCompromiso.frx":FCAA
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   6255
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   120
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   212
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
      Height          =   975
      Left            =   6960
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
      Begin VB.CheckBox chkProductoProveedor 
         Caption         =   "Mostrar productos de proveedor."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkProductoSeleccionado 
         Caption         =   "Mostrar items a reponer."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
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

Private strFichero              As String

Private bolResumenCargado       As Boolean

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

Private Sub inicializarControles()
    txtbusqueda.Text = vbNullString
    
    dtpDesde.MaxDate = Date
    dtpDesde.MinDate = CDate(ModUtilitario.sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l"))
    dtpDesde.Value = Date 'DateSerial(Year(Date), Month(Date) + 0, 1)
    
    dtpHasta.MaxDate = Date
    dtpHasta.MinDate = CDate(ModUtilitario.sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l"))
    dtpHasta.Value = Date
    
    chkProductoSeleccionado.Enabled = False
    chkProductoProveedor.Enabled = False
    
    tlbResumen.Tools("ID_Inicio").Edit.Text = "00:00:00"
    tlbResumen.Tools("ID_Fin").Edit.Text = "00:00:00"
    
    tlbResumen.Tools("ID_Cantidad").Edit.Text = "0.00"
    
    tlbResumen.Tools("Seleccionar").Enabled = False
    tlbResumen.Tools("QuitarSeleccion").Enabled = False
    tlbResumen.Tools("Descartar").Enabled = False
    
    fraProveedor.Enabled = IIf(VerificaAutorizaciones("OCN", wusuario) <> "''", True, False)
    
    txtCodProveedor.Text = vbNullString
        lblProveedor.Caption = vbNullString
    
    cmbColocarEnOrden.ListIndex = -1
End Sub

Public Sub descargarResumenCompromisoAfectado()
    On Error GoTo errDescargarResumenCompromisoAfectado
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "DELETE FROM TMPUTILRESUMENCOMPROMISOAFECTADO"
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPUTILRESUMENCOMPROMISOAFECTADO("
    SqlCad = SqlCad & "LLAVE, NROPEDIDO, CLIENTE, FEMISION, FENTREGA, "
    SqlCad = SqlCad & "VENDEDOR, CODPRODUCTO, NOMPRODUCTO, UM, NOMPRODUCTOUM, "
    SqlCad = SqlCad & "OBSERVACION, USUARIO, FECHA, CANTIDAD, DESCARTAR, REPOSICION"
    SqlCad = SqlCad & ") "
    
    SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
    
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "RES.IDREPOSICION, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "PED.CS_NOMREF, "
    SqlCad = SqlCad & "PED.CS_FECHA, "
    SqlCad = SqlCad & "PED.CS_FENTREGA, "
    SqlCad = SqlCad & "PED.CS_CODSOLICITANTE, "
    SqlCad = SqlCad & "RES.IDINSUMO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO, "
    SqlCad = SqlCad & "MED.F7SIGMED, "
    SqlCad = SqlCad & "(PROD.F5NOMPRO + ' ( ' + MED.F7SIGMED + ' )') AS NOMPRODUCTOUM, "
    SqlCad = SqlCad & "RES.OBSERVACION, "
    SqlCad = SqlCad & "RES.USUDESCOMPROMISO, "
    SqlCad = SqlCad & "RES.FECDESCOMPROMISO, "
    SqlCad = SqlCad & "RES.CANTIDAD AS CANTIDAD, "
    SqlCad = SqlCad & "RES.DESCARTARREPOSICION, "
    SqlCad = SqlCad & "RES.REPOSICIONENCURSO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "((SF1REPOSICIONCOMPROMISO AS RES "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = RES.IDINSUMO) "
    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
    SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS PED ON PED.COD_SOLICITUD = RES.NROPEDIDO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "(CVDATE(RES.FECDESCOMPROMISO) BETWEEN CVDATE('" & dtpDesde.Value & "') AND CVDATE('" & dtpHasta.Value & "')) AND "
    SqlCad = SqlCad & "RES.DESCARTARREPOSICION = FALSE AND "
    SqlCad = SqlCad & "RES.REPOSICIONENCURSO = FALSE"
    
    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (1/5)..."
    
    cnn_dbbancos.Execute SqlCad
    
    SqlCad = vbNullString

    Exit Sub
errDescargarResumenCompromisoAfectado:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName
    
    Err.Clear
End Sub

Private Sub descargarSaldoEnOPsDeCompromisoAfectado()
    On Error GoTo errDescargarSaldoEnOPsDeCompromisoAfectado
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen(0).Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "

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
    
    DoEvents

    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (2/5)..."
    
    abrirCnTemporal

    If rstTemporal.State = 1 Then rstTemporal.Close

    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly

    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst

        DoEvents

        fraProceso.Caption = "Contabilizando registros consultados (2/5)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0

        DoEvents

        fraProceso.Caption = "Descargando Saldo en OP's (2/5)..."
        
        Do While Not rstTemporal.EOF
            
            dblCantidad = ModMilano.devolverSaldoRequerimientoProduccion(Trim(rstTemporal!NroPedido & ""), Trim(rstTemporal!CodProducto & ""))
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "SALDO = " & IIf(dblCantidad < 0, 0, dblCantidad) & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "NROPEDIDO = '" & Trim(rstTemporal!NroPedido & "") & "' AND "
            SqlCad = SqlCad & "CODPRODUCTO = '" & Trim(rstTemporal!CodProducto & "") & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Saldo en OP's (2/5)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If

    If rstTemporal.State = 1 Then rstTemporal.Close

    fraProceso.Visible = False

    Exit Sub
errDescargarSaldoEnOPsDeCompromisoAfectado:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName

    Err.Clear
End Sub

Private Sub descargarAtencionCompromisoAfectado()
    On Error GoTo errDescargarAtencionCompromisoAfectado
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    dbgResumen(0).Dataset.Close

    abrirCnTemporal

    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENATENCIONCOMPROMISOAFECTADO"
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "NROPEDIDO, "
    SqlCad = SqlCad & "CODPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(CODPRODUCTO & '') <> '' "

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
    
    DoEvents

    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (3/5)..."
    
    abrirCnTemporal

    If rstTemporal.State = 1 Then rstTemporal.Close

    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst

        DoEvents

        fraProceso.Caption = "Contabilizando registros consultados (3/5)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0

        DoEvents

        fraProceso.Caption = "Descargando Atencion de Requerimiento (3/5)..."

        Do While Not rstTemporal.EOF
            With objAyudaSolicitud
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .Codigo = Trim(rstTemporal!NroPedido & "")
                .CodProducto = Trim(rstTemporal!CodProducto & "")
                
                .descargarResumenAtencionPorProductoResCompromisoAfectado "F"
        
                'abrirCnTemporal
                
                .descargarResumenAtencionPorProductoResCompromisoAfectado "V"
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Atencion de Requerimiento (3/5)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1
            
            rstTemporal.MoveNext
        Loop
    End If

    If rstTemporal.State = 1 Then rstTemporal.Close

    fraProceso.Visible = False

    Exit Sub
errDescargarAtencionCompromisoAfectado:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName

    Err.Clear
End Sub

Private Sub cargarStockDeProducto()
    On Error GoTo errCargarStockProducto

    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double

    dbgResumen(0).Dataset.Close

    abrirCnTemporal

    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO"
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CODPRODUCTO, "
    SqlCad = SqlCad & "NOMPRODUCTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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
    
    DoEvents

    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (4/5)..."
    
    abrirCnTemporal

    If rstTemporal.State = 1 Then rstTemporal.Close

    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly

    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst

        DoEvents

        fraProceso.Caption = "Contabilizando registros consultados (4/5)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0

        DoEvents

        fraProceso.Caption = "Descargando Stock de Productos (4/5)..."

        Do While Not rstTemporal.EOF
            With objAyudaVale
                .CodigoProducto = Trim(rstTemporal!CodProducto & "")

                .descargarStockProductoCompAfectado "CEA"

                .descargarStockProductoCompAfectado "CPL"

                .descargarStockProductoCompAfectado "LEA"

                .descargarStockProductoCompAfectado "LPL"
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Stock de Productos (4/5)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1

            rstTemporal.MoveNext
        Loop
    End If

    If rstTemporal.State = 1 Then rstTemporal.Close

    rstTemporal.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly

    DoEvents

    fraProceso.Visible = True
    fraProceso.Caption = "Ejecutando consulta (5/5)..."

    If Not rstTemporal.EOF Then
        rstTemporal.MoveFirst

        DoEvents

        fraProceso.Caption = "Contabilizando registros consultados (5/5)..."
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
        pgbProceso.Value = 0

        DoEvents

        fraProceso.Caption = "Actualizando Stock de Productos (5/5)..."

        Do While Not rstTemporal.EOF
            With objAyudaVale
                .inicializarEntidadesDetalle
                .inicializarEntidadesAdicionales

                .CodigoProducto = Trim(rstTemporal!CodProducto & "")

                .CompromisoEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CEA'"))

                .CompromisoPLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'CPL'"))

                .LibreEAG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LEA'"))

                .LibrePLG = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENSTOCKCOMPROMISOAFECTADO", "CODPRODUCTO", .CodigoProducto, "T", "AND TIPO = 'LPL'"))

                .StockEAG = .CompromisoEAG + .LibreEAG

                .StockPLG = .CompromisoPLG + .LibrePLG

                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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
            fraProceso.Caption = "Actualizando Stock de Productos (5/5)... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) 'pgbProceso.value + 1

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
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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
        SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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
                    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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

'Listar Vista 1 de Compromisos Afectados en Grilla (QuamtumGrid)
Private Sub cargarResumenCompromisoAfectadoVista1()
    On Error GoTo errCargarResumenCompromisoAfectadoVista1
    
    SqlCad = vbNullString
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "(RESU.NROPEDIDO & RESU.CODPRODUCTO) AS LLAVE, "
    SqlCad = SqlCad & "('[ No. Requerimiento: ' & RESU.NROPEDIDO & ' ]" & Space(20) & "[ Cliente: ' & RESU.CLIENTE & ' ]" & Space(20) & "[ Fec. Emisión: ' & RESU.FEMISION & ' ]" & Space(20) & "[ Fec. Entrega: ' & RESU.FENTREGA & ' ]') AS INFO, "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.CODPRODUCTO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
    SqlCad = SqlCad & "RESU.UM, "
    SqlCad = SqlCad & "SUM(RESU.CANTIDAD) AS CANTIDAD, "
    
    SqlCad = SqlCad & "VAL(COMPR.CANT & '') AS COMPROMISOEA, "
    SqlCad = SqlCad & "VAL(PORL.CANT & '') AS COMPROMISOPL, "
    
    SqlCad = SqlCad & "RESU.SALDO, "
    
    SqlCad = SqlCad & "RESU.COMPROMISOEAG, "
    SqlCad = SqlCad & "RESU.COMPROMISOPLG, "
    SqlCad = SqlCad & "RESU.LIBREEAG, "
    SqlCad = SqlCad & "RESU.LIBREPLG, "
    SqlCad = SqlCad & "RESU.STOCKEAG, "
    SqlCad = SqlCad & "RESU.STOCKPLG, "
    
    SqlCad = SqlCad & "(IIF(RESU.SALDO >= SUM(RESU.CANTIDAD), SUM(RESU.CANTIDAD), RESU.SALDO) - SUM(RESU.CANTIDADPC)) AS COMPRAR, "
    
    SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA, "
    SqlCad = SqlCad & "RESU.ATENDIDOPORPROV2 "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "(TMPUTILRESUMENCOMPROMISOAFECTADO AS RESU "

    SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, VAL(FORMAT(CANTIDAD, '#0.00')) AS CANT FROM TMPUTILRESUMENATENCIONCOMPROMISOAFECTADO WHERE TIPO = 'F') AS COMPR "
    SqlCad = SqlCad & "ON COMPR.NROPEDIDO = RESU.NROPEDIDO AND COMPR.CODPRODUCTO = RESU.CODPRODUCTO) "
    SqlCad = SqlCad & "LEFT JOIN (SELECT NROPEDIDO, CODPRODUCTO, VAL(FORMAT(CANTIDAD, '#0.00')) AS CANT FROM TMPUTILRESUMENATENCIONCOMPROMISOAFECTADO WHERE TIPO = 'V') AS PORL "
    SqlCad = SqlCad & "ON PORL.NROPEDIDO = RESU.NROPEDIDO AND PORL.CODPRODUCTO = RESU.CODPRODUCTO "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(RESU.CODPRODUCTO & '') <> '' AND "
    SqlCad = SqlCad & "RESU.DESCARTAR = FALSE "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "RESU.NROPEDIDO LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If

'        If CBool(chkProductoProveedor.value) Then
'            SqlCad = SqlCad & "AND RESU.ATENDIDOPORPROV = TRUE "
'        End If
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "RESU.NROPEDIDO, "
    SqlCad = SqlCad & "RESU.CLIENTE, "
    SqlCad = SqlCad & "RESU.FEMISION, "
    SqlCad = SqlCad & "RESU.FENTREGA, "
    SqlCad = SqlCad & "RESU.CODPRODUCTO, "
    SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
    SqlCad = SqlCad & "RESU.UM, "
    
    SqlCad = SqlCad & "VAL(COMPR.CANT & ''), "
    SqlCad = SqlCad & "VAL(PORL.CANT & ''), "
    
    SqlCad = SqlCad & "RESU.SALDO, "
    
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
            'Columna Informacion del Requerimiento
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Información de Requerimiento"
                .DisableEditor = True
                .FieldName = "INFO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColInfo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Nro Pedido
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "No. Requerimiento"
                .DisableEditor = True
                .FieldName = "NROPEDIDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColNroPedido"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
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
                .Caption = "Cant. a Reponer"
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
            
            'Columna Saldo en Produccion
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 1
                .Caption = "Saldo en OP's"
                .Color = vbWhite
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "SALDO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColSaldo"
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
                .Caption = "Cant. a Comprar"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "COMPRA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCompra"
                .SummaryFooterType = cstSum
                .Width = 70
            End With
            
            abrirCnTemporal

            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString

            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltOptimistic
            .Dataset.ADODataset.CommandText = SqlCad
            .Dataset.Active = True
            .Dataset.Refresh

            .KeyField = "LLAVE"
            
            .Columns.ColumnByFieldName("INFO").GroupIndex = 0
            
            .Columns.HeaderFont.Bold = True
            
            .m.FullExpand
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
        End With
    End If
    
    SqlCad = vbNullString

    Exit Sub
errCargarResumenCompromisoAfectadoVista1:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnTemporal

            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - CargarResumenCompromisoAfectadoVista1"
    End Select
    
    Err.Clear
End Sub

'Listar Vista 2 de Compromisos Afectados en Grilla (QuamtumGrid)
Private Sub cargarResumenCompromisoAfectadoVista2()

    On Error GoTo errCargarResumenCompromisoAfectadoVista2
    
    SqlCad = vbNullString
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "RES.LLAVE, "
    SqlCad = SqlCad & "('"
    SqlCad = SqlCad & "[Cliente: ' & RES.CLIENTE & ']" & Space(20)
    SqlCad = SqlCad & "[No Pedido: ' & RES.NROPEDIDO & ']" & Space(20)
    SqlCad = SqlCad & "[Fec. Emision: ' & RES.FEMISION & ']" & Space(20)
    SqlCad = SqlCad & "[Fec. Entrega: ' & RES.FENTREGA & ']" & Space(20)
    SqlCad = SqlCad & "[Vendedor: ' & RES.VENDEDOR & ']"
    SqlCad = SqlCad & "') AS INFO, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "RES.CODPRODUCTO, "
    SqlCad = SqlCad & "RES.NOMPRODUCTO, "
    SqlCad = SqlCad & "RES.UM, "
    SqlCad = SqlCad & "RES.NOMPRODUCTOUM, "
    SqlCad = SqlCad & "RES.USUARIO, "
    SqlCad = SqlCad & "RES.FECHA, "
    SqlCad = SqlCad & "RES.OBSERVACION, "
    SqlCad = SqlCad & "RES.CANTIDAD, "

    SqlCad = SqlCad & "RES.COMPROMISOEAG, "
    SqlCad = SqlCad & "RES.COMPROMISOPLG, "
    SqlCad = SqlCad & "RES.LIBREEAG, "
    SqlCad = SqlCad & "RES.LIBREPLG, "
    SqlCad = SqlCad & "RES.STOCKEAG, "
    SqlCad = SqlCad & "RES.STOCKPLG, "
    SqlCad = SqlCad & "RES.CANTIDADPC, "
    SqlCad = SqlCad & "RES.REPOSICION, "
    SqlCad = SqlCad & "RES.ATENDIDOPORPROV "
    
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO AS RES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "RES.NROPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "RES.CODPRODUCTO = '" & strCodProducto & "' AND "
    SqlCad = SqlCad & "RES.DESCARTAR = FALSE "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "RES.LLAVE LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "RES.USUARIO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
'        If CBool(chkProductoSeleccionado.value) Then
'            SqlCad = SqlCad & "AND RES.REPOSICION = TRUE "
'        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RES.FECHA"
    
    If Not dbgResumen Is Nothing Then

        With dbgResumen(1)
            .Dataset.Close

            .Columns.DestroyColumns
        End With

        Dim gColumn As dxGridColumn

        With dbgResumen(1)
            'Columna Llave
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "ID de Reposicion"
                .DisableEditor = True
                .FieldName = "LLAVE"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColLlave"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
            End With
            
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

            'Columna Usuario
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Redistribuido Por"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "USUARIO"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColUsuario"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 68
            End With

            'Columna Fecha
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Fecha Redist."
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "FECHA"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColFecha"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With

            'Columna Observacion
            Set gColumn = .Columns.Add(gedTextEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Observación de Redistribución"
                .Color = vbWhite
                .DisableEditor = True
                .FieldName = "OBSERVACION"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColObservacion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
            End With

            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)

            With gColumn
                .Alignment = taRightJustify
                .BandIndex = 0
                .Caption = "Cantidad"
                .Color = &HFFFFC0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANTIDAD"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
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
            
            'Columna Reposicion
            Set gColumn = .Columns.Add(gedCheckEdit)

            With gColumn
                .Alignment = taCenter
                .BandIndex = 5
                .Caption = "Reponer"
                .FieldName = "REPOSICION"
                .Font.Charset = 0
                .HeaderAlignment = taCenter
                .ObjectName = "ColReposicion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
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
            
            .Columns.ColumnByFieldName("INFO").GroupIndex = 0

            .Columns.ColumnByFieldName("CANTIDAD").SummaryFooterType = cstSum
            
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterType = cstCount
            .Columns.ColumnByFieldName("NOMPRODUCTO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount

            .m.FullExpand
        End With
    End If

    SqlCad = vbNullString
    
    Exit Sub
errCargarResumenCompromisoAfectadoVista2:
    Select Case Err.Number
        Case 3704, 3709
            abrirCnTemporal

            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - CargarResumenCompromisoAfectadoVista2"
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
    
    If MsgBox("¿Desea " & IIf(bolEstado, "reponer ", "cancelar la reposición d") & "el Item seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
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
    SqlCad = SqlCad & "RES.LLAVE, "
    SqlCad = SqlCad & "RES.CODPRODUCTO, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "RES.CANTIDAD, "
    SqlCad = SqlCad & "RES.CANTIDADPC, "
    SqlCad = SqlCad & "RES.REPOSICION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO AS RES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "RES.NROPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "RES.CODPRODUCTO = '" & strCodProducto & "' AND "
    SqlCad = SqlCad & "RES.DESCARTAR = FALSE "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "RES.LLAVE LIKE '%" & txtbusqueda.Text & "%' OR "
            SqlCad = SqlCad & "RES.USUARIO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
'        If CBool(chkProductoSeleccionado.value) Then
'            SqlCad = SqlCad & "AND RES.REPOSICION = TRUE "
'        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RES.FECHA"
    
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
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "CANTIDADPC = " & IIf(bolEstado, Val(rstTemporal!Cantidad & ""), "0") & ", "
            SqlCad = SqlCad & "REPOSICION = " & IIf(bolEstado, "TRUE", "FALSE") & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "LLAVE = '" & Trim(rstTemporal!LLAVE & "") & "'"
            
            abrirCnTemporal

            cnDBTemp.Execute SqlCad
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "SF1REPOSICIONCOMPROMISO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "REPOSICIONENCURSO = " & IIf(bolEstado, "TRUE", "FALSE") & ", "
            SqlCad = SqlCad & "USUREPOSICION = " & IIf(bolEstado, "'" & wusuario & "'", "NULL") & ", "
            SqlCad = SqlCad & "FECREPOSICION = " & IIf(bolEstado, "CVDATE('" & Now & "')", "NULL") & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "IDREPOSICION = " & Trim(rstTemporal!LLAVE & "")
            
            cnn_dbbancos.Execute SqlCad
            
            Actualiza_Log SqlCad, StrConexDbBancos
            
            If bolEstado Then
                tlbResumen.Tools("ID_Cantidad").Edit.Text = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) - IIf(Val(rstTemporal!Cantidad & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(rstTemporal!Cantidad & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))

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
    tlbResumen.Tools("Descartar").Enabled = Not bolEstado
    
    Me.MousePointer = vbDefault
    
    If dbgResumen(0).Visible Then
        cargarResumenCompromisoAfectadoVista1
    ElseIf dbgResumen(1).Visible Then
        cargarResumenCompromisoAfectadoVista2
    End If
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub descartarSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errDescartarSeleccion
    
    If MsgBox("¿Desea descartar la reposición del Item seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    dbgResumen(1).Dataset.Close
    
    Me.MousePointer = vbHourglass
    
    Dim rstTemporal As New ADODB.Recordset
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "RES.LLAVE, "
    SqlCad = SqlCad & "RES.CODPRODUCTO, "
    SqlCad = SqlCad & "RES.NROPEDIDO, "
    SqlCad = SqlCad & "RES.CANTIDAD, "
    SqlCad = SqlCad & "RES.CANTIDADPC, "
    SqlCad = SqlCad & "RES.REPOSICION "

    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO AS RES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "RES.NROPEDIDO = '" & strNroPedido & "' AND "
    SqlCad = SqlCad & "RES.CODPRODUCTO = '" & strCodProducto & "' AND "
    SqlCad = SqlCad & "RES.DESCARTAR = FALSE "
        
        If txtbusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            SqlCad = SqlCad & "RES.LLAVE LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & "RES.USUARIO LIKE '%" & txtbusqueda.Text & "%'"
            SqlCad = SqlCad & ") "
        End If
        
'        If CBool(chkProductoSeleccionado.value) Then
'            SqlCad = SqlCad & "AND RES.REPOSICION = TRUE "
'        End If
    
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "RES.FECHA"
    
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
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "DESCARTAR = " & IIf(bolEstado, "TRUE", "FALSE") & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "LLAVE = '" & Trim(rstTemporal!LLAVE & "") & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad

            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "SF1REPOSICIONCOMPROMISO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "DESCARTARREPOSICION = " & IIf(bolEstado, "TRUE", "FALSE") & ", "
            SqlCad = SqlCad & "USUDESCARTE = " & IIf(bolEstado, "'" & wusuario & "'", "NULL") & ", "
            SqlCad = SqlCad & "FECDESCARTE = " & IIf(bolEstado, "CVDATE('" & Now & "')", "NULL") & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "IDREPOSICION = " & Trim(rstTemporal!LLAVE & "")
            
            cnn_dbbancos.Execute SqlCad
    
            Actualiza_Log SqlCad, StrConexDbBancos
            
            pgbProgresoBusqueda.Value = pgbProgresoBusqueda.Value + 1

            rstTemporal.MoveNext
        Loop
    End If

    If rstTemporal.State = 1 Then rstTemporal.Close

    Set rstTemporal = Nothing

    pgbProgresoBusqueda.Visible = False

    SqlCad = vbNullString
    
    tlbResumen.Tools("Descartar").Enabled = False
    tlbResumen.Tools("Seleccionar").Enabled = False
    tlbResumen.Tools("QuitarSeleccion").Enabled = False
        
    Me.MousePointer = vbDefault
    
    cargarResumenCompromisoAfectadoVista2
    
    Exit Sub
errDescartarSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub evaluarCantidadDeCompra(ByVal dblSaldoActualEnOPs As Double, _
                                    ByVal dblCantidadaReponer As Double, _
                                    ByVal dblCantidadaComprar As Double)
    
    If dblSaldoActualEnOPs >= dblCantidadaReponer Then
        tlbResumen.Tools("ID_Cantidad").Edit.Text = dblCantidadaReponer
    Else
        tlbResumen.Tools("ID_Cantidad").Edit.Text = dblSaldoActualEnOPs
    End If
    
    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) - dblCantidadaComprar, "#0.00")
End Sub

Private Sub chkProductoProveedor_Click()
    cargarResumenCompromisoAfectadoVista1
End Sub

Private Sub chkProductoSeleccionado_Click()
    'cargarResumenCompromisoAfectadoVista2
End Sub

Private Sub cmbAlmacen_Click()
    'cargarResumenRequerimiento
End Sub

'Procedimiento Declarado para Selección y Vista de Detalle de Registro
Private Sub dbgResumen_RowColChange(Index As Integer)
    If dbgResumen(Index).Dataset.RecordCount > 0 Then
        Select Case Index
            Case 0
                strNroPedido = Trim(dbgResumen(0).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                strCodProducto = Trim(dbgResumen(0).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                
                evaluarCantidadDeCompra Val(dbgResumen(0).Columns.ColumnByFieldName("SALDO").Value & ""), _
                                        Val(dbgResumen(0).Columns.ColumnByFieldName("CANTIDAD").Value & ""), _
                                        Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRA").Value & "")
                
                If Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") = 0 And Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRA").Value & "") = 0 Then
                    tlbResumen.Tools("Seleccionar").Enabled = False
                    tlbResumen.Tools("QuitarSeleccion").Enabled = False
                Else
                    tlbResumen.Tools("Seleccionar").Enabled = IIf(Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") > 0, True, False)
                    tlbResumen.Tools("QuitarSeleccion").Enabled = IIf(Val(dbgResumen(0).Columns.ColumnByFieldName("COMPRAR").Value & "") = 0, True, False)
                End If
                
                'txtRegistro.Text = Trim(dbgResumen(Index).Dataset.RecNo & "")
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
    
    Me.MousePointer = vbHourglass
    
    cmdGenerar.Enabled = False

    Dim rstTemporal As New ADODB.Recordset
    Dim dblItem As Double
    Dim strUltimaDescripcion As String
    Dim dblUltimoPrecioSinIGv As Double
    'Dim dblUltimoDescuento As Double

    Dim strCuentaContable As String
    
    dbgResumen(0).Dataset.Close
    dbgResumen(1).Dataset.Close
    
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
            .Observacion = "REPOSICION DE COMPROMISOS AFECTADOS"
            
            .SUBTOTAL = 0
            .TotalInafecto = 0
            .TotalImpuesto = 0
            .TotalFacturado = 0
            
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
                SqlCad = SqlCad & "RESU.LLAVE, "
                SqlCad = SqlCad & "RESU.NROPEDIDO, "
                SqlCad = SqlCad & "RESU.CODPRODUCTO, "
                SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
                SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO AS RESU "
                SqlCad = SqlCad & "GROUP BY "
                SqlCad = SqlCad & "RESU.LLAVE, "
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
                            .ObservacionPorItem = "REPOSICION DE COMPROMISO AFECTADO."

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

                MsgBox "Orden Registrada con No." & .NumeroOrden & ".", vbInformation + vbOKOnly, App.ProductName

                txtCodProveedor.Text = vbNullString: lblProveedor.Caption = vbNullString

                listarOrdenesEnCombo

                SqlCad = vbNullString
                SqlCad = SqlCad & "DELETE FROM "
                SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "DESCARTAR = TRUE OR "
                SqlCad = SqlCad & "REPOSICION = TRUE"
                
                abrirCnTemporal

                cnDBTemp.Execute SqlCad

                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
                SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"
                
                abrirCnTemporal
                
                cnDBTemp.Execute SqlCad
                
                descargarAtencionCompromisoAfectado
                
                cargarResumenCompromisoAfectadoVista1

                'dbgResumen_OnClick 0
            End If
        Else
            If rstTemporal.State = 1 Then rstTemporal.Close
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            'SqlCad = SqlCad & "(RESU.NROPEDIDO & RESU.CODPRODUCTO) AS LLAVE, "
            SqlCad = SqlCad & "RESU.LLAVE, "
            SqlCad = SqlCad & "RESU.NROPEDIDO, "
            SqlCad = SqlCad & "RESU.CODPRODUCTO, "
            SqlCad = SqlCad & "RESU.NOMPRODUCTO, "
            SqlCad = SqlCad & "SUM(RESU.CANTIDADPC) AS COMPRA "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO AS RESU "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "RESU.LLAVE, "
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

                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOP 1 VAL(ITEM & '')", "IF3ORDEN", "F4LOCAL", .TipoOrden, "T", "AND F4NUMORD = '" & .NumeroOrden & "' ORDER BY VAL(ITEM & '') DESC"))
                
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
                        .ObservacionPorItem = "REPOSICION DE COMPROMISO AFECTADO."
                        
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

            MsgBox "Orden Registrada con No." & .NumeroOrden & ".", vbInformation + vbOKOnly, App.ProductName

            txtCodProveedor.Text = vbNullString: lblProveedor.Caption = vbNullString

            listarOrdenesEnCombo

            SqlCad = vbNullString
            SqlCad = SqlCad & "DELETE FROM "
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "DESCARTAR = TRUE OR "
            SqlCad = SqlCad & "REPOSICION = TRUE"
            
            abrirCnTemporal

            cnDBTemp.Execute SqlCad

            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "ATENDIDOPORPROV = FALSE, "
            SqlCad = SqlCad & "ATENDIDOPORPROV2 = 0"

            abrirCnTemporal

            cnDBTemp.Execute SqlCad
            
            descargarAtencionCompromisoAfectado

            cargarResumenCompromisoAfectadoVista1

            'dbgResumen_OnClick 0
        End If
        
        .inicializarEntidades
        .inicializarEntidadesDetalle
    End With
    
    cmdGenerar.Enabled = True
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgResumen_OnChangeNode(Index As Integer, ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    dbgResumen_RowColChange Index
End Sub

Private Sub dbgResumen_OnClick(Index As Integer)
    dbgResumen_RowColChange Index
End Sub

Private Sub dbgResumen_OnCheckEditToggleClick(Index As Integer, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "REPOSICION"
            With dbgResumen(1)
                If .Dataset.State = dsEdit Then
                    .Dataset.Post
                End If
                
                .Dataset.Edit
                
                If CBool(.Columns.ColumnByFieldName("REPOSICION").Value) Then
                    If MsgBox("¿Desea reponer el Item seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                        .Columns.ColumnByFieldName("REPOSICION").Value = False
                    End If
                End If
                
                If CBool(.Columns.ColumnByFieldName("REPOSICION").Value) Then
                    If Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") = 0 Then
                        .Columns.ColumnByFieldName("CANTIDADPC").Value = IIf(Val(.Columns.ColumnByFieldName("CANTIDAD").Value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(.Columns.ColumnByFieldName("CANTIDAD").Value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text))
                        
                        tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) - Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), "#0.00")
                    End If
                Else
                    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) + Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), "#0.00")
                    
                    .Columns.ColumnByFieldName("CANTIDADPC").Value = 0
                End If
                
                .Dataset.Post
                
                Me.MousePointer = vbHourglass
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "SF1REPOSICIONCOMPROMISO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "REPOSICIONENCURSO = " & IIf(CBool(.Columns.ColumnByFieldName("REPOSICION").Value), "TRUE", "FALSE") & ", "
                SqlCad = SqlCad & "USUREPOSICION = " & IIf(CBool(.Columns.ColumnByFieldName("REPOSICION").Value), "'" & wusuario & "'", "NULL") & ", "
                SqlCad = SqlCad & "FECREPOSICION = " & IIf(CBool(.Columns.ColumnByFieldName("REPOSICION").Value), "CVDATE('" & Now & "')", "NULL") & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDREPOSICION = " & Val(.Columns.ColumnByFieldName("LLAVE").Value & "")
                
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
                
                If Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0 Then
                    tlbResumen.Tools("Seleccionar").Enabled = True
                    tlbResumen.Tools("QuitarSeleccion").Enabled = False
                Else
                    tlbResumen.Tools("Seleccionar").Enabled = False
                    tlbResumen.Tools("QuitarSeleccion").Enabled = True
                End If
                
                Me.MousePointer = vbDefault
            End With
    End Select
End Sub

Private Sub dbgResumen_OnCustomDrawCell(Index As Integer, ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "SALDO"
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
            If Val(Node.Values(16) & "") = 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(75, 172, 198)
            End If
    End Select
End Sub

Private Sub dbgResumen_OnCustomDrawFooter(Index As Integer, ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "COMPRA", "CANTIDADPC"
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
        Case "COMPROMISOEAG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString

                .DeshabilitarRedistribucion = True 'IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "") > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                '.CantidadMaximaParaPedido = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)) 'Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                
                .Show 1
            End With
        Case "COMPROMISOPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString

                .DeshabilitarRedistribucion = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                '.CantidadMaximaParaPedido = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)) 'Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)

                .Show 1
            End With
        Case "LIBREEAG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString

                .DeshabilitarRedistribucion = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                '.CantidadMaximaParaPedido = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)) 'Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)
                
                .Show 1
            End With
        Case "LIBREPLG"
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgResumen(Index).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
                .CodigoAlmacen = vbNullString

                .DeshabilitarRedistribucion = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0, False, True)
                .NroPedidoSolicitante = Trim(dbgResumen(Index).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
                '.CantidadMaximaParaPedido = IIf(Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "") <= Val(tlbResumen.Tools("ID_Cantidad").Edit.Text), Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & ""), Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)) 'Val(dbgResumen(Index).Columns.ColumnByFieldName("CANTIDAD").value & "")
                .CantidadMaximaParaPedido = Val(tlbResumen.Tools("ID_Cantidad").Edit.Text)

                .Show 1
            End With
        Case "COMPRAR", "COMPRA"
            If dbgResumen(0).Visible Then
                If CBool(tlbResumen.Tools("Seleccionar").Enabled) Then
                    tlbResumen_ToolClick tlbResumen.Tools("Seleccionar")
                ElseIf CBool(tlbResumen.Tools("QuitarSeleccion").Enabled) Then
                    tlbResumen_ToolClick tlbResumen.Tools("QuitarSeleccion")
                End If
            End If
        Case Else
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

                descargarAtencionCompromisoAfectado
                
                evaluarCantidadDeCompra Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SALDO", "TMPUTILRESUMENCOMPROMISOAFECTADO", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' GROUP BY NROPEDIDO, CODPRODUCTO")), _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "CANTIDAD", "TMPUTILRESUMENCOMPROMISOAFECTADO", "CODPRODUCTO", strCodProducto, "T", "AND NROPEDIDO = '" & strNroPedido & "' GROUP BY NROPEDIDO, CODPRODUCTO")), _
                                        Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENCOMPROMISOAFECTADO", "CODPRODUCTO", strCodProducto, "T", "GROUP BY CODPRODUCTO, NOMPRODUCTO, UM"))
                
                With objAyudaVale
                    .CodigoProducto = strCodProducto
                    
                    .verificarStockProducto

                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "TMPUTILRESUMENCOMPROMISOAFECTADO "
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
                
                'cargarResumenCompromisoAfectadoVista2
                Select Case Index
                    Case 0
                        cargarResumenCompromisoAfectadoVista1
                    Case 1
                        cargarResumenCompromisoAfectadoVista2
                        
                        tlbResumen.Tools("Anterior").Enabled = True
                End Select
            End If
    End Select
    
    If dbgResumen(Index).Dataset.RecordCount >= nSaveRecNo Then
        dbgResumen(Index).Dataset.RecNo = nSaveRecNo
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgResumen_OnEdited(Index As Integer, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgResumen(1).Columns.FocusedColumn.FieldName
        Case "CANTIDADPC"
            With dbgResumen(1)
                .Dataset.Edit
                
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

                If Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").Value & "") > Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDAD").Value & "") Then
                    MsgBox "La cantidad no puede exceder al Saldo por Atender, verifique.", vbInformation + vbOKOnly, App.ProductName

                    .Dataset.Cancel

                    Exit Sub
                End If

'                If Val(.Columns.ColumnByFieldName("CANTIDADPC").value & "") > 0 And (Val(.Columns.ColumnByFieldName("CANTIDADPC").SummaryFooterValue & "") + Val(.Columns.ColumnByFieldName("CANTIDADPC").value & "")) > Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) Then
'                    MsgBox "La cantidad no puede exceder la Cantidad Maxima a Comprar, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                    .Dataset.Cancel
'
'                    Exit Sub
'                End If

                .Columns.ColumnByFieldName("PROCESAR").Value = IIf(Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & "") <= 0, False, True)

                If CBool(.Columns.ColumnByFieldName("PROCESAR").Value) Then
                    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) - Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), "#0.00")
                Else
                    tlbResumen.Tools("ID_Cantidad").Edit.Text = Format(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) + Val(.Columns.ColumnByFieldName("CANTIDADPC").Value & ""), "#0.00")
                End If

                .Dataset.Post
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "SF1REPOSICIONCOMPROMISO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "REPOSICIONENCURSO = " & IIf(CBool(.Columns.ColumnByFieldName("PROCESAR").Value), "TRUE", "FALSE") & ", "
                SqlCad = SqlCad & "USUREPOSICION = " & IIf(CBool(.Columns.ColumnByFieldName("PROCESAR").Value), "'" & wusuario & "'", "NULL") & ", "
                SqlCad = SqlCad & "FECREPOSICION = " & IIf(CBool(.Columns.ColumnByFieldName("PROCESAR").Value), "CVDATE('" & Now & "')", "NULL") & " "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDREPOSICION = " & Val(.Columns.ColumnByFieldName("LLAVE").Value & "")
                
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
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

Private Sub dtpDesde_Change()
    'dtpDesde_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtpDesde_CloseUp()
    'dtpDesde_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CDate(dtpDesde.Value) > CDate(dtpHasta.Value) Then
                MsgBox "El campo 'Desde' no puede ser mayor al campo 'Hasta'.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If MsgBox("¿Desea ejecutar la consulta?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                bolResumenCargado = False
        
                Form_Activate
            End If
    End Select
End Sub


Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CDate(dtpHasta.Value) < CDate(dtpDesde.Value) Then
                MsgBox "El campo 'Hasta' no puede ser menor al campo 'Desde'.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If MsgBox("¿Desea ejecutar la consulta?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                bolResumenCargado = False
        
                Form_Activate
            End If
    End Select
End Sub


Private Sub Form_Activate()
    If Not bolResumenCargado Then
        FraBusqueda.Enabled = bolResumenCargado
        fraRango.Enabled = bolResumenCargado
        fraOpciones.Enabled = bolResumenCargado
        tlbResumen.Enabled = bolResumenCargado
        fraProveedor.Enabled = bolResumenCargado
        dbgResumen(0).Enabled = bolResumenCargado
        
        Me.MousePointer = vbHourglass
        
        bolResumenCargado = True
        
        tlbResumen.Tools("ID_Inicio").Edit.Text = "Inicio: " & Format(Now, "hh:mm:ss AM/PM")
                
        descargarResumenCompromisoAfectado
        
        descargarSaldoEnOPsDeCompromisoAfectado
        
        descargarAtencionCompromisoAfectado
        
        cargarStockDeProducto
        
        Me.MousePointer = vbDefault
        
        cargarResumenCompromisoAfectadoVista1
        
        FraBusqueda.Enabled = bolResumenCargado
        fraRango.Enabled = bolResumenCargado
        fraOpciones.Enabled = bolResumenCargado
        tlbResumen.Enabled = bolResumenCargado
        fraProveedor.Enabled = IIf(VerificaAutorizaciones("OCN", wusuario) <> "''", bolResumenCargado, False)
        
        dbgResumen(0).Enabled = bolResumenCargado
        
        tlbResumen.Tools("ID_Fin").Edit.Text = "Fin: " & Format(Now, "hh:mm:ss AM/PM")
        
        dbgResumen_OnClick 0
        
        txtbusqueda.SetFocus
    End If
End Sub

Private Sub Form_Load()
    bolResumenCargado = False
    
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    inicializarControles
    
'    'Para Control de Anterior/Siguiente
'    tmrTemporizador.Enabled = False
'    tmrTemporizador.Interval = 0
'    tlbResumen.Tools("Anterior").Enabled = False
'    tlbResumen.Tools("Siguiente").Enabled = True
'
'    intIndiceGrilla = 0
'    intIndiceVisible = 0
'    intIndiceOculto = 0
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilReposicionCompromiso
    
    'Activar Control de Apertura de Formulario
    '(Para evitar abrir mas de una vez, el mismo formulario en diferentes Instancias del Programa)
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", "1"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If strNroPedido <> vbNullString Then
        If MsgBox("¿Desea salir de la Consulta actual?" & vbNewLine & "RECUERDA: La actual consulta se provee de datos Externos, por lo cual tiende a tardar algunos minutos en cargar.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
            
            dbgResumen(0).Dataset.ADODataset.Requery
                        
            dbgResumen(0).m.ResetFullRefresh
            
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

'    dblFactorAncho = dbgResumen(intIndiceGrilla).Width / 40
End Sub

Private Sub tlbResumen_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
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
        Case "Descartar"
            descartarSeleccion True
            
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
            
            Me.Caption = "Reposición de Compromisos Afectados"
            
            dbgResumen(1).Dataset.Close
            
            dbgResumen(1).Visible = False
            dbgResumen(0).Visible = True
            
            strNroPedido = vbNullString
            strCodProducto = vbNullString
            
            txtbusqueda.Text = vbNullString
            
            chkProductoSeleccionado.Enabled = False
            chkProductoProveedor.Enabled = True

            tlbResumen.Tools("ID_Cantidad").Edit.Text = "0.00"
            tlbResumen.Tools("ID_EstadoAtte").Edit.Text = "--"

            tlbResumen.Tools("Seleccionar").Enabled = False
            tlbResumen.Tools("QuitarSeleccion").Enabled = False
            tlbResumen.Tools("Descartar").Enabled = False
            
            tlbResumen.Tools("ID_Salir").Enabled = True

            If VerificaAutorizaciones("OCN", wusuario) <> "''" Then
                fraProveedor.Enabled = True
            End If

            cargarResumenCompromisoAfectadoVista1

            dbgResumen_OnClick 0
            
            Screen.MousePointer = vbDefault
        Case "Siguiente"
            Screen.MousePointer = vbHourglass

            tlbResumen.Tools("Siguiente").Enabled = False
            tlbResumen.Tools("Anterior").Enabled = True
            
            If VerificaAutorizaciones("OCN", wusuario) <> "''" Then
                fraProveedor.Enabled = False
            End If
            
            Me.Caption = Trim(dbgResumen(0).Columns.ColumnByFieldName("INFO").Value & "")
            dbgResumen(1).Bands(0).Caption = Trim(dbgResumen(0).Columns.ColumnByFieldName("NOMPRODUCTO").Value & "") & " ( " & Trim(dbgResumen(0).Columns.ColumnByFieldName("UM").Value & "") & " )"
            
            strNroPedido = Trim(dbgResumen(0).Columns.ColumnByFieldName("NROPEDIDO").Value & "")
            strCodProducto = Trim(dbgResumen(0).Columns.ColumnByFieldName("CODPRODUCTO").Value & "")
            
            dbgResumen(0).Dataset.Close
            
            dbgResumen(0).Visible = False
            dbgResumen(1).Visible = True
            
            txtbusqueda.Text = vbNullString
            
            chkProductoSeleccionado.Enabled = True
            chkProductoProveedor.Enabled = False
            
            tlbResumen.Tools("ID_Salir").Enabled = False
            
            cargarResumenCompromisoAfectadoVista2
            
            tlbResumen.Tools("Seleccionar").Enabled = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0 And Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").SummaryFooterValue & "") = 0, True, False)
            tlbResumen.Tools("QuitarSeleccion").Enabled = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0 And Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").SummaryFooterValue & "") > 0, True, False)
            tlbResumen.Tools("Descartar").Enabled = IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) = 0 And Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").SummaryFooterValue & "") = 0, True, IIf(Val(tlbResumen.Tools("ID_Cantidad").Edit.Text) > 0 And Val(dbgResumen(1).Columns.ColumnByFieldName("CANTIDADPC").SummaryFooterValue & "") = 0, True, False))
            
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
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "SUM(CANTIDADPC)", "TMPUTILRESUMENCOMPROMISOAFECTADO", vbNullString, vbNullString, vbNullString, "TRIM(CODPRODUCTO & '') <> ''")) > 0 Then
                'If MsgBox("Se cuenta con Items marcados para compra, " & vbNewLine & _
                '            "¿Desea salir sin generar Orden de Compra?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                '
                    
                    MsgBox "Se cuenta con Items marcados para compra, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    If dbgResumen(0).Visible Then
                        dbgResumen(0).Dataset.ADODataset.Requery
                        
                        dbgResumen(0).m.ResetFullRefresh
                    End If
                    
                    Exit Sub
                'End If
            End If
            
            Unload Me
    End Select
End Sub

'Private Sub tmrTemporizador_Timer()
'    On Error GoTo errTmrTemporizador
'
'    If tmrTemporizador.Interval = 40 Then
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
                cargarResumenCompromisoAfectadoVista1
            ElseIf CBool(dbgResumen(1).Visible) Then
                cargarResumenCompromisoAfectadoVista2
            End If
        Case vbKeyEscape
            If CBool(dbgResumen(1).Visible) Then
                tlbResumen_ToolClick tlbResumen.Tools("Anterior")
            End If
    End Select
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

                cargarProductoAtendidoPorProveedor

                cargarResumenCompromisoAfectadoVista1

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

                    cargarResumenCompromisoAfectadoVista1
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

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtRegistro_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub


