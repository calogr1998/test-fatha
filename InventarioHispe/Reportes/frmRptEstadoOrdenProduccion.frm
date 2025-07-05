VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRptEstadoOrdenProduccion 
   Caption         =   "Estado de Ordenes de Producción"
   ClientHeight    =   8265
   ClientLeft      =   225
   ClientTop       =   1785
   ClientWidth     =   15375
   Icon            =   "frmRptEstadoOrdenProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin DXDBGRIDLibCtl.dxDBGrid dbgPedido 
      Height          =   2895
      Left            =   1080
      OleObjectBlob   =   "frmRptEstadoOrdenProduccion.frx":058A
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Proceso en curso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3960
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   8175
      Begin MSComctlLib.ProgressBar pgbProceso2 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbProceso1 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
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
         TabIndex        =   19
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label lblProceso2 
         Caption         =   "Proceso 2"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   7695
      End
   End
   Begin VB.Frame fraNroPedidoFiltro 
      Caption         =   " No. Pedido "
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
      Height          =   975
      Left            =   13920
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1365
      Begin VB.TextBox txtNroPedidoFiltro 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1080
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
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtAnno 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8160
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNroPedido 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFecValidez 
         Height          =   285
         Left            =   7800
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   114819073
         CurrentDate     =   41939
      End
      Begin VB.Label lblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Entrega Valido"
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblNroPedido 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "No. Pedido"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
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
      Left            =   9480
      TabIndex        =   0
      Top             =   120
      Width           =   4365
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3960
      End
   End
   Begin MSComDlg.CommonDialog cmdlgOrden 
      Left            =   0
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgOrden 
      Height          =   4395
      Left            =   120
      OleObjectBlob   =   "frmRptEstadoOrdenProduccion.frx":1205
      TabIndex        =   2
      Top             =   1200
      Width           =   15135
   End
   Begin MSComctlLib.ImageList imgLstEstado 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":1E76
            Key             =   "Estado 1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":2410
            Key             =   "Estado 2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":29AA
            Key             =   "Estado 3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":2F44
            Key             =   "Estado 4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":34DE
            Key             =   "Estado 5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":3A78
            Key             =   "Estado 6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":4012
            Key             =   "Estado 7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptEstadoOrdenProduccion.frx":45AC
            Key             =   "Estado 8"
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars tlbOrden 
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
      Tools           =   "frmRptEstadoOrdenProduccion.frx":4B46
      ToolBars        =   "frmRptEstadoOrdenProduccion.frx":115FC
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgOrdenDetalle 
      Height          =   2250
      Left            =   120
      OleObjectBlob   =   "frmRptEstadoOrdenProduccion.frx":117CB
      TabIndex        =   3
      Top             =   5640
      Width           =   15105
   End
End
Attribute VB_Name = "frmRptEstadoOrdenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda                                As Boolean

Private strFichero                              As String
Private strTablaUsuario                         As String
Private strFechaCorteValidezCompromiso          As String
Private intCantidadMesesDeValidezCompromiso     As Integer


Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double

Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property



'Procedimiento Declarado para Selección y Vista de Detalle de Registro
Private Sub dbgDocumento_RowColChange()
    If dbgOrden.Dataset.RecordCount > 0 Then
        listarOrdenDetalle
    End If
End Sub

Private Sub dbgOrden_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgOrden.Columns.FocusedColumn.FieldName
        Case "CATEGORIA", "NROOP"
            dbgDocumento_RowColChange
    End Select
End Sub

Private Sub dbgOrden_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "F4VB1"
'            With dbgOrden
'                .Dataset.Edit
'
'                With objAyudaOrden
'                    .inicializarEntidades
'
'                    .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
'                    .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
'
'                    .obtenerConfigOrden
'
'                    If .Estado = 8 Then
'                        MsgBox "Imposible realizar esta acción, registro Anulado.", vbInformation + vbOKOnly, App.ProductName
'
'                        dbgOrden.Dataset.Cancel
'
'                        Exit Sub
'                    End If
'
'                    If .Estado >= 3 Then
'                        MsgBox "Imposible realizar esta acción, registro en etapa superior.", vbInformation + vbOKOnly, App.ProductName
'
'                        dbgOrden.Dataset.Cancel
'
'                        Exit Sub
'                    End If
'                End With
'
'                If Not CBool(objAyudaOrden.VB1) Then
'                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0006'") = "0006" Then
'                        If MsgBox("¿Desea DAR su VºBº al registro seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'                            .Columns.ColumnByFieldName("F4ESTADO").value = 2
'                            .Columns.ColumnByFieldName("F4VB1").value = True
'                            .Columns.ColumnByFieldName("F4VBUSER1").value = UCase(wusuario)
'                            .Columns.ColumnByFieldName("F4VBFECHA1").value = Now
'
''                            objAyudaOrden.Estado = 2
''                            objAyudaOrden.VB1 = True
''                            objAyudaOrden.VB1Usuario = UCase(wusuario)
''                            objAyudaOrden.VB1Fecha = Now
''
''                            'If objAyudaOrden.aprobarOrden Then
''
''                            'Else
''                                .Columns.ColumnByFieldName("F4VB1").value = False
''                            'End If
'                        Else
'                            '.Columns.ColumnByFieldName("F4VB1").value = False
'
'                            .Dataset.Cancel
'
'                            Exit Sub
'                        End If
'                    Else
'                        MsgBox "Ud. no cuenta con permisos para realizar esta acción.", vbInformation + vbOKOnly, App.ProductName
'
'                        .Dataset.Cancel
'
'                        Exit Sub
'                    End If
'                Else
'                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0011'") = "0011" Then
'                        If MsgBox("¿Desea QUITAR su VºBº al registro seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
''                            objAyudaOrden.Estado = 1
''                            objAyudaOrden.VB1 = False
''                            objAyudaOrden.VB1Usuario = vbNullString
''                            objAyudaOrden.VB1Fecha = vbNullString
''
''                            objAyudaOrden.aprobarOrden
'
'                            .Columns.ColumnByFieldName("F4ESTADO").value = 1
'                            .Columns.ColumnByFieldName("F4VB1").value = False
'                            .Columns.ColumnByFieldName("F4VBUSER1").value = vbNullString
'                            .Columns.ColumnByFieldName("F4VBFECHA1").value = Null
'                        Else
'                            '.Columns.ColumnByFieldName("F4VB1").value = True
'
'                            .Dataset.Cancel
'
'                            Exit Sub
'                        End If
'                    Else
'                        MsgBox "Ud. no cuenta con permisos para realizar esta acción.", vbInformation + vbOKOnly, App.ProductName
'
'                        .Dataset.Cancel
'
'                        Exit Sub
'                    End If
'                End If
'
'                .Dataset.Post
'
'                With objAyudaOrden
'                    .Estado = Val(dbgOrden.Columns.ColumnByFieldName("F4ESTADO").value & "")
'                    .VB1 = CBool(dbgOrden.Columns.ColumnByFieldName("F4VB1").value)
'                    .VB1Usuario = UCase(Trim(dbgOrden.Columns.ColumnByFieldName("F4VBUSER1").value & ""))
'                    .VB1Fecha = IIf(Not IsNull(dbgOrden.Columns.ColumnByFieldName("F4VBFECHA1").value), Trim(dbgOrden.Columns.ColumnByFieldName("F4VBFECHA1").value & ""), vbNullString)
'
'                    .aprobarOrden
'                End With
'            End With
    End Select
End Sub

Private Sub dbgOrden_OnClick()
    Select Case dbgOrden.Columns.FocusedColumn.FieldName
        Case "CATEGORIA", "NROOP"
            dbgDocumento_RowColChange
    End Select
End Sub

Private Sub dbgOrden_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "FENTREGA"
            Text = Format(Text, "Short Date")
            
            If CDate(Text) >= CDate(strFechaCorteValidezCompromiso) Then
                Color = RGB(0, 176, 80)
                Font.Bold = True
                FontColor = vbWhite
            Else
                Color = vbRed
                Font.Bold = True
                FontColor = vbWhite
            End If
    End Select
End Sub

Private Sub dbgOrden_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyUp
            If dbgOrden.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub dbgOrden_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, r As Single, b As Single, NeedShowTip As Boolean)
    Select Case Column.FieldName
        Case "DESCRIPCIONOP"
            NeedShowTip = True
        Case "FENTREGA"
            NeedShowTip = True
            
            If CDate(Format(Node.Values(7), "Short Date")) >= CDate(strFechaCorteValidezCompromiso) Then
                TipText = "O/P Vigente"
            Else
                TipText = "O/P Vencida"
            End If
        Case Else
            NeedShowTip = False
    End Select
End Sub

Private Sub dbgOrdenDetalle_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "COD_SOLICITUD"
            If Trim(Text) <> "STOCK LIBRE" Then
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(79, 129, 189)
            End If
        Case "CANTIDAD"
            Text = Format(Val(Text), "#,0.00;(#,0.00)")
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
            
            Text = Format(Val(Text), "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgPedido_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyEscape
            ModUtilitario.seleccionarTextoCaja txtNroPedido
            
            dbgPedido.Dataset.Close
            
            dbgPedido.Visible = False
            
            txtNroPedido.SetFocus
        Case vbKeyReturn
            txtNroPedido.Text = Trim(dbgPedido.Columns.ColumnByFieldName("IDPEDIDO").value & "")
            lblNroPedido.Caption = Trim(dbgPedido.Columns.ColumnByFieldName("RESUMEN").value & "")
            
            dbgPedido.Dataset.Close
            
            dbgPedido.Visible = False
            
            txtNroPedido.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    'abrirCnDBMilano
    
    'Instanciar ruta de Fichero de Configuracion por Usuario
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    strTablaUsuario = "TMPCPRESUMENPRODUCCIONESTADOOP" & wusuario
    
    intCantidadMesesDeValidezCompromiso = Val(ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigSQLCliente, "ConfigServidorSQLCliente", "CantidadMesesDeValidezCompromiso", "l"))
    strFechaCorteValidezCompromiso = Format(Date - (30 * intCantidadMesesDeValidezCompromiso), "Short Date")
    
    inicializarControles
    
    listarOrden
End Sub

Private Sub inicializarControles()
    txtCodProducto.Text = vbNullString: lblProducto.Caption = vbNullString
    txtNroPedido.Text = vbNullString: lblNroPedido.Caption = vbNullString
    
    dtpFecValidez.value = strFechaCorteValidezCompromiso
    
    txtBusqueda.Text = vbNullString
    
    txtAnno.Text = Format(strFechaCorteValidezCompromiso, "yyyy")
End Sub

Private Sub consultarEstadoOP()
    If Trim(txtNroPedido.Text) = vbNullString Then
        MsgBox "Ingrese Nº de Pedido para realizar la consulta.", vbInformation + vbOKOnly, App.ProductName
        
        txtNroPedido.SetFocus
        
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    dbgOrdenDetalle.Dataset.Close
    dbgOrden.Dataset.Close
    
    
    fraDatos.Enabled = False
    fraBusqueda.Enabled = False
    fraNroPedidoFiltro.Enabled = False
    'fraProceso.Visible = True
    
'    If ModMilano.importarResumenRequerimientoProduccionV2(Nothing, _
                                                            Nothing, _
                                                            strTablaUsuario, _
                                                            Trim(txtNroPedido.Text), _
                                                            Trim(txtCodProducto.Text), _
                                                            vbNullString, _
                                                            dtpFecValidez.value, _
                                                            True) Then
    
    If ModMilano.importarResumenRequerimientoProduccionV2(Nothing, _
                                                            Nothing, _
                                                            strTablaUsuario, _
                                                            Trim(txtNroPedido.Text), _
                                                            Trim(txtCodProducto.Text), _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            True) Then
        listarOrden
    End If
    
'    If ModMilano.importarResumenRequerimientoProduccionV3(lblProceso1, _
'                                                            pgbProceso1, _
'                                                            lblProceso2, _
'                                                            pgbProceso2, _
'                                                            strTablaUsuario, _
'                                                            Trim(txtNroPedido.Text), _
'                                                            Trim(txtCodProducto.Text), _
'                                                            vbNullString, _
'                                                            dtpFecValidez.value, _
'                                                            True) Then
'
'        listarOrden
'    End If
    
    fraDatos.Enabled = True
    fraBusqueda.Enabled = True
    fraNroPedidoFiltro.Enabled = True
    'fraProceso.Visible = False
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub listarOrden()
    If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NAME", "SYSOBJECTS", "NAME", strTablaUsuario, "T") = vbNullString Then
        MsgBox "Presionar [Consultar] para cargar la consulta.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    dbgOrden.Dataset.Close
    
    ModMilano.vistaEstadoOrdenProduccion dbgOrden, strTablaUsuario, vbNullString, txtBusqueda.Text, Trim(txtNroPedidoFiltro.Text)
    
    dbgDocumento_RowColChange
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub listarOrdenDetalle()
    If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NAME", "SYSOBJECTS", "NAME", strTablaUsuario, "T") = vbNullString Then
        Exit Sub
    End If
    
    ModMilano.vistaEstadoOrdenProduccionDetalle dbgOrdenDetalle, strTablaUsuario, Trim(dbgOrden.Columns.ColumnByFieldName("IDOP").value & ""), vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbOrden_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ConsultarOP"
            consultarEstadoOP
        Case "Filtrar"
            If Tool.State = ssChecked Then
                dbgOrden.Filter.FilterActive = True
            Else
                dbgOrden.Filter.FilterActive = False
            End If
        Case "Agrupar"
            If Tool.State = ssChecked Then
                dbgOrden.Options.Set (egoShowGroupPanel)
            Else
                dbgOrden.Options.Unset (egoShowGroupPanel)
            End If
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgOrden
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dbgOrden.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Movimiento"
            Dim rpt As New rptOPMovimiento
            
            With rpt
                .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("IDOP").value & "")
                
                .Show 1
            End With
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim dblPromH As Double
    
    dblPromH = Me.ScaleHeight / 6
    
    fraProceso.Move (Me.ScaleWidth / 2) - (fraProceso.Width / 2), fraProceso.top, fraProceso.Width, fraProceso.Height
    
    fraDatos.Move 0, 0, fraDatos.Width, fraDatos.Height
    fraBusqueda.Move fraDatos.Width + 20, 0, fraBusqueda.Width, fraBusqueda.Height
    fraNroPedidoFiltro.Move fraDatos.Width + fraBusqueda.Width + 20, 0, fraNroPedidoFiltro.Width, fraNroPedidoFiltro.Height
    
    dbgOrden.Move 0, fraDatos.Height, Me.ScaleWidth, (dblPromH * 1.5) - 400
    
    dbgOrdenDetalle.Move 0, fraDatos.Height + dbgOrden.Height, Me.ScaleWidth, (dblPromH * 4.5) - 600
    
    'fraBusqueda.Width = dbgOrden.Width
    'txtBusqueda.Width = dbgOrden.Width - 350
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarOrden
        Case vbKeyDown
            dbgOrden.SetFocus
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
                    lblProducto.Caption = objAyudaBien.Descripcion
                    lblProducto.ToolTipText = lblProducto.Caption
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.ToolTipText = lblProducto.Caption
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            lblProducto.Caption = "Todos los Productos (*)"
            lblProducto.ToolTipText = vbNullString
        End If
    Else
        lblProducto.Caption = "Todos los Productos (*)"
        lblProducto.ToolTipText = vbNullString
    End If
End Sub

Private Sub txtNroPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'abrirCnDBMilano
            
'            lblNroPedido.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "(PER.NOMBRE + ' ( FEC. PEDIDO: ' + CONVERT(CHAR(10), PED.FECHAEMISION, 103) + ' / FEC. ENTREGA: ' + CONVERT(CHAR(10), PED.FECHAENTREGA, 103) + ')') AS RESUMEN", "PEDIDO AS PED LEFT JOIN PERSONA AS PER ON PER.IDPERSONA = PED.IDPERSONA", "PED.IDPEDIDO", Trim(txtNroPedido.Text), "T")
'            lblNroPedido.ToolTipText = lblNroPedido.Caption
'
'            If lblNroPedido.Caption = vbNullString Then
'                MsgBox "No. de Pedido no encontrado o inválido.", vbInformation + vbOKOnly, App.ProductName
'
'                txtNroPedido.SetFocus
'            Else
'                ModUtilitario.pulsarTecla vbKeyTab
'            End If
            
            dbgPedido.Visible = True
            
            ModMilano.vistaPedidoDetalle dbgPedido, Trim(txtAnno.Text), txtNroPedido.Text
            
            dbgPedido.SetFocus
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

Private Sub txtNroPedidoFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarOrden
        Case vbKeyDown
            dbgOrden.SetFocus
    End Select
End Sub


Private Sub txtNroPedidoFiltro_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub


