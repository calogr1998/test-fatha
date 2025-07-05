VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptResComprasProveedores 
   Caption         =   "Resumen de Compras por Proveedores"
   ClientHeight    =   5595
   ClientLeft      =   345
   ClientTop       =   1920
   ClientWidth     =   16200
   Icon            =   "frmRptResComprasProveedores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   16200
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlgReporte 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraReporte 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   15975
      Begin VB.CheckBox chkVerColumnaFechaCompra 
         Caption         =   "Ver Columnas de Fechas Adicionales."
         Height          =   255
         Left            =   12480
         TabIndex        =   17
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtProducto 
         Height          =   285
         Left            =   9840
         TabIndex        =   15
         Text            =   "Text1"
         ToolTipText     =   "Ingrese cadena a buscar"
         Top             =   1440
         Width           =   5500
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         ToolTipText     =   "Seleccione Producto (F2)"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtBusqueda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2400
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   8880
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   42103
      End
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Movimientos de Salidas por Servicios de Terceros."
         Height          =   255
         Index           =   2
         Left            =   8280
         TabIndex        =   5
         Tag             =   "X3R"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtNomProveedor 
         Height          =   1095
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "frmRptResComprasProveedores.frx":058A
         Top             =   600
         Width           =   6855
      End
      Begin VB.TextBox txtCodProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   6855
      End
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Movimientos de Compra."
         Height          =   255
         Index           =   0
         Left            =   8280
         TabIndex        =   3
         Tag             =   "XC0"
         Top             =   600
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Movimientos de Ingresos por Servicios de Terceros."
         Height          =   255
         Index           =   1
         Left            =   8280
         TabIndex        =   4
         Tag             =   "XST"
         Top             =   840
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   11280
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   42103
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Height          =   255
         Left            =   8280
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Filtrar por Producto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8520
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   10680
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmRptResComprasProveedores.frx":0590
      ToolBars        =   "frmRptResComprasProveedores.frx":90DE
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReporte 
      Height          =   3480
      Left            =   120
      OleObjectBlob   =   "frmRptResComprasProveedores.frx":920F
      TabIndex        =   6
      Top             =   2040
      Width           =   15930
   End
End
Attribute VB_Name = "frmRptResComprasProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub limpiarCajas()
    txtCodProveedor.Text = "*"
    txtNomProveedor.Text = "(*) - Todos los Proveedores"
    
    dtpDesde.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpDesde.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpHasta.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpHasta.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpDesde.Value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtpHasta.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    chkMovimiento(0).Value = vbChecked
    chkMovimiento(1).Value = vbChecked
    chkMovimiento(2).Value = vbChecked
    
    txtCodProducto.Text = vbNullString
    txtProducto.Text = vbNullString
    
    txtBusqueda.Text = vbNullString
End Sub

Private Sub procesarConsulta()
    Dim i As Integer
    Dim strConceptosConsultados As String
    
    strConceptosConsultados = vbNullString
    
    For i = 0 To chkMovimiento.Count - 1
        If CBool(chkMovimiento(i).Value) Then
            If strConceptosConsultados = vbNullString Then
                strConceptosConsultados = chkMovimiento(i).Tag
            Else
                strConceptosConsultados = strConceptosConsultados & "','" & chkMovimiento(i).Tag
            End If
        End If
    Next i
    
    Screen.MousePointer = vbHourglass
    
    objAyudaVale.listarResumenComprasPorProveedores dbgReporte, Nothing, _
                                                    Replace(IIf(Trim(txtCodProveedor.Text) = "*", vbNullString, Trim(txtCodProveedor.Text)), ",", "','"), _
                                                    dtpDesde.Value, _
                                                    dtpHasta.Value, _
                                                    strConceptosConsultados, _
                                                    Trim(txtCodProducto.Text), _
                                                    txtProducto.Text, _
                                                    CBool(chkVerColumnaFechaCompra.Value)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkMovimiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub dbgReporte_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "PORCENTAJEDSCTO"
            FontColor = vbWhite
            Font.Bold = True
            Text = Text & " %"
        Case "CANTIDAD", "COSTOMN", "COSTOME", "SUBTOTALMN", "SUBTOTALME", "IMPUESTOMN", "IMPUESTOME", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
    End Select
End Sub

Private Sub dbgReporte_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "PORCENTAJEDSCTO", "COSTOMN", "COSTOME", "SUBTOTALMN", "SUBTOTALME", "IMPUESTOMN", "IMPUESTOME", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
            FontColor = vbBlue
            Color = vbWhite
            Font.Bold = True
    End Select
End Sub

Private Sub dbgReporte_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "PORCENTAJEDSCTO", "COSTOMN", "COSTOME", "SUBTOTALMN", "SUBTOTALME", "IMPUESTOMN", "IMPUESTOME", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
            FontColor = vbBlue
            Color = vbWhite
            Font.Bold = True
    End Select
End Sub


Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    limpiarCajas
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraReporte.Move 0, 0, Me.ScaleWidth, 1815
    
    dbgReporte.Move 0, fraReporte.Height, Me.ScaleWidth, (Me.ScaleHeight - fraReporte.Height) - 200
End Sub

Private Sub tlbReporte_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Consultar"
            procesarConsulta
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
'            With dbgReporte
'                .Columns.ColumnByFieldName("F2NOMPROV").GroupIndex = -1
'                .Columns.ColumnByFieldName("F1NOMORI").GroupIndex = -1
'                .Columns.ColumnByFieldName("DATOS").GroupIndex = -1
'            End With
            
            With cmdlgReporte
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dbgReporte.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
'            With dbgReporte
'                .Columns.ColumnByFieldName("F2NOMPROV").GroupIndex = 0
'                .Columns.ColumnByFieldName("F1NOMORI").GroupIndex = 1
'                .Columns.ColumnByFieldName("DATOS").GroupIndex = 2
'
'                .m.FullExpand
'            End With
            
            Screen.MousePointer = vbDefault
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub txtCodProveedor_DblClick()
    txtCodProveedor_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Me.MousePointer = vbHourglass
            
            wcodcliprov = vbNullString
            
            With Ayuda_Proveedores
                .Show 1
            End With
            
            If wcodcliprov <> vbNullString Then
                If InStr(1, Trim(txtCodProveedor.Text), wcodcliprov) > 0 Then
                    MsgBox "El Proveedor ya esta seleccionado.", vbInformation, App.ProductName
                Else
                    If Trim(txtCodProveedor.Text) = vbNullString Or Trim(txtCodProveedor.Text) = "*" Then
                        txtCodProveedor.Text = wcodcliprov
                    ElseIf Trim(txtCodProveedor.Text) <> vbNullString Or Trim(txtCodProveedor.Text) <> "*" Then
                        txtCodProveedor.Text = txtCodProveedor.Text & "," & wcodcliprov
                    End If
                    
                    ModUtilitario.validarCodigosConsecutivosTexto txtCodProveedor, _
                                                txtNomProveedor, _
                                                "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", _
                                                vbNullString
                End If
'                txtCodProveedor.Text = wcodcliprov
'                lblProveedor.Caption = wnomcliprov
                
                ModUtilitario.pulsarTecla vbKeyTab
            End If
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            If Trim(txtCodProveedor.Text) <> vbNullString Then
                ModUtilitario.validarCodigosConsecutivosTexto txtCodProveedor, _
                                                txtNomProveedor, _
                                                "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", _
                                                vbNullString
                                                
'                lblProveedor.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtCodProveedor.Text), "T")
'
'                If Trim(lblProveedor.Caption) = vbNullString Then
'                    MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                    ModUtilitario.seleccionarTextoCaja txtCodProveedor
'
'                    Exit Sub
'                End If
            End If

            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
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
                .InsumoOP = False
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
                Else
                    txtCodProducto.Text = vbNullString
                    txtProducto.ToolTipText = "Ingrese cadena a buscar"
                End If
            End With
        Case vbKeyReturn
            'ModUtilitario.pulsarTecla vbKeyTab
            procesarConsulta
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        txtProducto.ToolTipText = txtProducto.Text
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            txtProducto.Text = vbNullString
            txtProducto.ToolTipText = "Ingrese cadena a buscar"
        End If
    Else
        txtProducto.Text = vbNullString
        txtProducto.ToolTipText = "Ingrese cadena a buscar"
    End If
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            procesarConsulta
    End Select
End Sub

