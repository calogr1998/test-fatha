VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptResCompras 
   Caption         =   "Resumen de Compras"
   ClientHeight    =   5595
   ClientLeft      =   210
   ClientTop       =   1785
   ClientWidth     =   16905
   Icon            =   "frmRptResCompras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   16905
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   " Filtrar Por "
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
      Left            =   10200
      TabIndex        =   11
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtNomFiltro 
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "frmRptResCompras.frx":058A
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txtFiltro 
         Height          =   285
         Left            =   1680
         MaxLength       =   255
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   120
         Width           =   4695
      End
      Begin VB.CheckBox chkDetalle 
         Caption         =   "Ver Detalle"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "< Etiqueta >"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraAgruparPor 
      Caption         =   " Agrupar por "
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
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton optAgruparPor 
         Caption         =   "Ninguno"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optAgruparPor 
         Caption         =   "Sub-Familia"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optAgruparPor 
         Caption         =   "Familia"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optAgruparPor 
         Caption         =   "Tipo de Existencia"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cmdlgReporte 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   " Periodo a Consultar "
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   42103
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   42103
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   615
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
      Tools           =   "frmRptResCompras.frx":0590
      ToolBars        =   "frmRptResCompras.frx":90DE
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReporte 
      Height          =   4440
      Left            =   120
      OleObjectBlob   =   "frmRptResCompras.frx":920F
      TabIndex        =   2
      Top             =   1080
      Width           =   16650
   End
End
Attribute VB_Name = "frmRptResCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intOpcionResumen As Integer

Private Sub limpiarCajas()
    dtpDesde.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpDesde.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpHasta.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpHasta.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpDesde.Value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtpHasta.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    optAgruparPor(0).Value = False
    optAgruparPor(1).Value = False
    optAgruparPor(2).Value = False
    optAgruparPor(3).Value = True
End Sub

Private Sub procesarConsulta()
    Screen.MousePointer = vbHourglass
    
    'objAyudaVale.listarResumenCompras dbgReporte, _
                                        dtpDesde.value, _
                                        dtpHasta.value, _
                                        intOpcionResumen, _
                                        vbNullString
    
    objAyudaVale.listarResumenComprasV2 dbgReporte, _
                                        dtpDesde.Value, _
                                        dtpHasta.Value, _
                                        intOpcionResumen, _
                                        IIf(CBool(chkDetalle.Value), 1, 0), _
                                        Replace(IIf(Trim(txtFiltro.Text) = "*", vbNullString, "'" & Trim(txtFiltro.Text) & "'"), ",", "','"), _
                                        vbNullString
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub dbgReporte_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
    End Select
End Sub

Private Sub dbgReporte_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTIDAD", "TOTALMN", "TOTALME"
            Text = Format(Text, "#,0.00")
            FontColor = vbBlue
            Color = vbWhite
            Font.Bold = True
    End Select
End Sub

'Private Sub dbgReporte_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
'    Select Case UCase(Column.FieldName)
'        Case "CANTIDAD", "PORCENTAJEDSCTO", "COSTOMN", "COSTOME", "SUBTOTALMN", "SUBTOTALME", "IMPUESTOMN", "IMPUESTOME", "TOTALMN", "TOTALME"
'            Text = Format(Text, "#,0.00")
'            FontColor = vbBlue
'            Color = vbWhite
'            Font.Bold = True
'    End Select
'End Sub

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
    
    'fraPeriodo.Move 0, 0, Me.ScaleWidth, 1000
    
    dbgReporte.Move 0, fraPeriodo.Height + 200, Me.ScaleWidth, (Me.ScaleHeight - fraPeriodo.Height) - 200
End Sub

Private Sub optAgruparPor_Click(Index As Integer)
    intOpcionResumen = Index
    
    Select Case Index
        Case 3
            lblEtiqueta.Caption = "Producto :"
            
            chkDetalle.Value = vbUnchecked
            chkDetalle.Enabled = False
        Case Else
            lblEtiqueta.Caption = optAgruparPor(Index).Caption & " :"
            
            chkDetalle.Value = vbChecked
            chkDetalle.Enabled = True
    End Select
    
    Select Case intOpcionResumen
        Case 0
            txtFiltro.Text = "*"
            txtNomFiltro.Text = "(*) - Todos los Tipos de Existencias."
        Case 1
            txtFiltro.Text = "*"
            txtNomFiltro.Text = "(*) - Todos las Familias."
        Case 2
            txtFiltro.Text = "*"
            txtNomFiltro.Text = "(*) - Todos las Sub-Familias."
        Case 3
            txtFiltro.Text = "*"
            txtNomFiltro.Text = "(*) - Todos los Productos."
    End Select
End Sub

Private Sub tlbReporte_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Consultar"
            procesarConsulta
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
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
            
            Screen.MousePointer = vbDefault
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub txtFiltro_DblClick()
    txtFiltro_KeyDown vbKeyF2, 0
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Me.MousePointer = vbHourglass
            
            Select Case intOpcionResumen
                Case 0 'Tipo de Existencia
                    objAyudaTipoExistencia.inicializarEntidades
                    
                    With frmListaTipoExistencia
                        .Ayuda = True
                        
                        .Show 1
                    End With
                    
                    If objAyudaTipoExistencia.Codigo <> vbNullString Then
                        If InStr(1, Trim(txtFiltro.Text), objAyudaTipoExistencia.Codigo) > 0 Then
                            MsgBox "Tipo de Existencia ya esta seleccionado.", vbInformation, App.ProductName
                        Else
                            If Trim(txtFiltro.Text) = vbNullString Or Trim(txtFiltro.Text) = "*" Then
                                txtFiltro.Text = objAyudaTipoExistencia.Codigo
                            ElseIf Trim(txtFiltro.Text) <> vbNullString Or Trim(txtFiltro.Text) <> "*" Then
                                txtFiltro.Text = txtFiltro.Text & "," & objAyudaTipoExistencia.Codigo
                            End If
                            
                            ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "DESCRIPCION", "EF2TIPOEXISTENCIA", "CODIGO", _
                                                        vbNullString
                        End If
                        
                        ModUtilitario.pulsarTecla vbKeyTab
                    End If
                Case 1 'Familia
                    objAyudaFamilia.inicializarEntidades
                    
                    With frmListaFamilia
                        .Ayuda = True
                        
                        .Show 1
                    End With
                    
                    If objAyudaFamilia.Codigo <> vbNullString Then
                        If InStr(1, Trim(txtFiltro.Text), objAyudaFamilia.Codigo) > 0 Then
                            MsgBox "Familia ya esta seleccionada.", vbInformation, App.ProductName
                        Else
                            If Trim(txtFiltro.Text) = vbNullString Or Trim(txtFiltro.Text) = "*" Then
                                txtFiltro.Text = objAyudaFamilia.Codigo
                            ElseIf Trim(txtFiltro.Text) <> vbNullString Or Trim(txtFiltro.Text) <> "*" Then
                                txtFiltro.Text = txtFiltro.Text & "," & objAyudaFamilia.Codigo
                            End If
                            
                            ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F7DESCON", "SF7NIVEL01", "F7CODCON", _
                                                        vbNullString
                        End If
                        
                        ModUtilitario.pulsarTecla vbKeyTab
                    End If
                Case 2 'Sub-Familia
                    objAyudaSubFamilia.inicializarEntidades
                    
                    With frmListaSubFamilia
                        .Ayuda = True
                        
                        .Show 1
                    End With
                    
                    If objAyudaSubFamilia.Codigo <> vbNullString Then
                        If InStr(1, Trim(txtFiltro.Text), objAyudaSubFamilia.Codigo) > 0 Then
                            MsgBox "Sub-Familia ya esta seleccionada.", vbInformation, App.ProductName
                        Else
                            If Trim(txtFiltro.Text) = vbNullString Or Trim(txtFiltro.Text) = "*" Then
                                txtFiltro.Text = objAyudaSubFamilia.Codigo
                            ElseIf Trim(txtFiltro.Text) <> vbNullString Or Trim(txtFiltro.Text) <> "*" Then
                                txtFiltro.Text = txtFiltro.Text & "," & objAyudaSubFamilia.Codigo
                            End If
                            
                            ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F7DESCON", "SF7NIVEL02", "F7CODCON", _
                                                        vbNullString
                        End If
                        
                        ModUtilitario.pulsarTecla vbKeyTab
                    End If
                Case 3 'Producto
                    objAyudaBien.inicializarEntidades
                    
                    With frmListaBien
                        '.Ayuda = True
                        '.TieneMovimientoAlmacen = True
                        '.SoloServicios = False
                        '.InsumoOP = False
                        '.CadenaCorte = vbNullString
                        
                        
                        .Ayuda = True
                        .InsumoOP = False
                        .ParaVenta = False
                        .TieneMovimientoAlmacen = True
                        .CadenaCorte = vbNullString
                        .FiltroAdicional = vbNullString
                        .TipoBienMostrar = "P"
                        
                        .Show 1
                    End With
                    
                    If objAyudaBien.Codigo <> vbNullString Then
                        If InStr(1, Trim(txtFiltro.Text), objAyudaBien.Codigo) > 0 Then
                            MsgBox "Sub-Familia ya esta seleccionada.", vbInformation, App.ProductName
                        Else
                            If Trim(txtFiltro.Text) = vbNullString Or Trim(txtFiltro.Text) = "*" Then
                                txtFiltro.Text = objAyudaBien.Codigo
                            ElseIf Trim(txtFiltro.Text) <> vbNullString Or Trim(txtFiltro.Text) <> "*" Then
                                txtFiltro.Text = txtFiltro.Text & "," & objAyudaBien.Codigo
                            End If
                            
                            ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F5NOMPRO", "IF5PLA", "F5CODPRO", _
                                                        vbNullString
                        End If
                        
                        ModUtilitario.pulsarTecla vbKeyTab
                    End If
            End Select
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            If Trim(txtFiltro.Text) <> vbNullString Then
                Select Case intOpcionResumen
                    Case 0 'Tipo de Existencia
                        ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "DESCRIPCION", "EF2TIPOEXISTENCIA", "CODIGO", _
                                                        vbNullString
                    Case 1 'Familia
                        ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F7DESCON", "SF7NIVEL01", "F7CODCON", _
                                                        vbNullString
                    Case 2 'Sub-Familia
                        ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F7DESCON", "SF7NIVEL02", "F7CODCON", _
                                                        vbNullString
                    Case 3 'Producto
                        ModUtilitario.validarCodigosConsecutivosTexto txtFiltro, _
                                                        txtNomFiltro, _
                                                        "F5NOMPRO", "IF5PLA", "F5CODPRO", _
                                                        vbNullString
                End Select
            End If
            
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub
