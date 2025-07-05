VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptAtencionOCPorProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atención de O/C's por Proveedor"
   ClientHeight    =   3015
   ClientLeft      =   390
   ClientTop       =   1755
   ClientWidth     =   9480
   Icon            =   "frmRptAtencionOCPorProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9480
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
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      Begin VB.CheckBox chkIncluirOCAtencionTotalYOrdenCerrada 
         Caption         =   "Incluir las O/C's con estados 'Atencion Total' y 'Orden Cerrada':"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Seleccione Producto (F2)"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkMostrarNroRequerimiento 
         Caption         =   "Mostrar el No. de Requerimiento del producto solicitado (Detalle de Compromiso)."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   7695
      End
      Begin VB.CheckBox chkSoloProductoConSaldo 
         Caption         =   "Mostrar solo productos pendientes de atención."
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Value           =   1  'Checked
         Width           =   7695
      End
      Begin VB.TextBox txtCodProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   5760
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   42103
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   7800
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   42103
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblProveedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
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
      ToolsCount      =   9
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmRptAtencionOCPorProv.frx":058A
      ToolBars        =   "frmRptAtencionOCPorProv.frx":8433
   End
End
Attribute VB_Name = "frmRptAtencionOCPorProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub limpiarCajas()
    txtCodProveedor.Text = vbNullString
    lblproveedor.Caption = vbNullString
    
    txtCodProducto.Text = vbNullString
    lblProducto.Caption = vbNullString
    
    chkSoloProductoConSaldo.value = vbChecked
    chkMostrarNroRequerimiento.value = vbUnchecked
    chkIncluirOCAtencionTotalYOrdenCerrada.value = vbUnchecked
    
    dtpDesde.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpDesde.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpHasta.MinDate = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpHasta.MaxDate = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpDesde.value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtpHasta.value = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    dtpDesde.Enabled = False
    dtpHasta.Enabled = False
End Sub

Private Sub chkIncluirOCAtencionTotalYOrdenCerrada_Click()
    If CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value) Then
        chkSoloProductoConSaldo.value = vbUnchecked
        chkSoloProductoConSaldo.Enabled = Not CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
        
        chkMostrarNroRequerimiento.value = vbUnchecked
        chkMostrarNroRequerimiento.Enabled = Not CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
        
        dtpDesde.Enabled = True
        dtpHasta.Enabled = True
    Else
        chkSoloProductoConSaldo.value = vbChecked
        chkSoloProductoConSaldo.Enabled = Not CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
        
        chkMostrarNroRequerimiento.value = vbUnchecked
        chkMostrarNroRequerimiento.Enabled = Not CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
        
        dtpDesde.Enabled = False
        dtpHasta.Enabled = False
    End If
End Sub

Private Sub chkMostrarNroRequerimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub chkSoloProductoConSaldo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    limpiarCajas
End Sub

Private Sub tlbReporte_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "Detalle"
            If Trim(txtCodProveedor.Text) = vbNullString Then
                MsgBox "Seleccione el Proveedor.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If Trim(lblproveedor.Caption) = vbNullString Then
                MsgBox "Verifique el Proveedor, no existe.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Dim rpt1 As New rptOrdenPendienteAtencion
            
            With rpt1
                .CodigoProveedor = Trim(txtCodProveedor.Text)
                .CodigoProducto = Trim(txtCodProducto.Text)
                .MostrarNroRequerimiento = CBool(chkMostrarNroRequerimiento.value)
                .SoloProductoConSaldo = CBool(chkSoloProductoConSaldo.value)
                .IncluirOCAtencionTotalYOrdenCerrada = CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
                .Desde = Format(dtpDesde.value, "Short Date")
                .Hasta = Format(dtpHasta.value, "Short Date")
                
                .Show
            End With
        Case "Consolidado"
            If Trim(txtCodProveedor.Text) = vbNullString Then
                MsgBox "Seleccione el Proveedor.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If Trim(lblproveedor.Caption) = vbNullString Then
                MsgBox "Verifique el Proveedor, no existe.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Dim rpt2 As New rptOrdenPendienteAtencionResumen
            
            With rpt2
                .CodigoProveedor = Trim(txtCodProveedor.Text)
                .CodigoProducto = Trim(txtCodProducto.Text)
                .SoloProductoConSaldo = CBool(chkSoloProductoConSaldo.value)
                .IncluirOCAtencionTotalYOrdenCerrada = CBool(chkIncluirOCAtencionTotalYOrdenCerrada.value)
                .Desde = Format(dtpDesde.value, "Short Date")
                .Hasta = Format(dtpHasta.value, "Short Date")
                
                .Show
            End With
        Case "Salir"
            Unload Me
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
                    lblProducto.Caption = objAyudaBien.Descripcion
                    lblProducto.ToolTipText = lblProducto.Caption
                Else
                    txtCodProducto.Text = vbNullString
                    lblProducto.ToolTipText = vbNullString
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
            'procesarConsulta
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.ToolTipText = lblProducto.Caption
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            lblProducto.Caption = vbNullString
            lblProducto.ToolTipText = vbNullString
        End If
    Else
        lblProducto.Caption = vbNullString
        lblProducto.ToolTipText = vbNullString
    End If
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
                txtCodProveedor.Text = wcodcliprov
                lblproveedor.Caption = wnomcliprov
                
                ModUtilitario.pulsarTecla vbKeyTab
            Else
                MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                ModUtilitario.seleccionarTextoCaja txtCodProveedor
                
                Me.MousePointer = vbDefault
                
                Exit Sub
            End If
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            If Trim(txtCodProveedor.Text) <> vbNullString Then
                lblproveedor.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtCodProveedor.Text), "T")
                
                If Trim(lblproveedor.Caption) = vbNullString Then
                    MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    ModUtilitario.seleccionarTextoCaja txtCodProveedor
                    
                    Exit Sub
                End If
            End If
            
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub
