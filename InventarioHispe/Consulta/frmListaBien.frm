VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaBien 
   Caption         =   "Lista de Productos y Servicios"
   ClientHeight    =   6945
   ClientLeft      =   5445
   ClientTop       =   2190
   ClientWidth     =   13560
   Icon            =   "frmListaBien.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   13560
   Begin VB.CheckBox s 
      Caption         =   "Lista Servicios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8760
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   1920
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   8505
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8280
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   120
         TabIndex        =   5
         Top             =   670
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   25
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   855
      Left            =   8640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgBien 
      Height          =   5865
      Left            =   75
      OleObjectBlob   =   "frmListaBien.frx":058A
      TabIndex        =   1
      Top             =   1080
      Width           =   13365
   End
   Begin ActiveToolBars.SSActiveToolBars tlbBien 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      Tools           =   "frmListaBien.frx":3339
      ToolBars        =   "frmListaBien.frx":E3FB
   End
End
Attribute VB_Name = "frmListaBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda        As Boolean
Private bolInsumoOP     As Boolean
Private bolParaVenta    As Boolean
Private bolTieneMovimientoAlmacen As Boolean
Private strCadenaCorte  As String
Private strFiltroAdicional As String


Private strTipoBienMostrar As String

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let InsumoOP(ByVal Value As Boolean)
    bolInsumoOP = Value
End Property

Public Property Get InsumoOP() As Boolean
    InsumoOP = bolInsumoOP
End Property

Public Property Let ParaVenta(ByVal Value As Boolean)
    bolParaVenta = Value
End Property

Public Property Get ParaVenta() As Boolean
    ParaVenta = bolParaVenta
End Property

Public Property Let TieneMovimientoAlmacen(ByVal Value As Boolean)
    bolTieneMovimientoAlmacen = Value
End Property

Public Property Get TieneMovimientoAlmacen() As Boolean
    TieneMovimientoAlmacen = bolTieneMovimientoAlmacen
End Property


'Propiedad Cadena de Corte de Informacion
Public Property Let CadenaCorte(ByVal Value As String)
    strCadenaCorte = Value
End Property

Public Property Get CadenaCorte() As String
    CadenaCorte = strCadenaCorte
End Property

'Propiedad Filtro Adicional
Public Property Let FiltroAdicional(ByVal Value As String)
    strFiltroAdicional = Value
End Property

Public Property Get FiltroAdicional() As String
    FiltroAdicional = strFiltroAdicional
End Property

'Propiedad Tipo de Bien as Mostrar en la Ayuda
Public Property Let TipoBienMostrar(ByVal Value As String)
    strTipoBienMostrar = Value
End Property

Public Property Get TipoBienMostrar() As String
    TipoBienMostrar = strTipoBienMostrar
End Property



Public Sub listarBien()
    Screen.MousePointer = vbHourglass
    
    dbgBien.Dataset.Close
    
    
        objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, bolTieneMovimientoAlmacen
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub chkServicios_Click()

    objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, 0
End Sub


Private Sub dbgBien_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaBien.Codigo = Trim(dbgBien.Columns.ColumnByFieldName("F5CODPRO").Value & "")
        objAyudaBien.Descripcion = Trim(dbgBien.Columns.ColumnByFieldName("F5NOMPRO").Value & "")
        
        Me.Hide
    Else
        If ModUtilitario.validarFormAbierto("frmMantBien") Then
            Unload frmMantBien
        End If
        
        With frmMantBien
            .Ayuda = bolAyuda
            .Codigo = Trim(dbgBien.Columns.ColumnByFieldName("F5CODPRO").Value & "")
            
            .Show 1
            
            listarBien
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgBien_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgBien_OnDblClick
        Case vbKeyUp
            If dbgBien.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    Else
        Me.top = (Screen.Height / 2) - (Me.Height / 2)
        Me.left = (Screen.Width / 2) - (Me.Width / 2)
    End If
    
    txtBusqueda.Text = strCadenaCorte
    
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
    
    listarBien
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgBien.Move 0, fraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height)
    
    fraBusqueda.left = 0
    fraBusqueda.top = 0
    
    fraProceso.left = fraBusqueda.Width + 100
    fraProceso.top = fraBusqueda.top
    fraProceso.Width = (dbgBien.Width - fraBusqueda.Width) - 100
    fraProceso.Height = fraBusqueda.Height
    pgbProceso.Width = fraProceso.Width - 1000
    pgbProceso.left = (fraProceso.Width - pgbProceso.Width) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
    
    dbgBien.Dataset.Close
End Sub

Private Sub s_Click()
   If s.Value = False Then
        objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, True, True
    Else
        objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, False, False
    End If
'    listarBien
End Sub

Private Sub s_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If s.Value = False Then
'        objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, False, True
'    Else
'        objAyudaBien.listarGrillaBien dbgBien, txtBusqueda.Text, bolAyuda, bolInsumoOP, False, False
'    End If
End Sub


Private Sub s_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If s.Value = True Then
'        objAyudaBien.listarGrillaBien dbgBien, txtbusqueda.Text, bolAyuda, bolInsumoOP, False, True
'    Else
'        objAyudaBien.listarGrillaBien dbgBien, txtbusqueda.Text, bolAyuda, bolInsumoOP, True, False
'    End If
End Sub


Private Sub tlbBien_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            With frmMantBien
                .Ayuda = bolAyuda
                .Codigo = vbNullString
                
                .Show 1
                
                If Not bolAyuda Then
                    listarBien
                Else
                    Unload Me
                End If
            End With
        Case "Imprimir"
                objAyudaBien.listarGrillaBien Nothing, txtBusqueda.Text, bolAyuda, bolInsumoOP
                
                With rptListaBien
                    .dtcBien.ConnectionString = StrConexDbBancos
                    .dtcBien.Source = objAyudaBien.SQLSelectAlter
                    
                    .fldFecha.Text = Format(Date, "DD/MM/YYYY")
                    .lblempresa.Caption = wnomcia
                    .Caption = "Relación de Productos"
                    
                    .Show 1
                End With
        Case "Filtrar"
            dbgBien.Filter.FilterActive = CBool(Tool.State)
        Case "Agrupar"
            If Tool.State = ssChecked Then
                dbgBien.Options.Set (egoShowGroupPanel)
            Else
                dbgBien.Options.Unset (egoShowGroupPanel)
            End If
        Case "Importar"
            If MsgBox("¿Dese ejecutar el proceso de Importación de Codigos PT?" & vbNewLine & vbNewLine & _
                        "IMPORTANTE: Recuerde que este proceso puede tomar varios minutos.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
            
            dbgBien.Dataset.Close
            
            
            listarBien
        Case "VerificarCtasContables"
            If MsgBox("¿Desea verificar y actualizar las Cuentas Contables de Productos Importados?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            
                dbgBien.Dataset.Close
                
                abrirCnnDbBancos
                
                objAyudaBien.actualizarCtasContablesLogistica fraProceso, pgbProceso
                
                listarBien
            End If
        Case "Salir"
            objAyudaBienColor.inicializarEntidades
            
            Unload Me
    End Select
End Sub

'Private Sub timTemporizador_Timer()
'    If timTemporizador.Interval = 25 Then
'        listarBien
'
'        timTemporizador.Enabled = False
'        pgbProgresoBusqueda.value = 0
'        pgbProgresoBusqueda.Visible = False
'    Else
'        timTemporizador.Interval = timTemporizador.Interval + 1
'        pgbProgresoBusqueda.value = timTemporizador.Interval
'    End If
'End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            listarBien
        Case vbKeyDown
            dbgBien.SetFocus
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

