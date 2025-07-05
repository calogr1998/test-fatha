VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vale_salida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vale de Salida de Almacen"
   ClientHeight    =   8295
   ClientLeft      =   135
   ClientTop       =   1710
   ClientWidth     =   15735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "vale_salida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   15735
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   12240
      TabIndex        =   56
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      Begin MSComctlLib.ProgressBar pgbProceso 
         Height          =   200
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraOrdenProduccion 
      Caption         =   " O. Producción "
      Height          =   1815
      Left            =   12240
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtNroOrdenProduccion 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   720
         MaxLength       =   30
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbCategoriaTipo 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Text            =   "cmbCategoriaTipo"
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtIDOrdenProduccion 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   53
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblIdCategoriaTipo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID CategoriaTipo"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "O.P."
         Height          =   210
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ID O.P."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   510
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   465
      Left            =   9720
      TabIndex        =   32
      Top             =   120
      Width           =   5835
      _Version        =   65536
      _ExtentX        =   10292
      _ExtentY        =   820
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtdestino 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   48
         Top             =   90
         Width           =   1725
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   795
         TabIndex        =   14
         Top             =   90
         Width           =   1485
      End
      Begin VB.Label lblNumeroValeExterno 
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
         Left            =   2280
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label LblDestino 
         AutoSize        =   -1  'True
         Caption         =   "Vale.Dest."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2880
         TabIndex        =   49
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Vale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   135
         Width           =   570
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   465
      Left            =   180
      TabIndex        =   30
      Top             =   180
      Width           =   9405
      _Version        =   65536
      _ExtentX        =   16589
      _ExtentY        =   820
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   12
         ShowShortcutsInToolTips=   -1  'True
         Tools           =   "vale_salida.frx":058A
         ToolBars        =   "vale_salida.frx":AA24
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2880
      Left            =   120
      TabIndex        =   29
      Top             =   615
      Width           =   12105
      _Version        =   65536
      _ExtentX        =   21352
      _ExtentY        =   5080
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbalmacen 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         ItemData        =   "vale_salida.frx":AC69
         Left            =   1440
         List            =   "vale_salida.frx":AC6B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3300
      End
      Begin VB.TextBox txtproveedor 
         Height          =   315
         Left            =   3600
         MaxLength       =   11
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtnomprov 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   58
         Top             =   1080
         Width           =   5775
      End
      Begin VB.ComboBox cmbTipoAuxiliar 
         Height          =   330
         ItemData        =   "vale_salida.frx":AC6D
         Left            =   1440
         List            =   "vale_salida.frx":AC77
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2115
      End
      Begin VB.CheckBox chkExportarVale 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportar Vale"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   51
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbconcepto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "vale_salida.frx":AD5D
         Left            =   1440
         List            =   "vale_salida.frx":AD5F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3300
      End
      Begin VB.TextBox txtnumdoc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtserie 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtnumorden 
         Height          =   315
         Left            =   12240
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtusuario 
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   3390
         Width           =   960
      End
      Begin VB.TextBox txtalmacendes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbalmacendes 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtobserva 
         Height          =   675
         Left            =   8040
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2070
         Width           =   3810
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2400
         Width           =   2925
      End
      Begin VB.TextBox txtserfac 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2400
         Width           =   465
      End
      Begin VB.TextBox txtnumfac 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   9
         Top             =   2400
         Width           =   1410
      End
      Begin VB.ComboBox cmbmoneda 
         Height          =   330
         ItemData        =   "vale_salida.frx":AD61
         Left            =   8370
         List            =   "vale_salida.frx":AD6B
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3915
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtccosto 
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox txttc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   8415
         MaxLength       =   11
         TabIndex        =   26
         Text            =   "1.0"
         Top             =   3645
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtconcepto 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin Threed.SSPanel pnlalmacen 
         Height          =   315
         Left            =   5265
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.TextBox txtalmacen 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "01"
         Top             =   225
         Width           =   495
      End
      Begin Threed.SSPanel pnlconcepto 
         Height          =   315
         Left            =   5265
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlccosto 
         Height          =   315
         Left            =   2880
         TabIndex        =   24
         Top             =   1440
         Width           =   7380
         _Version        =   65536
         _ExtentX        =   13017
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlsolicitante 
         Height          =   315
         Left            =   2295
         TabIndex        =   23
         Top             =   3390
         Width           =   3300
         _Version        =   65536
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlorden 
         Height          =   315
         Left            =   12240
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   180
         _Version        =   65536
         _ExtentX        =   317
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin MSComCtl2.DTPicker abofecha 
         Height          =   315
         Left            =   10560
         TabIndex        =   28
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117112833
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker dtpFechaDoc 
         Height          =   300
         Left            =   6360
         TabIndex        =   60
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117112833
         CurrentDate     =   40611
      End
      Begin VB.Label Label23 
         Caption         =   "Fecha Documento"
         Height          =   195
         Left            =   6360
         TabIndex        =   61
         Top             =   2160
         Width           =   1635
      End
      Begin VB.Label lblFechaMensaje 
         Caption         =   "Uso Fecha Predeterminada."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9480
         TabIndex        =   59
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Guía de Remisión"
         Height          =   210
         Left            =   45
         TabIndex        =   47
         Top             =   2025
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   210
         Left            =   1440
         TabIndex        =   46
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   210
         Left            =   2280
         TabIndex        =   45
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Partida / Control"
         Height          =   210
         Left            =   12240
         TabIndex        =   44
         Top             =   2520
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Alm. Destino"
         Height          =   210
         Left            =   5640
         TabIndex        =   42
         Top             =   675
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   210
         Left            =   8040
         TabIndex        =   41
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label lbldocum 
         AutoSize        =   -1  'True
         Caption         =   "Serie/Número"
         Height          =   210
         Left            =   4560
         TabIndex        =   40
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   210
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   210
         Left            =   7785
         TabIndex        =   38
         Top             =   3960
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T/C"
         Height          =   210
         Left            =   8010
         TabIndex        =   36
         Top             =   3690
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   9480
         TabIndex        =   35
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   675
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Almacén"
         Height          =   195
         Left            =   135
         TabIndex        =   31
         Top             =   315
         Width           =   645
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4545
      Left            =   120
      OleObjectBlob   =   "vale_salida.frx":AD7F
      TabIndex        =   11
      Top             =   3600
      Width           =   15555
   End
End
Attribute VB_Name = "vale_salida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw_activate         As Boolean
Dim cconex_form         As String
Dim wpartida            As String
Dim sw_ayuda            As Boolean
Dim sw_nuevo_item       As Boolean
Dim cnombase            As String
Dim cnomtabla           As String
Dim cnomtabla1          As String
Dim Values()            As Variant
'Dim amovs_cab(0 To 18)  As a_grabacion
Dim amovs_cab(0 To 22)  As a_grabacion
'Dim amovs_det(0 To 13)  As a_grabacion
Dim amovs_det(0 To 16)  As a_grabacion
Dim ctipo               As String * 1
Dim cvalores            As String
Dim cmes                As String * 2
Dim RSDETALLE           As New ADODB.Recordset
Dim nfil                As Integer
Dim sw_cabecera         As Boolean
Dim sw_detalle          As Boolean
Dim sw_ayuda_prod       As Boolean
Dim sw_habil            As Boolean
Dim wprecio             As String
Dim sw_ingreso          As Boolean


Rem SK ADD:
Private bolAyuda            As Boolean
Private strCodAlmacen       As String
Private strNumeroVale       As String

Private strFichero          As String
Private bolObviarCierre     As Boolean

Private objVale             As ClsVale



Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let CodigoAlmacen(ByVal value As String)
    strCodAlmacen = value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property

Public Property Let NumeroVale(ByVal value As String)
    strNumeroVale = value
End Property

Public Property Get NumeroVale() As String
    NumeroVale = strNumeroVale
End Property



Private Sub listarGrilla()
    abrirCnTemporal
    
    With dxDBGrid1.Dataset
        .Active = False
        .ADODataset.ConnectionString = cnDBTemp
        .ADODataset.CommandText = "SELECT * FROM TMPVALESALIDA ORDER BY ITEM, DESCRIPCION"
        .Active = True
        
        dxDBGrid1.KeyField = "ITEM"
        
        .Close
        .Open
    End With
End Sub
Rem SK ADD:----------------------------------------------------------------------------------------------------------
Private Sub copiarSeleccionAyudaProductos()
    Dim rstProducto As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "COD_SOLICITUD, F5CODPRO, F5CODFAB, F5NOMPRO, F5NOMPRO_ING, F7CODMED, "
    CadSql = CadSql & "IIF(TRIM(F5AFECTO & '') = '*', TRUE, FALSE) AS AFECTO, "
    CadSql = CadSql & "F2MONEDA, F5VTANET, F5VTANETDOL, F5FOB, F5FOBMAX "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPPRODUCTOS "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "F4PERINT = -1"
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    rstProducto.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProducto.EOF Then
        rstProducto.MoveFirst
        
        Do While Not rstProducto.EOF
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALESALIDA", "CODPROD", Trim(rstProducto!f5codpro & ""), "T", "AND COD_SOLICITUD = '" & Trim(rstProducto!COD_SOLICITUD & "") & "' AND TRIM(F4NUMORD & '') = ''") = vbNullString Then
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALESALIDA(ITEM, COD_SOLICITUD, "
                CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, "
                CadSql = CadSql & "CANTIDAD, CANTIDADMAX) "
                CadSql = CadSql & "VALUES(" & dblItem & ", "
                CadSql = CadSql & "'" & Trim(rstProducto!COD_SOLICITUD & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!f5codpro & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!f5codpro & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!F5NOMPRO & "") & "', "
                CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", Trim(rstProducto!f7codmed & ""), "T") & "', "
                CadSql = CadSql & Val(rstProducto!F5FOB & "") & ", "
                CadSql = CadSql & "0)"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALESALIDA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CANTIDAD = " & Val(rstProducto!F5FOB & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = 0 "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstProducto!COD_SOLICITUD & "") & "' AND "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!f5codpro & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProducto.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    Set rstProducto = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionDevolucionOC()
    Dim rstProducto As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILDEVOLUCIONOC "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE"
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    rstProducto.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProducto.EOF Then
        rstProducto.MoveFirst
        
        CadSql = vbNullString
        CadSql = CadSql & "DELETE FROM TMPVALESALIDA "
        CadSql = CadSql & "WHERE "
        CadSql = CadSql & "F4NUMORD NOT IN (SELECT NROOC FROM TMPUTILDEVOLUCIONOC WHERE CODPROVEEDOR = '" & Trim(rstProducto!CodProveedor & "") & "' GROUP BY NROOC)"
        
        cnDBTemp.Execute CadSql
        
        Do While Not rstProducto.EOF
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALESALIDA", "CODPROD", Trim(rstProducto!CodProducto & ""), "T", _
                                            "AND COD_SOLICITUD = '" & Trim(rstProducto!NroPedido & "") & "' " & _
                                            "AND F4NUMORD = '" & Trim(rstProducto!NROOC & "") & "'") = vbNullString Then
                         
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                If dblItem = 1 Then
                    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') = ''"
                End If
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALESALIDA("
                CadSql = CadSql & "ITEM, F4NUMORD, COD_SOLICITUD, "
                CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, "
                CadSql = CadSql & "UMEDIDA, CANTIDAD, CANTIDADMAX) "
                CadSql = CadSql & "VALUES("
                CadSql = CadSql & dblItem & ", "
                CadSql = CadSql & "'" & Trim(rstProducto!NROOC & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!NroPedido & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!CodProducto & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!NOMPRODUCTO & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!um & "") & "', "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALESALIDA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CANTIDAD = " & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = " & Val(rstProducto!CANTIDADDESTINO & "") & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "F4NUMORD = '" & Trim(rstProducto!NROOC & "") & "' AND "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstProducto!NroPedido & "") & "' AND "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!CodProducto & "") & "' AND "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProducto.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    Set rstProducto = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionStockDisponible()
    Dim rstProducto As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILSTOCKDISPONIBLE "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE"
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    rstProducto.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProducto.EOF Then
        rstProducto.MoveFirst
        
'        CadSql = vbNullString
'        CadSql = CadSql & "DELETE FROM TMPVALESALIDA "
'        CadSql = CadSql & "WHERE "
'        CadSql = CadSql & "F4NUMORD NOT IN (SELECT NROOC FROM TMPUTILSTOCKDISPONIBLE WHERE CODPROVEEDOR = '" & Trim(rstProducto!CodProveedor & "") & "' GROUP BY NROOC)"
'
'        cnDBTemp.Execute CadSql
        
        Do While Not rstProducto.EOF
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALESALIDA", "CODPROD", Trim(rstProducto!CodProducto & ""), "T", _
                                            "AND COD_SOLICITUD = '" & Trim(rstProducto!NroPedido & "") & "' " & _
                                            "AND F4NUMORD = '" & Trim(rstProducto!NROOC & "") & "'") = vbNullString Then
                
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALESALIDA("
                CadSql = CadSql & "ITEM, F4NUMORD, COD_SOLICITUD, "
                CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, "
                CadSql = CadSql & "UMEDIDA, CANTIDAD, CANTIDADMAX) "
                CadSql = CadSql & "VALUES("
                CadSql = CadSql & dblItem & ", "
                CadSql = CadSql & "'" & Trim(rstProducto!NROOC & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!NroPedido & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!CodProducto & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & IIf(Trim(rstProducto!NOMPRODUCTO & "") <> vbNullString, "'" & Trim(rstProducto!NOMPRODUCTO & "") & "'", "NULL") & ", "
                CadSql = CadSql & IIf(Trim(rstProducto!um & "") <> vbNullString, "'" & Trim(rstProducto!um & "") & "'", "NULL") & ", "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALESALIDA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!CodProducto & "") & "', "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & "DESCRIPCION = " & IIf(Trim(rstProducto!NOMPRODUCTO & "") <> vbNullString, "'" & Trim(rstProducto!NOMPRODUCTO & "") & "'", "NULL") & ", "
                CadSql = CadSql & "UMEDIDA = " & IIf(Trim(rstProducto!um & "") <> vbNullString, "'" & Trim(rstProducto!um & "") & "'", "NULL") & ", "
                CadSql = CadSql & "CANTIDAD = " & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = " & Val(rstProducto!CANTIDADDESTINO & "") & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "F4NUMORD = '" & Trim(rstProducto!NROOC & "") & "' AND "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstProducto!NroPedido & "") & "' AND "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!CodProducto & "") & "' AND "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProducto.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    Set rstProducto = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionStockDisponibleSql()
    Dim rstProducto As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
    
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') = ''"
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPCPSTOCKDISPONIBLE" & UCase(wusuario) & " "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = 1"
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    rstProducto.Open CadSql, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProducto.EOF Then
        rstProducto.MoveFirst
        
'        CadSql = vbNullString
'        CadSql = CadSql & "DELETE FROM TMPVALESALIDA "
'        CadSql = CadSql & "WHERE "
'        CadSql = CadSql & "F4NUMORD NOT IN (SELECT NROOC FROM TMPUTILSTOCKDISPONIBLE WHERE CODPROVEEDOR = '" & Trim(rstProducto!CodProveedor & "") & "' GROUP BY NROOC)"
'
'        cnDBTemp.Execute CadSql
        
        Do While Not rstProducto.EOF
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALESALIDA", "CODPROD", Trim(rstProducto!CodProducto & ""), "T") = vbNullString Then
                
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALESALIDA("
                CadSql = CadSql & "ITEM, "
                CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, "
                CadSql = CadSql & "UMEDIDA, CANTIDAD, CANTIDADMAX) "
                CadSql = CadSql & "VALUES("
                CadSql = CadSql & dblItem & ", "
                CadSql = CadSql & "'" & Trim(rstProducto!CodProducto & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & IIf(Trim(rstProducto!NOMPRODUCTO & "") <> vbNullString, "'" & Trim(rstProducto!NOMPRODUCTO & "") & "'", "NULL") & ", "
                CadSql = CadSql & IIf(Trim(rstProducto!um & "") <> vbNullString, "'" & Trim(rstProducto!um & "") & "'", "NULL") & ", "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & Val(rstProducto!CANTIDADDESTINO & "") & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALESALIDA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!CodProducto & "") & "', "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & "DESCRIPCION = " & IIf(Trim(rstProducto!NOMPRODUCTO & "") <> vbNullString, "'" & Trim(rstProducto!NOMPRODUCTO & "") & "'", "NULL") & ", "
                CadSql = CadSql & "UMEDIDA = " & IIf(Trim(rstProducto!um & "") <> vbNullString, "'" & Trim(rstProducto!um & "") & "'", "NULL") & ", "
                CadSql = CadSql & "CANTIDAD = " & Val(rstProducto!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = " & Val(rstProducto!CANTIDADDESTINO & "") & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProducto!CodProducto & "") & "' AND "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstProducto!CODPRODUCTOORIGINAL & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProducto.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    Set rstProducto = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub verificarAtencionOrden(ByVal strCodAlmacen As String, _
                                    ByVal strNumeroVale As String, _
                                    Optional ByVal bolValeEliminado As Boolean)
                                        
    Dim rstValeDet As New ADODB.Recordset
    Dim dblCantidadOP As Double
    
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    If Not bolValeEliminado Then
        rstValeDet.Open "SELECT F4NUMORD FROM IF3VALES WHERE F2CODALM = '" & strCodAlmacen & "' AND F4NUMVAL = '" & strNumeroVale & "' GROUP BY F4NUMORD", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    Else
        rstValeDet.Open "SELECT F4NUMORD FROM TMPVALESALIDA GROUP BY F4NUMORD", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not rstValeDet.EOF Then
        rstValeDet.MoveFirst
        
        Do While Not rstValeDet.EOF
            With objAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = "OC"
                .NumeroOrden = Trim(rstValeDet!F4NUMORD & "")
                
                If .obtenerOrden Then
                    If .Estado <> 7 And .Estado <> 8 Then
                        .atencionOrden
                    End If
                End If
                
                .inicializarEntidades
            End With
            
            rstValeDet.MoveNext
        Loop
    End If
End Sub
'-----------------------------------------------------------------------------------------------------


Private Sub IMPRIMIR_VALES(Tipo As Integer)
Dim csql            As String
Dim CSQL1           As String
Dim Csql2           As String
Dim Csql3           As String
Dim prov            As String
Dim costo           As String
Dim RegEmpresa      As New ADODB.Recordset
Dim RegCosto        As New ADODB.Recordset
Dim ccod_almacen    As String
Dim cnum_vale       As String
Dim ctipo_vale      As String * 1

    ccod_almacen = Trim(txtalmacen.Text)
    cnum_vale = Trim(txtnumero.Text)
    costo = Trim(txtccosto.Text)
    If Tipo = 1 Then
        With acr_vales
            .DataControl1.ConnectionString = cconex_dbbancos
            ctipo_vale = "S"
            .Lbl_vale.Caption = " VALE DE SALIDA "
            .lblprov.Visible = False
            .lblpunto.Visible = False
            .fldprov.Visible = False
            If wprecio = "1" Then
                csql = "SELECT DISTINCTROW A.F2CODALM, A.F4NUMVAL, A.F1CODORI, D.F1NOMORI, A.F2CODPROV, A.F4NUMDOC, A.F4OBSERVA, A.F4CENTRO, A.F4FECVAL, B.F5CODPRO  AS CODIGO, B.F3CANPRO, C.F5NOMPRO, C.F5CODFAB, C.F7CODMED, C.F5CODPRO, B.F3PUNIT AS F5PREVTA, B.F3PUNIT*B.F3CANPRO AS PRECIO " & _
                       "FROM ((IF4VALES AS A INNER JOIN IF3VALES AS B ON (A.F2CODALM = B.F2CODALM) AND (A.F4NUMVAL = B.F4NUMVAL)) INNER JOIN IF5PLA AS C ON B.F5CODPRO = C.F5CODPRO) INNER JOIN SF1ORIGENES AS D ON A.F1CODORI = D.F1CODORI " & _
                       "WHERE (((A.F2CODALM)='" & ccod_almacen & "') AND ((A.F4NUMVAL)='" & cnum_vale & "') AND ((A.F1CODORI)=[D].[F1CODORI]) AND ((B.F5CODPRO)=[c].[F5CODPRO]) AND ((C.F5CODPRO)=[B].[F5CODPRO]) ) " & _
                       "ORDER BY A.F4NUMVAL, B.F5CODPRO, C.F5CODFAB;"
            Else
                csql = "SELECT DISTINCTROW A.F2CODALM, A.F4NUMVAL, A.F1CODORI, D.F1NOMORI, A.F2CODPROV, A.F4NUMDOC, A.F4OBSERVA, A.F4CENTRO, A.F4FECVAL, B.F5CODPRO  AS CODIGO, B.F3CANPRO, C.F5NOMPRO, C.F5CODFAB, C.F7CODMED, C.F5CODPRO, C.F5PREVTA, C.F5PREVTA*B.F3CANPRO AS PRECIO " & _
                       "FROM ((IF4VALES AS A INNER JOIN IF3VALES AS B ON (A.F2CODALM = B.F2CODALM) AND (A.F4NUMVAL = B.F4NUMVAL)) INNER JOIN IF5PLA AS C ON B.F5CODPRO = C.F5CODPRO) INNER JOIN SF1ORIGENES AS D ON A.F1CODORI = D.F1CODORI " & _
                       "WHERE (((A.F2CODALM)='" & ccod_almacen & "') AND ((A.F4NUMVAL)='" & cnum_vale & "') AND ((A.F1CODORI)=[D].[F1CODORI]) AND ((B.F5CODPRO)=[c].[F5CODPRO]) AND ((C.F5CODPRO)=[B].[F5CODPRO]) ) " & _
                       "ORDER BY A.F4NUMVAL, B.F5CODPRO, C.F5CODFAB;"
            End If
    
            .DataControl1.Source = csql
            .fldAlmacen.Text = pnlalmacen.Caption
            .fldnomcosto.Text = pnlccosto.Caption
            .fldempresa.Text = wnomcia
            .fldFecha.Text = Format(Date, "dd/mm/yyyy")
            .fldvale.Text = cnum_vale
            .fldalma.Text = ccod_almacen
            .lblpie2.Caption = "Hecho por"
            .Show vbModal
        End With
    Else
        With acr_vales_p_s
            .DataControl1.ConnectionString = cconex_dbbancos
            ctipo_vale = "S"
            .Lbl_vale.Caption = " VALE DE SALIDA "
            .lblprov.Visible = False
            .lblpunto.Visible = False
            .fldprov.Visible = False
            If wprecio = "1" Then
                csql = "SELECT DISTINCTROW A.F2CODALM, A.F4NUMVAL, A.F1CODORI, D.F1NOMORI, A.F2CODPROV, A.F4NUMDOC, A.F4OBSERVA, A.F4CENTRO, A.F4FECVAL, B.F5CODPRO  AS CODIGO, B.F3CANPRO, C.F5NOMPRO, C.F5CODFAB, C.F7CODMED, C.F5CODPRO, B.F3PUNIT AS F5PREVTA, B.F3PUNIT*B.F3CANPRO AS PRECIO " & _
                       "FROM ((IF4VALES AS A INNER JOIN IF3VALES AS B ON (A.F2CODALM = B.F2CODALM) AND (A.F4NUMVAL = B.F4NUMVAL)) INNER JOIN IF5PLA AS C ON B.F5CODPRO = C.F5CODPRO) INNER JOIN SF1ORIGENES AS D ON A.F1CODORI = D.F1CODORI " & _
                       "WHERE (((A.F2CODALM)='" & ccod_almacen & "') AND ((A.F4NUMVAL)='" & cnum_vale & "') AND ((A.F1CODORI)=[D].[F1CODORI]) AND ((B.F5CODPRO)=[c].[F5CODPRO]) AND ((C.F5CODPRO)=[B].[F5CODPRO]) ) " & _
                       "ORDER BY A.F4NUMVAL, B.F5CODPRO, C.F5CODFAB;"
            Else
                csql = "SELECT DISTINCTROW A.F2CODALM, A.F4NUMVAL, A.F1CODORI, D.F1NOMORI, A.F2CODPROV, A.F4NUMDOC, A.F4OBSERVA, A.F4CENTRO, A.F4FECVAL, B.F5CODPRO  AS CODIGO, B.F3CANPRO, C.F5NOMPRO, C.F5CODFAB, C.F7CODMED, C.F5CODPRO, C.F5PREVTA, C.F5PREVTA*B.F3CANPRO AS PRECIO " & _
                       "FROM ((IF4VALES AS A INNER JOIN IF3VALES AS B ON (A.F2CODALM = B.F2CODALM) AND (A.F4NUMVAL = B.F4NUMVAL)) INNER JOIN IF5PLA AS C ON B.F5CODPRO = C.F5CODPRO) INNER JOIN SF1ORIGENES AS D ON A.F1CODORI = D.F1CODORI " & _
                       "WHERE (((A.F2CODALM)='" & ccod_almacen & "') AND ((A.F4NUMVAL)='" & cnum_vale & "') AND ((A.F1CODORI)=[D].[F1CODORI]) AND ((B.F5CODPRO)=[c].[F5CODPRO]) AND ((C.F5CODPRO)=[B].[F5CODPRO]) ) " & _
                       "ORDER BY A.F4NUMVAL, B.F5CODPRO, C.F5CODFAB;"
            End If
    
            .DataControl1.Source = csql
            .fldAlmacen.Text = pnlalmacen.Caption
            .fldnomcosto.Text = pnlccosto.Caption
            .fldempresa.Text = wnomcia
            .fldFecha.Text = Format(Date, "dd/mm/yyyy")
            .fldvale.Text = cnum_vale
            .fldalma.Text = ccod_almacen
            .lblpie2.Caption = "Hecho por"
            .Show vbModal
       End With
    End If
End Sub

Private Function VALIDA_CONCEPTO_INV(pconcepto As String)
Dim sw_e    As Boolean

    If rsconcepto_inv.State = adStateOpen Then rsconcepto_inv.Close
    rsconcepto_inv.Open "SELECT F1PRECIO,F1NOMORI,F1PARTIDA FROM SF1ORIGENES WHERE F1CODORI='" & pconcepto & "'", cnn_dbbancos
    If Not rsconcepto_inv.EOF Then
        wnomconcepto = Trim(rsconcepto_inv.Fields("F1NOMORI") & "")
        wpartida = Trim(rsconcepto_inv.Fields("F1PARTIDA") & "")
        wprecio = Trim(rsconcepto_inv.Fields("F1PRECIO") & "")
        sw_e = True
    Else
        sw_e = False
    End If
    rsconcepto_inv.Close
    VALIDA_CONCEPTO_INV = sw_e

End Function

Private Sub abofecha_CloseUp()
    If IsDate(abofecha.value) Then
'        If rscambios.State = adStateOpen Then rscambios.Close
'        If ctipoadm_bd = "A" Then
'            rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.value & "')", cnn_dbbancos
'        Else
'            rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.value & "')", cnn_dbbancos
'        End If
'        If Not rscambios.EOF Then
'            txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
'        Else
'            txttc.Text = Format(1, "0.000")
'        End If
'        rscambios.Close
        
        txttc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", abofecha.value, "F")), "#.000")
    Else
        MsgBox "Fecha incorrecta. Verifique.", vbCritical, "Atención"
        
        abofecha.SetFocus
    End If
End Sub

Private Sub abofecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub abofecha_LostFocus()
    On Error Resume Next
    
    If Trim(txtalmacen.Text) <> vbNullString Then
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(txtalmacen.Text)
            
            .FechaInicioMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 0, 1)
            .FechaFinMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 1, 0)
            
            If .verificarCierreVale Then
                MsgBox "Imposible registrar Vale, periodo ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                abofecha.SetFocus
            End If
        End With
    End If
End Sub


Private Sub cmbalmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbalmacen_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    cmbconcepto.SetFocus
'    End If
End Sub

Private Sub cmbalmacen_LostFocus()
    txtalmacen.Text = right(cmbalmacen.Text, 2)
End Sub


Private Sub cmbalmacendes_Click()
    txtalmacendes.Text = right(cmbalmacendes.Text, 2)
End Sub

Private Sub cmbalmacendes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtusuario.SetFocus
    End If
End Sub

Private Sub cmbCategoriaTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbCategoriaTipo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim LenText As Long, ret As Long
    
    'Si los caracteres presionados están entre el 0 y la Z
    If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
        ret = SendMessage(cmbCategoriaTipo.HWnd, &H14C&, -1, ByVal cmbCategoriaTipo.Text)
        
        If ret >= 0 Then
            LenText = Len(cmbCategoriaTipo.Text) 'InStr(1, cmbCategoriaTipo.Text, Space(150)) - 1
            
            cmbCategoriaTipo.ListIndex = ret
            cmbCategoriaTipo.Text = cmbCategoriaTipo.List(ret)
            cmbCategoriaTipo.SelStart = LenText
            cmbCategoriaTipo.SelLength = Len(cmbCategoriaTipo.Text) - LenText
        End If
    End If
End Sub


Private Sub cmbCategoriaTipo_LostFocus()
    If Trim(cmbCategoriaTipo.Text) <> vbNullString Then
        'ModMilano.abrirCnDBMilano
        
        lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "CATEGORIATIPO", "NOMBRE", Trim(cmbCategoriaTipo.Text), "T")
        
        If Trim(lblIdCategoriaTipo.Caption) = vbNullString Then
            MsgBox "Categoria no identificada.", vbInformation + vbOKOnly, App.ProductName
            
            cmbCategoriaTipo.SetFocus
        End If
    End If
End Sub


Private Sub cmbConcepto_Click()
    With objAyudaOrigen
        .inicializarEntidades
        
        .Codigo = Trim(Mid(cmbconcepto.Text, 200))
        
        If .obtenerOrigen Then
            txtconcepto.Text = .Codigo
            
            cmbalmacendes.Visible = .TieneAlmacenDestino
            txtalmacendes.Visible = .TieneAlmacenDestino
            Label6.Visible = .TieneAlmacenDestino
            
            dxDBGrid1.Columns.ColumnByFieldName("PUNIT").DisableEditor = Not .RegistrarCosto
            dxDBGrid1.Columns.ColumnByFieldName("PUNIT").Color = IIf(.RegistrarCosto, &H80000005, &HE0E0E0)
            
            If .TieneAlmacenDestino Then
                habilita_almacen_destino txtconcepto.Text, txtalmacen.Text
                
                If cmbalmacendes.ListCount > 0 Then
                    cmbalmacendes.Visible = .TieneAlmacenDestino
                    txtalmacendes.Visible = .TieneAlmacenDestino
                    Label6.Visible = .TieneAlmacenDestino
                End If
            End If
            
'            With dxDBGrid1
'                If .Dataset.RecordCount > 0 Then
'                    .Dataset.Close
'
'                    abrirCnTemporal
'
'                    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA"
'
'                    AdicionaItem
'
'                    listarGrilla
'                End If
'            End With
        End If
    End With
        
'    txtconcepto.Text = Trim(Mid(cmbconcepto.Text, 200))
'    If cmbalmacen.ListIndex <> -1 Then
'        habilita_almacen_destino txtconcepto.Text, txtalmacen.Text
'        If cmbalmacendes.ListCount <> 0 Then
'            cmbalmacendes.Visible = True
'            txtalmacendes.Visible = True
'            Label6.Visible = True
'        Else
'            cmbalmacendes.Visible = False
'            txtalmacendes.Visible = False
'            Label6.Visible = False
'        End If
'        txtalmacendes.Text = ""
'

'    End If
End Sub

Private Sub cmbconcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbconcepto_LostFocus()
    OCULTA_PRECIO
End Sub

Private Function OCULTA_PRECIO()
        If sw_ayuda = False Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            If Len(Trim(txtconcepto.Text)) > 0 Then
                If VALIDA_CONCEPTO_INV(txtconcepto.Text) = True Then
'                    If wprecio = "1" Then
'                        dxDBGrid1.Columns.ColumnByFieldName("PUNIT").Visible = True
'                        dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = True
'                    Else
'                        dxDBGrid1.Columns.ColumnByFieldName("PUNIT").Visible = False
'                        dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = False
'                    End If
                Else
                    MsgBox "Código de concepto no existe. Verifique", vbCritical, "Atención"
                    txtconcepto.SetFocus
                End If
            Else
                If cmbalmacendes.ListIndex <> -1 Then
                    MsgBox "Falta ingresar el código del concepto. Verifique", vbCritical, "Atención"
                    txtconcepto.SetFocus
                End If
            End If
        End If
    End If

End Function

Private Sub cmbmoneda_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub cmbtipo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        If txtserfac.Visible = True Then
'            txtserfac.SetFocus
'        Else
'            txtnumfac.SetFocus
'        End If
'    End If

End Sub

Private Sub cmbtipo_LostFocus()

'    If right(cmbtipo.Text, 2) = "01" Or right(cmbtipo.Text, 1) = "03" Then
'        lbldocum.Caption = "Serie/Número"
'        txtserfac.Visible = True
'        txtserfac.Enabled = True
'        txtserfac.SetFocus
'    Else
'        txtserfac.Text = "001"
'        lbldocum.Caption = "Número"
'        txtserfac.Visible = False
'        txtserfac.Enabled = False
'        txtnumfac.SetFocus
'    End If

End Sub

Private Sub cmbTipoAuxiliar_Click()
    txtproveedor.Text = vbNullString: txtnomprov.Text = vbNullString
End Sub

Private Sub cmbTipoAuxiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub


Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    'If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            sw_detalle = True
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
                        
            dxDBGrid1.Columns.FocusedIndex = 4
        End If
    'End If
           
End Sub


Private Sub configuraGrilla()
    
    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        '.Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        '.Set (egoRowSelect)
    End With
   
'    dxDBGrid1.Columns(1).Visible = False
'
'    Select Case wvisualiza_cod
'        Case "I"
'            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = False
'        Case "F"
'            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = True
'        Case "T"
'            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = True
'        Case Else
'            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = False
'    End Select
'
'    dxDBGrid1.Columns.ColumnByFieldName("PUNIT").Visible = False
'    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = False
'    dxDBGrid1.Columns.ColumnByFieldName("STOCKACTUAL").Visible = False

    
End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer

    dxDBGrid1.Dataset.Active = False
    
    If sw_nuevo_documento = True Then
        'DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
        abrirCnTemporal
        
        DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
        
        If Len(Trim(dxDBGrid1.Dataset.ADODataset.CommandText)) > 0 Then
            dxDBGrid1.Dataset.Open
            dxDBGrid1.Dataset.ADODataset.Requery
        End If
        dxDBGrid1.Dataset.Refresh
    End If
    
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp 'cnDBTemp
    dxDBGrid1.Dataset.ADODataset.CommandText = "SELECT * FROM TMPVALESALIDA"
    dxDBGrid1.Dataset.Open
    
    dxDBGrid1.Dataset.Active = True
    
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    dxDBGrid1.Dataset.Open
    
    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
    
        For i = 1 To 1
            If sw_nuevo_temp = False Then
                If sw_nuevo_documento = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("ITEM") = i
            .FieldValues("CODPROD") = ""
            .FieldValues("CODFAB") = ""
            .FieldValues("MARCA") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UMEDIDA") = ""
            '.FieldValues("STOCKACTUAL") = Format(0, "###,##0.00")
            .FieldValues("CANTIDAD") = Format(0, "###,##0.00")
            .FieldValues("PUNIT") = Format(0, "###,##0.00")
            .FieldValues("TOTAL") = Format(0, "###,##0.00")
        Next
        .Post
        
        sw_nuevo_item = False
        
    End With
    
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    'If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value & "")) = 0 Then
                    Allow = False
                Else
                    dxDBGrid1.Columns.FocusedIndex = 1
                End If
            End If
        End If
        If Action = daDelete Then
            sw_detalle = True
            dxDBGrid1.Dataset.Refresh
        End If
    'End If
    If Action = daEdit Then
        isaldo = Val("" & dxDBGrid1.Columns.ColumnByFieldName("cantidad").value)
    End If
End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case Column.FieldName
        Case "CODPROD"
            If Trim(txtconcepto.Text) = vbNullString Then
                MsgBox "Seleccione el Concepto.", vbInformation + vbOKOnly, App.ProductName
                
                cmbconcepto.SetFocus
                
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            With objAyudaOrigen
                .inicializarEntidades
                
                .Codigo = Trim(txtconcepto.Text)
                
                .obtenerConfigOrigen
            
                Select Case .CodigoAyudaProducto
                    Case "1" '"XJ1" 'Ajuste de Salida
'                        If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                            Unload ayuda_productos
'                        End If
                        
                        With ayuda_productos_salida
'                            .CodigoAuxiliar = vbNullString
'                            .CodigoRequerimiento = vbNullString
'                            .CodigoProducto = vbNullString
'
'                            .CadenaCorte = "" 'InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                            Con_Ayu = 2
                            .Show 1
                        End With
                        
'                        abrirCnTemporal
                        If Len(Trim(wcodproducto)) > 0 Then
                        dxDBGrid1.Dataset.Edit
                        sw_detalle = True
                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = wcodproducto
                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = wcodfab
                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = wmedida
                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = wmarca
                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
                        dxDBGrid1.Dataset.Post
                        dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
        End If
'                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPPRODUCTOS", "F4PERINT", "-1", "N") & "") <> 0 Then
'                            copiarSeleccionAyudaProductos
'                        End If
                    Case "2" '"XRQ", "X3R", "XV1" 'Por Requerimiento, Servicios a Terceros, Venta
                        If ModUtilitario.validarFormAbierto("frmUtilStockDisponible") Then
                            Unload frmUtilStockDisponible
                        End If
                        
                        With frmUtilStockDisponible
'                            .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                            
                            .Show 1
                            
                            cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                            cmbalmacen.Enabled = False
                        End With
                        
                        abrirCnTemporal
                        
                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILSTOCKDISPONIBLE", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                            copiarSeleccionStockDisponible
                        End If
                    Case Else
                        MsgBox "Ayuda de Concepto de Movimiento no configurado, verifique.", vbInformation + vbOKOnly, App.ProductName
                End Select
            End With
            
            listarGrilla
            
            Me.MousePointer = vbDefault
    End Select
    
'    If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNCODPROD" Or dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNCODFAB" Then
'        wcod_alm = txtalmacen.Text
'        wcodproducto = ""
'        wcodpartida = txtnumorden.Text
'        wcodpresupuesto = txtccosto.Text
'        If Trim(wcod_alm) = "" Then
'            MsgBox "Ingrese Almacén", vbInformation, "Sistema de Logistica"
'            Exit Sub
'        End If
'        sw_ayuda_prod = True
'        Con_Ayu = 2
'        ayuda_productos.Show 1
'        Me.MousePointer = vbdefault
'        '-----
'        With ayuda_productos.dxDBGrid1
'            If Len(Trim(wcodproducto)) = 0 Then
'                .Dataset.Filtered = True
'                .Dataset.Filter = "F4PERINT = -1"
'                .Dataset.First
'                X = 0
'                    Do While Not .Dataset.EOF
'                        z = .Dataset.RecordCount
'                        If z = 0 Then Exit Sub
'                        X = X + 1
'                        If dxDBGrid1.Columns.ColumnByFieldName("COD_PRODUCTO").Value = "" Then
'                            dxDBGrid1.Dataset.Edit
'                        Else
'                            dxDBGrid1.Dataset.Append
'                        End If
'                        dxDBGrid1.Dataset.Edit
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = .Columns.ColumnByFieldName("f5codpro").Value
'                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = .Columns.ColumnByFieldName("f5codpro").Value
'                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = .Columns.ColumnByFieldName("f5nompro").Value
'                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = .Columns.ColumnByFieldName("F7SIGMED").Value
'                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = .Columns.ColumnByFieldName("f5fob").Value
'                         dxDBGrid1.Dataset.Post
'                        .Dataset.Next
'                        Loop
'                        If X = 0 And Len(Trim(wcodproducto)) > 0 Then
'                            dxDBGrid1.Dataset.Edit
'                            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = wcodproducto
'                            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = wcodfab
'                            dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
'                            dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = wdesproducto
'                            dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida, "T", cnn_dbbancos)
'                            dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
'                            dxDBGrid1.Dataset.Post
'                        End If
'                    Else
'                        dxDBGrid1.Dataset.Edit
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = wcodproducto
'                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = wcodfab
'                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
'                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = wdesproducto
'                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida, "T", cnn_dbbancos)
'                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
'                        dxDBGrid1.Dataset.Post
'                    End If
'                Unload ayuda_productos
'            End With
'            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
''''        Unload ayuda_productos
''''        If Len(Trim(wcodproducto)) > 0 Then
''''            dxDBGrid1.Dataset.Edit
''''            sw_detalle = True
''''            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = wcodproducto
''''            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = wcodfab
''''            dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = wdesproducto
''''            dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = wmedida
''''            dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
''''            dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
''''            dxDBGrid1.Dataset.Post
''''            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
''''        End If
'    End If
    If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNELIMINAR" Then
        If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Count = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
                sw_detalle = False
                'SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
            Else
                dxDBGrid1.Dataset.Delete
            End If
            sw_nuevo_item = False
        End If
    ElseIf dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNPARTIDA" Then
        wnumordentrab = ""
        wcod_alm = ""
        sw_ayuda_prod = True
        wllamada = 0
        ayuda_orden_trab.Show 1
        Unload ayuda_orden_trab
        If Len(Trim(wnumordentrab)) > 0 Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("PARTIDA").value = wnumordentrab
            dxDBGrid1.Columns.ColumnByFieldName("DESPARTIDA").value = wobservacion
            dxDBGrid1.Dataset.Post
            'Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Index
        End If

    End If

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Rem SK ADD:
    Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
        Case "CANTIDAD", "PUNIT"
            With objAyudaOrden
                'DATOS
                .PorcentajeImpuesto = IIf(right(cmbtipo.Text, 2) = "02", gretenc, wwigv) / 100
                .SignoImpuesto = IIf(right(cmbtipo.Text, 2) = "02", -1, 1)
                
                .Cantidad = Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & "")
                .CantidadMaxima = Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDADMAX").value & "")
                .PorcentajeDemasia = 0
                
                Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
                    Case "PUNIT"
                        .PrecioSinImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value & "")
                        .PrecioConImpuesto = 0
'                    Case "PVUNIT"
'                        .PrecioSinImpuesto = 0
'                        .PrecioConImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value & "")
                    Case Else
                        .PrecioSinImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value & "")
                        '.PrecioConImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value & "")
                End Select
                
'                Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
'                    Case "PORDESC"
'                        .PorcentajeDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("PORDESC").value & "") / 100
'                        .TotalDscto = 0
'                    Case "VALDESC"
'                        .PorcentajeDscto = 0
'                        .TotalDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("VALDESC").value & "")
'                    Case Else
'                        .PorcentajeDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("PORDESC").value & "") / 100
'                        .TotalDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("VALDESC").value & "")
'                End Select
                
                If Trim(dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value & "") = "*" Then
                    .Afecto = True
                Else
                    .Afecto = False
                End If

'                'RESULTADOS
                
                dxDBGrid1.Dataset.Edit
                
                If .CantidadMaxima > 0 Then
                    If .Cantidad > .CantidadMaxima Then
                        MsgBox "La cantidad no puede exceder al origen seleccionado, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        dxDBGrid1.Dataset.Cancel
                        
                        Exit Sub
                    End If
                End If
                
                'CALCULOS
                .calculosPorItem
                
                dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value = .PrecioSinImpuesto
                'dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = .PrecioConImpuesto
                'dxDBGrid1.Columns.ColumnByFieldName("PORDESC").value = Val(Format(.PorcentajeDscto * 100, "#0.00"))
                'dxDBGrid1.Columns.ColumnByFieldName("F3CANPROFINAL").value = .CantidadFinal
                'dxDBGrid1.Columns.ColumnByFieldName("VALDESC").value = .TotalDscto
                
                'dxDBGrid1.Columns.ColumnByFieldName("COSTOUNINETO").value = .PrecioNetoSinImpuesto
                
                'dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = .BasePorItem
                'dxDBGrid1.Columns.ColumnByFieldName("F3MONINA").value = .ExoneradoPorItem
                'dxDBGrid1.Columns.ColumnByFieldName("IGV").value = .ImpuestoPorItem
                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = .TotalPorItem
                
                dxDBGrid1.Dataset.Post
            End With
    End Select
    
    
    
'    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
'        wcod_alm = txtalmacen.Text
'
'        If Trim(wcod_alm) = "" Then
'            MsgBox "Ingrese Almacen", vbInformation, "Sistema de Logistica"
'            Exit Sub
'        End If
'
'        If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODPROD" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CODFAB" Then
'
'            If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODPROD" Then
'                wcodproducto = "" & dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value
'            Else
'                wcodproducto = "" & dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value
'            End If
'            If rsif5pla.State = adStateOpen Then rsif5pla.Close
'            'OR (((A.F5CODFAB)='" & wcodproducto & "'))
'            'sql = "SELECT A.F5CODPRO, A.F5CODFAB, A.F5NOMPRO, A.F5MARCA, A.F5VALVTA, A.F5PRECOS, C.F7SIGMED, D.F2DESMAR, A.F7CODMED FROM ((IF5PLA AS A INNER JOIN IF6ALMA AS B ON A.F5CODPRO = B.F5CODPRO) INNER JOIN EF7MEDIDAS AS C ON A.F7CODMED = C.F7CODMED) INNER JOIN EF2MARCAS AS D ON A.F5MARCA = D.F2CODMAR WHERE (((A.F5CODPRO)='" & wcodproducto & "') AND ((A.F5MARCA)=[D].[F2CODMAR]) AND ((A.F7CODMED)=[C].[F7CODMED]) AND ((B.F2CODALM)='" & wcod_alm & "')) ORDER BY A.F5NOMPRO;"
'            sql = "SELECT A.F5CODPRO, A.F5CODFAB, A.F5NOMPRO, A.F5VALVTA, A.F5PRECOS, C.F7SIGMED, D.F2DESMAR FROM (IF5PLA AS A LEFT JOIN EF7MEDIDAS AS C ON A.F7CODMED = C.F7CODMED) LEFT JOIN EF2MARCAS AS D ON A.F5MARCA = D.F2CODMAR WHERE (((A.F5CODPRO)='" & wcodproducto & "'))"
'            rsif5pla.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
'            If Not rsif5pla.EOF Then
'                wVariosProductos = False
'                gcodpro = wcodproducto
'                If rsif5pla.RecordCount > 1 Then      'Existe más de 1 producto con el mismo codigo de fabricante
'                    frmdetalle.Caption = "Producto " & gcodpro
'                    frmdetalle.grddetalle.Rows = 1
'                    wVariosProductos = True
'                    wf5codpro = ""
'                    Do While Not rsif5pla.EOF
'                        frmdetalle.grddetalle.Rows = frmdetalle.grddetalle.Rows + 1
'                        frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 1) = "" & rsif5pla("f2desmar")
'                        frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 2) = "" & rsif5pla("f5codpro")
'                        frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 3) = "" & rsif5pla("f5nompro")
'                        rsif5pla.MoveNext
'                    Loop
'                    rsif5pla.MoveFirst
'                    frmdetalle.Show vbModal
'                End If
'                If wVariosProductos Then
'                    If Len(Trim(wf5codpro)) > 0 Then
'                        cad = "f5codpro='" & wf5codpro & "'"
'                        rsif5pla.Find cad
'                    Else
'                        dxDBGrid1.Dataset.Edit
'                        sw_detalle = True
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value = 0#
'                        dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = ""
'                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
'                        dxDBGrid1.Dataset.Post
'                        wcodproducto = ""
'                        dxDBGrid1.Columns.FocusedIndex = 1
'                        rsif5pla.Close
'                        Exit Sub
'                    End If
'                End If
'
'                dxDBGrid1.Dataset.Edit
'                sw_detalle = True
'                dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = rsif5pla.Fields("F5CODPRO")
'                xcodproducto = "" & rsif5pla.Fields("F5CODPRO")
'                dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = "" & rsif5pla.Fields("F5CODFAB")
'                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = "" & rsif5pla.Fields("F5NOMPRO")
'                dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = "" & rsif5pla.Fields("F7SIGMED")
'                dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
'    '            If Rs.State = adStateOpen Then Rs.Close
'    '            Rs.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & "" & rsif5pla.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    '            If Not Rs.EOF Then
'    '                wmarca = Trim(Rs.Fields("F2DESMAR") & "")
'    '            End If
'    '            Rs.Close
'                dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = "" & rsif5pla.Fields("F2DESMAR")
'                dxDBGrid1.Dataset.Post
'                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
'            Else
'                dxDBGrid1.Dataset.Edit
'                MsgBox "El Producto No Existe", vbInformation, "Atención"
'                dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = ""
'                dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = ""
'                wcodproducto = ""
'                dxDBGrid1.Dataset.Post
'                dxDBGrid1.Columns.FocusedIndex = 0
'            End If
'            rsif5pla.Close
'        End If
'        If dxDBGrid1.Columns.FocusedColumn.FieldName = "PUNIT" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CANTIDAD" Then
'
'            If dxDBGrid1.Columns.FocusedColumn.FieldName = "CANTIDAD" Then
'                istock = Val("" & dxDBGrid1.Columns.ColumnByFieldName("cantidad").value)
'                If istock > 0 Then
'                    'Verifica si hay Existencias del producto en Almacén
'                    wf1evalua_stock = 1
'                    If Len(Trim(wf1evalua_stock)) = 0 Then
'                        xcodproducto = "" & dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value
'                        nstock = CalculaExistencia(txtalmacen.Text, xcodproducto, abofecha.value)
'                        If (nstock + isaldo) < istock Then
'                            MsgBox "La Cantidad por salir " & istock & " es mayor a la cantidad en Stock " & nstock, vbInformation, "Sistema de Logistica"
'                            dxDBGrid1.Dataset.Edit
'                            dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = ""
'                            dxDBGrid1.Dataset.Post
'                            dxDBGrid1.Columns.FocusedIndex = 4
'                        End If
'                    End If
'                End If
'            Else
'                dxDBGrid1.Dataset.Edit
'                sw_detalle = True
'                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value
'                dxDBGrid1.Dataset.Post
'            End If
'
'        End If
'    End If
    
'    If dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0 And dxDBGrid1.Count = 1 Then
'        SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
'    Else
'        SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'    End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyF2
            'dxDBGrid1_OnEditButtonClick dxDBGrid1.Columns.FocusedColumn, Nothing
        Case vbKeyTab
            Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
                Case "DESCRIPCION" '"CODFAB", "CODPROD"
                    If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then dxDBGrid1.Dataset.Post
                    
                    If Trim(txtconcepto.Text) = vbNullString Then
                        MsgBox "Seleccione el Concepto.", vbInformation + vbOKOnly, App.ProductName
                        
                        cmbconcepto.SetFocus
                        
                        Exit Sub
                    End If
                    
                    Me.MousePointer = vbHourglass
                    
                    With objAyudaOrigen
                        .inicializarEntidades
                        
                        .Codigo = Trim(txtconcepto.Text)
                        
                        .obtenerConfigOrigen
                    
                        Select Case .CodigoAyudaProducto
                            Case "1" '"XJ1" 'Ajuste de Salida
'                                If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                                    Unload ayuda_productos
'                                End If
'
'                                With ayuda_productos
'                                    .CodigoAuxiliar = vbNullString
'                                    .CodigoRequerimiento = vbNullString
'                                    .CodigoProducto = vbNullString
'
'                                    .CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value & "") 'InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
'
'                                    .Show 1
'                                End With
'
'                                abrirCnTemporal
'
'                                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPPRODUCTOS", "F4PERINT", "-1", "N") & "") <> 0 Then
'                                    copiarSeleccionAyudaProductos
'                                End If
                                
                                If dxDBGrid1.Dataset.State = dsEdit Then
                                    dxDBGrid1.Dataset.Post
                                End If
                                
                                If ModUtilitario.validarFormAbierto("frmListaBien") Then
                                    Unload frmListaBien
                                End If
                                
                                With frmListaBien
                                    '.Ayuda = True
                                    '.TieneMovimientoAlmacen = True
                                    '.InsumoOP = False
                                    '.CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value & "")
                                    
                                    .Ayuda = True
                                    .InsumoOP = False
                                    .ParaVenta = False
                                    .TieneMovimientoAlmacen = True
                                    .CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value & "")
                                    .FiltroAdicional = vbNullString
                                    .TipoBienMostrar = "P"
                                    
                                    objAyudaBien.inicializarEntidades
                                    
                                    .Show 1
                                    
                                    If objAyudaBien.Codigo <> vbNullString Then
                                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(CODPROD)", "TMPVALESALIDA", "CODPROD", objAyudaBien.Codigo, "T", "AND TRIM(F4NUMORD & '') = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "") & "' AND TRIM(COD_SOLICITUD & '') = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value & "") & "'")) > 0 Then
                                            MsgBox "Producto ya seleccionado, verifique.", vbInformation + vbOKOnly, App.ProductName
                                            
                                            Me.MousePointer = vbDefault
                                            
                                            Exit Sub
                                        End If
                                        
                                        
                                        objAyudaBien.obtenerConfigBien
                                        
                                        With dxDBGrid1
                                            .Dataset.Edit
                                            
                                            .Columns.ColumnByFieldName("CODPROD").value = objAyudaBien.Codigo
                                            .Columns.ColumnByFieldName("CODPRODORIGINAL").value = objAyudaBien.Codigo
                                            .Columns.ColumnByFieldName("DESCRIPCION").value = objAyudaBien.Descripcion
                                            .Columns.ColumnByFieldName("UMEDIDA").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                                            .Columns.ColumnByFieldName("CANTIDAD").value = 0
                                            .Columns.ColumnByFieldName("CANTIDADMAX").value = 0
                                            
                                            .Dataset.Post
                                        End With
                                    End If
                                End With
                                
                            Case "2" '"XRQ", "X3R", "XV1" 'Por Requerimiento, Servicios a Terceros, Venta
                                If ModUtilitario.validarFormAbierto("frmUtilStockDisponible") Then
                                    Unload frmUtilStockDisponible
                                End If
                                
                                With frmUtilStockDisponible
                                    .CodigoAlmacen = right(cmbalmacen.Text, 2)
                                    .CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value & "") 'InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                                    
                                    .Show 1
                                    
                                    'cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                                End With
                                
                                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                    If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(*)", "TMPCPSTOCKDISPONIBLE" & UCase(wusuario), "PROCESAR", "1", "N") & "") <> 0 Then
                                        copiarSeleccionStockDisponibleSql
                                        
                                        cmbalmacen.Enabled = False
                                    Else
                                        cmbalmacen.Enabled = True
                                    End If
                                Else
                                    abrirCnTemporal
                                    
                                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILSTOCKDISPONIBLE", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                                        copiarSeleccionStockDisponible
                                        
                                        cmbalmacen.Enabled = False
                                    Else
                                        cmbalmacen.Enabled = True
                                    End If
                                End If
                            Case Else
                                MsgBox "Ayuda de Concepto de Movimiento no configurado, verifique.", vbInformation + vbOKOnly, App.ProductName
                        End Select
                    End With
                    
                    listarGrillaVale
                    
                    Me.MousePointer = vbDefault
            End Select
    End Select
    
    If KeyCode = 113 Then
        If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNCODPROD" Or dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNCODFAB" Then
            wcod_alm = txtalmacen.Text
            wcodproducto = ""
            sw_ayuda_prod = True
            Con_Ayu = 2
            ayuda_productos_salida.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                sw_detalle = True
                dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = wcodfab
                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = wmarca
                dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
            End If
        End If
    End If
'    If KeyCode = 115 Or KeyCode = 46 Then
'        If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
'            If dxDBGrid1.Dataset.RecNo = 1 Then
'                dxDBGrid1.Dataset.Delete
'                AdicionaItem
'            Else
'                dxDBGrid1.Dataset.Delete
'            End If
'        End If
'    End If
    
End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)
    
    
'    Select Case KeyCode
'        Case 123:
'            wsw_codbarra = wsw_codbarra + 1
'            If wsw_codbarra = 1 Then
'                wcodproducto = "": wdesproducto = "": wcodmar = "": wcodfab = "": wmarca = ""
'                wstock = 0# ':  gvalvta = 0#: wigv = 0#: gprevta = 0#: wfactor = 0#
'                wcodigo_barra = ""
'                lee_codigosbarra.Show 1
'                If Len(Trim(wcodigo_barra)) > 0 Then
'                    wcodproducto = wcodigo_barra
'                    'wcodfab = wcodigo_barra
'                    sw_nuevo_item = True
'                    dxDBGrid1.Dataset.Edit
'                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
'                    rsif5pla.Open "SELECT F5CODPRO,F5CODFAB,F5NOMPRO,F7CODMED,F5GRUPO,F5TIPO,F5IGVVTA,F5VALVTA,F5PREVTA,F5FACTOR,F5tipoporc,F5AFECTO,F5MARCA FROM IF5PLA WHERE F5CODPRO='" & "" & Trim(wcodproducto) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                    'rsif5pla.MoveFirst
'                    If Not rsif5pla.EOF Then
'                        'rs.Open "SELECT COUNT(*) AS CANTI FROM IF5PLA WHERE F5CODFAB='" & "" & Trim(wcodfab) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                        'If Not rs.EOF Then
'                        '    wcant = rs.Fields("CANTI")
'                        'Else
'                        '    wcant = 0
'                        'End If
'                        'If wcant = 1 Then
'                            wcodproducto = "" & rsif5pla.Fields("F5CODPRO")
'                            wcodfab = "" & rsif5pla.Fields("F5CODFAB")
'                            wdesproducto = "" & rsif5pla.Fields("F5NOMPRO")
'                            wcodmar = "" & rsif5pla.Fields("F7CODMED")
'                            wmarca = "" & rsif5pla.Fields("F5MARCA")
'                        'Else
'                        '    MsgBox "INGRESE LA MARCA"
'                        'End If
'                    Else
'                        MsgBox "Código no existe", vbInformation + vbDefaultButton1, "Atención"
'                        wcodproducto = "": wdesproducto = "": wcodmar = "": wcodfab = "": wmarca = ""
'                        wstock = 0# ': gvalvta = 0#: wigv = 0#: gprevta = 0#: wfactor = 0#
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = wcodproducto
'                        dxDBGrid1.Dataset.Post
'                        Exit Sub
'                    End If
'                    rsif5pla.Close
'                    dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = wcodproducto
'                    dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = wmarca
'                    dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
'                    dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = wcodmar
'                    dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = wcodfab
'                    sw_detalle = True
'                    '----------------------------------------------------------------
'                    dxDBGrid1.Dataset.Post
'                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
'                    sw_nuevo_item = False
'                    wcodigo_barra = ""
'                End If
'                Unload lee_codigosbarra
'            Else
'                wsw_codbarra = 0
'            End If
'    End Select
End Sub

Private Sub Form_Activate()
'    If sw_activate = True Then
'        sw_activate = False
'        If wmoneda_productos = "S" Then
'            cmbmoneda.ListIndex = 0
'        Else
'            cmbmoneda.ListIndex = 1
'        End If
'        'txtalmacen.SetFocus
'    End If
End Sub

Private Sub Form_Load()
'    Dim CadSql          As String
'    Dim pnumvale        As String
'    Dim palmacen        As String
'
'    Me.MousePointer = vbHourglass
'    sw_nuevo_item = True
'    Me.left = 1600
'    Me.top = 1150
'
'    cmbmoneda.AddItem "Soles"
'    cmbmoneda.AddItem "Dólares"
'
'    sw_activate = True
'    'cnombase = wusuario & "VALES" & Format(Time, "hh_mm_ss") & ".MDB"
'    'CREATEDATABASE_N wrutatemp & "\", cnombase
'    'cnombase = wusuario & "VALES" & Format(Time, "hh_mm_ss") & ".MDB"
'
''    cnombase = "TEMPLUS.mdb"
''
''    "TMPVALESALIDA" = "tmpValeSalida"
''    If cnDBTemp.State = adStateOpen Then cnDBTemp.Close
''    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
''    cnDBTemp.Open cconex_form
''    DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
'
'    abrirCnTemporal
'
'    DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
'
'    'CadSql = "(ITEM TEXT(4),CODPROD TEXT(15),CODFAB TEXT(15),MARCA TEXT(4),DESCRIPCION TEXT(100)," & _
'             "UMEDIDA TEXT(3),STOCKACTUAL DOUBLE,CANTIDAD DOUBLE,PUNIT DOUBLE,TOTAL DOUBLE)"
'    'CREATETABLE_N "TMPVALESALIDA", CadSql, cnDBTemp
'    configuraGrilla
'
'    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
'
'    rsdocumentos.Open "SELECT * FROM DOCUMENTOS ORDER BY F2CODDOC", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'
'    If Not rsdocumentos.EOF Then
'        rsdocumentos.MoveFirst
'
'        Do While Not rsdocumentos.EOF
'            cmbtipo.AddItem rsdocumentos.Fields("F2DESDOC") & "" & Space(100) & rsdocumentos.Fields("F2CODDOC") & ""
'
'            rsdocumentos.MoveNext
'        Loop
'    End If
'    rsdocumentos.Close
'
'    ModMilano.listarCategoriaTipo cmbCategoriaTipo
'
'    If Rs.State = adStateOpen Then Rs.Close
'    Rs.Open "select f2codalm,f2nomalm from ef2almacenes order by f2nomalm asc", cnn_dbbancos, adOpenStatic, adLockReadOnly
'    X = 0
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        Do While Not Rs.EOF
'            cmbalmacen.AddItem Rs.Fields("f2nomalm") & "" & Space(50) & Rs.Fields("F2CODALM") & ""
'            Rs.MoveNext
'            X = X + 1
'        Loop
'    End If
'    Rs.Close
'
'    If X = 3 Then
'        cmbalmacen.ListIndex = 0
'    End If
'    sw_detalle = False
'
'    If sw_nuevo_documento = True Then
'        nuevo
'        AdicionaItem
'        'dxDBGrid1.Enabled = False
'    Else
'        palmacen = lista_vales.dxDBGrid1.Columns.ColumnByFieldName("f2codalm").value
'        'palmacen_destino = walmacen_destino
'        pnumvale = lista_vales.dxDBGrid1.Columns.ColumnByFieldName("f4numval").value
'        BUSCA_VALE palmacen, pnumvale
'        sw_cabecera = False
'        sw_nuevo_documento = False
'        'ADICIONAR PARA BLOQUEAR LOS CONTROLES QUE NO PERMITA MODIFICACION
'
'    End If
'    wsw_codbarra = 0
'    xvale = 2
'    Me.MousePointer = vbDefault

    Me.MousePointer = vbHourglass
    
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    configuraGrilla
    
    listarAlmacenEnCombo
    
    listarTipoDocumentoEnCombo
    
'    ModMilano.listarCategoriaTipo cmbCategoriaTipo
    
    consultarVale
    
    'Activar Control de Apertura de Formulario
    '(Para evitar abrir mas de una vez, el mismo formulario en diferentes Instancias del Programa)
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeSalidaAbierto", "1"
    
    Me.MousePointer = vbDefault
End Sub

Private Sub nuevo()
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = True
    
    sw_nuevo_documento = True
    txtnumero.Text = ""
    txtserie.Text = ""
    txtnumdoc.Text = ""
    pnlalmacen.Caption = ""
    abofecha.value = Format(Date, "DD/MM/YYYY")
    txtconcepto.Text = "": pnlconcepto.Caption = ""
    
    If Trim(wtiposalida) = "*" Then  '----- LA EMPRESA ES CONSTRUCTORA (AIC)
        txtalmacen.Text = "01"
        If VALIDA_ALMACEN(txtalmacen.Text) = True Then
            pnlalmacen.Caption = wnomalmacen
        End If
        txtconcepto.Text = wconc_compra: pnlconcepto.Caption = ""
    End If
            
    txttc.Text = "1.000"
    txtccosto.Text = "": pnlccosto.Caption = ""
    txtobserva.Text = ""
    cmbalmacen.ListIndex = -1
    cmbconcepto.ListIndex = -1
    cmbalmacendes.ListIndex = -1
    txtalmacen.Text = ""
    txtconcepto.Text = ""
    txtalmacendes.Text = ""
    pnlsolicitante.Caption = ""
    txtusuario.Text = ""
    txtnumorden.Text = ""
    pnlorden.Caption = ""
    pnlccosto.Caption = ""
    txtserfac.Text = ""
    txtnumfac.Text = ""
    
    sw_detalle = False
    sw_cabecera = False
    'tc
   If IsDate(abofecha.value) Then
        If rscambios.State = adStateOpen Then rscambios.Close
        If ctipoadm_bd = "A" Then
            rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.value & "')", cnn_dbbancos
        Else
            rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.value & "')", cnn_dbbancos
        End If
        If Not rscambios.EOF Then
            txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
        Else
            txttc.Text = Format(1, "0.000")
        End If
        rscambios.Close
    End If
    '''''''''
    Rem SK ADD:
    fraOrdenProduccion.Enabled = True
    txtIDOrdenProduccion.Text = vbNullString
        cmbCategoriaTipo.ListIndex = -1
        txtNroOrdenProduccion.Text = vbNullString
    
    lblNumeroValeExterno.Caption = "< ID Externo >"
    chkExportarVale.value = vbChecked: chkExportarVale.Enabled = True
    
    cmbalmacen.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Desea salir del registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Cancel = 1
        
        'Exit Sub
    Else
        Me.dxDBGrid1.Dataset.Close
        
        With lista_vales
            .listarVale
        End With
        
        ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeSalidaAbierto", "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    sw_nuevo_item = True
'    dxDBGrid1.Dataset.Close
'    'ELIMINA_BD_N wrutatemp, cnombase
'
    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
'
'    lista_vales.dxDBGrid1.Dataset.Active = False
'    lista_vales.dxDBGrid1.Dataset.Refresh
'    lista_vales.dxDBGrid1.Dataset.Active = True
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.ID
        Case "ID_Nuevo":
'            If dxDBGrid1.Dataset.RecordCount > 1 Or Trim(txtIDOrdenProduccion.Text) <> vbNullString Or Trim(txtproveedor.Text) <> vbNullString Then
'                If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'                    Exit Sub
'                End If
'            End If
'
'            Me.MousePointer = vbHourglass
'            sw_nuevo_documento = True
'            AdicionaItem
'            nuevo
'            cmbalmacen.SetFocus
'            sw_nuevo_documento = True
'            Me.MousePointer = vbDefault
            
            If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            strCodAlmacen = vbNullString
            strNumeroVale = vbNullString
            
            consultarVale
            
            Me.MousePointer = vbDefault
        Case "ID_Grabar":
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'            'Validacion del Punto (PC) que origina el Vale
'            ModMilano.abrirCnDBMilano
'
'            If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'                MsgBox "Su computador no esta habilitado para generar registros de SALIDA." & vbNewLine & vbNewLine & _
'                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'                Exit Sub
'            End If
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                dxDBGrid1.Dataset.Post
'                sw_detalle = True
'            End If
'
'            If MsgBox("¿Desea grabar el Vale de Salida?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                Me.MousePointer = vbHourglass
'
'                grabar
'
'                If sw_GRABA_REGISTRO_logistica Then
'                    If ModMilano.exportarValeAserverSQL(Trim(txtalmacen.Text), Trim(txtnumero.Text), lblNumeroValeExterno) Then
'                        MsgBox "Vale Exportado.", vbInformation + vbOKOnly, App.ProductName
'                    End If
'                End If
'
'                BUSCA_VALE Trim(txtalmacen.Text), Trim(txtnumero.Text)
'
'                sw_detalle = False
'                sw_cabecera = False
'
'                Me.MousePointer = vbDefault
'            End If
            
            Me.MousePointer = vbHourglass
            
            validarCajas
            
            Me.MousePointer = vbDefault
        Case "ID_Eliminar":
'            Me.MousePointer = vbHourglass
'
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'            'Restricción de Anulación de Vale
'            If Month(CDate(abofecha.value)) < Month(Date) Then
'                MsgBox "Imposible eliminar Vale. Fuera del Periodo Actual." & vbNewLine & vbNewLine & _
'                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'                Me.MousePointer = vbDefault
'
'                Exit Sub
'            End If
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'
'            elimina txtnumero.Text, txtalmacen.Text
'
'
'
'            Me.MousePointer = vbDefault
            
            Me.MousePointer = vbHourglass
            
            eliminarVale
            
            Me.MousePointer = vbDefault
        Case "ID_ImprimirA4":
            With objAyudaVale
                .TipoVale = "S"
                .CodigoAlmacen = Trim(txtalmacen.Text)
                .NumeroVale = Trim(txtnumero.Text)
                
                If Not .verificarExistencia Then
                    MsgBox "Vale no registrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            End With
            
            'IMPRIMIR_VALES (1)
            With rptValeIngreso
                .TipoVale = "S"
                .CodAlmacen = Trim(txtalmacen.Text)
                .NumeroVale = Trim(txtnumero.Text)
                
                'ModMilano.abrirCnDBMilano
                
'                .fldCategoria.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CT.NOMBRE", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
'                .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
                .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4CENTRO", "IF4VALES", "F4NUMVAL", Trim(txtnumero.Text), "T"), "T")
                .Show 1
            End With
        Case "ID_ImprimirA5":
            IMPRIMIR_VALES (2)
        Case "ID_CargarData":
            frmExcel.Show 1
        Case "ID_Lista":
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                dxDBGrid1.Dataset.Post
'                sw_detalle = True
'            End If
'            If sw_detalle = True And sw_cabecera = True Then
'                If (dxDBGrid1.Count >= 1 And dxDBGrid1.Columns.ITEM(1).value <> "" And sw_nuevo_documento = True) Then
'                    If MsgBox("El Vale de Salida no ha sido grabado ... Desea Grabar ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'                        grabar
'                        sw_detalle = False
'                    End If
'                End If
'            End If
'            Unload Me
            
            Unload Me
        Case "ID_Calculadora":
            'Calculadora.Show 1
            x = Shell("calc.exe", vbNormalFocus)
        Case "CerrarVale"
            If bolObviarCierre Then
                Exit Sub
            End If
            
            If CBool(Tool.State) Then
                cerrarVale
            Else
                abrirVale
            End If
        Case "ID_Salir":
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                dxDBGrid1.Dataset.Post
'                sw_detalle = True
'            End If
'            If sw_detalle = True And sw_cabecera = True Then
'                If (dxDBGrid1.Count >= 1 And dxDBGrid1.Columns.ITEM(1).value <> "" And sw_nuevo_documento = True) Then
'                    If MsgBox("El Vale de Salida no ha sido grabado ... Desea Grabar ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'                        grabar
'                        sw_detalle = False
'                    End If
'                End If
'            End If
'            Unload Me
'            Unload lista_vales
            
            Unload Me
        Case "DevolucionCompra" 'Opcion para Devolucion de Mercaderia Ingresada por O/C
            MsgBox "Opción en revisión, proceda a la Devolución mediante la ayuda habilitada para el concepto desde el detalle.", vbInformation + vbOKOnly, App.ProductName
            
            Me.MousePointer = vbHourglass
            
            If ModUtilitario.validarFormAbierto("frmUtilDevolucionOC") Then
                Unload frmUtilDevolucionOC
            End If
            
            With frmUtilDevolucionOC
                .TipoVale = "S"
                
                .Show 1
                
                cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XNC", "DER", 3)
            End With
            
            abrirCnTemporal
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILDEVOLUCIONOC", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                copiarSeleccionDevolucionOC
            End If
            
            listarGrilla
            
            Me.MousePointer = vbDefault
    End Select

End Sub

Private Sub grabar()
    'On Error GoTo HndError
    Dim calma_obra      As String
    
    sw_ingreso = False
'   'cnn_dbbancos.BeginTrans
    GRABA_SAL_ALMACEN_CENTRAL
    
    If cmbalmacendes.Text <> "" Then
        sw_ingreso = True
        GRABA_SAL_ALMACEN_CENTRAL
    End If
    ''cnn_dbbancos.CommitTrans
    sw_nuevo_documento = False
    
    Exit Sub
    
HndError:
'    'cnn_dbbancos.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox "Ha Ocurrido el siguiente error:" & Chr(13) & Chr(13) & Err.Description & "." & Chr(13) & "La Operación de Actualización no se Realizó. Consulte al Proveedor.", vbCritical, "Sistema de Logistica"
    Exit Sub
    
End Sub

Private Sub txtalmacen_Change()
    pnlalmacen.Caption = ""

    If Trim(txtalmacen.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtAlmacen_DblClick()

    txtAlmacen_KeyDown 113, 0
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtalmacen.SelStart = 0: txtalmacen.SelLength = Len(txtalmacen.Text)

End Sub

Private Sub txtAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        'hlp_almacenes.Show 1
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            txtalmacen.Text = wcod_alm
            pnlalmacen.Caption = wnomalmacen
            
            If Rs.State = adStateOpen Then Rs.Close
            Rs.Open "select f2codalm,f2nomalm from ef2almacenes order by f2nomalm asc", cnn_dbbancos, adOpenStatic, adLockReadOnly
            x = 0
            If Not Rs.EOF Then
                Rs.MoveFirst
                cmbalmacen.Clear
                Do While Not Rs.EOF
                    cmbalmacen.AddItem Rs.Fields("f2nomalm") & "" & Space(50) & Rs.Fields("F2CODALM") & ""
                    Rs.MoveNext
                Loop
            End If
            Rs.Close
            
            habilita_conceptos txtalmacen.Text
            For i = 0 To cmbalmacen.ListCount - 1
                If txtalmacen.Text = Trim(right(cmbalmacen.List(i), 2)) Then
                    cmbalmacen.ListIndex = i
                    Exit For
                End If
            Next

            txtAlmacen_KeyPress 13
        End If
    End If

End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtAlmacen_LostFocus()

    'If sw_ayuda = True Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            If VALIDA_ALMACEN(txtalmacen.Text) = True Then
                cmbalmacen.Text = wnomalmacen & Space(50) & txtalmacen.Text
                dxDBGrid1.Enabled = True
            Else
                MsgBox "El Almacén no Existe", vbInformation, "Sistema de Logística"
                dxDBGrid1.Enabled = False
                cmbalmacen.ListIndex = -1
                txtalmacen.Text = ""
                txtalmacen.SetFocus
            End If
        Else
            cmbalmacen.ListIndex = -1
        End If
    'End If
    
End Sub

Private Sub txtalmacendes_Change()
    If Trim(txtalmacendes.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtccosto_Change()
    If Trim(txtccosto.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtconcepto_Change()
    If Trim(txtconcepto.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtconcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtNroOrdenProduccion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errTxtNroOrdenProduccion_KeyDown
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtnumero.Text) <> vbNullString Then
                Exit Sub
            End If
            
            'If cmbCategoriaTipo.ListIndex = -1 Then
            If Trim(lblIdCategoriaTipo.Caption) = vbNullString Then
                MsgBox "Seleccione la Categoria de la O.P.", vbInformation + vbOKOnly, App.ProductName
                
                cmbCategoriaTipo.SetFocus
                
                Exit Sub
            End If
            
            If Trim(txtNroOrdenProduccion.Text) = vbNullString Then
                MsgBox "Ingrese el Número de Orden de Producción.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            
            '------------------------------------------------------------------
            '------------------------------------------------------------------
            
            If ModMilano.verificarDatosOP(0, Trim(lblIdCategoriaTipo.Caption), Trim(txtNroOrdenProduccion.Text), txtIDOrdenProduccion) Then
                If ModMilano.importarOPServidorExternoV2(Trim(txtIDOrdenProduccion.Text), fraProceso, pgbProceso) Then
                    dxDBGrid1.Dataset.Close
                    
                    With frmUtilDescargaOrdenProduccion
                        .IdOrdenProduccion = Trim(txtIDOrdenProduccion.Text)
                        
                        .Caption = .Caption & ": " & Trim(cmbCategoriaTipo.Text) & " - " & txtNroOrdenProduccion.Text
                        
                        .Show vbModal
                        
                        If Not .cmbalmacen.Enabled Then
                            cmbalmacen.ListIndex = Val(ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2))
                            cmbconcepto.ListIndex = Val(ModUtilitario.seleccionarItem(cmbconcepto, "XDP", "DER", 3))
                            
                            If IsDate(.dtpFechaDespacho.value) Then
                                abofecha.value = .dtpFechaDespacho.value
                            End If
                        Else
                            cmbalmacen.ListIndex = -1
                            txtIDOrdenProduccion.Text = vbNullString
                            
                            If IsDate(.dtpFechaDespacho.value) Then
                                abofecha.value = .dtpFechaDespacho.value 'Format(Date, "Short Date")
                            End If
                        End If
                    End With
                    
                    Unload frmUtilDescargaOrdenProduccion
                    
                    dxDBGrid1.Dataset.Open
                    
                    If dxDBGrid1.Dataset.RecordCount = 0 Then
                        AdicionaItem
                    End If
                Else
                    
                    txtIDOrdenProduccion.Text = vbNullString
                End If
                
                ModUtilitario.pulsarTecla vbKeyTab
            Else
                cmbalmacen.ListIndex = 0
                cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XDP", "DER", 3)
            End If
            
            '------------------------------------------------------------------
            '------------------------------------------------------------------
            
'            txtIDOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDORDENPRODUCCION", "ORDENPRODUCCION", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T", "AND OP = '" & Trim(txtNroOrdenProduccion.Text) & "' AND ANULADO = 0")
'
'            If Trim(txtIDOrdenProduccion.Text) = vbNullString Then
'                MsgBox "O.P. no existe o esta anulada.", vbInformation + vbOKOnly, App.ProductName
'            Else
'                If MsgBox("¿Desea descargar la O.P.?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'                    If ModMilano.importarOPServidorExternoV2(Trim(txtIDOrdenProduccion.Text), fraProceso, pgbProceso) Then
'                        dxDBGrid1.Dataset.Close
'
'                        With frmUtilDescargaOrdenProduccion
'                            .IdOrdenProduccion = Trim(txtIDOrdenProduccion.Text)
'
'                            .Caption = .Caption & ": " & Trim(cmbCategoriaTipo.Text) & " - " & txtNroOrdenProduccion.Text
'
'                            .Show vbModal
'
'                            If Not .cmbalmacen.Enabled Then
'                                cmbalmacen.ListIndex = Val(ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2))
'                                cmbconcepto.ListIndex = Val(ModUtilitario.seleccionarItem(cmbconcepto, "XDP", "DER", 3))
'
'                                If IsDate(.dtpFechaDespacho.value) Then
'                                    abofecha.value = .dtpFechaDespacho.value
'                                End If
'                            Else
'                                cmbalmacen.ListIndex = -1
'                                txtIDOrdenProduccion.Text = vbNullString
'
'                                If IsDate(.dtpFechaDespacho.value) Then
'                                    abofecha.value = .dtpFechaDespacho.value
'                                End If
'                            End If
'                        End With
'
'                        Unload frmUtilDescargaOrdenProduccion
'
'                        dxDBGrid1.Dataset.Open
'
'                        If dxDBGrid1.Dataset.RecordCount = 0 Then
'                            AdicionaItem
'                        End If
'                    Else
'
'                        txtIDOrdenProduccion.Text = vbNullString
'                    End If
'                Else
'                    cmbalmacen.ListIndex = 0
'                    cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XDP", "DER", 3)
'
'                    'txtIDOrdenProduccion.Text = vbNullString
'                End If
'
'                ModUtilitario.pulsarTecla vbKeyTab
'            End If
            
            Screen.MousePointer = vbDefault
    End Select
    
    Exit Sub
errTxtNroOrdenProduccion_KeyDown:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Salida: TxtNroOrdenProduccion_KeyDown"
    
    Err.Clear
End Sub

Private Sub txtNroOrdenProduccion_LostFocus()
    If Trim(txtNroOrdenProduccion.Text) = vbNullString Then
        txtIDOrdenProduccion.Text = vbNullString
    End If
End Sub

Private Sub txtnumdoc_Change()
    If Trim(txtnumdoc.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub TxtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'cmbtipo.SetFocus
'End If
End Sub

Private Sub TxtNumDoc_LostFocus()
    txtnumdoc.Text = Format(txtnumdoc.Text, "0000000")
End Sub

Private Sub txtnumero_LostFocus()
    txtnumero.Text = Format(txtnumero.Text, "0000000")
End Sub

Private Sub txtnumfac_Change()
    If Trim(txtnumfac.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtnumfac_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtnumfac_LostFocus()
    txtnumfac.Text = Format(txtnumfac.Text, "0000000")
End Sub

Private Sub txtnumorden_Change()
    If Trim(txtnumorden.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtnumorden_DblClick()
    txtnumorden_KeyDown 113, 0
End Sub

Private Sub txtnumorden_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wllamada = 0
        wnumordentrab = "": wobservacion = ""
        ayuda_orden_trab.Show 1
        sw_ayuda = False
        If Len(Trim(wnumordentrab)) > 0 Then
            txtnumorden.Text = wnumordentrab
            txtnumorden_KeyPress 13
        End If
    End If

End Sub
Private Sub txtnumorden_LostFocus()
    
    If sw_ayuda = False Then
            If Len(Trim(txtnumorden.Text)) > 0 Then
                If rsordtrab.State = adStateOpen Then rsordtrab.Close
                rsordtrab.Open "SELECT Descripcion FROM dbo_partida WHERE CodPartida='" & txtnumorden.Text & "'", cnn_dbbancos
                If Not rsordtrab.EOF Then
                    pnlorden.Caption = Trim("" & rsordtrab.Fields("Descripcion"))
                Else
                    MsgBox "Código de la partida no existe. Verifique.", vbCritical, "Atención"
                    txtnumorden.SetFocus
                End If
                rsordtrab.Close
            Else
                pnlorden.Caption = ""
            End If
    End If

End Sub
Private Sub txtnumorden_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtserie.SetFocus
    End If

End Sub

Private Sub txtobserva_Change()
    If Trim(txtobserva.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtobserva_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
            
            dxDBGrid1.Columns.FocusedIndex = 3
    End Select
End Sub

Private Sub txtproveedor_DblClick()
    txtproveedor_KeyDown vbKeyF2, 0
End Sub

Private Sub txtproveedor_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtproveedor
End Sub

Private Sub txtproveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Me.MousePointer = vbHourglass
            
            wcodcliprov = vbNullString
            
            Select Case cmbTipoAuxiliar.ListIndex
                Case 0 'Clientes
                    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = False
                    
                    With ayuda_clientes
                        .Show 1
                    End With
                Case 1 'Proveedores
                    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = True
                    
                    With Ayuda_Proveedores
                        .Show 1
                    End With
                Case Else
                    MsgBox "Seleccione el Tipo de Persona.", vbInformation + vbOKOnly, App.ProductName
                    
                    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = True
                    
                    cmbTipoAuxiliar.SetFocus
                    
                    Me.MousePointer = vbDefault
                    
                    Exit Sub
            End Select
            
            If wcodcliprov <> vbNullString Then
                txtproveedor.Text = wcodcliprov
                txtnomprov.Text = wnomcliprov
            End If
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtproveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtproveedor_LostFocus()
    If Trim(txtproveedor.Text) <> vbNullString Then
        Select Case cmbTipoAuxiliar.ListIndex
            Case 0 'Clientes
                txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMCLI", "EF2CLIENTES", "F2CODCLI", Trim(txtproveedor.Text), "T")
            Case 1 'Proveedores
                txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
                
                If Trim(txtconcepto.Text) = "XNC" And Not ModUtilitario.validarFormAbierto("DevolucionCompra") Then
                    'SSActiveToolBars1_ToolClick 'SSActiveToolBars1.Tools("DevolucionCompra")
                End If
            Case Else
                MsgBox "Seleccione el Tipo de Persona.", vbInformation + vbOKOnly, App.ProductName
                
                cmbTipoAuxiliar.SetFocus
                
                Exit Sub
        End Select
    End If
End Sub

Private Sub txtserfac_Change()
    If Trim(txtserfac.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtserfac_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtserfac_LostFocus()
    txtserfac.Text = Format(txtserfac.Text, "000")
End Sub

Private Sub txtserie_Change()
    If Trim(txtserie.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
End Sub

Private Sub txtserie_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtnumdoc.SetFocus
'End If
End Sub

Private Sub txtserie_LostFocus()
    txtserie.Text = Format(txtserie.Text, "000")
End Sub

Private Sub txtusuario_Change()
    If Trim(txtusuario.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtusuario_DblClick()
    txtusuario_KeyDown 113, 0
End Sub

Private Sub txtusuario_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodusuario = "": wnomusuario = ""
        ayuda_usuarios.Show 1
        sw_ayuda = False
        If Len(Trim(wcodusuario)) > 0 Then
            txtusuario.Text = wcodusuario
            txtusuario_KeyPress 13
        End If
    End If

End Sub
Private Sub txtusuario_LostFocus()
    
    If sw_ayuda = False Then
            If Len(Trim(txtusuario.Text)) > 0 Then
                If rsusuarios.State = adStateOpen Then rsusuarios.Close
                rsusuarios.Open "SELECT F2NOMUSER FROM EF2USERS WHERE F2CODUSER='" & txtusuario.Text & "'", cnn_dbbancos
                If Not rsusuarios.EOF Then
                    pnlsolicitante.Caption = Trim("" & rsusuarios.Fields("F2NOMUSER"))
                Else
                    'MsgBox "Código de usuario no existe. Verifique.", vbCritical, "Atención"
                    'txtusuario.SetFocus
                End If
                rsusuarios.Close
            Else
                pnlsolicitante.Caption = ""
            End If
    End If

End Sub
Private Sub txtusuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtccosto.SetFocus
    End If

End Sub


Private Sub txtccosto_DblClick()

    txtccosto_KeyDown 113, 0
    
End Sub

Private Sub txtccosto_GotFocus()
    
    txtccosto.SelStart = 0: txtccosto.SelLength = Len(txtccosto.Text)
    
End Sub

Private Sub txtccosto_KeyDown(KeyCode As Integer, Shift As Integer)

'    If KeyCode = 113 Then
'
'    End If
    
    Select Case KeyCode
        Case vbKeyF2
            sw_ayuda = True
            wcodcosto = "": wdescosto = ""
            'hlp_centros.Show 1
            Ayuda_Centros.Show 1
            sw_ayuda = False
            If Len(Trim(wcodcosto)) > 0 Then
                txtccosto.Text = wcodcosto
                'txtccosto_KeyPress 13
            End If
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtccosto_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        'txtnumorden.SetFocus
'    End If

End Sub

Private Sub txtccosto_LostFocus()
    
    If sw_ayuda = False Then
        If Val(txttc.Text & "") > 0# Then
            If Len(Trim(txtccosto.Text)) > 0 Then
                If rsccosto.State = adStateOpen Then rsccosto.Close
                rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos
                If Not rsccosto.EOF Then
                    pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
                Else
                    MsgBox "Código de centro de costo no existe. Verifique.", vbCritical, "Atención"
                    txtccosto.SetFocus
                End If
                rsccosto.Close
            Else
                pnlccosto.Caption = ""
            End If
        End If
    End If

End Sub

Private Sub txtconcepto_DblClick()

    txtconcepto_KeyDown 113, 0
    
End Sub

Private Sub txtconcepto_GotFocus()

    txtconcepto.SelStart = 0: txtconcepto.SelLength = Len(txtconcepto.Text)
    
End Sub

Private Sub txtconcepto_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = 113 Then
''        sw_ayuda = True
''        wcod_alm = ""
''        'hlp_almacenes.Show 1
''        ayuda_almacen.Show 1
''        sw_ayuda = False
''        If Len(Trim(wcod_alm)) > 0 Then
''            txtalmacen.Text = wcod_alm
''            pnlalmacen.Caption = wnomalmacen
''
''            If rs.State = adStateOpen Then rs.Close
''            rs.Open "select f2codalm,f2nomalm from ef2almacenes order by f2nomalm asc", cnn_dbbancos, adOpenStatic, adLockReadOnly
''            X = 0
''            If Not rs.EOF Then
''                rs.MoveFirst
''                cmbalmacen.Clear
''                Do While Not rs.EOF
''                    cmbalmacen.AddItem rs.Fields("f2nomalm") & "" & Space(50) & rs.Fields("F2CODALM") & ""
''                    rs.MoveNext
''                Loop
''            End If
''            rs.Close
''
''            habilita_conceptos txtalmacen.Text
''            For i = 0 To cmbalmacen.ListCount - 1
''                If txtalmacen.Text = Trim(Right(cmbalmacen.List(i), 2)) Then
''                    cmbalmacen.ListIndex = i
''                    Exit For
''                End If
''            Next
''
''            txtalmacen_KeyPress 13
''        End If
''    End If
    
    If KeyCode = 113 Then
         sw_ayuda = True
         wconcepto = ""
         wtipmov = "S": wnomconcepto = ""
        ayuda_conceptos.Show 1
        sw_ayuda = False
            
        If Len(Trim(wconcepto)) > 0 Then
            habilita_conceptos txtalmacen.Text
            txtconcepto.Text = wconcepto
            For i = 0 To cmbconcepto.ListCount - 1
                If txtconcepto.Text = Trim(Mid(cmbconcepto.List(i), 200)) Then
                    cmbconcepto.ListIndex = i
                    Exit For
                End If
            Next

            txtconcepto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtconcepto_LostFocus()

            If Len(Trim(txtconcepto.Text)) > 0 Then
                If VALIDA_CONCEPTO_INV(txtconcepto.Text) = True Then
                    txtconcepto_KeyDown 113, 0
                    dxDBGrid1.Enabled = True
                Else
                    MsgBox "El Concepto no Existe", vbInformation, "Sistema de Logística"
                    dxDBGrid1.Enabled = False
                    cmbconcepto.ListIndex = -1
                    txtconcepto.Text = ""
                    txtconcepto.SetFocus
                End If
            Else
                cmbconcepto.ListIndex = -1
            End If

End Sub

Private Sub Txtnumfac_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        txtobserva.SetFocus
'    End If

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
''jcg        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'        dxDBGrid1.Columns.FocusedIndex = 1
'    End If
    
End Sub

Private Sub Txtserfac_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtnumfac.SetFocus
'    End If

End Sub

Private Sub txttc_Change()
    txttc.BackColor = vbWhite
    If Trim(txttc.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txttc_GotFocus()

    txttc.SelStart = 0:  txttc.SelLength = Len(txttc.Text)
    
End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txttc.Text = Format(txttc.Text, "0.000")
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Function GENERA_NUMVALE(palmacen As String, pmes As String, ptipo As String)
'Dim cnumvale    As String
'
'    If rsalmacen.State = adStateOpen Then rsalmacen.Close
'    rsalmacen.Open "SELECT * FROM EF2ALMACENES WHERE F2CODALM='" & palmacen & "'", cnn_dbbancos
'    If Not rsalmacen.EOF Then
'        If ptipo = "I" Then
'            cnumvale = Mid(rsalmacen.Fields("F1VALING" & pmes) & "", 1, 4) & Format(Val(Mid(rsalmacen.Fields("F1VALING" & pmes) & "", 5, 4)) + 1, "0000")
'        Else
'            cnumvale = Mid(rsalmacen.Fields("F1VALSAL" & pmes) & "", 1, 4) & Format(Val(Mid(rsalmacen.Fields("F1VALSAL" & pmes) & "", 5, 4)) + 1, "0000")
'        End If
'    End If
'    rsalmacen.Close
'
'    GENERA_NUMVALE = cnumvale
    With objAyudaVale
        .inicializarEntidades
        
        .CodigoAlmacen = Trim(txtalmacen.Text)
        .Fecha = Trim(abofecha.value)
        .TipoVale = ptipo
        
        GENERA_NUMVALE = .generarNumeroVale
    End With
End Function

Private Sub GRABA_SAL_ALMACEN_CENTRAL()
    Dim cnumvale        As String
    Dim ccampo          As String
    Dim nsoles          As Double
    Dim ndolar          As Double
    Dim rsdet           As New ADODB.Recordset
    Dim dbbase          As Database
    Dim tbbase          As Recordset
    Dim nitems          As Integer
    Dim cant_alm        As Double
    Dim existe          As Boolean
    
    If Trim(txtalmacen.Text) = "" Then
        MsgBox "Ingrese Almacen", vbInformation, "Sistema de Logistica"
        txtalmacen.SetFocus
        Exit Sub
    End If
    
    If Trim(txtconcepto.Text) = "" Then
        MsgBox "Ingrese Concepto", vbInformation, "Sistema de Logistica"
        'txtconcepto.SetFocus
        Exit Sub
    End If
    
'    If sw_nuevo_documento = True Then
'        If sw_ingreso = False Then
'            cnumvale = GENERA_NUMVALE(txtalmacen.Text, Format(Month(abofecha.Value), "00"), "S")
'            txtnumero.Text = cnumvale
'        Else
'            cnumvale = GENERA_NUMVALE(txtalmacendes.Text, Format(Month(abofecha.Value), "00"), "I")
'            'txtnumero.Text = cnumvale
'        End If
'        ctipo = "A"
'    Else
'        If sw_ingreso = False Then
'            cnumvale = txtnumero.Text
'        Else
'            If Rs.State = adStateOpen Then Rs.Close
'            Rs.Open "select valing from if4vales where f4numval = '" & txtnumero.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not Rs.EOF Then cnumvale = Rs.Fields("valing").Value & ""
'            Rs.Close
'        End If
'        ctipo = "M"
'    End If
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = IIf(Not sw_ingreso, "S", "I")
        .CodigoAlmacen = IIf(Not sw_ingreso, Trim(txtalmacen.Text), Trim(txtalmacendes.Text))
        .NumeroVale = IIf(Not sw_ingreso, Trim(txtnumero.Text), Trim(txtdestino.Text))
        .Fecha = abofecha.value
        
        If Not .verificarExistencia Then
            cnumvale = .generarNumeroVale
            
            txtnumero.Text = cnumvale
            ctipo = "A"
            
            .NumeroVale = cnumvale
        Else
            cnumvale = txtnumero.Text
            ctipo = "M"
        End If
        
        
    End With
    
    '-------------------------------------------------------
    '------------------------- ASIGNA DATOS DE LA CABECERA
    amovs_cab(0).campo = "F4NUMVAL": amovs_cab(0).valor = cnumvale: amovs_cab(0).Tipo = "T"
    If sw_ingreso = False Then
        amovs_cab(1).campo = "F2CODALM": amovs_cab(1).valor = txtalmacen.Text: amovs_cab(1).Tipo = "T"
    Else
        amovs_cab(1).campo = "F2CODALM": amovs_cab(1).valor = txtalmacendes.Text: amovs_cab(1).Tipo = "T"
    End If
    amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = abofecha.value: amovs_cab(2).Tipo = "F"
    If txtconcepto.Text = "XTA" And left(cnumvale, 1) = "I" Then
        amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = "XT3": amovs_cab(3).Tipo = "T"
    Else
        amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = txtconcepto.Text: amovs_cab(3).Tipo = "T"
    End If
    amovs_cab(4).campo = "F4TIPCAM": amovs_cab(4).valor = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", abofecha.value, "F")): amovs_cab(4).Tipo = "N"
    amovs_cab(5).campo = "F2CODPROV": amovs_cab(5).valor = txtproveedor.Text: amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "F4CENTRO": amovs_cab(6).valor = txtccosto.Text: amovs_cab(6).Tipo = "T"
    amovs_cab(7).campo = "F4MONEDA": amovs_cab(7).valor = "S": amovs_cab(7).Tipo = "T"
    amovs_cab(8).campo = "F4SERGUIA": amovs_cab(8).valor = txtserie.Text: amovs_cab(8).Tipo = "T"
    amovs_cab(9).campo = "F4NUMGUIA": amovs_cab(9).valor = txtnumdoc.Text: amovs_cab(9).Tipo = "T"
    amovs_cab(10).campo = "F4TIPDOC": amovs_cab(10).valor = right(cmbtipo.Text, 2): amovs_cab(10).Tipo = "T"
    amovs_cab(11).campo = "F4SERDOC": amovs_cab(11).valor = txtserfac.Text: amovs_cab(11).Tipo = "T"
    amovs_cab(12).campo = "F4NUMDOC": amovs_cab(12).valor = txtnumfac.Text: amovs_cab(12).Tipo = "T"
    
    If ctipo = "A" Then
        amovs_cab(13).campo = "F4FECGRA": amovs_cab(13).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(13).Tipo = "F"
        amovs_cab(14).campo = "F4USEGRA": amovs_cab(14).valor = wusuario: amovs_cab(14).Tipo = "T"
    Else
        amovs_cab(13).campo = "F4FECMOD": amovs_cab(13).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(13).Tipo = "F"
        amovs_cab(14).campo = "F4USEMOD": amovs_cab(14).valor = wusuario: amovs_cab(14).Tipo = "T"
    End If
    
    amovs_cab(15).campo = "F4OBSERVA": amovs_cab(15).valor = txtobserva.Text: amovs_cab(15).Tipo = "T"
    If sw_ingreso = False Then
        amovs_cab(16).campo = "CODALMDES": amovs_cab(16).valor = txtalmacendes.Text: amovs_cab(16).Tipo = "T"
    Else
        amovs_cab(16).campo = "CODALMDES": amovs_cab(16).valor = "": amovs_cab(16).Tipo = "T"
    End If
    amovs_cab(17).campo = "F2CODUSER": amovs_cab(17).valor = txtusuario.Text: amovs_cab(17).Tipo = "T"
    amovs_cab(18).campo = "NUMORDEN": amovs_cab(18).valor = txtnumorden.Text: amovs_cab(18).Tipo = "T"
    Rem SK ADD:
    amovs_cab(19).campo = "F4TIPOVALE": amovs_cab(19).valor = IIf(sw_ingreso, "I", "S"): amovs_cab(19).Tipo = "T"
    amovs_cab(20).campo = "EXPORTARVALE": amovs_cab(20).valor = IIf(CBool(chkExportarVale.value), -1, 0): amovs_cab(20).Tipo = "N"
    amovs_cab(21).campo = "F4ORDTRA": amovs_cab(21).valor = Trim(txtIDOrdenProduccion.Text): amovs_cab(21).Tipo = "T"
    amovs_cab(22).campo = "F1TIPPRV": amovs_cab(22).valor = right(cmbTipoAuxiliar.Text, 1): amovs_cab(22).Tipo = "T"
    

    
    '-------------------------------------------------------
    '------------------------- ASIGNA DATOS DEL DETALLE
    
    amovs_det(0).campo = "F4NUMVAL": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
    amovs_det(1).campo = "F5CODPRO": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
    amovs_det(2).campo = "F3CANPRO": amovs_det(2).valor = "": amovs_det(2).Tipo = "N"
    amovs_det(3).campo = "F3VALVTA": amovs_det(3).valor = "": amovs_det(3).Tipo = "N"
    amovs_det(4).campo = "F3IGV": amovs_det(4).valor = "": amovs_det(4).Tipo = "N"
    amovs_det(5).campo = "F3TOTITE": amovs_det(5).valor = "": amovs_det(5).Tipo = "N"
    amovs_det(6).campo = "F2CODALM": amovs_det(6).valor = "": amovs_det(6).Tipo = "T"
    amovs_det(7).campo = "F4FECVAL": amovs_det(7).valor = "": amovs_det(7).Tipo = "F"
    amovs_det(8).campo = "F3VALDOL": amovs_det(8).valor = "": amovs_det(8).Tipo = "N"
    amovs_det(9).campo = "F3IGVDOL": amovs_det(9).valor = "": amovs_det(9).Tipo = "N"
    amovs_det(10).campo = "F3TOTDOL": amovs_det(10).valor = "": amovs_det(10).Tipo = "N"
    amovs_det(11).campo = "TIPO": amovs_det(11).valor = "": amovs_det(11).Tipo = "T"
    amovs_det(12).campo = "F3PUNIT": amovs_det(12).valor = "": amovs_det(12).Tipo = "N"
    amovs_det(13).campo = "PARTIDA": amovs_det(13).valor = "": amovs_det(13).Tipo = "T"
    Rem SK ADD:
    amovs_det(14).campo = "F5CODPROORIGINAL": amovs_det(14).valor = "": amovs_det(14).Tipo = "T"
    amovs_det(15).campo = "F4NUMORD": amovs_det(15).valor = "": amovs_det(15).Tipo = "T"
    amovs_det(16).campo = "COD_SOLICITUD": amovs_det(16).valor = "": amovs_det(16).Tipo = "T"
    '------------------- CALCULA NUMERO DE FILAS
    
    nitems = 0
    
'    dxDBGrid1.Dataset.ADODataset.Requery
'    dxDBGrid1.Dataset.Edit
'    dxDBGrid1.Dataset.Post
'    dxDBGrid1.Dataset.Refresh
'    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
'    RSDETALLE.Open "SELECT COUNT(ITEM) AS NITEM FROM " & "TMPVALESALIDA" & " WHERE LEN(TRIM(CODPROD)) > 0 OR NOT ISNULL(CODPROD)", cnDBTemp, adOpenDynamic, adLockOptimistic
'    If Not RSDETALLE.EOF Then
'        nitems = Val("" & RSDETALLE.Fields("NITEM"))
'    End If
'    RSDETALLE.Close
    nitems = dxDBGrid1.Dataset.RecordCount
    '---------------------------------------------
    
    With objAyudaOrigen
        .inicializarEntidades
        
        .Codigo = Trim(txtconcepto.Text)
        
        .obtenerConfigOrigen
    End With
    
    'ReDim Values(13, nitems)
    ReDim Values(16, nitems)
    
'    If rsdet.State = adStateOpen Then rsdet.Close
'    rsdet.Open "SELECT * FROM " & "TMPVALESALIDA" & "", cnDBTemp, adOpenDynamic, adLockOptimistic
'    If Not rsdet.EOF Then
        nfil = 0
'        rsdet.MoveFirst
'        Do While Not rsdet.EOF
'            dxDBGrid1.Dataset.RecNo = nfil + 1
        dxDBGrid1.Dataset.First
        Do While Not dxDBGrid1.Dataset.EOF
            If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value & "")) > 0 Then
                Values(0, nfil) = cnumvale
                Values(1, nfil) = dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value
                Values(2, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & "")
                
                If objAyudaOrigen.RegistrarCosto Then
                    Values(3, nfil) = Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value, "0.00"))
                    Values(4, nfil) = 0
                    Values(5, nfil) = Val(Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value, "0.00")) * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & ""), "0.00"))
                    Values(8, nfil) = Val(Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value, "0.00")) / 1, "0.00"))
                    Values(9, nfil) = 0
                    Values(10, nfil) = Val(Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value, "0.00")) * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & "") / 1, "0.00"))
                Else
                    nsoles = 0: ndolar = 0
                    
                    With objAyudaVale
                        .CodigoProducto = Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value & "")
                        .Fecha = abofecha.value
                        
                        .CodigoMoneda = "S"
                        
                        nsoles = objAyudaVale.calcularCostoPromedio 'Costo_Unitario(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value & ""), abofecha.value, "S")
                        
                        .CodigoMoneda = "D"
                        
                        ndolar = objAyudaVale.calcularCostoPromedio 'Costo_Unitario(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value & ""), abofecha.value, "D")
                    End With
                    
                    Values(3, nfil) = nsoles
                    Values(4, nfil) = 0
                    Values(5, nfil) = Val(Format(nsoles * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & ""), "0.00"))
                    Values(8, nfil) = Val(Format(nsoles / 1, "0.00"))
                    Values(9, nfil) = 0
                    Values(10, nfil) = Val(Format(nsoles * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & "") / 1, "0.00"))
                End If
                
                If sw_ingreso = False Then
                    Values(6, nfil) = txtalmacen.Text
                Else
                    Values(6, nfil) = txtalmacendes.Text
                End If
                
                Values(7, nfil) = Format(abofecha.value, "dd/mm/yyyy")
                Values(11, nfil) = IIf(sw_ingreso, "I", "S")
                Values(12, nfil) = Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PUNIT").value, "0.00"))
                'Values(13, nfil) = dxDBGrid1.Columns.ColumnByFieldName("PARTIDA").Value
                
                Rem SK ADD:
                Values(14, nfil) = dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").value & "0"
                Values(15, nfil) = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value
                Values(16, nfil) = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value
                
                nfil = nfil + 1
            End If
            dxDBGrid1.Dataset.Next
            'rsdet.MoveNext
        Loop
'    Else
'    End If
'
'    rsdet.Close
    
    cvalores = "11111111111110111"
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    cmes = Format(Month(abofecha.value), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        'GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 18, cnn_dbbancos, ""
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 22, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            'GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 13, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 16, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
'        If sw_ingreso = False Then
'            ccampo = "F1VALSAL" & cmes
'            ACTUALIZA_ALMA_VALE cnumvale, ccampo, txtalmacen.Text
'        Else
'            ccampo = "F1VALING" & cmes
'            ACTUALIZA_ALMA_VALE cnumvale, ccampo, txtalmacendes.Text
'        End If
        
        
    Else    '--- Modificación
        '-------------------------------------------------------
        '------- GRABA CABECERA
        If sw_ingreso = False Then
            'GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 18, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtalmacen.Text & "'"
            GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 22, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtalmacen.Text & "'"
        Else
            'GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 18, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtalmacendes.Text & "'"
            GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 22, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtalmacendes.Text & "'"
        End If
        '-------------------------------------------------------
        '------- RESTA LOS SALDOS
        'If rsif3vales.State = adStateOpen Then rsif3vales.Close
        'rsif3vales.Open "SELECT * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        'If Not rsif3vales.EOF Then
        '    Do While Not rsif3vales.EOF
        '        GRABA_SALDO_ALM rsif3vales.Fields("F5CODPRO") & "", rsif3vales.Fields("F3CANPRO"), rsif3vales.Fields("F3TOTITE"), cmes, "E", cnn_dbbancos, txtalmacen.Text, rsif3vales.Fields("F3TOTDOL"), "R"
        '        rsif3vales.MoveNext
        '    Loop
        'End If
        'rsif3vales.Close
        '-------------------------------------------------------
        '------- GRABA DETALLE
        If sw_ingreso = False Then
            sql = ("DELETE FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            'GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 13, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 16, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
        Else
            sql = ("DELETE FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacendes.Text & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            'GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 13, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacendes.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 16, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacendes.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
        End If
    End If
    
'    For i = 0 To nfil - 1
'        variable = Values(1, i)
'        costo = Costo_Unitario2(variable, "S")(0)
'        Cant = Costo_Unitario2(variable, "S")(1)
'
'
'        If sw_ingreso = False Then
'            cant_alm = Calcula_Cantidad(variable, "S", txtAlmacen.Text)
'            csql = "UPDATE IF6ALMA SET F5COSPRO =  " & costo & ", F6STOCKACT = " & cant_alm & " WHERE F5CODPRO = '" & variable & "' and f2codalm ='" & txtAlmacen.Text & "'"
'            cnn_dbbancos.Execute csql
'            AlmacenaQuery_sql csql, cnn_dbbancos
'
'            csql = "UPDATE IF5PLA SET F5STOCKACT = " & Cant & " WHERE F5CODPRO = '" & variable & "'"
'            cnn_dbbancos.Execute csql
'            AlmacenaQuery_sql csql, cnn_dbbancos
'        Else
'            cant_alm = Calcula_Cantidad(variable, "S", txtalmacendes.Text)
'            existe = False
'            If Rs.State = adStateOpen Then Rs.Close
'            Rs.Open "select distinct f2codalm,f5codpro from if6alma where f2codalm = '" & txtalmacendes.Text & "' and f5codpro = '" & variable & "'", cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
'            Do While Not Rs.EOF
'                cant_alm = Calcula_Cantidad(variable, "S", txtalmacendes.Text)
'                csql = "UPDATE IF6ALMA SET F6STOCKACT = " & cant_alm & " WHERE F5CODPRO = '" & variable & "'and f2codalm ='" & txtalmacendes.Text & "'"
'                cnn_dbbancos.Execute csql
'                AlmacenaQuery_sql csql, cnn_dbbancos
'
'                existe = True
'                Rs.MoveNext
'            Loop
'            Rs.Close
'            If existe = False Then
'                csql = "insert into if6alma (f2codalm,f5codpro,f6stockact) values ('" & txtalmacendes.Text & "','" & variable & "','" & cant_alm & "')"
'                cnn_dbbancos.Execute csql
'                AlmacenaQuery_sql csql, cnn_dbbancos
'            End If
'            amovs_cab(0).campo = "valing": amovs_cab(0).valor = cnumvale: amovs_cab(0).Tipo = "T"
'            GRABA_REGISTRO_logistica amovs_cab, "if4vales", "M", 0, cnn_dbbancos, "f4numval = '" & txtnumero.Text & "'"
'
'            'cnn_dbbancos.Execute "update if4vales set valing = '" & cnumvale & "' where f4numval = '" & txtnumero.Text & "'"
'            'cnn_dbbancos.Execute "update if4vales set F4NumVal = '" & cnumvale & "' where f4numval = '" & txtnumero.Text & "'"
'
'        End If
'    Next
    
    If sw_ingreso = False Then
        MsgBox "Se ha Actualizado el Vale de Salida " & txtnumero.Text, vbInformation, "Sistema de Logística"
    Else
        MsgBox "Se ha Actualizado el Vale de Ingreso " & cnumvale, vbInformation, "Sistema de Logística"
    End If
    '-------------------------------------------------------
    '-------------------------------------------------------
End Sub

Private Sub ACTUALIZA_ALMA_VALE(pnumvale As String, pcampo As String, palmacen As String)
'Dim csql    As String
'
'    csql = "UPDATE EF2ALMACENES SET " & pcampo & " =  '" & pnumvale & "' WHERE '" & pnumvale & "' > " & pcampo & " AND F2CODALM='" & palmacen & "'"
'    cnn_dbbancos.Execute csql
'    AlmacenaQuery_sql csql, cnn_dbbancos
End Sub

Private Sub BUSCA_VALE(palmacen As String, pnumvale As String)
    Dim ncontador       As Long
    Dim cmedida         As String
    Dim i               As Integer
    Dim sw_nuevo_temp   As Boolean

    If rsif4vales.State = adStateOpen Then rsif4vales.Close
    
    rsif4vales.Open "SELECT * FROM IF4VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rsif4vales.EOF Then
        
        sw_nuevo_documento = False
        
        txtnumero.Text = rsif4vales.Fields("f4numval")
        txtserie.Text = "" & rsif4vales.Fields("f4serguia")
        txtnumdoc.Text = "" & rsif4vales.Fields("f4numguia")
        txtalmacen.Text = "" & rsif4vales.Fields("F2CODALM")
        txtdestino.Text = "" & rsif4vales.Fields("VALING")
        
        Rem SK ADD:
        'abrirCnDBMilano
        
        fraOrdenProduccion.Enabled = False
        txtIDOrdenProduccion.Text = Trim(rsif4vales!F4ORDTRA & "")
            If Trim(rsif4vales!F4ORDTRA & "") <> vbNullString Then
                cmbCategoriaTipo.ListIndex = ModUtilitario.seleccionarItem(cmbCategoriaTipo, ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", Trim(rsif4vales!F4ORDTRA & ""), "N"), "DER", 10)
                txtNroOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Trim(rsif4vales!F4ORDTRA & ""), "N")
            Else
                cmbCategoriaTipo.ListIndex = -1
                txtNroOrdenProduccion.Text = vbNullString
            End If
        
        lblNumeroValeExterno.Caption = Trim(rsif4vales!NUMENSAM & "")
        chkExportarVale.value = IIf(CBool(rsif4vales!ExportarVale), vbChecked, vbUnchecked): chkExportarVale.Enabled = False
        
        
        If VALIDA_ALMACEN(txtalmacen.Text) = True Then
            pnlalmacen.Caption = wnomalmacen
            cmbalmacen.Text = wnomalmacen & Space(50) & txtalmacen.Text
        End If
        
        For i = 0 To cmbalmacen.ListCount - 1
            If txtalmacen.Text = right(cmbalmacen.List(i), 2) Then
                cmbalmacen.ListIndex = i
                Exit For
            End If
        Next

        abofecha.value = Format(rsif4vales.Fields("F4FECVAL"), "DD/MM/YYYY")
        txtconcepto.Text = "" & rsif4vales.Fields("F1CODORI")
        If VALIDA_CONCEPTO_INV(txtconcepto.Text) = True Then
            pnlconcepto.Caption = wnomconcepto
        End If
        
        For i = 0 To cmbconcepto.ListCount - 1
            If txtconcepto.Text = Trim(Mid(cmbconcepto.List(i), 200)) Then
                cmbconcepto.ListIndex = i
                OCULTA_PRECIO
                Exit For
            End If
        Next

        For i = 0 To cmbalmacendes.ListCount - 1
            If rsif4vales.Fields("codalmdes").value = right(cmbalmacendes.List(i), 2) Then
                cmbalmacendes.ListIndex = i
                Exit For
            End If
        Next
        
        For i = 0 To cmbtipo.ListCount - 1
            If rsif4vales.Fields("f4tipdoc").value = right(cmbtipo.List(i), 2) Then
                cmbtipo.ListIndex = i
                Exit For
            End If
        Next
        
        txttc.Text = Format(rsif4vales.Fields("F4TIPCAM"), "0.000")
        txtobserva.Text = Trim("" & rsif4vales.Fields("F4OBSERVA"))
        
        cmbTipoAuxiliar.ListIndex = ModUtilitario.seleccionarItem(cmbTipoAuxiliar, Trim(rsif4vales!F1TIPPRV & ""), "DER", 1)
        
        txtproveedor.Text = Trim(rsif4vales!F2CODPROV & "")
        
        Select Case cmbTipoAuxiliar.ListIndex
            Case 0 'Clientes
                txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMCLI", "EF2CLIENTES", "F2CODCLI", Trim(txtproveedor.Text), "T")
            Case 1 'Proveedores
                txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
            Case Else
                txtnomprov.Text = vbNullString
        End Select
        
'        txtusuario.Text = "" & rsif4vales.Fields("F2CODUSER")
'        If rsusuarios.State = adStateOpen Then rsusuarios.Close
'        rsusuarios.Open "select f2nomuser from ef2users where f2coduser='" & txtusuario.Text & "'", cnn_dbbancos
'        If Not rsusuarios.EOF Then
'            pnlsolicitante.Caption = Trim("" & rsusuarios.Fields("f2nomuser"))
'        End If
'        rsusuarios.Close
        
        txtnumorden.Text = "" & rsif4vales.Fields("numorden")
        If Len(Trim(txtnumorden.Text)) > 0 Then
                If rsordtrab.State = adStateOpen Then rsordtrab.Close
                rsordtrab.Open "SELECT Descripcion FROM dbo_partida WHERE CodPartida='" & txtnumorden.Text & "'", cnn_dbbancos
                If Not rsordtrab.EOF Then
                    pnlorden.Caption = Trim("" & rsordtrab.Fields("Descripcion"))
                Else
                    MsgBox "Código de la partida no existe. Verifique.", vbCritical, "Atención"
                    txtnumorden.SetFocus
                End If
                rsordtrab.Close
            Else
                pnlorden.Caption = ""
            End If
        If rsordtrab.State = adStateOpen Then rsordtrab.Close
        rsordtrab.Open "select observacion from ordentrab_cab where numorden='" & txtnumorden.Text & "'", cnn_dbbancos
        If Not rsordtrab.EOF Then
            pnlorden.Caption = Trim("" & rsordtrab.Fields("observacion"))
        End If
        rsordtrab.Close
        
        txtccosto.Text = "" & rsif4vales.Fields("F4CENTRO")
        If rsccosto.State = adStateOpen Then rsccosto.Close
        rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos
        If Not rsccosto.EOF Then
            pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
        End If
        rsccosto.Close
            
        If "" & rsif4vales.Fields("F4MONEDA") = "S" Then
            cmbmoneda.ListIndex = 0
        Else
            cmbmoneda.ListIndex = 1
        End If
        
        If Trim(rsif4vales!F4TIPDOC & "") <> vbNullString Then
'            For i = 0 To cmbtipo.ListCount
'                If right(cmbtipo.Text, 2) = "" & rsif4vales.Fields("F4TIPDOC") Then
'                    cmbtipo.ListIndex = i
'                    Exit For
'                End If
'            Next
            cmbtipo.ListIndex = ModUtilitario.seleccionarItem(cmbtipo, Trim(rsif4vales!F4TIPDOC & ""), "DER", 2)
        Else
            cmbtipo.ListIndex = -1
        End If
        
        txtserfac.Text = "" & rsif4vales.Fields("F4SERDOC")
        txtnumfac.Text = "" & rsif4vales.Fields("F4NUMDOC")
        
        '--------------------------------------------------
        If sw_nuevo_documento = False Then
            DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
            AdicionaItem
            sw_nuevo_documento = True
        End If
        
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
        dxDBGrid1.Dataset.Active = True
    
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        
        Rem NSE dxDBGrid1.Option = egoSmartRefresh
        dxDBGrid1.OptionEnabled = False
        dxDBGrid1.Dataset.DisableControls
        With dxDBGrid1.Dataset
            
            If rsif3vales.State = adStateOpen Then rsif3vales.Close
            rsif3vales.Open "SELECT * FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsif3vales.EOF Then
                i = 1
                sw_nuevo_item = True
                rsif3vales.MoveFirst
                Do While Not rsif3vales.EOF
                    If sw_nuevo_temp = False Then
                        If sw_nuevo_documento = True Then
                            .Edit
                        Else
                            .Append
                        End If
                        sw_nuevo_temp = True
                    Else
                        .Append
                    End If
                    .FieldValues("ITEM") = i
                    .FieldValues("CODPROD") = rsif3vales.Fields("F5CODPRO") & ""
                    'If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    'rsif5pla.Open "SELECT F5NOMPRO,F7CODMED,F5CODFAB,F5MARCA FROM IF5PLA WHERE F5CODPRO='" & rsif3vales.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    'If Not rsif5pla.EOF Then
                    '    .FieldValues("UMEDIDA") = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsif5pla.Fields("F7CODMED"), "T", cnn_dbbancos)
                    '    .FieldValues("DESCRIPCION") = "" & rsif5pla.Fields("F5NOMPRO")
                    '    .FieldValues("CODFAB") = "" & rsif5pla.Fields("F5CODFAB")
                    '   '.FieldValues("MARCA") = "" & rsif5pla.Fields("F5MARCA")
                    'End If
                    'rsif5pla.Close
                    
                    .FieldValues("DESCRIPCION") = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(rsif3vales!f5codpro & ""), "T")
                    .FieldValues("CODFAB") = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODFAB", "IF5PLA", "F5CODPRO", Trim(rsif3vales!f5codpro & ""), "T")
                    .FieldValues("UMEDIDA") = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODMED", "IF5PLA", "F5CODPRO", Trim(rsif3vales!f5codpro & ""), "T")
                    .FieldValues("UMEDIDA") = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", Trim(.FieldValues("UMEDIDA") & ""), "T")
                    
                    .FieldValues("CANTIDAD") = Format(rsif3vales.Fields("F3CANPRO"), "###,###,##0.00")
                    .FieldValues("PUNIT") = Format(rsif3vales.Fields("F3VALVTA"), "###,###,##0.0000")
                    .FieldValues("TOTAL") = Val(Format(rsif3vales.Fields("F3CANPRO"), "0.00")) * Val(Format(rsif3vales.Fields("F3PUNIT"), "0.00"))
                    '.FieldValues("PARTIDA") = "" & rsif3vales.Fields("PARTIDA")
                    '.FieldValues("DESPARTIDA") = ObtenerCampo("dbo_partida", "Descripcion", "CodPartida", "" & rsif3vales.Fields("PARTIDA"), "T", cnn_dbbancos)
                    
                    Rem SK ADD:
                    .FieldValues("CODPRODORIGINAL") = rsif3vales.Fields("F5CODPROORIGINAL") & ""
                    .FieldValues("F4NUMORD") = rsif3vales.Fields("F4NUMORD") & ""
                    .FieldValues("COD_SOLICITUD") = rsif3vales.Fields("COD_SOLICITUD") & ""
                    
                    rsif3vales.MoveNext
                    i = i + 1
                Loop
                .Post
                sw_nuevo_item = False
            End If
            rsif3vales.Close
            
        End With
        dxDBGrid1.Dataset.EnableControls
        'dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        Rem NSE dxDBGrid1.Dataset.Refresh
        dxDBGrid1.OptionEnabled = True
        '--------------------------------------------------
        '--------------------------------------------------
                
    Else
        '----- No existe la guía
        
        sw_nuevo_documento = False
        nuevo
        AdicionaItem
        AdicionaItem
        sw_nuevo_documento = True
        
    End If
    rsif4vales.Close
    
    Rem SK ADD:
    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = False
End Sub

Private Sub elimina(pnumvale As String, palmacen As String)
On Error GoTo ERROR_ELIMINA
ReDim amovs(0 To 0) As a_grabacion
Dim cmes            As String * 2
Dim regAfec         As Integer

    If Len(Trim("" & txtnumero.Text)) = 0 Then
        MsgBox "El Vale de Salida no ha sido grabado. Verifique", vbCritical, "Atención"
        Exit Sub
    End If
    
    If MsgBox("Está seguro(a) de eliminar el Vale de Salida ?", vbYesNo + vbQuestion, "Atención") = vbYes Then
        
        Rem SK ADD:
        If Val(lblNumeroValeExterno.Caption) > 0 Then
'            If ModMilano.anularValeExterno("S", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T"), fraProceso, pgbProceso) Then
'                Me.MousePointer = vbDefault
'
'                Exit Sub
'            End If
        End If
        
        
        sql = ("DELETE * FROM IF4VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        
        sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        
        
        
        If wIndEnvia = "*" Then
            sql = "delete from if4vales where f2codalm='" & palmacen & "' and f4numval='" & pnumvale & "'"
            cnn_dbEnvia.Execute sql, regAfec
             AlmacenaQuery_sql sql, cnn_dbbancos
            If regAfec = 0 Then
                'vales
                'guardar el almacen,numvale en una tabla(de vales eliminados)
                'al recibir leer de esa tabla y ejecutar la setencia con los datos
                sql = "insert into VALESELIMINADOS(F2CODALM, F4NUMVAL) " & _
                                    " values('" & palmacen & "','" & pnumvale & "')"
                cnn_dbEnvia.Execute sql
                AlmacenaQuery_sql sql, cnn_dbbancos
                
            End If
        End If
        
        'nuevo
        'AdicionaItem
        
        If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4NUMVAL", "IF4VALES", "F4NUMVAL", Trim(txtnumero.Text), "T", "AND F2CODALM = '" & Trim(txtalmacen.Text) & "'") = vbNullString Then
            sw_nuevo_documento = True
            nuevo
            dxDBGrid1.Dataset.Close
            DELETEREC_LOG "TMPVALESALIDA", cnDBTemp
            AdicionaItem
        End If
    End If
    
    Exit Sub
    
ERROR_ELIMINA:
    MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atención"
    Resume Next
    
End Sub
Private Sub cmbAlmacen_Click()
    txtalmacen.Text = right(cmbalmacen.Text, 2)
'    If Val(wcod_alm) = Val(txtalmacen.Text) Then
        wcod_alm = txtalmacen.Text
        habilita_conceptos txtalmacen.Text
'    End If
End Sub

Private Sub habilita_almacen_destino(codori As String, codalm As String)
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    cmbalmacendes.Clear
    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES WHERE F2CODALM NOT IN ('" & codalm & "')", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            cmbalmacendes.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(75) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
    End If
    
'    cmbalmacendes.Clear
'
'    Dim sql As String
'
'    sw_habil = False
'
    cmbalmacendes.Enabled = True
'    cmbalmacendes.Clear
'
'    If rsconcepto_inv.State = adStateOpen Then rsconcepto_inv.Close
'
'    rsconcepto_inv.Open "select codalmdes from sf1origenes where f1codori = '" & codori & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'
'    If Not rsconcepto_inv.EOF Then
'        If rsconcepto_inv.Fields("codalmdes").value = "99" Then
'            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
'            RSCONSULTA.Open "select f2codalm,f2nomalm from ef2almacenes where f2codalm <>'" & codalm & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not RSCONSULTA.EOF Then
'                RSCONSULTA.MoveFirst
'                Do While Not RSCONSULTA.EOF
'                    cmbalmacendes.AddItem RSCONSULTA.Fields("f2nomalm") & "" & Space(75) & RSCONSULTA.Fields("f2codalm") & ""
'                    RSCONSULTA.MoveNext
'                Loop
'                sw_habil = True
'            End If
'            RSCONSULTA.Close
'        End If
'    End If
'    rsconcepto_inv.Close
'
'    If Rs.State = adStateOpen Then Rs.Close
'    sql = "SELECT EF2ALMACENES.F2CODALM, EF2ALMACENES.F2NOMALM " & _
'          "FROM EF2ALMACENES INNER JOIN SF1ORIGENES ON EF2ALMACENES.F2CODALM = " & _
'          "SF1ORIGENES.CODALMDES where SF1ORIGENES.f1codori = '" & codori & "';"
'
'    Rs.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        Do While Not Rs.EOF
'            cmbalmacendes.AddItem Rs.Fields("f2nomalm") & "" & Space(75) & Rs.Fields("F2CODALM") & ""
'            Rs.MoveNext
'        Loop
'        sw_habil = True
'    End If
'    Rs.Close
    
End Sub

Private Sub habilita_conceptos(codalma As String)
    cmbconcepto.Clear
    txtconcepto.Text = ""
    cmbalmacendes.Visible = False
    txtalmacendes.Visible = False
    Label6.Visible = False
    
    If Rs.State = adStateOpen Then Rs.Close
    
    'sql = "SELECT SF1ORIGENES.F1CODORI, SF1ORIGENES.F1NOMORI " & _
          "FROM ALMACEN_CONCEPTO INNER JOIN SF1ORIGENES ON ALMACEN_CONCEPTO.F1CODORI = " & _
          "SF1ORIGENES.F1CODORI where ALMACEN_CONCEPTO.f2codalm = '" & codalma & "';"
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "ORI.F1CODORI, "
    SqlCad = SqlCad & "ORI.F1NOMORI "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ALMACEN_CONCEPTO AS AC "
    SqlCad = SqlCad & "LEFT JOIN SF1ORIGENES AS ORI "
    SqlCad = SqlCad & "ON ORI.F1CODORI = AC.F1CODORI "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "AC.F2CODALM = '" & codalma & "' AND "
    SqlCad = SqlCad & "ORI.F1TIPMOV = 'S' AND "
    SqlCad = SqlCad & "ORI.F1CODORI NOT IN ('XCS') "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "ORI.F1NOMORI "
    
    'Rs.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
    Rs.Open SqlCad, cnn_dbbancos, adOpenStatic, adLockReadOnly
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        Do While Not Rs.EOF
            cad = Space(255)
            Mid(cad, 1) = "" & Rs.Fields("f1nomori")
            Mid(cad, 200) = "" & Rs.Fields("f1codori")
            
            cmbconcepto.AddItem cad
            
            Rs.MoveNext
        Loop
    End If
    
    Rs.Close
End Sub




'-------------------------------------------------------------------------------------------------------------

Private Sub listarAlmacenEnCombo()
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    cmbalmacen.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            cmbalmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If cmbalmacen.ListCount > 0 Then
                cmbalmacen.ListIndex = 0
            End If
    End If
End Sub

Private Sub listarTipoDocumentoEnCombo()
    Dim rstDocumento As New ADODB.Recordset
    
    If rstDocumento.State = 1 Then rstDocumento.Close
    
    rstDocumento.Open "SELECT F2CODDOC, F2DESDOC FROM DOCUMENTOS WHERE TRIM(CODEXT3 & '') <> '' ORDER BY F2DESDOC", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    cmbtipo.Clear
    
    If Not rstDocumento.EOF Then
        rstDocumento.MoveFirst
        
        Do While Not rstDocumento.EOF
            cmbtipo.AddItem Trim(rstDocumento!F2DESDOC & "") & Space(100) & Trim(rstDocumento!F2CODDOC & "")
            
            rstDocumento.MoveNext
        Loop
            If cmbtipo.ListCount > 0 Then
                cmbtipo.ListIndex = 0
            End If
    End If
End Sub

Private Sub listarGrillaVale()
    With dxDBGrid1.Dataset
        abrirCnTemporal
        
        .Active = False
        .Refresh
        
        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        .ADODataset.CommandText = "SELECT * FROM TMPVALESALIDA ORDER BY ITEM, DESCRIPCION"
        .Active = False
        .Active = True
        
        dxDBGrid1.KeyField = "ITEM"
        
        .Close
        .Open
    End With
End Sub

Private Sub adicionarItemVale()
    With dxDBGrid1.Dataset
        .Close
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE * FROM TMPVALESALIDA"
        
'        abrirCnTemporal
'
'        cnDBTemp.Execute "INSERT INTO TMPVALESALIDA(ITEM, CODPROD, CODPRODORIGINAL) VALUES(1, NULL, NULL)"

        .Active = False

        .Refresh

        abrirCnTemporal

        .ADODataset.ConnectionString = cnDBTemp

        .Active = True
        .Close
        .Open

        If .State = dsEdit Or .State = dsInsert Then
            .Append
        Else
            .Edit
        End If
        
        .FieldValues("ITEM") = .RecordCount + 1
        .FieldValues("CODPROD") = vbNullString
        .FieldValues("CODPRODORIGINAL") = vbNullString
        
        .Post
        
        .Close
        .Open
    End With
End Sub

Private Sub renumerarItemVale()
    Dim rstTemporalRenumerarS As New ADODB.Recordset
    Dim dblItem As Double
    
    If rstTemporalRenumerarS.State = 1 Then rstTemporalRenumerarS.Close
    
    rstTemporalRenumerarS.Open "SELECT * FROM TMPVALESALIDA ORDER BY ITEM, DESCRIPCION", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporalRenumerarS.EOF Then
        rstTemporalRenumerarS.MoveFirst
        
        dblItem = 0
        
        'dxDBGrid1.Dataset.Close
        
        Do While Not rstTemporalRenumerarS.EOF
            dblItem = dblItem + 1
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPVALESALIDA "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "ITEM = " & dblItem & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstTemporalRenumerarS!F4NUMORD & "") & "' AND "
            SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstTemporalRenumerarS!COD_SOLICITUD & "") & "' AND "
            SqlCad = SqlCad & "TRIM(CODPROD & '') = '" & Trim(rstTemporalRenumerarS!codprod & "") & "' AND "
            SqlCad = SqlCad & "TRIM(CODPRODORIGINAL & '') = '" & Trim(rstTemporalRenumerarS!CODPRODORIGINAL & "") & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            rstTemporalRenumerarS.MoveNext
        Loop
            'dxDBGrid1.Dataset.Open
    End If
    
    If rstTemporalRenumerarS.State = 1 Then rstTemporalRenumerarS.Close
    
    Set rstTemporalRenumerarS = Nothing
    
    dblItem = 0
End Sub

Private Sub limpiarCajas()
    Dim strFichero As String
    
    Me.Caption = "Vale de Salida de Almacen"
    
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    txtnumero.Text = vbNullString
        lblNumeroValeExterno.Caption = "< ID Externo >"
    
    cmbalmacen.ListIndex = -1
        cmbalmacen.Enabled = True
        txtalmacen.Text = vbNullString
    cmbconcepto.ListIndex = -1
        txtconcepto.Text = vbNullString
    
    cmbTipoAuxiliar.ListIndex = -1
        txtproveedor.Text = vbNullString: txtnomprov.Text = vbNullString
    txtccosto.Text = vbNullString: pnlccosto.Caption = vbNullString
    txtserie.Text = vbNullString: txtnumdoc.Text = vbNullString
    
    cmbtipo.ListIndex = -1
    txtserfac.Text = vbNullString: txtnumfac.Text = vbNullString
    dtpFechaDoc.value = Format(Date, "Short Date")
    
    fraOrdenProduccion.Enabled = True
    cmbCategoriaTipo.ListIndex = -1
    cmbCategoriaTipo.Text = vbNullString
    lblIdCategoriaTipo.Caption = vbNullString
    txtNroOrdenProduccion.Text = vbNullString
    txtIDOrdenProduccion.Text = vbNullString
    
    'abofecha.value = Date
    
    With abofecha
        .value = Format(Date, "Short Date")
        .CalendarBackColor = vbWhite
        .CalendarForeColor = vbBlack
        .Font.Bold = False
        
        lblFechaMensaje.Visible = False
        
        If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaUsarFechaPredeterminada", "l") = "1" Then
            .value = ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaFechaPredeterminada", "l")
            .CalendarBackColor = vbRed
            .CalendarForeColor = vbWhite
            .CalendarTrailingForeColor = vbGreen
            .Font.Bold = True
            
            lblFechaMensaje.Visible = True
        End If
    End With
    
    
    cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, "S", "IZQ", 1)
    txttc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", abofecha.value, "F")), "#.000")
    
    If Val(txttc.Text) = 0 Then
        txttc.Text = "1"
    End If
    
    txtobserva.Text = vbNullString
    
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
    SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
    SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
    SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
    'SSActiveToolBars1.Tools("DevolucionCompra").Enabled = False
    
    bolObviarCierre = True
    
    SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
    SSActiveToolBars1.Tools.ITEM("CerrarVale").Enabled = False
    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
    
    bolObviarCierre = False
    
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA"
    
    SSFrame1.Enabled = True
    dxDBGrid1.Enabled = True
End Sub

Private Sub consultarVale()
    On Error GoTo errConsultarVale
    
    Set objVale = New ClsVale
    
    limpiarCajas
    
    dxDBGrid1.Dataset.Close
    
    With objVale
        .inicializarEntidades
        
        .CodigoAlmacen = strCodAlmacen
        .NumeroVale = strNumeroVale
        
        If .obtenerVale Then
            txtnumero.Text = .NumeroVale
                lblNumeroValeExterno.Caption = .NumeroValeExterno
            
            cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, .CodigoAlmacen, "DER", 2): cmbalmacen.Enabled = False
                txtalmacen.Text = .CodigoAlmacen
            cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, .CodigoOrigen, "DER", 3)
                txtconcepto.Text = .CodigoOrigen
            
            cmbTipoAuxiliar.ListIndex = ModUtilitario.seleccionarItem(cmbTipoAuxiliar, .TipoPersona, "DER", 1)
                txtproveedor.Text = .CodigoProveedor
                
                Select Case .TipoPersona
                    Case "C"
                        txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMCLI", "EF2CLIENTES", "F2CODCLI", .CodigoProveedor, "T")
                    Case "P"
                        txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                End Select
                
            txtccosto.Text = .CentroCosto
                pnlccosto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", .CentroCosto, "T")
            txtserie.Text = .SerieGuia: txtnumdoc.Text = .NumeroGuia
            
            cmbtipo.ListIndex = ModUtilitario.seleccionarItem(cmbtipo, .CodTipoComprobante, "DER", 2)
            txtserfac.Text = .SerieDocumento: txtnumfac.Text = .NumeroDocumento
            dtpFechaDoc.value = IIf(.FechaUltima <> vbNullString, .FechaUltima, .Fecha)
            
            'abrirCnDBMilano
            
            fraOrdenProduccion.Enabled = False
            txtIDOrdenProduccion.Text = .OrdenTrabajo
                If .OrdenTrabajo <> vbNullString Then
                    'cmbCategoriaTipo.ListIndex = ModUtilitario.seleccionarItem(cmbCategoriaTipo, ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N"), "DER", 10)
                    lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                    'cmbCategoriaTipo.ListIndex = ModUtilitario.seleccionarItem(cmbCategoriaTipo, ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N"), "DER", 10)
                    cmbCategoriaTipo.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T")
                    txtNroOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                End If
            
            abofecha.value = .Fecha
            cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, .CodigoMoneda, "IZQ", 1)
            txttc.Text = Format(.TipoCambio, "#.000")
            
            txtobserva.Text = .observaciones
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) > 0 Then
                renumerarItemVale
            End If
            
            listarGrillaVale
            
            If dxDBGrid1.Dataset.RecordCount = 0 Then
                dxDBGrid1.Dataset.Close
                
                adicionarItemVale
            End If
            
            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
            
            If .VB1 Then
                Me.Caption = Me.Caption & " ( VALE CERRADO )"
            End If
            
            bolObviarCierre = True
            
            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = IIf(.VB1, ssChecked, ssUnchecked)
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0015'") = "0015" Then
                SSActiveToolBars1.Tools.ITEM("CerrarVale").Enabled = True
            End If
            
            If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
            Else
                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
            End If
            
            bolObviarCierre = False
            
            Rem SK ADD:
            If .VB1 Then
                SSFrame1.Enabled = False
                dxDBGrid1.Enabled = False
                
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
                'SSActiveToolBars1.Tools("ID_OC").Enabled = False
            End If
        Else
            listarGrillaVale
            
            adicionarItemVale
        End If
    End With
    
    Set objVale = Nothing
    
    Exit Sub
errConsultarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Salida: ConsultarVale"
    
    Err.Clear
End Sub

Private Sub validarCajas()
    On Error GoTo errValidarCajas
    
    'Validacion del Punto (PC) que origina el Vale
    'ModMilano.abrirCnDBMilano
    
'    If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'        MsgBox "Su computador no esta registrado y/o habilitado. Consulte con su" & vbNewLine & vbNewLine & _
'                "administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        Exit Sub
'    End If
'
'    If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")), "T", _
'                                                                    "AND TABLA = 'SALIDA'")) = 0 Then
'
'        MsgBox "El Punto de Venta no cuenta con correlativo habilitado de SALIDA." & vbNewLine & vbNewLine & _
'                "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        Exit Sub
'    End If
    
    If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
        dxDBGrid1.Dataset.Post
    End If
    
    If cmbalmacen.ListIndex = -1 Then
        MsgBox "El Campo Almacén es obligatorio.", vbInformation + vbOKOnly, App.ProductName
        
        cmbalmacen.SetFocus
        
        Exit Sub
    End If
    
'    If ModMilano.verificarCierreDeMesEnServidorExterno(Year(CDate(abofecha.value)), Month(CDate(abofecha.value)), Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T"))) Then
'        MsgBox "Imposible registrar Vale, el mes se encuentra cerrado. Verifique la fecha" & vbNewLine & vbNewLine & _
'                "seleccionada o consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        abofecha.SetFocus
'
'        Exit Sub
'    End If
    
    
    
    
    If cmbconcepto.ListIndex = -1 Then
        MsgBox "El Campo Concepto es obligatorio.", vbInformation + vbOKOnly, App.ProductName
        
        cmbconcepto.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoAuxiliar.ListIndex = -1 And (Trim(txtproveedor.Text) <> vbNullString And Trim(txtnomprov.Text) <> vbNullString) Then
        MsgBox "Seleccione el Tipo de Persona.", vbInformation + vbOKOnly, App.ProductName
        
        cmbTipoAuxiliar.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtproveedor.Text) <> vbNullString And (txtnomprov.Text) = vbNullString Then
        MsgBox "Proveedor ingresado no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        txtproveedor.SetFocus
        
        Exit Sub
    End If
    
    'If Val(txttc.Text) <= 0 Then
    '    MsgBox "Tipo de Cambio incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName

        'txttc.SetFocus
        
    '    Exit Sub
    'End If
    
    If Trim(txtconcepto.Text) = "XDP" And Trim(txtIDOrdenProduccion.Text) = vbNullString Then
        MsgBox "ID de Orden de Produccion incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        cmbCategoriaTipo.SetFocus
        
        Exit Sub
    End If

    If Trim(txtconcepto.Text) <> "XDP" And Trim(txtIDOrdenProduccion.Text) <> vbNullString Then
        MsgBox "Concepto incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName

        cmbconcepto.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtconcepto.Text) = "X3R" Then
        If Trim(txtnumfac.Text) = vbNullString Then
            MsgBox "Ingrese el Nro. de Guia Interna.", vbInformation + vbOKOnly, App.ProductName
            
            txtnumfac.SetFocus
            
            Exit Sub
        End If
    End If
    
    If CDate(abofecha.value) > CDate(Date) Then
        MsgBox "Fecha ingresada inválida, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        abofecha.SetFocus
        
        Exit Sub
    End If
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) = 0 Then
        MsgBox "Registro no cuenta con Detalle, verifique.", vbInformation + vbOKOnly, App.ProductName

        dxDBGrid1.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                .CodigoAlmacen = Trim(txtalmacen.Text)
                
                .FechaInicioMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 0, 1)
                .FechaFinMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 1, 0)
                
                If .verificarCierreVale Then
                    MsgBox "Imposible registrar Vale, periodo ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                    
                    .inicializarEntidades
                    .inicializarEntidadesAdicionales
                    
                    Exit Sub
                End If
            End With
        Else
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                .CodigoAlmacen = Trim(txtalmacen.Text)
                
                .FechaInicioMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 0, 1)
                .FechaFinMes = DateSerial(Year(CDate(abofecha.value)), Val(Month(CDate(abofecha.value))) + 1, 0)
                
                If .verificarCierreVale Then
                    MsgBox "Imposible registrar Vale, periodo ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                    
                    .inicializarEntidades
                    .inicializarEntidadesAdicionales
                    
                    Exit Sub
                End If
            End With
        End If
        
        'guardarVale
        grabar
    End If
    
    Exit Sub
    Resume
errValidarCajas:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Salida: ValidarCajas"
    
    Err.Clear
End Sub

Private Sub guardarVale()
    On Error GoTo errGuardarVale
    
    Dim rstTmpValeDetalleS As New ADODB.Recordset
    Dim dblItem As Double
    Dim dblItemNoGuardado As Double
    Dim strDescripcionItemNoGuardado As String
    
    Set objVale = New ClsVale
    
    dxDBGrid1.Dataset.Close
    
    With objVale
        .inicializarEntidades
        
        .TipoVale = "S"
        .NumeroVale = Trim(txtnumero.Text)
        .NumeroValeExterno = Trim(lblNumeroValeExterno.Caption)
        
        .CodigoAlmacen = Trim(txtalmacen.Text)
        .CodigoOrigen = Trim(txtconcepto.Text)
        
        .TipoPersona = right(cmbTipoAuxiliar.Text, 1)
            .CodigoProveedor = Trim(txtproveedor.Text)
            
        .CentroCosto = Trim(txtccosto.Text)
        .SerieGuia = Trim(txtserie.Text)
        .NumeroGuia = Trim(txtnumdoc.Text)
        
        .CodTipoComprobante = right(cmbtipo.Text, 2)
        .SerieDocumento = Trim(txtserfac.Text)
        .NumeroDocumento = Trim(txtnumfac.Text)
        
        If .CodTipoComprobante <> vbNullString And .NumeroDocumento <> vbNullString Then
            .FechaUltima = Format(dtpFechaDoc.value, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = Trim(txtIDOrdenProduccion.Text)
        
        .Fecha = Format(abofecha.value, "Short Date")
        .CodigoMoneda = left(cmbmoneda.Text, 1)
        .TipoCambio = Val(txttc.Text)
        
        .observaciones = Trim(txtobserva.Text)
        
        .ExportarVale = CBool(chkExportarVale.value)
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        Me.MousePointer = vbHourglass
        
        If .guardarVale Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            'If Trim(txtNumero.Text) <> vbNullString Then
                .SQLSelectAlter = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                cnn_dbbancos.Execute .SQLSelectAlter
                
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                
                    cnBdCPlus.Execute .SQLSelectAlter
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            'End If
            
            'If rstTmpValeDetalleS.State = 1 Then rstTmpValeDetalleS.Close
            
            'rstTmpValeDetalleS.Open "SELECT * FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') <> '' AND CANTIDAD > 0", cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "F4NUMORD, "
            SqlCad = SqlCad & "COD_SOLICITUD, "
            SqlCad = SqlCad & "CODPROD, "
            SqlCad = SqlCad & "CODPRODORIGINAL, "
            SqlCad = SqlCad & "DESCRIPCION, "
            SqlCad = SqlCad & "SUM(CANTIDAD) AS CANTIDAD, "
            SqlCad = SqlCad & "PUNIT "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "TMPVALESALIDA "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "TRIM(CODPROD & '') <> '' AND "
            SqlCad = SqlCad & "CANTIDAD > 0 "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "F4NUMORD, "
            SqlCad = SqlCad & "COD_SOLICITUD, "
            SqlCad = SqlCad & "CODPROD, "
            SqlCad = SqlCad & "CODPRODORIGINAL, "
            SqlCad = SqlCad & "DESCRIPCION, "
            SqlCad = SqlCad & "PUNIT"
            
            abrirCnTemporal
            
            If rstTmpValeDetalleS.State = 1 Then rstTmpValeDetalleS.Close
            
            rstTmpValeDetalleS.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTmpValeDetalleS.EOF Then
                rstTmpValeDetalleS.MoveFirst
                
                With objAyudaOrigen
                    .inicializarEntidades
                    
                    .Codigo = objVale.CodigoOrigen
                    
                    .obtenerConfigOrigen
                End With
                
                dblItem = 0
                dblItemNoGuardado = 0
                strDescripcionItemNoGuardado = vbNullString
                
                fraProceso.Visible = True
                pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstTmpValeDetalleS)
                pgbProceso.value = 0
                fraProceso.Caption = "Guardando Detalle..."
                
                Do While Not rstTmpValeDetalleS.EOF
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstTmpValeDetalleS!codprod & "")
                    .CodigoProductoOriginal = Trim(rstTmpValeDetalleS!CODPRODORIGINAL & "")
                    .Cantidad = Val(rstTmpValeDetalleS!Cantidad & "")
                    
                    If objAyudaOrigen.RegistrarCosto Then
                        Select Case .CodigoMoneda
                            Case "S"
                                .ValorVenta = Val(rstTmpValeDetalleS!PUNIT & "")
                                .IGV = 0
                                .TOTAL = Val(rstTmpValeDetalleS!PUNIT & "") * Val(rstTmpValeDetalleS!Cantidad & "")
                                
                                .ValorVentaDol = Val(rstTmpValeDetalleS!PUNIT & "") / .TipoCambio
                                .IgvDol = 0
                                .TotalDol = (Val(rstTmpValeDetalleS!PUNIT & "") * Val(rstTmpValeDetalleS!Cantidad & "")) / .TipoCambio
                            Case Else
                                .ValorVenta = Val(rstTmpValeDetalleS!PUNIT & "") * .TipoCambio
                                .IGV = 0
                                .TOTAL = (Val(rstTmpValeDetalleS!PUNIT & "") * Val(rstTmpValeDetalleS!Cantidad & "")) * .TipoCambio
                                
                                .ValorVentaDol = Val(rstTmpValeDetalleS!PUNIT & "")
                                .IgvDol = 0
                                .TotalDol = Val(rstTmpValeDetalleS!PUNIT & "") * Val(rstTmpValeDetalleS!Cantidad & "")
                        End Select
                    Else
                        objAyudaVale.CodigoAlmacen = .CodigoAlmacen
                        objAyudaVale.CodigoMoneda = .CodigoMoneda
                        objAyudaVale.CodigoProducto = .CodigoProducto
                        objAyudaVale.Fecha = .Fecha
                        
                        .ValorVenta = objAyudaVale.calcularCostoPromedio
                        
'                        If .ValorVenta <= 0 Then
'                            .ValorVenta = objAyudaVale.obtenerUltimoCostoPromedio
'                        End If
                        
                        .IGV = 0
                        .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                        
                        .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                        .IgvDol = 0
                        .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    End If
                    
                    .NumeroOrdenCompra = Trim(rstTmpValeDetalleS!F4NUMORD & "")
                    .Requerimiento = Trim(rstTmpValeDetalleS!COD_SOLICITUD & "")
                    
                    Rem SK: COMENTADO TEMPORALMENTE
'                    If .verificarStockProductoFisicoCorL(.CodigoProducto, _
                                                            .CodigoAlmacen, _
                                                            .NumeroOrdenCompra, _
                                                            .Requerimiento, _
                                                            .Cantidad, _
                                                            .Fecha) Then
                        
                        
'                    If .devuelveStockFisicoDeProducto(vbNullString, True) >= .Cantidad Then

                    
                    With objAyudaVale
                        .inicializarEntidades
                        .inicializarEntidadesAdicionales
                        
                        .Fecha = objVale.Fecha
                        .CodigoProducto = objVale.CodigoProducto
                        .CodigoAlmacen = objVale.CodigoAlmacen
                    End With
                    
                    If objAyudaVale.devuelveStockFisicoDeProducto >= .Cantidad Then
                    
                        dblItem = dblItem + 1
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    Else
'                        If MsgBox("El producto '" & Trim(rstTmpValeDetalleS!Descripcion & "") & "' no cuenta con Stock al '" & .Fecha & "'." & vbNewLine & _
'                                    "¿Desea descargar el producto a pesar de ello?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'
'                            dblItem = dblItem + 1
'
'                            .ITEM = dblItem
'
'                            .guardarValeDetalleOneByOne
'
'                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
'                        Else
''                            Actualiza_Log "Producto '" & Trim(rstTmpValeDetalleS!Descripcion & "") & "' no descargado en el Vale '" & .CodigoAlmacen & "/" & .NumeroVale & "'. Usuario de Registro: " & wusuario & ".", StrConexDbBancos
''
                            dblItemNoGuardado = dblItemNoGuardado + 1
                            
                            If strDescripcionItemNoGuardado = vbNullString Then
                                strDescripcionItemNoGuardado = left(Trim(rstTmpValeDetalleS!Descripcion & ""), 50)
                            Else
                                strDescripcionItemNoGuardado = strDescripcionItemNoGuardado & ", " & left(Trim(rstTmpValeDetalleS!Descripcion & ""), 50)
                            End If
'                        End If
                    End If
                    
                    DoEvents
                    
                    pgbProceso.value = pgbProceso.value + 1
                    fraProceso.Caption = "Guardando Detalle... " & FormatPercent(pgbProceso.value / pgbProceso.Max, 3)
                    
                    rstTmpValeDetalleS.MoveNext
                Loop
            End If
            
            If Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COUNT(*)", "IF3VALES", "F2CODALM", .CodigoAlmacen, "T", "AND F4NUMVAL = '" & .NumeroVale & "' AND TRIM(F5CODPRO & '') <> ''")) > 0 Then
                'Exportar el Vale
'                If ModMilano.exportarValeAserverSQLv2(.CodigoAlmacen, .NumeroVale, lblNumeroValeExterno, fraProceso, pgbProceso) Then
'
'                End If
'
'                If Val(txtIDOrdenProduccion.Text) > 0 And Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) > 0 Then
'                    ModMilano.actualizarEstadoDescargadoOP txtIDOrdenProduccion.Text, True
'                End If
'
'                strCodAlmacen = .CodigoAlmacen
'                strNumeroVale = .NumeroVale
'
'                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                    guardarValeSql
'                End If
                
                consultarVale
                
                MsgBox "Se ha Actualizado el Vale de Ingreso " & .NumeroVale & "." & vbNewLine & _
                        IIf(dblItemNoGuardado = 0, vbNullString, vbNewLine & vbNewLine & _
                        "ATENCIÓN: " & dblItemNoGuardado & " item(s) no registrado(s) por falta de Stock Disponible al momento de la Grabación." & vbNewLine & _
                        "Items: " & strDescripcionItemNoGuardado), vbInformation + vbOKOnly, App.ProductName
                
                Actualiza_Log "< Vale de Salida: '" & strCodAlmacen & "'-'" & strNumeroVale & "' Registrado." & _
                                IIf(dblItemNoGuardado = 0, vbNullString, "ATENCIÓN: " & dblItemNoGuardado & " item(s) no registrado(s) por falta de Stock Disponible al momento de la Grabación. " & _
                                "Items: " & strDescripcionItemNoGuardado & " > "), StrConexDbBancos
            Else
                If Val(lblNumeroValeExterno.Caption) > 0 Then
'                    If Not ModMilano.anularValeExterno("S", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T"), fraProceso, pgbProceso) Then
'                        Me.MousePointer = vbDefault
'
'                        Exit Sub
'                    End If
                End If
                
                If .eliminarVale Then
                    Actualiza_Log "< Vale de Salida no registrado." & _
                                    "ATENCIÓN: " & dblItemNoGuardado & " item(s) no registrado(s) por falta de Stock Disponible al momento de la Grabación. Items: " & strDescripcionItemNoGuardado & " > " & .SQLSelectAlter, StrConexDbBancos
                    
                    MsgBox "Vale de Salida no registrado." & vbNewLine & _
                            "ATENCIÓN: " & dblItemNoGuardado & " item(s) no registrado(s) por falta de Stock Disponible al momento de la Grabación." & vbNewLine & _
                            "Items: " & strDescripcionItemNoGuardado, vbInformation + vbOKOnly, App.ProductName
                End If
            End If
            
            listarGrillaVale
            
            If dxDBGrid1.Dataset.RecordCount = 0 Then
                adicionarItemVale
            End If
            
            fraProceso.Visible = False
        Else
            listarGrillaVale
        End If
        
        .CodigoAlmacen = ""
        
        Me.MousePointer = vbDefault
    End With
    
    Set objVale = Nothing
    
    If rstTmpValeDetalleS.State = 1 Then rstTmpValeDetalleS.Close
    
    Set rstTmpValeDetalleS = Nothing
    
    dblItem = 0
    dblItemNoGuardado = 0
    strDescripcionItemNoGuardado = vbNullString
    
    Exit Sub
errGuardarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Salida: GuardarVale"
    
    Actualiza_Log "Guardado de Vale [Numero: " & objVale.CodigoAlmacen & " / " & objVale.NumeroVale & "] [ID Externo: " & lblNumeroValeExterno.Caption & "] cancelada por el siguiente error: " & _
                    "[Numero Error: " & Err.Number & "] [Descripción: " & Err.Description & "]", StrConexDbBancos
    
    Err.Clear
End Sub

Private Sub eliminarVale()
    On Error GoTo errEliminarVale
    
    Set objVale = New ClsVale
    
    With objVale
        .CodigoAlmacen = Trim(txtalmacen.Text)
        .NumeroVale = Trim(txtnumero.Text)
        
        .obtenerConfigVale
        
        If Not objVale.obtenerVale Then
            MsgBox "Registro no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
'        'Restricción de Anulación de Vale
'        If Month(CDate(abofecha.value)) < Month(Date) Then
'            MsgBox "Imposible eliminar Vale. Fuera del Periodo Actual." & vbNewLine & vbNewLine & _
'                    "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'            Exit Sub
'        End If
        
        'Validacion del Punto (PC) que elimina el Vale
        'ModMilano.abrirCnDBMilano
        
'        If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'            MsgBox "Su computador no esta registrado y/o habilitado. Consulte con su" & vbNewLine & vbNewLine & _
'                    "administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'            Exit Sub
'        End If
        
'        If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")), "T", _
'                                                                        "AND TABLA = 'INGRESO'")) = 0 Then
'
'            MsgBox "El Punto de Venta no cuenta con correlativo habilitado de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
'                    "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'            Exit Sub
'        End If
        
        If MsgBox("¿Desea eliminar el Vale con No. " & .NumeroVale & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If Val(lblNumeroValeExterno.Caption) > 0 Then
'                If Not ModMilano.anularValeExterno("S", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T"), fraProceso, pgbProceso) Then
'                    Me.MousePointer = vbDefault
'
'                    Exit Sub
'                End If
            End If
            
            If .eliminarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
'                If .CodigoOrigen = "XNC" Then
'                    verificarAtencionOrden .CodigoAlmacen, .NumeroVale, True
'                End If
                
                strCodAlmacen = .CodigoAlmacen
                strNumeroVale = .NumeroVale
                
                consultarVale
                
                MsgBox "Registro eliminado.", vbInformation + vbOKOnly, App.ProductName
            End If
        End If
    End With
    
    Set objVale = Nothing
    
    Exit Sub
    Resume
errEliminarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Salida: EliminarVale"
    
    Err.Clear
End Sub

Private Sub exportarRegistroCompra()
    On Error GoTo errExportarRegistroCompra
    
    Dim rstVale As ADODB.Recordset
    Dim rstValeDet As ADODB.Recordset
    Dim dblItem As Double
    Dim dblMontoCancelado As Double
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = "S"
        .CodigoProveedor = Trim(txtproveedor.Text)
        .CodTipoComprobante = right(cmbtipo.Text, 2)
        .SerieDocumento = Trim(txtserfac.Text)
        .NumeroDocumento = Trim(txtnumfac.Text)
        
        Set rstVale = .obtenerRstValeCompraPorProvYdocumento
        Set rstValeDet = .obtenerRstValeDetalleCompraPorProvYdocumento
        
        If Not rstVale.EOF Then
            rstVale.MoveFirst
            
            With objAyudaCompra
                .inicializarEntidades
                
                If Trim(rstVale!F4REGCOM & "") <> vbNullString Then
                    .MesMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), 1, InStr(1, Trim(rstVale!F4REGCOM & ""), "-") - 1)
                    .NumeroMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), InStr(1, Trim(rstVale!F4REGCOM & ""), "-") + 1)
                Else
                    .MesMovimiento = Year(CDate(dtpFechaDoc.value)) & Format(Month(CDate(dtpFechaDoc.value)), "00")
                    .NumeroMovimiento = vbNullString
                End If
                
                .CodProveedor = Trim(rstVale!F2CODPROV & "")
                .NomProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                .DireccionProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                .TelefonoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TELPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                .TipoDocAuxiliar = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TIPPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                .RucProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                
                    If .TipoDocAuxiliar = vbNullString Then
                        Select Case Len(.RucProveedor)
                            Case 8
                                .TipoDocAuxiliar = "1"
                            Case 11
                                .TipoDocAuxiliar = "6"
                            Case Else
                                .TipoDocAuxiliar = "0"
                        End Select
                    End If
                    
                .TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                .SerieDocumento = Trim(rstVale!F4SERDOC & "")
                .NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                
                .CodigoCategoria = 1
                
                .FechaRegistro = Format(Date, "Short Date")
                .FechaDocumento = Format(Trim(rstVale!F4FECULT & ""), "Short Date")
                
                .CodMoneda = Trim(rstVale!F4MONEDA & "")
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaDocumento, "F"))
                .ConceptoCompra = Trim(rstVale!F4OBSERVA & "")
                
                .CodFormaPago = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                
                With objAyudaFormaPago
                    .inicializarEntidades
                    
                    .Codigo = objAyudaCompra.CodFormaPago
                    
                    .obtenerConfigFormaPago
                End With
                
                If Val(objAyudaFormaPago.Dias) > 0 Then
                    .FechaVencimiento = CDate(.FechaDocumento) + Val(objAyudaFormaPago.Dias)
                Else
                    .FechaVencimiento = .FechaDocumento
                End If
                
                Select Case .CodMoneda
                    Case "S"
                        .CodigoGasto = "PRO"
                        .CuentaContable = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", .CodigoGasto, "T")
                    Case Else
                        .CodigoGasto = "PROD"
                        .CuentaContable = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", .CodigoGasto, "T")
                End Select
                
                .PorcentajeIGV = wIgv
                
                .FechaReg = Format(Date, "Short Date")
                .UsuarioReg = wusuario
                .FechaMod = Format(Date, "Short Date")
                .UsuarioMod = wusuario
                
                If .guardarCompra(True) Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM REGISMOV WHERE F4MESMOV = '" & .MesMovimiento & "' AND F4NUMMOV = '" & .NumeroMovimiento & "'"
                    
                    cnn_dbbancos.Execute .SQLSelectAlter
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    If Not rstValeDet.EOF Then
                        rstValeDet.MoveFirst
                        
                        dblItem = 0
                        
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .ITEM = dblItem
                            
                            .CtaContableDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CTACON", "IF5PLA", "F5CODPRO", Trim(rstValeDet!CodProducto & ""), "T")
                            .CodigoGastoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T")
                            
                            .NumeroOrden = Trim(rstValeDet!NROOC & "")
                            .CodigoProducto = Trim(rstValeDet!CodProducto & "")
                            .ConceptoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", .CodigoProducto, "T")
                            
                            .Cantidad = Val(rstValeDet!Cantidad & "")
                            .PrecioUnitario = Val(rstValeDet!costo & "")
                            .SubTotalDet = Val(rstValeDet!TOTAL & "")
                            .Afecto = IIf(Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5AFECTO", "IF5PLA", "F5CODPRO", .CodigoProducto, "T")) = "*", True, False)
                            
                            .DebHab = "D"
                            
                            'Acumular
                            .BaseImponible = .BaseImponible + (.Cantidad * .PrecioUnitario)
                            .MontoInafecto = .MontoInafecto + IIf(.Afecto, 0, (.Cantidad * .PrecioUnitario))
                            .TotalIGV = .TotalIGV + Val(rstValeDet!IGV & "")
                            
                            .Descuento = .Descuento + Val(rstValeDet!DSCTO & "")
                            .TotalFacturado = .TotalFacturado + ((.BaseImponible + .MontoInafecto + .TotalIGV) - .Descuento)
                            
                            'Concatenar
                            If InStr(1, .OrdenCompra, .NumeroOrden) = 0 Then
                                If .OrdenCompra = vbNullString Then
                                    .OrdenCompra = .NumeroOrden
                                Else
                                    .OrdenCompra = .OrdenCompra & "," & .NumeroOrden
                                End If
                            End If
                            
                            .guardarCompraDetalleOneByOne
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            rstValeDet.MoveNext
                        Loop
                            'ACTUALIZAR POSTERIOR A LA GRABACION
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4BASIMP = " & .BaseImponible & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MONINA = " & .MontoInafecto & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & .TotalIGV & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4DCTO = " & .Descuento & " "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                            
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = (VAL(F4BASIMP & '') + VAL(F4MONINA & '') + VAL(F4IGV & '') + VAL(F4OTRIMP & '') + VAL(F4REDSUMA & '')) - (VAL(F4FONAVI & '') + VAL(F4DCTO & '') + VAL(F4MONTORET & '') + VAL(F4REDRESTA & '')) "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                            
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End If
                
                If .obtenerCompra Then
                
                    With objAyudaPagDcto
                        .Correlativo = objAyudaCompra.Correlativo
                        
                        .TipoIngreso = "1"
                        .ITEM = 1
                        .NumeroComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2ABREV", "DOCUMENTOS", "F2CODDOC", objAyudaCompra.TipoDocumento, "T")
                        
                        .NumeroComprobante = .NumeroComprobante & objAyudaCompra.SerieDocumento & "/" & objAyudaCompra.NumeroDocumento
                        
                        .FechaComprobante = objAyudaCompra.FechaDocumento
                        .FechaVencimiento = objAyudaCompra.FechaVencimiento
                        .CodProveedor = objAyudaCompra.CodProveedor
                        .RucProveedor = objAyudaCompra.RucProveedor
                        .NomProveedor = objAyudaCompra.NomProveedor
                        .CodMoneda = objAyudaCompra.CodMoneda
                        .TotalFacturado = objAyudaCompra.TotalFacturado
                        .SaldoFacturado = objAyudaCompra.TotalFacturado
                        .TipoCambio = objAyudaCompra.TipoCambio
                        .Debe_Haber = "H"
                        
                        .Grupo = objAyudaCompra.CodigoGasto
                        .CtaContable = objAyudaCompra.CuentaContable
                        .AnnoRegCompra = left(objAyudaCompra.MesMovimiento, 4)
                        .MovRegCompra = objAyudaCompra.NumeroMovimiento
                        .TipoDocumento = objAyudaCompra.TipoDocumento
                        .SerieDocumento = objAyudaCompra.SerieDocumento
                        .NumeroDocumento = Val(objAyudaCompra.NumeroDocumento)
                        .Notas = "CUENTA POR PAGAR GENERADA DESDE EL MODULO DE LOGISTICA."
                        .Concepto = "INGRESO DE COMPRAS A ALMACEN."
                        .Detalle = "INGRESO DE COMPRAS A ALMACEN."
                        .referencia = "INGRESO DE COMPRAS A ALMACEN."
                            
                        If .obtenerPagDcto Then
                            If .SaldoFacturado = .TotalFacturado Then
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                .SaldoFacturado = objAyudaCompra.TotalFacturado
                            Else
                                .SaldoFacturado = objAyudaCompra.TotalFacturado - (.TotalFacturado - .SaldoFacturado)
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                            End If
                        End If
                        
                        Call .guardarPagDcto(False)
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        If objAyudaCompra.Correlativo = 0 Then
                            'ACTUALIZAR POSTERIOR A LA GRABACION
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4CORRELA = " & .Correlativo & " "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & objAyudaCompra.MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & objAyudaCompra.NumeroMovimiento & "'"
                            
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        End If
                        
                        .inicializarEntidades
                    End With
                    
                    .ValeIngreso = vbNullString
                    
                    Do While Not rstVale.EOF
                        If .ValeIngreso = vbNullString Then
                            .ValeIngreso = Trim(rstVale!f2codalm & "") & "/" & Trim(rstVale!F4NUMVAL & "")
                        Else
                            .ValeIngreso = .ValeIngreso & "," & Trim(rstVale!f2codalm & "") & "/" & Trim(rstVale!F4NUMVAL & "")
                        End If
                        
                        'ACTUALIZAR VALE(S) REFERENCIA DE REGISTRO DE COMPRA
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                        .SQLSelectAlter = .SQLSelectAlter & "IF4VALES "
                        .SQLSelectAlter = .SQLSelectAlter & "SET "
                        .SQLSelectAlter = .SQLSelectAlter & "F4REGCOM = '" & .MesMovimiento & "-" & .NumeroMovimiento & "' "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F2CODALM = '" & Trim(rstVale!f2codalm & "") & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMVAL = '" & Trim(rstVale!F4NUMVAL & "") & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        rstVale.MoveNext
                    Loop
                        'ACTUALIZAR POSTERIOR A LA GRABACION
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                        .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                        .SQLSelectAlter = .SQLSelectAlter & "SET "
                        .SQLSelectAlter = .SQLSelectAlter & "F4VALESING = '" & left(.ValeIngreso, 100) & "' "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                    
                    MsgBox "Ingreso exportado a Registro de Compras:" & vbNewLine & "Mes de Registro: " & .MesMovimiento & vbNewLine & "Numero de Movimiento: " & .NumeroMovimiento, vbInformation + vbOKOnly, App.ProductName
                End If
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
        End If
        
        .inicializarEntidades
    End With
    
    Exit Sub
errExportarRegistroCompra:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description & vbNewLine & vbNewLine & "Registro de Compra NO EXPORTADO correctamente a Tesoreria, intente volver a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub cerrarVale()
    Dim strFechaCorteInicialDeValesParaCP As String
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    strFechaCorteInicialDeValesParaCP = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With objAyudaVale
            .inicializarEntidades
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If Val(Year(CDate(.Fecha)) & Format(Month(CDate(.Fecha)), "00")) > Val(Format(CDate(strFechaCorteInicialDeValesParaCP), "yyyymm")) Then
                .inicializarEntidadesAdicionales
                
                intAnnoCorte = Val(Year(CDate(.Fecha))) - IIf(Val(Month(CDate(.Fecha))) > 1, 0, 1)
                intMesCorte = IIf(Val(Month(CDate(.Fecha))) > 1, Val(Month(CDate(.Fecha))) - 1, 12)
                
                .inicializarEntidades
                
                .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
                .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
                
                If Not .verificarCierreVale Then
                    MsgBox "Imposible cerrar el Vale; ya que el anterior Periodo aun se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    bolObviarCierre = True
                    
                    SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
                    
                    If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                        SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                    Else
                        SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                    End If
                    
                    bolObviarCierre = False
                    
                    Exit Sub
                End If
            End If
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If .VB1 Then
                MsgBox "Vale ya se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    Else
        With objAyudaVale
            .inicializarEntidades
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If Val(Year(CDate(.Fecha)) & Format(Month(CDate(.Fecha)), "00")) > Val(Format(CDate(strFechaCorteInicialDeValesParaCP), "yyyymm")) Then
                .inicializarEntidadesAdicionales
                
                intAnnoCorte = Val(Year(CDate(.Fecha))) - IIf(Val(Month(CDate(.Fecha))) > 1, 0, 1)
                intMesCorte = IIf(Val(Month(CDate(.Fecha))) > 1, Val(Month(CDate(.Fecha))) - 1, 12)
                
                .inicializarEntidades
                
                .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
                .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
                
                If Not .verificarCierreVale Then
                    MsgBox "Imposible cerrar el Vale; ya que el anterior Periodo aun se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    bolObviarCierre = True
                    
                    SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
                    
                    If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                        SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                    Else
                        SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                    End If
                    
                    bolObviarCierre = False
                    
                    Exit Sub
                End If
            End If
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If .VB1 Then
                MsgBox "Vale ya se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    End If
    
    With objAyudaVale
        If MsgBox("¿Desea cerrar el Vale?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            .VB1 = True
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            If .cerrarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                MsgBox "Vale cerrado correctamente.", vbInformation + vbOKOnly, App.ProductName
            End If
            
            strCodAlmacen = .CodigoAlmacen
            strNumeroVale = .NumeroVale
            
            consultarVale
        Else
            bolObviarCierre = True
            
            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
            
            bolObviarCierre = False
        End If
    End With
End Sub

Private Sub abrirVale()
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            intAnnoCorte = Val(Year(CDate(.Fecha))) + IIf(Val(Month(CDate(.Fecha))) < 12, 0, 1)
            intMesCorte = IIf(Val(Month(CDate(.Fecha))) < 12, Val(Month(CDate(.Fecha))) + 1, 1)
            
            .inicializarEntidades
            
            .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
            .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
            
            If .verificarCierreVale Then
                MsgBox "Imposible abrir el Vale; ya que el Periodo posterior se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If Not .VB1 Then
                MsgBox "Vale se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    Else
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            intAnnoCorte = Val(Year(CDate(.Fecha))) + IIf(Val(Month(CDate(.Fecha))) < 12, 0, 1)
            intMesCorte = IIf(Val(Month(CDate(.Fecha))) < 12, Val(Month(CDate(.Fecha))) + 1, 1)
            
            .inicializarEntidades
            
            .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
            .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
            
            If .verificarCierreVale Then
                MsgBox "Imposible abrir el Vale; ya que el Periodo posterior se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = strCodAlmacen
            .NumeroVale = strNumeroVale
            
            .obtenerConfigVale
            
            If Not .VB1 Then
                MsgBox "Vale se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
                
                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
                Else
                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    End If
    
    With objAyudaVale
        If MsgBox("¿Desea abrir el Vale?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            .VB1 = False
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            If .cerrarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                MsgBox "Vale abierto correctamente.", vbInformation + vbOKOnly, App.ProductName
            End If
            
            strCodAlmacen = .CodigoAlmacen
            strNumeroVale = .NumeroVale
            
            consultarVale
        Else
            bolObviarCierre = True
            
            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
            
            bolObviarCierre = False
        End If
    End With
End Sub


