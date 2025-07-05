VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vale_ingreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vale de Ingreso a Almacen"
   ClientHeight    =   8550
   ClientLeft      =   2205
   ClientTop       =   1680
   ClientWidth     =   16455
   Icon            =   "vale_ingreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   16455
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   11400
      TabIndex        =   73
      Top             =   2280
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar pgbProceso 
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   690
      Left            =   5640
      TabIndex        =   47
      Top             =   7800
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   1217
      _StockProps     =   14
      Caption         =   "Totales"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox txtDscto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   4440
         TabIndex        =   67
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtTotvv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtTotigv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   6585
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtTotpv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   210
         Left            =   3120
         TabIndex        =   69
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   210
         Left            =   4125
         TabIndex        =   68
         Top             =   285
         Width           =   165
      End
      Begin VB.Label Lbl3 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   210
         Left            =   9120
         TabIndex        =   56
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Lbl2 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   210
         Left            =   1350
         TabIndex        =   55
         Top             =   315
         Width           =   165
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   210
         Left            =   6270
         TabIndex        =   54
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valor Venta"
         Height          =   210
         Left            =   225
         TabIndex        =   53
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "IGV"
         Height          =   210
         Left            =   5865
         TabIndex        =   52
         Top             =   315
         Width           =   270
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Precio Venta"
         Height          =   210
         Left            =   8010
         TabIndex        =   51
         Top             =   315
         Width           =   930
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   555
      Left            =   10680
      TabIndex        =   26
      Top             =   75
      Width           =   4485
      _Version        =   65536
      _ExtentX        =   7911
      _ExtentY        =   979
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   90
         Width           =   1905
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
         Left            =   2880
         TabIndex        =   59
         Top             =   165
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   27
         Top             =   180
         Width           =   570
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   465
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   820
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Autosize        =   3
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   12
         ShowShortcutsInToolTips=   -1  'True
         Tools           =   "vale_ingreso.frx":058A
         ToolBars        =   "vale_ingreso.frx":B687
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3000
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   5292
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox chkVerObservaciones 
         Caption         =   "Visualizar Observaciones por Item."
         Height          =   255
         Left            =   11160
         TabIndex        =   66
         Top             =   2520
         Width           =   2895
      End
      Begin VB.ComboBox cmbTipoAuxiliar 
         Height          =   330
         ItemData        =   "vale_ingreso.frx":B8F5
         Left            =   1560
         List            =   "vale_ingreso.frx":B8FF
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2115
      End
      Begin VB.Frame fraOrdenProduccion 
         Caption         =   " O. Producción "
         Height          =   1335
         Left            =   7320
         TabIndex        =   61
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
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
            TabIndex        =   64
            Top             =   960
            Width           =   2535
         End
         Begin VB.ComboBox cmbCategoriaTipo 
            Height          =   330
            Left            =   120
            TabIndex        =   17
            Text            =   "cmbCategoriaTipo"
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtNroOrdenProduccion 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   720
            MaxLength       =   30
            TabIndex        =   18
            Top             =   600
            Width           =   2535
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
            TabIndex        =   63
            Top             =   960
            Width           =   510
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "O.P."
            Height          =   210
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   540
         End
      End
      Begin VB.CheckBox chkExportarVale 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportar Vale"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7920
         TabIndex        =   60
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtnomprov 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   70
         TabIndex        =   58
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txtOcompra 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   11100
         TabIndex        =   14
         Top             =   960
         Width           =   3690
      End
      Begin VB.TextBox txtobserva 
         Height          =   915
         Left            =   11100
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1560
         Width           =   3690
      End
      Begin VB.ComboBox cmbmoneda 
         BackColor       =   &H0080FFFF&
         Height          =   330
         ItemData        =   "vale_ingreso.frx":B9E2
         Left            =   11100
         List            =   "vale_ingreso.frx":B9EC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1320
      End
      Begin VB.TextBox txtconcepto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5880
         TabIndex        =   20
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox txtalmacen 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   5880
         TabIndex        =   19
         Top             =   180
         Width           =   495
      End
      Begin VB.ComboBox cmbconcepto 
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   560
         Width           =   4290
      End
      Begin VB.ComboBox cmbalmacen 
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   4290
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   330
         ItemData        =   "vale_ingreso.frx":BA00
         Left            =   1560
         List            =   "vale_ingreso.frx":BA02
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtserfac 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   8
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtnumfac 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtserie 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1890
         Width           =   855
      End
      Begin VB.TextBox txtnumdoc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1880
         Width           =   1695
      End
      Begin VB.TextBox txtccosto 
         Height          =   315
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1305
         Width           =   1095
      End
      Begin VB.TextBox txttc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   13500
         MaxLength       =   11
         TabIndex        =   13
         Text            =   "4.05"
         Top             =   600
         Width           =   1320
      End
      Begin VB.TextBox txtproveedor 
         Height          =   315
         Left            =   3720
         MaxLength       =   11
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin Threed.SSPanel pnlccosto 
         Height          =   315
         Left            =   2760
         TabIndex        =   33
         Top             =   1305
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.ComboBox cmbmarca 
         Height          =   330
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3480
         Visible         =   0   'False
         Width           =   885
      End
      Begin Threed.SSCommand cmdimporta 
         Height          =   330
         Left            =   1440
         TabIndex        =   46
         Top             =   3195
         Visible         =   0   'False
         Width           =   480
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Importación"
      End
      Begin Threed.SSPanel pnluupp 
         Height          =   315
         Left            =   990
         TabIndex        =   41
         Top             =   3195
         Visible         =   0   'False
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.TextBox txtuupp 
         Height          =   315
         Left            =   765
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3195
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComCtl2.DTPicker abofecha 
         Height          =   315
         Left            =   11085
         TabIndex        =   11
         Top             =   240
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
         Format          =   117702657
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker dtpFechaDoc 
         Height          =   300
         Left            =   4320
         TabIndex        =   10
         Top             =   2595
         Width           =   1695
         _ExtentX        =   2990
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
         CheckBox        =   -1  'True
         Format          =   117702657
         CurrentDate     =   40611
      End
      Begin VB.Label lblIdCategoriaTipo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID CategoriaTipo"
         Height          =   255
         Left            =   9480
         TabIndex        =   72
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   12480
         TabIndex        =   71
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label23 
         Caption         =   "Fecha de Documento"
         Height          =   195
         Left            =   2640
         TabIndex        =   70
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Orden de compra"
         Height          =   210
         Left            =   9720
         TabIndex        =   57
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label lbloc 
         AutoSize        =   -1  'True
         Caption         =   "O/Compra"
         Height          =   210
         Left            =   9450
         TabIndex        =   45
         Top             =   3330
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   210
         Left            =   9720
         TabIndex        =   43
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observ."
         Height          =   210
         Left            =   11160
         TabIndex        =   40
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   210
         Left            =   5280
         TabIndex        =   39
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   210
         Left            =   4320
         TabIndex        =   38
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   210
         Left            =   2520
         TabIndex        =   37
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   210
         Left            =   1560
         TabIndex        =   36
         Top             =   1665
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Guía de Remisión"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   1935
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "C. Costo"
         Height          =   210
         Left            =   225
         TabIndex        =   32
         Top             =   1350
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo/cambio"
         Height          =   210
         Left            =   12480
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblproveedor 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         Height          =   210
         Left            =   225
         TabIndex        =   30
         Top             =   990
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   9720
         TabIndex        =   29
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   225
         TabIndex        =   28
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Almacén"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lbluupp 
         AutoSize        =   -1  'True
         Caption         =   "UUPP"
         Height          =   210
         Left            =   360
         TabIndex        =   42
         Top             =   3240
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "vale_ingreso.frx":BA04
      TabIndex        =   65
      Top             =   3600
      Width           =   16200
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid12 
      Height          =   1350
      Left            =   135
      OleObjectBlob   =   "vale_ingreso.frx":1342A
      TabIndex        =   16
      Top             =   3660
      Width           =   15000
   End
   Begin Threed.SSPanel pnlRegistroCompra 
      Height          =   555
      Left            =   60
      TabIndex        =   75
      Top             =   7890
      Width           =   4485
      _Version        =   65536
      _ExtentX        =   7911
      _ExtentY        =   979
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtRegistroCompra 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Registro de Compra"
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
         TabIndex        =   77
         Top             =   180
         Width           =   1665
      End
   End
End
Attribute VB_Name = "vale_ingreso"
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
Dim Values()            As Variant
'Dim amovs_cab(0 To 17)  As a_grabacion
Dim amovs_cab(0 To 21)  As a_grabacion
'Dim amovs_det(0 To 12)  As a_grabacion
Dim amovs_det(0 To 15)  As a_grabacion
Dim ctipo               As String * 1
Dim cvalores            As String
Dim cmes                As String * 2
Dim RSDETALLE           As New ADODB.Recordset
Dim nfil                As Integer
Dim sw_cabecera         As Boolean
Dim sw_detalle          As Boolean
Dim sw_ayuda_prod       As Boolean
Dim compra              As Boolean 'se usa para comprobar si se graba o no el ultimo precio de compra
Dim Rs                  As New ADODB.Recordset
Dim rst                 As New ADODB.Recordset
Dim I                   As Integer
Dim wopcion             As Byte
Dim pnumvale            As String
Dim palmacen            As String
Dim wnumsord(999)       As String
Dim sw_Orden            As Boolean
Dim rsOr                 As New ADODB.Recordset
Dim sw_Ord As Boolean

Rem SK ADD:
Private bolAyuda            As Boolean
Private strCodAlmacen       As String
Private strNumeroVale       As String

Private strFichero          As String
Private bolObviarCierre     As Boolean

Private objVale             As ClsVale



Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let CodigoAlmacen(ByVal Value As String)
    strCodAlmacen = Value
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = strCodAlmacen
End Property

Public Property Let NumeroVale(ByVal Value As String)
    strNumeroVale = Value
End Property

Public Property Get NumeroVale() As String
    NumeroVale = strNumeroVale
End Property

Rem SK ADD:----------------------------------------------------------------------------------------------------------
Private Sub copiarSeleccionAyudaProductos()
    Dim rstProductoIngreso As New ADODB.Recordset
    Dim dblItem As Double
    Dim dblPrecioSinIGV As Double, dblPrecionConIgv As Double, dblSubTotalPorItem As Double, dblIgvPorItem As Double, dblTotalPorItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
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
    
    If rstProductoIngreso.State = 1 Then rstProductoIngreso.Close
    
    rstProductoIngreso.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProductoIngreso.EOF Then
        rstProductoIngreso.MoveFirst
        
        Do While Not rstProductoIngreso.EOF
            dblPrecioSinIGV = 0
            dblPrecionConIgv = 0
            dblSubTotalPorItem = 0
            dblIgvPorItem = 0
            dblTotalPorItem = 0
            
            'Verificar si Concepto admite registro de Costo, Solo en caso de Proveedores
            If cmbTipoAuxiliar.ListIndex = 1 Then
                With objAyudaOrigen
                    .inicializarEntidades
                    
                    .Codigo = right(cmbconcepto.Text, 3)  'Trim(Mid(cmbconcepto.Text, 200, 3))
                    
                    .obtenerConfigOrigen
                    
                    If .RegistrarCosto Then
                        'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                        With objAyudaVale
                            .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
                            .CodigoProducto = Trim(rstProductoIngreso!f5codpro & "")
                            
                            .obtenerUltimoPrecioSinIgvProductoDeProveedor
                            
                            Select Case .CodigoMoneda
                                Case "S"
                                    dblPrecioSinIGV = Val(Format(.ValorVenta / IIf(left(cmbmoneda.Text, 1) = "S", 1, Val(txttc.Text)), "#0.0000"))
                                Case Else
                                    dblPrecioSinIGV = Val(Format(.ValorVentaDol * IIf(left(cmbmoneda.Text, 1) = "D", 1, Val(txttc.Text)), "#0.0000"))
                            End Select
                            
                            'dblPrecionConIgv = Val(Format(dblPrecioSinIGV * IIf(CBool(rstProductoIngreso!Afecto), (1 + (wIgv / 100)), 0), "#0.0000"))
                            'dblSubTotalPorItem = Val(Format(dblPrecioSinIGV * Val(rstProductoIngreso!F5FOB & ""), "#0.00"))
                            'dblIgvPorItem = Val(Format(dblSubTotalPorItem * IIf(CBool(rstProductoIngreso!Afecto), (wIgv / 100), 0), "#0.00"))
                            'dblTotalPorItem = Val(Format(dblSubTotalPorItem + dblIgvPorItem, "#0.00"))
                        End With
                    End If
                End With
            End If
            
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALEINGRESO", "CODPROD", Trim(rstProductoIngreso!f5codpro & ""), "T", "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstProductoIngreso!COD_SOLICITUD & "") & "' AND TRIM(F4NUMORD & '') = ''") = vbNullString Then
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
'                If dblItem = 1 Then
'                    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
'                End If
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALEINGRESO("
                CadSql = CadSql & "ITEM, COD_SOLICITUD, CODPROD, CODPRODORIGINAL, "
                CadSql = CadSql & "CODFAB, DESCRIPCION, UMEDIDA, AFECTO, "
                'CadSql = CadSql & "AFECTO, CANTIDAD, CANTIDADMAX, COSTOUNI, PVUNIT, VVTOTAL, IGV, TOTAL) "
                CadSql = CadSql & "CANTIDAD, CANTIDADMAX, COSTOUNI, PORDESC) "
                CadSql = CadSql & "VALUES(" & dblItem & ", '" & Trim(rstProductoIngreso!COD_SOLICITUD & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProductoIngreso!f5codpro & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProductoIngreso!f5codpro & "") & "', "
                CadSql = CadSql & IIf(Trim(rstProductoIngreso!f5codfab & "") <> vbNullString, "'" & Trim(rstProductoIngreso!f5codfab & "") & "'", "NULL") & ", "
                CadSql = CadSql & "" & IIf(Trim(rstProductoIngreso!F5NOMPRO & "") <> vbNullString, "'" & Trim(rstProductoIngreso!F5NOMPRO & "") & "'", "NULL") & ", "
                CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", Trim(rstProductoIngreso!f7codmed & ""), "T") & "', "
                CadSql = CadSql & "'" & IIf(CBool(rstProductoIngreso!Afecto), "*", vbNullString) & "', "
                'CadSql = CadSql & Val(rstProductoIngreso!F5FOB & "") & ", 0, " & dblPrecioSinIGV & ", " & dblPrecionConIgv & ", " & dblSubTotalPorItem & ", "
                CadSql = CadSql & Val(rstProductoIngreso!F5FOB & "") & ", 0, " & dblPrecioSinIGV & ", " & objAyudaVale.PorcentajeDscto & ")"
                'CadSql = CadSql & dblIgvPorItem & ", " & dblTotalPorItem & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALEINGRESO "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CANTIDAD = " & Val(rstProductoIngreso!F5FOB & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = 0, "
                CadSql = CadSql & "COSTOUNI = " & dblPrecioSinIGV & ", "
                CadSql = CadSql & "PORDESC = " & objAyudaVale.PorcentajeDscto & " "
                'CadSql = CadSql & "PVUNIT = " & dblPrecionConIgv & ", "
                'CadSql = CadSql & "VVTOTAL = " & dblSubTotalPorItem & ", "
                'CadSql = CadSql & "IGV = " & dblIgvPorItem & ", "
                'CadSql = CadSql & "TOTAL = " & dblTotalPorItem & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "TRIM(F4NUMORD & '') = '' AND "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstProductoIngreso!COD_SOLICITUD & "") & "' AND "
                CadSql = CadSql & "CODPROD = '" & Trim(rstProductoIngreso!f5codpro & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProductoIngreso.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProductoIngreso.State = 1 Then rstProductoIngreso.Close
    
    Set rstProductoIngreso = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionStockDisponible()
    Dim rstStockDisponible As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILSTOCKDISPONIBLE "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE"
    
    If rstStockDisponible.State = 1 Then rstStockDisponible.Close
    
    rstStockDisponible.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStockDisponible.EOF Then
        rstStockDisponible.MoveFirst
        
'        CadSql = vbNullString
'        CadSql = CadSql & "DELETE FROM TMPVALESALIDA "
'        CadSql = CadSql & "WHERE "
'        CadSql = CadSql & "F4NUMORD NOT IN (SELECT NROOC FROM TMPUTILSTOCKDISPONIBLE WHERE CODPROVEEDOR = '" & Trim(rstStockDisponible!CodProveedor & "") & "' GROUP BY NROOC)"
'
'        cnDBTemp.Execute CadSql
        
        Do While Not rstStockDisponible.EOF
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "CODPROD", "TMPVALESALIDA", "CODPROD", Trim(rstStockDisponible!CodProducto & ""), "T", _
                                            "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstStockDisponible!NroPedido & "") & "' " & _
                                            "AND TRIM(F4NUMORD & '') = '" & Trim(rstStockDisponible!NROOC & "") & "'") = vbNullString Then
                
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALESALIDA", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
'                If dblItem = 1 Then
'                    cnDBTemp.Execute "DELETE FROM TMPVALESALIDA WHERE TRIM(CODPROD & '') = ''"
'                End If
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPVALESALIDA(ITEM, F4NUMORD, COD_SOLICITUD, "
                CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, CANTIDAD, "
                CadSql = CadSql & "CANTIDADMAX) "
                CadSql = CadSql & "VALUES("
                CadSql = CadSql & dblItem & ", "
                CadSql = CadSql & "'" & Trim(rstStockDisponible!NROOC & "") & "', "
                CadSql = CadSql & "'" & Trim(rstStockDisponible!NroPedido & "") & "', "
                CadSql = CadSql & "'" & Trim(rstStockDisponible!CodProducto & "") & "', "
                CadSql = CadSql & "'" & Trim(rstStockDisponible!CODPRODUCTOORIGINAL & "") & "', "
                CadSql = CadSql & IIf(Trim(rstStockDisponible!NOMPRODUCTO & "") <> vbNullString, "'" & Trim(rstStockDisponible!NOMPRODUCTO & "") & "'", "NULL") & ", "
                CadSql = CadSql & IIf(Trim(rstStockDisponible!um & "") <> vbNullString, "'" & Trim(rstStockDisponible!um & "") & "'", "NULL") & ", "
                CadSql = CadSql & Val(rstStockDisponible!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & Val(rstStockDisponible!CANTIDADDESTINO & "") & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPVALESALIDA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "CANTIDAD = " & Val(rstStockDisponible!CANTIDADDESTINO & "") & ", "
                CadSql = CadSql & "CANTIDADMAX = " & Val(rstStockDisponible!CANTIDADDESTINO & "") & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "F4NUMORD = '" & Trim(rstStockDisponible!NROOC & "") & "' AND "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstStockDisponible!NroPedido & "") & "' AND "
                CadSql = CadSql & "CODPROD = '" & Trim(rstStockDisponible!CodProducto & "") & "' AND "
                CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstStockDisponible!CODPRODUCTOORIGINAL & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstStockDisponible.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstStockDisponible.State = 1 Then rstStockDisponible.Close
    
    Set rstStockDisponible = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionRecepcionOC()
    Dim rstRecepcionOC As New ADODB.Recordset
    Dim dblItem As Double
    
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILDEVOLUCIONOC "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE"
    
    If rstRecepcionOC.State = 1 Then rstRecepcionOC.Close
    
    rstRecepcionOC.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstRecepcionOC.EOF Then
        rstRecepcionOC.MoveFirst
        
'        CadSql = vbNullString
'        CadSql = CadSql & "DELETE FROM TMPVALEINGRESO "
'        CadSql = CadSql & "WHERE "
'        CadSql = CadSql & "F4NUMORD NOT IN (SELECT NROOC FROM TMPUTILDEVOLUCIONOC WHERE CODPROVEEDOR = '" & Trim(rstRecepcionOC!CodProveedor & "") & "' GROUP BY NROOC)"
'
'        abrirCnTemporal
'
'        cnDBTemp.Execute CadSql
        
        Do While Not rstRecepcionOC.EOF
            With objAyudaOrden
                .inicializarEntidadesDetalle
                
                .TipoOrden = "OC"
                .NumeroOrden = Trim(rstRecepcionOC!NROOC & "")
                .CodigoProducto = Trim(rstRecepcionOC!CodProducto & "")
                .Requerimiento = Trim(rstRecepcionOC!NroPedido & "")
                
                If Trim(txtOcompra.Text) = vbNullString Then
                    txtOcompra.Text = .NumeroOrden
                Else
                    If InStr(1, txtOcompra.Text, .NumeroOrden) = 0 Then
                        txtOcompra.Text = txtOcompra.Text & "," & .NumeroOrden
                    End If
                End If
                
                .obtenerConfigOrdenDetalleOnebyOne
            
            
                If ModUtilitario.ObtenerCampoV2(cnDBTemp, "TRIM(CODPROD & '')", "TMPVALEINGRESO", "TRIM(CODPROD & '')", Trim(rstRecepcionOC!CodProducto & ""), "T", _
                                                "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstRecepcionOC!NroPedido & "") & "' " & _
                                                "AND TRIM(F4NUMORD & '') = '" & Trim(rstRecepcionOC!NROOC & "") & "'") = vbNullString Then
                             
                    dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                    
                    'If dblItem = 1 Then
                    '    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
                    'End If
                    
                    CadSql = vbNullString
                    CadSql = CadSql & "INSERT INTO TMPVALEINGRESO(ITEM, F4NUMORD, COD_SOLICITUD, "
                    CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, CANTIDAD, "
                    'CadSql = CadSql & "CANTIDADMAX, AFECTO, COSTOUNI, PVUNIT, VVTOTAL, IGV, TOTAL) "
                    CadSql = CadSql & "CANTIDADMAX, AFECTO, COSTOUNI, PORDESC) "
                    CadSql = CadSql & "VALUES("
                    CadSql = CadSql & dblItem & ", "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NROOC & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NroPedido & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!CodProducto & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!CODPRODUCTOORIGINAL & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NOMPRODUCTO & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!um & "") & "', "
                    CadSql = CadSql & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "'*', "
                    'CadSql = CadSql & .PrecioNetoSinImpuesto & ", "
                    CadSql = CadSql & Val(rstRecepcionOC!PrecioUnitario & "") & ", "
                    CadSql = CadSql & .PorcentajeDscto
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & Val(Format(Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000"))
                    CadSql = CadSql & ")"
                Else
                    CadSql = vbNullString
                    CadSql = CadSql & "UPDATE TMPVALEINGRESO "
                    CadSql = CadSql & "SET "
                    CadSql = CadSql & "CANTIDAD = " & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "CANTIDADMAX = " & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "AFECTO = '*', "
                    CadSql = CadSql & "COSTOUNI = " & Val(rstRecepcionOC!PrecioUnitario & "") & ", "
                    CadSql = CadSql & "PORDESC = " & .PorcentajeDscto & " "
                    'CadSql = CadSql & "PVUNIT = " & Val(Format(Val(rstRecepcionOC!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & "VVTOTAL = " & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & "IGV = " & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & "TOTAL = " & Val(Format(Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000")) & " "
                    CadSql = CadSql & "WHERE "
                    CadSql = CadSql & "F4NUMORD = '" & Trim(rstRecepcionOC!NROOC & "") & "' AND "
                    CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstRecepcionOC!NroPedido & "") & "' AND "
                    CadSql = CadSql & "CODPROD = '" & Trim(rstRecepcionOC!CodProducto & "") & "' AND "
                    CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstRecepcionOC!CODPRODUCTOORIGINAL & "") & "'"
                End If
            End With
            
            cnDBTemp.Execute CadSql
            
            rstRecepcionOC.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstRecepcionOC.State = 1 Then rstRecepcionOC.Close
    
    Set rstRecepcionOC = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionRecepcionOCSql()
    Dim rstRecepcionOC As New ADODB.Recordset
    Dim dblItem As Double
    
    Dim dblUltimoPrecioSinIGv As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPCPRECEPCIONOC" & UCase(wusuario) & " "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = 1"
    
    If rstRecepcionOC.State = 1 Then rstRecepcionOC.Close
    
    rstRecepcionOC.Open CadSql, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstRecepcionOC.EOF Then
        rstRecepcionOC.MoveFirst
        
        Do While Not rstRecepcionOC.EOF
            With objSqlAyudaOrden
                .inicializarEntidadesDetalle
                
                .TipoOrden = "OC"
                .NumeroOrden = Trim(rstRecepcionOC!NROOC & "")
                .CodigoProducto = Trim(rstRecepcionOC!CodProducto & "")
                .Requerimiento = Trim(rstRecepcionOC!NroPedido & "")
                
                If Trim(txtOcompra.Text) = vbNullString Then
                    txtOcompra.Text = .NumeroOrden
                Else
                    If InStr(1, txtOcompra.Text, .NumeroOrden) = 0 Then
                        txtOcompra.Text = txtOcompra.Text & "," & .NumeroOrden
                    End If
                End If
                
                .obtenerConfigOrdenDetalleOnebyOne
                
                If Val(rstRecepcionOC!PrecioUnitario & "") = 0 Then
                    'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                    With objAyudaVale
                        .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")   'Trim(Txt_Prove.Text)
                        .CodigoProducto = objSqlAyudaOrden.CodigoProducto
                        
                        .obtenerUltimoPrecioSinIgvProductoDeProveedor
                        
                        Select Case .CodigoMoneda
                            Case "S"
                                dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(cmbmoneda.Text, 1) = "S", 1, Val(txttc.Text))), "#0.0000")
                            Case Else
                                dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(cmbmoneda.Text, 1) = "D", 1, Val(txttc.Text))), "#0.0000")
                        End Select
                    End With
                Else
                    dblUltimoPrecioSinIGv = Val(rstRecepcionOC!PrecioUnitario & "")
                End If
                
                
                If ModUtilitario.ObtenerCampoV2(cnDBTemp, "TRIM(CODPROD & '')", "TMPVALEINGRESO", "TRIM(CODPROD & '')", Trim(rstRecepcionOC!CodProducto & ""), "T", _
                                                "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstRecepcionOC!NroPedido & "") & "' " & _
                                                "AND TRIM(F4NUMORD & '') = '" & Trim(rstRecepcionOC!NROOC & "") & "'") = vbNullString Then
                             
                    dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                    
                    'If dblItem = 1 Then
                    '    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
                    'End If
                    
                    CadSql = vbNullString
                    CadSql = CadSql & "INSERT INTO TMPVALEINGRESO(ITEM, F4NUMORD, COD_SOLICITUD, "
                    CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, CANTIDAD, "
                    'CadSql = CadSql & "CANTIDADMAX, AFECTO, COSTOUNI, PVUNIT, VVTOTAL, IGV, TOTAL) "
                    CadSql = CadSql & "CANTIDADMAX, AFECTO, COSTOUNI, PORDESC) "
                    CadSql = CadSql & "VALUES("
                    CadSql = CadSql & dblItem & ", "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NROOC & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NroPedido & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!CodProducto & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!CODPRODUCTOORIGINAL & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!NOMPRODUCTO & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstRecepcionOC!um & "") & "', "
                    CadSql = CadSql & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "'" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & .PrecioNetoSinImpuesto & ", "
                    'CadSql = CadSql & Val(rstRecepcionOC!PrecioUnitario & "") & ", "
                    CadSql = CadSql & dblUltimoPrecioSinIGv & ", "
                    CadSql = CadSql & .PorcentajeDscto
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & Val(Format(Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000"))
                    CadSql = CadSql & ")"
                Else
                    CadSql = vbNullString
                    CadSql = CadSql & "UPDATE TMPVALEINGRESO "
                    CadSql = CadSql & "SET "
                    CadSql = CadSql & "CANTIDAD = " & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "CANTIDADMAX = " & Val(rstRecepcionOC!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "AFECTO = '" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & "COSTOUNI = " & Val(rstRecepcionOC!PrecioUnitario & "") & ", "
                    CadSql = CadSql & dblUltimoPrecioSinIGv & ", "
                    CadSql = CadSql & "PORDESC = " & .PorcentajeDscto & " "
                    'CadSql = CadSql & "PVUNIT = " & Val(Format(Val(rstRecepcionOC!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & "VVTOTAL = " & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & "IGV = " & Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & "TOTAL = " & Val(Format(Val(Format(Val(rstRecepcionOC!CANTIDADDESTINO & "") * Val(rstRecepcionOC!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000")) & " "
                    CadSql = CadSql & "WHERE "
                    CadSql = CadSql & "F4NUMORD = '" & Trim(rstRecepcionOC!NROOC & "") & "' AND "
                    CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstRecepcionOC!NroPedido & "") & "' AND "
                    CadSql = CadSql & "CODPROD = '" & Trim(rstRecepcionOC!CodProducto & "") & "' AND "
                    CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstRecepcionOC!CODPRODUCTOORIGINAL & "") & "'"
                End If
            End With
            
            cnDBTemp.Execute CadSql
            
            rstRecepcionOC.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstRecepcionOC.State = 1 Then rstRecepcionOC.Close
    
    Set rstRecepcionOC = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionDevolucionOP()
    Dim rstDevolucionOP As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILDEVOLUCIONOP "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE"
    
    If rstDevolucionOP.State = 1 Then rstDevolucionOP.Close
    
    rstDevolucionOP.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDevolucionOP.EOF Then
        rstDevolucionOP.MoveFirst
            
        Do While Not rstDevolucionOP.EOF
            With objAyudaOrden
                .inicializarEntidadesDetalle
                
                If ModUtilitario.ObtenerCampoV2(cnDBTemp, "TRIM(CODPROD & '')", "TMPVALEINGRESO", "TRIM(CODPROD & '')", Trim(rstDevolucionOP!CodProducto & ""), "T", _
                                                "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstDevolucionOP!NroPedido & "") & "' " & _
                                                "AND TRIM(F4NUMORD & '') = '" & Trim(rstDevolucionOP!NROOC & "") & "'") = vbNullString Then
                             
                    dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                    
                    CadSql = vbNullString
                    CadSql = CadSql & "INSERT INTO TMPVALEINGRESO(ITEM, F4NUMORD, COD_SOLICITUD, "
                    CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, CANTIDAD, "
                    CadSql = CadSql & "CANTIDADMAX) "
                    'CadSql = CadSql & "AFECTO, COSTOUNI, PVUNIT, VVTOTAL, IGV, TOTAL) "
                    CadSql = CadSql & "VALUES("
                    CadSql = CadSql & dblItem & ", "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!NROOC & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!NroPedido & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!CodProducto & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!CODPRODUCTOORIGINAL & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!NOMPRODUCTO & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!um & "") & "', "
                    CadSql = CadSql & Val(rstDevolucionOP!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & Val(rstDevolucionOP!CANTIDADDESTINO & "")
                    'CadSql = CadSql & "'" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & Val(rstDevolucionOP!PRECIOUNITARIO & "") & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & Val(Format(Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000"))
                    CadSql = CadSql & ")"
                Else
                    CadSql = vbNullString
                    CadSql = CadSql & "UPDATE TMPVALEINGRESO "
                    CadSql = CadSql & "SET "
                    CadSql = CadSql & "CANTIDAD = " & Val(rstDevolucionOP!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "CANTIDADMAX = " & Val(rstDevolucionOP!CANTIDADDESTINO & "") & " "
                    'CadSql = CadSql & "AFECTO = '" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & "COSTOUNI = " & Val(rstDevolucionOP!PRECIOUNITARIO & "") & ", "
                    'CadSql = CadSql & "PVUNIT = " & Val(Format(Val(rstDevolucionOP!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & "VVTOTAL = " & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & "IGV = " & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & "TOTAL = " & Val(Format(Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000")) & " "
                    CadSql = CadSql & "WHERE "
                    CadSql = CadSql & "F4NUMORD = '" & Trim(rstDevolucionOP!NROOC & "") & "' AND "
                    CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstDevolucionOP!NroPedido & "") & "' AND "
                    CadSql = CadSql & "CODPROD = '" & Trim(rstDevolucionOP!CodProducto & "") & "' AND "
                    CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstDevolucionOP!CODPRODUCTOORIGINAL & "") & "'"
                End If
            End With
            
            cnDBTemp.Execute CadSql
            
            rstDevolucionOP.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstDevolucionOP.State = 1 Then rstDevolucionOP.Close
    
    Set rstDevolucionOP = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionDevolucionOPSql()
    Dim rstDevolucionOP As New ADODB.Recordset
    Dim dblItem As Double
    
    Me.MousePointer = vbHourglass
                
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM " 'tmpCPDevolucionOp
    CadSql = CadSql & "TMPCPDEVOLUCIONOP" & UCase(wusuario) & " "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = 1"
    
    If rstDevolucionOP.State = 1 Then rstDevolucionOP.Close
    
    rstDevolucionOP.Open CadSql, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDevolucionOP.EOF Then
        rstDevolucionOP.MoveFirst
            
        Do While Not rstDevolucionOP.EOF
            With objAyudaOrden
                .inicializarEntidadesDetalle
                
                'AND TRIM(F4NUMORD & '') = '" & Trim(rstDevolucionOP!NROOC & "") & "'
                
                If ModUtilitario.ObtenerCampoV2(cnDBTemp, "TRIM(CODPROD & '')", "TMPVALEINGRESO", "TRIM(CODPROD & '')", Trim(rstDevolucionOP!CodProducto & ""), "T", _
                                                "AND TRIM(COD_SOLICITUD & '') = '" & Trim(rstDevolucionOP!NroPedido & "") & "' " & _
                                                "") = vbNullString Then
                    
                    dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1
                    
                    CadSql = vbNullString
                    CadSql = CadSql & "INSERT INTO TMPVALEINGRESO(ITEM, F4NUMORD, COD_SOLICITUD, "
                    CadSql = CadSql & "CODPROD, CODPRODORIGINAL, DESCRIPCION, UMEDIDA, CANTIDAD, "
                    CadSql = CadSql & "CANTIDADMAX) "
                    'CadSql = CadSql & "AFECTO, COSTOUNI, PVUNIT, VVTOTAL, IGV, TOTAL) "
                    CadSql = CadSql & "VALUES("
                    CadSql = CadSql & dblItem & ", "
                    'CadSql = CadSql & "'" & Trim(rstDevolucionOP!NROOC & "") & "', "
                    CadSql = CadSql & "'', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!NroPedido & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!CodProducto & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!CODPRODUCTOORIGINAL & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!NOMPRODUCTO & "") & "', "
                    CadSql = CadSql & "'" & Trim(rstDevolucionOP!um & "") & "', "
                    CadSql = CadSql & Val(rstDevolucionOP!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & Val(rstDevolucionOP!CANTIDADDESTINO & "")
                    'CadSql = CadSql & "'" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & Val(rstDevolucionOP!PRECIOUNITARIO & "") & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & Val(Format(Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000"))
                    CadSql = CadSql & ")"
                Else
                    CadSql = vbNullString
                    CadSql = CadSql & "UPDATE TMPVALEINGRESO "
                    CadSql = CadSql & "SET "
                    CadSql = CadSql & "CANTIDAD = " & Val(rstDevolucionOP!CANTIDADDESTINO & "") & ", "
                    CadSql = CadSql & "CANTIDADMAX = " & Val(rstDevolucionOP!CANTIDADDESTINO & "") & " "
                    'CadSql = CadSql & "AFECTO = '" & IIf(.Afecto, "*", vbNullString) & "', "
                    'CadSql = CadSql & "COSTOUNI = " & Val(rstDevolucionOP!PRECIOUNITARIO & "") & ", "
                    'CadSql = CadSql & "PVUNIT = " & Val(Format(Val(rstDevolucionOP!PRECIOUNITARIO & "") * (1 + wwigv / 100), "#0.0000")) & ", "
                    'CadSql = CadSql & "VVTOTAL = " & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) & ", "
                    'CadSql = CadSql & "IGV = " & Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (wwigv / 100) & ", "
                    'CadSql = CadSql & "TOTAL = " & Val(Format(Val(Format(Val(rstDevolucionOP!CANTIDADDESTINO & "") * Val(rstDevolucionOP!PRECIOUNITARIO & ""), "#0.0000")) * (1 + wwigv / 100), "#0.0000")) & " "
                    CadSql = CadSql & "WHERE "
                    CadSql = CadSql & "F4NUMORD = '" & Trim(rstDevolucionOP!NROOC & "") & "' AND "
                    CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstDevolucionOP!NroPedido & "") & "' AND "
                    CadSql = CadSql & "CODPROD = '" & Trim(rstDevolucionOP!CodProducto & "") & "' AND "
                    CadSql = CadSql & "CODPRODORIGINAL = '" & Trim(rstDevolucionOP!CODPRODUCTOORIGINAL & "") & "'"
                End If
            End With
            
            cnDBTemp.Execute CadSql
            
            rstDevolucionOP.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstDevolucionOP.State = 1 Then rstDevolucionOP.Close
    
    Set rstDevolucionOP = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub listarGrilla()
    abrirCnTemporal
    
    With dxDBGrid1.Dataset
        .Active = False
        .ADODataset.ConnectionString = cnDBTemp  'CnTmp
        .ADODataset.CommandText = "SELECT * FROM TMPVALEINGRESO ORDER BY ITEM, DESCRIPCION"
        .Active = True
        
        dxDBGrid1.KeyField = "ITEM"
        
        .Close
        .Open
    End With
    
    txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
    txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
    txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
End Sub

Private Sub verificarProductoDevueltoOP(ByVal strCodAlmacen As String, _
                                        ByVal strNumeroVale As String)
                                        
    Dim rstValeDet As New ADODB.Recordset
    Dim dblCantidadOP As Double
    
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    rstValeDet.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & strCodAlmacen & "' AND F4NUMVAL = '" & strNumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeDet.EOF Then
        rstValeDet.MoveFirst
        
        'abrirCnDBMilano
        
        Do While Not rstValeDet.EOF
            If ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDINSUMO", "ORDENPRODUCCIONDESCARGO", "IDINSUMO", Trim(rstValeDet!f5codpro & ""), "T", "AND IDORDENPRODUCCION = " & Trim(txtIDOrdenProduccion.Text)) <> vbNullString Then
                dblCantidadOP = Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CANTIDAD", "ORDENPRODUCCIONDESCARGO", "IDINSUMO", Trim(rstValeDet!f5codpro & ""), "T", "AND IDORDENPRODUCCION = " & Trim(txtIDOrdenProduccion.Text)))
                
                If Val(rstValeDet!F3CANPRO & "") < dblCantidadOP Then
                    Call ModMilano.modificarProductoEnOP(Trim(txtIDOrdenProduccion.Text), Trim(rstValeDet!f5codpro & ""), Trim(rstValeDet!f5codpro & ""), dblCantidadOP, dblCantidadOP - Val(rstValeDet!F3CANPRO & ""), "DEVOLUCION DE OP - PROCESO AUTOMATICO DE CP PARA AJUSTE DE CANTIDAD CONSUMIDA REALMENTE EN OP.")
                End If
            End If
            
            rstValeDet.MoveNext
        Loop
    End If
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
        rstValeDet.Open "SELECT F4NUMORD FROM TMPVALEINGRESO GROUP BY F4NUMORD", cnDBTemp, adOpenForwardOnly, adLockReadOnly
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
            End With
            
            rstValeDet.MoveNext
        Loop
    End If
End Sub

Private Sub calcularDescuentoTotal()
    Dim dblSubTotal As Double
    Dim dblDsctoTotal As Double
    Dim dblPorcentajeDscto As Double
    
    dblSubTotal = Val(Format(txtTotvv.Text, "#0.00"))
    dblDsctoTotal = Val(Format(txtDscto.Text, "#0.00"))
    
    If dblDsctoTotal > 0 Then
        dblPorcentajeDscto = Val(Format(dblDsctoTotal / dblSubTotal, "#0.0000"))
    Else
        dblPorcentajeDscto = 0
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPVALEINGRESO "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "PORDESC = " & Val(Format(dblPorcentajeDscto * 100, "#0.00"))
    
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad
    
    txtDscto.Text = "0.00"
    
    listarGrillaVale
    
    recalcularItems
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
    Dim WMONEDA As String * 1
    
    Me.MousePointer = vbHourglass
    
    ccod_almacen = right(txtAlmacen.Text, 2)
    cnum_vale = Trim(txtnumero.Text)
    costo = Trim(txtccosto.Text)
    
        WMONEDA = IIf(cmbmoneda.ListIndex = 0, "S", "D")
        If WMONEDA = "D" Then
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALDOL AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valdol] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE (((IF4VALES.F2CODALM)='" & txtAlmacen.Text & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "') AND ((IF4VALES.F1CODORI)='" & txtconcepto.Text & "')) " & _
                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        Else
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALVTA AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valvta] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE (((IF4VALES.F2CODALM)='" & txtAlmacen.Text & "') AND ((IF4VALES.F4NUMVAL)='" & cnum_vale & "') AND ((IF4VALES.F1CODORI)='" & txtconcepto.Text & "')) " & _
                " ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        End If
    If Tipo = 1 Then
        With acr_vales
        .DataControl1.ConnectionString = cnn_dbbancos
        If left(txtnumero.Text, 1) = "I" Then
            prov = Trim(txtproveedor.Text)
            ctipo_vale = "I"
            .Lbl_vale.Caption = " VALE DE INGRESO "
            .lblprov.Visible = True
            .lblpunto.Visible = True
            .fldprov.Visible = True
            .lblpie2.Caption = "Entregado por"
        Else
            ctipo_vale = "S"
            .Lbl_vale.Caption = " VALE DE SALIDA "
            .lblprov.Visible = False
            .lblpunto.Visible = False
            .fldprov.Visible = False
            .lblpie2.Caption = "Hecho por"
        End If
        
        .DataControl1.Source = csql
        .fldnomprov.Text = txtnomprov.Text
        .fldAlmacen.Text = Mid(cmbalmacen.Text, 200)
        .fldnomcosto.Text = pnlccosto.Caption
        .fldempresa.Text = wnomcia
        .fldFecha.Text = Format(Date, "dd/mm/yyyy")
        .fldvale.Text = cnum_vale
        .fldalma.Text = ccod_almacen
       
        .Show vbModal
    End With

    Else
        With acr_vales_p
        .DataControl1.ConnectionString = cnn_dbbancos
        If left(txtnumero.Text, 1) = "I" Then
            prov = Trim(txtproveedor.Text)
            ctipo_vale = "I"
            .Lbl_vale.Caption = " VALE DE INGRESO "
            .lblprov.Visible = True
            .lblpunto.Visible = True
            .fldprov.Visible = True
            .lbldocum.Caption = cmbtipo.Text
            .lblmoneda.Caption = cmbmoneda.Text
            If cmbmoneda.Text = "Soles" Then
                .Lblsigno.Caption = "S/" '"S/"
            Else
                .Lblsigno.Caption = "US$"
            End If
        Else
            ctipo_vale = "S"
            .Lbl_vale.Caption = " VALE DE SALIDA "
            .lblprov.Visible = False
            .lblpunto.Visible = False
            .fldprov.Visible = False
            .lblpie2.Caption = "Hecho por"
        End If
        
        .DataControl1.Source = csql
        .fldnomprov.Text = txtnomprov.Text
        .fldAlmacen.Text = left(cmbalmacen.Text, 50)
        .fldnomcosto.Text = pnlccosto.Caption
        .fldempresa.Text = wnomcia
        .fldFecha.Text = Format(Now, "dd/mm/yyyy hh:mm")
        .fldvale.Text = cnum_vale
        .fldalma.Text = ccod_almacen
        
        .Show vbModal
    End With

    End If
Me.MousePointer = vbDefault
End Sub

Private Sub SUMA_CANT_OCOMPRA(pnumorden As Double, palmacen As String, pnumvale As String)
    
    If rsif3vales.State = adStateOpen Then rsif3vales.Close
    rsif3vales.Open "SELECT * FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsif3vales.EOF Then
        Do While Not rsif3vales.EOF
            sql = ("UPDATE IF3ORDEN SET F3CANFAL = F3CANFAL + " & rsif3vales.Fields("F3CANPRO") & " WHERE F4NUMORD=" & pnumorden & " AND F3CODPRO = '" & rsif3vales.Fields("F5CODPRO") & "'") 'rstempo.Fields("CODPROD") & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            rsif3vales.MoveNext
        Loop
    End If
    rsif3vales.Close
    
End Sub

'Private Function VALIDA_PROVEEDOR(pproveedor As String)
'Dim sw_e    As Boolean
'
'    If RsProveedor.State = adStateOpen Then RsProveedor.Close
'    RsProveedor.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & pproveedor & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    If Not RsProveedor.EOF Then
'        wnomprov = Trim(RsProveedor.Fields("F2NOMPROV") & "")
'        sw_e = True
'    Else
'        sw_e = False
'    End If
'    RsProveedor.Close
'    VALIDA_PROVEEDOR = sw_e
'
'End Function

Private Function VALIDA_CONCEPTO_INV(pconcepto As String)
    Dim sw_e    As Boolean
    
    If rsconcepto_inv.State = adStateOpen Then rsconcepto_inv.Close
    
    rsconcepto_inv.Open "SELECT F1NOMORI,F1PARTIDA FROM SF1ORIGENES WHERE F1CODORI='" & pconcepto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rsconcepto_inv.EOF Then
        wnomconcepto = Trim(rsconcepto_inv.Fields("F1NOMORI") & "")
        wpartida = Trim(rsconcepto_inv.Fields("F1PARTIDA") & "")
        sw_e = True
    Else
        sw_e = False
    End If
    
    rsconcepto_inv.Close
    
    VALIDA_CONCEPTO_INV = sw_e
End Function

Private Sub habilita_conceptos(codalma As String)
    cmbconcepto.Clear
    txtconcepto.Text = ""
    
    If Rs.State = adStateOpen Then Rs.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "ORI.F1CODORI, "
    SqlCad = SqlCad & "ORI.F1NOMORI, "
    SqlCad = SqlCad & "ORI.F1COSTO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ALMACEN_CONCEPTO AS AC "
    SqlCad = SqlCad & "LEFT JOIN SF1ORIGENES AS ORI "
    SqlCad = SqlCad & "ON ORI.F1CODORI = AC.F1CODORI "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "AC.F2CODALM = '" & codalma & "' AND "
    SqlCad = SqlCad & "ORI.F1TIPMOV = 'I' AND "
    SqlCad = SqlCad & "ORI.F1CODORI NOT IN ('XCS') "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "ORI.F1NOMORI "
    
    Rs.Open SqlCad, cnn_dbbancos, adOpenStatic, adLockReadOnly
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        Do While Not Rs.EOF
            'cad = Space(255)
            'Mid(cad, 1) = Trim(Rs!F1NOMORI & "")
            'Mid(cad, 200) = Trim(Rs!F1CODORI & "")
            'Mid(cad, 203) = Trim(Rs!F1COSTO & "")
            
            cmbconcepto.AddItem Trim(Rs!F1NOMORI & "") & Space(100) & Trim(Rs!F1CODORI & "")
            
            Rs.MoveNext
        Loop
    End If
    
    Rs.Close
End Sub

Private Sub abofecha_Change()
     sw_cabecera = True
End Sub

Private Sub abofecha_CloseUp()
    If IsDate(abofecha.Value) Then
'        If rscambios.State = adStateOpen Then rscambios.Close
'        If ctipoadm_bd = "M" Then
'            sql = "SELECT * FROM CAMBIOS WHERE FECHA='" & abofecha.value & "'"
'        Else
'            sql = "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.value & "')"
'        End If
'        rscambios.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If sw_nuevo_documento = True Then
'            If Not rscambios.EOF And sw_nuevo_documento = True Then
'                txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
'            Else
'                txttc.Text = Format(3.64, "0.000")
'            End If
'        End If
'        rscambios.Close
        txttc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", abofecha.Value, "F")), "#.000")
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
    
    If Trim(txtAlmacen.Text) <> vbNullString Then
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(txtAlmacen.Text)
            
            .FechaInicioMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 0, 1)
            .FechaFinMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 1, 0)
            
            If .verificarCierreVale Then
                MsgBox "Imposible registrar Vale, periodo ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                abofecha.SetFocus
            End If
        End With
    End If
End Sub

Private Sub chkVerObservaciones_Click()
    dxDBGrid1.Columns.ColumnByFieldName("OBSERVACIONES").Visible = CBool(chkVerObservaciones.Value)
End Sub

Private Sub cmbalmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbalmacen_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        abofecha.SetFocus
'    End If
    'KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
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
        'lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "CATEGORIATIPO", "NOMBRE", Trim(cmbCategoriaTipo.Text), "T")
        
        If Trim(lblIdCategoriaTipo.Caption) = vbNullString Then
            MsgBox "Categoria no identificada.", vbInformation + vbOKOnly, App.ProductName
            
            cmbCategoriaTipo.SetFocus
        End If
    End If
End Sub

Private Sub cmbconcepto_Change()
    sw_cabecera = True
End Sub

Private Sub cmbconcepto_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'    End If
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub cmbmarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtserie.SetFocus
End Sub

Private Sub cmbmoneda_Change()
    sw_cabecera = True
End Sub

Private Sub Cmbmoneda_Click()
    If cmbmoneda.Text = "Dolares" Then
        cmbmoneda.BackColor = &HC0FFC0
        txtDscto.BackColor = &HC0FFC0
        txtTotigv.BackColor = &HC0FFC0
        txtTotpv.BackColor = &HC0FFC0
        txtTotvv.BackColor = &HC0FFC0
        Lbl1.Caption = "US$"
        Label21.Caption = "US$"
        Lbl2.Caption = "US$"
        Lbl3.Caption = "US$"
        Lbl1.ForeColor = &HC0FFC0
        Lbl2.ForeColor = &HC0FFC0
        Lbl3.ForeColor = &HC0FFC0
    ElseIf cmbmoneda.Text = "Soles" Then
        cmbmoneda.BackColor = &HC0FFFF
        txtDscto.BackColor = &HC0FFFF
        txtTotigv.BackColor = &HC0FFFF
        txtTotpv.BackColor = &HC0FFFF
        txtTotvv.BackColor = &HC0FFFF
        Lbl1.Caption = "S/" '"S/"
        Label21.Caption = "S/" '"S/"
        Lbl2.Caption = "S/" '"S/"
        Lbl3.Caption = "S/" '"S/"
        Lbl1.ForeColor = &HC0FFFF
        Lbl2.ForeColor = &HC0FFFF
        Lbl3.ForeColor = &HC0FFFF
    Else
        cmbmoneda.BackColor = vbRed
    End If
End Sub

Private Sub cmbmoneda_keypress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
''        If txtuupp.Visible = True Then
''            txtuupp.SetFocus
''        Else
''            txtccosto.SetFocus
''           ' TxtSerie.SetFocus
''        End If
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'    End If
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub cmbAlmacen_Click()
    txtAlmacen.Text = right(cmbalmacen.Text, 2)
    wcod_alm = txtAlmacen.Text
    
    Rem SK ADD:
    habilita_conceptos txtAlmacen.Text
End Sub

Private Sub cmbConcepto_Click()
    On Error Resume Next
    
    With objAyudaOrigen
        .inicializarEntidades
        
        .Codigo = right(cmbconcepto.Text, 3)  'Trim(Mid(cmbconcepto.Text, 200, 3))
        
        If .obtenerOrigen Then
            txtconcepto.Text = .Codigo
            
            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").DisableEditor = Not .RegistrarCosto
                dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Color = IIf(.RegistrarCosto, &H80000005, &HE0E0E0)
                
            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").DisableEditor = Not .RegistrarCosto
                dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Color = IIf(.RegistrarCosto, &H80000005, &HE0E0E0)
                
            dxDBGrid1.Columns.ColumnByFieldName("PORDESC").DisableEditor = Not .RegistrarCosto
                dxDBGrid1.Columns.ColumnByFieldName("PORDESC").Color = IIf(.RegistrarCosto, &H80000005, &HE0E0E0)
            
            'dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = .RegistrarCosto
            
            'dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = .RegistrarCosto
            
            SSFrame2.Visible = .RegistrarCosto
            
            cmbconcepto.SetFocus
        End If
    End With
    
'    Dim csql        As String
'    Dim cf1costo    As String
'
'    If right(RTrim(cmbconcepto.Text), 1) = "*" Then
'        cf1costo = right(RTrim(cmbconcepto.Text), 1)
'        X = 1
'    End If
'
'    Select Case cf1costo
'        Case "*":
'            If UCase(left(cmbconcepto.Text, 6)) = "COMPRA" Then compra = True
'
'            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").ColIndex = 7
'            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ColIndex = 10
'            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").ColIndex = 8
'            dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").ColIndex = 9
'        Case "1":
'            compra = False
'            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Visible = True
'            dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = True
'
'        Case Else
'            compra = False
'            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = False
'
'    End Select
'
'    txtconcepto.Text = Trim(Mid(cmbconcepto.Text, 200, 3))
End Sub

Private Sub cmbtipo_Change()
    sw_cabecera = True
End Sub

Private Sub cmbtipo_Click()
    If cmbtipo.ListIndex = -1 Then Exit Sub
    
    If cnDBTemp Is Nothing Then Exit Sub
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(CODPROD)", "TMPVALEINGRESO", "", "", "", "TRIM(CODPROD & '') <> ''")) > 0 Then
        recalcularItems
    End If
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        If Trim(cmbtipo.Text) <> "" And sw_cabecera = False Then
'            sw_cabecera = True
'        End If
'
'        txtserfac.SetFocus
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

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


Private Sub cmdimporta_Click()

    hlp_importaciones.Show vbModal
    If hlp_importaciones.DataGrid.SelBookmarks.Count > 0 Then
        For x = 0 To hlp_importaciones.DataGrid.SelBookmarks.Count - 1
            hlp_importaciones.DataGrid.Bookmark = hlp_importaciones.DataGrid.SelBookmarks.ITEM(x)
            BuscaImportacion (Trim(hlp_importaciones.DataGrid.Columns(0)))
        Next
    End If
    'Unload hlp_importaciones

End Sub

Private Sub cmdoc_Click()

End Sub

Private Sub dtpFechaDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    'If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> '' ORDER BY ITEM DESC") & "") + 1 'dxDBGrid1.Dataset.RecordCount + 1
            
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
        .Set (egoExpandOnDblClick)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
'    dxDBGrid1.Columns.ColumnByFieldName("ITEM").Visible = True
'    dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Visible = False
'    dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterType = cstSum
'    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterType = cstSum
End Sub

Private Sub AdicionaItem2()
Dim sw_nuevo_temp   As Boolean
Dim I               As Integer

    dxDBGrid1.Dataset.Active = False
    'If sw_nuevo_documento = False Then
        DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
        dxDBGrid1.Dataset.Refresh
    'End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
    
        For I = 1 To 1
            'If sw_nuevo_temp = False Then
            If sw_nuevo_documento = False Then
                If sw_nuevo_documento = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("ITEM") = I
            .FieldValues("CODPROD") = ""
            .FieldValues("CODFAB") = ""
            .FieldValues("marca") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UMEDIDA") = ""
            .FieldValues("CANTIDAD") = Format(0, "###,##0.0000")
            .FieldValues("COSTOUNI") = Format(0, "###,##0.0000")
            .FieldValues("IGV") = Format(0, "###,##0.00")
            .FieldValues("TOTAL") = Format(0, "###,##0.00")
            Cantidad = .FieldValues("CANTIDAD")
        Next
        .Post
        sw_nuevo_item = False
    End With
    
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
End Sub

Private Sub AdicionaItem()
Dim x As Double
Dim sw_nuevo_temp   As Boolean
Dim I               As Integer

    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_documento = False Then
        abrirCnTemporal
        
        DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
        
        dxDBGrid1.Dataset.Refresh
    End If

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
    dxDBGrid1.Dataset.ADODataset.CommandText = "SELECT * FROM TMPVALEINGRESO"
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    dxDBGrid1.Dataset.Open
    
    
    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
    
        For I = 1 To 1
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
            .FieldValues("ITEM") = I
            .FieldValues("CODPROD") = ""
            .FieldValues("CODFAB") = ""
            .FieldValues("MARCA") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UMEDIDA") = ""
            .FieldValues("CANTIDAD") = Format(0, "###,##0.0000")
            .FieldValues("COSTOUNI") = Format(0, "###,##0.0000")
            .FieldValues("PVUNIT") = Format(0, "###,##0.0000")
            .FieldValues("VVTOTAL") = Format(0, "###,##0.00")
            .FieldValues("TOTAL") = Format(0, "###,##0.00")

            
            .FieldValues("AFECTO") = ""
            
        Next
        
        .Post
        
        sw_nuevo_item = False
        
    End With
'    For X = 1 To 999999
'    Next
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
End Sub

Private Sub abofecha_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        cmbconcepto.SetFocus
'    End If

End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    'If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "")) = 0 Then
                    Allow = False
                Else
                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CODPROD").ColIndex
                End If
            End If
        End If
        If Action = daDelete Then
            sw_detalle = True
            dxDBGrid1.Dataset.Refresh
        End If
    'End If

End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case Column.Caption
    Case "Igv", "Total": Text = Format(Text, "#,###,###0.00") ', "P.V. Uni." se saca para dar formato
    Case "Costo Unit.": Text = Format(Text, "#,###,###0.0000")
    Case "Costo Total": Text = Format(Text, "#,###,###0.00")
    Case "P.V. Uni.": Text = Format(Text, "#,###,###0.0000")
End Select
End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case Column.FieldName
        Case "CODFAB", "CODPROD"
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
                
                If .RegistrarCosto Then
                    If Trim(txtproveedor.Text) = vbNullString Then
                        If MsgBox("No se ha seleccionado el Proveedor, ¿Desea continuar?", vbInformation + vbYesNo, App.ProductName) = vbNo Then
                            
                            .inicializarEntidades
                            
                            Exit Sub
                        End If
                    End If
                End If
            
                Select Case .CodigoAyudaProducto
                    Case "1"
                        If ModUtilitario.validarFormAbierto("ayuda_productos") Then
                            Unload ayuda_productos
                        End If
                        
                        With ayuda_productos
                            .CodigoAuxiliar = IIf(objAyudaOrigen.RegistrarCosto, Trim(txtproveedor.Text), vbNullString)
                            .CodigoRequerimiento = vbNullString
                            .CodigoProducto = vbNullString
                            
                            .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                            
                            .Show 1
                        End With
                        
                        abrirCnTemporal
                        
                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPPRODUCTOS", "F4PERINT", "-1", "N") & "") <> 0 Then
                            copiarSeleccionAyudaProductos
                            
                            SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                        End If
                    Case "2"
                        If ModUtilitario.validarFormAbierto("frmUtilStockDisponible") Then
                            Unload frmUtilStockDisponible
                        End If
                        
                        With frmUtilStockDisponible
                            .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                            
                            .Show 1
                            
                            cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                            cmbalmacen.Enabled = False
                        End With
                        
                        abrirCnTemporal
                        
                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILSTOCKDISPONIBLE", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                            copiarSeleccionStockDisponible
                            
                            SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                        End If
                    Case Else
                        MsgBox "Ayuda de Concepto de Movimiento no configurado, verifique.", vbInformation + vbOKOnly, App.ProductName
                End Select
            End With
            
            listarGrilla
            
            Me.MousePointer = vbDefault
    End Select
'    Dim cad As String
'
'    If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODFAB" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CODPROD" Then
'        wcod_alm = txtalmacen.Text
'        wcodproducto = ""
'        sw_ayuda_prod = True
'        wmarca = Trim(Mid(cmbmarca.Text, 200))
'        cad = ""
'        If Not wmarca = "" Then
'            cad = " and a.f5marca='" & wmarca & "'"
'        End If
'        Me.MousePointer = vbhourglass
'        Con_Ayu = 1
'        ayuda_productos.Show 1
'        Me.MousePointer = vbdefault
'        '-----
'        With ayuda_productos.dbgProducto
'            .Dataset.Filtered = True
'            .Dataset.Filter = "F4PERINT = -1"
'            .Dataset.First
'            X = 0
'                Do While Not .Dataset.EOF
'                    z = .Dataset.RecordCount
'                    If z = 0 Then Exit Sub
'                    X = X + 1
'                    If dxDBGrid1.Columns.ColumnByFieldName("COD_PRODUCTO").Value = "" Then
'                        dxDBGrid1.Dataset.Edit
'                    Else
'                        dxDBGrid1.Dataset.Append
'                    End If
'                    dxDBGrid1.Dataset.Edit
'                    dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = .Columns.ColumnByFieldName("f5codpro").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").Value = .Columns.ColumnByFieldName("f5codpro").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = .Columns.ColumnByFieldName("f5codpro").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = ""
'                    dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = .Columns.ColumnByFieldName("f5nompro").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = .Columns.ColumnByFieldName("F7SIGMED").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value = IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value)
'                    dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Value = "*"
'                    dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = .Columns.ColumnByFieldName("f5fob").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value = Format(IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value) * 1.18, "0.00")
'                    dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = Format(.Columns.ColumnByFieldName("f5fob").Value * IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value), "0.00")
'                    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = Format(.Columns.ColumnByFieldName("f5fob").Value * IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value) * 1.18, "0.00")
'                     dxDBGrid1.Dataset.Post
'                    .Dataset.Next
'                    Loop
'                    If X = 0 And Len(Trim(wcodproducto)) > 0 Then
'                        dxDBGrid1.Dataset.Edit
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = wcodproducto
'                        dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").Value = wcodproducto
'                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = wcodfab
'                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
'                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = wdesproducto
'                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = wmedida
'                        dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value = IIf(cmbmoneda.Text = "Soles", wprecos, wprecosdol)
'                        dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Value = wafecto
'                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
'                        dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value = 0#
'                        dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = 0#
'                        dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = 0#
'                        dxDBGrid1.Dataset.Post
'                    End If
'                Unload ayuda_productos
'            End With
'            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
'    End If

    Select Case Column.ObjectName
        Case "ColBotonEliminar"
            If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                sw_nuevo_item = True
                
                dxDBGrid1.Dataset.Delete
                
                If dxDBGrid1.Count = 0 Then
                    AdicionaItem
                    
                    sw_detalle = False
                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                End If
                
                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
                
                sw_nuevo_item = False
            End If
    End Select
    
    Rem SK ADD:
    Select Case Column.Caption
        Case "?" 'Productos Alternos
            With frmListaBienAlterno
                Select Case Trim(txtconcepto.Text)
                    Case "XC0"
                        If Trim(txtOcompra.Text) = vbNullString Then
                            MsgBox "Opción solo aplica para Ingresos por Ordenes de Compra generadas por el sistema.", vbInformation, "Sistema de Logística"
                            
                            'txtocompra.SetFocus
                            
                            Exit Sub
                        End If
                    Case "XOP"
                        If Val(txtIDOrdenProduccion.Text) = 0 Then
                            MsgBox "Opción solo aplica para Ingresos por Devolución de OP.", vbInformation, "Sistema de Logística"
                            
                            Exit Sub
                        End If
                End Select
                
                If dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").Value = vbNullString Then
                    MsgBox "Opción no aplica para Item seleccionado.", vbInformation + vbOKOnly, App.ProductName
                    
                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").ColIndex
                    
                    Exit Sub
                End If
                
                If ModUtilitario.validarFormAbierto("frmListaBienAlterno") Then
                    Unload frmListaBienAlterno
                End If
                
                objAyudaBienAlterno.inicializarEntidades
                
                .Ayuda = True
                .CodigoBien = Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").Value & "")
                
                .Show vbModal
                
                If objAyudaBienAlterno.CodigoBienAlterno <> vbNullString Then
                    With dxDBGrid1.Dataset
                        .Edit
                        
                        .FieldValues("CODPROD") = objAyudaBienAlterno.CodigoBienAlterno
                        .FieldValues("DESCRIPCION") = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", objAyudaBienAlterno.CodigoBienAlterno, "T")
                        
                        .Post
                        
                        dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").ColIndex
                    End With
                End If
                
                objAyudaBienAlterno.inicializarEntidades
                
                .Ayuda = False
            End With
        Case "H" 'Historial de Precios de Producto
            If Trim(txtconcepto.Text) <> "XC0" Then
                MsgBox "Opción solo aplica para Ingresos por Compra.", vbInformation, "Sistema de Logística"
                
                Exit Sub
            End If
            
            If Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "") = vbNullString Then
                MsgBox "Producto no especificado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("ayuda_prov_prod") Then
                Unload ayuda_prov_prod
            End If
            
            With ayuda_prov_prod
                objAyudaOrden.inicializarEntidades
                objAyudaOrden.inicializarEntidadesDetalle
                
                .CodigoProducto = Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "")
                
                .Show 1
            End With
            
            With objAyudaOrden
                If .PrecioSinImpuesto > 0 Then
                    'Grid.Dataset.Edit
                    
                    Select Case .CodMoneda
                        Case "S"
                            .PrecioSinImpuesto = Val(Format(.PrecioSinImpuesto / IIf(left(cmbmoneda.Text, 1) = "S", 1, Val(txttc.Text)), "#0.0000"))
                        Case Else
                            .PrecioSinImpuesto = Val(Format(.PrecioSinImpuesto * IIf(left(cmbmoneda.Text, 1) = "D", 1, Val(txttc.Text)), "#0.0000"))
                    End Select
                    
                    Dim rstTemporalEditButtonI As New ADODB.Recordset
                    
                    If rstTemporalEditButtonI.State = 1 Then rstTemporalEditButtonI.Close
                    
                    rstTemporalEditButtonI.Open "SELECT * FROM TMPVALEINGRESO WHERE CODPROD = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "") & "'", cnDBTemp, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstTemporalEditButtonI.EOF Then
                        rstTemporalEditButtonI.MoveFirst
                        
                        dxDBGrid1.Dataset.Close
                        
                        Do While Not rstTemporalEditButtonI.EOF
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE "
                            SqlCad = SqlCad & "TMPVALEINGRESO "
                            SqlCad = SqlCad & "SET "
                            SqlCad = SqlCad & "COSTOUNI = " & .PrecioSinImpuesto & ", "
                            SqlCad = SqlCad & "PORDESC = " & .PorcentajeDscto & " "
                            SqlCad = SqlCad & "WHERE "
                            SqlCad = SqlCad & "CODPROD = '" & Trim(rstTemporalEditButtonI!codprod & "") & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            rstTemporalEditButtonI.MoveNext
                        Loop
                            listarGrillaVale
                            
                            recalcularItems
                    End If
                End If
            End With
    End Select
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Rem SK ADD:
    Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
        Case "CANTIDAD", "COSTOUNI", "PVUNIT", "PORDESC", "VALDESC", "AFECTO"
            With objAyudaOrden
                'DATOS
                .PorcentajeImpuesto = IIf(right(cmbtipo.Text, 2) = "02", gretenc, IIf(right(cmbtipo.Text, 2) = "03", 0, wwigv)) / 100
                .SignoImpuesto = IIf(right(cmbtipo.Text, 2) = "02", -1, 1)
                
                .Cantidad = Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value & "")
                .CantidadMaxima = Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDADMAX").Value & "")
                .PorcentajeDemasia = 0
                
                Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
                    Case "COSTOUNI"
                        .PrecioSinImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "")
                        .PrecioConImpuesto = 0
                    Case "PVUNIT"
                        .PrecioSinImpuesto = 0
                        .PrecioConImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value & "")
                    Case Else
                        .PrecioSinImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "")
                        .PrecioConImpuesto = Val(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value & "")
                End Select
                
                Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
                    Case "PORDESC"
                        .PorcentajeDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("PORDESC").Value & "") / 100
                        .TotalDscto = 0
                    Case "VALDESC"
                        .PorcentajeDscto = 0
                        .TotalDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("VALDESC").Value & "")
                    Case Else
                        .PorcentajeDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("PORDESC").Value & "") / 100
                        .TotalDscto = Val(dxDBGrid1.Columns.ColumnByFieldName("VALDESC").Value & "")
                End Select
                
                If right(cmbtipo.Text, 2) = "03" Then
                    .Afecto = False
                Else
                    If Trim(dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Value & "") = "*" Then
                        .Afecto = True
                    Else
                        .Afecto = False
                    End If
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
                
                dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value = .PrecioSinImpuesto
                dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value = .PrecioConImpuesto
                dxDBGrid1.Columns.ColumnByFieldName("PORDESC").Value = Val(Format(.PorcentajeDscto * 100, "#0.00"))
                'dxDBGrid1.Columns.ColumnByFieldName("F3CANPROFINAL").value = .CantidadFinal
                dxDBGrid1.Columns.ColumnByFieldName("VALDESC").Value = .TotalDscto
                
                dxDBGrid1.Columns.ColumnByFieldName("COSTOUNINETO").Value = .PrecioNetoSinImpuesto
                
                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = .BasePorItem
                'dxDBGrid1.Columns.ColumnByFieldName("F3MONINA").value = .ExoneradoPorItem
                dxDBGrid1.Columns.ColumnByFieldName("IGV").Value = .ImpuestoPorItem
                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = .TotalPorItem
                
                dxDBGrid1.Dataset.Post
            End With
                        
            txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
            txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
            txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
    End Select
    
    
'    Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
'        Case "CODFAB", "CODPROD"
'
'        Case "CANTIDAD", "COSTOUNI"
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                If Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDADMAX").value & "") > 0 Then
'                    If Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value & "") > Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDADMAX").value & "") Then
'                        MsgBox "La cantidad no puede exceder al origen seleccionado.", vbInformation + vbOKOnly, App.ProductName
'
'                        dxDBGrid1.Dataset.Cancel
'
'                        dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
'
'                        Exit Sub
'                    End If
'                End If
'
'                If Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value & "") = 0 Then
'                    MsgBox "Debe Ingresar el Costo del Producto " & dxDBGrid1.Columns.ColumnByFieldName("descripcion").value, vbInformation + vbOKOnly, App.ProductName
'
'                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").ColIndex
'
'                    Exit Sub
'                End If
'
'                dxDBGrid1.Dataset.Edit
'
'                sw_detalle = True
'
'                If dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = "*" Then
'                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format((Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.0000")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000"))) * (wwigv / 100), "#0.0000")
'                    dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) * (1 + wwigv / 100), "#0.0000")
'                Else
'                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format(0, "0.00")
'                End If
'
'                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value
'                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.0000")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").value, "0.0000")), "#0.00")
'
'                dxDBGrid1.Dataset.Post
'
'                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
'                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
'                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
'
'                SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'            End If
'        Case "PVUNIT"
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                dxDBGrid1.Dataset.Edit
'
'                sw_detalle = True
'
'                If dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = "*" Then
'                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = (dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value) - ((dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value) / (1 + (wwigv / 100)))
'                    dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value, "0.0000")) / (1 + wwigv / 100), "###,###,##0.0000")
'                Else
'                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format(0, "0.00")
'                End If
'
'                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value
'                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").value, "0.00")), "###,###,##0.00")
'
'                dxDBGrid1.Dataset.Post
'
'                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
'                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
'                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
'            End If
'    End Select
    
'''    Dim rst As New ADODB.Recordset
'''    Dim cad     As String
'''    Dim codprod As String
'''
'''    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
'''        'If sw_nuevo_item = False Then
'''            If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODFAB" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CODPROD" Then
'''                wcodproducto = ""
'''                wcod_alm = txtalmacen.Text
'''                wcodproducto = "" & dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value
'''                codprod = "" & dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value
'''                wmarca = Trim(Mid(cmbmarca.Text, 200))
'''                cad = ""
'''                If Not wmarca = "" Then
'''                    cad = " and a.f5marca='" & wmarca & "'"
'''                End If
'''
'''                sql = "Select A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta,A.F5PRECOS,A.F5VTANET,A.F5VTANETDOL,C.F7SIGMED,D.F2DESMAR,A.F7CODMED,A.AFECTO " & _
'''                      "FROM IF5PLA AS A,EF7MEDIDAS AS C,EF2MARCAS AS D " & _
'''                      "WHERE (A.F5CODFAB='" & wcodproducto & "' or A.F5CODPRO = '" & codprod & "' ) " & _
'''                      "AND A.F7CODMED=C.F7CODMED AND A.F5MARCA=D.F2CODMAR " & cad
'''                If rsif5pla.State = adStateOpen Then rsif5pla.Close
'''                rsif5pla.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
'''                If Not rsif5pla.EOF Then
'''                    wVariosProductos = False
'''                    gcodpro = wcodproducto
'''                    If Len(Trim(wcodproducto)) > 0 Or Len(Trim(codprod)) > 0 Then
'''                        If rsif5pla.RecordCount > 1 Then      'Existe más de 1 producto con el mismo codigo de fabricante
'''                            'frmdetalle.Caption = "Producto " & gcodpro
'''                            'frmdetalle.grddetalle.Rows = 1
'''                            wVariosProductos = True
'''                            wf5codpro = ""
'''                            Do While Not rsif5pla.EOF
'''                                'frmdetalle.grddetalle.Rows = frmdetalle.grddetalle.Rows + 1
'''                                'frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 1) = "" & rsif5pla("f2desmar")
'''                                'frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 2) = "" & rsif5pla("f5codpro")
'''                                'frmdetalle.grddetalle.TextMatrix(frmdetalle.grddetalle.Rows - 1, 3) = "" & rsif5pla("f5nompro")
'''                                rsif5pla.MoveNext
'''                            Loop
'''                            rsif5pla.MoveFirst
'''                            'frmdetalle.Show vbModal
'''                        End If
'''
'''                        If wVariosProductos Then
'''                            If Len(Trim(wf5codpro)) > 0 Then
'''                                cad = "f5codpro='" & wf5codpro & "'"
'''                                rsif5pla.Find cad
'''                            Else
'''                                dxDBGrid1.Dataset.Edit
'''                                dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = 0#
'''                                dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = ""
'''                                dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
'''                                dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = 0#
'''                                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = 0#
'''                                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = 0#
'''                                dxDBGrid1.Dataset.Post
'''                                wcodproducto = ""
'''                                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("codfab").ColIndex
'''                                rsif5pla.Close
'''                                Exit Sub
'''                            End If
'''                        End If
'''
'''                        dxDBGrid1.Dataset.Edit
'''                        dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = "" & rsif5pla.Fields("F5CODPRO")
'''                        dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = "" & rsif5pla.Fields("F5CODFAB")
'''                        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = "" & rsif5pla.Fields("F5NOMPRO")
'''                        dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = "" & rsif5pla.Fields("F7SIGMED")
'''                        dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = "" & rsif5pla.Fields("F2DESMAR")
'''                        wcosto = Val(Costo_Unitario("" & rsif5pla.Fields("F5CODPRO"), abofecha.value, left(cmbmoneda.Text, 1)) & "")
'''                        If left(cmbmoneda.Text, 1) = "S" Then
'''                            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = Format("" & rsif5pla.Fields("F5VTANET"), "#,###,##0.0000")
'''                        Else
'''                            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = Format("" & rsif5pla.Fields("F5VTANETDOL"), "#,###,##0.0000")
'''                        End If
'''                        dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = "" & rsif5pla.Fields("AFECTO")
'''                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
'''                        dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = 0#
'''                        dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = 0#
'''                        dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = 0#
'''
'''                        sw_nuevo_item = True
'''                        dxDBGrid1.Dataset.Post
'''                        sw_nuevo_item = False
'''                        dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
'''                    End If
'''                Else
'''                    MsgBox "El Producto No Existe", vbInformation, "Atención"
'''                    dxDBGrid1.Dataset.Edit
'''                    dxDBGrid1.Columns.ColumnByFieldName("CODPROD").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("CODFAB").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("MARCA").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = 0#
'''                    dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = ""
'''                    dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
'''                    dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = 0#
'''                    dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = 0#
'''                    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = 0#
'''
'''                    dxDBGrid1.Dataset.Post
'''                    wcodproducto = ""
'''                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CODPROD").ColIndex
'''                End If
'''                rsif5pla.Close
'''            End If
''
'''            If dxDBGrid1.Columns.FocusedColumn.FieldName = "COSTOUNI" Then
'''                If Val("" & dxDBGrid1.Columns.ColumnByFieldName("costouni").value) = 0 Then
'''                    MsgBox "Debe Ingresar el Costo del Producto " & dxDBGrid1.Columns.ColumnByFieldName("descripcion").value, vbInformation, "Aviso"
'''                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
'''                    Exit Sub
'''                End If
'''            End If
'''
'''
'''            If dxDBGrid1.Columns.FocusedColumn.FieldName = "COSTOUNI" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CANTIDAD" Then
'''                dxDBGrid1.Dataset.Edit
'''                sw_detalle = True
'''                If dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = "*" Then
'''                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format((Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.0000")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000"))) * (wwigv / 100), "###,###,##0.0000")
'''                    dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) * (1 + wwigv / 100), "###,###,##0.0000")
'''                Else
'''                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format(0, "0.00")
'''                End If
'''                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value
'''                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.0000")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").value, "0.0000")), "###,###,##0.00")
'''                dxDBGrid1.Dataset.Post
'''                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#0.00")
'''                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#0.00")
'''                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#0.00")
'''                SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'''            End If
'''
'''
'''            If dxDBGrid1.Columns.FocusedColumn.FieldName = "PVUNIT" Then
'''                dxDBGrid1.Dataset.Edit
'''                sw_detalle = True
'''                If dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = "*" Then
'''                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = (dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value) - ((dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value) / (1 + (wwigv / 100)))
'''                    dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").value, "0.0000")) / (1 + wwigv / 100), "###,###,##0.0000")
'''                Else
'''                    dxDBGrid1.Columns.ColumnByFieldName("IGV").value = Format(0, "0.00")
'''                End If
'''                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value
'''                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").value, "0.00")), "###,###,##0.00")
'''                dxDBGrid1.Dataset.Post
'''                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#0.00")
'''                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#0.00")
'''                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#0.00")
'''
'''            End If
'''            If dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value <> 0 Then
'''                SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'''            End If
'''
'''
'''        'End If
'''    End If
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
                        
                        If .RegistrarCosto Then
                            If Trim(txtproveedor.Text) = vbNullString And left(cmbconcepto.Text, 7) = "Compras" Then
                                If MsgBox("No se ha seleccionado el Proveedor, ¿Desea continuar?", vbInformation + vbYesNo, App.ProductName) = vbNo Then
                                    
                                    .inicializarEntidades
                                    
                                    Me.MousePointer = vbDefault
                                    
                                    Exit Sub
                                End If
                            End If
                        End If
                    
                        Select Case .CodigoAyudaProducto
                            Case "1"
'                                If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                                    Unload ayuda_productos
'                                End If
'
'                                With ayuda_productos
'                                    .CodigoAuxiliar = IIf(objAyudaOrigen.RegistrarCosto, Trim(txtproveedor.Text), vbNullString)
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
'
'                                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
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
                                    .CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value & "")
                                    .FiltroAdicional = vbNullString
                                    .TipoBienMostrar = "P"
                                    
                                    objAyudaBien.inicializarEntidades
                                    
                                    .Show 1
                                    
                                    If objAyudaBien.Codigo <> vbNullString Then
                                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(CODPROD)", "TMPVALEINGRESO", "CODPROD", objAyudaBien.Codigo, "T", "AND TRIM(F4NUMORD & '') = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "") & "' AND TRIM(COD_SOLICITUD & '') = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value & "") & "'")) > 0 Then
                                            MsgBox "Producto ya seleccionado, verifique.", vbInformation + vbOKOnly, App.ProductName
                                            
                                            Me.MousePointer = vbDefault
                                            
                                            Exit Sub
                                        End If
                                        
                                        
                                        objAyudaBien.obtenerConfigBien
                                        
'                                        Dim strUltimaDescripcion As String
                                        Dim dblUltimoPrecioSinIGv As Double
                                        
'                                        'Obtener la Descripción del Producto en la Ultima Compra (Ordenes de Compra)
'                                        With objAyudaOrden
'                                            .CodProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
'                                            .CodigoProducto = objAyudaBien.Codigo
'
'                                            strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
'                                        End With

                                        'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                                        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                            With objAyudaVale
                                                .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
                                                .CodigoProducto = objAyudaBien.Codigo
                                                
                                                .obtenerUltimoPrecioSinIgvProductoDeProveedor
                                                
                                                Select Case .CodigoMoneda
                                                    Case "S"
                                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(cmbmoneda.Text, 1) = "S", 1, Val(txttc.Text))), "#0.0000")
                                                    Case Else
                                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(cmbmoneda.Text, 1) = "D", 1, Val(txttc.Text))), "#0.0000")
                                                End Select
                                            End With
                                        Else
                                            With objAyudaVale
                                                .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
                                                .CodigoProducto = objAyudaBien.Codigo
                                                
                                                .obtenerUltimoPrecioSinIgvProductoDeProveedor
                                                
                                                Select Case .CodigoMoneda
                                                    Case "S"
                                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(cmbmoneda.Text, 1) = "S", 1, Val(txttc.Text))), "#0.0000")
                                                    Case Else
                                                        dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(cmbmoneda.Text, 1) = "D", 1, Val(txttc.Text))), "#0.0000")
                                                End Select
                                            End With
                                        End If
                                        
                                        With dxDBGrid1
                                            .Dataset.Edit
                                            
                                            .Columns.ColumnByFieldName("CODPROD").Value = objAyudaBien.Codigo
                                            '.Columns.ColumnByFieldName("CODPRODORIGINAL").value = objAyudaBien.Codigo
                                            .Columns.ColumnByFieldName("DESCRIPCION").Value = objAyudaBien.Descripcion
                                            .Columns.ColumnByFieldName("UMEDIDA").Value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                                            .Columns.ColumnByFieldName("CANTIDAD").Value = 0
                                            .Columns.ColumnByFieldName("CANTIDADMAX").Value = 0
                                            .Columns.ColumnByFieldName("COSTOUNI").Value = dblUltimoPrecioSinIGv
                                            .Columns.ColumnByFieldName("AFECTO").Value = IIf(objAyudaBien.Afecto, "*", vbNullString)
                                            
                                            .Dataset.Post
                                        End With
                                    End If
                                End With
                            Case "2"
                                If ModUtilitario.validarFormAbierto("frmUtilStockDisponible") Then
                                    Unload frmUtilStockDisponible
                                End If
                                
                                With frmUtilStockDisponible
                                    .CadenaCorte = Trim(dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value & "") 'InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
                                    
                                    .Show 1
                                    
                                    cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                                    cmbalmacen.Enabled = False
                                End With
                                
                                abrirCnTemporal
                                
                                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILSTOCKDISPONIBLE", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                                    copiarSeleccionStockDisponible
                                    
                                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                                End If
                            Case Else
                                MsgBox "Ayuda de Concepto de Movimiento no configurado, verifique.", vbInformation + vbOKOnly, App.ProductName
                        End Select
                    End With
                    
                    listarGrilla
                    
                    recalcularItems
                    
                    Me.MousePointer = vbDefault
            End Select
        Case vbKeyF3
            Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
                Case "COSTOUNI" 'Sin IGV
                    Dim dblNuevoPrecioSinIgv As Double
                    
                    If dxDBGrid1.Dataset.State = dsEdit Then
                        dxDBGrid1.Dataset.Post
                    End If
                    
                    dblNuevoPrecioSinIgv = Val(InputBox("Ingrese el Precio S/Igv del Producto para el Proveedor:", "Reemplazar Precio S/Igv", Trim(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "")))
                    
                    If dblNuevoPrecioSinIgv > 0 Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPVALEINGRESO SET COSTOUNI = " & dblNuevoPrecioSinIgv & ", PVUNIT = 0 WHERE CODPROD = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "") & "'"
                            
                            dxDBGrid1.Dataset.Close
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrilla
                            
                            recalcularItems
                        End If
                    End If
                Case "PVUNIT" 'Con IGV
                    Dim dblNuevoPrecioConIgv As Double
                    
                    If dxDBGrid1.Dataset.State = dsEdit Then
                        dxDBGrid1.Dataset.Post
                    End If
                    
                    dblNuevoPrecioConIgv = Val(InputBox("Ingrese el Precio C/Igv del Producto para el Proveedor:", "Reemplazar Precio C/Igv", Trim(dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value & "")))
                    
                    If dblNuevoPrecioConIgv > 0 Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPVALEINGRESO SET COSTOUNI = 0, PVUNIT = " & dblNuevoPrecioConIgv & " WHERE CODPROD = '" & Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "") & "'"
                            
                            dxDBGrid1.Dataset.Close
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrilla
                            
                            recalcularItems
                        End If
                    End If
            End Select
    End Select
    
'    Dim cad As String
'    If KeyCode = 113 Then
'        If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODFAB" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "CODPROD" Then
'            wcod_alm = txtalmacen.Text
'            wcodproducto = ""
'            sw_ayuda_prod = True
'            wmarca = Trim(Mid(cmbmarca.Text, 200))
'            cad = ""
'            If Not wmarca = "" Then
'                cad = " and a.f5marca='" & wmarca & "'"
'            End If
'            Me.MousePointer = vbhourglass
'            If dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = "" Then
'            Con_Ayu = 1
'            ayuda_productos.Show 1
'            End If
'            Me.MousePointer = vbdefault
'            If Len(Trim(wcodproducto)) > 0 Then
'                dxDBGrid1.Dataset.Edit
'                dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value = wcodproducto
'                dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Value = wcodfab
'                dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
'                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").Value = wdesproducto
'                dxDBGrid1.Columns.ColumnByFieldName("UMEDIDA").Value = wmedida
'                dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value = IIf(cmbmoneda.Text = "Soles", wprecos, wprecosdol)
'                dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Value = wafecto
'                dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
'                dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value = 0#
'                dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = 0#
'                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = 0#
'                dxDBGrid1.Dataset.Post
'                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
'            End If
'        End If
'    End If
'    If KeyCode = 115 Or KeyCode = 46 Then
'        If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNCODPROD" Then
'            If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
'                sw_nuevo_item = True
'                If dxDBGrid1.Count = 1 Then
'                    dxDBGrid1.Dataset.Delete
'                    AdicionaItem
'                    sw_detalle = False
'                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
'                Else
'                    dxDBGrid1.Dataset.Delete
'                End If
'                txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
'                txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
'                txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
'                sw_nuevo_item = False
'            End If
'        End If
'    End If
End Sub


Private Sub Form_Activate()
'    If sw_nuevo_documento = True Then
'        If sw_activate = True Then
'            sw_activate = False
'            sw_Ord = False
'            cmbmoneda.ListIndex = 0
'        End If
'    Else
'        dxDBGrid1.Dataset.ADODataset.Requery
'    End If
End Sub

Private Sub Form_Load()
'    Dim CadSql          As String
'    Dim pnumvale        As String
'    Dim palmacen        As String
'    Dim Ni As Single
'    Dim Num As Single
'    Dim X As Integer
'
'    Me.MousePointer = vbHourglass
'    sw_Orden = False
'    wopcion = 0
'    If wf1uupp = "*" Then
'        lbluupp.Caption = "UUPP"
'        cmbmarca.Visible = False
'        txtuupp.Visible = True
'        pnluupp.Visible = True
'    Else
'        lbluupp.Caption = "Marca"
'        cmbmarca.Visible = False
'        txtuupp.Visible = False
'        pnluupp.Visible = False
'        CargarMarca
'    End If
'    sw_nuevo_item = False
'    Me.left = 1600
'    Me.top = 1150
'
'    If Rs.State = adStateOpen Then Rs.Close
'
'    Rs.Open "SELECT f2codalm,f2nomalm FROM ef2almacenes order by f2codalm asc", cnn_dbbancos, adOpenStatic, adLockReadOnly
'    X = 0
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        Do While Not Rs.EOF
'            cmbalmacen.AddItem Rs.Fields("f2nomalm") & "" & Space(100) & Rs.Fields("F2CODALM") & ""
'            Rs.MoveNext
'            X = X + 1
'        Loop
'    End If
'    Rs.Close
'
'    If X = 1 Then
'        cmbalmacen.ListIndex = 0
'    End If
'
'    ModMilano.listarCategoriaTipo cmbCategoriaTipo
'
'    Rem SK ADD:
''    If rst.State = adStateOpen Then rst.Close
''
''    SqlCad = vbNullString
''    SqlCad = SqlCad & "SELECT "
''    SqlCad = SqlCad & "F1CODORI, F1COSTO, F1NOMORI "
''    SqlCad = SqlCad & "FROM "
''    SqlCad = SqlCad & "SF1ORIGENES "
''    SqlCad = SqlCad & "WHERE "
''    SqlCad = SqlCad & "F1TIPMOV = 'I' AND "
''    SqlCad = SqlCad & "TRIM(F1CODORIEXTERNO) <> '' "
''    SqlCad = SqlCad & "ORDER BY "
''    SqlCad = SqlCad & "F1COSTO DESC, F1NOMORI"
''
''    'rst.Open "Select F1CODORI,F1COSTO, F1NOMORI FROM SF1ORIGENES WHERE F1TIPMOV='I' ORDER BY F1COSTO DESC,F1NOMORI", cnn_dbbancos, adOpenStatic, adLockReadOnly
''    rst.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
''
''    If Not rst.EOF Then
''        rst.MoveFirst
''
''        Do While Not rst.EOF
''            cad = Space(255)
''            Mid(cad, 1) = "" & rst.Fields("f1nomori")
''            Mid(cad, 200) = "" & rst.Fields("f1codori")
''            Mid(cad, 203) = "" & rst.Fields("f1costo")
''
''            cmbconcepto.AddItem cad
''
''            rst.MoveNext
''        Loop
''    End If
''    rst.Close
''    cmbconcepto.ListIndex = 0
'    sw_activate = True
'
'    'cnombase = wusuario & "VALES" & Format(Time, "hh_mm_ss") & ".MDB"
''    cnombase = "TEMPLUS.mdb"
''    'CREATEDATABASE_N wrutatemp & "\", cnombase
''    "TMPVALEINGRESO" = "TMPVALEINGRESO"
''    If cnDBTemp.State = adStateOpen Then cnDBTemp.Close
''    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
''    cnDBTemp.Open cconex_form
'
'
'
'    'CadSql = "(ITEM TEXT(4),CODPROD TEXT(15),CODFAB TEXT(20),DESCRIPCION TEXT(100),MARCA TEXT(30), " & _
'             "UMEDIDA TEXT(4),CANTIDAD DOUBLE,COSTOUNI DOUBLE,IGV DOUBLE,TOTAL DOUBLE,AFECTO TEXT(1))"
'
'    'CREATETABLE_N "TMPVALEINGRESO", CadSql, cnDBTemp
''    DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
''    DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
'
'    abrirCnTemporal
'
'    DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
'
'    configuraGrilla
'
'    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
'    rsdocumentos.Open "SELECT * FROM DOCUMENTOS ORDER BY F2CODDOC", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    If Not rsdocumentos.EOF Then
'        rsdocumentos.MoveFirst
'        Do While Not rsdocumentos.EOF
'            cmbtipo.AddItem rsdocumentos.Fields("F2DESDOC") & "" & Space(50) & rsdocumentos.Fields("F2CODDOC") & ""
'            rsdocumentos.MoveNext
'        Loop
'    End If
'    rsdocumentos.Close
'
'    sw_detalle = False
'
'    If sw_nuevo_documento = True Then
'        nuevo
'        AdicionaItem
'        sw_cabecera = False
'    Else 'modificacion del vale
'        sw_cabecera = True
'        palmacen = lista_vales.dxDBGrid1.Columns(1).value
'        pnumvale = lista_vales.dxDBGrid1.Columns(2).value
'        BUSCA_VALE palmacen, pnumvale
'        sw_cabecera = False
'        sw_nuevo_documento = False
'        cmbalmacen.Enabled = False
'        txtproveedor.Enabled = False
'    End If
'
'    If sw_nuevo_documento = True Then
'        SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'        SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
'        SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
'        SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
'    Else
'        SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
'        SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
'        SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
'        SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
'
'        Rem SK ADD:
'        If Trim(txtconcepto.Text) = "XCS" Then
'            SSFrame1.Enabled = False
'            dxDBGrid1.Enabled = False
'
'            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
'            SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
'            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
'            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
'            SSActiveToolBars1.Tools("ID_OC").Enabled = False
'        End If
'    End If
'
'    xvale = 1
'    'txtccosto.SetFocus
'    Me.MousePointer = vbDefault
    
    Me.MousePointer = vbHourglass
    
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    'abrirCnTemporal
    
    'cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO"
    
    abrirCnContaTabla
    
    configuraGrilla
    
    listarAlmacenEnCombo
    
    listarTipoDocumentoEnCombo
    
'    ModMilano.listarCategoriaTipo cmbCategoriaTipo
    
    consultarVale
    
    'Activar Control de Apertura de Formulario
    '(Para evitar abrir mas de una vez, el mismo formulario en diferentes Instancias del Programa)
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeIngresoAbierto", "1"
    
    Me.MousePointer = vbDefault
End Sub

Private Sub nuevo()
     SSActiveToolBars1.Tools("ID_OC").Enabled = True
     
    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
    cmbalmacen.Enabled = True
    cmbTipoAuxiliar.ListIndex = -1
    txtproveedor.Enabled = True
    txtnumero.Text = ""
    'txtalmacen.Text = ""
    abofecha.Value = Format(Date, "DD/MM/YYYY")
    txtconcepto.Text = ""
    txtserie.Text = ""
    txtnumdoc.Text = ""
    txtserfac.Text = ""
    txtnumfac.Text = ""
    txtTotigv.Text = ""
    txtTotpv.Text = ""
    txtTotvv.Text = ""
    cmbconcepto.ListIndex = 0
    
    
'    If right(RTrim(cmbconcepto.Text), 1) = "*" Then
'        If UCase(left(cmbconcepto.Text, 6)) = "COMPRA" Then compra = True
'
'        dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = True
'    End If
'
'    txtconcepto.Text = Trim(Mid(cmbconcepto.Text, 200, 3))
'
    dxDBGrid1.Columns.ColumnByFieldName("IGV").Visible = False
    
    cmbalmacen.ListIndex = 0
    txtAlmacen.Text = right(cmbalmacen.Text, 2)
    
    With objAyudaOrigen
        .inicializarEntidades
        
        .Codigo = Trim(Mid(cmbconcepto.Text, 200, 3))
        
        If .obtenerOrigen Then
            txtconcepto.Text = .Codigo
            
            dxDBGrid1.Columns.ColumnByFieldName("PUNIT").Visible = .RegistrarCosto
            
            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Visible = .RegistrarCosto
            'dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").ColIndex = 7
            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Visible = .RegistrarCosto
            'dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ColIndex = 10
            dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Visible = .RegistrarCosto
            'dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").ColIndex = 8
            dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Visible = .RegistrarCosto
            'dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").ColIndex = 9
            SSFrame2.Visible = .RegistrarCosto
        End If
    End With
            
    txttc.Text = "3.377"
    txtproveedor.Text = "": txtnomprov.Text = ""
    txtccosto.Text = "": pnlccosto.Caption = ""
    txtobserva.Text = ""
    txtnomprov.Text = "": pnluupp.Caption = ""
    
    sw_cabecera = False
    sw_detalle = False
            
    'tc
    If IsDate(abofecha.Value) Then
        If rscambios.State = adStateOpen Then rscambios.Close
        If ctipoadm_bd = "M" Then
            sql = "SELECT * FROM CAMBIOS WHERE FECHA='" & abofecha.Value & "'"
        Else
            sql = "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.Value & "')"
        End If
        rscambios.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rscambios.EOF Then
            txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
        Else
            txttc.Text = Format(0, "3.377")
        End If
        rscambios.Close
    End If
    ''''''''
    Rem SK ADD:
    fraOrdenProduccion.Enabled = True
    txtNroOrdenProduccion.Text = vbNullString
        cmbCategoriaTipo.ListIndex = -1
        txtIDOrdenProduccion.Text = vbNullString
        
    lblNumeroValeExterno.Caption = "< ID Externo >"
    chkExportarVale.Value = vbChecked: chkExportarVale.Enabled = True
    
    SSFrame1.Enabled = True
    dxDBGrid1.Enabled = True
    
    txtOcompra.Text = vbNullString
    
    SSActiveToolBars1.Tools("ID_OC").Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Desea salir del registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Cancel = 1
    Else
        Me.dxDBGrid1.Dataset.Close
        
        With lista_vales
            .listarVale
        End With
        
        ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeIngresoAbierto", "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    If cnDBTemp.State = 1 Then cnDBTemp.Close
'
'    sw_nuevo_item = True
'    dxDBGrid1.Dataset.Close
'    ELIMINA_BD_N wrutatemp, cnombase
'
'    If sw_ayuda_prod = True Then
'        Unload ayuda_productos
'    End If
'
'    If sw_importa_valedeingreso = False Then
'        lista_vales.dxDBGrid1.Dataset.Active = False
'        lista_vales.dxDBGrid1.Dataset.Refresh
'        lista_vales.dxDBGrid1.Dataset.Active = True
'    End If
    
End Sub

Private Sub pnlproveedor_Click()

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error Resume Next
    
'    Dim rst As ADODB.Recordset
    
    Select Case Tool.ID
        Case "ID_Nuevo"
'            If dxDBGrid1.Dataset.RecordCount > 1 Or Trim(txtIDOrdenProduccion.Text) <> vbNullString Or Trim(txtproveedor.Text) <> vbNullString Then
'                If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'                    Exit Sub
'                End If
'            End If
'
'            Me.MousePointer = vbHourglass
'            sw_nuevo_documento = False
'            sw_detalle = False
'            nuevo
'            AdicionaItem
'            AdicionaItem
'            sw_nuevo_documento = True
'            Me.MousePointer = vbDefault
'            cmbalmacen.Enabled = True
'            'cmbalmacen.SetFocus
'            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
'            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
        
            If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            strCodAlmacen = vbNullString
            strNumeroVale = vbNullString
            
            consultarVale
            
            Me.MousePointer = vbDefault
        Case "ID_Grabar"
'            'Me.MousePointer = vbHourglass
'
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'            'Validacion del Punto (PC) que origina el Vale
'            ModMilano.abrirCnDBMilano
'
'            If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'                MsgBox "Su computador no esta registrado. Consulte con su" & vbNewLine & vbNewLine & _
'                        "administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
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
'            'If sw_cabecera = True Or sw_detalle = True Then
'                '-----------------------------------------------
'                If cmbalmacen.ListIndex < 0 Then
'                    MsgBox "Debe seleccionar el almacén.", vbDefaultButton1 + vbInformation, "Atención"
'                    cmbalmacen.SetFocus
'                    Exit Sub
'                End If
'                '-----------------------------------------------
'                If cmbconcepto.ListIndex < 0 Then
'                    MsgBox "Debe seleccionar el concepto de ingreso.", vbDefaultButton1 + vbInformation, "Atención"
'                    cmbconcepto.SetFocus
'                    Exit Sub
'                End If
'                '-----------------------------------------------
'                If Val(txttc.Text & "") = 0 Then
'                    MsgBox "Ingrese el tipo de cambio ", vbDefaultButton1 + vbInformation, "Atención"
'                    If txttc.Enabled = True Then
'                        txttc.SetFocus
'                     Exit Sub
'                     End If
''                    txttc.Text = "3.45"
'                End If
'                '-----------------------------------------------
'''                Set rst = New ADODB.Recordset
'''                'Valida Existencias de Productos que ingresan a Almacén
'''                cnDBTemp.Execute "update " & "TMPVALEINGRESO" & " set cantidad = cantidad"
'''                If rst.State = adStateOpen Then rst.Close
'''                SQL = "SELECT sum(cantidad) as cant FROM " & "TMPVALEINGRESO"
'''                rst.Open SQL, cnDBTemp, adOpenStatic, adLockOptimistic
'''                If Not rst.EOF Then
'''                    xcant = Val("" & rst("cant"))
'''                    If xcant = 0 Then
'''                        Me.MousePointer = vbDefault
'''                        MsgBox "Debe Ingresar Cantidad del Producto(s)", vbInformation, "Sistema de Logística"
'''                        rst.Close
'''                        Exit Sub
'''                    End If
'''                End If
'''                rst.Close
'                '-----------------------------------------------
'                If Trim(txtconcepto.Text) = "XOP" And Trim(txtIDOrdenProduccion.Text) = vbNullString Then
'                    MsgBox "ID de Orden de Produccion incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                    cmbCategoriaTipo.SetFocus
'
'                    Exit Sub
'                End If
'
'                If Trim(txtconcepto.Text) <> "XOP" And Trim(txtIDOrdenProduccion.Text) <> vbNullString Then
'                    MsgBox "Concepto incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                    cmbconcepto.SetFocus
'
'                    Exit Sub
'                End If
'
'                If Trim(txtconcepto.Text) <> "XC0" Then
'                    If Trim(txtserie.Text) = vbNullString And Trim(txtnumdoc.Text) = vbNullString And _
'                       Trim(txtserfac.Text) = vbNullString And Trim(txtnumfac.Text) = vbNullString Then
'
'                        If MsgBox("¿Desea guardar la Compra sin numeros de documento de referencia?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                If MsgBox("¿Desea grabar el Vale de Ingreso?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                    Me.MousePointer = vbHourglass
'
'                    grabar
'
'                    If sw_GRABA_REGISTRO_logistica Then
'                        If Trim(txtconcepto.Text) = "XOP" And Trim(txtIDOrdenProduccion.Text) <> vbNullString Then
'                            verificarProductoDevueltoOP
'                        End If
'
'                        If ModMilano.exportarValeAserverSQL(Trim(txtalmacen.Text), Trim(txtnumero.Text), lblNumeroValeExterno) Then
'                            MsgBox "Vale Exportado.", vbInformation + vbOKOnly, App.ProductName
'                        End If
'                    End If
'
'                    BUSCA_VALE Trim(txtalmacen.Text), Trim(txtnumero.Text)
'
'                    sw_detalle = False
'                    sw_cabecera = False
'                End If
'            'End If
'
'            Me.MousePointer = vbDefault
'
'            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
'            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
'            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
    
            Me.MousePointer = vbHourglass
            
            validarCajas
            
            Me.MousePointer = vbDefault
        Case "ID_Eliminar":
'
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'            'Restricción de Anulación de Vale
'            If Month(CDate(abofecha.value)) < Month(Date) Then
'                MsgBox "Imposible eliminar Vale. Fuera del Periodo Actual." & vbNewLine & vbNewLine & _
'                        "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'                Exit Sub
'            End If
'            Rem SK ADD:-------------------------------------------------------------------------------------------------------------------------
'
'            If Len(Trim("" & txtnumero.Text)) = 0 Then
'                MsgBox "El Vale de Ingreso no ha sido grabado. Verifique", vbCritical, "Atención"
'
'                Exit Sub
'            End If
'
'            If MsgBox("¿Está seguro(a) de eliminar el Vale de Ingreso ?", vbYesNo + vbInformation, "Atención") = vbYes Then
'                Rem SK ADD:
'                If Val(lblNumeroValeExterno.Caption) > 0 Then
'                    If Not ModMilano.anularValeExterno("I", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtalmacen.Text), "T")) Then
'                        Me.MousePointer = vbDefault
'
'                        Exit Sub
'                    End If
'                End If
'
'                elimina txtnumero.Text, txtalmacen.Text
'            Else
'                Exit Sub
'            End If
'
'            sw_nuevo_documento = True
'
'            nuevo
'
'            dxDBGrid1.Dataset.Close
'            DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
'            AdicionaItem
            
            Me.MousePointer = vbHourglass
            
            eliminarVale

            Me.MousePointer = vbDefault
        Case "ID_ImprimirA4":
            With objAyudaVale
                .TipoVale = "I"
                .CodigoAlmacen = Trim(txtAlmacen.Text)
                .NumeroVale = Trim(txtnumero.Text)
                
                If Not .verificarExistencia Then
                    MsgBox "Vale no registrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
            End With
            
            'IMPRIMIR_VALES (1)
            With rptValeIngreso
                .TipoVale = "I"
                .CodAlmacen = Trim(txtAlmacen.Text)
                .NumeroVale = Trim(txtnumero.Text)
                
                'ModMilano.abrirCnDBMilano
                
'                .fldCategoria.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CT.NOMBRE", "ORDENPRODUCCION AS OP LEFT JOIN CATEGORIATIPO AS CT ON CT.IDCATEGORIATIPO = OP.IDCATEGORIATIPO", "OP.IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
'                .fldOP.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Val(txtIDOrdenProduccion.Text), "N")
                
                .Show 1
            End With
        Case "ID_ImprimirA5":
            IMPRIMIR_VALES (2)
        'Case "ID_CargarData":
        '    frmExcel.Show 1
        Case "ID_Lista":
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                dxDBGrid1.Dataset.Post
'                sw_detalle = True
'            End If
'            If sw_detalle = True Or sw_cabecera = True Then
'                If (dxDBGrid1.Count >= 1 And dxDBGrid1.Columns.ITEM(2).value <> "" And sw_nuevo_documento = True) Then
'                    If MsgBox("El Vale de Ingreso no ha sido grabado ... Desea Grabar ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'                        grabar
'                        sw_detalle = False
'                    End If
'                End If
'            End If
'            Unload Me
            
            Unload Me
        Case "ID_OC": 'Opcion para Descargar Mercaderia de O/C
'            If Trim(txtconcepto.Text) <> "XC0" Then
'                MsgBox "Opción disponible solo para Compras.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
'
'            If Trim(txtproveedor.Text) = vbNullString Then
'                MsgBox "Seleccione el Proveedor.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
'
'            wrucprov = ObtenerCampo("EF2PROVEEDORES", "f2newruc", "F2CODPROV", Trim(txtproveedor.Text), "T", cnn_dbbancos)
'            whelpoc = "S"
'            wtipoc = ""
'            wcodcosto = vbNullString
'            importar_ocompra_logistica.Show 1
'            whelpoc = "N"
'
'            If wcodcosto <> vbNullString Then
'                SELECCIONA_OCOMPRAS
'
'                'cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XC0", "DER", 3)
'            End If
'
'            Unload importar_ocompra_logistica
            
            If Trim(txtproveedor.Text) = vbNullString Then
                MsgBox "Seleccione el Proveedor.", vbInformation + vbOKOnly, App.ProductName
                
                txtproveedor.SetFocus
                
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            If ModUtilitario.validarFormAbierto("frmUtilDevolucionOC") Then
                Unload frmUtilDevolucionOC
            End If
            
            With frmUtilDevolucionOC
                .TipoVale = "I"
                .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(txtproveedor.Text), "T")
                
                .Show 1
            End With
            
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(*)", "TMPCPRECEPCIONOC" & UCase(wusuario), "PROCESAR", "1", "N") & "") <> 0 Then
                    'cmbAlmacen.ListIndex = ModUtilitario.seleccionarItem(cmbAlmacen, right(frmUtilDevolucionOC.cmbAlmacen.Text, 2), "DER", 2)
                    cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XC0", "DER", 3)
                    
                    cmbConcepto_Click
                    
                    copiarSeleccionRecepcionOCSql
                    
                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                End If
            Else
                abrirCnTemporal
                
                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILDEVOLUCIONOC", "PROCESAR", "TRUE", "N") & "") <> 0 Then
'                    cmbAlmacen.ListIndex = ModUtilitario.seleccionarItem(cmbAlmacen, right(frmUtilDevolucionOC.cmbAlmacen.Text, 2), "DER", 2)
'                    cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XC0", "DER", 3)
                    
                    cmbConcepto_Click
                    
                    copiarSeleccionRecepcionOC
                    
                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                End If
            End If
            
            listarGrilla
            
            recalcularItems
            
            txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
            txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
            txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
            
            Me.MousePointer = vbDefault
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
'            If sw_detalle = True Or sw_cabecera = True Then
'                If (dxDBGrid1.Count >= 1 And dxDBGrid1.Columns.ITEM(1).value <> "" And sw_nuevo_documento = True) Then
'                    If MsgBox("El Vale de Ingreso no ha sido grabado ... ¿Desea Grabar?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'                        grabar
'                        sw_detalle = False
'                    End If
'                End If
'            End If
'            Unload Me
'            Unload lista_vales
            
            Unload Me
    End Select

End Sub

Private Sub grabar()
    On Error GoTo HndError
    
    Dim calma_obra      As String
    Dim costo As Double
    Dim Cant As Integer
    
    'If rsccosto.State = adStateOpen Then rsccosto.Close
    'rsccosto.Open "SELECT F3ALMACEN FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    'If Not rsccosto.EOF Then
    '    calma_obra = Trim("" & rsccosto.Fields("F3ALMACEN"))
    'End If
    'rsccosto.Close
    If ctipoadm_bd <> "M" Then
        'cnn_dbbancos.BeginTrans
    End If
    
    GRABA_ING_ALMACEN_CENTRAL
    'arreglo = Costo_Unitario2(wcodproducto, wmoneda_productos)
    'costo = arreglo(0)
    'Cant = Conversion.CInt(arreglo(1))
    

    'If Trim(wtiposalida) = "*" Then  '----- LA EMPRESA ES CONSTRUCTORA (AIC)
    '    If Len(Trim(calma_obra)) > 0 Then
    '        GRABA_SALIDA_X_TRANSF
    '        GRABA_ING_ALMACEN_OBRA calma_obra
    '    Else
    '        GRABA_VALE_SALIDA
    '    End If
    'End If
    
    sw_nuevo_documento = False
    
    If ctipoadm_bd <> "M" Then
        'cnn_dbbancos.CommitTrans
    End If
    
    MsgBox "Se ha Actualizado el Vale de Ingreso " & txtnumero.Text, vbInformation, "Sistema de Logística"
    
    Exit Sub
HndError:
    MsgBox "Ha Ocurrido el siguiente error:" & vbNewLine & vbNewLine & _
            Err.Description & "." & vbNewLine & _
            "La Operación de Actualización no se Realizó. Consulte al Proveedor.", vbCritical, "Sistema de Logistica"
    
    Me.MousePointer = vbDefault
    
    'cnn_dbbancos.RollbackTrans
    
    Err.Clear
End Sub



Private Sub txtalmacen_Change()
    
    If Trim(txtAlmacen.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtAlmacen.SelStart = 0: txtAlmacen.SelLength = Len(txtAlmacen.Text)

End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtAlmacen.Text)) > 0 Then
            abofecha.SetFocus
        Else
            txtAlmacen.SetFocus
        End If
    End If
    
End Sub

Private Sub txtccosto_Change()

    If Trim(txtccosto.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtccosto_DblClick()

    txtccosto_KeyDown 113, 0
    
End Sub

Private Sub txtccosto_GotFocus()
    
    txtccosto.SelStart = 0: txtccosto.SelLength = Len(txtccosto.Text)
    
End Sub

Private Sub txtccosto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodcosto = "": wdescosto = ""
        Ayuda_Centros.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            txtccosto.Text = wcodcosto
            pnlccosto.Caption = wdescosto
            txtccosto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtccosto_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
         ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub txtccosto_LostFocus()
    
    If sw_ayuda = False Then
        'If Val(txttc.Text & "") > 0# Then
            If Len(Trim(txtccosto.Text)) > 0 Then
                If rsccosto.State = adStateOpen Then rsccosto.Close
                rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsccosto.EOF Then
                    pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
                Else
                    MsgBox "Código de centro de costo no existe. Verifique.", vbCritical, "Atención"
                    txtccosto.SetFocus
                End If
                rsccosto.Close
            Else
                'MsgBox "Falta ingresar el centro de costo.", vbCritical, "Atención"
                'txtccosto.SetFocus
            End If
        'End If
    End If

End Sub

Private Sub txtconcepto_Change()
    
    If Trim(txtconcepto.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtconcepto_GotFocus()

    txtconcepto.SelStart = 0: txtconcepto.SelLength = Len(txtconcepto.Text)
    
End Sub

Private Sub txtconcepto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       ' txttc.SetFocus
    End If

End Sub

Private Sub txtDscto_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtDscto
End Sub

Private Sub txtDscto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If MsgBox("¿Desea aplicar el Descuento Total ingresado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                calcularDescuentoTotal
            Else
                txtDscto.Text = "0.00"
            End If
    End Select
End Sub

Private Sub txtNroOrdenProduccion_KeyDown(KeyCode As Integer, Shift As Integer)
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
            
            'ModMilano.abrirCnDBMilano
            
            'txtIDOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDORDENPRODUCCION", "ORDENPRODUCCION", "IDCATEGORIATIPO", right(Trim(cmbCategoriaTipo.Text), 10), "T", "AND OP = '" & Trim(txtNroOrdenProduccion.Text) & "' AND ANULADO = 0")
            txtIDOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDORDENPRODUCCION", "ORDENPRODUCCION", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T", "AND OP = '" & Trim(txtNroOrdenProduccion.Text) & "' AND ANULADO = 0")
            
            If Trim(txtIDOrdenProduccion.Text) = vbNullString Then
                MsgBox "O.P. no existe o esta anulada.", vbInformation + vbOKOnly, App.ProductName
            Else
                'Me.MousePointer = vbHourglass
                
                If MsgBox("¿Desea verificar las descargas de la O/P?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                    If ModUtilitario.validarFormAbierto("frmUtilSalidaOP") Then
                        Unload frmUtilSalidaOP
                    End If
                    
                    With frmUtilSalidaOP
                        .TipoVale = "I"
                        .IdOrdenProduccion = Trim(txtIDOrdenProduccion.Text)
                        '.NroOP = Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", right(Trim(cmbCategoriaTipo.Text), 10), "T")) & "-" & Trim(txtNroOrdenProduccion.Text)
                        .NroOP = Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T")) & "-" & Trim(txtNroOrdenProduccion.Text)
                        
                        .Show 1
                        
                        cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, right(.cmbalmacen.Text, 2), "DER", 2)
                        cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, "XOP", "DER", 3)
                    End With
                    
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        If Val(ModUtilitario.ObtenerCampoV2(cnBdCPlus, "COUNT(*)", "TMPCPDEVOLUCIONOP" & UCase(wusuario), "PROCESAR", "1", "N") & "") <> 0 Then
                            copiarSeleccionDevolucionOPSql
                        End If
                    Else
                        abrirCnTemporal
                        
                        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILDEVOLUCIONOP", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                            copiarSeleccionDevolucionOP
                        End If
                    End If
                    
                    listarGrilla
                End If
                
                'Me.MousePointer = vbDefault
            End If
            
            ModUtilitario.pulsarTecla vbKeyTab
            
            Screen.MousePointer = vbDefault
    End Select
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

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub TxtNumDoc_LostFocus()
    With objAyudaEmpresa
        .CodigoEmpresa = wF1Dir
        
        .obtenerConfigEmpresa
    End With
    
    txtnumdoc.Text = Format(txtnumdoc.Text, objAyudaEmpresa.FormatoNumDocCompra)
End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dxDBGrid1.SetFocus
'    sw_nuevo_item = True
    dxDBGrid1.Columns.FocusedIndex = 1
End If
End Sub

Private Sub txtnumfac_Change()

    If Trim(txtnumfac.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub Txtnumfac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
End Sub

'Private Sub Txtnumfac_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        txtobserva.SetFocus
'    End If
'
'End Sub

Private Sub txtnumfac_LostFocus()
    With objAyudaEmpresa
        .CodigoEmpresa = wF1Dir
        
        .obtenerConfigEmpresa
    End With
    
    txtnumfac.Text = Format(txtnumfac.Text, objAyudaEmpresa.FormatoNumDocCompra)
    
    If strNumeroVale = vbNullString Then
        If Trim(txtproveedor.Text) <> vbNullString And _
            cmbtipo.ListIndex <> -1 And _
            Trim(txtnumfac.Text) <> vbNullString Then
            
            With objAyudaVale
                .CodigoProveedor = Trim(txtproveedor.Text)
                
                .SerieGuia = Trim(txtserie.Text)
                .NumeroGuia = Trim(txtnumdoc.Text)
                
                .CodTipoComprobante = right(cmbtipo.Text, 2)
                .SerieDocumento = Trim(txtserfac.Text)
                .NumeroDocumento = Trim(txtnumfac.Text)
                
                If .verificarExistenciaPorNumRef Then
                    MsgBox "Ya existe un Vale de Ingreso registrado con los Numeros de Documentos de Referencia, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    ModUtilitario.seleccionarTextoCaja txtnumfac
                End If
            End With
        End If
    End If
End Sub

'Private Sub txtobserva_Change()
'
'    If Trim(txtobserva.Text) <> "" And sw_cabecera = False Then
'        sw_cabecera = True
'    End If
'
'End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 1
    End If
    
End Sub

'Private Sub txtproveedor_Change()
'    If Trim(txtproveedor.Text) <> "" Or sw_cabecera = False Then
'        sw_cabecera = True
'        SSActiveToolBars1.Tools("ID_OC").Enabled = True
'    End If
'    If txtproveedor.Text = "99999999999" Then
'        Txtnomprov.Enabled = True
'        Txtnomprov.BackColor = vbWhite
'    Else
'        Txtnomprov.Enabled = False
'        Txtnomprov.BackColor = &H8000000F
'    End If
'End Sub
'
'Private Sub txtproveedor_DblClick()
'
'    txtproveedor_KeyDown 113, 0
'
'End Sub
'
'Private Sub txtproveedor_GotFocus()
'    txtproveedor.SelStart = 0: txtproveedor.SelLength = Len(txtproveedor.Text)
'End Sub
'
'Private Sub txtproveedor_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        sw_ayuda = True
'        wcodprov = "": wrucprov = "": wnomprov = ""
'        'hlp_proveedores.Show 1
'        Ayuda_Proveedores.Show 1
'        sw_ayuda = False
'        If Len(Trim(wcodcliprov)) > 0 Then
'            txtproveedor.Text = wcodcliprov
'            Txtnomprov.Text = wnomcliprov
'            txtproveedor_KeyPress 13
'            CmbMoneda.Text = IIf(wmoneda_productos = "S", "Soles", "Dolares")
'
'            If Trim(txtconcepto.Text) = "XC0" Then
'                SSActiveToolBars1_ToolClick SSActiveToolBars1.Tools("ID_OC")
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub txtproveedor_KeyPress(KeyAscii As Integer)
'    On Error Resume Next
'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'    End If
'
'End Sub
'
'Private Sub txtproveedor_LostFocus()
'
'    If sw_ayuda = False Then
'        'If Val(txttc.Text & "") > 0# Then
'            If Len(Trim(txtproveedor.Text)) > 0 And txtproveedor.Text <> "99999999999" Then
'                If VALIDA_PROVEEDOR(txtproveedor.Text) = True Then
'                    Txtnomprov.Text = wnomprov
'                Else
'                    MsgBox "El proveedor no existe. Verifique.", vbCritical, "Atención"
'                    txtproveedor.SetFocus
'                End If
'            End If
'        'End If
'    End If
'
'End Sub

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
                    SSActiveToolBars1.Tools("ID_OC").Enabled = False
                    With ayuda_usuarios
                        .Show 1
                    End With
                Case 1 'Proveedores
                    SSActiveToolBars1.Tools("ID_OC").Enabled = True
                    
                    With Ayuda_Proveedores
                        .Show 1
                    End With
                Case Else
                    MsgBox "Seleccione el Tipo de Persona.", vbInformation + vbOKOnly, App.ProductName
                    
                    SSActiveToolBars1.Tools("ID_OC").Enabled = False
                    
                    cmbTipoAuxiliar.SetFocus
                    
                    Me.MousePointer = vbDefault
                    
                    Exit Sub
            End Select
            
            If wcodcliprov <> vbNullString Then
                txtproveedor.Text = wcodcliprov
                txtnomprov.Text = wnomcliprov
                
                Select Case cmbTipoAuxiliar.ListIndex
                    Case 1 'Proveedores
                        With objAyudaProveedor
                            .inicializarEntidades
                            
                            .Codigo = Trim(txtproveedor.Text)
                            
                            .obtenerConfigProveedor
                            
                            cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, .CodigoMoneda, "IZQ", 1)
                            
                            .inicializarEntidades
                        End With
                End Select
                
                'SSActiveToolBars1_ToolClick SSActiveToolBars1.Tools("ID_OC")
            End If
            
            Me.MousePointer = vbDefault
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
                
                If Trim(txtconcepto.Text) = "XC0" And Not ModUtilitario.validarFormAbierto("importar_ocompra_logistica") Then
                    'SSActiveToolBars1_ToolClick SSActiveToolBars1.Tools("ID_OC")
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

Private Sub Txtserfac_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub txtserfac_LostFocus()
    ' Convertir a mayúsculas
    txtserfac.Text = UCase(txtserfac.Text)
    
    ' Asegurarse de que el texto tenga 4 dígitos
    If Len(txtserfac.Text) <> 4 Then
        MsgBox "El campo debe contener exactamente 4 dígitos.", vbExclamation, "Advertencia"
        txtserfac.SetFocus  ' Regresar al campo para que el usuario lo corrija
    Else
        ' Formatear a 4 dígitos
        txtserfac.Text = Format(txtserfac.Text, "0000")
    End If
End Sub

Private Sub txtserie_Change()

    If Trim(txtserie.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub txtserie_LostFocus()
    
    txtserie.Text = Format(txtserie.Text, "0000")
    
End Sub

Private Sub txttc_Change()
    txttc.BackColor = vbWhite
    If Trim(txttc.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txttc_GotFocus()

    txttc.SelStart = 0: txttc.SelLength = Len(txttc.Text)
    
End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        If txtproveedor.Visible = True Then
'            txtproveedor.SetFocus
'        Else
'            txtccosto.SetFocus
'        End If
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txttc_LostFocus()
    On Error Resume Next
    
    txttc.Text = Format(txttc.Text, "0.000")
    If Len(Trim(txtconcepto.Text)) > 0 Then
        If Val(Format(txttc.Text, "0.000")) > 0 Then
        
        Else
'            MsgBox "El tipo de cambio no puede ser cero. Verifique.", vbInformation, "Sistema de Logistica"
            txttc.SetFocus
       End If
    End If

End Sub

Private Function GENERA_NUMVALE(palmacen As String, pmes As String, ptipo As String)
    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "IF4VALES.F2CODALM, "
    strSQL = strSQL & "IF4VALES.F4NUMVAL "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "IF4VALES "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "MONTH(IF4VALES.F4FECVAL) = " & pmes & " AND "
    strSQL = strSQL & "LEFT(IF4VALES.F4NUMVAL, 1) = '" & ptipo & "' AND "
    strSQL = strSQL & "F2CODALM = '" & palmacen & "' "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "IF4VALES.F4NUMVAL DESC"
    
    rsalmacen.Open strSQL, cnn_dbbancos, 3, 1
    
    If Not rsalmacen.EOF Then
        rsalmacen.MoveFirst
        
        cnumvale = Mid(rsalmacen.Fields(1).Value, 1, 4) & Format(Val(Mid(rsalmacen.Fields(1), 5, 4)) + 1, "0000")
        
        GENERA_NUMVALE = cnumvale
    Else
        If pmes = "01" Then
            cnumvale = ptipo & "-010001"
            
            GENERA_NUMVALE = cnumvale
        Else
            pmes = Format(Val(pmes), "00")
            
            If rsalmacen.State = adStateOpen Then rsalmacen.Close
            
            strSQL = vbNullString
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "IF4VALES.F2CODALM, "
            strSQL = strSQL & "IF4VALES.F4NUMVAL "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "IF4VALES "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "MID(IF4VALES.F4NUMVAL, 3, 2) = '" & pmes & "' AND "
            strSQL = strSQL & "LEFT(IF4VALES.F4NUMVAL, 1) = '" & ptipo & "' AND "
            strSQL = strSQL & "F2CODALM = '" & palmacen & "' "
            strSQL = strSQL & "ORDER BY "
            strSQL = strSQL & "IF4VALES.F4NUMVAL DESC"
            
            rsalmacen.Open strSQL, cnn_dbbancos, 3, 1
            
            If Not rsalmacen.EOF Then
                rsalmacen.MoveFirst
                cnumvale = Mid(rsalmacen.Fields(1).Value, 1, 4) & Format(Val(Mid(rsalmacen.Fields(1), 5, 4)) + 1, "0000")
            Else
                cnumvale = ptipo & "-" & Format(Val(pmes), "00") & Format(Val(1), "0000")
            End If
            
            rsalmacen.Close
            
            GENERA_NUMVALE = cnumvale
        End If
    End If
End Function

Private Sub SELECCIONA_OCOMPRAS()
    Dim x       As Integer

    If importar_ocompra_logistica.dxDBGrid1.Count > 0 Then
        s$ = importar_ocompra_logistica.dxDBGrid1.Columns(5).Value
        
        BUSCA_OCOMPRA wcodcosto   'Trim(importar_ocompra_logistica.dxDBGrid1.Columns(0).Value)
        
        sw_Ord = True
    Else
        sw_Ord = False
    End If

End Sub

Private Sub BUSCA_OCOMPRA(pocompra As String)
    Dim I As Integer
    
    If rsif4orden.State = adStateOpen Then rsif4orden.Close
    
    rsif4orden.Open "SELECT * FROM IF4ORDEN WHERE F4NUMORD = '" & pocompra & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rsif4orden.EOF Then
'--------------------- C. COSTO
        txtccosto.Text = Trim("" & rsif4orden.Fields("F4CENTRO"))
        If rsccosto.State = adStateOpen Then rsccosto.Close
        rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsccosto.EOF Then
            pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
        End If
        rsccosto.Close
        '-----------------------------------------------------
        If Trim("" & rsif4orden.Fields("F4TIPMON")) = "S" Then
            cmbmoneda.ListIndex = 0
        ElseIf Trim("" & rsif4orden.Fields("F4TIPMON")) = "D" Then
            cmbmoneda.ListIndex = 1
        End If
        
        'Txtnomprov.Text = rsif4orden.Fields("F4REFERE") & ""
'        txtuupp.Text = rsif4orden.Fields("F4UUPP") & ""
'        If VALIDA_UUPP(txtuupp.Text) = True Then
'            pnluupp.Caption = wdeslocalidad
'        End If
'
        If rsif3orden.State = adStateOpen Then rsif3orden.Close
        
        Rem SK ADD:
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, DET.COD_SOLICITUD, DET.F3CODPRO, DET.F3CODFAB, "
        SqlCad = SqlCad & "DET.F5NOMPRO, MED.F7SIGMED, DET.F5AFECTO, "
        SqlCad = SqlCad & "(DET.F3PRENETO / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS PRECOS, "
        SqlCad = SqlCad & "DET.F3IGV, "
        SqlCad = SqlCad & "DET.F3TOTAL, "
        
        '(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)))
        
        'SqlCad = SqlCad & "(DET.F3CANPRO * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD, "
        SqlCad = SqlCad & "((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD, "
        
        'SqlCad = SqlCad & "((DET.F3CANPRO * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(MOVPROD.CANTIDAD & '')) AS SALDO "
        'SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(MOVPROD.CANTIDAD & '')) AS SALDO "
        SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
        
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        'SqlCad = SqlCad & "CAB.NUMORDEN, "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDAD "
        'SqlCad = SqlCad & "DET.F3CANPRO AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
        'SqlCad = SqlCad & "GROUP BY "
        'SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS MOVPROD "
        SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') <> '' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL"
        SqlCad = SqlCad & ") AS INGRESOS "
        'SqlCad = SqlCad & "ON MOVPROD.NUMORDEN = DET.F4NUMORD AND MOVPROD.F5CODPROORIGINAL = DET.F3CODPRO "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & pocompra & "' AND "
        'SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(MOVPROD.CANTIDAD & '')) > 0 "
        SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "DET.F3CODPRO"
        
        'rsif3orden.Open "SELECT * FROM IF3ORDEN WHERE F4NUMORD = '" & pocompra & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        rsif3orden.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        
        If Not rsif3orden.EOF Then
            sw_Orden = True
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
            dxDBGrid1.Dataset.Open
            dxDBGrid1.Dataset.Active = True
        
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
                        
            dxDBGrid1.OptionEnabled = False
            dxDBGrid1.Dataset.DisableControls
            With dxDBGrid1.Dataset
                I = IIf(dxDBGrid1.Count > 1, dxDBGrid1.Count + 1, 1)
                sw_nuevo_item = True
                rsif3orden.MoveFirst
                Do While Not rsif3orden.EOF
                    If sw_detalle = False Then
                        .Edit
                    Else
                        .Append
                    End If
                    
                    .FieldValues("ITEM") = I
                    
                    .FieldValues("F4NUMORD") = Trim("" & rsif3orden.Fields("F4NUMORD"))
                    .FieldValues("COD_SOLICITUD") = Trim("" & rsif3orden.Fields("COD_SOLICITUD"))
                    
                    .FieldValues("CODPROD") = Trim("" & rsif3orden.Fields("F3CODPRO"))
                    .FieldValues("CODPRODORIGINAL") = Trim("" & rsif3orden.Fields("F3CODPRO"))
                    .FieldValues("CODFAB") = Trim("" & rsif3orden.Fields("F3CODFAB"))
                    
                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    
                    rsif5pla.Open "SELECT F5NOMPRO,F7CODMED,F5MARCA,F5CODFAB FROM IF5PLA WHERE F5CODPRO = '" & Trim("" & rsif3orden.Fields("F3CODPRO")) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    
                    If Not rsif5pla.EOF Then
                        .FieldValues("DESCRIPCION") = "" & rsif5pla.Fields("F5NOMPRO")
                        .FieldValues("UMEDIDA") = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsif5pla.Fields("F7CODMED"), "T", cnn_dbbancos)  '"" & rsif5pla.Fields("F7CODMED")
                        .FieldValues("marca") = "" & rsif5pla.Fields("f5marca")
                        .FieldValues("CODFAB") = "" & rsif5pla.Fields("F5CODFAB")
                    End If
                    
                    rsif5pla.Close
                    
                    .FieldValues("CANTIDAD") = Format(Val("" & rsif3orden.Fields("SALDO")), "###,##0.0000") 'Format(Val("" & rsif3orden.Fields("F3CANPRO")), "###,##0.0000")
                    .FieldValues("COSTOUNI") = Format(Val("" & rsif3orden.Fields("PRECOS")), "###,##0.0000") 'Format(Val("" & rsif3orden.Fields("F3PRECOS")), "###,##0.0000")
                    .FieldValues("AFECTO") = rsif3orden.Fields("F5AFECTO")
                    .FieldValues("IGV") = Format(Val("" & rsif3orden.Fields("F3IGV")), "###,##0.00")
                    .FieldValues("TOTAL") = Format(Val("" & rsif3orden.Fields("F3TOTAL")), "###,##0.00")
                    rsif3orden.MoveNext
'Calcula pvunitario
                    dxDBGrid1.Dataset.Edit
                    sw_detalle = True
                    If dxDBGrid1.Columns.ColumnByFieldName("AFECTO").Value = "*" Then
                        dxDBGrid1.Columns.ColumnByFieldName("IGV").Value = Format((Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000"))) * (wwigv / 100), "###,###,##0.00")
                        dxDBGrid1.Columns.ColumnByFieldName("PVUNIT").Value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000")) * (1 + wwigv / 100), "###,###,##0.0000")
                    Else
                        dxDBGrid1.Columns.ColumnByFieldName("IGV").Value = Format(0, "0.00")
                    End If
                    dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value
                    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value, "0.00")), "###,###,##0.00")
                    dxDBGrid1.Dataset.Post

                    
                    I = I + 1
                Loop
                .Edit
                .Post

                sw_nuevo_item = False
            End With
            dxDBGrid1.Dataset.EnableControls
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
            txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
            txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
            txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")

            dxDBGrid1.OptionEnabled = True
            
            'Acumula los codigos de orden de compra en el arreglo wnumsord
            txtOcompra.Text = (pocompra)
            For J = 0 To 999
                If wnumsord(J) = "" Then
                        wnumsord(J) = (pocompra)
                        J = 999
                End If
            Next
            
            SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = True
        End If
        rsif3orden.Close
    End If
    rsif4orden.Close
End Sub

Private Sub GRABA_ING_ALMACEN_CENTRAL()
    Dim cnumvale        As String
    Dim ccampo          As String
    Dim nitems          As Integer
    Dim cdocum$
    Dim cmarca_costo    As String
    Dim costo As Double
    Dim Cant, cant_alm As Double
    Dim variable
    Dim precom
    Dim Fecha
    Dim sql, csql            As String
    Dim tmpVale         As ADODB.Recordset
    Dim regVale         As ADODB.Recordset
    Dim rsOrden         As ADODB.Recordset
    
    Dim nuevo As Boolean
    
    Rem SK ADD:
    Dim nsoles          As Double
    Dim ndolar          As Double
    
'    If sw_nuevo_documento = True Then
'        cnumvale = GENERA_NUMVALE(txtalmacen.Text, Format(Month(abofecha.Value), "00"), "I")
'
'        txtnumero.Text = cnumvale
'        ctipo = "A"
'    Else
'        cnumvale = txtnumero.Text
'        ctipo = "M"
'    End If
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .CodigoAlmacen = Trim(txtAlmacen.Text)
        .NumeroVale = Trim(txtnumero.Text)
        .Fecha = abofecha.Value
        
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
    amovs_cab(0).campo = "F4NUMVAL": amovs_cab(0).valor = txtnumero.Text: amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "F2CODALM": amovs_cab(1).valor = txtAlmacen.Text: amovs_cab(1).Tipo = "T"
    amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = abofecha.Value: amovs_cab(2).Tipo = "F"
    amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = txtconcepto.Text: amovs_cab(3).Tipo = "T"
    amovs_cab(4).campo = "F4TIPCAM": amovs_cab(4).valor = txttc.Text: amovs_cab(4).Tipo = "N"
    amovs_cab(5).campo = "F2CODPROV": amovs_cab(5).valor = txtproveedor: amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "F4CENTRO": amovs_cab(6).valor = txtccosto.Text: amovs_cab(6).Tipo = "T"
    amovs_cab(7).campo = "F4MONEDA": amovs_cab(7).valor = IIf(cmbmoneda.Text = "Soles", "S", "D"): amovs_cab(7).Tipo = "T"
    amovs_cab(8).campo = "F4SERGUIA": amovs_cab(8).valor = txtserie.Text: amovs_cab(8).Tipo = "T"
    amovs_cab(9).campo = "F4NUMGUIA": amovs_cab(9).valor = txtnumdoc.Text: amovs_cab(9).Tipo = "T"
    amovs_cab(10).campo = "F4TIPDOC": amovs_cab(10).valor = right(cmbtipo.Text, 2): amovs_cab(10).Tipo = "T"
    amovs_cab(11).campo = "F4SERDOC": amovs_cab(11).valor = txtserfac.Text: amovs_cab(11).Tipo = "T"
    
    If sw_importa_valedeingreso = True Then
        cdocum$ = Importaciones.txtnumero.Text
    Else
        cdocum$ = txtnumfac.Text
    End If
    
    amovs_cab(12).campo = "F4NUMDOC": amovs_cab(12).valor = cdocum$: amovs_cab(12).Tipo = "T"
    If ctipo = "A" Then
        amovs_cab(13).campo = "F4FECGRA": amovs_cab(13).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(13).Tipo = "F"
        amovs_cab(14).campo = "F4USEGRA": amovs_cab(14).valor = wusuario: amovs_cab(14).Tipo = "T"
    Else
        amovs_cab(13).campo = "F4FECMOD": amovs_cab(13).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(13).Tipo = "F"
        amovs_cab(14).campo = "F4USEMOD": amovs_cab(14).valor = wusuario: amovs_cab(14).Tipo = "T"
    End If
    amovs_cab(15).campo = "F4OBSERVA": amovs_cab(15).valor = txtobserva.Text: amovs_cab(15).Tipo = "T"
    '----------- OJO  REVISAR ESTO
'    If Trim(txtocompra.Text) <> "" Then
'        txtocompra.Text = Val(txtocompra.Text)
'    Else
'        txtocompra.Text = "0"
'    End If
    amovs_cab(16).campo = "numorden": amovs_cab(16).valor = txtOcompra.Text: amovs_cab(16).Tipo = "T"
    '------------------------------
    amovs_cab(17).campo = "F4REFERE": amovs_cab(17).valor = txtnomprov.Text: amovs_cab(17).Tipo = "T"
    
    Rem SK ADD:
    amovs_cab(18).campo = "F4TIPOVALE": amovs_cab(18).valor = "I": amovs_cab(18).Tipo = "T"
    amovs_cab(19).campo = "F4ORDTRA": amovs_cab(19).valor = Trim(txtIDOrdenProduccion.Text): amovs_cab(19).Tipo = "T"
    amovs_cab(20).campo = "EXPORTARVALE": amovs_cab(20).valor = IIf(CBool(chkExportarVale.Value), -1, 0): amovs_cab(20).Tipo = "N"
    amovs_cab(21).campo = "F1TIPPRV": amovs_cab(21).valor = right(cmbTipoAuxiliar.Text, 1): amovs_cab(21).Tipo = "T"
    
    
    
    
    
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
    amovs_det(12).campo = "F3JCG": amovs_det(12).valor = "": amovs_det(12).Tipo = "T"
    
    Rem SK ADD:
    amovs_det(13).campo = "F5CODPROORIGINAL": amovs_det(13).valor = "": amovs_det(13).Tipo = "T"
    amovs_det(14).campo = "F4NUMORD": amovs_det(14).valor = "": amovs_det(14).Tipo = "T"
    amovs_det(15).campo = "COD_SOLICITUD": amovs_det(15).valor = "": amovs_det(15).Tipo = "T"
    
    '------------------- CALCULA NUMERO DE FILAS
    nitems = 0
    'If cnDBTemp.State = adStateOpen Then cnDBTemp.Close
    'cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\TEMPLUS.MDB;Persist Security Info=False"
    'cnDBTemp.Open cconex_form
    abrirCnTemporal
    
    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
    
    RSDETALLE.Open "SELECT COUNT(ITEM) AS NITEM FROM " & "TMPVALEINGRESO" & " WHERE LEN(TRIM(CODPROD)) > 0 OR NOT ISNULL(CODPROD)", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not RSDETALLE.EOF Then
        nitems = Val("" & RSDETALLE.Fields("NITEM"))
    End If
    
    RSDETALLE.Close
    '---------------------------------------------
    cmarca_costo = ""
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select F1COSTO FROM SF1ORIGENES WHERE F1CODORI='" & txtconcepto.Text & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        cmarca_costo = "" & rst.Fields("F1COSTO")
    End If
    rst.Close
    '---------------------------------------------
    
    
    With objAyudaOrigen
        .inicializarEntidades
        
        .Codigo = Trim(txtconcepto.Text)
        
        .obtenerConfigOrigen
    End With
    
    'ReDim Values(12, nitems)
    ReDim Values(15, nitems)
    
    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
    
    RSDETALLE.Open "SELECT * FROM TMPVALEINGRESO", cnDBTemp
    
    If Not RSDETALLE.EOF Then
        nfil = 0
        RSDETALLE.MoveFirst
        Do While Not RSDETALLE.EOF
            dxDBGrid1.Dataset.RecNo = nfil + 1
            If Len(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "") > 0 Then
                Values(0, nfil) = txtnumero.Text
                Values(1, nfil) = dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & ""
                Values(2, nfil) = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value & ""
                
                If objAyudaOrigen.RegistrarCosto Then
                    If cmbmoneda.ListIndex = 0 Then
                        Values(3, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "") 'RSDETALLE.Fields("COSTOUNI") & "")
                        Values(4, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value & "")
                        Values(5, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value & "")
                        Values(8, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "") / Val(Format(txttc.Text, "0.0000"))
                        Values(9, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value & "") / Val(Format(txttc.Text, "0.00"))
                        Values(10, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value & "") / Val(Format(txttc.Text, "0.00"))
                    Else
                        Values(3, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "") * Val(Format(txttc.Text, "0.0000"))
                        Values(4, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value & "") * Val(Format(txttc.Text, "0.00"))
                        Values(5, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value & "") * Val(Format(txttc.Text, "0.00"))
                        Values(8, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "")
                        Values(9, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value & "")
                        Values(10, nfil) = Val(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value & "")
                    End If
                Else
                    nsoles = 0: ndolar = 0
                    
                    With objAyudaVale
                        .CodigoProducto = Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "")
                        .Fecha = abofecha.Value
                        
                        .CodigoMoneda = "S"
                        
                        nsoles = objAyudaVale.calcularCostoPromedio
                        
                        .CodigoMoneda = "D"
                        
                        ndolar = objAyudaVale.calcularCostoPromedio
                    End With
                    
                    Values(3, nfil) = nsoles
                    Values(4, nfil) = 0
                    Values(5, nfil) = Val(Format(nsoles * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value & ""), "0.00"))
                    Values(8, nfil) = Val(Format(nsoles / 1, "0.00"))
                    Values(9, nfil) = 0
                    Values(10, nfil) = Val(Format(nsoles * Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value & "") / 1, "0.00"))
                End If
                    
                Values(6, nfil) = txtAlmacen.Text
                Values(7, nfil) = Format(abofecha.Value, "dd/mm/yyyy")
                Values(11, nfil) = "I"
                
                Values(12, nfil) = cmarca_costo
                
                Rem SK:
                Values(13, nfil) = dxDBGrid1.Columns.ColumnByFieldName("CODPRODORIGINAL").Value & ""
                Values(14, nfil) = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & ""
                Values(15, nfil) = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value & ""
                
                nfil = nfil + 1
                
            End If
            
            RSDETALLE.MoveNext
        Loop
   
    End If
    RSDETALLE.Close
    
    cvalores = "1111111111111111"
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    cmes = Format(Month(abofecha.Value), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        'GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 17, cnn_dbbancos, ""
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 21, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            'GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 12, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 15, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
        ccampo = "F1VALING" & cmes
        'ACTUALIZA_ALMA_VALE cnumvale, ccampo, txtalmacen.Text
    Else    '--- Modificación
        
        '-------------------------------------------------------
        '------- GRABA CABECERA
        'GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 17, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtalmacen.Text & "'"
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 21, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & txtAlmacen.Text & "'"

        '-------------------------------------------------------
        
        '-------------------------------------------------------
        '------- GRABA DETALLE
        
''        Set tmpVale = New ADODB.Recordset
''        Set regVale = New ADODB.Recordset
''        Set rsOrden = New ADODB.Recordset
''        'Set rsOr = New ADODB.Recordset
''        BUSCA_VALE wcod_alm, cnumvale
''
''        SQL = "SELECT * FROM TMPVALEINGRESO"
''        tmpVale.Open SQL, cnDBTemp, adOpenStatic, adLockOptimistic
''        If ctipoadm_bd = "M" Then
''            cnn_dbbancos.Execute ("DELETE FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'")
''        Else
''            cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'")
''        End If
''        tmpVale.MoveFirst
''        Do While Not tmpVale.EOF
''            If ctipoadm_bd = "M" Then
''                SQL = "SELECT f3valvta, f4fecval From if3vales WHERE f2codalm='" & wcod_alm & "' and f5codpro='" & tmpVale.Fields("CODPROD") & "' ORDER BY f4fecval DESC limit 1;"
''            Else
''                SQL = "SELECT top 1 f3valvta, f4fecval From if3vales WHERE f2codalm='" & wcod_alm & "' and f5codpro='" & tmpVale.Fields("CODPROD") & "'"
''            End If
''
''                regVale.Open SQL, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
''            If Not (regVale.Bof Or regVale.EOF) Then
''            precom = regVale.Fields("f3valvta")
''            fecha = regVale.Fields("f4fecval")
''            Else
''            fecha = ""
''            precom = ""
''            End If
''
''            costo = Costo_Unitario2(tmpVale.Fields("CODPROD"), "S")(0)
''            Cant = Costo_Unitario2(tmpVale.Fields("CODPROD"), "S")(1)
''
''            csql = "UPDATE IF6ALMA SET F5COSPRO =  " & costo & ", F6STOCKACT = " & Cant & " WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "' and f2codalm ='" & wcod_alm & "'"
''            cnn_dbbancos.Execute csql
''            If compra = True Then
''                csql = "UPDATE IF5PLA SET F5ULTTC = " & Val(txttc.Text) & " , F5PRECOS =  " & costo & ", F5STOCKACT = " & Cant & ", F5FECUC ='" & fecha & "', F5VTANET = " & precom & " WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "'"
''            Else
''                csql = "UPDATE IF5PLA SET F5STOCKACT = " & Cant & " WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "'"
''            End If
''            cnn_dbbancos.Execute csql
''
''            regVale.Close
''            tmpVale.MoveNext
''        Loop
''        tmpVale.Close
''
''
        sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        
        'GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 12, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtalmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
        GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 15, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
        
                
    End If
    
    I = 0
    
    Do While wnumsord(I) <> ""
        csql = "INSERT INTO VALES_ORDENES (F4NUMORD, f4numval, f4codalm) values (" & Val(wnumsord(I)) & ", '" & cnumvale & "','" & wcod_alm & "')"
        cnn_dbbancos.Execute csql
        AlmacenaQuery_sql csql, cnn_dbbancos
        I = I + 1
    Loop
    
''    If sw_nuevo_documento = False Then
''        rsOrden.Open "SELECT numorden FROM VALES_ORDENES where f4numval = '" & cnumvale & "' and f2codalm = '" & wcod_alm & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
''        Do While Not rsOrden.EOF
''            wnumsord(i) = rsOrden.Fields("numorden").Value
''            i = i + 1
''            rsOrden.MoveNext
''        Loop
''        rsOrden.Close
''    End If
        
        For I = 0 To nfil - 1
            variable = Values(1, I)
            If left(cmbmoneda.Text, 1) = "S" Then
                precom = Values(3, I)
            Else
                precom = Values(3, I) / Val(Format(txttc.Text, "0.0000"))
            End If
            Fecha = Values(7, I)
            costo = Costo_Unitario2(variable, "S")(0)
            Cant = Costo_Unitario2(variable, "S")(1)
            cant_alm = Calcula_Cantidad(variable, "S", wcod_alm)
                                            
'            csql = "UPDATE IF6ALMA SET F5COSPRO =  " & costo & ", F6STOCKACT = " & cant_alm & " WHERE F5CODPRO = '" & variable & "' and F2CODALM ='" & wcod_alm & "'"
'
'            cnn_dbbancos.Execute csql
'            AlmacenaQuery_sql csql, cnn_dbbancos
            
'            If compra = True Then
'                If Left(CmbMoneda.Text, 1) = "S" Then
'                    csql = "UPDATE IF5PLA SET F5ULTTC = " & Val(txttc.Text) & " , F5PRECOS =  " & Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "") & ", F5STOCKACT = " & Cant & ", F5FECUC ='" & fecha & "', F5VTANET = " & precom & " WHERE F5CODPRO = '" & variable & "'"
'                Else
'                    csql = "UPDATE IF5PLA SET F5ULTTC = " & Val(txttc.Text) & " , F5PRECOS =  " & Val(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value & "") & ", F5STOCKACT = " & Cant & ", F5FECUC ='" & fecha & "', F5VTANETDOL = " & precom & " WHERE F5CODPRO = '" & variable & "'"
'                End If
'            Else
'                csql = "UPDATE IF5PLA SET F5STOCKACT = " & Cant & " WHERE F5CODPRO = '" & variable & "'"
'            End If
'            cnn_dbbancos.Execute csql
'            AlmacenaQuery_sql csql, cnn_dbbancos
          J = 0
          ' "OJO"  "SE MODIFICO PARA ACTUALIZAR ORDEN DE COMPRA
          
'           Do While wnumsord(j) <> 0
'            csql = "update if3orden set f3canfal = f3canpro  - " & Values(2, i) & " where numorden = " & txtOcompra & " and f3codpro = '" & variable & "'"
'            cnn_dbbancos.Execute csql
'            j = j + 1
'           Loop
          '---------------------
          'NUEVO PARA ACTUALIZAR SALDO DE ORDEN DE COMPRA
           
          ' rsOr.Open "SELECT * FROM If3Orden where numorden = " & txtOcompra & " and f3codpro = '" & variable & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
           If sw_Ord = True Then
            rsOr.Open "SELECT * FROM If3Orden where F4NUMORD = '" & txtOcompra & "' and f3codpro = '" & variable & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
           Else
                rsOr.Open "SELECT * FROM If3Orden where  f3codpro = '" & variable & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
           End If
            If Not (rsOr.EOF Or rsOr.Bof) Then
                If sw_Ord = True Then
                    If rsOr!F3CANFAL > 0 And rsOr!F3CANPRO = rsOr!F3CANFAL Then
                        csql = "update if3orden set f3canfal = f3canpro  - " & Values(2, I) & " where F4NUMORD = '" & txtOcompra & "' and f3codpro = '" & variable & "'"
                    Else
                        csql = "update if3orden set f3canfal = f3canfal - " & Values(2, I) & " where F4NUMORD = '" & txtOcompra & "' and f3codpro = '" & variable & "'"
                        
                    End If
                     cnn_dbbancos.Execute csql
                     AlmacenaQuery_sql csql, cnn_dbbancos
'                Else
'                    If rsOr!f3canfal > 0 And rsOr!F3CANPRO = rsOr!f3canfal Then
'                        csql = "update if3orden set f3canfal = f3canpro  - " & Values(2, i) & " where  f3codpro = '" & variable & "'"
'                    Else
'                        csql = "update if3orden set f3canfal = f3canfal - " & Values(2, i) & " where  f3codpro = '" & variable & "'"
'
'                    End If
'                     cnn_dbbancos.Execute csql
                End If
                
            End If
            rsOr.Close
    Next

    '-------------------------------------------------------
    
    '''graba envio
    If wIndEnvia = "*" Then 'cnn_dbEnvia
        sql = "delete FROM if4vales where f2codalm='" & txtAlmacen.Text & "' and f4numval='" & cnumvale & "'"
        cnn_dbEnvia.Execute sql
        AlmacenaQuery_sql sql, cnn_dbEnvia
        
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", "A", 21, cnn_dbEnvia, ""
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 15, cnn_dbEnvia, "", Values(), nfil - 1, cvalores, cmes, ""
        End If
    End If
    ''''''''''''''
    'BUSCA_VALE wcod_alm, cnumvale
End Sub

Private Sub GRABA_SALIDA_X_TRANSF()
Dim cnumvale        As String
Dim ccampo          As String
Dim I               As Integer

    If sw_nuevo_documento = True Then
        cnumvale = GENERA_NUMVALE(txtAlmacen.Text, Format(Month(abofecha.Value), "00"), "S")
        txtnumero.Text = cnumvale
        ctipo = "A"
    Else
        cnumvale = txtnumero.Text
        ctipo = "M"
    End If
    
    amovs_cab(0).campo = "F4NUMVAL": amovs_cab(0).valor = cnumvale: amovs_cab(0).Tipo = "T"
    amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = wconc_salxtransf: amovs_cab(3).Tipo = "T"
    For I = 0 To nfil - 1
        Values(0, I) = cnumvale
        Values(11, I) = "S"
    Next I
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    cmes = Format(Month(abofecha.Value), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 11, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
        
        ccampo = "F1VALSAL" & cmes
        ACTUALIZA_ALMA_VALE cnumvale, ccampo, txtAlmacen.Text
        
    Else    '--- Modificación
        
        '-------------------------------------------------------
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'"
        '-------------------------------------------------------
        '------- GRABA DETALLE
        sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        
        GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 11, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
    End If
    '-------------------------------------------------------
    '-------------------------------------------------------
End Sub

Private Sub GRABA_ING_ALMACEN_OBRA(palmacen As String)
Dim cnumvale        As String
Dim ccampo          As String
Dim I               As Integer

    If sw_nuevo_documento = True Then
        cnumvale = GENERA_NUMVALE(palmacen, Format(Month(abofecha.Value), "00"), "I")
        txtnumero.Text = cnumvale
        ctipo = "A"
    Else
        cnumvale = txtnumero.Text
        ctipo = "M"
    End If
    
    amovs_cab(0).campo = "F4NUMVAL": amovs_cab(0).valor = cnumvale: amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "F2CODALM": amovs_cab(1).valor = palmacen: amovs_cab(1).Tipo = "T"
    amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = wconc_ing_obra: amovs_cab(3).Tipo = "T"

    For I = 0 To nfil - 1
        Values(0, I) = cnumvale
        Values(6, I) = palmacen
        Values(11, I) = "I"
    Next I

    '-------------------------------------------------------
    '-------------------------------------------------------
    cmes = Format(Month(abofecha.Value), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 11, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
        ccampo = "F1VALING" & cmes
        ACTUALIZA_ALMA_VALE cnumvale, ccampo, palmacen
        
    Else    '--- Modificación
        
        '-------------------------------------------------------
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'"
        '-------------------------------------------------------
        '------- GRABA DETALLE
        sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
       GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 11, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'", Values(), nfil - 1, cvalores, cmes, "A"
    End If
    '-------------------------------------------------------

End Sub

Private Sub GRABA_VALE_SALIDA()
Dim cnumvale        As String
Dim ccampo          As String
Dim I               As Integer

    If sw_nuevo_documento = True Then
        cnumvale = GENERA_NUMVALE(txtAlmacen.Text, Format(Month(abofecha.Value), "00"), "S")
        txtnumero.Text = cnumvale
        ctipo = "A"
    Else
        cnumvale = txtnumero.Text
        ctipo = "M"
    End If
    
    amovs_cab(0).campo = "F4NUMVAL": amovs_cab(0).valor = cnumvale: amovs_cab(0).Tipo = "T"
    amovs_cab(3).campo = "F1CODORI": amovs_cab(3).valor = wconc_salida: amovs_cab(3).Tipo = "T"
    For I = 0 To nfil - 1
        Values(0, I) = cnumvale
        Values(11, I) = "S"
    Next I
    
    '-------------------------------------------------------
    cmes = Format(Month(abofecha.Value), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", ctipo, 11, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
        
        ccampo = "F1VALSAL" & cmes
        ACTUALIZA_ALMA_VALE cnumvale, ccampo, txtAlmacen.Text
        
    Else    '--- Modificación
        '-------------------------------------------------------
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", ctipo, 16, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'"
        '------- GRABA DETALLE
        sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'")
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 11, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & txtAlmacen.Text & "'", Values(), nfil - 1, cvalores, cmes, "A"
    End If
    '-------------------------------------------------------
    '-------------------------------------------------------

End Sub

Private Sub ACTUALIZA_ALMA_VALE(pnumvale As String, pcampo As String, palmacen As String)
Dim csql    As String
        
    csql = "UPDATE EF2ALMACENES SET " & pcampo & " =  '" & pnumvale & "' WHERE '" & pnumvale & "' > " & pcampo & " AND F2CODALM='" & palmacen & "'"
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
End Sub

Private Sub BUSCA_VALE(palmacen As String, pnumvale As String)
    Dim ncontador       As Long
    Dim cmedida         As String
    Dim I               As Integer
    Dim sw_nuevo_temp   As Boolean

    If rsif4vales.State = adStateOpen Then rsif4vales.Close
    
    rsif4vales.Open "SELECT * FROM IF4VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rsif4vales.EOF Then
        sw_nuevo_documento = False
        
        txtnumero.Text = pnumvale
        txtAlmacen.Text = "" & rsif4vales.Fields("F2CODALM")
        
        For I = 0 To cmbalmacen.ListCount - 1
            If txtAlmacen.Text = right(cmbalmacen.List(I), 2) Then
                cmbalmacen.ListIndex = I
                Exit For
            End If
        Next
        
        abofecha.Value = Format(rsif4vales.Fields("F4FECVAL"), "DD/MM/YYYY")
        txtconcepto.Text = "" & rsif4vales.Fields("F1CODORI")
        
        Rem SK ADD:
'        If VALIDA_CONCEPTO_INV(txtconcepto.Text) = True Then
'            If wpartida = "1" Then
'                lblproveedor.Visible = True
'                txtproveedor.Visible = True
'                txtnomprov.Visible = True
'            Else
'                lblproveedor.Visible = False
'                txtproveedor.Visible = False
'                txtnomprov.Visible = False
'            End If
'        End If
        For I = 0 To cmbconcepto.ListCount - 1
            If txtconcepto.Text = Trim(Mid(cmbconcepto.List(I), 200, 3)) Then
                cmbconcepto.ListIndex = I
                Exit For
            End If
        Next
        
        txttc.Text = Format(rsif4vales.Fields("F4TIPCAM"), "0.000")
        txtOcompra.Text = "" & rsif4vales.Fields("numorden")
        
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
        chkExportarVale.Value = IIf(CBool(rsif4vales!ExportarVale), vbChecked, vbUnchecked): chkExportarVale.Enabled = False
            
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
        
        
'        txtproveedor.Text = Trim(rsif4vales.Fields("F2CODPROV") & "")
'
'        If Len(Trim(txtproveedor.Text)) > 0 And txtproveedor.Text <> "99999999999" Then
'            If VALIDA_PROVEEDOR(txtproveedor.Text) = True Then
'                txtnomprov.Text = wnomprov
'            End If
'        Else
'            txtnomprov.Text = "" & rsif4vales.Fields("F4REFERE")
'        End If
        
        txtccosto.Text = "" & rsif4vales.Fields("F4CENTRO")
        If rsccosto.State = adStateOpen Then rsccosto.Close
        rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsccosto.EOF Then
            pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
        End If
        rsccosto.Close
            
        If "" & rsif4vales.Fields("F4MONEDA") = "S" Then
            cmbmoneda.ListIndex = 0
        Else
            cmbmoneda.ListIndex = 1
        End If
        
        txtserie.Text = "" & rsif4vales.Fields("F4SERGUIA")
        txtnumdoc.Text = "" & rsif4vales.Fields("F4NUMGUIA")
        txtobserva.Text = "" & rsif4vales.Fields("F4OBSERVA")
        
        If Len(Trim("" & rsif4vales.Fields("F4TIPDOC"))) > 0 Then
            For I = 0 To cmbtipo.ListCount
                If right(cmbtipo.List(I), 2) = "" & rsif4vales.Fields("F4TIPDOC") Then
                    cmbtipo.ListIndex = I
                    Exit For
                End If
            Next
        Else
            cmbtipo.ListIndex = -1
        End If
        txtserfac.Text = "" & rsif4vales.Fields("F4SERDOC")
        txtnumfac.Text = "" & rsif4vales.Fields("F4NUMDOC")
        
'        txtuupp.Text = rsif4vales.Fields("F4UUPP") & ""
'        If VALIDA_UUPP(txtuupp.Text) = True Then
'            pnluupp.Caption = wdeslocalidad
'        End If
        
        
        '--------------------------------------------------
        dxDBGrid1.Dataset.Close
        
        If sw_nuevo_documento = False Then
            DELETEREC_LOG "TMPVALEINGRESO", cnDBTemp
            AdicionaItem
            sw_nuevo_documento = True
        End If
        
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
        dxDBGrid1.Dataset.Active = True
    
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        
        dxDBGrid1.OptionEnabled = False
        dxDBGrid1.Dataset.DisableControls
        With dxDBGrid1.Dataset
            
            If rsif3vales.State = adStateOpen Then rsif3vales.Close
            rsif3vales.Open "SELECT * FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsif3vales.EOF Then
                I = 1
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
                    .FieldValues("ITEM") = I
                    
                    .FieldValues("F4NUMORD") = rsif3vales.Fields("F4NUMORD") & ""
                    .FieldValues("COD_SOLICITUD") = rsif3vales.Fields("COD_SOLICITUD") & ""
                    
                    .FieldValues("CODPROD") = rsif3vales.Fields("F5CODPRO") & ""
                    .FieldValues("CODPRODORIGINAL") = rsif3vales.Fields("F5CODPROORIGINAL") & ""
                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    rsif5pla.Open "SELECT IF5PLA.F5NOMPRO,IF5PLA.F7CODMED,IF5PLA.F5CODFAB, '' as DESMAR, IF5PLA.F5AFECTO FROM IF5PLA WHERE F5CODPRO='" & rsif3vales.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsif5pla.EOF Then
                        .FieldValues("UMEDIDA") = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsif5pla.Fields("F7CODMED"), "T", cnn_dbbancos)  '"" & rsif5pla.Fields("F7CODMED")
                        .FieldValues("DESCRIPCION") = "" & rsif5pla.Fields("F5NOMPRO")
                        .FieldValues("CODFAB") = "" & rsif5pla.Fields("F5CODFAB")
                        .FieldValues("marca") = ""
                        .FieldValues("afecto") = "" & rsif5pla.Fields("F5AFECTO")
                    End If
                    rsif5pla.Close
                    .FieldValues("CANTIDAD") = Format(rsif3vales.Fields("F3CANPRO"), "###,###,##0.000")
                    If "" & rsif4vales.Fields("F4MONEDA") = "S" Then
                        .FieldValues("COSTOUNI") = Format(rsif3vales.Fields("F3VALVTA"), "###,###,##0.0000")
                        .FieldValues("IGV") = Format(rsif3vales.Fields("F3IGV"), "###,###,##0.00")
                        .FieldValues("PVUNIT") = Format(rsif3vales.Fields("F3VALVTA") * (1 + wwigv / 100), "###,###,##0.00")
                        .FieldValues("VVTOTAL") = Format(rsif3vales.Fields("F3CANPRO") * rsif3vales.Fields("F3VALVTA"), "###,###,##0.00")
                        .FieldValues("TOTAL") = Format(rsif3vales.Fields("F3TOTITE"), "###,###,##0.00")
                    Else
                        .FieldValues("COSTOUNI") = Format(rsif3vales.Fields("F3VALDOL"), "###,###,##0.0000")
                        .FieldValues("IGV") = Format(rsif3vales.Fields("F3IGVDOL"), "###,###,##0.00")
                        .FieldValues("PVUNIT") = Format(rsif3vales.Fields("F3VALDOL") * wwigv, "###,###,##0.0000")
                        .FieldValues("VVTOTAL") = Format(rsif3vales.Fields("F3CANPRO") * rsif3vales.Fields("F3VALDOL"), "###,###,##0.00")
                        .FieldValues("TOTAL") = Format(rsif3vales.Fields("F3TOTDOL"), "###,###,##0.00")
                    End If
                    
                    rsif3vales.MoveNext
                    
                    I = I + 1
                Loop
                .Post

                sw_nuevo_item = False
            End If
            rsif3vales.Close
            
        End With
        With dxDBGrid1
            .Dataset.EnableControls
            .Dataset.Open
            .OptionEnabled = True
            txtTotvv.Text = Format(.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
            txtTotigv.Text = Format(.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
            txtTotpv.Text = Format(.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
        End With
        '--------------------------------------------------
    Else
        '----- No existe
        sw_nuevo_documento = False
        nuevo
        AdicionaItem
        AdicionaItem
        sw_nuevo_documento = True
    End If
    
    rsif4vales.Close
    
    Rem SK ADD:
    SSActiveToolBars1.Tools("ID_OC").Enabled = False
    
    listarGrilla
End Sub


Private Sub elimina(pnumvale As String, palmacen As String)
    'On Error GoTo ERROR_ELIMINA
    ReDim amovs(0 To 0) As a_grabacion
    Dim cmes            As String * 2
    Dim regAfec         As Integer
    Dim tmpVale         As ADODB.Recordset
    Dim regVale         As ADODB.Recordset
    Dim sql, csql            As String

    Set tmpVale = New ADODB.Recordset
    Set regVale = New ADODB.Recordset
       
        sql = "SELECT * FROM TMPVALEINGRESO"
        If tmpVale.State = adStateOpen Then tmpVale.Close
        tmpVale.Open sql, cnDBTemp, adOpenDynamic, adLockOptimistic
                
        If ctipoadm_bd = "M" Then
            sql = ("DELETE FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            
            sql = ("DELETE FROM IF4VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
        Else
            sql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            
            sql = ("DELETE * FROM IF4VALES WHERE F4NUMVAL = '" & pnumvale & "' AND F2CODALM = '" & palmacen & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
        End If
        If tmpVale.RecordCount > 0 Then
            tmpVale.MoveFirst
        End If
        
        
        
        ''REVISAR JCG
''        Do While Not tmpVale.EOF
''
''            If ctipoadm_bd = "M" Then
''                SQL = "SELECT f3valvta, f4fecval From if3vales WHERE f2codalm='" & wcod_alm & "' and f5codpro='" & tmpVale.Fields("CODPROD") & "' ORDER BY f4fecval DESC limit 1;"
''            Else
''                SQL = "SELECT TOP 1 f3valvta, f4fecval From if3vales WHERE f2codalm='" & wcod_alm & "' and f5codpro='" & tmpVale.Fields("CODPROD") & "' ORDER BY f4fecval DESC;"
''            End If
''            If regVale.State = adStateOpen Then regVale.Close
''            regVale.Open SQL, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
''
''            precom = regVale.Fields("f3valvta")
''            fecha = regVale.Fields("f4fecval")
''            costo = Costo_Unitario2(tmpVale.Fields("CODPROD"), "S")(0)
''            Cant = Costo_Unitario2(tmpVale.Fields("CODPROD"), "S")(1)
''
''            csql = "UPDATE IF6ALMA SET F5COSPRO =  " & costo & ", F6STOCKACT = " & Cant & " WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "' and f2codalm ='" & wcod_alm & "'"
''            cnn_dbbancos.Execute csql
''            If compra = True Then
''                csql = "UPDATE IF5PLA SET F5PRECOS =  " & costo & ", F5STOCKACT = " & Cant & ", F5FECUC ='" & fecha & "', F5VTANET = " & precom & " WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "'"
''            Else
''                csql = "UPDATE IF5PLA SET F5PRECOS =  " & costo & ", F5STOCKACT = " & Cant & ", F5FECUC ='" & fecha & "' WHERE F5CODPRO = '" & tmpVale.Fields("CODPROD") & "'"
''            End If
''            cnn_dbbancos.Execute csql
''            regVale.Close
''            tmpVale.MoveNext
''        Loop
        tmpVale.Close

        
        If wIndEnvia = "*" Then
        
            sql = "delete FROM if4vales where f2codalm='" & palmacen & "' and f4numval='" & pnumvale & "'"
            cnn_dbEnvia.Execute sql, regAfec
            AlmacenaQuery_sql sql, cnn_dbEnvia
            If regAfec = 0 Then
                'vales
                'guardar el almacen,numvale en una tabla(de vales eliminados)
                'al recibir leer de esa tabla y ejecutar la setencia con los datos
                sql = "insert into VALESELIMINADOS(F2CODALM, F4NUMVAL) " & _
                                    " values('" & palmacen & "','" & pnumvale & "')"
                cnn_dbEnvia.Execute sql
                AlmacenaQuery_sql sql, cnn_dbEnvia
            End If
        End If
        
        'nuevo
        'AdicionaItem
    
    Exit Sub
    
ERROR_ELIMINA:
    MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atención"
    Resume Next
    
End Sub

Private Sub ACTUALIZA_CANT_OCOMPRA(pnumorden As Double)
Dim rstempo     As New ADODB.Recordset
    
    If rstempo.State = adStateOpen Then rstempo.Close
    rstempo.Open "SELECT * FROM " & "TMPVALEINGRESO" & "", cnDBTemp, adOpenDynamic, adLockOptimistic
    If Not rstempo.EOF Then
        rstempo.MoveFirst
        Do While Not rstempo.EOF
            If Len(Trim(rstempo.Fields("CODPROD") & "")) > 0 Then
                
                sql = ("UPDATE IF3ORDEN SET F3CANFAL = F3CANFAL - " & rstempo.Fields("CANTIDAD") & " WHERE F4NUMORD=" & pnumorden & " AND F3CODPRO = '" & rstempo.Fields("CODPROD") & "'")
                cnn_dbbancos.Execute sql
                AlmacenaQuery_sql sql, cnn_dbbancos
            End If
            rstempo.MoveNext
        Loop
    End If
    rstempo.Close

End Sub

Private Sub txtuupp_DblClick()

    txtuupp_KeyDown 113, 0
    
End Sub

Private Sub txtuupp_KeyDown(KeyCode As Integer, Shift As Integer)

'    If KeyCode = 113 Then
'        wcodlocalidad = "": wdeslocalidad = ""
'        hlp_uupp.Show 1
'        If Len(Trim(wcodlocalidad)) > 0 Then
'            txtuupp.Text = Trim(wcodlocalidad)
'            pnluupp.Caption = Trim(wdeslocalidad)
'            txtuupp_KeyPress 13
'        End If
'    End If

End Sub

Private Sub txtuupp_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtserie.SetFocus
    End If

End Sub

Private Sub txtuupp_LostFocus()

    If Len(Trim(txtuupp.Text)) > 0 Then
        If VALIDA_UUPP(txtuupp.Text) = True Then
            pnluupp.Caption = wdeslocalidad
        Else
            MsgBox "Unidad de producción no existe", vbInformation + vbDefaultButton1, "Atención"
            txtuupp.Text = "": txtuupp.SetFocus
        End If
    End If

End Sub

Public Sub CargarMarca()
Dim rst As ADODB.Recordset

    Set rst = New ADODB.Recordset
    sql = "SELECT f2codmar, f2desmar FROM ef2marcas order by f2desmar"
    If rst.State = adStateOpen Then rst.Close
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cnn_dbbancos.Open "Provider =Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB;Persist Security Info=False"
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        cmbmarca.AddItem "** Todos **"
        Do While Not rst.EOF
            wmarca = Space(255)
            Mid(wmarca, 1) = rst!f2desmar
            Mid(wmarca, 200) = rst!f2codmar
            cmbmarca.AddItem wmarca
            rst.MoveNext
        Loop
        cmbmarca.ListIndex = 0
    End If
    rst.Close

End Sub

Private Sub BuscaImportacion(pimportacion As String)
Dim I                   As Integer
Dim sw_nuevo_temp       As Boolean

    sw_nuevo_temp = False
    If rst.State = adStateOpen Then rst.Close
    sql = "SELECT IMPORT_CAB.*, IMPORT_DET.* " _
    & "FROM IMPORT_CAB INNER JOIN IMPORT_DET ON IMPORT_CAB.F4NumImp = IMPORT_DET.F4NumImp " _
    & "where import_cab.f4numimp='" & pimportacion & "'"
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        With dxDBGrid1.Dataset
            dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
            dxDBGrid1.Dataset.Active = True
            
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
                    
            dxDBGrid1.OptionEnabled = False
            dxDBGrid1.Dataset.DisableControls
    
            sw_nuevo_item = True
            I = 1
            Do While Not rst.EOF
                If sw_nuevo_temp = False Then
                    sw_nuevo_temp = True
                    .Edit
                Else
                    .Append
                End If
                .FieldValues("ITEM") = I
                .FieldValues("CODPROD") = Trim("" & rst.Fields("F5CODPRO"))
                .FieldValues("CODFAB") = Trim("" & rst.Fields("F3CODFAB"))
        
                If rsif5pla.State = adStateOpen Then rsif5pla.Close
                rsif5pla.Open "SELECT F5NOMPRO,F7CODMED,f5marca FROM IF5PLA WHERE F5CODPRO = '" & Trim("" & rst.Fields("F5CODPRO")) & "'", cnn_dbbancos, adOpenStatic, adLockOptimistic
                If Not rsif5pla.EOF Then
                    .FieldValues("DESCRIPCION") = "" & rsif5pla.Fields("F5NOMPRO")
                    .FieldValues("UMEDIDA") = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsif5pla.Fields("F7CODMED"), "T", cnn_dbbancos)  '"" & rsif5pla.Fields("F7CODMED")
                    .FieldValues("marca") = "" & rsif5pla.Fields("f5marca")
                End If
                rsif5pla.Close
                wcantidad = Val("" & rst.Fields("F3CANTIDAD"))
                wcostouni = Val("" & rst.Fields("F3PREUNI"))
                xigv = (wcantidad * wcostouni) * (wIgv / 100)
                wtot = wcantidad * wcostouni + xigv
                .FieldValues("CANTIDAD") = Format(Val("" & rst.Fields("F3CANTIDAD")), "###,##0.0000")
                .FieldValues("COSTOUNI") = Format(Val("" & rst.Fields("F3PREUNI")), "###,##0.0000")
                .FieldValues("IGV") = Format(xigv, "###,##0.00")
                .FieldValues("TOTAL") = Format(wtot, "###,##0.00")
                I = I + 1
                rst.MoveNext
            Loop
            .Post
            dxDBGrid1.Dataset.EnableControls
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
            dxDBGrid1.OptionEnabled = True
            cmbconcepto.ListIndex = 7
        End With
    End If
    If rst.State = adStateOpen Then rst.Close

    Exit Sub
    sw_nuevo_temp = False
    If rsif4orden.State = adStateOpen Then rsif4orden.Close
    rsif4orden.Open "SELECT * FROM IF4ORDEN WHERE F4NUMORD = " & Val(pocompra) & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsif4orden.EOF Then
        '--------------------- C. COSTO
        txtccosto.Text = Trim("" & rsif4orden.Fields("F4CENTRO"))
        If rsccosto.State = adStateOpen Then rsccosto.Close
        rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsccosto.EOF Then
            pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
        End If
        rsccosto.Close
        '-----------------------------------------------------
        If Trim("" & rsif4orden.Fields("F4TIPMON")) = "S" Then
            cmbmoneda.ListIndex = 0
        ElseIf Trim("" & rsif4orden.Fields("F4TIPMON")) = "D" Then
            cmbmoneda.ListIndex = 1
        End If
        
        txtuupp.Text = rsif4orden.Fields("F4UUPP") & ""
        If VALIDA_UUPP(txtuupp.Text) = True Then
            pnluupp.Caption = wdeslocalidad
        End If
        
        If rsif3orden.State = adStateOpen Then rsif3orden.Close
        rsif3orden.Open "SELECT * FROM IF3ORDEN WHERE F4NUMORD = " & Val(pocompra) & " AND F3CANFAL > 0 ", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif3orden.EOF Then
            dxDBGrid1.Dataset.ADODataset.ConnectionString = cnDBTemp
            dxDBGrid1.Dataset.Active = True
        
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
                        
            dxDBGrid1.OptionEnabled = False
            dxDBGrid1.Dataset.DisableControls
            With dxDBGrid1.Dataset
                I = 1
                sw_nuevo_item = True
                rsif3orden.MoveFirst
                Do While Not rsif3orden.EOF
                    If sw_nuevo_temp = False Then
                        sw_nuevo_temp = True
                        .Edit
                    Else
                        .Append
                    End If
                    .FieldValues("ITEM") = I
                    .FieldValues("CODPROD") = Trim("" & rsif3orden.Fields("F3CODPRO"))
                    .FieldValues("CODFAB") = Trim("" & rsif3orden.Fields("F3CODFAB"))
                    
                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    rsif5pla.Open "SELECT F5NOMPRO,F7CODMED,f5marca FROM IF5PLA WHERE F5CODPRO = '" & Trim("" & rsif3orden.Fields("F3CODPRO")) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsif5pla.EOF Then
                        .FieldValues("DESCRIPCION") = "" & rsif5pla.Fields("F5NOMPRO")
                        .FieldValues("UMEDIDA") = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsif5pla.Fields("F7CODMED"), "T", cnn_dbbancos)  '"" & rsif5pla.Fields("F7CODMED")
                        .FieldValues("marca") = "" & rsif5pla.Fields("f5marca")
                    End If
                    rsif5pla.Close
                    .FieldValues("CANTIDAD") = Format(Val("" & rsif3orden.Fields("F3CANFAL")), "###,##0.0000")
                    .FieldValues("COSTOUNI") = Format(Val("" & rsif3orden.Fields("F3PRECOS")), "###,##0.0000")
                    .FieldValues("IGV") = Format(Val("" & rsif3orden.Fields("F3IGV")), "###,##0.00")
                    .FieldValues("TOTAL") = Format(Val("" & rsif3orden.Fields("F3TOTAL")), "###,##0.00")
                    rsif3orden.MoveNext
                    I = I + 1
                Loop
                .Post
                sw_nuevo_item = False
            End With
            dxDBGrid1.Dataset.EnableControls
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
            dxDBGrid1.OptionEnabled = True
        End If
        rsif3orden.Close
    End If
    rsif4orden.Close
    
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
        .ADODataset.CommandText = "SELECT * FROM TMPVALEINGRESO ORDER BY ITEM, DESCRIPCION"
        
        .Active = False
        .Active = True
        
        dxDBGrid1.KeyField = "ITEM"
        
        .Close
        .Open
    End With
    
    txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
    txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
    txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
End Sub

Private Sub adicionarItemVale()
    With dxDBGrid1.Dataset
        .Close
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE * FROM TMPVALEINGRESO"
        
'        abrirCnTemporal
'
'        cnDBTemp.Execute "INSERT INTO TMPVALEINGRESO(ITEM, CODPROD, CODPRODORIGINAL) VALUES(1, NULL, NULL)"

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
    Dim rstTemporalRenumerarI As New ADODB.Recordset
    Dim dblItem As Double
    
    If rstTemporalRenumerarI.State = 1 Then rstTemporalRenumerarI.Close
    
    rstTemporalRenumerarI.Open "SELECT * FROM TMPVALEINGRESO ORDER BY ITEM, DESCRIPCION", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTemporalRenumerarI.EOF Then
        rstTemporalRenumerarI.MoveFirst
        
        dblItem = 0
        
        'dxDBGrid1.Dataset.Close
        
        Do While Not rstTemporalRenumerarI.EOF
            dblItem = dblItem + 1
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "TMPVALEINGRESO "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "ITEM = " & dblItem & " "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "TRIM(F4NUMORD & '') = '" & Trim(rstTemporalRenumerarI!F4NUMORD & "") & "' AND "
            SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstTemporalRenumerarI!COD_SOLICITUD & "") & "' AND "
            SqlCad = SqlCad & "TRIM(CODPROD & '') = '" & Trim(rstTemporalRenumerarI!codprod & "") & "' AND "
            SqlCad = SqlCad & "TRIM(CODPRODORIGINAL & '') = '" & Trim(rstTemporalRenumerarI!CODPRODORIGINAL & "") & "'"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            rstTemporalRenumerarI.MoveNext
        Loop
            'dxDBGrid1.Dataset.Open
    End If
End Sub

Private Sub recalcularItems()
    With objAyudaOrden
        If dxDBGrid1.Dataset.RecordCount = 0 Then Exit Sub
        
        dxDBGrid1.Dataset.First
        
        Do While Not dxDBGrid1.Dataset.EOF
            .inicializarEntidadesDetalle
            
            'Entregar Datos a Clase
            .PorcentajeImpuesto = IIf(right(cmbtipo.Text, 2) = "02", gretenc, IIf(right(cmbtipo.Text, 2) = "03", 0, wwigv)) / 100
            .SignoImpuesto = IIf(right(cmbtipo.Text, 2) = "02", -1, 1)
            
            .Cantidad = Val(dxDBGrid1.Dataset.FieldValues("CANTIDAD") & "")
            .CantidadMaxima = Val(dxDBGrid1.Dataset.FieldValues("CANTIDADMAX") & "")

            .PorcentajeDemasia = 0 'Val(dxDBGrid1.Dataset.FieldValues("F3PORCDEMASIA") & "") / 100
            
            .PrecioSinImpuesto = Val(dxDBGrid1.Dataset.FieldValues("COSTOUNI") & "")
            .PrecioConImpuesto = Val(dxDBGrid1.Dataset.FieldValues("PVUNIT") & "")
            
            .PorcentajeDscto = Val(dxDBGrid1.Dataset.FieldValues("PORDESC") & "") / 100
            .TotalDscto = Val(dxDBGrid1.Dataset.FieldValues("VALDESC") & "")

            .Afecto = IIf(right(cmbtipo.Text, 2) = "03", False, IIf(Trim(dxDBGrid1.Dataset.FieldValues("AFECTO") & "") = "*", True, False))
            
            'Calcular
            .calculosPorItem
            
            'Copiar Resultados
            dxDBGrid1.Dataset.Edit

            dxDBGrid1.Dataset.FieldValues("COSTOUNI") = .PrecioSinImpuesto
            dxDBGrid1.Dataset.FieldValues("PVUNIT") = .PrecioConImpuesto
            dxDBGrid1.Dataset.FieldValues("PORDESC") = Val(Format(.PorcentajeDscto * 100, "#0.00"))
            'dxDBGrid1.Dataset.FieldValues("F3CANPROFINAL") = .CantidadFinal
            dxDBGrid1.Dataset.FieldValues("VALDESC") = .TotalDscto
            
            dxDBGrid1.Dataset.FieldValues("COSTOUNINETO") = .PrecioNetoSinImpuesto
            
            dxDBGrid1.Dataset.FieldValues("VVTOTAL") = .BasePorItem
            'dxDBGrid1.Dataset.FieldValues("F3MONINA") = .ExoneradoPorItem
            dxDBGrid1.Dataset.FieldValues("IGV") = .ImpuestoPorItem
            dxDBGrid1.Dataset.FieldValues("TOTAL") = .TotalPorItem
            
            dxDBGrid1.Dataset.Post
            
            dxDBGrid1.Dataset.Next
        Loop
            txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,0.00")
            txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,0.00")
            txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,0.00")
    End With
End Sub

Private Sub limpiarCajas()
    Me.Caption = "Vale de Ingreso a Almacen"
    
    txtnumero.Text = vbNullString
        lblNumeroValeExterno.Caption = "< ID Externo >"
    
    cmbalmacen.ListIndex = -1: cmbalmacen.Enabled = True
        txtAlmacen.Text = vbNullString
    cmbconcepto.ListIndex = -1
        txtconcepto.Text = vbNullString
    
    cmbTipoAuxiliar.ListIndex = -1
        txtproveedor.Text = vbNullString: txtnomprov.Text = vbNullString
    txtccosto.Text = vbNullString: pnlccosto.Caption = vbNullString
    txtserie.Text = vbNullString: txtnumdoc.Text = vbNullString
    
    cmbtipo.ListIndex = -1
    txtserfac.Text = vbNullString: txtnumfac.Text = vbNullString
    dtpFechaDoc.Value = Format(Date, "Short Date")
    dtpFechaDoc.Value = Null
    
    fraOrdenProduccion.Enabled = True
    cmbCategoriaTipo.ListIndex = -1
    cmbCategoriaTipo.Text = vbNullString
    lblIdCategoriaTipo.Caption = vbNullString
    txtNroOrdenProduccion.Text = vbNullString
    txtIDOrdenProduccion.Text = vbNullString
    
    With abofecha
        .Value = Format(Date, "Short Date")
        .CalendarBackColor = vbWhite
        .CalendarForeColor = vbBlack
        .Font.Bold = False
        
        lblFechaMensaje.Visible = False
        
        If ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoUsarFechaPredeterminada", "l") = "1" Then
            .Value = ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoFechaPredeterminada", "l")
            .CalendarBackColor = vbRed
            .CalendarForeColor = vbWhite
            .CalendarTrailingForeColor = vbGreen
            .Font.Bold = True
            
            lblFechaMensaje.Visible = True
        End If
    End With
        
    cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, "S", "IZQ", 1)
    txttc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", abofecha.Value, "F")), "#.000")
    
    txtOcompra.Text = vbNullString
    txtobserva.Text = vbNullString
    
    chkVerObservaciones.Value = vbUnchecked
    
    chkVerObservaciones_Click
    
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
    SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
    SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
    SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
    SSActiveToolBars1.Tools("ID_OC").Enabled = False
    
    dxDBGrid1.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPVALEINGRESO"
    
    SSFrame1.Enabled = True
    dxDBGrid1.Enabled = True
    
    SSFrame2.Visible = False
    
    txtTotvv.Text = "0.00"
    txtDscto.Text = "0.00"
    txtTotigv.Text = "0.00"
    txtTotpv.Text = "0.00"
    
    bolObviarCierre = True
    
    SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
    SSActiveToolBars1.Tools.ITEM("CerrarVale").Enabled = False
    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
    
    bolObviarCierre = False
    
    pnlRegistroCompra.Visible = False
    txtRegistroCompra.Text = vbNullString
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
                txtAlmacen.Text = .CodigoAlmacen
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
            dtpFechaDoc.Value = IIf(.FechaUltima <> vbNullString, .FechaUltima, Null)
            
            'abrirCnDBMilano
            
            fraOrdenProduccion.Enabled = False
            txtIDOrdenProduccion.Text = .OrdenTrabajo
                If .OrdenTrabajo <> vbNullString Then
                    lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                    'cmbCategoriaTipo.ListIndex = ModUtilitario.seleccionarItem(cmbCategoriaTipo, ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N"), "DER", 10)
                    cmbCategoriaTipo.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T")
                    txtNroOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                End If
            
            abofecha.Value = .Fecha
            cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, .CodigoMoneda, "IZQ", 1)
            txttc.Text = Format(.TipoCambio, "#.000")
            
            txtOcompra.Text = .NumeroOrdenCompra
            txtobserva.Text = .observaciones
            
            If .RegistroCompra <> vbNullString Then
                pnlRegistroCompra.Visible = True
                txtRegistroCompra.Text = .RegistroCompra
            End If
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) > 0 Then
                renumerarItemVale
            End If
            
            listarGrillaVale
            
            If dxDBGrid1.Dataset.RecordCount = 0 Then
                dxDBGrid1.Dataset.Close
                
                adicionarItemVale
            Else
                recalcularItems
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
            If Trim(txtconcepto.Text) = "XCS" Then
                SSFrame1.Enabled = False
                dxDBGrid1.Enabled = False
                
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
                SSActiveToolBars1.Tools("ID_OC").Enabled = False
                
                chkVerObservaciones.Value = vbChecked
                
                chkVerObservaciones_Click
            ElseIf .VB1 Then
                SSFrame1.Enabled = False
                dxDBGrid1.Enabled = False
                
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
                SSActiveToolBars1.Tools("ID_OC").Enabled = False
                
                chkVerObservaciones.Value = vbChecked
                
                chkVerObservaciones_Click
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
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: ConsultarVale"
    
    Err.Clear
End Sub

Private Sub validarCajas()
    On Error Resume Next
    
    'Validacion del Punto (PC) que origina el Vale
    'ModMilano.abrirCnDBMilano
    
'    If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'        MsgBox "Su computador no esta registrado y/o habilitado. Consulte con su" & vbNewLine & vbNewLine & _
'                "administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'        Exit Sub
'    End If
    
'    If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")), "T", _
'                                                                    "AND TABLA = 'INGRESO'")) = 0 Then
'
'        MsgBox "El Punto de Venta no cuenta con correlativo habilitado de INGRESO." & vbNewLine & vbNewLine & _
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

    If Trim(txtproveedor.Text) <> vbNullString And Trim(lblproveedor.Caption) = vbNullString Then
        MsgBox "Persona seleccionada inválida, verifique..", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
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
    
    If Val(txttc.Text) <= 0 Then
'        MsgBox "Tipo de Cambio incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName
        obtiene_tipodecambio (abofecha.Value)
        txttc.Text = TCVenta
'        txttc.SetFocus
'
'        Exit Sub
    End If
    
    If Trim(txtconcepto.Text) = "XOP" And Trim(txtIDOrdenProduccion.Text) = vbNullString Then
        MsgBox "ID de Orden de Produccion incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName

        cmbCategoriaTipo.SetFocus

        Exit Sub
    End If

    If Trim(txtconcepto.Text) <> "XOP" And Trim(txtIDOrdenProduccion.Text) <> vbNullString Then
        MsgBox "Concepto incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName

        cmbconcepto.SetFocus

        Exit Sub
    End If

    If Trim(txtconcepto.Text) = "XC0" Then
        If Trim(txtproveedor.Text) = vbNullString Then
            MsgBox "El Campo Proveedor es obligatorio.", vbInformation + vbOKOnly, App.ProductName
            
            txtproveedor.SetFocus
            
            Exit Sub
        End If
        
        If Trim(txtserie.Text) = vbNullString And Trim(txtnumdoc.Text) = vbNullString And _
           Trim(txtserfac.Text) = vbNullString And Trim(txtnumfac.Text) = vbNullString Then

            If MsgBox("¿Desea guardar la Compra sin numeros de documento de referencia?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
        End If
        txtserie.Text = UCase(txtserie.Text)
        If Len(Trim(txtserie.Text)) <> 4 Then
            MsgBox "La Serie debe tener 4 dígitos.", vbInformation + vbOKOnly, App.ProductName
            txtserie.SetFocus
            Exit Sub
        End If
        
        If cmbtipo.ListIndex <> -1 And Trim(txtnumfac.Text) <> vbNullString Then
            With objAyudaComprobante
                .Codigo = right(cmbtipo.Text, 2)
                
                .obtenerConfigComprobante
                
                If .EsOficial Then
                    If Not IsDate(dtpFechaDoc.Value) Then
                        MsgBox "El Campo Fecha Documento es obligatorio.", vbInformation + vbOKOnly, App.ProductName
                        
                        dtpFechaDoc.SetFocus
                        
                        Exit Sub
                    Else
                        If CDate(abofecha.Value) < CDate(dtpFechaDoc.Value) Then
                            If MsgBox("La Fecha de Ingreso es menor a la Fecha de Documento" & vbNewLine & _
                                        "¿Desea guardar la Compra?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                                
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End With
            
            If Trim(txtnumfac.Text) = vbNullString Then
                If MsgBox("¿Desea guardar la Compra sin CONSIGNAR el NUMERO del Documento?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                    Exit Sub
                End If
            Else
                With objAyudaEmpresa
                    .CodigoEmpresa = wF1Dir
                    
                    .obtenerConfigEmpresa
                End With
                
                txtnumfac.Text = Format(txtnumfac.Text, objAyudaEmpresa.FormatoNumDocCompra)
            End If
            
            If strNumeroVale = vbNullString Then
                With objAyudaVale
                    .inicializarEntidades
                    
                    .CodigoProveedor = Trim(txtproveedor.Text)
                    
                    .SerieGuia = Trim(txtserie.Text)
                    .NumeroGuia = Trim(txtnumdoc.Text)
                    
                    .CodTipoComprobante = right(cmbtipo.Text, 2)
                    .SerieDocumento = Trim(txtserfac.Text)
                    .NumeroDocumento = Trim(txtnumfac.Text)
                    
                    If .verificarExistenciaPorNumRef Then
                        MsgBox "Ya existe un Vale de Ingreso registrado con los Numeros de Documentos de Referencia, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        txtnumdoc.SetFocus
                        
                        Exit Sub
                    End If
                    
                    .inicializarEntidades
                End With
            End If
        End If
        
        If cmbmoneda.ListIndex <> -1 Then
            With objAyudaProveedor
                .inicializarEntidades
                
                .Codigo = Trim(txtproveedor.Text)
                
                .obtenerConfigProveedor
                
                If UCase(left(cmbmoneda.Text, 1)) <> .CodigoMoneda Then
                    If MsgBox("El proveedor seleccionado tiene como moneda predeterminada '" & IIf(.CodigoMoneda = "S", "Soles", "Dólares") & "' ¿Desea guardar la Compra con la moneda en '" & Trim(cmbmoneda.Text) & "'?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                        Exit Sub
                    End If
                End If
                
                .inicializarEntidades
            End With
        End If
        
        If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(COSTOUNI) AS CANTIDAD", "TMPVALEINGRESO", "VAL(COSTOUNI & '')", "0", "N", "AND TRIM(CODPROD & '') <> '' GROUP BY COSTOUNI")) > 0 Then
            If MsgBox("Se han detectado Items sin Precio ingresado, ¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                dxDBGrid1.SetFocus
                
                Exit Sub
            End If
        End If
    End If
    
    If Trim(txtOcompra.Text) <> vbNullString And Trim(txtconcepto.Text) <> "XC0" Then
        MsgBox "Concepto seleccionado erroneo, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        cmbconcepto.SetFocus
        
        Exit Sub
    End If
    
    If CDate(abofecha.Value) > CDate(Date) Then
        MsgBox "Fecha ingresada inválida, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        abofecha.SetFocus
        
        Exit Sub
    End If
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) = 0 Then
        MsgBox "Registro no cuenta con Detalle, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        dxDBGrid1.SetFocus
        
        Exit Sub
    End If
        
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                .CodigoAlmacen = Trim(txtAlmacen.Text)
                
                .FechaInicioMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 0, 1)
                .FechaFinMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 1, 0)
                
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
                
                .CodigoAlmacen = Trim(txtAlmacen.Text)
                
                .FechaInicioMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 0, 1)
                .FechaFinMes = DateSerial(Year(CDate(abofecha.Value)), Val(Month(CDate(abofecha.Value))) + 1, 0)
                
                If .verificarCierreVale Then
                    MsgBox "Imposible registrar Vale, periodo ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                    
                    .inicializarEntidades
                    .inicializarEntidadesAdicionales
                    
                    Exit Sub
                End If
            End With
        End If
        
        guardarVale
    End If
End Sub

Private Sub guardarVale()
    On Error GoTo errGuardarVale
    
    Dim rstTemporalGuardarValeI As New ADODB.Recordset
    Dim dblItem As Double
    
    Set objVale = New ClsVale
    
    dxDBGrid1.Dataset.Close
    
    With objVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .NumeroVale = Trim(txtnumero.Text)
        .NumeroValeExterno = Trim(lblNumeroValeExterno.Caption)
        
        .CodigoAlmacen = Trim(txtAlmacen.Text)
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
            .FechaUltima = Format(dtpFechaDoc.Value, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = Trim(txtIDOrdenProduccion.Text)
        
        .Fecha = Format(abofecha.Value, "Short Date")
        .CodigoMoneda = left(cmbmoneda.Text, 1)
        .TipoCambio = Val(txttc.Text)
        
        .NumeroOrdenCompra = Trim(txtOcompra.Text)
        .observaciones = Trim(txtobserva.Text)
        
        .RegistroCompra = Trim(txtRegistroCompra.Text)
        
        .ExportarVale = CBool(chkExportarVale.Value)
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If .guardarVale Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnn_dbbancos.Execute .SQLSelectAlter
            
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            If rstTemporalGuardarValeI.State = 1 Then rstTemporalGuardarValeI.Close
            
            rstTemporalGuardarValeI.Open "SELECT * FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') <> '' AND CANTIDAD > 0 ORDER BY ITEM, DESCRIPCION", cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTemporalGuardarValeI.EOF Then
                rstTemporalGuardarValeI.MoveFirst
                
                With objAyudaOrigen
                    .inicializarEntidades
                    
                    .Codigo = objVale.CodigoOrigen
                    
                    .obtenerConfigOrigen
                End With
                
                dblItem = 0
                
                Do While Not rstTemporalGuardarValeI.EOF
                    .inicializarEntidadesDetalle
                    
                    dblItem = dblItem + 1
                    
                    .ITEM = dblItem
                    
                    .CodigoProducto = Trim(rstTemporalGuardarValeI!codprod & "")
                    .CodigoProductoOriginal = Trim(rstTemporalGuardarValeI!CODPRODORIGINAL & "")
                    .Cantidad = Val(rstTemporalGuardarValeI!Cantidad & "")
                    .CantidadMaxima = Val(rstTemporalGuardarValeI!CANTIDADMAX & "")
                    
                    .ValorVenta = 0
                    .IGV = 0
                    .TOTAL = 0
                    
                    .IgvDol = 0
                    .TotalDol = 0
                    
                    If objAyudaOrigen.RegistrarCosto Then
                        .PorcentajeDscto = Val(Format(Val(rstTemporalGuardarValeI!PORDESC & ""), "#0.00"))
                        
                        Select Case .CodigoMoneda
                            Case "S"
                                .ValorVenta = Val(Format(Val(rstTemporalGuardarValeI!COSTOUNI & ""), "#0.0000"))
                                .IGV = Val(Format(Val(rstTemporalGuardarValeI!IGV & ""), "#0.0000"))
                                .TOTAL = Val(Format(Val(rstTemporalGuardarValeI!TOTAL & ""), "#0.00"))
                                
                                .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#.0000"))
                                .IgvDol = Val(Format(.IGV / .TipoCambio, "#0.0000"))
                                .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                                
                                .MontoDscto = IIf(.PorcentajeDscto > 0, Val(Format((.Cantidad * .ValorVenta) * (.PorcentajeDscto / 100), "#.0000")), 0)
                            Case Else
                                .ValorVentaDol = Val(Format(Val(rstTemporalGuardarValeI!COSTOUNI & ""), "#0.0000"))
                                .IgvDol = Val(Format(Val(rstTemporalGuardarValeI!IGV & ""), "#0.0000"))
                                .TotalDol = Val(Format(Val(rstTemporalGuardarValeI!TOTAL & ""), "#0.00"))
                                
                                .ValorVenta = Val(Format(.ValorVentaDol * .TipoCambio, "#.0000"))
                                .IGV = Val(Format(.IgvDol * .TipoCambio, "#0.0000"))
                                .TOTAL = Val(Format(.TotalDol * .TipoCambio, "#0.00"))
                                
                                .MontoDscto = IIf(.PorcentajeDscto > 0, Val(Format((.Cantidad * .ValorVentaDol) * (.PorcentajeDscto / 100), "#.0000")), 0)
                        End Select
                    Else
                        objSqlAyudaVale.CodigoAlmacen = .CodigoAlmacen
                        objSqlAyudaVale.CodigoMoneda = .CodigoMoneda
                        objSqlAyudaVale.CodigoProducto = .CodigoProducto
                        objSqlAyudaVale.Fecha = .Fecha
                        
                        .ValorVenta = 0 'objSqlAyudaVale.calcularCostoPromedioV2
'
'                        If .ValorVenta <= 0 Then
'                            .ValorVenta = objSqlAyudaVale.obtenerUltimoCostoPromedioV2
'                        End If
                        
                        .IGV = 0
                        .TOTAL = 0 'Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                        
                        .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                        .IgvDol = 0
                        .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    End If
                    
                    .NumeroOrdenCompra = Trim(rstTemporalGuardarValeI!F4NUMORD & "")
                    .Requerimiento = Trim(txtccosto.Text)
                    
                    .ObservacionesPorItem = Trim(rstTemporalGuardarValeI!observaciones & "")
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    rstTemporalGuardarValeI.MoveNext
                Loop
            End If
            
            'Exportar el Vale
'            If ModMilano.exportarValeAserverSQLv2(.CodigoAlmacen, .NumeroVale, lblNumeroValeExterno, fraProceso, pgbProceso) Then
'                'MsgBox "ID Ingreso en Sistema Externo: " & lblNumeroValeExterno.Caption & ".", vbInformation + vbOKOnly, App.ProductName
'
'                If .CodigoOrigen = "XOP" And .OrdenTrabajo <> vbNullString Then
'                    verificarProductoDevueltoOP .CodigoAlmacen, .NumeroVale
'                End If
'            End If
            
            strCodAlmacen = .CodigoAlmacen
            strNumeroVale = .NumeroVale
            
'            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                guardarValeSql
'            End If
            
            Select Case .CodigoOrigen
                Case "XC0"
                    If Trim(txtOcompra.Text) <> vbNullString Then
                        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                            verificarAtencionOrdenSql .CodigoAlmacen, .NumeroVale
                        Else
                            verificarAtencionOrden .CodigoAlmacen, .NumeroVale
                        End If
                    End If
                    
                    If cmbtipo.ListIndex <> -1 And Trim(txtnumfac.Text) <> vbNullString Then
                        With objAyudaComprobante
                            .Codigo = right(cmbtipo.Text, 2)
                            
                            .obtenerConfigComprobante
                            
                            If .EsOficial Then
                                'Generar Registro de Compra
                                
                                If .CodCompraRegistro <> vbNullString Then
                                    'exportarRegistroCompra
                                    
                                    exportarRegistroCompraDAO
                                Else
                                    MsgBox "Imposible Generar el Registro de Compra, el Tipo de Documento seleccionado no cuenta con la Configuracion 'Tipo de Registro de Compra', verifique en el Sistema de Bancos: [Mantenimientos] >> [Documentos] y vuelva a GUARDAR el Vale.", vbInformation + vbOKOnly, App.ProductName
                                End If
                            End If
                        End With
                    End If
            End Select
            
            strCodAlmacen = .CodigoAlmacen
            strNumeroVale = .NumeroVale
            
'            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                guardarValeSql
'            End If
            
            consultarVale
            
            MsgBox "Se ha Actualizado el Vale de Ingreso " & .NumeroVale & "." & vbNewLine & _
                    ".", vbInformation + vbOKOnly, App.ProductName
        Else
            listarGrillaVale
        End If
    End With
    
    Set objVale = Nothing
    
    Exit Sub
errGuardarVale:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: GuardarVale"
    
    Err.Clear
End Sub

Private Sub eliminarVale()
    On Error GoTo errEliminarVale
    
    Set objVale = New ClsVale
    
    With objVale
        .CodigoAlmacen = Trim(txtAlmacen.Text)
        .NumeroVale = Trim(txtnumero.Text)
        
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
        
        'No permitir Eliminacion de Vale, si este genero Registro de Compra
        With objAyudaCompra
            .inicializarEntidades
            
            .CodProveedor = objVale.CodigoProveedor
            .TipoDocumento = objVale.CodTipoComprobante
            .SerieDocumento = objVale.SerieDocumento
            .NumeroDocumento = objVale.NumeroDocumento
            wrucprov = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", objVale.CodigoProveedor, "T")
            
            'IDENTIFICAR el Registro de Compra en base al Proveedor, Tipo, Numero y Serie de Documento
            .obtenerRegistroCompraProvYdocumento
            
            If .MesMovimiento <> vbNullString And .NumeroMovimiento <> vbNullString Then
'                MsgBox "Imposible Eliminar Vale, cuenta con Registro de Compra generado." & vbNewLine & vbNewLine & _
'                        "Recomendaciones:" & vbNewLine & _
'                        "- Ingresar a partir del Vale, otro documento para el reuso del Registro de Compra." & vbNewLine & _
'                        "- Coordinar con Contabilidad y Tesoreria, la Eliminacion del Registro de Compra (Perdida del Correlativo).", vbInformation + vbOKOnly, App.ProductName
'
'                .inicializarEntidades
'
'                Exit Sub
                .eliminarCompra
                
                sql = vbNullString
                sql = sql & "DELETE * FROM SCFDOCU1 WHERE "
                sql = sql & "RUCDOC = '" & wrucprov & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = '" & .NumeroDocumento & "'"

                
                cnn_dbbancos.Execute sql
                
                sql = vbNullString
                sql = sql & "DELETE * FROM SCFDOCU2 WHERE "
                sql = sql & "RUCDOC = '" & wrucprov & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = '" & .NumeroDocumento & "'"

                
                cnn_dbbancos.Execute sql
                
            End If
            
            .inicializarEntidades
        End With
        
        If MsgBox("¿Desea eliminar el Vale con No. " & .NumeroVale & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If Val(lblNumeroValeExterno.Caption) > 0 Then
'                If Not ModMilano.anularValeExterno("I", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T"), fraProceso, pgbProceso) Then
'                    Me.MousePointer = vbDefault
'
'                    Exit Sub
'                End If
            End If
            
            If .eliminarVale Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                If .CodigoOrigen = "XC0" And .NumeroOrdenCompra <> vbNullString Then
'                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                        verificarAtencionOrdenSql .CodigoAlmacen, .NumeroVale, True
'                    Else
                        verificarAtencionOrden .CodigoAlmacen, .NumeroVale, True
'                    End If
                End If
                
                strCodAlmacen = .CodigoAlmacen
                strNumeroVale = .NumeroVale
                
'                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                    eliminarValeSql
'                End If
                
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
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: EliminarVale"
    
    Err.Clear
End Sub

Private Sub exportarRegistroCompra()
    On Error GoTo errExportarRegistroCompra
    
    'Dim rstVale As ADODB.Recordset
    Dim rstValeDet As ADODB.Recordset
    
    Dim cnValeDAO As DAO.Database
    Dim rstVale As DAO.Recordset
    
    Dim strConceptoCompra As String
    Dim dblItem As Double
    Dim dblMontoCancelado As Double
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .CodigoProveedor = Trim(txtproveedor.Text)
        .CodTipoComprobante = right(cmbtipo.Text, 2)
        .SerieDocumento = Trim(txtserfac.Text)
        .NumeroDocumento = Trim(txtnumfac.Text)
        
        Set cnValeDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        Set rstVale = cnValeDAO.OpenRecordset("IF4VALES")  '"SELECT * FROM IF4VALES ORDER BY F4REGCOM DESC")
        
        rstVale.Index = "IndiceProvComprobante"
        
        rstVale.Seek "=", .TipoVale, .CodigoProveedor, .CodTipoComprobante, .SerieDocumento, .NumeroDocumento
        
        Set rstVale = .obtenerRstValeCompraPorProvYdocumento
        Set rstValeDet = .obtenerRstValeDetalleCompraPorProvYdocumento
        
        strConceptoCompra = "Ingreso de Compras a Almacén" '.obtenerConceptoCompraPorProvYdocumento
        
        'MsgBox "CORTE 1: OBTIENE RST"
        
        'If Not rstVale.EOF Then
        If Not rstVale.NoMatch Then
            'rstVale.MoveFirst
            
            Do While (Trim(rstVale!F4TIPOVALE & "") = .TipoVale And Trim(rstVale!F2CODPROV & "") = .CodigoProveedor And Trim(rstVale!F4TIPDOC & "") = .CodTipoComprobante And Trim(rstVale!F4SERDOC & "") = .SerieDocumento And Trim(rstVale!F4NUMDOC & "") = .NumeroDocumento)
                
                
                rstVale.MoveNext
            Loop
            
            'rstVale.MoveFirst
            
            With objAyudaCompra
                .inicializarEntidades
                
                If Trim(rstVale!F4REGCOM & "") <> vbNullString Then
                    .MesMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), 1, InStr(1, Trim(rstVale!F4REGCOM & ""), "-") - 1)
                    .NumeroMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), InStr(1, Trim(rstVale!F4REGCOM & ""), "-") + 1)
                Else
                    .MesMovimiento = Year(CDate(dtpFechaDoc.Value)) & Format(Month(CDate(dtpFechaDoc.Value)), "00")
                    
                    Rem SK ADD: ADICIONAR LINEAS DE EVALUACION:
                    'SI EL MES DEL DOCUMENTO DE LA COMPRA SE ENCUENTRA CERRADO, PASARLO AUTOMATICAMENTE AL MES SIGUIENTE.
                    If .verificarCierreCompra Then
                        MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                        
                        .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                            IIf(Month(CDate(dtpFechaDoc.Value)) = 12, "01", Format(Month(CDate(dtpFechaDoc.Value)) + 1, "00"))
                        
                        If .verificarCierreCompra Then
                            MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                    "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                    "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                                                        
                            .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                                IIf(Month(CDate(dtpFechaDoc.Value)) >= 11, Format(CDate(dtpFechaDoc.Value) - 10, "00"), Format(Month(CDate(dtpFechaDoc.Value)) + 2, "00"))
                                                
                        End If
                    End If
                    '---------------------------------------------------------------------
                    .TipoRegistro = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCOMPRAREG", "DOCUMENTOS", "F2CODDOC", Trim(rstVale!F4TIPDOC & ""), "T")
                    
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
                
                '.FechaRegistro = Format(Date, "Short Date")
                .FechaDocumento = Format(Trim(rstVale!F4FECULT & ""), "Short Date")
                
                
                If Val(Format(.FechaDocumento, "YYYYMM")) = Val(.MesMovimiento) Then
                    .FechaRegistro = .FechaDocumento
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) < Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 0, 1)
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) > Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 1, 0)
                End If
                
                
                .CodMoneda = Trim(rstVale!F4MONEDA & "")
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaDocumento, "F"))
                '.ConceptoCompra = strConceptoCompra 'Trim(rstVale!F4OBSERVA & "")
                
                .CodFormaPago = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                
                With objAyudaFormaPago
                    .inicializarEntidades
                    
                    .Codigo = objAyudaCompra.CodFormaPago
                    
                    .obtenerConfigFormaPago
                End With
                
                .ConceptoCompra = left(strConceptoCompra, 255) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(rstValeDet!CUENTA & ""), "T")
                
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
                .CentroCosto = Trim(rstVale!F4CENTRO & "")
                .PorcentajeIGV = wIgv
                
                .FechaReg = Format(rstVale!F4FECVAL, "Short Date")
                .UsuarioReg = wusuario
                .FechaMod = Format(Date, "Short Date")
                .UsuarioMod = wusuario
                
                MsgBox "CORTE 2: VERIFICACION DE CIERRE Y ASIGNACION DE DATOS."
                
                If .guardarCompra(True) Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM REGISMOV WHERE F4MESMOV = '" & .MesMovimiento & "' AND F4NUMMOV = '" & .NumeroMovimiento & "'"
                    
                    cnn_dbbancos.Execute .SQLSelectAlter
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    If Not rstValeDet.EOF Then
                        rstValeDet.MoveFirst
                        
                        dblItem = 0
                        
                        fraProceso.Visible = True
                        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstValeDet)
                        pgbProceso.Value = 0
                        fraProceso.Caption = "Registrando Detalle de R.C. ..."
                        
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .ITEM = dblItem
                            
                            .CtaContableDet = Trim(rstValeDet!CUENTA & "")
                            .AuxiliarDet = Trim(rstValeDet!AUXILIAR & "")
                            
                            If .CtaContableDet <> vbNullString Then
                                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T") = vbNullString Then
                                    With objAyudaGasto
                                        .inicializarEntidades
                                        
                                        .Codigo = vbNullString
                                        .Base = "G"
                                        .CuentaContable = objAyudaCompra.CtaContableDet
                                        .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", objAyudaCompra.CtaContableDet, "T")
                                        .TipoGasto = "P"
                                        .Moneda = left(cmbmoneda.Text, 1)
                                        .GrupoFlujo = vbNullString
                                        
                                        If .guardarGasto Then
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                        End If
                                        
                                        .inicializarEntidades
                                    End With
                                End If
                            End If
                            
                            .CodigoGastoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T")
                            
                            .NumeroOrden = Trim(rstValeDet!NROOC & "")
                            .ConceptoDet = left(strConceptoCompra, 100) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableDet, "T")
                            
                            .Cantidad = 1
                            .PrecioUnitario = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .SubTotalDet = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .Afecto = IIf(Trim(rstValeDet!F5AFECTO & "") = "*", True, False)
                            
                            .DebHab = "D"
                            
                            'Acumular
                            .BaseImponible = .BaseImponible + (.SubTotalDet * IIf(.Afecto, 1, 0))
                            .MontoInafecto = .MontoInafecto + (.SubTotalDet * IIf(.Afecto, 0, 1))
                            .TotalIGV = .TotalIGV + (Val(rstValeDet!IGV & "") * IIf(.Afecto, 1, 0))
                            '.Descuento = .Descuento + Val(rstValeDet!DSCTO & "")
                            .TotalFacturado = .TotalFacturado + Val(rstValeDet!TOTAL & "") '((.BaseImponible + .MontoInafecto + .TotalIGV) - .Descuento)
                            
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
                            
                            DoEvents
                            
                            pgbProceso.Value = pgbProceso.Value + 1
                            fraProceso.Caption = "Registrando Detalle de R.C. ... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
                                                        
                            rstValeDet.MoveNext
                        Loop
                            
                            MsgBox "CORTE 3: DETALLE GABRADO DE RC"
                            
                            'ACTUALIZAR POSTERIOR A LA GRABACION
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4OCOMPRA = '" & left(.OrdenCompra, 255) & "', "
                            .SQLSelectAlter = .SQLSelectAlter & "F4BASIMP = " & Val(Format(.BaseImponible, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MONINA = " & Val(Format(.MontoInafecto, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & Val(Format(.TotalIGV, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4DCTO = " & Val(Format(.Descuento, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = " & Val(Format(.TotalFacturado, "#0.00")) & " "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                            
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
'                            .SQLSelectAlter = vbNullString
'                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
'                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
'                            .SQLSelectAlter = .SQLSelectAlter & "SET "
'                            .SQLSelectAlter = .SQLSelectAlter & "F4OCOMPRA = '" & left(.OrdenCompra, 255) & "', "
'                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = (VAL(F4BASIMP & '') + VAL(F4MONINA & '') + VAL(F4IGV & '') + VAL(F4OTRIMP & '') + VAL(F4REDSUMA & '')) - (VAL(F4FONAVI & '') + VAL(F4DCTO & '') + VAL(F4MONTORET & '') + VAL(F4REDRESTA & '')) "
'                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
'                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
'                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
'
'                            cnn_dbbancos.Execute .SQLSelectAlter
'
'                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End If
                
                If .obtenerCompra Then
                    
                    With objAyudaPagDcto
                        .Correlativo = IIf(objAyudaCompra.Correlativo = 0, -1, objAyudaCompra.Correlativo)
                        
                        .TipoIngreso = "1"
                        .ITEM = 1
                        .NumeroComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2ABREV", "DOCUMENTOS", "F2CODDOC", objAyudaCompra.TipoDocumento, "T")
                        
                        .NumeroComprobante = .NumeroComprobante & IIf(objAyudaCompra.SerieDocumento <> vbNullString, objAyudaCompra.SerieDocumento & "/", vbNullString) & objAyudaCompra.NumeroDocumento
                        
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
                        .Concepto = left(strConceptoCompra, 255)  '"INGRESO DE COMPRAS A ALMACEN."
                        .Detalle = left(strConceptoCompra, 255) '"INGRESO DE COMPRAS A ALMACEN."
                        .referencia = left(strConceptoCompra, 255) '"INGRESO DE COMPRAS A ALMACEN."
                        
                        If .obtenerPagDcto Then
                            If .SaldoFacturado = .TotalFacturado Then
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                .SaldoFacturado = objAyudaCompra.TotalFacturado
                            ElseIf .SaldoFacturado < .TotalFacturado Then
                                .SaldoFacturado = objAyudaCompra.TotalFacturado - (.TotalFacturado - .SaldoFacturado)
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                            End If
                        End If
                        
                        If .guardarPagDcto(False) Then
                            
                            MsgBox "CORTE 4: PAG_DCTO GRABADO"
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            'If objAyudaCompra.Correlativo = 0 Then
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
                            'End If
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
                        .SQLSelectAlter = .SQLSelectAlter & "F4VALESING = '" & Trim(left(.ValeIngreso, 255)) & "' "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        sql = vbNullString
                        sql = sql & "DELETE * FROM SCFDOCU1 WHERE "
                        sql = sql & "RUCDOC = '" & wrucprov & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = '" & .NumeroDocumento & "'"
                        
                        cnn_dbbancos.Execute sql
                        
                        sql = vbNullString
                        sql = sql & "DELETE * FROM SCFDOCU2 WHERE "
                        sql = sql & "RUCDOC = '" & wrucprov & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = '" & .NumeroDocumento & "'"
                
                        cnn_dbbancos.Execute sql
                        
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "INSERT INTO SCFDOCU1 ( RUCDOC, TIPDOC, SERIE, NRODOC, FCHEMI, FCHVTO, TIPMON, TIPADQ, BASIMP, VALIGV, IMPTOT, NOMBRE, TIPDOC_O, SERIE_O, NRODOC_O, FCHDOC_O,FCHING,NRODET,FCHDET ) "
                        .SQLSelectAlter = .SQLSelectAlter & "SELECT REGISDOC.F4RUCPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, Format(F4FECHA,'ddmmyyyy') AS Expr1, Format(F4FECVEN,'ddmmyyyy') AS Expr2, IIf(F4MONEDA='D',2,1) AS Expr3, 'A' AS Expr4, REGISDOC.F4BASIMP, REGISDOC.F4IGV, REGISDOC.F4TOTAL, REGISDOC.F4NOMPRV, REGISDOC.TIPODOCREF, REGISDOC.SERDOCREF, REGISDOC.NUMDOCREF, REGISDOC.FECDOCREF, Format(REGISDOC.F4FECHING,'ddmmyyyy') AS Expr5,REGISDOC.NUMDETRACCION, REGISDOC.FECHADETRACCION "
                        .SQLSelectAlter = .SQLSelectAlter & "FROM REGISDOC "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
     
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "INSERT INTO SCFDOCU2 ( RUCDOC, TIPDOC, SERIE, NRODOC, SECUEN, CUENTA, CENCOS, DESCRI, BASIMP ) "
                        .SQLSelectAlter = .SQLSelectAlter & "SELECT REGISDOC.F4RUCPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, REGISMOV.F3ITEM, REGISMOV.F3CTACON, REGISDOC.F4OBRA, LEFT(REGISMOV.F3CONCEPTO,40), FORMAT(REGISMOV.F3IMPORTE,'0.00') "
                        .SQLSelectAlter = .SQLSelectAlter & "FROM REGISDOC  INNER JOIN REGISMOV ON (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) AND (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                    
                    MsgBox "Ingreso exportado a Registro de Compras:" & vbNewLine & _
                            "Mes de Registro: " & .MesMovimiento & vbNewLine & _
                            "Numero de Movimiento: " & .NumeroMovimiento, vbInformation + vbOKOnly, App.ProductName
                End If
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
        End If
        
        .inicializarEntidades
    End With
    
    fraProceso.Visible = False
    
    Exit Sub
errExportarRegistroCompra:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description & vbNewLine & vbNewLine & "Registro de Compra NO EXPORTADO correctamente a Tesoreria, intente volver a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub exportarRegistroCompraDAO()
    On Error GoTo errExportarRegistroCompraDAO
    
    Dim rstVale As New ADODB.Recordset
    Dim rstValeDet As New ADODB.Recordset
    
    Dim cnDAO As DAO.Database
    
    Dim rstValeDAO As DAO.Recordset
    Dim rstValeDetDAO As DAO.Recordset
    
    Dim strConceptoCompra As String
    Dim dblItem As Double
    Dim dblMontoCancelado As Double
    Dim strMensajeErrorPagDcto As String
    
    abrirCnContaTabla
    
    Set cnDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    With objAyudaProveedor
        .Codigo = Trim(txtproveedor.Text)
        
        .obtenerConfigProveedor
    End With
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .CodigoProveedor = Trim(txtproveedor.Text)
        .CodTipoComprobante = right(cmbtipo.Text, 2)
        .SerieDocumento = Trim(txtserfac.Text)
        .NumeroDocumento = Trim(txtnumfac.Text)
        
        Rem Descargar en el Temporal Cabecera y Detalle de los Vales que coincidan con los datos ingresados anteriormente.
        
        Set rstValeDAO = cnDAO.OpenRecordset("IF4VALES")
        
        rstValeDAO.Index = "IndiceProvComprobante"
        
        rstValeDAO.Seek "=", .TipoVale, .CodigoProveedor, .CodTipoComprobante, .SerieDocumento, .NumeroDocumento, "XC0"
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILVALECAB"
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILVALEDET"
        
        If Not rstValeDAO.NoMatch Then
            Do While (Trim(rstValeDAO!F4TIPOVALE & "") = .TipoVale And _
                        Trim(rstValeDAO!F2CODPROV & "") = .CodigoProveedor And _
                        Trim(rstValeDAO!F4TIPDOC & "") = .CodTipoComprobante And _
                        Trim(rstValeDAO!F4SERDOC & "") = .SerieDocumento And _
                        Trim(rstValeDAO!F4NUMDOC & "") = .NumeroDocumento)
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TMPUTILVALECAB "
                
                SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
                
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "* "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "IF4VALES "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODALM = '" & Trim(rstValeDAO!f2codalm & "") & "' AND "
                SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstValeDAO!F4NUMVAL & "") & "'"
                
                cnn_dbbancos.Execute SqlCad
                
                Set rstValeDetDAO = cnDAO.OpenRecordset("IF3VALES")

                rstValeDetDAO.Index = "IndiceValeDetalle"

                rstValeDetDAO.Seek "=", Trim(rstValeDAO!f2codalm & ""), Trim(rstValeDAO!F4NUMVAL & "")
                
                If Not rstValeDetDAO.NoMatch Then
                    Do While (Trim(rstValeDetDAO!f2codalm & "") = Trim(rstValeDAO!f2codalm & "") And _
                                Trim(rstValeDetDAO!F4NUMVAL & "") = Trim(rstValeDAO!F4NUMVAL & ""))
                        
                        With objAyudaBien
                            .Codigo = Trim(rstValeDetDAO!f5codpro & "")
                            
                            .obtenerConfigBien
                        End With
                        
                        With objAyudaUM
                            .Codigo = objAyudaBien.CodUM
                            
                            .obtenerConfigUM
                        End With
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO TMPUTILVALEDET("
                        SqlCad = SqlCad & "F2CODALM, F4NUMVAL, TIPO, F4FECVAL, F5CODPRO, F5CODPROORIGINAL, "
                        SqlCad = SqlCad & "F3CANPRO, F3VALVTA, F3IGV, F3TOTITE, F3IGVDOL, F3VALDOL, F3TOTDOL, "
                        SqlCad = SqlCad & "F3PORCENTAJEDSCTO, F3MONTODSCTO, CUENTA, AUXILIAR, AFECTO, "
                        SqlCad = SqlCad & "UM, SUBFAMILIA) "
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!f2codalm & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!F4NUMVAL & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!Tipo & "") & "', "
                        SqlCad = SqlCad & "CVDATE('" & Trim(rstValeDetDAO!F4FECVAL & "") & "'), "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!f5codpro & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!F5CODPROORIGINAL & "") & "', "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3CANPRO & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3VALVTA & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3IGV & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3TOTITE & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3IGVDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3VALDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3TOTDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3PORCENTAJEDSCTO & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3MONTODSCTO & "") & ", "
                        SqlCad = SqlCad & "'" & IIf(objAyudaProveedor.OrigenProveedor = "N", objAyudaBien.CtaContable, objAyudaBien.CtaContableImportacion) & "', "
                        SqlCad = SqlCad & "'" & IIf(objAyudaProveedor.OrigenProveedor = "N", objAyudaBien.Anexo, objAyudaBien.AnexoImportacion) & "', "
                        'SqlCad = SqlCad & "'" & IIf(objAyudaBien.Afecto, "*", vbNullString) & "', "
                        SqlCad = SqlCad & "'" & IIf(Val(rstValeDetDAO!F3IGV & "") = 0, vbNullString, "*") & "', "
                        SqlCad = SqlCad & "'" & objAyudaUM.Abreviatura & "', "
                        SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7DESCON", "SF7NIVEL02", "F7CODCON", objAyudaBien.CodigoSubFamilia, "T") & "'"
                        SqlCad = SqlCad & ")"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
                        
                        rstValeDetDAO.MoveNext
                        
                        If rstValeDetDAO.EOF Then Exit Do
                    Loop
                End If
                
                rstValeDAO.MoveNext
                
                If rstValeDAO.EOF Then Exit Do
            Loop
        End If
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "* "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILVALECAB "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "F4REGCOM DESC, "
        SqlCad = SqlCad & "F4FECVAL DESC"
        
        If rstVale.State = 1 Then rstVale.Close
        
        rstVale.Open SqlCad, cnDBTemp, adOpenDynamic, adLockOptimistic
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.CUENTA, "
        SqlCad = SqlCad & "DET.AUXILIAR, DET.f5codpro,"
        SqlCad = SqlCad & "SUM( VAL(FORMAT(DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL), '#0.0000')) ) AS SUBTOTAL, "
        SqlCad = SqlCad & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3IGV, DET.F3IGVDOL)) AS IGV, "
        SqlCad = SqlCad & "SUM( VAL(FORMAT((DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL)) * (DET.F3PORCENTAJEDSCTO / 100), '#0.0000')) ) AS DSCTO, "
        SqlCad = SqlCad & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE , DET.F3TOTDOL)) AS TOTAL, "
        SqlCad = SqlCad & "DET.AFECTO, "
        SqlCad = SqlCad & "DET.F4NUMORD AS NROOC "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILVALEDET AS DET "
        SqlCad = SqlCad & "LEFT JOIN TMPUTILVALECAB AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.CUENTA, "
        SqlCad = SqlCad & "DET.AUXILIAR, DET.f5codpro,"
        SqlCad = SqlCad & "DET.AFECTO, "
        SqlCad = SqlCad & "DET.F4NUMORD"
        
        If rstValeDet.State = 1 Then rstValeDet.Close
        
        rstValeDet.Open SqlCad, cnDBTemp, adOpenDynamic, adLockOptimistic
        
        strConceptoCompra = obtenerConceptoCompra
        
        If Not rstVale.EOF Then
            
            With objAyudaCompra
                .inicializarEntidades
                
                .CodProveedor = Trim(rstVale!F2CODPROV & "")
                .TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                .SerieDocumento = Trim(rstVale!F4SERDOC & "")
                .NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                
                'IDENTIFICAR el Registro de Compra en base al Proveedor, Tipo, Numero y Serie de Documento
                .obtenerRegistroCompraProvYdocumento
                
                If .MesMovimiento = vbNullString And .NumeroMovimiento = vbNullString Then
                    If Trim(rstVale!F4REGCOM & "") <> vbNullString Then
                        .MesMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), 1, InStr(1, Trim(rstVale!F4REGCOM & ""), "-") - 1)
                        .NumeroMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), InStr(1, Trim(rstVale!F4REGCOM & ""), "-") + 1)
                    Else
                        .MesMovimiento = Year(CDate(dtpFechaDoc.Value)) & Format(Month(CDate(dtpFechaDoc.Value)), "00")
                        
                        Rem SK ADD: ADICIONAR LINEAS DE EVALUACION:
                        'SI EL MES DEL DOCUMENTO DE LA COMPRA SE ENCUENTRA CERRADO, PASARLO AUTOMATICAMENTE AL MES SIGUIENTE.
                        If .verificarCierreCompra Then
                            MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                    "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                    "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                            
                            .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                                IIf(Month(CDate(dtpFechaDoc.Value)) = 12, "01", Format(Month(CDate(dtpFechaDoc.Value)) + 1, "00"))
                            
                            If .verificarCierreCompra Then
                                MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                        "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                        "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                                                            
                                .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                                    IIf(Month(CDate(dtpFechaDoc.Value)) >= 11, Format(CDate(dtpFechaDoc.Value) - 10, "00"), Format(Month(CDate(dtpFechaDoc.Value)) + 2, "00"))
                                                    
                            End If
                        End If
                        '---------------------------------------------------------------------
                        .TipoRegistro = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCOMPRAREG", "DOCUMENTOS", "F2CODDOC", Trim(rstVale!F4TIPDOC & ""), "T")
                        
                        .NumeroMovimiento = vbNullString
                    End If
                End If
                
                '.CodProveedor = Trim(rstVale!F2CODPROV & "")
                .NomProveedor = Replace(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T"), "'", "' & Chr(39) & '", 1)
                .DireccionProveedor = Replace(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T"), "'", "' & Chr(39) & '", 1)
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
                    
                '.TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                '.SerieDocumento = Trim(rstVale!F4SERDOC & "")
                '.NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                
                .CodigoCategoria = 1
                
                '.FechaRegistro = Format(Date, "Short Date")
                .FechaDocumento = Format(Trim(rstVale!F4FECULT & ""), "Short Date")
                
                
                If Val(Format(.FechaDocumento, "YYYYMM")) = Val(.MesMovimiento) Then
                    .FechaRegistro = .FechaDocumento
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) < Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 0, 1)
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) > Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 1, 0)
                End If
                
                
                .CodMoneda = Trim(rstVale!F4MONEDA & "")
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaDocumento, "F"))
                
                If .TipoCambio = 0 Then
                    obtiene_tipodecambio (.FechaDocumento)
                    .TipoCambio = TCVenta
                    If .TipoCambio = 0 Then
                        MsgBox "Imposible generar Registro de Compra, Tipo de Cambio del " & .FechaDocumento & " no registrado, verifique y vuelva a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
                        
                        .inicializarEntidades
                        .inicializarEntidadesDetalle
                        
                        Exit Sub
                    End If
                End If
                
                '.ConceptoCompra = strConceptoCompra 'Trim(rstVale!F4OBSERVA & "")
                
                .CodFormaPago = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                
                With objAyudaFormaPago
                    .inicializarEntidades
                    
                    .Codigo = objAyudaCompra.CodFormaPago
                    
                    .obtenerConfigFormaPago
                End With
                
                .ConceptoCompra = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(rstValeDet!CUENTA & ""), "T")
                
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
                
                .FechaReg = Format(rstVale!F4FECVAL, "Short Date")
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
                        
                        fraProceso.Visible = True
                        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstValeDet)
                        pgbProceso.Value = 0
                        fraProceso.Caption = "Registrando Detalle de R.C. ..."
                        
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .ITEM = dblItem
                            
                            .CtaContableDet = Trim(rstValeDet!CUENTA & "")
                            .AuxiliarDet = Trim(rstValeDet!AUXILIAR & "")
                            
'                            If CBool(ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5CC", "CF5PLA", "F5CODCTA", .CtaContableDet, "T")) Then
                            .CentroCostoDet = txtccosto.Text
'                            End If
                            
                            If .CtaContableDet <> vbNullString Then
                                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T") = vbNullString Then
                                    With objAyudaGasto
                                        .inicializarEntidades
                                        
                                        .Codigo = vbNullString
                                        .Base = "G"
                                        .CuentaContable = objAyudaCompra.CtaContableDet
                                        .Descripcion = Replace(ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", objAyudaCompra.CtaContableDet, "T"), "'", "' & Chr(39) & '", 1)
                                        .TipoGasto = "P"
                                        .Moneda = left(cmbmoneda.Text, 1)
                                        .GrupoFlujo = vbNullString
                                        
                                        If .guardarGasto Then
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                        End If
                                        
                                        .inicializarEntidades
                                    End With
                                End If
                            End If
                            
                            .CodigoGastoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T")
                            
                            .NumeroOrden = Trim(rstValeDet!NROOC & "")
                            .ConceptoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(rstValeDet!f5codpro & ""), "T") 'Replace(left(strConceptoCompra, 100), "'", "' & Chr(39) & '", 1) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableDet, "T")
                            
                            .Cantidad = 1
                            .PrecioUnitario = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .SubTotalDet = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .Afecto = IIf(Trim(rstValeDet!Afecto & "") = "*", True, False)
                            
                            .DebHab = "D"
                            
                            'Acumular
                            .BaseImponible = .BaseImponible + (.SubTotalDet * IIf(.Afecto, 1, 0))
                            .MontoInafecto = .MontoInafecto + (.SubTotalDet * IIf(.Afecto, 0, 1))
                            .TotalIGV = .TotalIGV + (Val(rstValeDet!IGV & "") * IIf(.Afecto, 1, 0))
                            '.Descuento = .Descuento + Val(rstValeDet!DSCTO & "")
                            .TotalFacturado = .TotalFacturado + Val(rstValeDet!TOTAL & "") '((.BaseImponible + .MontoInafecto + .TotalIGV) - .Descuento)
                            
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
                            
                            DoEvents
                            
                            pgbProceso.Value = pgbProceso.Value + 1
                            fraProceso.Caption = "Registrando Detalle de R.C. ... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
                                                        
                            rstValeDet.MoveNext
                        Loop
                            'ACTUALIZAR POSTERIOR A LA GRABACION
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4OCOMPRA = '" & left(.OrdenCompra, 255) & "', "
                            .SQLSelectAlter = .SQLSelectAlter & "F4BASIMP = " & Val(Format(.BaseImponible, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MONINA = " & Val(Format(.MontoInafecto, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & Val(Format((Val(Format(.BaseImponible, "#0.0000")) * (IIf(.TotalIGV = 0, 0, wIgv) / 100)), "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4DCTO = " & Val(Format(.Descuento, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = " & Val(Format(.BaseImponible, "#0.00")) + Val(Format(.MontoInafecto, "#0.00")) + Val(Format((Val(Format(.BaseImponible, "#0.0000")) * (IIf(.TotalIGV = 0, 0, wIgv) / 100)), "#0.00")) - Val(Format(.Descuento, "#0.00")) & " "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                                                        
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End If
                
                If .obtenerCompra Then
                    
                    With objAyudaPagDcto
                        .inicializarEntidades
                        
                        .Correlativo = IIf(objAyudaCompra.Correlativo = 0, -1, objAyudaCompra.Correlativo)
                        .TipoIngreso = "1"
                        .ITEM = 1
                        .NumeroComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2ABREV", "DOCUMENTOS", "F2CODDOC", objAyudaCompra.TipoDocumento, "T")
                        
                        .NumeroComprobante = .NumeroComprobante & IIf(objAyudaCompra.SerieDocumento <> vbNullString, objAyudaCompra.SerieDocumento & "/", vbNullString) & objAyudaCompra.NumeroDocumento
                        
                        .FechaComprobante = objAyudaCompra.FechaDocumento
                        .FechaVencimiento = objAyudaCompra.FechaVencimiento
                        .CodProveedor = objAyudaCompra.CodProveedor
                        .RucProveedor = objAyudaCompra.RucProveedor
                        .CodMoneda = objAyudaCompra.CodMoneda
                        '.TotalFacturado = objAyudaCompra.TotalFacturado
                        '.SaldoFacturado = objAyudaCompra.TotalFacturado
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
                        
                        .NomProveedor = Replace(objAyudaCompra.NomProveedor, "'", "' & Chr(39) & '", 1)
                        .Concepto = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1)  '"INGRESO DE COMPRAS A ALMACEN."
                        .Detalle = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) '"INGRESO DE COMPRAS A ALMACEN."
                        .referencia = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) '"INGRESO DE COMPRAS A ALMACEN."
                        
                        If .verificarExistencia Then
                            .TotalFacturado = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOTAL", "PAG_DCTO", "CORRELA", .Correlativo, "N"))
                            .SaldoFacturado = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "SALDO", "PAG_DCTO", "CORRELA", .Correlativo, "N"))
                            
                            If .SaldoFacturado = .TotalFacturado Then
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                .SaldoFacturado = objAyudaCompra.TotalFacturado
                            ElseIf .SaldoFacturado < .TotalFacturado Then
                                .SaldoFacturado = objAyudaCompra.TotalFacturado - (.TotalFacturado - .SaldoFacturado)
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                            End If
                        Else
                            .TotalFacturado = objAyudaCompra.TotalFacturado
                            .SaldoFacturado = objAyudaCompra.TotalFacturado
                        End If
                        
                        strMensajeErrorPagDcto = vbNullString
                        
                        If .guardarPagDcto(False) Then
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            'If objAyudaCompra.Correlativo = 0 Then
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
                            'End If
                        Else
                            strMensajeErrorPagDcto = "ATENCIÓN: No se genero correctamente la Cuenta por Pagar, por favor vuelva a guardar el Vale."
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
                        .SQLSelectAlter = .SQLSelectAlter & "F4VALESING = '" & Trim(left(.ValeIngreso, 255)) & "' "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                    
                    .obtenerConfigCompra
                    
                    MsgBox "Ingreso exportado a Registro de Compras:" & vbNewLine & _
                            "Mes de Registro: " & .MesMovimiento & vbNewLine & _
                            "Numero de Movimiento: " & .NumeroMovimiento & vbNewLine & vbNewLine & _
                            "Base Imponible = " & Format(.BaseImponible, "#,0.00") & vbNewLine & _
                            "Monto Inafecto = " & Format(.MontoInafecto, "#,0.00") & vbNewLine & _
                            "Total Impuesto = " & Format(.TotalIGV, "#,0.00") & vbNewLine & _
                            "Total Facturado = " & Format(.TotalFacturado, "#,0.00") & _
                            IIf(strMensajeErrorPagDcto <> vbNullString, vbNewLine & vbNewLine & strMensajeErrorPagDcto, vbNullString), _
                            vbInformation + vbOKOnly, App.ProductName
                End If
                
                        sql = vbNullString
                        sql = sql & "DELETE * FROM SCFDOCU1 WHERE "
                        sql = sql & "RUCDOC = '" & objAyudaCompra.RucProveedor & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = " & .NumeroDocumento & ""
                        
                        cnn_dbbancos.Execute sql
                        
                        sql = vbNullString
                        sql = sql & "DELETE * FROM SCFDOCU2 WHERE "
                        sql = sql & "RUCDOC = '" & objAyudaCompra.RucProveedor & "' AND SERIE = '" & .SerieDocumento & "' AND  NRODOC = " & .NumeroDocumento & ""
                
                        cnn_dbbancos.Execute sql
                        
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "INSERT INTO SCFDOCU1 ( RUCDOC, TIPDOC, SERIE, NRODOC, FCHEMI, FCHVTO, TIPMON, TIPADQ, BASIMP, VALIGV, IMPTOT, NOMBRE, TIPDOC_O, SERIE_O, NRODOC_O, FCHDOC_O ) "
                        .SQLSelectAlter = .SQLSelectAlter & "SELECT REGISDOC.F4RUCPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, Format(F4FECHA,'ddmmyyyy') AS Expr1, Format(F4FECVEN,'ddmmyyyy') AS Expr2, IIf(F4MONEDA='D',2,1) AS Expr3, 'A' AS Expr4, REGISDOC.F4BASIMP, REGISDOC.F4IGV, REGISDOC.F4TOTAL, REGISDOC.F4NOMPRV, REGISDOC.TIPODOCREF, REGISDOC.SERDOCREF, REGISDOC.NUMDOCREF, REGISDOC.FECDOCREF "
                        .SQLSelectAlter = .SQLSelectAlter & "FROM REGISDOC "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
     
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        .SQLSelectAlter = vbNullString
                        .SQLSelectAlter = .SQLSelectAlter & "INSERT INTO SCFDOCU2 ( RUCDOC, TIPDOC, SERIE, NRODOC, SECUEN, CUENTA, CENCOS, DESCRI, BASIMP ) "
                        .SQLSelectAlter = .SQLSelectAlter & "SELECT REGISDOC.F4RUCPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, REGISMOV.F3ITEM, REGISMOV.F3CTACON, REGISMOV.F3CENCOS, left(REGISMOV.F3CONCEPTO,40), REGISMOV.F3IMPORTE "
                        .SQLSelectAlter = .SQLSelectAlter & "FROM REGISDOC  INNER JOIN REGISMOV ON (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) AND (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "REGISDOC.F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "REGISDOC.F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
        End If
        
        .inicializarEntidades
    End With
    
    cnDBContaTabla.Close
    
    fraProceso.Visible = False
    
    Exit Sub
    Resume
errExportarRegistroCompraDAO:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description & vbNewLine & vbNewLine & "Registro de Compra NO EXPORTADO correctamente a Tesoreria, intente volver a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
    
    Actualiza_Log "Error en ExportarRegistroCompraDAO - Numero: " & Err.Number & " / Descripcion: " & Err.Description, StrConexDbBancos
    
    Err.Clear
End Sub

Private Sub exportarRegistroCompraADO()
    On Error GoTo errExportarRegistroCompraADO
    
    Dim rstVale As New ADODB.Recordset
    Dim rstValeDet As New ADODB.Recordset
    
    Dim cnDAO As DAO.Database
    
    Dim rstValeDAO As DAO.Recordset
    Dim rstValeDetDAO As DAO.Recordset
    
    Dim strConceptoCompra As String
    Dim dblItem As Double
    Dim dblMontoCancelado As Double
    Dim strMensajeErrorPagDcto As String
    
    abrirCnContaTabla
    
    Set cnDAO = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    With objAyudaProveedor
        .Codigo = Trim(txtproveedor.Text)
        
        .obtenerConfigProveedor
    End With
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .CodigoProveedor = Trim(txtproveedor.Text)
        .CodTipoComprobante = right(cmbtipo.Text, 2)
        .SerieDocumento = Trim(txtserfac.Text)
        .NumeroDocumento = Trim(txtnumfac.Text)
        
        Rem Descargar en el Temporal Cabecera y Detalle de los Vales que coincidan con los datos ingresados anteriormente.
        
        Set rstValeDAO = cnDAO.OpenRecordset("IF4VALES")
        
        rstValeDAO.Index = "IndiceProvComprobante"
        
        rstValeDAO.Seek "=", .TipoVale, .CodigoProveedor, .CodTipoComprobante, .SerieDocumento, .NumeroDocumento, "XC0"
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILVALECAB"
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE FROM TMPUTILVALEDET"
        
        If Not rstValeDAO.NoMatch Then
            Do While (Trim(rstValeDAO!F4TIPOVALE & "") = .TipoVale And _
                        Trim(rstValeDAO!F2CODPROV & "") = .CodigoProveedor And _
                        Trim(rstValeDAO!F4TIPDOC & "") = .CodTipoComprobante And _
                        Trim(rstValeDAO!F4SERDOC & "") = .SerieDocumento And _
                        Trim(rstValeDAO!F4NUMDOC & "") = .NumeroDocumento)
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "INSERT INTO TMPUTILVALECAB "
                
                SqlCad = SqlCad & "IN '" & wrutatemp & "Templus.mdb' "
                
                SqlCad = SqlCad & "SELECT "
                SqlCad = SqlCad & "* "
                SqlCad = SqlCad & "FROM "
                SqlCad = SqlCad & "IF4VALES "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "F2CODALM = '" & Trim(rstValeDAO!f2codalm & "") & "' AND "
                SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstValeDAO!F4NUMVAL & "") & "'"
                
                cnn_dbbancos.Execute SqlCad
                
                Set rstValeDetDAO = cnDAO.OpenRecordset("IF3VALES")

                rstValeDetDAO.Index = "IndiceValeDetalle"

                rstValeDetDAO.Seek "=", Trim(rstValeDAO!f2codalm & ""), Trim(rstValeDAO!F4NUMVAL & "")
                
                If Not rstValeDetDAO.NoMatch Then
                    Do While (Trim(rstValeDetDAO!f2codalm & "") = Trim(rstValeDAO!f2codalm & "") And _
                                Trim(rstValeDetDAO!F4NUMVAL & "") = Trim(rstValeDAO!F4NUMVAL & ""))
                        
                        With objAyudaBien
                            .Codigo = Trim(rstValeDetDAO!f5codpro & "")
                            
                            .obtenerConfigBien
                        End With
                        
                        With objAyudaUM
                            .Codigo = objAyudaBien.CodUM
                            
                            .obtenerConfigUM
                        End With
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO TMPUTILVALEDET("
                        SqlCad = SqlCad & "F2CODALM, F4NUMVAL, TIPO, F4FECVAL, F5CODPRO, F5CODPROORIGINAL, "
                        SqlCad = SqlCad & "F3CANPRO, F3VALVTA, F3IGV, F3TOTITE, F3IGVDOL, F3VALDOL, F3TOTDOL, "
                        SqlCad = SqlCad & "F3PORCENTAJEDSCTO, F3MONTODSCTO, CUENTA, AUXILIAR, AFECTO, "
                        SqlCad = SqlCad & "UM, SUBFAMILIA) "
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!f2codalm & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!F4NUMVAL & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!Tipo & "") & "', "
                        SqlCad = SqlCad & "CVDATE('" & Trim(rstValeDetDAO!F4FECVAL & "") & "'), "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!f5codpro & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDetDAO!F5CODPROORIGINAL & "") & "', "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3CANPRO & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3VALVTA & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3IGV & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3TOTITE & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3IGVDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3VALDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3TOTDOL & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3PORCENTAJEDSCTO & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDetDAO!F3MONTODSCTO & "") & ", "
                        SqlCad = SqlCad & "'" & IIf(objAyudaProveedor.OrigenProveedor = "N", objAyudaBien.CtaContable, objAyudaBien.CtaContableImportacion) & "', "
                        SqlCad = SqlCad & "'" & IIf(objAyudaProveedor.OrigenProveedor = "N", objAyudaBien.Anexo, objAyudaBien.AnexoImportacion) & "', "
                        'SqlCad = SqlCad & "'" & IIf(objAyudaBien.Afecto, "*", vbNullString) & "', "
                        SqlCad = SqlCad & "'" & IIf(Val(rstValeDetDAO!F3IGV & "") = 0, vbNullString, "*") & "', "
                        SqlCad = SqlCad & "'" & objAyudaUM.Abreviatura & "', "
                        SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7DESCON", "SF7NIVEL02", "F7CODCON", objAyudaBien.CodigoSubFamilia, "T") & "'"
                        SqlCad = SqlCad & ")"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
                        
                        rstValeDetDAO.MoveNext
                        
                        If rstValeDetDAO.EOF Then Exit Do
                    Loop
                End If
                
                rstValeDAO.MoveNext
                
                If rstValeDAO.EOF Then Exit Do
            Loop
        End If
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "* "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILVALECAB "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "F4REGCOM DESC, "
        SqlCad = SqlCad & "F4FECVAL DESC"
        
        If rstVale.State = 1 Then rstVale.Close
        
        rstVale.Open SqlCad, cnDBTemp, adOpenDynamic, adLockOptimistic
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.CUENTA, "
        SqlCad = SqlCad & "DET.AUXILIAR, "
        SqlCad = SqlCad & "SUM( VAL(FORMAT(DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL), '#0.0000')) ) AS SUBTOTAL, "
        SqlCad = SqlCad & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3IGV, DET.F3IGVDOL)) AS IGV, "
        SqlCad = SqlCad & "SUM( VAL(FORMAT((DET.F3CANPRO * IIF(CAB.F4MONEDA = 'S', DET.F3VALVTA, DET.F3VALDOL)) * (DET.F3PORCENTAJEDSCTO / 100), '#0.0000')) ) AS DSCTO, "
        SqlCad = SqlCad & "SUM(IIF(CAB.F4MONEDA = 'S', DET.F3TOTITE , DET.F3TOTDOL)) AS TOTAL, "
        SqlCad = SqlCad & "DET.AFECTO, "
        SqlCad = SqlCad & "DET.F4NUMORD AS NROOC "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "TMPUTILVALEDET AS DET "
        SqlCad = SqlCad & "LEFT JOIN TMPUTILVALECAB AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.CUENTA, "
        SqlCad = SqlCad & "DET.AUXILIAR, "
        SqlCad = SqlCad & "DET.AFECTO, "
        SqlCad = SqlCad & "DET.F4NUMORD"
        
        If rstValeDet.State = 1 Then rstValeDet.Close
        
        rstValeDet.Open SqlCad, cnDBTemp, adOpenDynamic, adLockOptimistic
        
        strConceptoCompra = obtenerConceptoCompra
        
        If Not rstVale.EOF Then
            
            With objAyudaCompra
                .inicializarEntidades
                
                .CodProveedor = Trim(rstVale!F2CODPROV & "")
                .TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                .SerieDocumento = Trim(rstVale!F4SERDOC & "")
                .NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                
                'IDENTIFICAR el Registro de Compra en base al Proveedor, Tipo, Numero y Serie de Documento
                .obtenerRegistroCompraProvYdocumento
                
                If .MesMovimiento = vbNullString And .NumeroMovimiento = vbNullString Then
                    If Trim(rstVale!F4REGCOM & "") <> vbNullString Then
                        .MesMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), 1, InStr(1, Trim(rstVale!F4REGCOM & ""), "-") - 1)
                        .NumeroMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), InStr(1, Trim(rstVale!F4REGCOM & ""), "-") + 1)
                    Else
                        .MesMovimiento = Year(CDate(dtpFechaDoc.Value)) & Format(Month(CDate(dtpFechaDoc.Value)), "00")
                        
                        Rem SK ADD: ADICIONAR LINEAS DE EVALUACION:
                        'SI EL MES DEL DOCUMENTO DE LA COMPRA SE ENCUENTRA CERRADO, PASARLO AUTOMATICAMENTE AL MES SIGUIENTE.
                        If .verificarCierreCompra Then
                            MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                    "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                    "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                            
                            .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                                IIf(Month(CDate(dtpFechaDoc.Value)) = 12, "01", Format(Month(CDate(dtpFechaDoc.Value)) + 1, "00"))
                            
                            If .verificarCierreCompra Then
                                MsgBox "El Periodo = " & .MesMovimiento & " se encuentra CERRADO por Contabilidad, " & vbNewLine & _
                                        "el Comprobante sera anexado al siguiente periodo." & vbNewLine & vbNewLine & _
                                        "Presione [OK] para continuar.", vbInformation + vbOKOnly, App.ProductName
                                                            
                                .MesMovimiento = IIf(Month(CDate(dtpFechaDoc.Value)) = 12, Year(CDate(dtpFechaDoc.Value)) + 1, Year(CDate(dtpFechaDoc.Value))) & _
                                                    IIf(Month(CDate(dtpFechaDoc.Value)) >= 11, Format(CDate(dtpFechaDoc.Value) - 10, "00"), Format(Month(CDate(dtpFechaDoc.Value)) + 2, "00"))
                                                    
                            End If
                        End If
                        '---------------------------------------------------------------------
                        .TipoRegistro = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODCOMPRAREG", "DOCUMENTOS", "F2CODDOC", Trim(rstVale!F4TIPDOC & ""), "T")
                        
                        .NumeroMovimiento = vbNullString
                    End If
                End If
                
                '.CodProveedor = Trim(rstVale!F2CODPROV & "")
                .NomProveedor = Replace(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T"), "'", "' & Chr(39) & '", 1)
                .DireccionProveedor = Replace(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T"), "'", "' & Chr(39) & '", 1)
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
                    
                '.TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                '.SerieDocumento = Trim(rstVale!F4SERDOC & "")
                '.NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                
                .CodigoCategoria = 1
                
                '.FechaRegistro = Format(Date, "Short Date")
                .FechaDocumento = Format(Trim(rstVale!F4FECULT & ""), "Short Date")
                
                
                If Val(Format(.FechaDocumento, "YYYYMM")) = Val(.MesMovimiento) Then
                    .FechaRegistro = .FechaDocumento
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) < Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 0, 1)
                ElseIf Val(Format(.FechaDocumento, "YYYYMM")) > Val(.MesMovimiento) Then
                    .FechaRegistro = DateSerial(Val(Mid(.MesMovimiento, 1, Len(.MesMovimiento) - 2)), Val(right(.MesMovimiento, 2)) + 1, 0)
                End If
                
                
                .CodMoneda = Trim(rstVale!F4MONEDA & "")
                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaDocumento, "F"))
                
                If .TipoCambio = 0 Then
                    obtiene_tipodecambio (.FechaDocumento)
                    .TipoCambio = TCVenta
                    If .TipoCambio = 0 Then
                        MsgBox "Imposible generar Registro de Compra, Tipo de Cambio del " & .FechaDocumento & " no registrado, verifique y vuelva a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
                        
                        .inicializarEntidades
                        .inicializarEntidadesDetalle
                        
                        Exit Sub
                    End If
                End If
                
                '.ConceptoCompra = strConceptoCompra 'Trim(rstVale!F4OBSERVA & "")
                
                .CodFormaPago = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                
                With objAyudaFormaPago
                    .inicializarEntidades
                    
                    .Codigo = objAyudaCompra.CodFormaPago
                    
                    .obtenerConfigFormaPago
                End With
                
                .ConceptoCompra = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(rstValeDet!CUENTA & ""), "T")
                
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
                        
                        fraProceso.Visible = True
                        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstValeDet)
                        pgbProceso.Value = 0
                        fraProceso.Caption = "Registrando Detalle de R.C. ..."
                        
                        Do While Not rstValeDet.EOF
                            .inicializarEntidadesDetalle
                            
                            dblItem = dblItem + 1
                            
                            .ITEM = dblItem
                            
                            .CtaContableDet = Trim(rstValeDet!CUENTA & "")
                            .AuxiliarDet = Trim(rstValeDet!AUXILIAR & "")
                            
                            If CBool(ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5CC", "CF5PLA", "F5CODCTA", .CtaContableDet, "T")) Then
                                .CentroCostoDet = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "CentroCostoAutomaticoRegCompraDesdeLogistica", "l")
                            End If
                            
                            If .CtaContableDet <> vbNullString Then
                                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T") = vbNullString Then
                                    With objAyudaGasto
                                        .inicializarEntidades
                                        
                                        .Codigo = vbNullString
                                        .Base = "G"
                                        .CuentaContable = objAyudaCompra.CtaContableDet
                                        .Descripcion = Replace(ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", objAyudaCompra.CtaContableDet, "T"), "'", "' & Chr(39) & '", 1)
                                        .TipoGasto = "P"
                                        .Moneda = left(cmbmoneda.Text, 1)
                                        .GrupoFlujo = vbNullString
                                        
                                        If .guardarGasto Then
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                        End If
                                        
                                        .inicializarEntidades
                                    End With
                                End If
                            End If
                            
                            .CodigoGastoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T")
                            
                            .NumeroOrden = Trim(rstValeDet!NROOC & "")
                            .ConceptoDet = Replace(left(strConceptoCompra, 100), "'", "' & Chr(39) & '", 1) 'ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableDet, "T")
                            
                            .Cantidad = 1
                            .PrecioUnitario = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .SubTotalDet = Val(rstValeDet!SUBTOTAL & "") - Val(rstValeDet!DSCTO & "")
                            .Afecto = IIf(Trim(rstValeDet!Afecto & "") = "*", True, False)
                            
                            .DebHab = "D"
                            
                            'Acumular
                            .BaseImponible = .BaseImponible + (.SubTotalDet * IIf(.Afecto, 1, 0))
                            .MontoInafecto = .MontoInafecto + (.SubTotalDet * IIf(.Afecto, 0, 1))
                            .TotalIGV = .TotalIGV + (Val(rstValeDet!IGV & "") * IIf(.Afecto, 1, 0))
                            '.Descuento = .Descuento + Val(rstValeDet!DSCTO & "")
                            .TotalFacturado = .TotalFacturado + Val(rstValeDet!TOTAL & "") '((.BaseImponible + .MontoInafecto + .TotalIGV) - .Descuento)
                            
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
                            
                            DoEvents
                            
                            pgbProceso.Value = pgbProceso.Value + 1
                            fraProceso.Caption = "Registrando Detalle de R.C. ... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
                                                        
                            rstValeDet.MoveNext
                        Loop
                            'ACTUALIZAR POSTERIOR A LA GRABACION
                            .SQLSelectAlter = vbNullString
                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                            .SQLSelectAlter = .SQLSelectAlter & "F4OCOMPRA = '" & left(.OrdenCompra, 255) & "', "
                            .SQLSelectAlter = .SQLSelectAlter & "F4BASIMP = " & Val(Format(.BaseImponible, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MONINA = " & Val(Format(.MontoInafecto, "#0.00")) & ", "
                            '.SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & Val(Format(.TotalIGV, "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & Val(Format((Val(Format(.BaseImponible, "#0.00")) * (IIf(.TotalIGV = 0, 0, wIgv) / 100)), "#0.00")) & ", "
                            .SQLSelectAlter = .SQLSelectAlter & "F4DCTO = " & Val(Format(.Descuento, "#0.00")) & ", "
                            '.SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = " & Val(Format(.TotalFacturado, "#0.00")) & " "
                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = " & Val(Format(((Val(Format(.BaseImponible, "#0.00")) * (1 + IIf(.TotalIGV = 0, 0, wIgv) / 100)) + .MontoInafecto) - Val(Format(.Descuento, "#0.00")), "#0.00")) & " "
                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                                                        
                            cnn_dbbancos.Execute .SQLSelectAlter
                            
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End If
                
                If .obtenerCompra Then
                    
                    With objAyudaPagDcto
                        .inicializarEntidades
                        
                        .Correlativo = IIf(objAyudaCompra.Correlativo = 0, -1, objAyudaCompra.Correlativo)
                        .TipoIngreso = "1"
                        .ITEM = 1
                        .NumeroComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2ABREV", "DOCUMENTOS", "F2CODDOC", objAyudaCompra.TipoDocumento, "T")
                        
                        .NumeroComprobante = .NumeroComprobante & IIf(objAyudaCompra.SerieDocumento <> vbNullString, objAyudaCompra.SerieDocumento & "/", vbNullString) & objAyudaCompra.NumeroDocumento
                        
                        .FechaComprobante = objAyudaCompra.FechaDocumento
                        .FechaVencimiento = objAyudaCompra.FechaVencimiento
                        .CodProveedor = objAyudaCompra.CodProveedor
                        .RucProveedor = objAyudaCompra.RucProveedor
                        .CodMoneda = objAyudaCompra.CodMoneda
                        '.TotalFacturado = objAyudaCompra.TotalFacturado
                        '.SaldoFacturado = objAyudaCompra.TotalFacturado
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
                        
                        .NomProveedor = Replace(objAyudaCompra.NomProveedor, "'", "' & Chr(39) & '", 1)
                        .Concepto = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1)  '"INGRESO DE COMPRAS A ALMACEN."
                        .Detalle = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) '"INGRESO DE COMPRAS A ALMACEN."
                        .referencia = Replace(left(strConceptoCompra, 255), "'", "' & Chr(39) & '", 1) '"INGRESO DE COMPRAS A ALMACEN."
                        
                        If .verificarExistencia Then
                            .TotalFacturado = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TOTAL", "PAG_DCTO", "CORRELA", .Correlativo, "N"))
                            .SaldoFacturado = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "SALDO", "PAG_DCTO", "CORRELA", .Correlativo, "N"))
                            
                            If .SaldoFacturado = .TotalFacturado Then
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                .SaldoFacturado = objAyudaCompra.TotalFacturado
                            ElseIf .SaldoFacturado < .TotalFacturado Then
                                .SaldoFacturado = objAyudaCompra.TotalFacturado - (.TotalFacturado - .SaldoFacturado)
                                .TotalFacturado = objAyudaCompra.TotalFacturado
                            End If
                        Else
                            .TotalFacturado = objAyudaCompra.TotalFacturado
                            .SaldoFacturado = objAyudaCompra.TotalFacturado
                        End If
                        
                        strMensajeErrorPagDcto = vbNullString
                        
                        If .guardarPagDcto(False) Then
                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            
                            'If objAyudaCompra.Correlativo = 0 Then
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
                            'End If
                        Else
                            strMensajeErrorPagDcto = "ATENCIÓN: No se genero correctamente la Cuenta por Pagar, por favor vuelva a guardar el Vale."
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
                        .SQLSelectAlter = .SQLSelectAlter & "F4VALESING = '" & Trim(left(.ValeIngreso, 255)) & "' "
                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                        
                        cnn_dbbancos.Execute .SQLSelectAlter
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                    
                    .obtenerConfigCompra
                    
                    MsgBox "Ingreso exportado a Registro de Compras:" & vbNewLine & _
                            "Mes de Registro: " & .MesMovimiento & vbNewLine & _
                            "Numero de Movimiento: " & .NumeroMovimiento & vbNewLine & vbNewLine & _
                            "Base Imponible = " & Format(.BaseImponible, "#,0.00") & vbNewLine & _
                            "Monto Inafecto = " & Format(.MontoInafecto, "#,0.00") & vbNewLine & _
                            "Total Impuesto = " & Format(.TotalIGV, "#,0.00") & vbNewLine & _
                            "Total Facturado = " & Format(.TotalFacturado, "#,0.00") & _
                            IIf(strMensajeErrorPagDcto <> vbNullString, vbNewLine & vbNewLine & strMensajeErrorPagDcto, vbNullString), _
                            vbInformation + vbOKOnly, App.ProductName
                End If
                
                .inicializarEntidades
                .inicializarEntidadesDetalle
            End With
        End If
        
        .inicializarEntidades
    End With
    
    cnDBContaTabla.Close
    
    fraProceso.Visible = False
    
    Exit Sub
errExportarRegistroCompraADO:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description & vbNewLine & vbNewLine & "Registro de Compra NO EXPORTADO correctamente a Tesoreria, intente volver a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
    
    Actualiza_Log "Error en ExportarRegistroCompraADO - Numero: " & Err.Number & " / Descripcion: " & Err.Description, StrConexDbBancos
    
    Err.Clear
End Sub

Private Function obtenerConceptoCompra() As String
    On Error GoTo errObtenerConceptoCompra
    
    Dim rstConceptoCompra As New ADODB.Recordset
    
    If rstConceptoCompra.State = 1 Then rstConceptoCompra.Close
    
    rstConceptoCompra.Open "SELECT SUBFAMILIA, UM, SUM(F3CANPRO) AS CANTIDAD FROM TMPUTILVALEDET GROUP BY SUBFAMILIA, UM", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstConceptoCompra.EOF Then
        Do While Not rstConceptoCompra.EOF
            If obtenerConceptoCompra = vbNullString Then
                obtenerConceptoCompra = Trim(rstConceptoCompra!SUBFAMILIA & "") & " ( " & Val(rstConceptoCompra!Cantidad & "") & " " & Trim(rstConceptoCompra!um & "") & " )"
            Else
                obtenerConceptoCompra = obtenerConceptoCompra & ", " & Trim(rstConceptoCompra!SUBFAMILIA & "") & " ( " & Val(rstConceptoCompra!Cantidad & "") & " " & Trim(rstConceptoCompra!um & "") & " )"
            End If
            
            rstConceptoCompra.MoveNext
        Loop
    End If
    
    Exit Function
errObtenerConceptoCompra:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description & vbNewLine & vbNewLine & _
            "Registro de Compra NO EXPORTADO correctamente a Tesoreria, intente volver a guardar el Vale.", vbInformation + vbOKOnly, App.ProductName
    
    obtenerConceptoCompra = vbNullString
    
    Err.Clear
End Function

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
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    cerrarValeSql
                End If
                
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
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    abrirValeSql
                End If
                
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


''--------------------------------------------------------------------------------------------------------------
''------------ SQL ---------------------------------------------------------------------------------------------
''--------------------------------------------------------------------------------------------------------------

Private Sub consultarValeSql()
    On Error GoTo errConsultarValeSql
    
    Set objSqlVale = New SqlClsVale
    
    limpiarCajas
    
    dxDBGrid1.Dataset.Close
    
    With objSqlVale
        .inicializarEntidades
        
        .CodigoAlmacen = strCodAlmacen
        .NumeroVale = strNumeroVale
        
        If .obtenerVale Then
            txtnumero.Text = .NumeroVale
                lblNumeroValeExterno.Caption = .NumeroValeExterno
            
            cmbalmacen.ListIndex = ModUtilitario.seleccionarItem(cmbalmacen, .CodigoAlmacen, "DER", 2): cmbalmacen.Enabled = False
                txtAlmacen.Text = .CodigoAlmacen
            cmbconcepto.ListIndex = ModUtilitario.seleccionarItem(cmbconcepto, .CodigoOrigen, "DER", 3)
                txtconcepto.Text = .CodigoOrigen
            
            cmbTipoAuxiliar.ListIndex = ModUtilitario.seleccionarItem(cmbTipoAuxiliar, .TipoPersona, "DER", 1)
                txtproveedor.Text = .CodigoProveedor
                
                Select Case .TipoPersona
                    Case "C"
                        txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2NOMCLI", "MAESTROS.EF2CLIENTES", "F2CODCLI", .CodigoProveedor, "T")
                    Case "P"
                        txtnomprov.Text = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2NOMPROV", "MAESTROS.EF2PROVEEDORES", "F2CODPROV", .CodigoProveedor, "T")
                End Select
                
            txtccosto.Text = .CentroCosto
                pnlccosto.Caption = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F3DESCRIP", "MAESTROS.CENTROS", "F3COSTO", .CentroCosto, "T")
            txtserie.Text = .SerieGuia: txtnumdoc.Text = .NumeroGuia
            
            cmbtipo.ListIndex = ModUtilitario.seleccionarItem(cmbtipo, .CodTipoComprobante, "DER", 2)
            txtserfac.Text = .SerieDocumento: txtnumfac.Text = .NumeroDocumento
            dtpFechaDoc.Value = IIf(.FechaUltima <> vbNullString, .FechaUltima, Null)
            
            'abrirCnDBMilano
            
            fraOrdenProduccion.Enabled = False
            txtIDOrdenProduccion.Text = .OrdenTrabajo
                If .OrdenTrabajo <> vbNullString Then
                    lblIdCategoriaTipo.Caption = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                    'cmbCategoriaTipo.ListIndex = ModUtilitario.seleccionarItem(cmbCategoriaTipo, ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N"), "DER", 10)
                    cmbCategoriaTipo.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", Trim(lblIdCategoriaTipo.Caption), "T")
                    txtNroOrdenProduccion.Text = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", .OrdenTrabajo, "N")
                End If
            
            abofecha.Value = .Fecha
            cmbmoneda.ListIndex = ModUtilitario.seleccionarItem(cmbmoneda, .CodigoMoneda, "IZQ", 1)
            txttc.Text = Format(.TipoCambio, "#.000")
            
            txtOcompra.Text = .NumeroOrdenCompra
            txtobserva.Text = .observaciones
            
            If .RegistroCompra <> vbNullString Then
                pnlRegistroCompra.Visible = True
                txtRegistroCompra.Text = .RegistroCompra
            End If
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPVALEINGRESO", vbNullString, vbNullString, vbNullString, "TRIM(CODPROD & '') <> ''")) > 0 Then
                renumerarItemVale
            End If
            
            listarGrillaVale
            
            If dxDBGrid1.Dataset.RecordCount = 0 Then
                dxDBGrid1.Dataset.Close
                
                adicionarItemVale
            Else
                recalcularItems
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
            
            If ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODTAREA", "MAESTROS.EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0015'") = "0015" Then
                SSActiveToolBars1.Tools.ITEM("CerrarVale").Enabled = True
            End If
            
            If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
            Else
                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
            End If
            
            bolObviarCierre = False
            
            Rem SK ADD:
            If Trim(txtconcepto.Text) = "XCS" Then
                SSFrame1.Enabled = False
                dxDBGrid1.Enabled = False
                
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
                SSActiveToolBars1.Tools("ID_OC").Enabled = False
                
                chkVerObservaciones.Value = vbChecked
                
                chkVerObservaciones_Click
            ElseIf .VB1 Then
                SSFrame1.Enabled = False
                dxDBGrid1.Enabled = False
                
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
                SSActiveToolBars1.Tools("ID_OC").Enabled = False
                
                chkVerObservaciones.Value = vbChecked
                
                chkVerObservaciones_Click
            End If
        Else
            listarGrillaVale
            
            adicionarItemVale
        End If
    End With
    
    Set objSqlVale = Nothing
    
    Exit Sub
errConsultarValeSql:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: ConsultarValeSql"
    
    Err.Clear
End Sub

Private Sub guardarValeSql()
    On Error GoTo errGuardarValeSql
    
    Dim rstTemporalGuardarValeISql As New ADODB.Recordset
    Dim dblItem As Double
    
    Set objSqlVale = New SqlClsVale
    
    'dxDBGrid1.Dataset.Close
    
    With objSqlVale
        .inicializarEntidades
        
        .TipoVale = "I"
        .NumeroVale = strNumeroVale  'Trim(txtNumero.Text)
        .NumeroValeExterno = Trim(lblNumeroValeExterno.Caption)
        
        .CodigoAlmacen = strCodAlmacen 'Trim(txtAlmacen.Text)
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
            .FechaUltima = Format(dtpFechaDoc.Value, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = Trim(txtIDOrdenProduccion.Text)
        
        .Fecha = Format(abofecha.Value, "Short Date")
        .CodigoMoneda = left(cmbmoneda.Text, 1)
        .TipoCambio = Val(txttc.Text)
        
        .NumeroOrdenCompra = Trim(txtOcompra.Text)
        .observaciones = Trim(txtobserva.Text)
        
        .RegistroCompra = Trim(txtRegistroCompra.Text)
        
        .ExportarVale = CBool(chkExportarVale.Value)
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        .FecMod = Format(Date, "Short Date")
        .UsuMod = wusuario
        
        If .guardarVale Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            If rstTemporalGuardarValeISql.State = 1 Then rstTemporalGuardarValeISql.Close
            
            rstTemporalGuardarValeISql.Open "SELECT * FROM TMPVALEINGRESO WHERE TRIM(CODPROD & '') <> '' AND CANTIDAD > 0 ORDER BY ITEM, DESCRIPCION", cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTemporalGuardarValeISql.EOF Then
                rstTemporalGuardarValeISql.MoveFirst
                
                With objSqlAyudaOrigen
                    .inicializarEntidades
                    
                    .Codigo = objSqlVale.CodigoOrigen
                    
                    .obtenerConfigOrigen
                End With
                
                dblItem = 0
                
                Do While Not rstTemporalGuardarValeISql.EOF
                    .inicializarEntidadesDetalle
                    
                    dblItem = dblItem + 1
                    
                    .ITEM = dblItem
                    
                    .CodigoProducto = Trim(rstTemporalGuardarValeISql!codprod & "")
                    .CodigoProductoOriginal = Trim(rstTemporalGuardarValeISql!CODPRODORIGINAL & "")
                    .Cantidad = Val(rstTemporalGuardarValeISql!Cantidad & "")
                    .CantidadMaxima = Val(rstTemporalGuardarValeISql!CANTIDADMAX & "")
                    
                    .ValorVenta = 0
                    .IGV = 0
                    .TOTAL = 0
                    
                    .IgvDol = 0
                    .TotalDol = 0
                    
                    If objAyudaOrigen.RegistrarCosto Then
                        .PorcentajeDscto = Val(Format(Val(rstTemporalGuardarValeISql!PORDESC & ""), "#0.00"))
                        
                        Select Case .CodigoMoneda
                            Case "S"
                                .ValorVenta = Val(Format(Val(rstTemporalGuardarValeISql!COSTOUNI & ""), "#0.0000"))
                                .IGV = Val(Format(Val(rstTemporalGuardarValeISql!IGV & ""), "#0.0000"))
                                .TOTAL = Val(Format(Val(rstTemporalGuardarValeISql!TOTAL & ""), "#0.00"))
                                
                                .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#.0000"))
                                .IgvDol = Val(Format(.IGV / .TipoCambio, "#0.0000"))
                                .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                                
                                .MontoDscto = IIf(.PorcentajeDscto > 0, Val(Format((.Cantidad * .ValorVenta) * (.PorcentajeDscto / 100), "#.0000")), 0)
                            Case Else
                                .ValorVentaDol = Val(Format(Val(rstTemporalGuardarValeISql!COSTOUNI & ""), "#0.0000"))
                                .IgvDol = Val(Format(Val(rstTemporalGuardarValeISql!IGV & ""), "#0.0000"))
                                .TotalDol = Val(Format(Val(rstTemporalGuardarValeISql!TOTAL & ""), "#0.00"))
                                
                                .ValorVenta = Val(Format(.ValorVentaDol * .TipoCambio, "#.0000"))
                                .IGV = Val(Format(.IgvDol * .TipoCambio, "#0.0000"))
                                .TOTAL = Val(Format(.TotalDol * .TipoCambio, "#0.00"))
                                
                                .MontoDscto = IIf(.PorcentajeDscto > 0, Val(Format((.Cantidad * .ValorVentaDol) * (.PorcentajeDscto / 100), "#.0000")), 0)
                        End Select
                    Else
                        '.ValorVenta = .calcularCostoPromedioV2
                        
                        '.ValorVenta = .calcularCostoPromedioV3conDAO
                        
                        .IGV = 0
                        .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                        
                        .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                        .IgvDol = 0
                        .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    End If
                    
                    .NumeroOrdenCompra = Trim(rstTemporalGuardarValeISql!F4NUMORD & "")
                    .Requerimiento = Trim(rstTemporalGuardarValeISql!COD_SOLICITUD & "")
                    
                    .ObservacionesPorItem = Trim(rstTemporalGuardarValeISql!observaciones & "")
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    rstTemporalGuardarValeISql.MoveNext
                Loop
            End If
            
'            'Exportar el Vale
'            If ModMilano.exportarValeAserverSQLv2(.CodigoAlmacen, .NumeroVale, lblNumeroValeExterno, fraProceso, pgbProceso) Then
'                'MsgBox "ID Ingreso en Sistema Externo: " & lblNumeroValeExterno.Caption & ".", vbInformation + vbOKOnly, App.ProductName
'
'                If .CodigoOrigen = "XOP" And .OrdenTrabajo <> vbNullString Then
'                    verificarProductoDevueltoOP .CodigoAlmacen, .NumeroVale
'                End If
'            End If
'
'            Select Case .CodigoOrigen
'                Case "XC0"
'                    If Trim(txtOcompra.Text) <> vbNullString Then
'                        verificarAtencionOrden .CodigoAlmacen, .NumeroVale
'                    End If
'
'                    If cmbtipo.ListIndex <> -1 And Trim(txtnumfac.Text) <> vbNullString Then
'                        With objSqlAyudaComprobante
'                            .Codigo = right(cmbtipo.Text, 2)
'
'                            .obtenerConfigComprobante
'
'                            If .EsOficial Then
'                                'Generar Registro de Compra
'
'                                If .CodCompraRegistro <> vbNullString Then
'                                    'exportarRegistroCompra
'
'                                    exportarRegistroCompraDAO
'                                Else
'                                    MsgBox "Imposible Generar el Registro de Compra, el Tipo de Documento seleccionado no cuenta con la Configuracion 'Tipo de Registro de Compra', verifique en el Sistema de Bancos: [Mantenimientos] >> [Documentos] y vuelva a GUARDAR el Vale.", vbInformation + vbOKOnly, App.ProductName
'                                End If
'                            End If
'                        End With
'                    End If
'            End Select
'
'
'            strCodAlmacen = .CodigoAlmacen
'            strNumeroVale = .NumeroVale
'
'            consultarVale
'
'            MsgBox "Se ha Actualizado el Vale de Ingreso " & .NumeroVale & "." & vbNewLine & _
'                    "ID Salida en Sistema Externo: " & lblNumeroValeExterno.Caption & ".", vbInformation + vbOKOnly, App.ProductName
'        Else
'            listarGrillaVale
        End If
    End With
    
    Set objSqlVale = Nothing
    
    Exit Sub
errGuardarValeSql:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: GuardarValeSql"
    
    Err.Clear
End Sub

Private Sub eliminarValeSql()
    On Error GoTo errEliminarValeSql
    
    Set objSqlVale = New SqlClsVale
    
    With objSqlVale
        .CodigoAlmacen = strCodAlmacen 'Trim(txtAlmacen.Text)
        .NumeroVale = strNumeroVale 'Trim(txtNumero.Text)
        
'        If Not objSqlVale.obtenerVale Then
'            MsgBox "Registro no existente.", vbInformation + vbOKOnly, App.ProductName
'
'            Exit Sub
'        End If
'
''        'Restricción de Anulación de Vale
''        If Month(CDate(abofecha.value)) < Month(Date) Then
''            MsgBox "Imposible eliminar Vale. Fuera del Periodo Actual." & vbNewLine & vbNewLine & _
''                    "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
''
''            Exit Sub
''        End If
'
'        'Validacion del Punto (PC) que elimina el Vale
'        ModMilano.abrirCnDBMilano
'
'        If Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")) = vbNullString Then
'            MsgBox "Su computador no esta registrado y/o habilitado. Consulte con su" & vbNewLine & vbNewLine & _
'                    "administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'            Exit Sub
'        End If
'
'        If Val(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "CORRELATIVO", "CORRELATIVOTABLA", "IDPUNTOVENTA", Trim(ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDPUNTOVENTA", "PUNTOVENTA", "NOMBREPC", modgeneral.ComputerName, "T", "AND ELIMINADO = 0")), "T", _
'                                                                        "AND TABLA = 'INGRESO'")) = 0 Then
'
'            MsgBox "El Punto de Venta no cuenta con correlativo habilitado de " & strTablaSQL & "." & vbNewLine & vbNewLine & _
'                    "Consulte con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName & " - " & wnomcia
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea eliminar el Vale con No. " & .NumeroVale & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'            If Val(lblNumeroValeExterno.Caption) > 0 Then
'                If Not ModMilano.anularValeExterno("I", lblNumeroValeExterno.Caption, ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2CODALMEXTERNO", "MAESTROS.EF2ALMACENES", "F2CODALM", Trim(txtalmacen.Text), "T"), fraProceso, pgbProceso) Then
'                    Me.MousePointer = vbDefault
'
'                    Exit Sub
'                End If
'            End If
'
            If .eliminarVale Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                
'                If .CodigoOrigen = "XC0" And .NumeroOrdenCompra <> vbNullString Then
'                    verificarAtencionOrdenSql .CodigoAlmacen, .NumeroVale, True
'                End If
'
'                strCodAlmacen = .CodigoAlmacen
'                strNumeroVale = .NumeroVale
'
'                consultarVale
'
'                MsgBox "Registro eliminado.", vbInformation + vbOKOnly, App.ProductName
            End If
'        End If
    End With
    
    Set objSqlVale = Nothing
    
    Exit Sub
errEliminarValeSql:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Vale Ingreso: EliminarValeSql"
    
    Err.Clear
End Sub

Private Sub verificarAtencionOrdenSql(ByVal strCodAlmacen As String, _
                                    ByVal strNumeroVale As String, _
                                    Optional ByVal bolValeEliminado As Boolean)
    
    Dim rstValeDet As New ADODB.Recordset
    Dim dblCantidadOP As Double
    
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    If Not bolValeEliminado Then
        rstValeDet.Open "SELECT F4NUMORD FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & strCodAlmacen & "' AND F4NUMVAL = '" & strNumeroVale & "' GROUP BY F4NUMORD", cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
        rstValeDet.Open "SELECT F4NUMORD FROM TMPVALEINGRESO GROUP BY F4NUMORD", cnDBTemp, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not rstValeDet.EOF Then
        rstValeDet.MoveFirst
        
        Do While Not rstValeDet.EOF
            With objSqlAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = "OC"
                .NumeroOrden = Trim(rstValeDet!F4NUMORD & "")
                
                If .obtenerOrden Then
                    If .Estado <> 7 And .Estado <> 8 Then
                        .atencionOrden
                        
                        abrirCnnDbBancos
                        
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "UPDATE "
                            SqlCad = SqlCad & "IF4ORDEN "
                        SqlCad = SqlCad & "SET "
                            SqlCad = SqlCad & "F4ESTADO = " & .Estado & " "
                        SqlCad = SqlCad & "WHERE "
                            SqlCad = SqlCad & "F4LOCAL = '" & .TipoOrden & "' AND "
                            SqlCad = SqlCad & "F4NUMORD = '" & .NumeroOrden & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        
                        Actualiza_Log SqlCad, StrConexDbBancos
                        
                        abrirCnnDbBancos
                    End If
                End If
            End With
            
            rstValeDet.MoveNext
        Loop
    End If
    
    If rstValeDet.State = 1 Then rstValeDet.Close
    
    Set rstValeDet = Nothing
End Sub

Private Sub cerrarValeSql()
    Dim strFechaCorteInicialDeValesParaCP As String
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    strFechaCorteInicialDeValesParaCP = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")
    
    With objAyudaVale
        .inicializarEntidades
        
        .CodigoAlmacen = strCodAlmacen
        .NumeroVale = strNumeroVale
        
'        .obtenerConfigVale
'
'        If Val(Year(CDate(.Fecha)) & Format(Month(CDate(.Fecha)), "00")) > Val(Format(CDate(strFechaCorteInicialDeValesParaCP), "yyyymm")) Then
'            .inicializarEntidadesAdicionales
'
'            intAnnoCorte = Val(Year(CDate(.Fecha))) - IIf(Val(Month(CDate(.Fecha))) > 1, 0, 1)
'            intMesCorte = IIf(Val(Month(CDate(.Fecha))) > 1, Val(Month(CDate(.Fecha))) - 1, 12)
'
'            .inicializarEntidades
'
'            .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
'            .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
'
'            If Not .verificarCierreVale Then
'                MsgBox "Imposible cerrar el Vale; ya que el anterior Periodo aun se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                bolObviarCierre = True
'
'                SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
'
'                If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
'                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
'                Else
'                    SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
'                End If
'
'                bolObviarCierre = False
'
'                Exit Sub
'            End If
'        End If
'
'        .CodigoAlmacen = strCodAlmacen
'        .NumeroVale = strNumeroVale
'
'        .obtenerConfigVale
'
'        If .VB1 Then
'            MsgBox "Vale ya se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'            bolObviarCierre = True
'
'            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
'
'            If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
'            Else
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
'            End If
'
'            bolObviarCierre = False
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea cerrar el Vale?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            .VB1 = True
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            If .cerrarVale Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                
'                MsgBox "Vale cerrado correctamente.", vbInformation + vbOKOnly, App.ProductName
            End If
            
'            strCodAlmacen = .CodigoAlmacen
'            strNumeroVale = .NumeroVale
'
'            consultarVale
'        Else
'            bolObviarCierre = True
'
'            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
'
'            bolObviarCierre = False
'        End If
    End With
End Sub

Private Sub abrirValeSql()
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    With objAyudaVale
        .inicializarEntidades
        .inicializarEntidadesAdicionales
        
        .CodigoAlmacen = strCodAlmacen
        .NumeroVale = strNumeroVale
        
'        .obtenerConfigVale
'
'        intAnnoCorte = Val(Year(CDate(.Fecha))) + IIf(Val(Month(CDate(.Fecha))) < 12, 0, 1)
'        intMesCorte = IIf(Val(Month(CDate(.Fecha))) < 12, Val(Month(CDate(.Fecha))) + 1, 1)
'
'        .inicializarEntidades
'
'        .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
'        .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0)
'
'        If .verificarCierreVale Then
'            MsgBox "Imposible abrir el Vale; ya que el Periodo posterior se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'            bolObviarCierre = True
'
'            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
'
'            If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
'            Else
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
'            End If
'
'            bolObviarCierre = False
'
'            Exit Sub
'        End If
'
'        .inicializarEntidades
'        .inicializarEntidadesAdicionales
'
'        .CodigoAlmacen = strCodAlmacen
'        .NumeroVale = strNumeroVale
'
'        .obtenerConfigVale
'
'        If Not .VB1 Then
'            MsgBox "Vale se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'            bolObviarCierre = True
'
'            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssUnchecked
'
'            If Not CBool(SSActiveToolBars1.Tools.ITEM("CerrarVale").State) Then
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Cerrar Vale"
'            Else
'                SSActiveToolBars1.Tools("CerrarVale").ChangeAll ssChangeAllName, "Abrir Vale"
'            End If
'
'            bolObviarCierre = False
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea abrir el Vale?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            .VB1 = False
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            If .cerrarVale Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                
'                MsgBox "Vale abierto correctamente.", vbInformation + vbOKOnly, App.ProductName
            End If
            
'            strCodAlmacen = .CodigoAlmacen
'            strNumeroVale = .NumeroVale
'
'            consultarVale
'        Else
'            bolObviarCierre = True
'
'            SSActiveToolBars1.Tools.ITEM("CerrarVale").State = ssChecked
'
'            bolObviarCierre = False
'        End If
    End With
End Sub
