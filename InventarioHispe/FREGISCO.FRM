VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmregiscom 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Compras"
   ClientHeight    =   7365
   ClientLeft      =   1635
   ClientTop       =   960
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7365
   ScaleMode       =   0  'User
   ScaleWidth      =   10125
   Begin Threed.SSPanel SSPanel4 
      Height          =   510
      Left            =   90
      TabIndex        =   82
      Top             =   0
      Width           =   9960
      _Version        =   65536
      _ExtentX        =   17568
      _ExtentY        =   900
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin Crystal.CrystalReport Report1 
         Left            =   8985
         Top             =   45
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin Threed.SSCommand BtnAddO 
         Height          =   375
         Left            =   90
         TabIndex        =   83
         ToolTipText     =   "Nuevo"
         Top             =   70
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0000
      End
      Begin Threed.SSCommand BtnEditO 
         Height          =   375
         Left            =   495
         TabIndex        =   84
         ToolTipText     =   "Grabar"
         Top             =   70
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0542
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   375
         Left            =   1710
         TabIndex        =   85
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0654
      End
      Begin Threed.SSCommand BtnDelO 
         Height          =   375
         Left            =   1305
         TabIndex        =   86
         ToolTipText     =   "Eliminar"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0B96
      End
      Begin Threed.SSCommand BtnExitO 
         Height          =   375
         Left            =   3690
         TabIndex        =   87
         ToolTipText     =   "Salir"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0CA0
      End
      Begin Threed.SSCommand BtnCalculadora 
         Height          =   375
         Left            =   3285
         TabIndex        =   88
         ToolTipText     =   "Calculadora"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "C"
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0DAA
      End
      Begin Threed.SSCommand cmdfacturas 
         Height          =   375
         Left            =   2070
         TabIndex        =   89
         ToolTipText     =   "Facturas"
         Top             =   75
         Visible         =   0   'False
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0DC6
      End
      Begin Threed.SSCommand cmdusuarios 
         Height          =   375
         Left            =   2880
         TabIndex        =   90
         ToolTipText     =   "Usuarios"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":0F60
      End
      Begin Threed.SSCommand cmdtrans 
         Height          =   375
         Left            =   2475
         TabIndex        =   91
         ToolTipText     =   "Transferencia"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":10FA
      End
      Begin Threed.SSCommand cmdmodifica 
         Height          =   375
         Left            =   900
         TabIndex        =   92
         ToolTipText     =   "Modificar"
         Top             =   75
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483630
         Picture         =   "fregisco.frx":11F4
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   780
      Left            =   90
      TabIndex        =   68
      Top             =   6480
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   1376
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.CheckBox Checkdatos 
         Height          =   240
         Left            =   5805
         TabIndex        =   93
         Top             =   405
         Width           =   240
      End
      Begin VB.TextBox txtdcto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7620
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtredresta 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6975
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtredsuma 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6210
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox TxtOtrImp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   405
         Width           =   1140
      End
      Begin VB.TextBox TxtIgv 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   405
         Width           =   1140
      End
      Begin Threed.SSPanel PnlTotal 
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   69
         Top             =   405
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   4
         Enabled         =   0   'False
      End
      Begin Threed.SSPanel PnlMonIna 
         Height          =   285
         Index           =   0
         Left            =   1710
         TabIndex        =   70
         Top             =   405
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   4
         Enabled         =   0   'False
      End
      Begin Threed.SSPanel PnlBasImp 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   71
         Top             =   405
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   4
         Enabled         =   0   'False
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "+ Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5670
         TabIndex        =   94
         Top             =   135
         Width           =   555
      End
      Begin VB.Label PnlSigMon 
         BackColor       =   &H80000004&
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   81
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9015
         TabIndex        =   80
         Top             =   135
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Dctos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7800
         TabIndex        =   78
         Top             =   135
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "-Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7065
         TabIndex        =   77
         Top             =   135
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "+ Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6300
         TabIndex        =   76
         Top             =   135
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label PnlOtrImp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Otros Impuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4320
         TabIndex        =   75
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label PnlImpuesto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "IGV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3015
         TabIndex        =   74
         Top             =   135
         Width           =   1125
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Monto Inafecto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1860
         TabIndex        =   73
         Top             =   135
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Base Imp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   72
         Top             =   135
         Width           =   705
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   960
      Left            =   2160
      TabIndex        =   61
      Top             =   525
      Width           =   7890
      _Version        =   65536
      _ExtentX        =   13917
      _ExtentY        =   1693
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox TxtCodPrv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   675
         MaxLength       =   4
         TabIndex        =   1
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox TxtNomPrv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4050
         TabIndex        =   4
         Top             =   180
         Width           =   3765
      End
      Begin VB.TextBox TxtDirPrv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   675
         TabIndex        =   5
         Top             =   540
         Width           =   3300
      End
      Begin VB.TextBox TxtRucPrv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1665
         MaxLength       =   11
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox TxtTelPrv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4860
         TabIndex        =   6
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3510
         MaxLength       =   4
         TabIndex        =   3
         Top             =   180
         Width           =   465
      End
      Begin VB.TextBox txtocompra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   7
         Top             =   540
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Prov."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   67
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "RUC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1260
         TabIndex        =   66
         Top             =   225
         Width           =   345
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Direcc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   65
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cta. Cont."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4140
         TabIndex        =   64
         Top             =   585
         Width           =   705
      End
      Begin VB.Label lblcod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cod."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3015
         TabIndex        =   63
         Top             =   225
         Width           =   330
      End
      Begin VB.Label lblocompra 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O. Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6030
         TabIndex        =   62
         Top             =   585
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   960
      Left            =   90
      TabIndex        =   59
      Top             =   540
      Width           =   2040
      _Version        =   65536
      _ExtentX        =   3598
      _ExtentY        =   1693
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox TxtNumMov 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   765
         MaxLength       =   7
         TabIndex        =   0
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txtmesmov 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   79
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "Nº MOVIMIENTO"
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   90
         TabIndex        =   60
         Top             =   135
         Width           =   1815
      End
   End
   Begin VB.TextBox txtfocus 
      Height          =   285
      Left            =   10305
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   5445
      Width           =   555
   End
   Begin VB.Data DataFactura 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4770
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtdepa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4365
      MaxLength       =   2
      TabIndex        =   56
      Top             =   4545
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Data DataAnticipo 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   585
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5355
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtctadestino 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4140
      MaxLength       =   12
      TabIndex        =   55
      Top             =   4905
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtdocpag 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11880
      MaxLength       =   11
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ComboBox cmbdocum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11760
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2355
      Left            =   90
      TabIndex        =   30
      Top             =   1560
      Width           =   9990
      _Version        =   65536
      _ExtentX        =   17621
      _ExtentY        =   4154
      _StockProps     =   14
      Caption         =   "Datos del Documento"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   3
      Begin VB.TextBox TxtPoliza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8550
         MaxLength       =   3
         TabIndex        =   11
         Top             =   315
         Width           =   1260
      End
      Begin MSMask.MaskEdBox TxtFecha 
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   1035
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtref 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   20
         Top             =   1530
         Width           =   1440
      End
      Begin VB.ComboBox cmbfpagos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2805
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1035
         Width           =   2160
      End
      Begin VB.ComboBox cmbigv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1530
         Width           =   4830
      End
      Begin VB.ComboBox CmbTipDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   315
         Width           =   2175
      End
      Begin VB.TextBox TxtRefere 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1035
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1935
         Width           =   8700
      End
      Begin VB.TextBox TxtNumDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         MaxLength       =   7
         TabIndex        =   10
         Top             =   315
         Width           =   1005
      End
      Begin VB.TextBox TxtSerDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   9
         Top             =   315
         Width           =   420
      End
      Begin VB.TextBox TxtTipCam 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         TabIndex        =   13
         Text            =   "0.000"
         Top             =   1035
         Width           =   960
      End
      Begin Threed.SSCheck chkgg 
         Height          =   195
         Left            =   10050
         TabIndex        =   31
         Top             =   405
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Gastos Generales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   600
         Left            =   90
         TabIndex        =   32
         Top             =   135
         Width           =   2130
         _Version        =   65536
         _ExtentX        =   3757
         _ExtentY        =   1058
         _StockProps     =   14
         Caption         =   "Moneda"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption OpcDol 
            Height          =   195
            Left            =   1080
            TabIndex        =   15
            Top             =   270
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Dólares"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption OpcSol 
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   270
            Width           =   780
            _Version        =   65536
            _ExtentX        =   1376
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin MSMask.MaskEdBox TxtFecVen 
         Height          =   285
         Left            =   5505
         TabIndex        =   17
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfecharec 
         Height          =   285
         Left            =   6840
         TabIndex        =   18
         Top             =   1035
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Lblpol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Póliza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   95
         Top             =   360
         Width           =   420
      End
      Begin VB.Line Line1 
         X1              =   6675
         X2              =   6555
         Y1              =   315
         Y2              =   555
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5955
         TabIndex        =   42
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Fecha Rec."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6930
         TabIndex        =   41
         Top             =   795
         Width           =   840
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2865
         TabIndex        =   40
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label LblFecVen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Fec.Venc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   38
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         TabIndex        =   37
         Top             =   1935
         Width           =   780
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "T/C Oficial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1605
         TabIndex        =   36
         Top             =   800
         Width           =   885
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Fecha Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   800
         Width           =   840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Serie/Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5130
         TabIndex        =   34
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2340
         TabIndex        =   33
         Top             =   360
         Width           =   345
      End
   End
   Begin Threed.SSPanel PnlOficial 
      Height          =   60
      Left            =   11880
      TabIndex        =   43
      Top             =   120
      Width           =   105
      _Version        =   65536
      _ExtentX        =   185
      _ExtentY        =   106
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox TxtIgv 
         Height          =   330
         Index           =   1
         Left            =   4410
         TabIndex        =   45
         Top             =   315
         Width           =   870
      End
      Begin VB.TextBox TxtOtrImp 
         Height          =   330
         Index           =   1
         Left            =   5670
         TabIndex        =   44
         Top             =   315
         Width           =   1275
      End
      Begin Threed.SSPanel PnlBasImp 
         Height          =   375
         Index           =   1
         Left            =   810
         TabIndex        =   46
         Top             =   315
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin Threed.SSPanel PnlMonIna 
         Height          =   285
         Index           =   1
         Left            =   2565
         TabIndex        =   47
         Top             =   360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel PnlTotal 
         Height          =   330
         Index           =   1
         Left            =   7245
         TabIndex        =   48
         Top             =   315
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label24 
         Caption         =   "Datos a Transferir al Registro de Compras Oficial"
         Height          =   240
         Left            =   270
         TabIndex        =   50
         Top             =   45
         Width           =   4695
      End
      Begin VB.Label PnlSigMon 
         Caption         =   "S/."
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   49
         Top             =   405
         Width           =   375
      End
   End
   Begin Threed.SSPanel PnlAyuda 
      Height          =   30
      Left            =   90
      TabIndex        =   51
      Top             =   7425
      Width           =   105
      _Version        =   65536
      _ExtentX        =   185
      _ExtentY        =   -53
      _StockProps     =   15
      Caption         =   "F2 --> Ayuda                    F4 ---> Eliminar      "
      ForeColor       =   12582912
      BackColor       =   -2147483648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11261.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      FloodColor      =   12582912
   End
   Begin Threed.SSCommand CmdGrabar 
      Height          =   15
      Left            =   2025
      TabIndex        =   52
      Top             =   6165
      Visible         =   0   'False
      Width           =   30
      _Version        =   65536
      _ExtentX        =   53
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Grabar"
   End
   Begin Threed.SSCommand CmdImprimir 
      Height          =   15
      Left            =   5490
      TabIndex        =   53
      Top             =   6165
      Visible         =   0   'False
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Imprimir"
   End
   Begin Threed.SSCommand CmdSalir 
      Height          =   15
      Left            =   7245
      TabIndex        =   54
      Top             =   6165
      Visible         =   0   'False
      Width           =   45
      _Version        =   65536
      _ExtentX        =   -79
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Salir"
   End
   Begin TrueOleDBGrid70.TDBGrid Grd_Docum 
      Bindings        =   "fregisco.frx":1306
      Height          =   2565
      Left            =   90
      TabIndex        =   22
      Top             =   3915
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   4524
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Item"
      Columns(0).DataField=   "F3ITEM"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Año"
      Columns(1).DataField=   "F3ANOORDEN"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Refer."
      Columns(2).DataField=   "F3ORDEN"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Gastos"
      Columns(3).DataField=   "F3GASTO"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Cta. Contable"
      Columns(4).DataField=   "F3CTACON"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "C.C."
      Columns(5).DataField=   "F3CENCOS"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Concepto"
      Columns(6).DataField=   "F3CONCEPTO"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Importe"
      Columns(7).DataField=   "F3IMPORTE"
      Columns(7).DefaultValue=   "0.00"
      Columns(7).DefaultValue.vt=   8
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Af."
      Columns(8).DataField=   "F3AFECTO"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "D/H"
      Columns(9).DataField=   "F3DEBHAB"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Departamento"
      Columns(10).DataField=   "F3DEPARTAMENTO"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Cta. Destino"
      Columns(11).DataField=   "F3CTADESTINO"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=65810"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=926"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=847"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=65808"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1058"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=979"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=65808"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1244"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1164"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=65808"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2037"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1958"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=65808"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=1693"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1614"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=65808"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=5636"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=5556"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=65808"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1879"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1799"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=65810"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=529"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=450"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=65808"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(9).Width=635"
      Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=556"
      Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=65808"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(52)=   "Column(10).Width=2037"
      Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=1958"
      Splits(0)._ColumnProps(55)=   "Column(10)._ColStyle=65808"
      Splits(0)._ColumnProps(56)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(58)=   "Column(11).Width=1826"
      Splits(0)._ColumnProps(59)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(11)._WidthInPix=1746"
      Splits(0)._ColumnProps(61)=   "Column(11)._ColStyle=65808"
      Splits(0)._ColumnProps(62)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "Splits(0).Style:id=43,.parent=1"
      _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=1,.valignment=2"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=1,.valignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
      _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44"
      _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
      _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
      _StyleDefs(79)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44"
      _StyleDefs(81)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
      _StyleDefs(82)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
      _StyleDefs(83)  =   "Named:id=33:Normal"
      _StyleDefs(84)  =   ":id=33,.parent=0"
      _StyleDefs(85)  =   "Named:id=34:Heading"
      _StyleDefs(86)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   ":id=34,.wraptext=-1"
      _StyleDefs(88)  =   "Named:id=35:Footing"
      _StyleDefs(89)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   "Named:id=36:Selected"
      _StyleDefs(91)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(92)  =   "Named:id=37:Caption"
      _StyleDefs(93)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(94)  =   "Named:id=38:HighlightRow"
      _StyleDefs(95)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(96)  =   "Named:id=39:EvenRow"
      _StyleDefs(97)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(98)  =   "Named:id=40:OddRow"
      _StyleDefs(99)  =   ":id=40,.parent=33"
      _StyleDefs(100) =   "Named:id=41:RecordSelector"
      _StyleDefs(101) =   ":id=41,.parent=34"
      _StyleDefs(102) =   "Named:id=42:FilterBar"
      _StyleDefs(103) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid grdanticipos 
      Bindings        =   "fregisco.frx":1333
      Height          =   1275
      Left            =   3480
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   2249
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=65808"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=65808"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(9)   =   ":id=2,.fgcolor=&H80000012&"
      _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(19)  =   "Splits(0).Style:id=43,.parent=1"
      _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(22)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(25)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmregiscom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DbTempfac    As DAO.Database
Dim db_tempo     As DAO.Database
Dim Tabla        As DAO.Recordset
Dim tem_afecto   As DAO.Recordset
Dim tem_inafecto As DAO.Recordset

Dim WNUMERO      As String
Dim tmpfactura   As String

Dim WITEMS       As Integer
Dim swigv        As Integer

Dim wbasimp      As Double
Dim wmonina      As Double
Dim wIMPUESTO    As Double
Dim wotrimp      As Double
Dim WGRABO       As Integer

Dim wctacble     As Integer
Dim wimporte     As Integer
Dim swmoneda     As Integer
Dim swsalir      As Integer
Dim winddepa     As String

Dim TEMP_CABE    As DAO.Recordset
Dim TEMP_DET     As DAO.Recordset

Dim swmodigv     As Integer
Dim swmodcta     As Integer
Dim swayuda      As Integer

Dim WCHANGE      As Integer
Dim keyenter     As String

Dim tbfpagos     As DAO.Recordset
Dim windigraba   As Integer

Dim swtc         As Integer

Dim sw_grabanse   As Integer
Dim sw_doc        As Integer
Dim swgraba_valing  As Integer

Dim sw_numdoc       As Integer
Dim wcosto          As Integer
Dim wnewnumero      As Integer
Dim xtipmonoc       As String
Dim xforpagoc       As String
Dim xfecvencoc      As Date
Dim xtotaloc        As Double
Dim cflag           As Integer
Dim xmonto_anterior As Double
Dim wvcodprov       As String
Dim wvtipdocu       As String
Dim wvserdoc        As String
Dim wvnumdoc        As String
Dim gobra           As String
Dim xflag           As Integer
Dim meslet          As String
Dim WF1INGCTA_CC    As String * 1
Dim pctacen         As String

Private Sub actualiza()
Dim TbConsulta  As DAO.Recordset
Dim I           As Integer
Dim xtipopago   As String
Dim Msql        As String

    Me.MousePointer = 11

    wvcodprov = "" & TbCabRegis.Fields("F4CODPRV")
    wvtipdocu = TbCabRegis.Fields("F4TIPDOC")
    wvserdoc = "" & TbCabRegis.Fields("F4SERDOC")
    wvnumdoc = "" & TbCabRegis.Fields("F4NUMDOC")
    
    TxtCodPrv.Text = "" & TbCabRegis.Fields("F4CODPRV")
    TxtNomPrv.Text = "" & TbCabRegis.Fields("F4NOMPRV")
    TxtDirPrv.Text = "" & TbCabRegis.Fields("F4DIRPRV")
    TxtRucPrv.Text = "" & TbCabRegis.Fields("F4RUCPRV")
    TxtTelPrv.Text = "" & TbCabRegis.Fields("F4CTACONT")

    txtref.Text = "" & TbCabRegis.Fields("F4DIARIO") & ""

    gcodppp = "" & TbCabRegis.Fields("F4GRUPO")
    Text1.Text = "" & TbCabRegis.Fields("F4GRUPO")
    gorden_cs = TbCabRegis.Fields("f4tiporden") & ""
    
    Rem txtocompra.Text = Val(tbcabregis.Fields("F4OCOMPRA") & "")
    
    If wocompra = "*" Then
        If wf1tipdoc_asoc = "V" Then
            txtocompra.Text = TbCabRegis.Fields("F4VALESING") & ""
        Else
            txtocompra.Text = TbCabRegis.Fields("F4OCOMPRA") & ""
        End If
    End If
    
    
    gobra = "" & TbCabRegis.Fields("F4OBRA")
    txtfecha.Text = "" & Format(TbCabRegis.Fields("F4FECHA"), "DD/MM/YYYY")
    chkgg.Value = IIf(TbCabRegis.Fields("f4gastgen") & "" = "*", True, False)
    
    '--------------------------- formas de pago
    For I = 0 To cmbfpagos.ListCount - 1
        If Left(Right(cmbfpagos.List(I), 4), 3) = TbCabRegis.Fields("f4forpag") Then
            cmbfpagos.ListIndex = I
        End If
    Next
    '---------------------------

    xtipopago = Right(cmbfpagos.Text, 1)

    If xtipopago = "C" Then
       cmbdocum.ListIndex = Val(TbCabRegis.Fields("F4TIPDOCPAG"))
       txtdocpag.Text = "" & TbCabRegis.Fields("F4DOCPAG")
    Else
        TxtFecVen.Text = "" & Format(TbCabRegis.Fields("F4FECVEN"), "DD/MM/YYYY")
    End If
    txtfecharec.Text = "" & Format(TbCabRegis.Fields("F4FECHAREC"), "DD/MM/YYYY")

    If wocompra = "*" Then
        lblocompra.Visible = True
        txtocompra.Visible = True
        If wf1tipdoc_asoc = "V" Then
            lblocompra.Caption = "Vale Ing."
        Else
            lblocompra.Caption = "O. Compra"
        End If
    Else
        lblocompra.Visible = False
        txtocompra.Visible = False
    End If


    For I% = 0 To CmbTipDoc.ListCount - 1
        If Left(Trim(Right(CmbTipDoc.List(I%), 5)), 2) = TbCabRegis.Fields("f4tipdoc") Then
            CmbTipDoc.ListIndex = I%
        End If
    Next

    For I% = 0 To cmbigv.ListCount - 1
        If Right(cmbigv.List(I%), 3) = TbCabRegis.Fields("f4codigv") Then
            cmbigv.ListIndex = I%
        End If
    Next
   
    TxtSerDoc.Text = "" & TbCabRegis.Fields("F4SERDOC")
    TxtNumDoc.Text = "" & TbCabRegis.Fields("F4NUMDOC")

    TxtTipCam.Text = TbCabRegis.Fields("F4TIPCAM")
    TxtRefere.Text = "" & TbCabRegis.Fields("F4REFERE")

    PnlBasImp(0).Caption = Format(TbCabRegis.Fields("F4BASIMP"), "###,##0.00")
    PnlMonIna(0).Caption = Format(TbCabRegis.Fields("F4MONINA"), "###,##0.00")
    If UCase(Right(CmbTipDoc.Text, 3)) = "HON" Then
        PnlImpuesto.Caption = "Retención"
        PnlOtrImp.Caption = "I.E.S."
        txtigv(0).Text = Format(TbCabRegis.Fields("F4MONTORET"), "###,##0.00")
        TxtOtrImp(0).Text = Format(TbCabRegis.Fields("F4FONAVI"), "###,##0.00")
    Else
        txtigv(0).Text = Format(TbCabRegis.Fields("F4IGV"), "###,##0.00")
        TxtOtrImp(0).Text = Format(TbCabRegis.Fields("F4OTRIMP"), "###,##0.00")
    End If
    txtredsuma.Text = Format(TbCabRegis.Fields("f4redsuma"), "###,##0.00")
    txtredresta.Text = Format(TbCabRegis.Fields("f4redresta"), "###,##0.00")
    txtdcto.Text = Format(TbCabRegis.Fields("f4dcto"), "###,##0.00")
    PnlTotal(0).Caption = Format(TbCabRegis.Fields("F4TOTAL"), "###,##0.00")

    If TbCabRegis.Fields("F4MONEDA") = "S" Then
       OpcSol.Value = True
    Else
       OpcDol.Value = True
       TbOfiRegis.Index = "IDMESNUM"
       TbOfiRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
       If Not TbOfiRegis.NoMatch Then
          PnlBasImp(1).Caption = Format(TbOfiRegis.Fields("F4BASIMP"), "###,##0.00")
          PnlMonIna(1).Caption = Format(TbOfiRegis.Fields("F4MONINA"), "###,##0.00")
          If UCase(Right(CmbTipDoc.Text, 3)) = "HON" Then
            txtigv(1).Text = Format(TbOfiRegis.Fields("F4MONTORET"), "###,##0.00")
            TxtOtrImp(1).Text = Format(TbOfiRegis.Fields("F4FONAVI"), "###,##0.00")
          Else
            txtigv(1).Text = Format(TbOfiRegis.Fields("F4IGV"), "###,##0.00")
            TxtOtrImp(1).Text = Format(TbOfiRegis.Fields("F4OTRIMP"), "###,##0.00")
          End If
          PnlTotal(1).Caption = Format(TbOfiRegis.Fields("F4TOTAL"), "###,##0.00")
       End If
    End If
    DbTempfac.Execute ("Delete From temp_det")
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Msql = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmesmov.Text + "' AND F4NUMMOV = '" + TxtNumMov.Text + "' Order by F3ITEM"
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    Do While Not TbConsulta.EOF
       DataFactura.Recordset.AddNew
       DataFactura.Recordset.Fields("F3ITEM") = 0 + Format(TbConsulta.Fields("F3ITEM"), "00")
       DataFactura.Recordset.Fields("F3CenCos") = "" & TbConsulta.Fields("F3CenCos")
       DataFactura.Recordset.Fields("F3GASTO") = "" & TbConsulta.Fields("F3GASTO")
       DataFactura.Recordset.Fields("F3CTACON") = "" & TbConsulta.Fields("F3CTACON")
       DataFactura.Recordset.Fields("F3CONCEPTO") = "" & TbConsulta.Fields("F3CONCEPTO")
       DataFactura.Recordset.Fields("F3IMPORTE") = 0# + Format(TbConsulta.Fields("F3IMPORTE"), "###,##0.00")
       DataFactura.Recordset.Fields("F3AFECTO") = "" & TbConsulta.Fields("F3AFECTO")
       DataFactura.Recordset.Fields("F3DEBHAB") = "" & TbConsulta.Fields("F3DEBHAB")
       DataFactura.Recordset.Fields("F3ANOORDEN") = "" & TbConsulta.Fields("F3ANOORDEN")
       DataFactura.Recordset.Fields("F3ORDEN") = "" & TbConsulta.Fields("F3ORDEN")
       DataFactura.Recordset.Fields("F3DEPARTAMENTO") = "" & TbConsulta.Fields("F3DEPARTAMENTO")
       DataFactura.Recordset.Fields("F3CTADESTINO") = "" & TbConsulta.Fields("F3CTADESTINO")
       DataFactura.Recordset.Update
       TbConsulta.MoveNext
       'BtnEditO.Visible = True
    Loop
    DataFactura.Refresh
    Grd_Docum.Refresh
    TbConsulta.Close
    
    Me.MousePointer = 1

End Sub

Private Sub Actualiza_DebHab(debhab As String)

    If DataFactura.Recordset.RecordCount > 0 Then
        DataFactura.Recordset.MoveFirst
        Do While Not DataFactura.Recordset.EOF
            DataFactura.Recordset.Edit
            DataFactura.Recordset.Fields("f3debhab") = "" & debhab
            DataFactura.Recordset.Update
            DataFactura.Recordset.MoveNext
        Loop
    End If
    DataFactura.Refresh
    Grd_Docum.Refresh

End Sub

Private Sub BtnAddO_Click()
        
    'If BtnEditO.Visible = True Then
        If WGRABO = 0 Then
            If MsgBox(" El registro no ha sido grabado . Desea grabar ?", 36, "Atención") = 6 Then grabar
        End If
    'End If
    WGRABO = 0
    Me.MousePointer = 11
    nuevo
    sw_doc = False
    Me.MousePointer = 1

End Sub

Private Sub BtnCalculadora_Click()
On Error GoTo BtnConsulta
Dim X   As Variant

    X = Shell("Calc.exe", 1)

BtnConsulta: Resume Next

End Sub

Private Sub BtnDelO_Click()

    Me.MousePointer = 11
    elimina
    Me.MousePointer = 1

End Sub

Private Sub BtnEditO_Click()
Dim swgg As Integer
Dim swcc As Integer

    '-----------------------------------------------------
    If Len(Trim(TxtTelPrv.Text)) = 0 Then
        MsgBox "Verifique la cuenta contable del total del documento está en blanco.", 48, "Compras"
    End If
    '-----------------------------------------------------
    
    If wf1inggasto = "*" Then
        swgg = False
        swcc = False
        If Tabla.RecordCount > 0 Then
            Tabla.MoveFirst
            Do While Not Tabla.EOF
                If Len(Trim(Tabla.Fields("f3gasto") & "")) = 0 Then
                    swgg = True
                    Exit Do
                End If
                If val_cc(Tabla.Fields("f3ctacon") & "") = True Then
                    If Len(Trim(Tabla.Fields("f3cencos") & "")) = 0 Then
                        swcc = True
                        Exit Do
                    Else
                        If valida_centros2(Tabla.Fields("f3ctacon") & "", Tabla.Fields("f3cencos") & "") Then
                            swcc = False
                        Else
                            swcc = True
                            Exit Do
                        End If
                    End If
                End If
                Tabla.MoveNext
            Loop
        End If
        If swgg = True Or swcc = True Then
            MsgBox "No se puede grabar, falta asignar el código de gasto o el centro de costo. Verifique.", 48, "Compras"
            Grd_Docum.SetFocus
        Else
            Me.MousePointer = 11
            grabar
            WGRABO = 1
            If sw_grabanse = True Then
                If Right(cmbfpagos.Text, 1) = "C" And wbancos = "*" Then
                    frmactban.Show 1
                End If
            End If
            Me.MousePointer = 1
            BtnAddO.SetFocus
        
            If Len(Trim(txtocompra.Text)) > 0 Then
                Rem Revisa_anticipos
            End If
        End If
    Else
        '------------------
        Me.MousePointer = 11
        grabar
        WGRABO = 1
        If sw_grabanse = True Then
            If Right(cmbfpagos.Text, 1) = "C" And wbancos = "*" Then
                frmactban.Show 1
            End If
        End If
        Me.MousePointer = 1
        BtnAddO.SetFocus
    
        If Len(Trim(txtocompra.Text)) > 0 Then
            Rem Revisa_anticipos
        End If
    End If

    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Or UCase(Right(Trim(CmbTipDoc.Text), 3)) = "DEB" Then
        cmdfacturas.Visible = True
        If sw_doc = False Then
            MsgBox "Falta asignar el Nº de Factura para este documento.", 48, "Compras"
        End If
    End If

End Sub

Private Sub BtnExitO_Click()
Dim X  As Integer
Dim swgg As Integer
Dim swcc As Integer

    swsalir = 0
    If WGRABO = 0 And DataFactura.Recordset.RecordCount > 0 Then
        X = MsgBox("El registro no ha sido grabado . Desea grabar?", 35)
        Select Case X
            Case 6:
                If wf1inggasto = "*" Then
                    swgg = False
                    swcc = False
                    If Tabla.RecordCount > 0 Then
                        Tabla.MoveFirst
                        Do While Not Tabla.EOF
                            If Len(Trim(Tabla.Fields("f3gasto") & "")) = 0 Then
                                swgg = True
                                Exit Do
                            End If
                            If val_cc(Tabla.Fields("f3ctacon") & "") = True Then
                                If Len(Trim(Tabla.Fields("f3cencos") & "")) = 0 Then
                                    swcc = True
                                    Exit Do
                                Else
                                    If valida_centros2(Tabla.Fields("f3ctacon") & "", Tabla.Fields("f3cencos") & "") Then
                                        swcc = False
                                    Else
                                        swcc = True
                                        Exit Do
                                    End If
                                End If
                            End If
                            Tabla.MoveNext
                        Loop
                    End If
                    If swgg = True Or swcc = True Then
                        MsgBox "No se puede grabar, falta asignar el código de gasto o el centro de costo. Verifique.", 48, "Compras"
                        Grd_Docum.SetFocus
                    Else
                        grabar
                        swsalir = 1
                        Unload Me
                    End If
                Else
                    grabar
                    swsalir = 1
                    Unload Me
                End If
            Case 7: swsalir = 1
                    Unload Me
            Case 2: Exit Sub
        End Select
    Else
        WGRABO = 0
        Unload Me
    End If
    
End Sub

Private Sub BtnPrint_Click()
Dim xflagnse        As Integer
Dim cnomfile        As String
Dim xcorre          As Double
Dim xvar            As Integer

    DbTempfac.Execute "DELETE * FROM VOU_CAB"
    DbTempfac.Execute "DELETE * FROM VOU_DET"

    Set TEMP_CABE = DbTempfac.OpenRecordset("vou_cab")
    Set temp_deta = DbTempfac.OpenRecordset("vou_det")

    LLENA_TEMPCAB
    LLENA_TEMPDET

    temp_deta.Close
    TEMP_CABE.Close

    xflagnse = False
    '---------------------------------------------------------------
    If wf1traslado = "*" Then '- TECPRO, despues de imprimir el voucher, lo traslada a ctasxpagar
        If gctapag = "*" And Right(cmbfpagos.Text, 1) = "F" And xflag = False Then
            TbCabRegis.Index = "IDMESNUM"
            TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
            If Not TbCabRegis.NoMatch Then
            
                Set DbCtaPag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
                Set tbcabcta = DbCtaPag.OpenRecordset("PAG_DCTO")
                xvar = False
                tbcabcta.Index = "nro_corre"
                tbcabcta.Seek "=", Val("" & TbCabRegis.Fields("f4correla"))
                If Not tbcabcta.NoMatch Then
                    If Val(TbCabRegis.Fields("f4correla") & "") = 0 Then
                        xvar = False
                    Else
                        If Val(Format(tbcabcta.Fields("SALDO"), "0.00")) <> Val(Format(tbcabcta.Fields("TOTALO"), "0.00")) Then
                            xflagnse = True
                            xvar = False
                        Else
                            xvar = True
                            xcorre = tbcabcta.Fields("correla")
                        End If
                    End If
                Else
                    xvar = False
                End If
                tbcabcta.Close
                DbCtaPag.Close
                
                If xvar = False Then
                    '-----------------------------------
                    trans_ctasxpagar "N", Me, 0
                    '-----------------------------------
                    TbCabRegis.Edit
                    TbCabRegis.Fields("f4correla") = gcorre
                    TbCabRegis.Fields("f4usuarioimp") = wusuario
                    TbCabRegis.Fields("f4fechaimp") = Format(Date, "dd/mm/yyyy")
                    TbCabRegis.Fields("f4horaimp") = Time
                    cnomfile = TbCabRegis.Name
                    TbCabRegis.Update
                    '-----------------
                Else
                    If xflagnse = True Then
                        MsgBox "Verifique el documento ya ha sido aplicado", 48, "Compras"
                    Else
                        trans_ctasxpagar "M", Me, xcorre
                        TbCabRegis.Edit
                        TbCabRegis.Fields("f4usuarioimp") = wusuario
                        TbCabRegis.Fields("f4fechaimp") = Format(Date, "dd/mm/yyyy")
                        TbCabRegis.Fields("f4horaimp") = Time
                        cnomfile = TbCabRegis.Name
                        TbCabRegis.Update
                        Rem NSE tbcabcta.Edit
                    End If
                End If
            End If
        End If
    End If

    '---------------------------------------------------------------
    Me.MousePointer = 11
    Report1.DataFiles(0) = wrutatemp & "\temp_com.mdb"
    If wf1formatov = "2" Then
        Report1.ReportFileName = wrutatemp & "\rpt_vou2.rpt"
    Else
        Report1.ReportFileName = wrutatemp & "\rpt_vou.rpt"
    End If
    Report1.Action = 1
    Me.MousePointer = 1
    '---------------------------------------------------------------

    BtnAddO.SetFocus

End Sub

Private Sub calcula()
Dim I As Integer
  
    I = 0
    wbasimp = 0#: wmonina = 0#: wIMPUESTO = 0#: wotrimp = 0#
    'For i = 0 To Grd_Docum.Row - 1
    '   Grd_Docum.Row = i
    '   If Val(Format(Grd_Docum.Columns(7), "#0.00")) <> 0 Then
    '   ' NO CALCULAR FONAVI SI AFECTO = *
    '       If Grd_Docum.Columns(8) = "*" Then
    '          wbasimp = wbasimp + Format(Grd_Docum.Columns(7), "#0.00")
    '          If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
    '             wIMPUESTO = wIMPUESTO + Format(Grd_Docum.Columns(7) * gretenc / 100, "#0.00")
    '             wotrimp = wotrimp + Format(Grd_Docum.Columns(7) * gfonavi / 100, "#0.00")
    '          Else
    '             wIMPUESTO = wIMPUESTO + Format(Grd_Docum.Columns(7) * wigv / 100, "#0.00")
    '          End If
    '          wIMPUESTO = Format(wIMPUESTO, "#0.00")
    '       Else
    '          Rem If UCase(Right(Trim(cmbtipdoc.Text), 3)) = "HON" Then
    '          Rem    wimpuesto = wimpuesto + Format(grd_docum.columns(7) * gretenc / 100, "#0.00")
    '          Rem    wOtrImp = wOtrImp + Format(grd_docum.columns(7) * gfonavi / 100, "#0.00")
    '          Rem End If
    '          wmonina = wmonina + Format(Grd_Docum.Columns(7), "#0.00")
    '       End If
    '   End If
    'Next i
    
    If Tabla.RecordCount > 0 Then
        Tabla.MoveFirst
        Do While Not Tabla.EOF
            If Val("" & Tabla.Fields("f3importe")) <> 0 Then
                If "" & Tabla.Fields("f3afecto") = "*" Then
                    wbasimp = wbasimp + Val("" & Tabla.Fields("f3importe"))
                    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                        wIMPUESTO = wIMPUESTO + (Val("" & Tabla.Fields("f3importe")) * gretenc / 100)
                        wotrimp = wotrimp + (Val("" & Tabla.Fields("f3importe")) * gfonavi / 100)
                    Else
                        wIMPUESTO = wIMPUESTO + (Val("" & Tabla.Fields("f3importe")) * wigv / 100)
                    End If
                Else
                    wmonina = wmonina + Val("" & Tabla.Fields("f3importe"))
                End If
            End If
            Tabla.MoveNext
        Loop
    End If
    
    
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "HON" Then
        wotrimp = Format(Val(Format(TxtOtrImp(0).Text & "", "0.00")), "#0.00")
    End If
    PnlBasImp(0).Caption = Format(wbasimp, "###,##0.00")
    PnlMonIna(0).Caption = Format(wmonina, "###,##0.00")
    'norma
    If swmodigv = True Then
        txtigv(0).Text = Format(wIMPUESTO, "###,##0.00")
    Else
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
            If Val(Format(txtigv(0).Text, "#0.00")) < 0 Then
                txtigv(0).Text = Val(Format(txtigv(0).Text, "#0.00")) * -1
            End If
        End If
    End If
    TxtOtrImp(0).Text = Format(wotrimp, "###,##0.00")

    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
       PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
    Else
       PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
    End If
    If OpcDol.Value = True Then  'dolares
       PnlBasImp(1).Caption = Format(wbasimp * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
       PnlMonIna(1).Caption = Format(wmonina * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
       'norma
       If swmodigv = True Then
          txtigv(1).Text = Format(wIMPUESTO * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
       End If
       TxtOtrImp(1).Text = Format(wotrimp * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
       If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
          PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) - Val(Format(TxtOtrImp(1).Text, "0.00")) - Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
       Else
          PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) + Val(Format(TxtOtrImp(1).Text, "0.00")) + Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
       End If
    Else   ' soles
       PnlBasImp(1).Caption = Format(wbasimp, "###,##0.00")
       PnlMonIna(1).Caption = Format(wmonina, "###,##0.00")
       'norma
       If swmodigv = True Then
          txtigv(1).Text = Format(wIMPUESTO, "###,##0.00")
       End If
       TxtOtrImp(1).Text = Format(wotrimp, "###,##0.00")
       If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
          PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) - Val(Format(TxtOtrImp(1).Text, "0.00")) - Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
       Else
          PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) + Val(Format(TxtOtrImp(1).Text, "0.00")) + Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
       End If
    End If

End Sub

Private Function Calcula_Item() As Integer

    TbDetRegis.Index = "IDMESNUMITE"
    TbDetRegis.Seek ">=", txtmesmov.Text, TxtNumMov.Text, 1
    If Not TbDetRegis.NoMatch Then
       Do While TbDetRegis.Fields("F4MESMOV") = txtmesmov.Text And TbDetRegis.Fields("F4NUMMOV") = TxtNumMov.Text
          TbDetRegis.MoveNext
          If TbDetRegis.EOF Then
             Exit Do
          End If
       Loop
       TbDetRegis.MovePrevious
       If Not TbDetRegis.Bof Then
           If TbDetRegis.Fields("F4MESMOV") = txtmesmov.Text And TbDetRegis.Fields("F4NUMMOV") = TxtNumMov.Text Then
              Calcula_Item = 0 + TbDetRegis.Fields("F3ITEM")
           Else
              Calcula_Item = 0
           End If
       Else
           Calcula_Item = 0
       End If
    End If

End Function

Private Sub Checkdatos_Click()
    If Checkdatos.Value = 1 Then
        Label19.Visible = True
        Label20.Visible = True
        Label21.Visible = True
        txtdcto.Visible = True
        txtredresta.Visible = True
        txtredsuma.Visible = True
        txtredsuma.SetFocus
    Else
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        txtdcto.Visible = False
        txtredresta.Visible = False
        txtredsuma.Visible = False
    End If
End Sub

Private Sub chkgg_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      TxtRefere.SetFocus
   End If

End Sub

Private Sub CmbDocum_KeyPress(KeyAscii As Integer)
    
    'If KeyAscii = 13 Then
    '    frmcta.Show 1
    'End If
    
End Sub

Private Sub cmbfpagos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtFecVen.SetFocus
    End If

End Sub

Private Sub cmbfpagos_LostFocus()

    tbfpagos.Index = "IDFORPAG"
    tbfpagos.Seek "=", Left(Right(cmbfpagos.Text, 4), 3)
    If Not tbfpagos.NoMatch Then
        TxtFecVen.Text = Format(CVDate(txtfecha.Text) + tbfpagos.Fields("f2dias"), "dd/mm/yyyy")
    End If

End Sub

Private Sub cmbigv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'If wocompra = "*" Then
        '    tbproveedor.Index = "idcodigo"
        '    tbproveedor.Seek "=", txtCODPRV.Text
        '    If Not tbproveedor.NoMatch Then
        '        If tbproveedor.Fields("f2orden") = True Then
        '            lblocompra.Visible = True
        '            txtocompra.Visible = True
        '            txtocompra.SetFocus
        '        End If
        '    End If
        'Else
            txtref.SetFocus
        'End If
    End If

End Sub

Private Sub CmbTipDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        '------------------------------------------------------
        If wf1numera = "*" Then
            tbmes.Index = "IDDOCMES"
            tbmes.Seek "=", "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00"), mes
            If Not tbmes.NoMatch Then
               txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
               WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
               TxtNumMov.Text = "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00") & Right(Format(WNUMERO, "0000000"), 5)
            End If
        End If
        '------------------------------------------------------
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            PnlImpuesto.Caption = "Retención"
            PnlOtrImp.Caption = "I.E.S."
        Else
            PnlImpuesto.Caption = "IGV"
            PnlOtrImp.Caption = "Otros Impuestos"
        End If
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
            Actualiza_DebHab ("H")
        Else
            Actualiza_DebHab ("D")
        End If
        TxtSerDoc.SetFocus
    End If

End Sub

Private Sub CmbTipDoc_LostFocus()

    '------------------------------------------------------
    If wf1numera = "*" Then
        tbmes.Index = "IDDOCMES"
        tbmes.Seek "=", "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00"), mes
        If Not tbmes.NoMatch Then
            txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
            WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
            TxtNumMov.Text = "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00") & Right(Format(WNUMERO, "0000000"), 5)
        End If
    End If
    '------------------------------------------------------
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
        PnlImpuesto.Caption = "Retención"
        PnlOtrImp.Caption = "I.E.S."
    Else
        PnlImpuesto.Caption = "IGV"
        PnlOtrImp.Caption = "Otros Impuestos"
    End If
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
        Actualiza_DebHab ("H")
    Else
        Actualiza_DebHab ("D")
    End If
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "POL" Then
        Lblpol.Visible = True
        TxtPoliza.Visible = True
    End If
Rem JCG    TxtSerDoc.SetFocus

End Sub

Private Sub cmdfacturas_Click()

    frmr_fac.Show 1
    sw_doc = True

End Sub

Private Sub cmdgrabar_Click()
    Me.MousePointer = 11
    grabar
    Me.MousePointer = 1
End Sub

Private Sub CmdImprimir_Click()
    Me.MousePointer = 11
    'Imprimir
    Me.MousePointer = 1
End Sub

Private Sub cmdmodifica_Click()

    frmhlmov.Show 1
    If Len(Trim(gnummov)) > 0 Then
        TxtNumMov.Text = Format(gnummov, "0000000")
        txtnummov_KeyPress 13
    End If

End Sub

Private Sub cmdsalir_Click()

    Unload Me
    
End Sub

Private Sub cmdtrans_Click()
Dim dbusuarios  As DAO.Database
Dim TbUser      As DAO.Recordset
Dim cnivel      As String
                          
    Set dbusuarios = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
    Set TbUser = dbusuarios.OpenRecordset("EF2USERS")
    TbUser.Index = "IDCODUSER"
    
    TbUser.Seek "=", wusuario
    If Not TbUser.NoMatch Then
        cnivel = "" & TbUser.Fields("F2NIVEL")
    End If
    TbUser.Close
    dbusuarios.Close

    If cnivel = "1" Then
        frmcnttr.Show 1
    Else
        MsgBox "Ud. no tiene acceso a esta opción."
    End If
    'TxtNumMov.SetFocus

End Sub

Private Sub cmdusuarios_Click()

    frmusuam.Show 1

End Sub

Private Sub elimina()
Dim nelem       As Integer
Dim xtempo      As String
Dim cnomfile    As String

    Beep
    If MsgBox("Está seguro de eliminar el registro", 36, "Atención") = 6 Then
        '--------------- cabecera
        TbCabRegis.Index = "IDMESNUM"
        TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
        If Not TbCabRegis.NoMatch Then
            '--------------------------------------
            elimina_provision
            Rem NSE evalua_elim_conta Trim(TbCabRegis.Fields("f4comp_conta") & "")
            '--------------------------------------
            '--------------- ctas x pagar
            Rem If gctapag = "*" And opccre.Value = True Then
            If gctapag = "*" And Right(cmbfpagos.Text, 1) = "F" Then
                Set DbCtaPag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
                Set tbcabcta = DbCtaPag.OpenRecordset("PAG_DCTO")
                tbcabcta.Index = "nro_corre"
                tbcabcta.Seek "=", TbCabRegis.Fields("f4correla")
                If Not tbcabcta.NoMatch Then
                    If Format(tbcabcta.Fields("total"), "0.00") <> Format(tbcabcta.Fields("saldo"), "0.00") Then
                        MsgBox "El documento ya ha sido cancelado en Ctas x Pagar.", 48, "Compras"
                    Else
                        tbcabcta.Delete
                    End If
                End If
                tbcabcta.Close
                DbCtaPag.Close
            End If
            '---------------
            TbCabRegis.Delete
        End If
        '--------------- detalle
        nelem = 1
        TbDetRegis.Index = "IDMESNUMITE"
        TbDetRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text, nelem
        If Not TbCabRegis.NoMatch Then
            Do While TbDetRegis.Fields("F4MESMOV") = Trim(txtmesmov.Text) And TbDetRegis.Fields("F4NUMMOV") = Trim(TxtNumMov.Text)
                TbDetRegis.Delete
                TbDetRegis.MoveNext
                If TbDetRegis.EOF Then Exit Do
                If TbDetRegis.Fields("F4MESMOV") <> Trim(txtmesmov.Text) Or TbDetRegis.Fields("F4NUMMOV") <> Trim(TxtNumMov.Text) Then Exit Do
            Loop
        End If
        '--------------- oficial
        TbOfiRegis.Index = "IDMESNUM"
        TbOfiRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
        If Not TbOfiRegis.NoMatch Then
            TbOfiRegis.Delete
        End If
        '--------------- meses
        If wf1numera = "*" Then
            tbmes.Index = "IDDOCMES"
            tbmes.Seek "=", "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00"), txtmesmov.Text
            If Not tbmes.NoMatch Then
                xtempo = Right(TxtNumMov.Text, 5)
                If tbmes.Fields("F2NUMMOV") = xtempo Then
                    tbmes.Edit
                    tbmes.Fields("F2NUMMOV") = Val(xtempo) - 1
                    cnomfile = tbmes.Name
                    tbmes.Update
                End If
            End If
        Else
            tbmes.Index = "IDNUMMES"
            tbmes.Seek "=", txtmesmov.Text
            If Not tbmes.NoMatch Then
                xtempo = TxtNumMov.Text
                If tbmes.Fields("F2NUMMOV") = xtempo Then
                    tbmes.Edit
                    tbmes.Fields("F2NUMMOV") = Val(xtempo) - 1
                    cnomfile = tbmes.Name
                    tbmes.Update
                End If
            End If
        End If
        '---------------
        nuevo
        '---------------
    End If

End Sub

Private Sub Elimina_Item()

    Beep
    If MsgBox("Está seguro de eliminar item", 36, "Atención") = 6 Then
       TbDetRegis.Index = "IDMESNUMITE"
       TbDetRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text, Grd_Docum.Columns(0)
       If Not TbDetRegis.NoMatch Then
          TbDetRegis.Delete
       End If
       DataFactura.Recordset.Delete
       DataFactura.Refresh
       Grd_Docum.Refresh

       swmodigv = True
       calcula

    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Form_Mes.Show 1
    End If
    Lblpol.Visible = False
    TxtPoliza.Visible = False
    
End Sub

Private Sub Form_Load()
Dim carchorigen     As String
Dim carchdestino    As String
Dim dbcntbanc       As DAO.Database
Dim tbcntbanc       As DAO.Recordset

    wnewnumero = True
    
    Set dbcntbanc = OpenDatabase(App.Path & "\cnt_banc.mdb")
    Set tbcntbanc = dbcntbanc.OpenRecordset("bf1cnt")
    tbcntbanc.Index = "bf1cnt"
    tbcntbanc.Seek "=", wempresa
    If Not tbcntbanc.NoMatch Then
        WF1INGCTA_CC = Trim(tbcntbanc.Fields("F1INGCTA_CC") & "")
    End If
    tbcntbanc.Close
    dbcntbanc.Close

    Set dbcntcont = OpenDatabase(wcontacnt & "\cnt_cont.mdb")
    Set tbcntcont = dbcntcont.OpenRecordset("cf1cnt")
    tbcntcont.Index = "cf1cnt"
    tbcntcont.Seek "=", wempresa
    If Not tbcntcont.NoMatch Then
        wcosto = Val(tbcntcont.Fields("f1costo") & "")
    End If
    '------------------------------------------------------------

    sw_numdoc = True

    meslet = Choose(Val(mes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
    Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno

    gcodppp = ""
    
    Set dbbanco = OpenDatabase(wrutabancos & "\db_tabla.mdb")
    Set tbtalon = dbbanco.OpenRecordset("bf7talon")

    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbfpagos = dbempresa.OpenRecordset("ef2forpag")
    tbfpagos.Index = "IDFORPAG"
    Set Tbproveedor = dbempresa.OpenRecordset("EF2PROVEEDORES")
    Tbproveedor.Index = "IDCODPROV"
    
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "IDMESNUM"
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
    TbOfiRegis.Index = "IDMESNUM"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    TbDetRegis.Index = "IDMESNUMITE"
    Set tbmes = dbcompras.OpenRecordset("MESES")
    tbmes.Index = "IDNUMMES"

    '------------------------------
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbcomtab = dbcomtabla.OpenRecordset("BF9GIN")
    Set tbdepa = dbcomtabla.OpenRecordset("departamentos")
    tbdepa.Index = "icodigo"
    Set tbcambios = dbcomtabla.OpenRecordset("cambios")
    Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
    tbgastos.Index = "IDCODIGO"
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "IDCODDOC"

    If Len(Trim(wrutaconta)) <> 0 Then
        Set dbconta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
        Set tbcf5pla = dbconta.OpenRecordset("cf5pla")
        Set tbcf5costo = dbconta.OpenRecordset("cf5costo")
        
        Set dbcentros = OpenDatabase(wrutaconta & "\db_costo.mdb")
        Set tbcentros = dbcentros.OpenRecordset("centros")
    End If
    '------------------------------

    TbDocumento.MoveFirst
    Do While Not TbDocumento.EOF
        CmbTipDoc.AddItem TbDocumento.Fields("F2DESDOC") + Space(100) + TbDocumento.Fields("F2CODDOC") + TbDocumento.Fields("F2ABREV")
        TbDocumento.MoveNext
    Loop
    CmbTipDoc.ListIndex = 0

    Set DbTempfac = OpenDatabase(wrutatemp & "\temp_com.mdb")
    DbTempfac.Execute ("Delete * From temp_det")
    Set Tabla = DbTempfac.OpenRecordset("temp_det")

    DataFactura.DatabaseName = wrutatemp & "\temp_com.Mdb"
    DataFactura.RecordSource = "select * from temp_det"
    DataFactura.Refresh
    Grd_Docum.Refresh
    'Nuevo_Item

    '-------------------------------------------------------------------
    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "idcodemp"
    tbparametro1.Seek "=", wempresa
    If Not tbparametro1.NoMatch Then
        If tbparametro1.Fields("f1mes" & mes) = "*" Then
            MsgBox "Mes está cerrado. Verifique.", 48, "Compras"
            Form_Mes.Show 1
            meslet = Choose(Val(mes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
            Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
        End If
    End If
    tbparametro1.Close
    dbcontrol.Close
    '-------------------------------------------------------------------
    'If wf1numera = "*" Then
    '    tbmes.Index = "IDDOCMES"
    '    TxtNumMov.Text = ""
        txtmesmov.Text = mes
    'Else
    '    tbmes.Index = "IDNUMMES"
    '    tbmes.Seek "=", MES
    '    If Not tbmes.NoMatch Then
    '       txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
    '       WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
    '       If Len(Trim(wnumobra)) <> 0 Then
    '            If Len(Trim(wnumobra)) = 2 Then
    '                WNUMERO = Format(wnumobra, "00") & Mid(Format(WNUMERO, "0000000"), 3, 5)
    '            End If
    '            If Len(Trim(wnumobra)) = 3 Then
    '                WNUMERO = Format(wnumobra, "000") & Mid(Format(WNUMERO, "0000000"), 4, 4)
    '            End If
    '       End If
    '       TxtNumMov.Text = Format(WNUMERO, "0000000")
    '    End If
    'End If
    If Month(Format(Date, "dd/mm/yyyy")) > Val(mes) Then
        txtfecha.Text = CDate(Format("01/" & mes & "/" & Year(Date), "dd/mm/yyyy"))
    Else
        txtfecha.Text = Format(Date, "dd/mm/yyyy")
    End If
    TxtFecVen.Text = Format(Date, "dd/mm/yyyy")
    txtfecharec.Text = Format(Date, "dd/mm/yyyy")

    cmbdocum.AddItem "Cheque"
    cmbdocum.AddItem "Efectivo"
    cmbdocum.AddItem "Letra"
    cmbdocum.AddItem "Otros"
    cmbdocum.ListIndex = 0

    cmbigv.AddItem "Destinadas a Ventas Gravadas Exclusivamente" & Space(50) & "001"
    cmbigv.AddItem "Destinadas a Ventas Gravadas y No Gravadas " & Space(50) & "002"
    cmbigv.AddItem "Destinadas a Ventas No Gravadas            " & Space(50) & "003"
    cmbigv.AddItem "Exonerados                                 " & Space(50) & "004"
    cmbigv.ListIndex = 0
    
    swmodigv = False
    swmodcta = False
    swtc = False
    Wnuevo = True
    swgraba_valing = False

    If wf1viscod = "*" Then
        lblcod.Visible = False
        Text1.Visible = False
    Else
        lblcod.Visible = True
        Text1.Visible = True
    End If

    '-------------------- llena formas de pago
    tbfpagos.MoveFirst
    Do While Not tbfpagos.EOF
        If Len(Trim(tbfpagos.Fields("f2tipo") & "")) > 0 Then
            If tbfpagos.Fields("f2tipo") & "" = "F" Or tbfpagos.Fields("f2tipo") & "" = "C" Then
                cmbfpagos.AddItem tbfpagos.Fields("f2despag") & Space(100) & tbfpagos.Fields("f2forpag") & tbfpagos.Fields("f2tipo") & ""
            End If
        End If
        tbfpagos.MoveNext
    Loop
    cmbfpagos.ListIndex = 0
    '--------------------
    
    DataFactura.Refresh
    Grd_Docum.Refresh
    Grd_Docum.col = 1
    'Grd_Docum.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    BtnExitO_Click
    If cflag = 1 Then Cancel = 1
    cflag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error_bd

    DataFactura.DatabaseName = ""
    DataFactura.RecordSource = ""
    DataFactura.Refresh

    '----------- DB_BANCOS.MDB
    tbcomtab.Close
    tbdepa.Close
    tbgastos.Close
    TbDocumento.Close
    tbcambios.Close
    dbcomtabla.Close
    '-----------

    '----------- temp_com.mdb
    DbTempfac.Execute ("Delete From temp_det")
    Tabla.Close
    DbTempfac.Close
    '-----------
    
    '----------- db_tabla.mdb
    tbtalon.Close
    dbbanco.Close

    '----------- DB_BANCOS.MDB
    tbfpagos.Close
    Tbproveedor.Close
    dbempresa.Close

    '----------- DB_BANCOS.MDB
    TbCabRegis.Close
    TbOfiRegis.Close
    TbDetRegis.Close
    tbmes.Close
    dbcompras.Close
    
    '----------- ctaspag.mdb
    If gctapag = "*" Then
        tbcabcta.Close
        tbdetcta.Close
        DbCtaPag.Close
    End If

    '----------- db_tabla.mdb
    '----------- db_costo.mdb
    If Len(Trim(wrutaconta)) <> 0 Then
        tbcf5pla.Close
        tbcf5costo.Close
        dbconta.Close
        tbcentros.Close
        dbcentros.Close
    End If

    Exit Sub

error_bd:
    Resume Next

End Sub

Private Sub graba_gridoc()
Dim DbInventa   As DAO.Database
Dim tbif5       As DAO.Recordset
Dim dbcomp      As DAO.Database
Dim tbf4oc      As DAO.Recordset
Dim tbf3oc      As DAO.Recordset
Dim dbtemp      As DAO.Database
Dim tbtemp      As DAO.Recordset

    'If Tabla.RecordCount = 0 Or Val(Format(Grd_Docum.Columns(7), "0.00")) = 0# Then
    '    swgraba_valing = True
    '    Set DBTEMP = OpenDatabase(wrutatemp & "\temp_com.mdb")
    '    DBTEMP.Execute ("delete * from temp_oc")
    '    Set tbtemp = DBTEMP.OpenRecordset("temp_oc")
    '    tbtemp.Index = "idgasto"

    '    Set DbInventa = OpenDatabase(ruta & "inventa.mdb")
    '    Set tbif5 = DbInventa.OpenRecordset("if5pla")
    '    tbif5.Index = "IdCodpro"

    '    Set dbcomp = OpenDatabase(ruta & "DB_BANCOS.MDB")
    '    Set tbf4oc = dbcomp.OpenRecordset("if4orden")
    '    tbf4oc.Index = "IDNUMORD"
    '    Set tbf3oc = dbcomp.OpenRecordset("if3orden")
    '    tbf3oc.Index = "IDNUMORD"
        
    '    If FrmAyudaOrdCom.datagrid.SelCount > 0 Then
    '        For i% = 0 To FrmAyudaOrdCom.datagrid.SelCount - 1
    '            FrmAyudaOrdCom.datagrid.Row = FrmAyudaOrdCom.datagrid.SelRow(i%)
    '
    '            tbf4oc.Seek "=", Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15))
    '            If Not tbf4oc.NoMatch Then
    '                xitem = 0
    '                tbf3oc.Seek "=", Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15))
    '                If Not tbf3oc.NoMatch Then
    '                    Do While tbf3oc.Fields("F4NUMORD") = Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15)) And Not tbf3oc.EOF
    '                        '---------------------------------------------------
    '                        xgasto = "": xnomgasto = "": xctacont = ""
    '                        tbif5.Seek "=", tbf3oc.Fields("f3codpro") & ""
    '                        If Not tbif5.NoMatch Then
    '                            xgasto = tbif5.Fields("f3gasto") & ""
    '                        End If
    '                        tbgastos.Index = "bf9gin"
    '                        tbgastos.Seek "=", "G", xgasto
    '                        If Not tbgastos.NoMatch Then
    '                            xnomgasto = Trim(tbgastos.Fields("nombre") & "")
    '                            xctacont = Trim(tbgastos.Fields("cuenta") & "")
    '                        End If
    '                        '---------------------------------------------------
    '                        tbtemp.Seek "=", xgasto
    '                        If Not tbtemp.NoMatch Then
    '                            tbtemp.Edit
    '                        Else
    '                            xitem = xitem + 1
    '                            tbtemp.AddNew
    '                            tbtemp.Fields("f3item") = xitem
    '                        End If
    '                        tbtemp.Fields("f3anoorden") = wanno
    '                        tbtemp.Fields("f3gasto") = xgasto
    '                        tbtemp.Fields("f3ctacon") = xctacont
    '                        tbtemp.Fields("f3concepto") = xnomgasto
    '                        If OpcSol.value = True Then
    '                            tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
    '                        Else
    '                            tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
    '                        End If
    '                        tbtemp.Fields("f3afecto") = "*"
    '                        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
    '                            tbtemp.Fields("F3DebHab") = "H"
    '                        Else
    '                            tbtemp.Fields("F3DebHab") = "D"
    '                        End If
    '                        tbtemp.Fields("f3cencos") = Trim(tbf4oc.Fields("f4centro") & "")
    '                        tbtemp.Update
    '                        tbf3oc.MoveNext
    '                        If tbf3oc.EOF Then Exit Do
    '                        If tbf3oc.Fields("F4NUMORD") <> Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15)) Then Exit Do
    '                    Loop
    '                End If
    '            End If
    '        Next i%
    '        '-----------------------------
    '        If tbtemp.RecordCount > 0 Then
    '            tbtemp.MoveFirst
    '            DBTEMPFAC.Execute ("Delete From temp_det")
    '            Do While Not tbtemp.EOF
    '                Tabla.AddNew
    '                Tabla.Fields("f3item") = tbtemp.Fields("f3item")
    '                Tabla.Fields("f3anoorden") = tbtemp.Fields("f3anoorden")
    '                Tabla.Fields("f3gasto") = tbtemp.Fields("f3gasto") & ""
    '                Tabla.Fields("f3ctacon") = tbtemp.Fields("f3ctacon") & ""
    '                Tabla.Fields("f3concepto") = tbtemp.Fields("f3concepto") & ""
    '                Tabla.Fields("f3importe") = tbtemp.Fields("f3importe")
    '                Tabla.Fields("f3afecto") = tbtemp.Fields("f3afecto") & ""
    '                Tabla.Fields("f3debhab") = tbtemp.Fields("f3debhab") & ""
    '                Tabla.Fields("f3cencos") = tbtemp.Fields("f3cencos") & ""
    '                Tabla.Update
    '                tbtemp.MoveNext
    '            Loop
    '            DataFactura.Refresh
    '            swmodigv = True
    '            swmodcta = True
    '            calcula
    '        End If
    '        '-----------------------------
    '    End If
    '
    '    tbf3oc.Close
    '    tbf4oc.Close
    '    dbcomp.Close

    '    tbif5.Close
    '    DbInventa.Close

    '    tbtemp.Close
    '    DBTEMP.Close

    'End If

End Sub

Private Sub grabar()
Dim FlagAdd         As Integer
Dim intcont         As Integer
Dim xtempo          As String
Dim dbbase          As DAO.Database
Dim TbCabOrden      As DAO.Recordset
Dim tbtrimestres    As DAO.Recordset
Dim xtrimestre      As Integer
Dim nmtonse         As Double
Dim tbf4oc          As DAO.Recordset
Dim sw_newreg       As Integer
Dim tbcabaux_regis  As DAO.Recordset
Dim numero          As Integer
Dim dbcomp          As DAO.Database
Dim cnomfile        As String
Dim xduplica        As Integer
Dim X               As Double
Dim xcorre          As Double
Dim xvar            As Integer

On Error GoTo Errores

    If wnewnumero = True Then
        wnewnumero = False
        If wf1numera = "*" Then
            tbmes.Index = "IDDOCMES"
            TxtNumMov.Text = ""
            txtmesmov.Text = mes
        Else
            tbmes.Index = "IDNUMMES"
            tbmes.Seek "=", mes
            If Not tbmes.NoMatch Then
                txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
                WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
                TxtNumMov.Text = Format(WNUMERO, "0000000")
            End If
        End If
    End If

    Set tbcabaux_regis = dbcompras.OpenRecordset("REGISDOC")
    tbcabaux_regis.Index = "IDMESNUM"
    
    sw_grabanse = False
    sw_newreg = False

    Me.MousePointer = 11
    If Trim(TxtCodPrv.Text) = "" Then
        MsgBox "Error: Código de Proveedor no válido...", 16, "Atención"
        TxtSerDoc.SetFocus
        TxtCodPrv.SetFocus
        TxtCodPrv_KeyUp 113, 0
        Exit Sub
    End If

    If Trim(TxtNumDoc.Text) = "" Then
        MsgBox "Error: Número de Documento no válido...", 16, "Atención"
        cflag = 1
        TxtNumDoc.SetFocus
        Exit Sub
    End If

    If Val(TxtTipCam.Text) = 0 Then '
        MsgBox "Error: El Tipo de Cambio no puede ser 0 ...", 16, "Atención"
        TxtTipCam.SetFocus
        Exit Sub
    End If
    '-------------------------------------------------------------------
    xmonto_anterior = 0#
    xflag = False
    TbCabRegis.Index = "IDMESNUM"
    TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
    If Not TbCabRegis.NoMatch Then
        If gctapag = "*" Then
            '------------- NSE
            Set DbCtaPag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
            Set tbcabcta = DbCtaPag.OpenRecordset("PAG_DCTO")
            tbcabcta.Index = "nro_corre"
            tbcabcta.Seek "=", TbCabRegis.Fields("f4correla")
            If Not tbcabcta.NoMatch Then
                If Val(Format(tbcabcta.Fields("SALDO"), "0.00")) <> Val(Format(tbcabcta.Fields("TOTALO"), "0.00")) Then
                    MsgBox "Verifique el documento ya ha sido aplicado", 48, "Compras"
                    xflag = True
                End If
            End If
            tbcabcta.Close
            DbCtaPag.Close
        End If
        
        FlagAdd = 0
        If windigraba = False Then
            xmonto_anterior = Val("" & TbCabRegis.Fields("f4total"))
            TbCabRegis.Edit
            TbCabRegis.Fields("F4FECHMOD") = Format(Date, "dd/mm/yyyy")
            TbCabRegis.Fields("F4USUARIOMOD") = wusuario
            TbCabRegis.Fields("F4HORAMOD") = Time
        Else
            '------------------------------- OBTIENE EL NUMERO DEL MOV.
            If wf1numera = "*" Then
                tbmes.Index = "IDDOCMES"
                tbmes.Seek "=", "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00"), mes
                If Not tbmes.NoMatch Then
                   txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
                   WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
                   TxtNumMov.Text = "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00") & Right(Format(WNUMERO, "0000000"), 5)
                End If
            Else
                tbmes.Index = "IDNUMMES"
                tbmes.Seek "=", mes
                If Not tbmes.NoMatch Then
                   txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
                   WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
                   TxtNumMov.Text = Format(WNUMERO, "0000000")
                End If
            End If
            '-------------------------------
            xduplica = True
            Do While xduplica = True
                TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
                If Not TbCabRegis.NoMatch Then
                    TxtNumMov.Text = Format(Val(TxtNumMov.Text) + 1, "0000000")
                Else
                    xduplica = False
                End If
            Loop
            sw_grabanse = True
            sw_newreg = True
            TbCabRegis.AddNew
            TbCabRegis.Fields("F4FECHING") = Format(Date, "dd/mm/yyyy")
            TbCabRegis.Fields("F4USUARIOING") = wusuario
            TbCabRegis.Fields("F4HORAING") = Time
            FlagAdd = 1
            windigraba = False
        End If
    Else
        '-------------------------------
        xduplica = True
        Do While xduplica = True
            TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
            If Not TbCabRegis.NoMatch Then
                TxtNumMov.Text = Format(Val(TxtNumMov.Text) + 1, "0000000")
            Else
                xduplica = False
            End If
        Loop
        '-------------------------------
        sw_grabanse = True
        sw_newreg = True
        TbCabRegis.AddNew
        TbCabRegis.Fields("F4FECHING") = Format(Date, "dd/mm/yyyy")
        TbCabRegis.Fields("F4USUARIOING") = wusuario
        TbCabRegis.Fields("F4HORAING") = Time
        FlagAdd = 1
        windigraba = False
    End If

    gcorre = 0
    '------------------------- graba en ctas x pagar
    If wf1traslado <> "*" Then  '- normal, despues de grabar en compras, lo traslada a ctasxpagar
        If gctapag = "*" And Right(cmbfpagos.Text, 1) = "F" And xflag = False Then
        
            Set DbCtaPag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
            Set tbcabcta = DbCtaPag.OpenRecordset("PAG_DCTO")
            tbcabcta.Index = "nro_corre"
            tbcabcta.Seek "=", Val("" & TbCabRegis.Fields("f4correla"))
            If Not tbcabcta.NoMatch Then
                If Val(TbCabRegis.Fields("f4correla") & "") = 0 Then
                    xvar = False
                Else
                    xvar = True
                    xcorre = tbcabcta.Fields("correla")
                End If
            Else
                xvar = False
            End If
            tbcabcta.Close
            DbCtaPag.Close
            
            If xvar = False Then
                trans_ctasxpagar "N", Me, 0
            Else
                trans_ctasxpagar "M", Me, xcorre
            End If
                
            TbCabRegis.Fields("F4CORRELA") = gcorre
            
        End If
    End If
    '---------------------------------------------------------------
    TbCabRegis.Fields("F4DIARIO") = Trim(txtref.Text & "")
    TbCabRegis.Fields("F4CODIGV") = Right(cmbigv.Text, 3)
    '-----------------------------------------------------------------------
    TbCabRegis.Fields("F4MESMOV") = "" & txtmesmov.Text
    '-----------------------------------------------------------------------
    TbCabRegis.Fields("F4CODPRV") = "" & TxtCodPrv.Text
    TbCabRegis.Fields("F4NOMPRV") = "" & Left(TxtNomPrv.Text, 100)
    TbCabRegis.Fields("F4DIRPRV") = "" & Left(Trim(TxtDirPrv.Text), 100)
    TbCabRegis.Fields("F4RUCPRV") = "" & TxtRucPrv.Text
    TbCabRegis.Fields("F4TIPDOC") = "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00")
    TbCabRegis.Fields("F4FECHA") = Format(txtfecha.Text, "DD/MM/YYYY")
    TbCabRegis.Fields("F4FORPAG") = Left(Right(cmbfpagos.Text, 4), 3)
    TbCabRegis.Fields("F4MONEDA") = IIf(OpcSol.Value = True, "S", "D")
    TbCabRegis.Fields("F4FECVEN") = Format(TxtFecVen.Text, "DD/MM/YYYY")
    TbCabRegis.Fields("F4TIPCAM") = 0 + TxtTipCam.Text
    TbCabRegis.Fields("F4REFERE") = "" & Trim(TxtRefere.Text)
    TbCabRegis.Fields("F4SERDOC") = "" & TxtSerDoc.Text
    TbCabRegis.Fields("F4NUMDOC") = "" & TxtNumDoc.Text
    TbCabRegis.Fields("F4MONINA") = Val(Format(PnlMonIna(0).Caption, "#0.00"))
    TbCabRegis.Fields("F4OTRIMP") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
    TbCabRegis.Fields("F4CTACONT") = "" & Trim(TxtTelPrv.Text)
    TbCabRegis.Fields("F4CTACONT") = "" & Left(Trim(TxtTelPrv.Text), 12)
    TbCabRegis.Fields("F4FECHAREC") = Format(txtfecharec.Text, "DD/MM/YYYY")
    'If wingcta = "*" Then
        TbCabRegis.Fields("F4GRUPO") = "" & Text1.Text
    'End If
    If wocompra = "*" Then
        If wf1tipdoc_asoc = "V" Then
            TbCabRegis.Fields("F4VALESING") = txtocompra.Text
        Else
            TbCabRegis.Fields("F4OCOMPRA") = Val("" & txtocompra.Text)
        End If
        TbCabRegis.Fields("F4TIPORDEN") = Left(gorden_cs, 3)
    End If
    TbCabRegis.Fields("F4GASTGEN") = IIf(chkgg.Value = True, "*", " ")
    If wingobra = "*" Then
       TbCabRegis.Fields("F4OBRA") = "" & gobra
    End If
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
        TbCabRegis.Fields("F4BASIMP") = Val(Format(PnlBasImp(0).Caption, "#0.00"))
        TbCabRegis.Fields("F4IGV") = Val(Format(txtigv(0).Text, "#0.00"))
        TbCabRegis.Fields("F4TOTAL") = Val(Format(PnlTotal(0).Caption, "#0.00"))
    Else
        TbCabRegis.Fields("F4BASIMP") = Val(Format(PnlBasImp(0).Caption, "#0.00")) * -1
        TbCabRegis.Fields("F4IGV") = Val(Format(txtigv(0).Text, "#0.00")) * -1
        TbCabRegis.Fields("F4TOTAL") = Val(Format(PnlTotal(0).Caption, "#0.00")) * -1
    End If

    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
        TbCabRegis.Fields("F4MONTORET") = Val(Format(txtigv(0).Text, "#0.00"))
        TbCabRegis.Fields("F4FONAVI") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
        TbCabRegis.Fields("F4IGV") = 0#
        TbCabRegis.Fields("F4OTRIMP") = 0#
    End If
    TbCabRegis.Fields("f4redsuma") = Val(Format(txtredsuma.Text, "#0.00"))
    TbCabRegis.Fields("f4redresta") = Val(Format(txtredresta.Text, "#0.00"))
    TbCabRegis.Fields("f4dcto") = Val(Format(txtdcto.Text, "#0.00"))
    cnomfile = TbCabRegis.Name
    '-----------------------------------------------------------------------
    If sw_newreg = True Then
        xduplica = True
        Do While xduplica = True
            tbcabaux_regis.Seek "=", txtmesmov.Text, TxtNumMov.Text
            If Not tbcabaux_regis.NoMatch Then
                TxtNumMov.Text = Format(Val(TxtNumMov.Text) + 1, "0000000")
            Else
                xduplica = False
            End If
        Loop
    End If
    TbCabRegis.Fields("F4NUMMOV") = "" & TxtNumMov.Text
    '-----------------------------------------------------------------------
    TbCabRegis.Update

    sw_newreg = False

    If Wnuevo = True Then
        wvcodprov = "" & TxtCodPrv.Text
        wvtipdocu = "" & Format(Left(Right(CmbTipDoc.Text, 5), 2), "00")
        wvserdoc = "" & TxtSerDoc.Text
        wvnumdoc = "" & TxtNumDoc.Text
    End If
    TbOfiRegis.Index = "IDMESNUM"
    TbOfiRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
    If Not TbOfiRegis.NoMatch Then
        TbOfiRegis.Edit
    Else
        TbOfiRegis.AddNew
    End If
    TbOfiRegis.Fields("F4CODIGV") = Right(cmbigv.Text, 3)
    TbOfiRegis.Fields("F4MESMOV") = "" & txtmesmov.Text
    TbOfiRegis.Fields("F4NUMMOV") = "" & TxtNumMov.Text
    TbOfiRegis.Fields("F4CODPRV") = "" & TxtCodPrv.Text
    TbOfiRegis.Fields("F4NOMPRV") = "" & Left(TxtNomPrv.Text, 100)
    TbOfiRegis.Fields("F4DIRPRV") = "" & Left(Trim(TxtDirPrv.Text), 100)
    TbOfiRegis.Fields("F4RUCPRV") = "" & TxtRucPrv.Text
    TbOfiRegis.Fields("F4TIPDOC") = "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00")
    TbOfiRegis.Fields("F4FECHA") = Format(txtfecha.Text, "DD/MM/YYYY")
    TbOfiRegis.Fields("F4MONEDA") = "S"
    TbOfiRegis.Fields("F4TIPCAM") = 0 + TxtTipCam.Text
    TbOfiRegis.Fields("F4REFERE") = "" & Trim(TxtRefere.Text)
    TbOfiRegis.Fields("F4SERDOC") = "" & TxtSerDoc.Text
    TbOfiRegis.Fields("F4NUMDOC") = "" & TxtNumDoc.Text
    TbOfiRegis.Fields("F4OTRIMP") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
    If OpcSol.Value = True Then
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            TbOfiRegis.Fields("F4FONAVI") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
            TbOfiRegis.Fields("F4OTRIMP") = 0#
        End If
    Else
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            TbOfiRegis.Fields("F4FONAVI") = Format(Val(Format(TxtOtrImp(0).Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            TbOfiRegis.Fields("F4OTRIMP") = 0#
        End If
    End If
    Rem NSE TbOfiRegis.Fields("F4MONINA") = Val(Format(PnlMonIna(1).Caption, "#0.00"))
    If OpcSol.Value = True Then
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                TbOfiRegis.Fields("F4MONTORET") = Val(Format(txtigv(0).Text, "#0.00"))
                TbOfiRegis.Fields("F4IGV") = 0#
            Else
                TbOfiRegis.Fields("F4IGV") = Val(Format(txtigv(0).Text, "#0.00"))
            End If
        Else
            TbOfiRegis.Fields("F4IGV") = Val(Format(txtigv(0).Text, "#0.00")) * -1
        End If
    End If
    If OpcDol.Value = True Then
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                TbOfiRegis.Fields("F4IGV") = 0#
                TbOfiRegis.Fields("F4MONTORET") = Format(Val(Format(txtigv(0).Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            Else
                TbOfiRegis.Fields("F4IGV") = Format(Val(Format(txtigv(0).Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            End If
        Else
            If OpcSol.Value = True Then
                TbOfiRegis.Fields("F4IGV") = Format(txtigv(0).Text, "#0.00") * -1
            Else
                TbOfiRegis.Fields("F4IGV") = Format(Val(Format(txtigv(0).Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")) * -1, "#0.00")
            End If
        End If
    End If

    If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
        If OpcSol.Value = True Then
            TbOfiRegis.Fields("F4MONINA") = Val(Format(PnlMonIna(0).Caption, "0.00"))
            TbOfiRegis.Fields("f4dcto") = Format(txtdcto.Text, "#0.00")
            TbOfiRegis.Fields("f4redsuma") = Format(txtredsuma.Text, "#0.00")
            TbOfiRegis.Fields("f4redresta") = Format(txtredresta.Text, "#0.00")
            TbOfiRegis.Fields("F4BASIMP") = Format(PnlBasImp(0).Caption, "#0.00")
            TbOfiRegis.Fields("F4TOTAL") = Format(PnlTotal(0).Caption, "#0.00")
        Else
            TbOfiRegis.Fields("F4MONINA") = Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "0.00")
            TbOfiRegis.Fields("f4redsuma") = Format(Val(Format(txtredsuma.Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            TbOfiRegis.Fields("f4redresta") = Format(Val(Format(txtredresta.Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            TbOfiRegis.Fields("F4BASIMP") = Format(Val(Format(PnlBasImp(0).Caption, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00")
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                TbOfiRegis.Fields("F4TOTAL") = Format(Val(Format(Val(Format(PnlBasImp(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(TxtOtrImp(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(txtigv(0).Text, "0.00")) * TxtTipCam.Text, "0.00")), "0.00")
            Else
                TbOfiRegis.Fields("F4TOTAL") = Format(Val(Format(Val(Format(PnlBasImp(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(TxtOtrImp(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(txtigv(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(txtredsuma.Text, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(txtredresta.Text, "0.00")) * TxtTipCam.Text, "0.00")), "0.00")
            End If
        End If
    Else
        If OpcSol.Value = True Then
            TbOfiRegis.Fields("F4MONINA") = Val(Format(PnlMonIna(0).Caption, "0.00")) * -1
            TbOfiRegis.Fields("f4redsuma") = Format(txtredsuma.Text, "#0.00") * -1
            TbOfiRegis.Fields("f4redresta") = Format(txtredresta.Text, "#0.00") * -1
            TbOfiRegis.Fields("F4BASIMP") = Format(PnlBasImp(0).Caption, "#0.00") * -1
            TbOfiRegis.Fields("F4TOTAL") = Format(PnlTotal(0).Caption, "#0.00") * -1
        Else
            TbOfiRegis.Fields("F4MONINA") = Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "0.00") * -1
            TbOfiRegis.Fields("f4redsuma") = Format(Val(Format(txtredsuma.Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00") * -1
            TbOfiRegis.Fields("f4redresta") = Format(Val(Format(txtredresta.Text, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00") * -1
            TbOfiRegis.Fields("F4BASIMP") = Format(Val(Format(PnlBasImp(0).Caption, "#0.00")) * Val(Format(TxtTipCam.Text, "#0.000")), "#0.00") * -1
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                TbOfiRegis.Fields("F4TOTAL") = Format(Val(Format(Val(Format(PnlBasImp(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(TxtOtrImp(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(txtigv(0).Text, "0.00")) * TxtTipCam.Text, "0.00")), "0.00")
            Else
                TbOfiRegis.Fields("F4TOTAL") = Format(Val(Format(Val(Format(PnlBasImp(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(PnlMonIna(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(TxtOtrImp(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(txtigv(0).Text, "0.00")) * TxtTipCam.Text, "0.00")) + Val(Format(Val(Format(txtredsuma.Text, "0.00")) * TxtTipCam.Text, "0.00")) - Val(Format(Val(Format(txtredresta.Text, "0.00")) * TxtTipCam.Text, "0.00")), "0.00") * -1
            End If
        End If
    End If
    cnomfile = TbOfiRegis.Name
    TbOfiRegis.Update

    TbDetRegis.Index = "IDMESNUMITE"
    intcont = 0
    
    Tabla.MoveFirst
    Do While Not Tabla.EOF
        TbDetRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text, Tabla.Fields("F3ITEM")
        If Not TbDetRegis.NoMatch Then
            TbDetRegis.Edit
        Else
            TbDetRegis.AddNew
        End If
        TbDetRegis.Fields("F4MESMOV") = "" & txtmesmov.Text
        TbDetRegis.Fields("F4NUMMOV") = "" & TxtNumMov.Text
        TbDetRegis.Fields("F3ITEM") = Tabla.Fields("F3ITEM")
        TbDetRegis.Fields("F3GASTO") = UCase(Trim("" & Tabla.Fields("F3GASTO")))
        TbDetRegis.Fields("F3CTACON") = Trim("" & Tabla.Fields("F3CTACON"))
        TbDetRegis.Fields("F3CENCOS") = Trim("" & Tabla.Fields("F3CENCOS"))
        TbDetRegis.Fields("F3CONCEPTO") = Trim("" & Tabla.Fields("F3CONCEPTO"))
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
            TbDetRegis.Fields("F3IMPORTE") = Val(Format(Tabla.Fields("F3IMPORTE"), "0.00"))
        Else
            TbDetRegis.Fields("F3IMPORTE") = Val(Format(Tabla.Fields("F3IMPORTE"), "0.00")) * -1
        End If
        TbDetRegis.Fields("F3IMPORTE") = Val(Format(Tabla.Fields("F3IMPORTE"), "0.00"))
        TbDetRegis.Fields("F3AFECTO") = Trim("" & Tabla.Fields("F3AFECTO"))
        TbDetRegis.Fields("F3DEBHAB") = Trim("" & Tabla.Fields("F3DEBHAB"))
        TbDetRegis.Fields("F3ANOORDEN") = Trim("" & Tabla.Fields("F3ANOORDEN"))
        TbDetRegis.Fields("F3ORDEN") = Trim("" & Tabla.Fields("F3ORDEN"))
        TbDetRegis.Fields("F3DEPARTAMENTO") = Trim("" & Tabla.Fields("F3DEPARTAMENTO"))
        TbDetRegis.Fields("F3CTADESTINO") = Trim("" & Tabla.Fields("F3CTADESTINO"))
        cnomfile = TbDetRegis.Name
        TbDetRegis.Update
        Tabla.MoveNext
    Loop

    '-----------------------------
    If Left(gorden_cs & "", 3) = "O/C" Or Left(gorden_cs & "", 3) = "O/S" Then
        Set dbbase = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        If Left(gorden_cs & "", 3) = "O/C" Then
            Set TbCabOrden = dbbase.OpenRecordset("if4orden")
        Else
            Set TbCabOrden = dbbase.OpenRecordset("if4servicio")
        End If
        TbCabOrden.Index = "IDNUMORD"
        TbCabOrden.Seek "=", txtocompra.Text
        If Not TbCabOrden.NoMatch Then
            TbCabOrden.Edit
            TbCabOrden.Fields("f4regcom") = txtmesmov.Text & TxtNumMov.Text
            cnomfile = TbCabOrden.Name
            TbCabOrden.Update
        End If
        TbCabOrden.Close
        dbbase.Close
    End If
    '-----------------------------
    
    If FlagAdd = 1 Then
        If wf1numera = "*" Then
            WNUMERO = TxtNumMov.Text
            tbmes.Index = "IDDOCMES"
            tbmes.Seek "=", "" & Format(Left(Trim(Right(Trim(CmbTipDoc.Text), 5)), 2), "00"), txtmesmov.Text
            If Not tbmes.NoMatch Then
                xtempo = Right(TxtNumMov.Text, 5)
                If Val(xtempo) > Val(tbmes.Fields("F2NUMMOV")) Then
                    tbmes.Edit
                    tbmes.Fields("F2NUMMOV") = xtempo
                    cnomfile = tbmes.Name
                    tbmes.Update
                End If
            End If
        Else
            WNUMERO = TxtNumMov.Text
            tbmes.Index = "IDNUMMES"
            tbmes.Seek "=", txtmesmov.Text
            If Not tbmes.NoMatch Then
                xtempo = TxtNumMov.Text
                If Val(xtempo) > Val(tbmes.Fields("F2NUMMOV")) Then
                    tbmes.Edit
                    tbmes.Fields("F2NUMMOV") = xtempo
                    cnomfile = tbmes.Name
                    tbmes.Update
                End If
            End If
        End If
    End If
    
    Me.MousePointer = 1
    WGRABO = 1
    '------------------------------------------
    
    If wf1trasoc = "*" Then
        Set dbcomp = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        If Left(gorden_cs, 3) = "O/C" Then
            Set tbf4oc = dbcomp.OpenRecordset("if4orden")
        Else
            Set tbf4oc = dbcomp.OpenRecordset("if4servicio")
        End If
        tbf4oc.Index = "IDNUMORD"
        
        tbf4oc.Seek "=", Val(txtocompra.Text & "")
        If Not tbf4oc.NoMatch Then
                tbf4oc.Edit
                tbf4oc.Fields("f4APLICADO") = Val(Format(PnlTotal(0).Caption, "0.00"))
                cnomfile = tbf4oc.Name
                tbf4oc.Update
            If Abs(xtotaloc - Val(Format(PnlTotal(0).Caption, "0.00"))) <= 0.2 Then
                tbf4oc.Edit
                tbf4oc.Fields("f4codtemp") = "1"
                tbf4oc.Fields("F4PARTIDACONTA") = txtmesmov.Text & TxtNumMov.Text
                cnomfile = tbf4oc.Name
                tbf4oc.Update
            End If
        End If
        
        tbf4oc.Close
        dbcomp.Close
    End If
    
    '------------------------------------------
    tbcabaux_regis.Close

    
    Rem NSE actualiza_provisiones

    Rem NSE EVALUA_CONTA txtmesmov.Text, TxtNumMov.Text

    Exit Sub

Errores:
    Select Case Err
        Case 3186:
            If MsgBox("La base de datos está bloqueada por otro usuario. Tabla : " & cnomfile & " ¿Reintentar?", vbYesNo) = vbYes Then
                For X = 1 To 10000
                Next
                Resume
            Else
                If MsgBox("No se grabarán los cambios efectuados. ¿ Continuar ?", vbYesNo) = vbYes Then
                    Exit Sub
                Else
                    Resume
                End If
            End If
        Case 3163:
            MsgBox Error(Err) & " Verifique la base de datos."
            Resume Next
    End Select
          
    'If sw_newreg = True Then
    '    xduplica = True
    '    Do While xduplica = True
    '        tbcabaux_regis.Seek "=", txtmesmov.Text, TxtNumMov.Text
    '        If Not tbcabaux_regis.NoMatch Then
    '            TxtNumMov.Text = Format(Val(TxtNumMov.Text) + 1, "0000000")
    '        Else
    '            xduplica = False
    '        End If
    '    Loop
    '    TbCabRegis.Fields("F4NUMMOV") = "" & TxtNumMov.Text
    'End If
    
    'NUMERO = NUMERO + 1
    'If NUMERO < 4 Then
    '    Resume
    'Else
    '    If MsgBox("Se ha producido el sgte. error : " & Error(Err) & "¿Desea reintentar?", 36, "Compras") = 6 Then
    '        NUMERO = 0
    '        Resume
    '    Else
    '        Exit Sub
    '    End If
    'End If
    
End Sub

Private Sub Grd_Docum_AfterColEdit(ByVal ColIndex As Integer)
     
    If Len(Trim(TxtCodPrv.Text)) > 0 Then
        Select Case ColIndex
            Case 3:
                If Valida_Gasto() = True Then
                    If winddepa = "*" Then
                        txtdepa.Visible = True
                        txtdepa.Text = Grd_Docum.Columns(10) & ""
                        txtdepa.SetFocus
                    End If
                    If Len(Trim(Grd_Docum.Columns(3))) > 0 Then
                    End If
                Else
                    Grd_Docum.Columns(3) = ""
                    Grd_Docum.col = 3
                    MsgBox "El código del gasto no existe. Verifique.", 48
                End If
            Case 4:
                swmodcta = True
                Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
                tbgastos.Index = "IDCUENTA"
                tbgastos.Seek "=", "G", Trim("" & Grd_Docum.Columns(4))
                If Not tbgastos.NoMatch Then
                    Grd_Docum.Columns(3) = "" & tbgastos.Fields("CODIGO")
                End If
                tbgastos.Close
            Case 5:
                If val_cc(Grd_Docum.Columns(4) & "") = True Then
                    If Len(Trim(Grd_Docum.Columns(5))) = 0 Then
                        MsgBox "Falta ingresar el centro de costo.", 48, "Compras"
                        Grd_Docum.col = 5
                    Else
                        If valida_centros() Then
                            '------------- Agropecuaria
                            If WF1INGCTA_CC = "*" Then
                                If Len(Trim(pctacen)) > 0 Then
                                    Grd_Docum.Columns(4) = pctacen & Mid(Trim(Grd_Docum.Columns(4)), 3, 10)
                                End If
                            End If
                            '-------------------------
                        Else
                            MsgBox "El centro de costo no existe. Verifique", 48, "Compras"
                            Grd_Docum.Columns(5) = ""
                            Grd_Docum.col = 5
                        End If
                    End If
                End If
            Case 7:
                swmodigv = True
        End Select
    End If
    
End Sub

Private Sub Grd_Docum_Change()

    If Grd_Docum.col = 5 Or Grd_Docum.col = 3 Then
       ' WCHANGE = 3
    End If
    If Grd_Docum.col = 4 Then
        swmodcta = True
    End If

End Sub

Private Sub Grd_Docum_Dblclick()

    Grd_Docum_KeyDown 113, 0

End Sub

Private Sub Grd_Docum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wbasimp As Double, wmonina As Double, wIMPUESTO As Double, wotrimp As Double
Dim xvar As String

    '--------------- Tecla DOWN Para Aumentar Un Item -------------------------
    If KeyCode = 13 Then
        Select Case Grd_Docum.col
            Case 4
                If Len(Trim(Grd_Docum.Columns(4))) = 0 And DataFactura.Recordset.RecordCount > 0 Then
                    If MsgBox("Desea Terminar ?", 36, "Información") = 6 Then
                        swigv = 1
                        txtigv(0).SetFocus
                    Else
                        TxtOtrImp(0).SetFocus
                        Grd_Docum.col = 3
                        Rem NSE Grd_Docum.Row = Grd_Docum.Row
                        Grd_Docum.SetFocus
                    End If
                Else
                    'txtfocus.SetFocus
                    If Len(Trim(wrutaconta)) <> 0 Then
                        If Valida_Contable(Grd_Docum.Columns(4) & "") = True Then
                            If tbcf5pla.Fields("F5CC") = True Then
                                'Grd_Docum.Col = 5 ' : Grd_Docum.SetFocus
                            Else
                                Grd_Docum.col = 6: Grd_Docum.SetFocus
                            End If
                            If swmodcta = False Or Len(Trim(Grd_Docum.Columns(6) & "")) = 0 Then
                                Grd_Docum.Columns(6) = "" & tbcf5pla.Fields("F5NOMCTA")
                            End If
                        Else
                            Grd_Docum.col = 4: Grd_Docum.SetFocus
                            MsgBox "La cuenta contable no existe. Verifique.", 48
                        End If
                    Else
                        Grd_Docum.col = 5: Grd_Docum.SetFocus
                    End If
                End If
        End Select
    End If
    If KeyCode = 117 Then
        'grd_docum.Col = 1: grd_docum.SetFocus
        'Nuevo_Item
    End If
    
    If KeyCode = 115 Then Elimina_Item
    
    If KeyCode = 113 Then
        Select Case Grd_Docum.col
            Case 3:
                GCODGAS = Grd_Docum.Columns(3)
                frmhlpgastos.Show 1
                wctacble = 1
                Grd_Docum.Columns(3) = GCODGAS
                Grd_Docum.Columns(4) = gcodcon
                Grd_Docum.Columns(6) = gnomgas
                gcodcon = ""
                gnomgas = ""
                Grd_Docum_KeyPress 13
            Case 4:
                gcodcon = Grd_Docum.Columns(4)
                FrmAyudaCon.Show 1
                Grd_Docum.Columns(4) = gcodcon
                Rem Grd_Docum.ColText(6) = gnomcon
                Grd_Docum_KeyPress 13
            Case 5:
                gcodcen = Grd_Docum.Columns(5)
                frmayucen.Show 1
                Grd_Docum.Columns(5) = gcodcen
                Grd_Docum_KeyPress 13
        End Select
    End If
       
End Sub

Private Sub Grd_Docum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Select Case Grd_Docum.col
'            Case 4:
'                If Len(Trim(Grd_Docum.Columns(4))) = 0 And DataFactura.Recordset.RecordCount > 0 Then
'                    If MsgBox("Desea Terminar ?", 36, "Información") = 6 Then
'                        Rem NSE TbDetRegis.Index = "IDMESNUMITE"
'                        Rem NSE TbDetRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text, Grd_Docum.Columns(0)
'                        Rem NSE If Not TbDetRegis.NoMatch Then
'                        Rem NSE     TbDetRegis.Delete
'                        Rem NSE End If
'                        Rem NSE DataFactura.Recordset.Delete
'                        Rem NSE DataFactura.Refresh
'                        swigv = 1
'                        txtigv(0).SetFocus
'                    Else
'                        txtotrimp(0).SetFocus
'                        Grd_Docum.col = 3
'                        'Grd_Docum.Row = Grd_Docum.Row
'                        Grd_Docum.SetFocus
'                    End If
'                Else
'                    Rem NSE txtfocus.SetFocus
'                    If Len(Trim(wrutaconta)) <> 0 Then
'                        If Valida_Contable(Grd_Docum.Columns(4) & "") = True Then
'                            If tbcf5pla.Fields("F5CC") = True Then
'                                Grd_Docum.col = 5: Grd_Docum.SetFocus
'                            Else
'                                Grd_Docum.col = 6: Grd_Docum.SetFocus
'                            End If
'                            If swmodcta = False Or Len(Trim(Grd_Docum.Columns(6) & "")) = 0 Then
'                                Grd_Docum.Columns(6) = "" & tbcf5pla.Fields("F5NOMCTA")
'                            End If
'                        Else
'                            Grd_Docum.col = 4: Grd_Docum.SetFocus
'                            MsgBox "La cuenta contable no existe. Verifique.", 48
'                        End If
'                    Else
'                        Grd_Docum.col = 5: Grd_Docum.SetFocus
'                    End If
'                End If
            Case 5:
                 'Grd_Docum.Columns(5).Text = gcodcen
                 Grd_Docum.Refresh
                 Grd_Docum.col = 6
            Case 7:
                Grd_Docum.col = 8
            Case 9:
                Grd_Docum.Row = Grd_Docum.Row + 1
                DataFactura.UpdateRecord
                calcula
                'txtfocus.SetFocus
                Grd_Docum.col = 3
                Grd_Docum.SetFocus
        End Select
    End If

End Sub

Private Sub Grd_Docum_OnAddNew()
    
    winddepa = ""
    txtdepa.Text = ""
    WITEMS = WITEMS + 1
    Grd_Docum.Columns(0) = WITEMS
    Grd_Docum.Columns(1) = wanno
    Grd_Docum.Columns(2) = ""
    'Grd_Docum.Columns(3) = ""
    'Grd_Docum.Columns(4) = ""
    If Len(Trim(wcentro)) <> 0 Then  'wnumobra
        Grd_Docum.Columns(5) = Format(wcentro, "000000")
    Else
        Grd_Docum.Columns(5) = ""
    End If
    'Grd_Docum.Columns(6) = ""
    Grd_Docum.Columns(7) = 0#
    Grd_Docum.Columns(8) = "*"
    If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
        Grd_Docum.Columns(9) = "H"
    Else
        Grd_Docum.Columns(9) = "D"
    End If
        
End Sub

Private Sub Grd_Docum_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 
 If Len(Trim(TxtCodPrv.Text)) > 0 Then
        Select Case LastCol
            Case 4:
   '             If Val(LastRow) = Grd_Docum.Row Then
                    If swigv = 1 Then Exit Sub
                        
                     If Valida_Contable(Grd_Docum.Columns(LastCol) & "") = False Then
                        MsgBox "La cuenta contable no existe.", 48, "Compras"
                        Grd_Docum.col = 4
                     End If
   '             End If
            Case 5:
                If val_cc(Grd_Docum.Columns(4) & "") = True Then
                    If Len(Trim(Grd_Docum.Columns(5))) = 0 Then
                       ' WCHANGE = 1
                        MsgBox "Falta ingresar el centro de costo.", 48, "Compras"
                        Grd_Docum.col = 5
                    Else
                        If valida_centros() Then
                            '------------- Agropecuaria
                            If WF1INGCTA_CC = "*" Then
                                If Len(Trim(pctacen)) > 0 Then
                                    Grd_Docum.Columns(4) = pctacen & Mid(Trim(Grd_Docum.Columns(4)), 3, 10)
                                End If
                            End If
                            '-------------------------
                        Else
                            MsgBox "El centro de costo no existe. Verifique", 48, "Compras"
                            Grd_Docum.Columns(5) = ""
                            Grd_Docum.col = 5
                        End If
                    End If
                End If
            Case 7:
                swmodigv = True
        End Select
    End If
    
End Sub

Private Sub grdanticipos_DblClick()
Dim wcorredoc   As Double
Dim nmonto      As Double
Dim xdifer      As Double
Dim xsaldodoc   As Double
Dim xsaldoant   As Double
Dim ximputado   As Double
Dim ximputaso   As Double

    If MsgBox("Está seguro de aplicar el anticipo ?", 36, "Compras") = 6 Then
        Set DbCtaPag = OpenDatabase(wrutabancos & "\ctaspag.mdb")
        Set tbcabcta = DbCtaPag.OpenRecordset("PAG_DCTO")
        Set tbdetcta = DbCtaPag.OpenRecordset("PAG_MVTO")
        
        tbcabcta.Index = "idprvdoc"
        tbcabcta.Seek "=", TxtCodPrv.Text, Right(Trim(CmbTipDoc.Text), 3) & Format(TxtSerDoc.Text, "000") & "/" & Format(Right(TxtNumDoc.Text, 7), "0000000")
        If Not tbcabcta.NoMatch Then
            wcorredoc = tbcabcta.Fields("correla")
        End If
        DataAnticipo.Recordset.MoveFirst
        Do While Not DataAnticipo.Recordset.EOF
            tbdetcta.AddNew
            tbdetcta.Fields("proveedor") = TxtCodPrv.Text
            tbdetcta.Fields("corr_comp") = DataAnticipo.Recordset.Fields("correla")
            tbdetcta.Fields("corr_dcto") = wcorredoc
            tbdetcta.Fields("tcambio") = TxtTipCam.Text
            If DataAnticipo.Recordset.Fields("monedao") = "S/." Then
                If OpcSol.Value = True Then
                    nmonto = Val(Format(PnlTotal(0).Caption, "0.00"))
                Else
                    nmonto = Format(Val(Format(PnlTotal(0).Caption, "0.00")) * TxtTipCam.Text, "0.00")
                End If
            Else
                If OpcDol.Value = True Then
                    nmonto = Val(Format(PnlTotal(0).Caption, "0.00"))
                Else
                    nmonto = Format(Val(Format(PnlTotal(0).Caption, "0.00")) / TxtTipCam.Text, "0.00")
                End If
            End If
            xdifer = DataAnticipo.Recordset.Fields("saldo") - nmonto
            If xdifer > 0 Then
                xsaldodoc = 0
                xsaldoant = xdifer
                ximputado = DataAnticipo.Recordset.Fields("saldo") - xdifer
                If DataAnticipo.Recordset.Fields("monedao") = "S/." Then
                    ximputaso = Format((DataAnticipo.Recordset.Fields("saldo") - xdifer) / TxtTipCam.Text, "0.00")
                Else
                    ximputaso = Format((DataAnticipo.Recordset.Fields("saldo") - xdifer) * TxtTipCam.Text, "0.00")
                End If
            Else
                xsaldodoc = Abs(xdifer)
                xsaldoant = 0
                ximputado = DataAnticipo.Recordset.Fields("saldo")
                If DataAnticipo.Recordset.Fields("monedao") = "S/." Then
                    ximputaso = Format(DataAnticipo.Recordset.Fields("saldo") / TxtTipCam.Text, "0.00")
                Else
                    ximputaso = Format(DataAnticipo.Recordset.Fields("saldo") * TxtTipCam.Text, "0.00")
                End If
            End If
            tbdetcta.Fields("imputado") = ximputado
            tbdetcta.Fields("imputaso") = ximputaso
            tbdetcta.Update
            '-------------- anticipo
            tbcabcta.Index = "nro_corre"
            tbcabcta.Seek "=", DataAnticipo.Recordset.Fields("correla")
            If Not tbcabcta.NoMatch Then
                tbcabcta.Edit
                tbcabcta.Fields("saldo") = xsaldoant
                tbcabcta.Update
            End If
            '-------------- documento
            tbcabcta.Index = "nro_corre"
            tbcabcta.Seek "=", wcorredoc
            If Not tbcabcta.NoMatch Then
                tbcabcta.Edit
                tbcabcta.Fields("saldo") = xsaldodoc
                tbcabcta.Update
            End If
            '--------------
            DataAnticipo.Recordset.MoveNext
        Loop
        tbdetcta.Close
        tbcabcta.Close
        DbCtaPag.Close
    End If
    grdanticipos.Visible = False

End Sub

Private Sub grdanticipos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        grdanticipos.Visible = False
    End If

End Sub

Private Sub Limpia_Datos()

    sw_numdoc = True

    txtref.Text = ""
    Text1.Text = ""
    WITEMS = 0
    TxtCodPrv.Text = ""
    TxtNomPrv.Text = ""
    TxtDirPrv.Text = ""
    TxtTelPrv.Text = ""
    TxtRucPrv.Text = ""
    TxtSerDoc.Text = ""
    TxtNumDoc.Text = ""

    TxtTipCam.Text = "0.000"
    TxtRefere.Text = ""
    gcodppp = ""
    txtocompra.Text = ""
    gobra = ""
    chkgg.Value = False

    lblocompra.Visible = False
    txtocompra.Visible = False
    

    lblfecven.Visible = True
    TxtFecVen.Visible = True
    txtfecha.Text = Format(Date, "dd/mm/yyyy")
    TxtFecVen.Text = Format(Date, "dd/mm/yyyy")
    txtfecharec.Text = Format(Date, "dd/mm/yyyy")

    PnlBasImp(0).Caption = "0.00"
    PnlBasImp(1).Caption = "0.00"
    PnlMonIna(0).Caption = "0.00"
    PnlMonIna(1).Caption = "0.00"
    txtigv(0).Text = "0.00"
    txtigv(1).Text = "0.00"
    TxtOtrImp(0).Text = "0.00"
    TxtOtrImp(1).Text = "0.00"
    PnlTotal(0).Caption = "0.00"
    PnlTotal(1).Caption = "0.00"
    txtredsuma.Text = "0.00"
    txtredresta.Text = "0.00"
    txtdcto.Text = "0.00"
    CmbTipDoc.ListIndex = 0
    cmbigv.ListIndex = 0
    cmbfpagos.ListIndex = 0

    DbTempfac.Execute ("Delete From temp_det")
    DataFactura.Refresh
    Grd_Docum.Refresh
    'TxtNumMov.SetFocus
    TxtNumMov.Text = ""

End Sub

Private Sub llena_gridoc()
'Dim DbInventa As DAO.DATABASE
'Dim tbif5  As DAO.RECORDSET
'Dim tbf4oc As DAO.RECORDSET
'Dim tbf3oc As DAO.RECORDSET
'Dim dbtemp As DAO.DATABASE
'Dim tbtemp As DAO.RECORDSET
'
'    If Tabla.RecordCount = 0 Or Val(Format(Grd_Docum.Columns(7), "0.00")) = 0# Then
'
'        swgraba_valing = True
'
'        Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
'        dbtemp.Execute ("delete * from temp_oc")
'        Set tbtemp = dbtemp.OpenRecordset("temp_oc")
'        tbtemp.Index = "idgasto"
'
'        Set DbInventa = OpenDatabase(ruta & "inventa.mdb")
'        Set tbif5 = DbInventa.OpenRecordset("if5pla")
'        tbif5.Index = "IdCodpro"
'
'        Set tbf4oc = DbInventa.OpenRecordset("if4vales")
'        tbf4oc.Index = "IDALMVAL"
'        Set tbf3oc = DbInventa.OpenRecordset("if3vales")
'        tbf3oc.Index = "IDALMVAL"
'
'        If frmhvali.DataGrid.SelCount > 0 Then
'            For i% = 0 To frmhvali.DataGrid.SelCount - 1
'                frmhvali.DataGrid.Row = frmhvali.DataGrid.SelRow(i%)
'
'                tbf4oc.Seek "=", frmhvali.DataGrid.Columns(0), frmhvali.DataGrid.Columns(1)
'                If Not tbf4oc.NoMatch Then
'                    XITEM = 0
'                    tbf3oc.Seek "=", frmhvali.DataGrid.Columns(0), frmhvali.DataGrid.Columns(1)
'                    If Not tbf3oc.NoMatch Then
'                        Do While tbf3oc.Fields("F2CODALM") = frmhvali.DataGrid.Columns(0) And tbf3oc.Fields("F4NUMVAL") = frmhvali.DataGrid.Columns(1) And Not tbf3oc.EOF
'                            '---------------------------------------------------
'                            xgasto = "": xnomgasto = "": xctacont = ""
'                            tbif5.Seek "=", tbf3oc.Fields("f5codpro") & ""
'                            If Not tbif5.NoMatch Then
'                                xgasto = tbif5.Fields("f3gasto") & ""
'                            End If
'                            tbgastos.Index = "bf9gin"
'                            tbgastos.Seek "=", "G", xgasto
'                            If Not tbgastos.NoMatch Then
'                                xnomgasto = Trim(tbgastos.Fields("nombre") & "")
'                                xctacont = Trim(tbgastos.Fields("cuenta") & "")
'                            End If
'                            '---------------------------------------------------
'                            tbtemp.Seek "=", xgasto
'                            If Not tbtemp.NoMatch Then
'                                tbtemp.Edit
'                            Else
'                                XITEM = XITEM + 1
'                                tbtemp.AddNew
'                                tbtemp.Fields("f3item") = XITEM
'                            End If
'                            tbtemp.Fields("f3anoorden") = wanno
'                            tbtemp.Fields("f3gasto") = xgasto
'                            tbtemp.Fields("f3ctacon") = xctacont
'                            tbtemp.Fields("f3concepto") = xnomgasto
'                            If OpcSol.Value = True Then
'                                tbtemp.Fields("f3importe") = tbf3oc.Fields("f3totite") + tbtemp.Fields("f3importe")
'                            Else
'                                tbtemp.Fields("f3importe") = tbf3oc.Fields("f3totdol") + tbtemp.Fields("f3importe")
'                            End If
'                            tbtemp.Fields("f3afecto") = "*"
'                            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
'                                tbtemp.Fields("F3DebHab") = "H"
'                            Else
'                                tbtemp.Fields("F3DebHab") = "D"
'                            End If
'                            tbtemp.Fields("f3cencos") = Trim(tbf4oc.Fields("f4centro") & "")
'                            tbtemp.Update
'                            tbf3oc.MoveNext
'                            If tbf3oc.EOF Then Exit Do
'                            If tbf3oc.Fields("F2CODALM") <> frmhvali.DataGrid.Columns(0) Or tbf3oc.Fields("F4NUMVAL") <> frmhvali.DataGrid.Columns(1) Then Exit Do
'                        Loop
'                    End If
'                End If
'            Next i%
'            '-----------------------------
'            If tbtemp.RecordCount > 0 Then
'                tbtemp.MoveFirst
'                DbTempFac.Execute ("Delete From temp_det")
'                Do While Not tbtemp.EOF
'                    Tabla.AddNew
'                    Tabla.Fields("f3item") = tbtemp.Fields("f3item")
'                    Tabla.Fields("f3anoorden") = tbtemp.Fields("f3anoorden")
'                    Tabla.Fields("f3gasto") = tbtemp.Fields("f3gasto") & ""
'                    Tabla.Fields("f3ctacon") = tbtemp.Fields("f3ctacon") & ""
'                    Tabla.Fields("f3concepto") = tbtemp.Fields("f3concepto") & ""
'                    Tabla.Fields("f3importe") = tbtemp.Fields("f3importe")
'                    Tabla.Fields("f3afecto") = tbtemp.Fields("f3afecto") & ""
'                    Tabla.Fields("f3debhab") = tbtemp.Fields("f3debhab") & ""
'                    Tabla.Fields("f3cencos") = tbtemp.Fields("f3cencos") & ""
'                    Tabla.Update
'                    tbtemp.MoveNext
'                Loop
'                DataFactura.Refresh
'                swmodigv = True
'                swmodcta = True
'                calcula
'            End If
'            '-----------------------------
'        End If
'
'        tbf3oc.Close
'        tbf4oc.Close
'        tbif5.Close
'        DbInventa.Close
'
'        tbtemp.Close
'        dbtemp.Close
'
'    End If

End Sub

Private Sub LLENA_TEMPCAB()
On Error GoTo error_llena

    TEMP_CABE.AddNew
    TEMP_CABE.Fields("F4COMPRO") = txtmesmov.Text & "/" & TxtNumMov.Text
    TEMP_CABE.Fields("F4TOTDEB") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4TOTHAB") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4CODORI") = worigen
    TEMP_CABE.Fields("F4ORIGEN") = "Registro de Compras"
    TEMP_CABE.Fields("F4FECHA") = txtfecha.Text
    TEMP_CABE.Fields("F4TIPCAMBD") = TxtTipCam.Text
    TEMP_CABE.Fields("F4PROVE") = TxtNomPrv.Text
    TEMP_CABE.Fields("F4IGV") = txtigv(0).Text
    TEMP_CABE.Fields("F4TOTAL") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4EMPRESA") = wnomcia
    TEMP_CABE.Fields("F4MONEDA") = IIf(OpcSol.Value = True, "SOLES", "DOLARES")
    TEMP_CABE.Fields("f4detalle") = TxtRefere.Text & ""
    TEMP_CABE.Fields("f4titulo") = "IMPRESION DEL COMPROBANTE"
    TEMP_CABE.Update

    Exit Sub

error_llena:
    Resume Next

End Sub

Private Sub LLENA_TEMPDET()
Dim RSDETALLE   As DAO.Recordset
Dim Busca       As String
Dim conta       As Integer
On Error GoTo error_llena_t

    Set RSDETALLE = DbTempfac.OpenRecordset("TEMP_DET")
    Busca = mes & "/" & TxtNumMov
    conta = 1
    RSDETALLE.MoveFirst
    Do While Not RSDETALLE.EOF
        temp_deta.AddNew
        temp_deta.Fields("F3COMPRO") = "" & Busca
        temp_deta.Fields("F3ELEMEN") = conta
        temp_deta.Fields("F3DETALL") = RSDETALLE.Fields("F3CONCEPTO")
        temp_deta.Fields("F3CODGAS") = RSDETALLE.Fields("F3GASTO")
        temp_deta.Fields("F5CODCTA") = RSDETALLE.Fields("F3CTACON")
        If RSDETALLE.Fields("f3debhab") = "D" Then
            temp_deta.Fields("F3DEBE") = Format(Abs(RSDETALLE.Fields("f3importe")), "#0.00")
        Else
            temp_deta.Fields("F3HABER") = Format(Abs(RSDETALLE.Fields("f3importe")), "#0.00")
        End If
        temp_deta.Fields("F3TIPCAMBD") = TxtTipCam.Text
        temp_deta.Fields("F3TIPDOC") = ""   ' TIPDOC
        temp_deta.Fields("F3COSTO") = RSDETALLE.Fields("F3CENCOS")
        temp_deta.Update
        RSDETALLE.MoveNext
        conta = conta + 1
        If RSDETALLE.EOF Then Exit Do
    Loop
    
    If UCase(Right(CmbTipDoc.Text, 3)) <> "HON" Then
        If Val(Format(txtigv(0).Text, "#0.00")) <> 0 Then
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "I.G.V."
            temp_deta.Fields("F3CODGAS") = "IGV"
            temp_deta.Fields("F5CODCTA") = wctaigv
            If UCase(Right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3HABER") = Abs(Val(Format(txtigv(0).Text, "#0.00")))
            Else
                temp_deta.Fields("F3DEBE") = Abs(Val(Format(txtigv(0).Text, "#0.00")))
            End If
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        If Val(Format(TxtOtrImp(0).Text, "#0.00")) <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "OTROS IMPUESTOS"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wctaotros
            If UCase(Right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3HABER") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
            Else
                temp_deta.Fields("F3DEBE") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
            End If
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        '---------------------------------------------------------
        If Val(Format(txtredsuma.Text, "0.00")) <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "REDONDEO"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wredsuma
            If UCase(Right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3HABER") = Val(Format(txtredsuma.Text, "#0.00"))
            Else
                temp_deta.Fields("F3DEBE") = Val(Format(txtredsuma.Text, "#0.00"))
            End If
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        '---------------------------------------------------------
        If Val(Format(txtredresta.Text, "0.00")) <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "REDONDEO"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wredresta
            If UCase(Right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3DEBE") = Val(Format(txtredresta.Text, "#0.00"))
            Else
                temp_deta.Fields("F3HABER") = Val(Format(txtredresta.Text, "#0.00"))
            End If
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        '---------------------------------------------------------
    Else
        If Val(Format(txtigv(0).Text, "#0.00")) <> 0 Then
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "RETENCION"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wctaret
            temp_deta.Fields("F3HABER") = Val(Format(txtigv(0).Text, "#0.00"))
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        If Val(Format(TxtOtrImp(0).Text, "#0.00")) <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "I.E.S."
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wctafon
            temp_deta.Fields("F3HABER") = Val(Format(TxtOtrImp(0).Text, "#0.00"))
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
    End If
           
    If Val(Format(PnlTotal(0).Caption, "#0.00")) <> 0 Then
        conta = conta + 1
        temp_deta.AddNew
        temp_deta.Fields("F3COMPRO") = "" & Busca
        temp_deta.Fields("F3ELEMEN") = conta
        temp_deta.Fields("F3DETALL") = Trim(TxtNomPrv.Text) & ""
        temp_deta.Fields("F3CODGAS") = ""
        temp_deta.Fields("F5CODCTA") = Trim(TxtTelPrv.Text) & ""
        If UCase(Right(CmbTipDoc.Text, 3)) <> "CRE" Then
            temp_deta.Fields("F3HABER") = Val(Format(PnlTotal(0).Caption, "#0.00"))
        Else
            temp_deta.Fields("F3DEBE") = Val(Format(PnlTotal(0).Caption, "#0.00"))
        End If
        temp_deta.Fields("F3TIPCAMBD") = TxtTipCam.Text
        temp_deta.Fields("F3TIPDOC") = Right(CmbTipDoc.Text, 3)
        temp_deta.Fields("F3COSTO") = ""
        temp_deta.Fields("F3DOCUM") = Right(CmbTipDoc.Text, 3) & TxtSerDoc.Text & "/" & TxtNumDoc.Text
        temp_deta.Update
    End If

    Exit Sub

error_llena_t:
    Resume Next

End Sub

Private Sub nuevo()
    
    cmdfacturas.Visible = False
    'BtnEditO.Visible = False
    WITEMS = 0
    
    'If wf1numera = "*" Then
    '    tbmes.Index = "IDDOCMES"
    '    TxtNumMov.Text = ""
    '    txtmesmov.Text = MES
    'Else
    '    tbmes.Index = "IDNUMMES"
    '    tbmes.Seek "=", MES
    '    If Not tbmes.NoMatch Then
    '       txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
    '       WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
    '       If Len(Trim(wnumobra)) <> 0 Then
    '            If Len(Trim(wnumobra)) = 2 Then
    '                WNUMERO = Format(wnumobra, "00") & Mid(Format(WNUMERO, "0000000"), 3, 5)
    '            End If
    '            If Len(Trim(wnumobra)) = 3 Then
    '                WNUMERO = Format(wnumobra, "000") & Mid(Format(WNUMERO, "0000000"), 4, 4)
    '            End If
    '       End If
    '       TxtNumMov.Text = Format(WNUMERO, "0000000")
    '    End If
    'End If
    
    OpcSol.Value = True
    Limpia_Datos
    'TxtNumMov.SetFocus
    Wnuevo = True
    windigraba = True
    swgraba_valing = False
    wnewnumero = True
    
    TxtCodPrv.SetFocus

End Sub

Private Sub Nuevo_Item()

        winddepa = ""
        txtdepa.Text = ""
        WITEMS = WITEMS + 1
        DataFactura.Recordset.AddNew
        DataFactura.Recordset.Fields("f3Item") = WITEMS
        DataFactura.Recordset.Fields("f3anoorden") = wanno
        DataFactura.Recordset.Fields("f3orden") = ""
        DataFactura.Recordset.Fields("F3Gasto") = ""
        DataFactura.Recordset.Fields("F3CtaCon") = ""
        If Len(Trim(wcentro)) <> 0 Then  'wnumobra
            DataFactura.Recordset.Fields("F3CENCOS") = Format(wcentro, "000000")
        Else
            DataFactura.Recordset.Fields("F3CENCOS") = ""
        End If
        DataFactura.Recordset.Fields("F3Concepto") = ""
        DataFactura.Recordset.Fields("F3Importe") = 0#
        DataFactura.Recordset.Fields("F3Afecto") = "*"
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
           DataFactura.Recordset.Fields("F3DebHab") = "H"
        Else
           DataFactura.Recordset.Fields("F3DebHab") = "D"
        End If
        DataFactura.Recordset.Update
    
End Sub


Private Sub OpcDol_Click(Value As Integer)
   PnlSigMon(0).Caption = "US$"
    'If opccre.Value = True Then
    '   opccre.SetFocus
    'Else
    '   opccon.SetFocus
    'end If
    
End Sub

Private Sub OpcDol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       'If opccre.Value = True Then
       '   opccre.SetFocus
       'Else
       '   opccon.SetFocus
       'End If
    End If

End Sub

Private Sub OpcSol_Click(Value As Integer)

    PnlSigMon(0).Caption = "S/."
    'If opccre.Value = True Then
    '   opccre.SetFocus
    'Else
    '   opccon.SetFocus
    'End If
    PnlOficial.Visible = False

End Sub

Private Sub OpcSol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
       'If opccre.Value = True Then
       '   opccre.SetFocus
       'Else
       '   opccon.SetFocus
       'End If
    End If
    
End Sub

Private Sub prov_oc()
Dim tbf4oc  As DAO.Recordset
Dim I       As Integer

    Set tbf4oc = dbcompras.OpenRecordset("if4orden")
    tbf4oc.Index = "IDNUMORD"
    
    tbf4oc.Seek "=", Val(txtocompra.Text & "")
    If Not tbf4oc.NoMatch Then
        
        For I% = 0 To cmbfpagos.ListCount - 1
            If Left(Right(cmbfpagos.List(I%), 4), 3) = tbf4oc.Fields("f4forpag") & "" Then
                cmbfpagos.ListIndex = I%
            End If
        Next
        
        TxtFecVen.Text = Format(tbf4oc.Fields("f4fecven") & "", "dd/mm/yyyy")
        'TxtRucPrv.Text = tbf4oc.Fields("f4codprv") & ""
        'TxtRucPrv_KeyPress 13
    End If
    tbf4oc.Close

End Sub

Private Sub Revisa_anticipos()
Dim Entra     As Integer
Dim dbtemprep As DAO.Database
Dim tbcabtmp  As DAO.Recordset
Dim dbctaspag As DAO.Database
'Dim tbcabcta  As DAO.RECORDSET
'Dim tbdetcta  As DAO.RECORDSET

    Set dbtemprep = OpenDatabase(wrutatemp & "\temppag.mdb")
    dbtemprep.Execute "Delete From Tmp_Dcto"
    Set tbcabtmp = dbtemprep.OpenRecordset("Tmp_Dcto")

    Set dbctaspag = OpenDatabase(wrutabancos & "ctaspag.mdb")
    Set tbcabcta = dbctaspag.OpenRecordset("pag_dcto")
    Set tbdetcta = dbctaspag.OpenRecordset("pag_mvto")

    DataAnticipo.DatabaseName = wrutatemp & "\temppag.Mdb"
    DataAnticipo.RecordSource = "Tmp_Dcto"
    DataAnticipo.Refresh

    Entra = 0
    tbcabcta.Index = "IDPRVDH"
    tbdetcta.Index = "IDCORRDCTO"
    tbcabcta.Seek "=", TxtCodPrv.Text, "D"
    If Not tbcabcta.NoMatch Then
        Do While tbcabcta.Fields("ProveedorO") = TxtCodPrv.Text And tbcabcta.Fields("Deb_Hab") = "D"
            If Format(tbcabcta.Fields("Saldo"), "0.00") > 0# And tbcabcta.Fields("f4ocompra") = Val(txtocompra.Text) Then
                tbcabtmp.AddNew
                Entra = 1
                tbcabtmp.Fields("Nro_Comp") = tbcabcta.Fields("Nro_Comp")
                tbcabtmp.Fields("Fch_Comp") = tbcabcta.Fields("fch_Comp")
                tbcabtmp.Fields("Dias") = CVDate(Date) - CVDate(tbcabcta.Fields("fch_Comp"))
                tbcabtmp.Fields("MonedaO") = IIf(tbcabcta.Fields("MonedaO") = "S", "S/.", "US$")
                tbcabtmp.Fields("TotalO") = Format(tbcabcta.Fields("TotalO"), "###,##0.00")
                tbcabtmp.Fields("Saldo") = Format(tbcabcta.Fields("Saldo"), "###,##0.00")
                tbcabtmp.Fields("deb_hab") = tbcabcta.Fields("Deb_Hab")
                tbcabtmp.Fields("correla") = tbcabcta.Fields("correla")
                tbcabtmp.Update
            End If
            tbcabcta.MoveNext
            If tbcabcta.EOF Then
                Exit Do
            End If
        Loop
        If tbcabtmp.RecordCount > 0 Then
            MsgBox "El proveedor registra los siguientes anticipos.", 48, "Compras"
            DataAnticipo.DatabaseName = wrutatemp & "\temppag.Mdb"
            DataAnticipo.RecordSource = "Tmp_Dcto"
            DataAnticipo.Refresh
            grdanticipos.Visible = True
        End If
    End If

    tbcabtmp.Close
    dbtemprep.Close
    
    tbdetcta.Close
    tbcabcta.Close
    dbctaspag.Close

End Sub

Private Sub Text1_DblClick()

        Text1_KeyDown 113, 0

End Sub

Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        llampro = 1
        Rem NSE FrmAyudaGas.Show 1
        frmhlpgastos.Show 1
        Text1.Text = Format(gcodppp, "000")
        Text1_KeyPress 13
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       gcodppp = Text1.Text
       
       Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
       Set tbcomtab = dbcomtabla.OpenRecordset("BF9GIN")
       tbcomtab.Index = "BF9GIN"
       tbcomtab.Seek "=", "G", gcodppp
       If Not tbcomtab.NoMatch Then
          
          gcueppp = tbcomtab.Fields("cuenta") & ""
          If Wnuevo = True Then
              If Trim(tbcomtab.Fields("moneda")) = "S" Then
                OpcSol.Value = True
                swmoneda = 3
              Else
                If Trim(tbcomtab.Fields("moneda")) = "D" Then
                    OpcDol.Value = True
                    swmoneda = 1
                Else
                    OpcSol.Value = True
                    swmoneda = 3
                End If
              End If
          End If
          If TxtCodPrv.Text = "9999" Then
            TxtNomPrv.SetFocus
          Else
            CmbTipDoc.SetFocus
          End If
       Else
          Text1.SetFocus
          MsgBox "El código ingresado no existe. Verifique.", 48, "Compras"
       End If
       If TxtCodPrv.Text = "9999" Then
            Tbproveedor.Index = "IDCODIGO"
            Tbproveedor.Seek "=", Trim(TxtCodPrv.Text)
            If Not Tbproveedor.NoMatch Then
                TxtTelPrv.Text = gcueppp & Tbproveedor.Fields("f2codcon") & ""
            End If
       Else
           tbcomtab.Index = "bf9gin"
           tbcomtab.Seek "=", "G", Trim(Text1.Text)
           If Not tbcomtab.NoMatch Then
            If tbcomtab.Fields("conta") & "" = "*" Then
                TxtTelPrv.Text = "" & Trim(gcueppp) & Trim(gsegppp)
            Else
                TxtTelPrv.Text = "" & Trim(gcueppp)
            End If
           End If
       End If
    End If

End Sub

Private Sub TxtCodPrv_DblClick()

    TxtCodPrv_KeyUp 113, 0

End Sub

Private Sub TxtCodPrv_KeyPress(KeyAscii As Integer)
Dim I   As Integer
 
     If KeyAscii = 13 Then
       If TxtCodPrv.Text = "" Then
          TxtRucPrv.SetFocus
       Else
          If TxtCodPrv.Text = "9999" Then
             TxtNomPrv.BackColor = &HFFFFFF
             TxtDirPrv.BackColor = &HFFFFFF
             TxtTelPrv.BackColor = &HFFFFFF
             TxtRucPrv.SetFocus
          Else
             If Wnuevo = True Then
                 Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
                 Set Tbproveedor = dbempresa.OpenRecordset("EF2PROVEEDORES")
                 Tbproveedor.Index = "IDCODIGO"
                 Tbproveedor.Seek "=", Trim(TxtCodPrv.Text)
                 If Not Tbproveedor.NoMatch Then
                    TxtNomPrv.Text = "" & Tbproveedor.Fields("F2NOMPROV")
                    TxtDirPrv.Text = "" & Tbproveedor.Fields("F2DIRPROV")
                    Rem WVR TxtRucPrv.Text = "" & TbProveedor.Fields("F2RUCPROV")
                    TxtRucPrv.Text = "" & Tbproveedor.Fields("F2NEWRUC")
    
                    gsegppp = Tbproveedor.Fields("F2CODCON") & ""
                    gmoneda = Tbproveedor.Fields("F2TIPMON") & ""
                    
                    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
                    Set tbcomtab = dbcomtabla.OpenRecordset("BF9GIN")
                    tbcomtab.Index = "idtipmon"
                    If Trim(gmoneda) = "S" Then
                        OpcSol.Value = True
                        swmoneda = 3
                        tbcomtab.Index = "IDTIPMON"
                        tbcomtab.Seek "=", "P", Trim(gmoneda)
                        If Not tbcomtab.NoMatch Then
                            Text1.Text = "" & tbcomtab.Fields("codigo")
                            gcueppp = "" & tbcomtab.Fields("cuenta")
                        Else
                            Text1.Text = ""
                        End If
                    Else
                        If Trim(gmoneda) = "D" Then
                            OpcDol.Value = True
                            swmoneda = 1
                            tbcomtab.Seek "=", "P", Trim(gmoneda)
                            If Not tbcomtab.NoMatch Then
                                Text1.Text = "" & tbcomtab.Fields("codigo")
                                gcueppp = "" & tbcomtab.Fields("cuenta")
                            Else
                                Text1.Text = ""
                            End If
                        Else
                            swmoneda = 3
                            OpcSol = True
                            If Len(Trim(Text1.Text)) = 0 Then
                                TxtTelPrv.Text = ""
                                Text1.Text = ""
                                llampro = 1
                                Rem NSE FrmAyudaGas.Show 1
                                frmhlpgastos.Show 1
                                Text1.Text = Format(gcodppp, "000")
                                Rem NSE Text1_KeyPress 13
                                If Text1.Visible = True Then
                                    Text1.SetFocus
                                Else
                                    CmbTipDoc.SetFocus
                                End If
                            Else
                                If wf1viscod <> "*" Then
                                    Text1.SetFocus
                                Else
                                    CmbTipDoc.SetFocus
                                End If
                            End If
                        End If
                    End If
                    tbcomtab.Index = "bf9gin"
                    tbcomtab.Seek "=", "G", Trim(Text1.Text)
                    If Not tbcomtab.NoMatch Then
                        If tbcomtab.Fields("conta") & "" = "*" Then
                            TxtTelPrv.Text = "" & gcueppp & gsegppp
                        Else
                            TxtTelPrv.Text = "" & gcueppp
                        End If
                    End If
                    If wf1viscod <> "*" Then
                        Text1.SetFocus
                    Else
'                        CmbTipDoc.SetFocus
                    End If
                 Else
                    TxtCodPrv.SetFocus
                 End If
            Else
                 'Set dbempresa = OpenDatabase(Ruta & "DB_BANCOS.MDB")
                 'Set tbproveedor = dbempresa.Openrecordset("EF2PROVEEDORES")
                 Tbproveedor.Index = "IDCODIGO"
                 Tbproveedor.Seek "=", Trim(TxtCodPrv.Text)
                 If Not Tbproveedor.NoMatch Then
                    gsegppp = Tbproveedor.Fields("F2CODCON") & ""
                    gmoneda = Tbproveedor.Fields("F2TIPMON") & ""
                 End If
                    
                    If wf1viscod <> "*" Then
                        Text1.SetFocus
                    Else
                        CmbTipDoc.SetFocus
                    End If
                
            End If
          End If
        '--- NSE
        '----------------------------------------------
        If wocompra = "*" Then
            lblocompra.Visible = True
            txtocompra.Visible = True
            If wf1tipdoc_asoc = "V" Then
                lblocompra.Caption = "Vale Ing."
            Else
                lblocompra.Caption = "O. Compra"
            End If
            Rem NSE txtocompra.SetFocus
            
            '----------------------------------------------------------
            '----------------------------------------------------------
            Tbproveedor.Index = "IDCODIGO"
            Tbproveedor.Seek "=", Trim(TxtCodPrv.Text)
            If Not Tbproveedor.NoMatch Then
                If Tbproveedor.Fields("f2orden") = True Then
                    If wf1tipdoc_asoc = "V" Then
                        'frmhvali.Show 1
                        'If frmhvali.datagrid.SelCount > 0 Then
                        '    Set DbInventa = OpenDatabase(ruta & "inventa.mdb")
                        '    Set tbf4oc = DbInventa.OpenRecordset("if4vales")
                        '    tbf4oc.Index = "IDALMVAL"
                        '    '----------------------------------------------------------
                        '    For i% = 0 To frmhvali.datagrid.SelCount - 1
                        '        frmhvali.datagrid.Row = frmhvali.datagrid.SelRow(i%)
                        '        If i% = 0 Then
                        '            ccade = "": cnumdoc = ""
                        '            tbf4oc.Seek "=", frmhvali.datagrid.Columns(0), frmhvali.datagrid.Columns(1)
                        '            If Not tbf4oc.NoMatch Then
                        '                For X% = 1 To Len(Trim(tbf4oc.Fields("f4numfac") & ""))
                        '                    If Mid(tbf4oc.Fields("f4numfac") & "", X%, 1) = "/" Then
                        '                        cserie = Format(ccade, "000")
                        '                        ccade = ""
                        '                    Else
                        '                        ccade = ccade & Mid(tbf4oc.Fields("f4numfac") & "", X%, 1)
                        '                    End If
                        '                Next X%
                        '                cnumdoc = Format(ccade, "0000000")
                        '            End If
                        '            txtocompra.Text = frmhvali.datagrid.Columns(1)
                        '        Else
                        '            txtocompra.Text = txtocompra.Text & " / " & frmhvali.datagrid.Columns(1)
                        '        End If
                        '    Next i%
                        '    '----------------------------------------------------------
                        '    tbf4oc.Close
                        '    DbInventa.Close
                
                        '    TxtSerDoc.Text = cserie
                        '    TxtNumDoc.Text = cnumdoc
                            
                        'End If
                    Else
                        FrmAyudaOrdCom.Show 1
                        If gcodord > 0 Then
                            txtocompra.Text = gcodord
                            TxtSerDoc.SetFocus
                            If wf1trasoc = "*" Then
                                If wf1tipdoc_asoc = "V" Then
                                    llena_gridoc
                                Else
                                    'graba_gridoc
                                    llena_oc
                                    If xtipmonoc = "S" Then
                                        OpcSol.Value = True
                                    Else
                                        OpcDol.Value = True
                                    End If
                                    For I = 0 To cmbfpagos.ListCount - 1
                                        If Left(Right(cmbfpagos.List(I), 4), 3) = xforpagoc Then
                                            cmbfpagos.ListIndex = I
                                        End If
                                    Next
                                    If IsDate(xfecvencoc) Then
                                        TxtFecVen.Text = Format(xfecvencoc, "DD/MM/YYYY")  'MIG
                                    End If
                                   Rem JCG TxtSerDoc.SetFocus
                                End If
                            End If
                            gcodord = 0
                        End If
                    End If
                Else
                  
                End If
            End If
            '----------------------------------------------------------
            '----------------------------------------------------------
        End If
        '----------------------------------------------

        '----------------------------------------------
        'If wocompra = "*" Then
        '    lblocompra.Visible = True
        '    txtocompra.Visible = True
        '    If wf1tipdoc_asoc = "V" Then
        '        lblocompra.Caption = "Vale Ing."
        '    Else
        '        lblocompra.Caption = "O. Compra"
        '    End If
        '    txtocompra.SetFocus
        'End If
        '----------------------------------------------
       End If
    End If

End Sub

Private Sub TxtCodPrv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Me.MousePointer = 11
        Text1.Text = ""
        gcodprov = "" & TxtCodPrv.Text
        FrmAyudaProv.Show 1
        TxtCodPrv.Text = gcodprov
        Me.MousePointer = 1
        TxtCodPrv_KeyPress 13
    End If

    If KeyCode = 116 Then
        If Len(Trim(TxtCodPrv.Text)) > 0 Then
            gcodprov = "" & TxtCodPrv.Text
            FrmProveedor.Show 1
        Else
            'MsgBox "Falta asignar el código del proveedor.", 48, "Compras"
        End If
    End If

End Sub

Private Sub TxtCodPrv_LostFocus()
'If txtcodprv.Text = "" Then
'txtcodprv.SetFocus
'End If

End Sub

Private Sub txtctadestino_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        txtctadestino.Visible = False
        Grd_Docum.col = 3: Grd_Docum.SetFocus
    End If

    If KeyCode = 113 Then
        gcoddepa = txtctadestino.Text
        Rem NSE frmacdes.Show 1
        txtctadestino.Text = gcoddepa
        txtctadestino_keypress 13
    End If
    
End Sub

Private Sub txtctadestino_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        tbcf5costo.Index = "CODCTACCOSTO1"
        tbcf5costo.Seek "=", Trim(Grd_Docum.Columns(4)), txtctadestino.Text
        If tbcf5costo.NoMatch Then
            MsgBox "La cuenta destino no existe. Verifique.", 48, "Compras"
            txtctadestino.SetFocus
        Else
            txtctadestino.Visible = False
            Grd_Docum.Columns(11) = txtctadestino.Text
            Grd_Docum.col = 5: Grd_Docum.SetFocus
        End If
    End If
    
End Sub

Private Sub txtdcto_GotFocus()

    txtdcto.SelStart = 0
    txtdcto.SelLength = Len(txtdcto.Text)
    
End Sub

Private Sub txtdcto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdcto.Text = Format(txtdcto.Text, "###,###,##0.00")
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
            If OpcDol.Value = True Then
                'TXTIGV(1).SetFocus
            Else
                BtnEditO.SetFocus
            End If
        Else
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
            BtnEditO.SetFocus
        End If
    End If
    
End Sub

Private Sub txtdepa_DblClick()

    txtdepa_KeyDown 113, 0

End Sub

Private Sub txtdepa_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        txtdepa.Visible = False
        Grd_Docum.col = 3: Grd_Docum.SetFocus
    End If

    If KeyCode = 113 Then
        gcoddepa = txtdepa.Text
        Rem NSE frmadepa.Show 1
        txtdepa.Text = gcoddepa
        txtdepa_keypress 13
    End If

End Sub

Private Sub txtdepa_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        tbdepa.Seek "=", txtdepa.Text
        If tbdepa.NoMatch Then
           MsgBox "El código del departamento no existe. Verifique.", 48, "Compras"
           txtdepa.SetFocus
        Else
            If Len(Trim(txtdepa.Text)) <> 0 Then
                txtdepa.Visible = False
                Grd_Docum.Columns(4) = txtdepa.Text & Mid(Grd_Docum.Columns(4), 3, 10)
                Grd_Docum.Columns(10) = txtdepa.Text
                Grd_Docum.col = 5: Grd_Docum.SetFocus
            End If
        End If
    End If

End Sub

Private Sub TxtDirPrv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        CmbTipDoc.SetFocus
    End If

End Sub

Private Sub txtdocpag_KeyPress(KeyAscii As Integer)
Dim entro   As Integer
    
    If KeyAscii = 13 Then
        tbtalon.Index = "idcodcta"
        tbtalon.Seek "=", gcodcta
        entro = 0
        If Not tbtalon.NoMatch Then
            Do
                If (Val(txtdocpag.Text) >= tbtalon.Fields("cdesde")) And (Val(txtdocpag.Text) <= tbtalon.Fields("chasta")) Then
                    entro = 1
                    TxtTipCam.SetFocus
                    Exit Do
                    
                Else
                    tbtalon.MoveNext
                End If
            Loop Until tbtalon.Fields("codcta") <> Val(gcodcta)
            If entro = 0 Then
                MsgBox "El Nro de documento no está en el rango del talonario", 64, "Información"
                Me.MousePointer = 1
                txtdocpag.SetFocus
            End If
        End If
    End If

End Sub

Private Sub txtfecha_GotFocus()

    txtfecha.SelStart = 0
    txtfecha.SelLength = 2

End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If

End Sub

Private Sub TxtFecha_LostFocus()

    If sw_numdoc = True Then
        If IsDate(txtfecha.Text) = True Then
            If Format(txtmesmov.Text, "00") <> Format(Mid(txtfecha.Text, 4, 2), "00") Then
                If wanno & Format(txtmesmov.Text, "00") < Year(txtfecha.Text) & Format(Mid(txtfecha.Text, 4, 2), "00") Then
                    swtc = True
                    MsgBox "La fecha del documento no puede ser mayor al mes de proceso. ", 48
                    txtfecha.SetFocus
                    swtc = False
                Else
                    MsgBox "El mes del documento no corresponde al mes de proceso. ", 48
                    tbcambios.Index = "cambios"
                    tbcambios.Seek "=", txtfecha.Text
                    If Not tbcambios.NoMatch Then
                        TxtTipCam.Text = tbcambios.Fields("cambio")
                    Else
                        TxtTipCam.Text = "0.000"
                    End If
                    TxtFecVen.Text = txtfecha.Text
                    TxtTipCam.SetFocus
                End If
            Else
                tbcambios.Index = "cambios"
                tbcambios.Seek "=", CVDate(txtfecha.Text)
                If Not tbcambios.NoMatch Then
                    TxtTipCam.Text = tbcambios.Fields("cambio")
                End If
                TxtFecVen.Text = txtfecha.Text
                TxtTipCam.SetFocus
            End If
        Else
            MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
            txtfecha.SetFocus
        End If
    End If

End Sub

Private Sub txtfecharec_GotFocus()

    txtfecharec.SelStart = 0
    txtfecharec.SelLength = Len(txtfecharec.Text)
    
End Sub

Private Sub txtfecharec_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If IsDate(txtfecharec.Text) = True Then
            cmbigv.SetFocus
        Else
            MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
            txtfecharec.SetFocus
        End If
    End If

End Sub

Private Sub TxtFecVen_GotFocus()

    TxtFecVen.SelStart = 0
    TxtFecVen.SelLength = 2
    
End Sub

Private Sub TxtFecVen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If IsDate(TxtFecVen.Text) = True Then
            txtfecharec.SetFocus
        Else
            MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
            TxtFecVen.SetFocus
        End If
    End If

End Sub

Private Sub TxtIgv_GotFocus(Index As Integer)
  
    txtigv(Index).SelStart = 0
    txtigv(Index).SelLength = Len(txtigv(Index).Text)
  
    If swigv = 0 Then
        MsgBox "Para validar el IGV debe poner ENTER en la columna vacía de Cuenta contable", 48, "Aviso"
        Grd_Docum.col = 3
        Rem NSE Grd_Docum.Row = Grd_Docum.Row
        Grd_Docum.SetFocus
    End If

End Sub

Private Sub txtigv_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
        Case 0:
            If KeyAscii = 13 Then
                txtigv(Index).Text = Format(txtigv(Index).Text, "###,###,##0.00")
                If Right(Trim(CmbTipDoc.Text), 3) = "Hon" Then
                    PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
                    TxtOtrImp(0).SetFocus
                Else
                    PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
                    If OpcSol.Value = True Then
                        PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
                    End If
                    TxtOtrImp(0).SetFocus
                End If
            End If
        Case 1:
            If KeyAscii = 13 Then
                txtigv(Index).Text = Format(txtigv(Index).Text, "###,###,##0.00")
                If Right(Trim(CmbTipDoc.Text), 3) = "Hon" Then
                    If OpcSol.Value = True Then
                        PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) - Val(Format(TxtOtrImp(1).Text, "0.00")) - Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
                    End If
                    TxtOtrImp(1).SetFocus
                Else
                    If OpcSol.Value = True Then
                        PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) + Val(Format(TxtOtrImp(1).Text, "0.00")) + Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
                    End If
                    TxtOtrImp(1).SetFocus
                End If
            End If
    End Select
  
End Sub

Private Sub TxtIgv_LostFocus(Index As Integer)

    If swigv = 1 Then
        'BtnEditO.Visible = True
        swigv = 0
    End If

End Sub

Private Sub txtmesmov_DblClick()

    TxtMesMov_KeyDown 113, 0

End Sub

Private Sub TxtMesMov_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        swayuda = True
        Form_Mes.Show 1
        swayuda = False
        meslet = Choose(Val(mes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
        Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
        DbTempfac.Execute ("Delete From temp_det")
        DataFactura.Refresh
        Grd_Docum.Refresh
        'Nuevo_Item
        If wf1numera = "*" Then
            tbmes.Index = "IDDOCMES"
            TxtNumMov.Text = ""
        Else
            tbmes.Index = "IDNUMMES"
            tbmes.Seek "=", mes
            If Not tbmes.NoMatch Then
                txtmesmov.Text = "" & tbmes.Fields("F2NUMMES")
                WNUMERO = "" & Val(Format(tbmes.Fields("F2NUMMOV"), "0000000")) + 1
                TxtNumMov.Text = Format(WNUMERO, "0000000")
            End If
        End If
        txtfecha.Text = Format("01/" & mes & "/" & wanno, "DD/MM/YYYY")
        TxtFecVen.Text = Format("01/" & mes & "/" & wanno, "DD/MM/YYYY")
        nuevo
        'TxtNumMov.SetFocus
        
        TxtCodPrv.SetFocus
        
    End If

End Sub

Private Sub TxtMesMov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtmesmov.Text = Format(mes, "00")
        SendKeys "{Tab}"
    End If

End Sub

Private Sub txtmesmov_LostFocus()

    If swayuda = False Then
        txtmesmov.Text = Format(mes, "00")
        SendKeys "{Tab}"
    End If

End Sub

Private Sub TxtNomPrv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtDirPrv.SetFocus
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtfecha.SetFocus
       'If Len(Trim(txtnumdoc.Text)) <> 0 Then
       '   txtnumdoc.Text = Format(txtnumdoc.Text, "0000000")
       '   tbcabregis.Index = "IDPROVDOC"
       '   If wnuevo = True Then
       '      Rem tbcabregis.Seek "=", Trim(TXTCODPRV.Text), Format(cmbtipdoc.ListIndex + 1, "00"), Trim(txtserdoc.Text), Trim(txtnumdoc.Text)
       '      tbcabregis.Seek "=", Trim(TxtCodPrv.Text), Format(Left(Right(Trim(cmbtipdoc.Text), 5), 2), "00"), Trim(txtserdoc.Text), Trim(txtnumdoc.Text)
       '      If Not tbcabregis.NoMatch Then
       '         MsgBox "El documento  ya  ha  sido  ingresado en  el Mes " & tbcabregis.Fields("F4MESMOV") & "  y  Nº " & tbcabregis.Fields("F4NUMMOV"), 48
       '         txtnumdoc.SetFocus
       '      Else
       '         TxtFecha.SetFocus
       '      End If
       '   Else
       '      'norma
       '      Rem If Trim(TXTCODPRV.text) <> Trim(wvcodprov) Or cmbtipdoc.ListIndex + 1 <> Val(wvtipdocu) Or Trim(txtserdoc.text) <> Trim(wvserdoc) Or Trim(txtnumdoc.text) <> Trim(wvnumdoc) Then
       '      If Trim(TxtCodPrv.Text) <> Trim(wvcodprov) Or Format(Left(Right(Trim(cmbtipdoc.Text), 5), 2), "00") <> Val(wvtipdocu) Or Trim(txtserdoc.Text) <> Trim(wvserdoc) Or Trim(txtnumdoc.Text) <> Trim(wvnumdoc) Then
       '         Rem tbcabregis.Seek "=", Trim(TXTCODPRV.text), Format(cmbtipdoc.ListIndex + 1, "00"), Trim(txtserdoc.text), Trim(txtnumdoc.text)
       '         tbcabregis.Seek "=", Trim(TxtCodPrv.Text), Format(Left(Right(Trim(cmbtipdoc.Text), 5), 2), "00"), Trim(txtserdoc.Text), Trim(txtnumdoc.Text)
       '         If Not tbcabregis.NoMatch Then
       '            MsgBox "El documento  ya  ha  sido  ingresado en  el Mes " & tbcabregis.Fields("F4MESMOV") & "  y  Nº " & tbcabregis.Fields("F4NUMMOV"), 48
       '            txtnumdoc.SetFocus
       '         Else
       '            TxtFecha.SetFocus
       '         End If
       '      Else
       '         TxtFecha.SetFocus
       '      End If
       '   End If
       '   tbcabregis.Index = "IDMESNUM"
       'Else
       '   MsgBox ("Numero de Documento no puede estar vacío")
       '   txtnumdoc.SetFocus
       'End If
    End If

End Sub

Private Sub TxtNumDoc_LostFocus()

       sw_numdoc = True

       If Len(Trim(TxtNumDoc.Text)) <> 0 Then
          TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000")
          TbCabRegis.Index = "IDPROVDOC"
          If Wnuevo = True Then
             Rem tbcabregis.Seek "=", Trim(TXTCODPRV.Text), Format(cmbtipdoc.ListIndex + 1, "00"), Trim(txtserdoc.Text), Trim(txtnumdoc.Text)
             TbCabRegis.Seek "=", Trim(TxtCodPrv.Text), Format(Left(Right(Trim(CmbTipDoc.Text), 5), 2), "00"), Trim(TxtSerDoc.Text), Trim(TxtNumDoc.Text)
             If Not TbCabRegis.NoMatch Then
                sw_numdoc = False
                MsgBox "El documento  ya  ha  sido  ingresado en  el Mes " & TbCabRegis.Fields("F4MESMOV") & "  y  Nº " & TbCabRegis.Fields("F4NUMMOV"), 48
                TxtNumDoc.SetFocus
                Rem TxtNumDoc_KeyPress 13
             Else
                txtfecha.SetFocus
             End If
          Else
             'norma
             Rem If Trim(TXTCODPRV.text) <> Trim(wvcodprov) Or cmbtipdoc.ListIndex + 1 <> Val(wvtipdocu) Or Trim(txtserdoc.text) <> Trim(wvserdoc) Or Trim(txtnumdoc.text) <> Trim(wvnumdoc) Then
             If Trim(TxtCodPrv.Text) <> Trim(wvcodprov) Or Format(Left(Right(Trim(CmbTipDoc.Text), 5), 2), "00") <> Val(wvtipdocu) Or Trim(TxtSerDoc.Text) <> Trim(wvserdoc) Or Trim(TxtNumDoc.Text) <> Trim(wvnumdoc) Then
                Rem tbcabregis.Seek "=", Trim(TXTCODPRV.text), Format(cmbtipdoc.ListIndex + 1, "00"), Trim(txtserdoc.text), Trim(txtnumdoc.text)
                TbCabRegis.Seek "=", Trim(TxtCodPrv.Text), Format(Left(Right(Trim(CmbTipDoc.Text), 5), 2), "00"), Trim(TxtSerDoc.Text), Trim(TxtNumDoc.Text)
                If Not TbCabRegis.NoMatch Then
                   sw_numdoc = False
                   MsgBox "El documento  ya  ha  sido  ingresado en  el Mes " & TbCabRegis.Fields("F4MESMOV") & "  y  Nº " & TbCabRegis.Fields("F4NUMMOV"), 48
                   TxtNumDoc.SetFocus
                Else
                   txtfecha.SetFocus
                End If
             Else
                txtfecha.SetFocus
             End If
          End If
          TbCabRegis.Index = "IDMESNUM"
       End If

End Sub

Private Sub TxtNumMov_DblClick()

    TxtNumMov_KeyDown 113, 0

End Sub

Private Sub TxtNumMov_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wayudamov = 0
        frmhlmov.Show 1
        wayudamov = 0
        If gnummov > 0 Then
            txtmesmov.Text = Format(mes, "00")
            TxtNumMov.Text = Format(gnummov, "0000000")
        End If
        txtnummov_KeyPress 13
    End If

End Sub

Private Sub txtnummov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.MousePointer = 11
        TxtNumMov.Text = Format(TxtNumMov.Text, "0000000")
        TbCabRegis.Index = "IDMESNUM"
        TbCabRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
        If Not TbCabRegis.NoMatch Then
            wnewnumero = False
            actualiza
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Or UCase(Right(Trim(CmbTipDoc.Text), 3)) = "DEB" Then
                cmdfacturas.Visible = True
            End If
            WITEMS = Calcula_Item()
            Wnuevo = False
            windigraba = False
            sw_doc = True
        Else
            wnewnumero = True
            WGRABO = 0
            WITEMS = 0
            Limpia_Datos
            Wnuevo = True
            windigraba = True
            cmdfacturas.Visible = False
            sw_doc = False
            
        End If
        Me.MousePointer = 1
        Rem SendKeys "{Tab}"
        TxtCodPrv.SetFocus
    End If
    
End Sub

Private Sub txtocompra_DblClick()

    txtocompra_KeyDown 113, 0

End Sub

Private Sub txtocompra_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tbf4oc  As DAO.Recordset
Dim ccade   As String
Dim cnumdoc As String
Dim cserie  As String

    If KeyCode = 113 Then
        If wf1tipdoc_asoc = "V" Then
'            frmhvali.Show 1
'            If frmhvali.DataGrid.SelCount > 0 Then
'                Set DbInventa = OpenDatabase(ruta & "inventa.mdb")
'                Set tbf4oc = DbInventa.OpenRecordset("if4vales")
'                tbf4oc.Index = "IDALMVAL"
'                '----------------------------------------------------------
'                For i% = 0 To frmhvali.DataGrid.SelCount - 1
'                    frmhvali.DataGrid.Row = frmhvali.DataGrid.SelRow(i%)
'                    If i% = 0 Then
'                        ccade = "": cnumdoc = ""
'                        tbf4oc.Seek "=", frmhvali.DataGrid.Columns(0), frmhvali.DataGrid.Columns(1)
'                        If Not tbf4oc.NoMatch Then
'                            For X% = 1 To Len(Trim(tbf4oc.Fields("f4numfac") & ""))
'                                If Mid(tbf4oc.Fields("f4numfac") & "", X%, 1) = "/" Then
'                                    cserie = Format(ccade, "000")
'                                    ccade = ""
'                                Else
'                                    ccade = ccade & Mid(tbf4oc.Fields("f4numfac") & "", X%, 1)
'                                End If
'                            Next X%
'                            cnumdoc = Format(ccade, "0000000")
'                        End If
'                        txtocompra.Text = frmhvali.DataGrid.Columns(1)
'                    Else
'                        txtocompra.Text = txtocompra.Text & " / " & frmhvali.DataGrid.Columns(1)
'                    End If
'                Next i%
'                '----------------------------------------------------------
'                tbf4oc.Close
'                DbInventa.Close
'
'                TxtSerDoc.Text = cserie
'                TxtNumDoc.Text = cnumdoc
'
'            End If
        Else
            FrmAyudaOrdCom.Show 1
        End If
        txtocompra_Keypress 13
    End If

End Sub

Private Sub txtocompra_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtocompra.Text & "")) > 0 Then
            If wf1trasoc = "*" Then
                prov_oc
            End If
        End If
        SendKeys "{Tab}"
    End If

End Sub

Private Sub txtocompra_LostFocus()

    If Len(Trim(txtocompra.Text)) > 0 Then
        If Len(Trim(gorden_cs)) = 0 Then
            gorden_cs = "O/C"
        End If
    End If

End Sub

Private Sub TxtOtrImp_GotFocus(Index As Integer)

    TxtOtrImp(Index).SelStart = 0
    TxtOtrImp(Index).SelLength = Len(TxtOtrImp(Index).Text)

End Sub

Private Sub TxtOtrImp_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
        Case 0:
            If KeyAscii = 13 Then
                TxtOtrImp(Index).Text = Format(TxtOtrImp(Index).Text, "###,###,##0.00")
                If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                    PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
                    If OpcDol.Value = True Then
                        'TXTIGV(1).SetFocus
                    Else
                        BtnEditO.SetFocus
                    End If
                Else
                    PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
                    If Checkdatos.Value = True Then
                        txtredsuma.SetFocus
                    Else
                        BtnEditO.SetFocus
                    End If
                End If
            End If
        Case 1:
            If KeyAscii = 13 Then
                TxtOtrImp(Index).Text = Format(TxtOtrImp(Index).Text, "###,###,##0.00")
                If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                    If OpcSol.Value = True Then
                        PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) - Val(Format(TxtOtrImp(1).Text, "0.00")) - Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
                    End If
                    CmdGrabar.SetFocus
                Else
                    If OpcSol.Value = True Then
                        PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(1).Caption, "0.00")) + Val(Format(PnlMonIna(1).Caption, "0.00")) + Val(Format(TxtOtrImp(1).Text, "0.00")) + Val(Format(txtigv(1).Text, "0.00")), "###,##0.00")
                    End If
                    BtnEditO.SetFocus
                End If
            End If
    End Select

End Sub

Private Sub txtredresta_GotFocus()

    txtredresta.SelStart = 0
    txtredresta.SelLength = Len(txtredresta.Text)
    
End Sub

Private Sub txtredresta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtredresta.Text = Format(txtredresta.Text, "###,###,##0.00")
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
            If OpcDol.Value = True Then
                'TXTIGV(1).SetFocus
            Else
                txtdcto.SetFocus
            End If
        Else
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
            txtdcto.SetFocus
        End If
    End If
    
End Sub

Private Sub txtredsuma_GotFocus()

    txtredsuma.SelStart = 0
    txtredsuma.SelLength = Len(txtredsuma.Text)
    
End Sub

Private Sub txtredsuma_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtredsuma.Text = Format(txtredsuma.Text, "###,###,##0.00")
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
            txtredresta.SetFocus
        Else
            PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
            If OpcSol.Value = True Then
                PnlTotal(1).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(txtigv(0).Text, "0.00")), "###,##0.00")
            End If
            txtredresta.SetFocus
        End If
    End If
    
End Sub

Private Sub txtref_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtRefere.SetFocus
    End If

End Sub

Private Sub TxtRefere_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'If wingobra = "*" Then
        '    frmobras.Show 1
        'End If
        'If wf1trasoc = "*" Then
        '    If wf1tipdoc_asoc = "V" Then
        '        llena_gridoc
        '    Else
        '        graba_gridoc
        '    End If
        'End If
        'KeyAscii = 0
        ''txtfocus.SetFocus
        
        'DataFactura.Refresh
        'Grd_Docum.Refresh
        Grd_Docum.col = 3
        Grd_Docum.SetFocus
    End If

End Sub

Private Sub TxtRucPrv_KeyPress(KeyAscii As Integer)
Dim I   As Integer

    If KeyAscii = 13 Then
       If TxtCodPrv.Text = "" And TxtRucPrv = "" Then
          TxtCodPrv.SetFocus
       Else
          If TxtCodPrv.Text = "9999" Then
             Text1.SetFocus
          Else
             Tbproveedor.Index = "IDCODPROV"
             Tbproveedor.Seek "=", Trim(TxtRucPrv.Text)
             If Not Tbproveedor.NoMatch Then
                TxtCodPrv.Text = "" & Tbproveedor.Fields("F2CODPROV")
                TxtNomPrv.Text = "" & Tbproveedor.Fields("F2NOMPROV")
                TxtDirPrv.Text = "" & Tbproveedor.Fields("F2DIRPROV")
                Rem WVR TxtRucPrv.Text = "" & TbProveedor.Fields("F2RUCPROV")
                TxtRucPrv.Text = "" & Tbproveedor.Fields("F2NEWRUC")

                gsegppp = Tbproveedor.Fields("F2CODCON") & ""
                gmoneda = Tbproveedor.Fields("F2TIPMON") & ""

                Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
                Set tbcomtab = dbcomtabla.OpenRecordset("BF9GIN")
                tbcomtab.Index = "idtipmon"
                If Trim(gmoneda) = "S" Then
                    OpcSol.Value = True
                    swmoneda = 3
                    tbcomtab.Index = "IDTIPMON"
                    tbcomtab.Seek "=", "P", Trim(gmoneda)
                    If Not tbcomtab.NoMatch Then
                        Text1.Text = "" & tbcomtab.Fields("codigo")
                        gcueppp = "" & tbcomtab.Fields("cuenta")
                    Else
                        Text1.Text = ""
                    End If
                Else
                    If Trim(gmoneda) = "D" Then
                        OpcDol.Value = True
                        swmoneda = 1
                        tbcomtab.Seek "=", "P", Trim(gmoneda)
                        If Not tbcomtab.NoMatch Then
                            Text1.Text = "" & tbcomtab.Fields("codigo")
                            gcueppp = "" & tbcomtab.Fields("cuenta")
                        Else
                            Text1.Text = ""
                        End If
                    Else
                        swmoneda = 3
                        OpcSol = True
                        If Len(Trim(Text1.Text)) = 0 Then
                            TxtTelPrv.Text = ""
                            Text1.Text = ""
                            llampro = 1
                            Rem NSE FrmAyudaGas.Show 1
                            frmhlpgastos.Show 1
                            Text1.Text = Format(gcodppp, "000")
                            Rem NSE Text1_KeyPress 13
                            If Text1.Visible = True Then
                                Text1.SetFocus
                            End If
                        Else
                            If wf1viscod <> "*" Then
                                Text1.SetFocus
                            Else
                                CmbTipDoc.SetFocus
                            End If
                        End If
                    End If
                End If
                tbcomtab.Index = "bf9gin"
                tbcomtab.Seek "=", "G", Trim(Text1.Text)
                If Not tbcomtab.NoMatch Then
                    If tbcomtab.Fields("conta") & "" = "*" Then
                        TxtTelPrv.Text = "" & gcueppp & gsegppp
                    Else
                        TxtTelPrv.Text = "" & gcueppp
                    End If
                End If
                If wf1viscod <> "*" Then
                    Text1.SetFocus
                Else
                    CmbTipDoc.SetFocus
                End If
                
                '----------------------------------------------------------
                '----------------------------------------------------------
                '----------------------------------------------------------
                If wocompra = "*" Then
                    lblocompra.Visible = True
                    txtocompra.Visible = True
                    If wf1tipdoc_asoc = "V" Then
                        lblocompra.Caption = "Vale Ing."
                    Else
                        lblocompra.Caption = "O. Compra"
                    End If
                    Rem NSE txtocompra.SetFocus
                    
                    '----------------------------------------------------------
                    '----------------------------------------------------------
                    Tbproveedor.Index = "IDCODIGO"
                    Tbproveedor.Seek "=", Trim(TxtCodPrv.Text)
                    If Not Tbproveedor.NoMatch Then
                        If Tbproveedor.Fields("f2orden") = True Then
                            If wf1tipdoc_asoc = "V" Then
                            Else
                                FrmAyudaOrdCom.Show 1
                                If gcodord > 0 Then
                                    txtocompra.Text = gcodord
                                    If wf1trasoc = "*" Then
                                        If wf1tipdoc_asoc = "V" Then
                                            llena_gridoc
                                        Else
                                            'graba_gridoc
                                            llena_oc
                                            If xtipmonoc = "S" Then
                                                OpcSol.Value = True
                                            Else
                                                OpcDol.Value = True
                                            End If
                                            For I = 0 To cmbfpagos.ListCount - 1
                                                If Left(Right(cmbfpagos.List(I), 4), 3) = xforpagoc Then
                                                    cmbfpagos.ListIndex = I
                                                End If
                                            Next
                                            If IsDate(xfecvencoc) Then
                                                TxtFecVen.Text = Format(xfecvencoc, "DD/MM/YYYY")  'MIG
                                            End If
                                            TxtSerDoc.SetFocus
                                        End If
                                    End If
                                    gcodord = 0
                                End If
                            End If
                        End If
                    End If
                    '----------------------------------------------------------
                    '----------------------------------------------------------
                End If
                '----------------------------------------------------------
                '----------------------------------------------------------
                '----------------------------------------------------------
             Else
                If MsgBox(" El Ruc No existe . Desea Registrarlo ?", 36, "Atención") = 6 Then
                   rucprv = TxtRucPrv.Text
                   FrmProveedor.Show 1
                   TxtRucPrv.Text = rucprv
                    
                   Else
                   TxtCodPrv.Text = "9999"
                   TxtCodPrv.SetFocus

                End If

                
             End If
          End If
       End If
       Tbproveedor.Index = "IDCODPROV"
    End If

End Sub

Private Sub TxtSerDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       TxtSerDoc.Text = Format(TxtSerDoc.Text, "000")
       TxtNumDoc.SetFocus
    End If

End Sub

Private Sub TxtSerDoc_LostFocus()

    If Len(Trim(TxtSerDoc.Text)) = 0 Then
        TxtSerDoc.Text = "000"
    End If

End Sub

Private Sub TxtTelPrv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       CmbTipDoc.SetFocus
    End If

End Sub

Private Sub TxtTipCam_GotFocus()

    TxtTipCam.SelStart = 0
    TxtTipCam.SelLength = Len(TxtTipCam.Text)
    
End Sub

Private Sub Txttipcam_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(TxtTipCam.Text) > 0# Then
            If OpcDol.Value = True Then
                 PnlBasImp(1).Caption = Format$(Val(Format(PnlBasImp(0).Caption, "0.00")) * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
                 PnlMonIna(1).Caption = Format$(Val(Format(PnlMonIna(0).Caption, "0.00")) * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
                 txtigv(1).Text = Format$(Val(Format(txtigv(0).Text, "0.00")) * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
                 TxtOtrImp(1).Text = Format$(Val(Format(TxtOtrImp(0).Text, "0.00")) * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
                 PnlTotal(1).Caption = Format$(Val(Format(PnlTotal(0).Caption, "0.00")) * Val(Format(TxtTipCam.Text, "0.000")), "###,##0.00")
            End If
            If swmoneda = 3 Or swmoneda = 1 Then       'JCG ARREGLAR
                Rem opccre.SetFocus
                cmbfpagos.SetFocus
            Else
                If OpcSol.Value = True Then
                    OpcSol.SetFocus
                Else
                    OpcDol.SetFocus
                End If
            End If
        End If
    End If

End Sub

Private Sub txttipcam_LostFocus()

    If Val(TxtTipCam.Text) = 0# And swtc = False Then
        MsgBox "Ingrese Tipo de Cambio", 48, "Aviso"
        'txttipcam.SetFocus
    End If

End Sub

Private Function val_cc(cuenta As String)
Dim sw As Integer

    sw = False
    If Len(Trim(wrutaconta)) <> 0 Then
        tbcf5pla.Index = "cf5pla"
        tbcf5pla.Seek "=", cuenta & ""
        If Not tbcf5pla.NoMatch Then
            If tbcf5pla.Fields("f5cc") = True Then
                sw = True
            Else
                sw = False
            End If
        End If
    Else
       sw = False
    End If
    val_cc = sw
          
End Function

Private Function valida_centros()
Dim sw As Integer

    sw = False
    If Len(Trim(wrutaconta)) <> 0 Then
        tbcf5pla.Index = "cf5pla"
        tbcf5pla.Seek "=", Grd_Docum.Columns(4) & ""
        If Not tbcf5pla.NoMatch Then
            If tbcf5pla.Fields("f5cc") = True Then
                tbcentros.Index = "centros"
                tbcentros.Seek "=", Grd_Docum.Columns(5) & ""
                If Not tbcentros.NoMatch Then
                    pctacen = tbcentros.Fields("f3ctacont") & ""
                    sw = True
                Else
                    sw = False
                End If
            Else
               sw = True
            End If
        End If
    Else
       sw = True
    End If
    valida_centros = sw

End Function

Private Function Valida_Contable(xx As String)
Dim sw As Integer

    sw = False
    If Len(Trim(wrutaconta)) <> 0 Then
        tbcf5pla.Index = "cf5pla"
        tbcf5pla.Seek "=", xx
        If Not tbcf5pla.NoMatch Then
           sw = True
        Else
           sw = False
        End If
    Else
        sw = True
    End If
    Valida_Contable = sw

End Function

Private Function Valida_Gasto()
Dim sw As Integer

    If Len(Trim("" & Grd_Docum.Columns(3))) <> 0 Then  ' ERROR *
       Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
       tbgastos.Index = "BF9GIN"
       tbgastos.Seek "=", "G", Trim("" & Grd_Docum.Columns(3))
       If Not tbgastos.NoMatch Then
          If tbgastos.Fields("tipo") & "" <> "P" Then
              winddepa = tbgastos.Fields("inddepa") & ""
              Grd_Docum.Columns(4) = "" & tbgastos.Fields("CUENTA")
              Grd_Docum.Columns(6) = "" & tbgastos.Fields("NOMBRE")
              sw = True
          Else
              sw = False
          End If
       Else
          sw = False
       End If
    Else
       sw = True
    End If
    Valida_Gasto = sw

End Function

Private Function valida_centros2(ccuenta As String, ccentro As String)
Dim sw As Integer

    sw = False
    If Len(Trim(wrutaconta)) <> 0 Then
        tbcf5pla.Index = "cf5pla"
        tbcf5pla.Seek "=", ccuenta
        If Not tbcf5pla.NoMatch Then
           If tbcf5pla.Fields("f5cc") = True Then
              tbcentros.Index = "centros"
              tbcentros.Seek "=", ccentro
              If Not tbcentros.NoMatch Then
                 sw = True
              Else
                 sw = False
              End If
           Else
              sw = True
           End If
        End If
    Else
       sw = True
    End If
    valida_centros2 = sw

End Function

Private Sub llena_oc()
Dim DbInventa   As DAO.Database
Dim tbif5       As DAO.Recordset
Dim dbcomp      As DAO.Database
Dim tbf4oc      As DAO.Recordset
Dim tbf3oc      As DAO.Recordset
Dim dbtemp      As DAO.Database
Dim tbtemp      As DAO.Recordset
Dim dbempreoc   As DAO.Database
Dim tbef2serv   As DAO.Recordset
Dim xnomgasto   As String
Dim xctacont    As String

    If Tabla.RecordCount = 0 Then
        swgraba_valing = True
        Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
        dbtemp.Execute ("delete * from temp_oc")
        Set tbtemp = dbtemp.OpenRecordset("temp_oc")
        tbtemp.Index = "idgasto"

        Set DbInventa = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
        Set tbif5 = DbInventa.OpenRecordset("if5pla")
        tbif5.Index = "IdCodpro"

        Set dbempreoc = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        Set tbef2serv = dbempreoc.OpenRecordset("ef2servicios")
        tbef2serv.Index = "IDCODSER"
        
        Set dbcomp = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        If Left(gorden_cs, 3) = "O/C" Then
            Set tbf4oc = dbcomp.OpenRecordset("if4orden")
            tbf4oc.Index = "IDNUMORD"
            Set tbf3oc = dbcomp.OpenRecordset("if3orden")
            tbf3oc.Index = "IDNUMORD"
        Else
            Set tbf4oc = dbcomp.OpenRecordset("if4servicio")
            tbf4oc.Index = "IDNUMORD"
            Set tbf3oc = dbcomp.OpenRecordset("if3servicio")
            tbf3oc.Index = "IDNUMORD"
        End If
        
        'If FrmAyudaOrdCom.datagrid.SelCount > 0 Then
        '    For i% = 0 To FrmAyudaOrdCom.datagrid.SelCount - 1
        '        FrmAyudaOrdCom.datagrid.Row = FrmAyudaOrdCom.datagrid.SelRow(i%)
        '        tbf4oc.Seek "=", Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15))
                tbf4oc.Seek "=", Val(txtocompra.Text & "")
                If Not tbf4oc.NoMatch Then
                    xtotaloc = Val(tbf4oc.Fields("f4monto") & "")
                    xtipmonoc = tbf4oc.Fields("f4tipmon") & ""
                    xforpagoc = tbf4oc.Fields("f4forpag") & ""
                    If IsDate(tbf4oc.Fields("f4fecven")) Then
                        xfecvencoc = Format(tbf4oc.Fields("f4fecven") & "", "dd/mm/yyyy")
                    End If
                    XITEM = 0
                    'tbf3oc.Seek "=", Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15))
                    tbf3oc.Seek "=", Val(txtocompra.Text & "")
                    If Not tbf3oc.NoMatch Then
                        'Do While tbf3oc.Fields("F4NUMORD") = Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15)) And Not tbf3oc.EOF
                        Do While tbf3oc.Fields("F4NUMORD") = Val(txtocompra.Text & "") And Not tbf3oc.EOF
                            '---------------------------------------------------
                            xgasto = "": xnomgasto = "": xctacont = ""
                            
                            '---------------------------------------------------
                            '---------------------------------------------------
                            If Left(gorden_cs, 3) = "O/C" Then
                                tbif5.Seek "=", tbf3oc.Fields("f3codpro") & ""
                                If Not tbif5.NoMatch Then
                                    xgasto = tbif5.Fields("f3gasto") & ""
                                End If
                            Else
                                tbef2serv.Seek "=", tbf3oc.Fields("f3codpro") & ""
                                If Not tbef2serv.NoMatch Then
                                    xgasto = tbef2serv.Fields("f3gasto") & ""
                                End If
                            End If
                            '---------------------------------------------------
                            '---------------------------------------------------
                            tbgastos.Index = "bf9gin"
                            tbgastos.Seek "=", "G", xgasto
                            If Not tbgastos.NoMatch Then
                                xnomgasto = Trim(tbgastos.Fields("nombre") & "")
                                xctacont = Trim(tbgastos.Fields("cuenta") & "")
                            End If
                            '---------------------------------------------------
                            tbtemp.Seek "=", xgasto
                            If Not tbtemp.NoMatch Then
                                tbtemp.Edit
                            Else
                                XITEM = XITEM + 1
                                tbtemp.AddNew
                                tbtemp.Fields("f3item") = XITEM
                            End If
                            tbtemp.Fields("f3anoorden") = wanno
                            tbtemp.Fields("f3gasto") = xgasto
                            tbtemp.Fields("f3ctacon") = xctacont
                            tbtemp.Fields("f3concepto") = xnomgasto
                            If Left(gorden_cs, 3) = "O/C" Then
                                If OpcSol.Value = True Then
                                    tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
                                Else
                                    tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
                                End If
                            Else
                                'If OpcSol.value = True Then
                                    tbtemp.Fields("f3importe") = Val(Format(tbf3oc.Fields("f3canpro") * tbf3oc.Fields("f3precos"), "0.00")) + tbtemp.Fields("f3importe")
                                'Else
                                '    tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
                                'End If
                            End If
                            tbtemp.Fields("f3afecto") = tbf3oc.Fields("f5afecto") & ""
                            If UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                                tbtemp.Fields("F3DebHab") = "H"
                            Else
                                tbtemp.Fields("F3DebHab") = "D"
                            End If
                            tbtemp.Fields("f3cencos") = Trim(tbf4oc.Fields("f4centro") & "")
                            tbtemp.Update
                            tbf3oc.MoveNext
                            If tbf3oc.EOF Then Exit Do
                            'If tbf3oc.Fields("F4NUMORD") <> Val(Mid(FrmAyudaOrdCom.datagrid.Columns(0), 4, 15)) Then Exit Do
                            If tbf3oc.Fields("F4NUMORD") <> Val(txtocompra.Text & "") Then Exit Do
                        Loop
                    End If
                End If
        '    Next i%
            '-----------------------------
            If tbtemp.RecordCount > 0 Then
                tbtemp.MoveFirst
                DbTempfac.Execute ("Delete From temp_det")
                Do While Not tbtemp.EOF
                    Tabla.AddNew
                    Tabla.Fields("f3item") = tbtemp.Fields("f3item")
                    Tabla.Fields("f3anoorden") = tbtemp.Fields("f3anoorden")
                    Tabla.Fields("f3gasto") = tbtemp.Fields("f3gasto") & ""
                    Tabla.Fields("f3ctacon") = tbtemp.Fields("f3ctacon") & ""
                    Tabla.Fields("f3concepto") = tbtemp.Fields("f3concepto") & ""
                    Tabla.Fields("f3importe") = tbtemp.Fields("f3importe")
                    Tabla.Fields("f3afecto") = tbtemp.Fields("f3afecto") & ""
                    Tabla.Fields("f3debhab") = tbtemp.Fields("f3debhab") & ""
                    Tabla.Fields("f3cencos") = tbtemp.Fields("f3cencos") & ""
                    Tabla.Update
                    tbtemp.MoveNext
                Loop
                DataFactura.Refresh
                swmodigv = True
                swmodcta = True
                calcula
            End If
            '-----------------------------
        'End If
        
        tbf3oc.Close
        tbf4oc.Close
        dbcomp.Close

        tbif5.Close
        DbInventa.Close

        tbtemp.Close
        dbtemp.Close

        tbef2serv.Close
        dbempreoc.Close

    End If

End Sub

Private Sub actualiza_provisiones()

    Tbproveedor.Index = "IDCODIGO"
    Tbproveedor.Seek "=", TxtCodPrv.Text
    If Not Tbproveedor.NoMatch Then
        If sw_grabanse = True Then
            Tbproveedor.Edit
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) + Val(Format(PnlTotal(0).Caption, "0.00"))
            Else
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) - Val(Format(PnlTotal(0).Caption, "0.00"))
            End If
            Tbproveedor.Update
        Else
            Tbproveedor.Edit
            If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) - xmonto_anterior
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) + Val(Format(PnlTotal(0).Caption, "0.00"))
            Else
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) + xmonto_anterior
                Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) - Val(Format(PnlTotal(0).Caption, "0.00"))
            End If
            Tbproveedor.Update
        End If
    End If

End Sub

Private Sub elimina_provision()

    Tbproveedor.Index = "IDCODIGO"
    Tbproveedor.Seek "=", TxtCodPrv.Text
    If Not Tbproveedor.NoMatch Then
        Tbproveedor.Edit
        If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
            Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) - Val(Format(PnlTotal(0).Caption, "0.00"))
        Else
            Tbproveedor.Fields("f2provi" & txtmesmov.Text) = Val("" & Tbproveedor.Fields("f2provi" & txtmesmov.Text)) + Val(Format(PnlTotal(0).Caption, "0.00"))
        End If
        Tbproveedor.Update
    End If

End Sub
