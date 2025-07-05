VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form mant_clientes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   7665
   ClientLeft      =   315
   ClientTop       =   1185
   ClientWidth     =   10410
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7665
   ScaleWidth      =   10410
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   -20144
      TabIndex        =   36
      Top             =   -16394
      Width           =   4785
      _Version        =   65536
      _ExtentX        =   8440
      _ExtentY        =   741
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Enabled         =   0   'False
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars2 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   7
         Tools           =   "mant_clientes.frx":0000
         ToolBars        =   "mant_clientes.frx":58B4
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   3210
      Left            =   -24779
      TabIndex        =   24
      Top             =   -19304
      Width           =   9510
      Begin VB.TextBox Txtcredis 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox Txtlincre 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2775
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   1035
         Width           =   1320
      End
      Begin VB.TextBox txtventasacum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7590
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtxcobrar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1440
         Width           =   1320
      End
      Begin VB.TextBox txttc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0.000"
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Línea de Crédito en US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   5
         Left            =   495
         TabIndex        =   33
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Crédito Disponible en US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   9
         Left            =   495
         TabIndex        =   32
         Top             =   2115
         Width           =   1860
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ventas Acumuladas en US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   12
         Left            =   5445
         TabIndex        =   31
         Top             =   2115
         Width           =   2040
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cuentas por Cobrar en US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   13
         Left            =   495
         TabIndex        =   30
         Top             =   1485
         Width           =   1995
      End
      Begin VB.Line Line1 
         X1              =   2670
         X2              =   4245
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   6315
         TabIndex        =   29
         Top             =   1080
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   5490
      Left            =   -24374
      TabIndex        =   16
      Top             =   -21689
      Width           =   8970
      Begin VB.TextBox txtemailcoti 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   7
         Top             =   3480
         Width           =   6210
      End
      Begin VB.TextBox txtcomision 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   48
         Top             =   4320
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox txtescuela 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         MaxLength       =   4
         TabIndex        =   8
         Top             =   4275
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtobserva 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   315
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4320
         Width           =   4860
      End
      Begin VB.TextBox txtcodpagxcobrar 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox txtabrev 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   2340
      End
      Begin VB.TextBox txtcobrador 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox TXTDIRCOB 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   6
         Top             =   2340
         Width           =   6210
      End
      Begin VB.TextBox Txtcodedi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         MaxLength       =   13
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox Txtcarleg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1920
         Width           =   3165
      End
      Begin VB.TextBox Txtrepleg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1440
         Width           =   6225
      End
      Begin Threed.SSPanel pnldespagxcobrar 
         Height          =   330
         Left            =   3690
         TabIndex        =   43
         Top             =   480
         Width           =   5010
         _Version        =   65536
         _ExtentX        =   8837
         _ExtentY        =   582
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
      Begin Threed.SSPanel pnlescuela 
         Height          =   330
         Left            =   3690
         TabIndex        =   46
         Top             =   4275
         Visible         =   0   'False
         Width           =   5010
         _Version        =   65536
         _ExtentX        =   8837
         _ExtentY        =   582
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
      Begin Threed.SSCheck chkcomision 
         Height          =   255
         Left            =   5280
         TabIndex        =   54
         Top             =   2880
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Comision"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlcobrador 
         Height          =   330
         Left            =   3675
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         Width           =   5010
         _Version        =   65536
         _ExtentX        =   8837
         _ExtentY        =   582
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
      Begin Threed.SSCheck chkagente 
         Height          =   255
         Left            =   6720
         TabIndex        =   115
         Top             =   2880
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Agente de Retención"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Email para Cotizacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   52
         Top             =   3480
         Width           =   1530
      End
      Begin VB.Label labelcc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   " % "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   3015
         TabIndex        =   50
         Top             =   4455
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label labelcc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Comisión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   49
         Top             =   4440
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   47
         Top             =   3960
         Width           =   1110
      End
      Begin VB.Label labelcc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número de carnet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   45
         Top             =   4005
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago por Cobrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   44
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abreviatura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hora de Recepcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   22
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Direccion de Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   315
         TabIndex        =   21
         Top             =   2385
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Codigo EDI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   6105
         TabIndex        =   20
         Top             =   1980
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cargo Rep. Legal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   315
         TabIndex        =   19
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Label labelcc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Rep. Legal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   315
         TabIndex        =   18
         Top             =   1440
         Width           =   765
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6405
      Left            =   180
      TabIndex        =   13
      Top             =   765
      Width           =   9645
      _Version        =   65536
      _ExtentX        =   17013
      _ExtentY        =   11298
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   360
         TabIndex        =   99
         Top             =   1650
         Visible         =   0   'False
         Width           =   9015
         Begin VB.ComboBox cmbtipvia 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "mant_clientes.frx":599F
            Left            =   1320
            List            =   "mant_clientes.frx":59CA
            TabIndex        =   60
            Top             =   285
            Width           =   975
         End
         Begin VB.TextBox txtint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7800
            TabIndex        =   63
            Top             =   250
            Width           =   1050
         End
         Begin VB.TextBox txtnumvia 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   62
            Top             =   250
            Width           =   735
         End
         Begin VB.TextBox txtnomvia 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   61
            Top             =   250
            Width           =   2175
         End
         Begin VB.TextBox txtcanal 
            DataField       =   "F5NOMPRO"
            DataSource      =   "DataProducto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8880
            MaxLength       =   2
            TabIndex        =   103
            Top             =   5040
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.TextBox txtmes 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4320
            TabIndex        =   102
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cmbsexo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "mant_clientes.frx":5A18
            Left            =   4320
            List            =   "mant_clientes.frx":5A22
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1200
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtdia 
            Height          =   315
            Left            =   3960
            TabIndex        =   100
            Text            =   "Text1"
            Top             =   1200
            Visible         =   0   'False
            Width           =   495
         End
         Begin Threed.SSCheck chkhijos 
            Height          =   240
            Left            =   4440
            TabIndex        =   104
            Top             =   1200
            Visible         =   0   'False
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Tiene Hijo(s)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlcanal 
            Height          =   330
            Left            =   8760
            TabIndex        =   105
            Top             =   4920
            Visible         =   0   'False
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   582
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
         Begin Threed.SSCheck chkescuela 
            Height          =   120
            Left            =   8760
            TabIndex        =   106
            Top             =   5040
            Visible         =   0   'False
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   212
            _StockProps     =   78
            Caption         =   "Escuela de cocina"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCommand cmdsalir 
            Height          =   345
            Left            =   80
            TabIndex        =   64
            ToolTipText     =   "Salir"
            Top             =   150
            Width           =   345
            _Version        =   65536
            _ExtentX        =   617
            _ExtentY        =   617
            _StockProps     =   78
            ForeColor       =   -2147483630
            Picture         =   "mant_clientes.frx":5A3B
         End
         Begin VB.Label Label16 
            Caption         =   "Via Tipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   114
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Via Nombre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   113
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "N°:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   112
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Interior"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7200
            TabIndex        =   111
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Canal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   90
            Left            =   8760
            TabIndex        =   110
            Top             =   5040
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   18
            Left            =   3840
            TabIndex        =   109
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "DD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   19
            Left            =   3480
            TabIndex        =   108
            Top             =   1200
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "MM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   20
            Left            =   3360
            TabIndex        =   107
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.TextBox txtemail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   85
         Top             =   4680
         Width           =   7995
      End
      Begin VB.TextBox txtweb 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   86
         Top             =   5130
         Width           =   7995
      End
      Begin VB.TextBox txtcontacto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1335
         MaxLength       =   100
         TabIndex        =   87
         Top             =   5580
         Width           =   7995
      End
      Begin VB.TextBox txtvendedor 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   82
         Top             =   3720
         Width           =   780
      End
      Begin VB.TextBox txtcodpag 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   81
         Top             =   3720
         Width           =   780
      End
      Begin VB.TextBox Txttelcli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   83
         Top             =   4200
         Width           =   3600
      End
      Begin VB.TextBox Txtfaxcli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   84
         Top             =   4200
         Width           =   3330
      End
      Begin VB.TextBox Txttercon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         MaxLength       =   4
         TabIndex        =   77
         Top             =   3120
         Width           =   1140
      End
      Begin VB.TextBox txtpais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4815
         TabIndex        =   76
         Top             =   3120
         Width           =   1920
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   520
         Left            =   360
         TabIndex        =   80
         Top             =   3000
         Width           =   3570
         Begin VB.OptionButton opttipo 
            Caption         =   "Extranjero"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2040
            TabIndex        =   75
            Top             =   200
            Width           =   1050
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Nacional"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   74
            Top             =   200
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.TextBox Txtzona2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   73
         Top             =   2640
         Width           =   690
      End
      Begin VB.TextBox Txtzona 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   68
         Top             =   2160
         Width           =   690
      End
      Begin VB.TextBox txtdocidentidad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         MaxLength       =   8
         TabIndex        =   70
         Top             =   2160
         Width           =   1260
      End
      Begin VB.TextBox Txtdircli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   120
         TabIndex        =   59
         Top             =   1680
         Width           =   7635
      End
      Begin VB.TextBox Txtcodcli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   65
         Top             =   1200
         Width           =   810
      End
      Begin VB.TextBox Txtnomcli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   58
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtnuevo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   56
         Top             =   1200
         Width           =   1530
      End
      Begin MSDataListLib.DataCombo dbtipo 
         Height          =   330
         Left            =   7440
         TabIndex        =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Data DataTipCliente 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8460
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   765
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   4620
         _Version        =   65536
         _ExtentX        =   8149
         _ExtentY        =   1349
         _StockProps     =   14
         Caption         =   "Persona"
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
         Begin Threed.SSOption Opcruccli 
            Height          =   195
            Left            =   600
            TabIndex        =   53
            Top             =   315
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Jurídica"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption OpcLibcli 
            Height          =   195
            Left            =   3000
            TabIndex        =   55
            Top             =   315
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Natural"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSPanel Txtdeszon 
         Height          =   330
         Left            =   2055
         TabIndex        =   69
         Top             =   2160
         Width           =   4305
         _Version        =   65536
         _ExtentX        =   7594
         _ExtentY        =   582
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
      Begin Threed.SSPanel txtdeszon2 
         Height          =   330
         Left            =   2055
         TabIndex        =   78
         Top             =   2640
         Width           =   7290
         _Version        =   65536
         _ExtentX        =   12859
         _ExtentY        =   582
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
      Begin Threed.SSPanel pnldespag 
         Height          =   330
         Left            =   2160
         TabIndex        =   92
         Top             =   3720
         Width           =   2730
         _Version        =   65536
         _ExtentX        =   4815
         _ExtentY        =   582
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
      Begin Threed.SSPanel pnlvendedor 
         Height          =   330
         Left            =   6840
         TabIndex        =   94
         Top             =   3720
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   582
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
      Begin Threed.SSCommand cmdayudaref 
         Height          =   320
         Left            =   9000
         TabIndex        =   116
         Top             =   1680
         Width           =   330
         _Version        =   65536
         _ExtentX        =   582
         _ExtentY        =   564
         _StockProps     =   78
         Caption         =   "..."
         ForeColor       =   -2147483630
      End
      Begin MSComCtl2.DTPicker TXTFECING 
         Height          =   315
         Left            =   7440
         TabIndex        =   12
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   134676481
         CurrentDate     =   40611
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   98
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sitio Web"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   360
         TabIndex        =   97
         Top             =   5160
         Width           =   675
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   16
         Left            =   360
         TabIndex        =   96
         Top             =   5640
         Width           =   645
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5160
         TabIndex        =   95
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   360
         TabIndex        =   93
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   91
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   5400
         TabIndex        =   90
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "T. Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   7200
         TabIndex        =   89
         Top             =   3120
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   17
         Left            =   4335
         TabIndex        =   88
         Top             =   3120
         Width           =   300
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   79
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   360
         TabIndex        =   72
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Identidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   15
         Left            =   6840
         TabIndex        =   71
         Top             =   2160
         Width           =   1140
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   67
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   7920
         TabIndex        =   66
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   11
         Left            =   360
         TabIndex        =   57
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   1260
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2715
      Left            =   -24869
      OleObjectBlob   =   "mant_clientes.frx":5B45
      TabIndex        =   34
      Top             =   -21959
      Width           =   9510
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   2625
      Left            =   -24869
      OleObjectBlob   =   "mant_clientes.frx":8B07
      TabIndex        =   35
      Top             =   -19139
      Width           =   9510
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid3 
      Height          =   5235
      Left            =   -24824
      OleObjectBlob   =   "mant_clientes.frx":B939
      TabIndex        =   37
      Top             =   -21899
      Width           =   9375
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   -18164
      TabIndex        =   42
      Top             =   -16139
      Width           =   2715
      _Version        =   65536
      _ExtentX        =   4789
      _ExtentY        =   741
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Enabled         =   0   'False
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars3 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   7
         Tools           =   "mant_clientes.frx":E8FB
         ToolBars        =   "mant_clientes.frx":141C7
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "mant_clientes.frx":14274
      ToolBars        =   "mant_clientes.frx":18E50
   End
   Begin VB.Label lblcliente 
      Alignment       =   2  'Center
      Caption         =   "Label10"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   -24779
      TabIndex        =   41
      Top             =   -16454
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Label lblcliente 
      Alignment       =   2  'Center
      Caption         =   "Label10"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   -24914
      TabIndex        =   40
      Top             =   -15944
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Label lblcliente 
      Alignment       =   2  'Center
      Caption         =   "Label10"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   -24914
      TabIndex        =   39
      Top             =   -15974
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label lblcliente 
      Alignment       =   2  'Center
      Caption         =   "Label10"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   -24689
      TabIndex        =   38
      Top             =   -16079
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "mant_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim wcodigo             As String * 4
Dim wtipcli             As String * 1
Dim sw                  As Integer
Dim sw_tab              As Boolean
Dim sw_ayuda            As Boolean
Dim sw_verifica         As Boolean
Dim CLIENTE             As String
Dim sql                 As String
Dim wgraba              As Integer
Dim cconex_formcli      As String
Dim cnn_formcli         As New ADODB.Connection
Dim cnombasecli         As String
Dim dbtablecli          As String
Dim ctablesuc           As String
Dim ctablelista         As String
Dim csql                As String
Dim rstemporal          As New ADODB.Recordset
Dim RSCONSULTA2         As New ADODB.Recordset
Dim est                 As Byte
Dim whijos              As String * 1
Dim wclicom             As Double
Dim wdia                As String
Dim wmes                As String
Dim wsexo               As String * 1
Dim indCli              As Boolean
Dim wfind               As Boolean
Dim Rs                  As New ADODB.Recordset

Private Sub Actualiza_Cliente(pcodcli As String)
Dim rsescuela   As New ADODB.Recordset
    If rsclientes.State = adStateOpen Then rsclientes.Close
    rsclientes.Open "SELECT * FROM EF2CLIENTES WHERE F2CODCLI='" & pcodcli & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsclientes.EOF Then
        TXTFECING.value = IIf(IsNull(rsclientes.Fields("fecing")), Format(Date, "DD/MM/YYYY"), Format(rsclientes.Fields("fecing"), "DD/MM/YYYY"))
        Txtcodcli.Text = "" & rsclientes.Fields("F2CODCLI")
        Txtcodedi.Text = "" & rsclientes.Fields("F2CODEDI")
        Txtnomcli.Text = "" & rsclientes.Fields("F2NOMCLI")
        txtabrev.Text = "" & rsclientes.Fields("F2ABREVCLI")
        Txtdircli.Text = "" & rsclientes.Fields("F2DIRCLI")
        cmbtipvia.Text = "" & rsclientes.Fields("VIA_TIPO")
        txtnomvia.Text = "" & rsclientes.Fields("VIA_NOMBRE")
        txtnumvia.Text = "" & rsclientes.Fields("VIA_NUM")
        txtint.Text = "" & rsclientes.Fields("VIA_INT")
        txtobserva.Text = "" & rsclientes.Fields("F2OBSERVA")
        txtrecepcio.value = "" & rsclientes.Fields("F2DIRENT")
        
        wnompag = ""
        txtcodpag.Text = "" & rsclientes.Fields("F2FORPAG")
        If VALIDA_FPAGO(txtcodpag.Text) = True Then
            pnldespag.Caption = wnompag
        End If
        
        wnompag = ""
        txtcodpagxcobrar.Text = "" & rsclientes.Fields("F2FORPAG_XCOBRAR")
        If VALIDA_FPAGO(txtcodpagxcobrar.Text) = True Then
            pnldespagxcobrar.Caption = wnompag
        End If
        
        Txtzona.Text = "" & rsclientes.Fields("F2ZONA")
        Txtzona2.Text = "" & rsclientes.Fields("F2ZONA_CLIENTE")
        sql = "SELECT F2DESZON FROM EF2ZONAS WHERE F2CODZON= '" & Txtzona.Text & "'"
        'SQL = "SELECT DISTRITO FROM EF2ZONAS_CLIENTE WHERE CODIGO= '" & Txtzona.Text & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_dbbancos, adOpenDynamic
        If Not rst.EOF Then
            Txtdeszon.Caption = "" & rst.Fields("f2deszon")
        Else
            Txtdeszon.Caption = ""
        End If
        rst.Close
        sql = "SELECT DISTRITO FROM EF2ZONAS_CLIENTE WHERE CODIGO= '" & Txtzona2.Text & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_dbbancos, adOpenDynamic
        If Not rst.EOF Then
            txtdeszon2.Caption = "" & rst.Fields("DISTRITO")
        Else
            txtdeszon2.Caption = ""
        End If
        rst.Close
        Txttelcli.Text = "" & rsclientes.Fields("F2TELCLI")
        Txtfaxcli.Text = "" & rsclientes.Fields("F2FAXCLI")
        Txtlincre.Text = Format(Val("" & rsclientes.Fields("F2LINCRE")), "###,###,##0.00")
        txtnuevo.Text = "" & rsclientes.Fields("F2NEWRUC")
        Txttercon.Text = "" & rsclientes.Fields("TERCON")
        Txtrepleg.Text = "" & rsclientes.Fields("F2REPLEG")
        Txtcarleg.Text = "" & rsclientes.Fields("F2CARLEG")
        indCli = False
        If rsclientes.Fields("F2TIPDOC") = "N" Then
            OpcLibcli.value = True
            txtdia.Text = Format("" & rsclientes.Fields("F2DIA"), "00")
            txtmes.Text = Format("" & rsclientes.Fields("F2MES"), "00")
            cmbsexo.ListIndex = IIf("" & rsclientes.Fields("f2sexo") = "M", 0, 1)
            chkhijos.value = IIf("" & rsclientes.Fields("f2hijos") = "S", True, False)
        End If
        If rsclientes.Fields("F2TIPDOC") = "J" Then
            OpcLibcli.value = False
            Opcruccli.value = True
        End If
        indCli = True
        txtemail.Text = "" & rsclientes.Fields("F2EMAIL")
        Txtweb.Text = "" & rsclientes.Fields("F2WEB")
'        TXTDIRCOB.Text = "" & rsclientes.Fields("F2DIRCOB")
        TXTDIRCOB.Text = "" & rsclientes.Fields("F2LUGENT")
        txtpais.Text = "" & rsclientes.Fields("F2PAIS")
        txtemailcoti.Text = "" & rsclientes.Fields("EMAILCOTI")
'        txtclicom.Text = "" & rsclientes.Fields("F2CLICOM")
        chkcomision.value = IIf("" & rsclientes.Fields("F2CLICOM") = 1, False, True)
        dbtipo.BoundText = "" & rsclientes.Fields("F2TIPOCLIE")
'        dbtipo.Text = "" & traerCampo("EF2TIPOS", "DESTIPCLIE", "CODTIPCLIE", "" & rsclientes.Fields("F2TIPOCLIE"))
        wnomven = ""
        txtvendedor.Text = "" & rsclientes.Fields("F2CODVEN")
        If Len(Trim(txtvendedor.Text)) > 0 Then
            If VALIDA_VENDEDOR(txtvendedor.Text) = True Then
                pnlvendedor.Caption = wnomven
            Else
                pnlvendedor.Caption = ""
            End If
        End If
        
        wnomven = ""
        txtcobrador.Text = "" & rsclientes.Fields("F2CODCOB")
        If Len(Trim(txtcobrador.Text)) > 0 Then
            If VALIDA_COBRADOR(txtcobrador.Text) = True Then
                pnlcobrador.Caption = wnomven
            Else
                pnlcobrador.Caption = ""
            End If
        End If
        txtdocidentidad.Text = Trim("" & rsclientes.Fields("F2DOCIDENTIDAD"))
        
        txtcanal.Text = "" & rsclientes.Fields("F2CANAL")
        If Len(Trim(txtcanal.Text)) > 0 Then
            If VALIDA_CANAL(txtcanal.Text) = True Then
                pnlcanal.Caption = wnomcanal
            Else
                pnlcanal.Caption = ""
                MsgBox "Código de canal no existe. Verifique.", vbInformation + vbDefaultButton1, "Atención"
                txtcanal.Text = "": txtcanal.SetFocus
            End If
        Else
            pnlcanal.Caption = ""
        End If
        txtcontacto.Text = Trim("" & rsclientes.Fields("F2CONTACTO"))
        If Trim("" & rsclientes.Fields("F2ESCUELA")) = "*" Then
            chkescuela.value = True
        Else
            chkescuela.value = False
        End If
'        txtcarnet.Text = Trim("" & rsclientes.Fields("F2CARNET"))
        opttipo(0).value = IIf(rsclientes.Fields("F2TIPOCLI") = "N", True, False)
        opttipo(1).value = IIf(rsclientes.Fields("F2TIPOCLI") = "E", True, False)
        
        txtescuela.Text = Trim("" & rsclientes.Fields("F2CODESCUELA"))
        
        If Trim("" & rsclientes.Fields("F2AGENTE")) = "*" Then
            chkagente.value = True
        Else
            chkagente.value = False
        End If
        
        If Len(Trim(txtescuela.Text)) > 0 Then
            If rsescuela.State = adStateOpen Then rsescuela.Close
            rsescuela.Open "SELECT F2NOMCLI FROM EF2CLIENTES WHERE F2CODCLI='" & txtescuela.Text & "'", cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rsescuela.EOF Then
                pnlescuela.Caption = Trim("" & rsescuela.Fields("F2NOMCLI"))
            End If
            rsescuela.Close
            Set rsescuela = Nothing
        End If
        
        txtcomision.Text = Val("" & rsclientes.Fields("F2COMISION"))
    Else
        txtnuevo.Enabled = True
        MsgBox "Código de cliente no existe. Verifique.", vbInformation, "Atención"
    End If
    rsclientes.Close
    wgraba = 0
End Sub

Private Function calcula_codigo()
Dim wnum2   As Integer

'    SQL = "select f2codcli from ef2clientes  order by f2codcli DESC"
'    If rst.State = adStateOpen Then rst.Close
'    rst.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    If Not rst.EOF Then
'        If rst.Fields("F2CODCLI") & "" = "9999" Then
'            rst.MoveNext
'            If rst.EOF Then
'                wnum2 = 1
'            Else
'                wnum2 = Val(rst.Fields("F2CODCLI") & "") + 1
'            End If
'        Else
'            wnum2 = Val(rst.Fields("F2CODCLI") & "") + 1
'        End If
'    Else
'        wnum2 = 1
'    End If
'    rst.Close
'    wcodigo = Format(wnum2, "0000")
'    Calcula_Codigo = wcodigo
'    Txttercon.Text = wcodigo

    sql = "select CLIENTES from CORRELATIVOS"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        wnum2 = Val(rst.Fields("CLIENTES") & "") + 1
    Else
        wnum2 = 1
    End If
    If Busca_codigo(Format(wnum2, "0000")) Then   ' Si existe en la tabla de clientes
        wcodigo = Obtener_Nuevo_Codigo(wLocal)
    Else
        wcodigo = Format(wnum2, "0000")
    End If
    rst.Close
    'wcodigo = Format(wnum2, "0000")
    calcula_codigo = wcodigo
    Txttercon.Text = wcodigo
End Function

Function Obtener_Nuevo_Codigo(plocal) As String
Dim sql As String
Dim rsbusca As New ADODB.Recordset
If plocal = "C" Then
   sql = "SELECT f2codcli FROM EF2CLIENTES where f2codcli NOT IN ('9999','8888')order by f2codcli desc"
   rsbusca.Open sql, cnn_dbbancos, adOpenDynamic, adLockReadOnly
   If Not rsbusca.EOF Then
      Obtener_Nuevo_Codigo = Format(Val("" & rsbusca!f2codcli) + 1, "0000")
   End If
Else
   sql = "SELECT f2codcli FROM EF2CLIENTES where f2codcli < '7000' order by f2codcli desc"
   rsbusca.Open sql, cnn_dbbancos, adOpenDynamic, adLockReadOnly
   If Not rsbusca.EOF Then
      Obtener_Nuevo_Codigo = Format(Val("" & rsbusca!f2codcli) + 1, "0000")
   End If
End If
rsbusca.Close: Set rsbusca = Nothing
End Function

Function Busca_codigo(pcodigo) As Boolean
Dim rsbusca  As New ADODB.Recordset
Dim sql As String
  sql = "Select f2codcli from ef2clientes where f2codcli='" & pcodigo & "'"
  rsbusca.Open sql, cnn_dbbancos, adOpenDynamic, adLockReadOnly
  If Not rsbusca.EOF Then
    Busca_codigo = True
  Else
    Busca_codigo = False
  End If
  rsbusca.Close: Set rsbusca = Nothing
End Function

Private Sub Elimina_Cliente()
    Beep
    If MsgBox("Está seguro de eliminar el cliente", 36, "Atención") = 6 Then
        sql = "SELECT F2CODCLI FROM EF2CLIENTES WHERE F2CODCLI= '" & Txtcodcli & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
        If Not rst.EOF Then
            sql = "DELETE * FROM EF2CLIENTES WHERE F2CODCLI = '" & Trim("" & Txtcodcli.Text) & "'"
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            
            If wIndEnvia = "*" Then  'cnn_dbEnvia
                cnn_dbEnvia.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbEnvia
            End If
        Else
            Beep
        End If
        rst.Close
        Nuevo_Cliente
    End If
End Sub

Private Sub chkhijos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Txttelcli.SetFocus
End If
End Sub

Private Sub cmbsexo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    chkhijos.SetFocus
End If
End Sub

Private Sub cmbtipvia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub cmdayudaref_Click()
Frame1.Visible = True
End Sub

Private Sub cmdsalir_Click()
Frame1.Visible = False
Txtdircli.SetFocus
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.FocusedIndex = 0
        End If
    End If
End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns(0).value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid1.Dataset.Refresh
        End If
    End If
    
End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName
        Case "CODIGO"
            gcodalm = "01"
            gcodpro = "": gnompro = "": gcodmar = "": gcodgrupo = ""
            gvalvta = 0#: gcodtipo = 0: wfactor = 0#
            wcodcanal = txtcanal.Text
            If Len(Trim(wcodcanal)) > 0 Then
                hlp_productos_canales.Show 1
            Else
                sw_ayuda_prod = True
                ayuda_productos.Show 1
            End If
            If Len(Trim(gcodpro)) > 0 Then
                csql = "SELECT CODIGO FROM " & dbtablecli & " WHERE CODIGO='" & gcodpro & "'"
                If rstemporal.State = adStateOpen Then rstemporal.Close
                rstemporal.Open csql, cnn_formcli, adOpenDynamic, adLockOptimistic
                If rstemporal.EOF Then
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value = gcodpro
                    dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = gnompro
                    dxDBGrid1.Columns.ColumnByFieldName("UNIDAD").value = gcodmar
                    dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = wafecto
                    dxDBGrid1.Columns.ColumnByFieldName("PRECIOUNIT").value = Format(gvalvta, "###,##0.00")
                    dxDBGrid1.Dataset.Post
                Else
                    MsgBox "El producto ya ha sido ingresado. Verifique.", vbInformation, "Atención"
                End If
                rstemporal.Close
            End If
            dxDBGrid1.Columns.FocusedIndex = 0
    End Select
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Inventario") = vbYes Then
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
        End If
    End If
End Sub

Private Sub dxDBGrid2_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim ccodigo_sucursal    As String
    If Action = daInsert Then
        ccodigo_sucursal = OBTIENE_CODSUCURSAL()
        dxDBGrid2.Columns.ColumnByFieldName("ITEM").value = dxDBGrid2.Dataset.RecordCount + 1
        dxDBGrid2.Columns.ColumnByFieldName("F2CODSUCURSAL").value = ccodigo_sucursal
        dxDBGrid2.Columns.FocusedIndex = 0
    End If
End Sub

Private Sub dxDBGrid2_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If dxDBGrid2.Dataset.RecordCount > 0 Then
            If Len(Trim(dxDBGrid2.Columns.ColumnByFieldName("F2CODSUCURSAL").value & "")) = 0 Then
                Allow = False
            End If
        End If
    End If
        
End Sub

Private Sub dxDBGrid2_OnDblClick()
Dim ccodsucursal        As String

    ccodsucursal = Trim("" & dxDBGrid2.Columns.ColumnByFieldName("F2CODSUCURSAL").value)
    '---------------------------------------------------------------
    If RSCONSULTA2.State = adStateOpen Then RSCONSULTA2.Close
    RSCONSULTA2.Open "Select * from IF4PRODCLI where F2CODCLI='" & Txtcodcli.Text & "' AND F2CODSUCURSAL='" & ccodsucursal & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If RSCONSULTA2.EOF Then
        AdicionaItem
    Else
        Actualiza_Productos Txtcodcli.Text, ccodsucursal
    End If
    RSCONSULTA2.Close
    
End Sub

Private Sub dxDBGrid2_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dxDBGrid2.Columns.FocusedColumn.FieldName
        Case "F2CODVENDEDOR"
            wcodven = "": wnomven = ""
            ayuda_vendedores.Show 1
            If Len(Trim(wcodven)) > 0 Then
                dxDBGrid2.Dataset.Edit
                dxDBGrid2.Columns.ColumnByFieldName("F2CODVENDEDOR").value = wcodven
                dxDBGrid2.Dataset.Post
            End If
            dxDBGrid2.Columns.FocusedIndex = 0
    End Select
End Sub

Private Sub dxDBGrid2_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case 115:
            If MsgBox("Desea eliminar el registro actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
                sw_detalle = True
                If dxDBGrid2.Dataset.RecNo = 1 Then
                    dxDBGrid2.Dataset.Delete
                    AdicionaItem2
                Else
                    dxDBGrid2.Dataset.Delete
                End If
            End If
    End Select
End Sub

Private Sub dxDBGrid3_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid3.Columns.ColumnByFieldName("ITEM").value = dxDBGrid3.Dataset.RecordCount + 1
            dxDBGrid3.Columns.FocusedIndex = 0
        End If
    End If
End Sub

Private Sub dxDBGrid3_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid3.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid3.Columns(0).value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid3.Dataset.Refresh
        End If
    End If
End Sub

Private Sub dxDBGrid3_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dxDBGrid3.Columns.FocusedColumn.FieldName
        Case "CODIGO"
            gcodalm = "01"
            gcodpro = "": gnompro = "": gcodmar = "": gcodgrupo = ""
            gvalvta = 0#: gcodtipo = 0: wfactor = 0#
            wcodcanal = txtcanal.Text
            If Len(Trim(wcodcanal)) > 0 Then
                hlp_productos_canales.Show 1
            Else
                sw_ayuda_prod = True
                ayuda_productos.Show 1
            End If
            If Len(Trim(gcodpro)) > 0 Then
                csql = "SELECT CODIGO FROM " & ctablelista & " WHERE CODIGO='" & gcodpro & "'"
                If rstemporal.State = adStateOpen Then rstemporal.Close
                rstemporal.Open csql, cnn_formcli, adOpenDynamic, adLockOptimistic
                If rstemporal.EOF Then
                    dxDBGrid3.Dataset.Edit
                    dxDBGrid3.Columns.ColumnByFieldName("CODIGO").value = gcodpro
                    dxDBGrid3.Columns.ColumnByFieldName("DESCRIPCION").value = gnompro
                    dxDBGrid3.Columns.ColumnByFieldName("UNIDAD").value = gcodmar
                    dxDBGrid3.Columns.ColumnByFieldName("AFECTO").value = wafecto
                    dxDBGrid3.Columns.ColumnByFieldName("PRECIOUNIT").value = Format(gvalvta, "###,##0.00")
                    dxDBGrid3.Dataset.Post
                Else
                    MsgBox "El producto ya ha sido ingresado. Verifique.", vbInformation, "Atención"
                End If
                rstemporal.Close
            End If
            dxDBGrid3.Columns.FocusedIndex = 0
    End Select
End Sub

Private Sub dxDBGrid3_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            If dxDBGrid3.Dataset.RecNo = 1 Then
                dxDBGrid3.Dataset.Delete
                AdicionaItem3
            Else
                dxDBGrid3.Dataset.Delete
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
'txtnuevo.SetFocus
End Sub

Private Sub Form_Load()
    Me.Height = 8115
    Me.Width = 10530
    Me.left = 1500
    Me.top = 980
    
    cnombasecli = "templus.mdb" '"TMP_CLI.MDB"
    cconex_formcli = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombasecli & ";Persist Security Info=False"
    If cnn_formcli.State = adStateOpen Then cnn_formcli.Close
    cnn_formcli.Open cconex_formcli
    
    dbtablecli = "DOCPRODCLI"
    ctablesuc = "Tempsucursales"
    ctablelista = "TBLISTAESP"
    
    sw_tab = True
    sw_ayuda = False
        
    If sw_load_mant = True Then
        SSActiveToolBars1.Tools.ITEM("ID_Lista").Visible = False
        SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Visible = False
        SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Visible = False
        SSActiveToolBars1.Tools.ITEM("ID_Salir").Visible = True
    Else
        SSActiveToolBars1.Tools.ITEM("ID_Lista").Visible = True
        SSActiveToolBars1.Tools.ITEM("ID_Salir").Visible = False
    End If
    
    wcodigo = ""
    If Rs.State = adStateOpen Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open "SELECT * FROM EF2TIPOS", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    Set dbtipo.RowSource = Rs.DataSource
    dbtipo.ListField = "destipclie"
    dbtipo.BoundColumn = "codtipclie"
    If sw_nuevo_mant = True Then
        Nuevo_Cliente
        wgraba = 1
        Txtnomcli.Text = wnomcli
        txtnuevo.Text = wruccli
        Txtdircli.Text = wdireccion
    Else
        CLIENTE = ayuda_clientes.dxDBGrid1.Columns(0).value
        Actualiza_Cliente ayuda_clientes.dxDBGrid1.Columns(0).value
    End If
    
    Conf_Grid
    Conf_Grid2
    Conf_Grid3
    
    fpTabProADO1.TabCount = 2
'    Opcruccli.Value = True
End Sub

Private Sub Graba_Cliente()
On Error GoTo graba
Dim cescuela    As String
    
    ReDim amovs(0 To 47) As a_grabacion
'    If Trim(Txtcodcli.Text) = "" Then
'        Txtcodcli.Text = Calcula_Codigo
'    End If
    
    wcodigo = "" & Trim(Txtcodcli.Text)
    wcodgenerado = "" & Trim(Txtcodcli.Text)
    wparam_ruc = txtnuevo.Text
    
    sql = "select F2CODCLI from ef2clientes where f2codcli = '" & wcodigo & "'"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        sw = 0
    Else
        sw = 1
    End If
    rst.Close
    
    amovs(0).campo = "F2CODCLI": amovs(0).valor = wcodigo: amovs(0).Tipo = "T"
    amovs(1).campo = "F2NOMCLI": amovs(1).valor = Trim("" & Txtnomcli.Text): amovs(1).Tipo = "T"
    amovs(2).campo = "F2DIRCLI": amovs(2).valor = Trim("" & Txtdircli.Text): amovs(2).Tipo = "T"
    amovs(3).campo = "F2ZONA": amovs(3).valor = Trim("" & Txtzona.Text): amovs(3).Tipo = "T"
    amovs(4).campo = "F2TELCLI": amovs(4).valor = Trim("" & Txttelcli.Text): amovs(4).Tipo = "T"
    amovs(5).campo = "F2FAXCLI": amovs(5).valor = Trim("" & Txtfaxcli.Text): amovs(5).Tipo = "T"
    If OpcLibcli.value = True Then
        amovs(6).campo = "F2TIPDOC": amovs(6).valor = "N": amovs(6).Tipo = "T"
    End If
    If Opcruccli.value = True Then
        amovs(6).campo = "F2TIPDOC": amovs(6).valor = "J": amovs(6).Tipo = "T"
    End If
    amovs(7).campo = "F2EMAIL": amovs(7).valor = Trim("" & txtemail.Text): amovs(7).Tipo = "T"
    amovs(8).campo = "F2NEWRUC": amovs(8).valor = Trim("" & txtnuevo.Text): amovs(8).Tipo = "T"
    amovs(9).campo = "TERCON": amovs(9).valor = Trim("" & Txttercon.Text): amovs(9).Tipo = "T"
    amovs(10).campo = "F2LINCRE": amovs(10).valor = Val(Format(Txtlincre.Text, "0.00")): amovs(10).Tipo = "N"
    amovs(11).campo = "F2WEB": amovs(11).valor = Trim("" & Txtweb.Text): amovs(11).Tipo = "T"
    amovs(12).campo = "F2REPLEG": amovs(12).valor = Trim("" & Txtrepleg.Text): amovs(12).Tipo = "T"
    amovs(13).campo = "F2CARLEG": amovs(13).valor = Trim("" & Txtcarleg.Text): amovs(13).Tipo = "T"
    amovs(14).campo = "F2CODEDI": amovs(14).valor = Trim("" & Txtcodedi.Text): amovs(14).Tipo = "T"
    amovs(15).campo = "F2DIRCOB": amovs(15).valor = "": amovs(15).Tipo = "T"
    amovs(16).campo = "FECING": amovs(16).valor = TXTFECING.value: amovs(16).Tipo = "F"
    amovs(17).campo = "F2ABREVCLI": amovs(17).valor = Trim("" & txtabrev.Text): amovs(17).Tipo = "T"
    amovs(18).campo = "F2FORPAG": amovs(18).valor = Trim("" & txtcodpag.Text): amovs(18).Tipo = "T"
    amovs(19).campo = "F2CODVEN": amovs(19).valor = Trim("" & txtvendedor.Text): amovs(19).Tipo = "T"
    amovs(20).campo = "F2CODCOB": amovs(20).valor = Trim("" & txtcobrador.Text): amovs(20).Tipo = "T"
    amovs(21).campo = "F2DISCLI": amovs(21).valor = Trim(left("" & Txtdeszon.Caption, 50)): amovs(21).Tipo = "T"
    amovs(22).campo = "F2DOCIDENTIDAD": amovs(22).valor = Trim("" & txtdocidentidad.Text): amovs(22).Tipo = "T"
    amovs(23).campo = "F2CANAL": amovs(23).valor = Trim("" & txtcanal.Text): amovs(23).Tipo = "T"
    amovs(24).campo = "F2CONTACTO": amovs(24).valor = Trim("" & txtcontacto.Text): amovs(24).Tipo = "T"
    amovs(25).campo = "F2FORPAG_XCOBRAR": amovs(25).valor = Trim("" & txtcodpagxcobrar.Text): amovs(25).Tipo = "T"
    wdia = ""
    If Val(txtdia.Text) > 0 Then
'        wdia = txtdia.Text
    End If
    amovs(26).campo = "F2DIA": amovs(26).valor = wdia: amovs(26).Tipo = "T"
    wmes = ""
    If Val(txtmes.Text) > 0 Then
        wmes = txtmes.Text
    End If
    amovs(27).campo = "F2MES": amovs(27).valor = wmes: amovs(27).Tipo = "T"
    If cmbsexo.ListIndex = 0 Then
        wsexo = "M"
    ElseIf cmbsexo.ListIndex = 1 Then
        wsexo = "F"
    Else
        wsexo = ""
    End If
    amovs(28).campo = "F2SEXO": amovs(28).valor = wsexo: amovs(28).Tipo = "T"
    If chkhijos.Enabled Then
        If chkhijos.value Then
            whijos = "S"
        Else
            whijos = "N"
        End If
    Else
        whijos = ""
    End If
    amovs(29).campo = "F2HIJOS": amovs(29).valor = whijos: amovs(29).Tipo = "T"
    cescuela = ""
    If chkescuela.value = True Then
        cescuela = "*"
    Else
        cescuela = ""
    End If
    amovs(30).campo = "F2ESCUELA": amovs(30).valor = cescuela: amovs(30).Tipo = "T"
    amovs(31).campo = "F2OBSERVA": amovs(31).valor = txtobserva.Text: amovs(31).Tipo = "T"
    amovs(32).campo = "F2CODESCUELA": amovs(32).valor = txtescuela.Text: amovs(32).Tipo = "T"
    amovs(33).campo = "F2COMISION": amovs(33).valor = txtcomision.Text: amovs(33).Tipo = "N"
    amovs(34).campo = "F2ZONA_CLIENTE": amovs(34).valor = Trim("" & Txtzona2.Text): amovs(34).Tipo = "T"
    amovs(35).campo = "F2TIPOCLIE": amovs(35).valor = Trim("" & dbtipo.BoundText): amovs(35).Tipo = "T"
    amovs(36).campo = "VIA_TIPO": amovs(36).valor = Trim("" & cmbtipvia.Text): amovs(36).Tipo = "T"
    amovs(37).campo = "VIA_NOMBRE": amovs(37).valor = Trim("" & txtnomvia.Text): amovs(37).Tipo = "T"
    amovs(38).campo = "VIA_NUM": amovs(38).valor = Trim("" & txtnumvia.Text): amovs(38).Tipo = "T"
    amovs(39).campo = "VIA_INT": amovs(39).valor = Trim("" & txtint.Text): amovs(39).Tipo = "T"
    amovs(40).campo = "F2DIRENT": amovs(40).valor = Trim("" & txtrecepcio.Text): amovs(40).Tipo = "T"
    amovs(41).campo = "F2LUGENT": amovs(41).valor = Trim("" & TXTDIRCOB.Text): amovs(41).Tipo = "T"
    amovs(42).campo = "F2PAIS": amovs(42).valor = Trim("" & txtpais.Text): amovs(42).Tipo = "T"
    amovs(43).campo = "F4USEGRA": amovs(43).valor = Trim("" & wusuario): amovs(43).Tipo = "T"
    amovs(44).campo = "EMAILCOTI": amovs(44).valor = Trim("" & txtemailcoti.Text): amovs(44).Tipo = "T"
    
    If chkcomision.value Then
        wclicom = 0.5
    Else
        wclicom = 1
    End If
    amovs(45).campo = "F2CLICOM": amovs(45).valor = Trim("" & wclicom): amovs(45).Tipo = "N"
    
    amovs(46).campo = "F2TIPOCLI"
    If opttipo(0).value = True Then
        amovs(46).valor = "N": amovs(46).Tipo = "T"
    End If
    If opttipo(1).value = True Then
        amovs(46).valor = "E": amovs(46).Tipo = "T"
    End If
    
    amovs(47).campo = "F2AGENTE": amovs(47).valor = IIf(chkagente.value = True, "*", " "): amovs(47).Tipo = "T"
    
    If sw = 1 Then
        GRABA_REGISTRO_logistica amovs(), "EF2CLIENTES", "A", 47, cnn_dbbancos, ""
        
        'actu correlativo
        sql = "update CORRELATIVOS set CLIENTES = '" & wcodigo & "'"
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
        
        sql = "update CORRELATIVOS set CLIENTES = '" & wcodigo & "'"
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
       
    
    Else
        GRABA_REGISTRO_logistica amovs(), "EF2CLIENTES", "M", 47, cnn_dbbancos, "F2CODCLI = '" & wcodigo & "'"
    End If
    
    If wIndEnvia = "*" Then  'cnn_dbEnvia
    
        csql = "delete from EF2CLIENTES where F2CODCLI = '" & wcodigo & "'"
        cnn_dbEnvia.Execute csql
        'AlmacenaQuery_sql csql, cnn_dbEnvia
        
        GRABA_REGISTRO_logistica amovs(), "EF2CLIENTES", "A", 47, cnn_dbEnvia, ""
    End If
    
    Txtcodcli.Enabled = False
    MsgBox "El Cliente Ha Sido Actualizado", vbInformation, "Atención"
    Exit Sub
    
graba:
    If Err = 3186 Then
        For i% = 1 To 10000
        Next i%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If
End Sub

Private Sub Nuevo_Cliente()
    Txtcodcli.Enabled = True
    Txtcodcli.Text = calcula_codigo
    'Actualiza_Cliente Txtcodcli
    txtnuevo.Enabled = True
    txtnuevo.Text = ""
    Txtnomcli.Text = ""
    Txtdircli.Text = ""
    Txtzona.Text = "": Txtdeszon.Caption = ""
    Txtzona2.Text = "": txtdeszon2.Caption = ""
    Txttelcli.Text = ""
    Txtfaxcli.Text = ""
    txtemail.Text = ""
    Txtweb.Text = ""
    txtcodpag.Text = "": pnldespag.Caption = ""
    txtcodpagxcobrar.Text = "": pnldespagxcobrar.Caption = ""
    txtabrev.Text = ""
    Txtrepleg.Text = ""
    Txtcodedi.Text = ""
    TXTDIRCOB.Text = ""
    txtdia.Text = ""
    txtpais.Text = ""
    txtvendedor.Text = "": pnlvendedor.Caption = ""
    txtcobrador.Text = "": pnlcobrador.Caption = ""
    Txtlincre.Text = "0.00"
    txtxcobrar.Text = "0.00"
    Txtcredis.Text = "0.00"
    txttc.Text = "0.00"
    txtventasacum.Text = "0.00"
    Txttercon.Text = Txtcodcli.Text
    Txtcarleg.Text = ""
    Opcruccli.value = True
    TXTFECING.value = Format(Date, "DD/MM/YYYY")
    txtdocidentidad.Text = ""
    chkescuela.Visible = False
    txtescuela.Text = "": pnlescuela.Caption = ""
    txtobserva.Text = ""
    txtcontacto.Text = ""
    txtcomision.Text = "0"
    cmbtipvia.Text = ""
    txtnomvia.Text = ""
    txtnumvia.Text = ""
    txtint.Text = ""
    dbtipo.Text = ""
    txtrecepcio.value = ""
    chkcomision.value = False
    opttipo(0).value = True
    opttipo_Click (0)
    AdicionaItem
    AdicionaItem2
    AdicionaItem3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rst As ADODB.Recordset
Dim cnn As New ADODB.Connection

    cnn_formcli.Close
    dxDBGrid1.Dataset.Close
    dxDBGrid2.Dataset.Close
    dxDBGrid3.Dataset.Close

    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    Else
        Set cnn = New ADODB.Connection
        Set rst = New ADODB.Recordset
        
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos.mdb ;Persist Security Info=False"
        
        sql = "select f2codcli from ef2clientes where f2codcli='" & wcodcli & "'"
        rst.Open sql, cnn, adOpenStatic, adLockOptimistic
        If rst.EOF Then
            wcodcli = ""
            wnomcli = ""
            wruccli = ""
            WDIRCLI = ""
            wforpag = ""
            wtipocli = ""
            wcontacto = ""
            wdocidentidad = ""
        End If
    End If
End Sub

Private Sub fpTabProADO1_TabActivate(TabToActivate As Integer)
Dim ntc As Double

    If TabToActivate <> 0 Then
        lblcliente(TabToActivate).Caption = Txtnomcli.Text
        lblcliente(TabToActivate).Visible = True
    End If

    If TabToActivate = 2 Then
        If sw_tab = True Then
            sw_tab = False
            If rscambios.State = adStateOpen Then rscambios.Close
            rscambios.Open "SELECT * FROM CAMBIOS WHERE FECHA= CVDate( '" & Format(Date, "DD/MM/YYYY") & "' )", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rscambios.EOF Then
                If wpartipcam = "V" Then
                    ntc = Val("" & rscambios.Fields("CAMBIO_VENTA"))
                Else
                    ntc = Val("" & rscambios.Fields("CAMBIO"))
                End If
            Else
                ntc = 2.65
            End If
            rscambios.Close
            If ntc = 0# Then ntc = 2.65
            txttc.Text = Format(ntc, "0.000")
            EVALUA_DATOS_FINANCIEROS Txtcodcli.Text, ntc
        End If
    End If
End Sub

Private Sub OpcLibcli_Click(value As Integer)
    If sw_load_mant = False Then '- lista
        If sw_nuevo_mant = True Then
            If value = True Then
                txtnuevo.Text = Format(Txtcodcli.Text, "00000000000")
            End If
        End If
    Else
        If gtipodocu = "B" Or gtipodocu = "P" Then
            If sw_nuevo_mant = True Then
                If value = True Then
                    txtnuevo.Text = Format(Txtcodcli.Text, "00000000000")
                End If
            End If
        End If
    End If
'    If sw_nuevo_mant = True Then
'         txtnuevo.SetFocus
'    End If
    est = 1
    txtdia.Enabled = True
    txtmes.Enabled = True
    cmbsexo.Enabled = True
    chkhijos.Enabled = True
    cmbsexo.ListIndex = 0
End Sub

Private Sub OpcLibcli_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtnuevo.SetFocus
End Sub

Private Sub Opcruccli_Click(value As Integer)
est = 2
'txtdia.Enabled = False
txtdia.Text = ""
txtmes.Enabled = False
txtmes.Text = ""
cmbsexo.Enabled = False
chkhijos.value = False
chkhijos.Enabled = False
cmbsexo.ListIndex = -1

'If indCli = False Then Exit Sub
If txtnuevo.Text <> "" Then
    If Val(left(txtnuevo.Text, 2)) = "00" And Opcruccli.value Then
        MsgBox "Ingrese el Ruc Correcto", vbInformation, "Sistema de Ventas"
        Opcruccli_KeyPress 13
'        OpcLibcli.Value = True
    End If
End If
End Sub

Private Sub Opcruccli_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtnuevo.SetFocus
End Sub

Private Sub opttipo_Click(Index As Integer)
If sw_nuevo_mant = True Then
    If opttipo(0).value Then
      txtpais.Text = "PERU"
    Else
      txtpais.Text = ""
'      txtpais.SetFocus
    End If
Else
    If opttipo(0).value And Len(Trim(txtpais)) = 0 Then
      txtpais.Text = "PERU"
    End If
End If
End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then txtpais.SetFocus
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ERRORCLI
    Select Case Tool.Id
        Case "ID_Nuevo":
            Me.MousePointer = vbHourglass
            If wgraba = 1 And Len(Trim(txtnuevo)) <> 0 Then
                If MsgBox("No ha guardado los cambios, Desea hacelo ahora", vbYesNo + vbInformation, "Clientes") = 6 Then
                    If Opcruccli.value And left(Trim(txtnuevo.Text), 2) = "00" Then
                        MsgBox "Ruc Incorrecto", vbInformation + vbOKOnly, "Atención"
                        txtnuevo.SetFocus
                        Exit Sub
                    End If
                    If Len(Trim(txtvendedor)) = 0 Then
                       MsgBox "Debe ingresar el vendedor", vbInformation, "Mensaje"
                       txtvendedor.SetFocus
                       Exit Sub
                    Else
                       If traerCampo("EF2USERS", "F2CODUSER", "F2CODUSER", Trim(txtvendedor.Text)) = "" Then
                          MsgBox "Codigo de vendedor no existe", vbInformation, "Mensaje"
                          txtvendedor.SetFocus
                          Exit Sub
                       End If
                    End If
                    Graba_Cliente
                End If
            End If
            Nuevo_Cliente
            sw_nuevo_mant = True
            TXTFECING.SetFocus
            wgraba = 1
            Me.MousePointer = vbDefault
        Case "ID_Grabar":
            If OpcLibcli.value = False And Opcruccli.value = False Then
                MsgBox "Falta indicar el tipo de cliente.", vbInformation, "Atención"
                Exit Sub
            End If
            If Len(Trim(Txtzona2)) > 0 Then
               If traerCampo("EF2ZONAS_CLIENTE", "DISTRITO", "CODIGO", Trim(Txtzona2.Text)) = "" Then
                  MsgBox "Codigo de zona no existe", vbInformation, "Mensaje"
                  Txtzona2.SetFocus
                  Exit Sub
               End If
            End If
            If Len(Trim(Txtzona)) > 0 Then
               If traerCampo("EF2ZONAS", "F2DESZON", "F2CODZON", Trim(Txtzona.Text)) = "" Then
                  MsgBox "Codigo de distrito no existe", vbInformation, "Mensaje"
                  Txtzona.SetFocus
                  Exit Sub
               End If
            End If
            If Len(Trim(txtnuevo.Text)) = 0 Then
                MsgBox "Falta ingresar el Ruc", vbCritical
                Exit Sub
            End If
            
            If (Val(txtdia.Text) > 0 And Val(txtmes.Text) = 0) Then
                MsgBox "Debe Ingresar Mes", vbInformation, "Mensaje"
                txtmes.SetFocus
                Exit Sub
            End If
            
            If (Val(txtdia.Text) = 0 And Val(txtmes.Text) > 0) Then
                MsgBox "Debe Ingresar Día", vbInformation, "Mensaje"
                txtdia.SetFocus
                Exit Sub
            End If
            
            If addCliFac = True Then
                If Len(Trim(Txtdircli.Text)) = 0 Then
                    MsgBox "Debe Ingresar Direccion del Cliente", vbInformation, "Sistema de Ventas"
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Len(Trim(Txttelcli.Text)) = 0 Then
                    MsgBox "Debe Ingresar el Telefono del Cliente", vbInformation, "Sistema de Ventas"
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
                            
            'n = traerCampo("EF2CLIENTES", "F2NEWRUC", "F2NEWRUC", Trim(txtnuevo.Text))
                                            
            If Len(Trim(txtvendedor)) = 0 Then
               MsgBox "Debe ingresar el vendedor", vbInformation, "Mensaje"
               txtvendedor.SetFocus
               Exit Sub
            Else
               If traerCampo("EF2USERS", "F2CODUSER", "F2CODUSER", Trim(txtvendedor.Text)) = "" Then
                  MsgBox "Codigo de vendedor no existe", vbInformation, "Mensaje"
                  txtvendedor.SetFocus
                  Exit Sub
               End If
            End If
            
            If Opcruccli.value And left(Trim(txtnuevo.Text), 2) = "00" Then
                MsgBox "Ruc Incorrecto", vbInformation + vbOKOnly, "Atención"
                txtnuevo.SetFocus
                Exit Sub
            End If
            
            'If traerCampo("EF2CLIENTES", "F2NEWRUC", "F2NEWRUC", Trim(txtnuevo.Text), "") = "" Then 'jcb
                Me.MousePointer = vbHourglass
                Graba_Cliente
                wgraba = 0
            'Else
            '    MsgBox "El Cliente con el RUC " & txtnuevo.Text & " existe, Ingrese otro", vbInformation, "Atención"
            'End If

            
            If sw_load_mant = True Then
                wcodcli = "" & Txtcodcli.Text
                wnomcli = "" & Txtnomcli.Text
                wruccli = "" & txtnuevo.Text
                WDIRCLI = "" & Txtdircli.Text
                wforpag = "" & txtcodpag.Text
                If OpcLibcli.value = True Then wtipocli = "N"
                If Opcruccli.value = True Then wtipocli = "J"
                wdocidentidad = "" & txtdocidentidad.Text
                If opttipo(0).value = True Then wcliext = "N"
                If opttipo(1).value = True Then wcliext = "E"
                Unload Me
            End If
            Me.MousePointer = vbDefault
        Case "ID_Eliminar"
            Me.MousePointer = vbHourglass
            Elimina_Cliente
            Me.MousePointer = vbDefault
        Case "ID_Lista":
            'If sw_nuevo_mant = False Then
'                lista_clientes.adoclientes.Refresh
            'Else
            '    wcodcli = "" & Txtcodcli.Text
            '    wnomcli = "" & txtnomcli.Text
            '    wruccli = "" & txtnuevo.Text
            '    wdircli = "" & txtdircli.Text
            '    wforpag = "" & Txtcodpag.Text
            '    If OpcLibcli.Value = True Then wtipocli = "" = "N"
            '    If Opcruccli.Value = True Then wtipocli = "" = "J"
            '    wdocidentidad = "" & txtdocidentidad.Text
            'End If
            Unload Me
        Case "ID_Salir"
            wcodcli = "" & Txtcodcli.Text
            wnomcli = "" & Txtnomcli.Text
            wruccli = "" & txtnuevo.Text
            WDIRCLI = "" & Txtdircli.Text
            wforpag = "" & txtcodpag.Text
            If OpcLibcli.value = True Then wtipocli = "N"
            If Opcruccli.value = True Then wtipocli = "J"
            wdocidentidad = "" & txtdocidentidad.Text
            If opttipo(0).value = True Then wcliext = "N"
            If opttipo(1).value = True Then wcliext = "E"
            Unload Me
    End Select
   Exit Sub
ERRORCLI:  Resume Next
End Sub

Private Sub SSActiveToolBars2_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Grabar":
            If dxDBGrid1.Dataset.State = dsInsert Or dxDBGrid1.Dataset.State = dsEdit Then
                dxDBGrid1.Dataset.Post
            End If
            GrabarProductos
        Case "ID_GrabarSucursales":
            If dxDBGrid2.Dataset.State = dsInsert Or dxDBGrid2.Dataset.State = dsEdit Then
                dxDBGrid2.Dataset.Post
            End If
            GrabarSucursales
    End Select
End Sub

Private Sub SSActiveToolBars3_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Grabar_Lista_Especifica":
            If dxDBGrid3.Dataset.State = dsInsert Or dxDBGrid3.Dataset.State = dsEdit Then
                dxDBGrid3.Dataset.Post
            End If
            Grabar_Lista_especifica
    End Select
End Sub

Private Sub txtabrev_GotFocus()
    txtabrev.SelStart = 0: txtabrev.SelLength = Len(txtabrev.Text)
End Sub

Private Sub txtabrev_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Txtrepleg.SetFocus
End Sub

Private Sub txtaño_Change()

End Sub

Private Sub txtclicom_Change()

End Sub

Private Sub txtclicom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

'Private Sub Txtcodcli_LostFocus()
''**** Buscar Cliente
''Actualiza_Cliente Txtcodcli
'End Sub

'Private Sub txtcomision_GotFocus()
'
'    txtcomision.SelStart = 0: txtcomision.SelLength = Len(txtcomision.Text)
'
'End Sub

Private Sub txtdia_Change()
If Not (Val(txtdia.Text) >= 1 And Val(txtdia.Text) <= 31 Or Trim(txtdia.Text) = "") Then
    MsgBox "Dia Incorrecto", vbInformation, "Mensaje"
    txtdia.Text = ""
End If
End Sub

Private Sub txtdia_GotFocus()
txtdia.SelStart = 0
txtdia.SelLength = Len(txtdia.Text)
End Sub

Private Sub txtdia_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtmes.SetFocus
Else
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtcanal_DblClick()
    txtcanal_KeyDown 113, 0
End Sub

Private Sub txtcanal_GotFocus()
    txtcanal.SelStart = 0: txtcanal.SelLength = Len(txtcanal.Text)
End Sub

Private Sub txtcanal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcodcanal = "": wnomcanal = ""
        sw_ayuda = True
        hlp_canales.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcanal)) > 0 Then
            txtcanal.Text = wcodcanal
            pnlcanal.Caption = wnomcanal
            txtcanal_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcanal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtcanal_LostFocus()
    If sw_ayuda = False Then
        If Len(Trim(txtcanal.Text)) > 0 Then
            If VALIDA_CANAL(txtcanal.Text) = True Then
                pnlcanal.Caption = wnomcanal
            Else
                pnlcanal.Caption = ""
                MsgBox "Código de canal no existe. Verifique.", vbInformation + vbDefaultButton1, "Atención"
                txtcanal.Text = "": txtcanal.SetFocus
            End If
        Else
            pnlcanal.Caption = ""
        End If
    End If
End Sub

Private Sub Txtcarleg_GotFocus()
    Txtcarleg.SelStart = 0: Txtcarleg.SelLength = Len(Txtcarleg.Text)
End Sub

Private Sub Txtcarleg_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtcobrador_DblClick()
    txtcobrador_KeyDown 113, 0
End Sub

Private Sub txtcobrador_GotFocus()
    txtcobrador.SelStart = 0: txtcobrador.SelLength = Len(txtcobrador.Text)
End Sub

Private Sub txtcobrador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcodusuario = "": wnomusuario = ""
        wtipo_usuario = "C"
        sw_ayuda = True
        ayuda_usuarios.Show 1
        sw_ayuda = False
        If Len(Trim(wcodusuario)) > 0 Then
            txtcobrador.Text = wcodusuario
            pnlcobrador.Caption = wnomusuario
            txtcobrador_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcobrador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtcobrador.Text) > 0 Then
           txtcodpag.SetFocus
        Else
            MsgBox "Ingrese una Cuenta", vbInformation, "Mensaje"
            txtcobrador.SetFocus
            txtcobrador.Text = ""
            pnlcobrador.Caption = ""
        End If
    End If
End Sub

Private Sub txtcobrador_LostFocus()
    If sw_ayuda = False Then
        If Len(Trim(txtcobrador.Text)) > 0 Then
            If VALIDA_COBRADOR(txtcobrador.Text) = True Then
                pnlcobrador.Caption = wnomven
            Else
                pnlcobrador.Caption = ""
                MsgBox "Código de cobrador no existe. Verifique.", vbInformation + vbDefaultButton1, "Atención"
                txtcobrador.Text = "": txtcobrador.SetFocus
            End If
        Else
            pnlcobrador.Caption = ""
        End If
    End If
End Sub

Private Sub Txtcodcli_DblClick()
    'ayuda_clientes.Show 1
    'If Len(wcodcli) > 0 Then
    '    Txtcodcli.Text = wcodcli
    '    Txtcodcli_KeyPress 13
    'End If
End Sub

Private Sub Txtcodcli_GotFocus()
    Txtcodcli.SelStart = 0: Txtcodcli.SelLength = Len(Txtcodcli.Text)
End Sub

Private Sub Txtcodcli_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 113 Then
    '    Txtcodcli_DblClick
    'End If
End Sub

Private Sub Txtcodcli_KeyPress(KeyAscii As Integer)
On Error GoTo ERRORCLI
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Txtcodcli.Text = Trim("" & Txtcodcli.Text)
        sql = "select * from ef2clientes where f2codcli = '" & Txtcodcli.Text & "'"
        If rsclientes.State = adStateOpen Then rsclientes.Close
        rsclientes.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsclientes.EOF Then
            Actualiza_Cliente Txtcodcli.Text
        Else
            Txtnomcli.SetFocus
        End If
        rsclientes.Close
    End If
    Exit Sub
ERRORCLI: Resume Next
End Sub

Private Sub Txtcodedi_GotFocus()
        Txtcodedi.SelStart = 0: Txtcodedi.SelLength = Len(Txtcodedi.Text)
End Sub

Private Sub Txtcodedi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TXTDIRCOB.SetFocus
End Sub

Private Sub txtcodpag_DblClick()
    txtcodpag_KeyDown 113, 0
End Sub

Private Sub txtcodpag_GotFocus()
    txtcodpag.SelStart = 0: txtcodpag.SelLength = Len(txtcodpag.Text)
End Sub

Private Sub txtcodpag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcodpag = "": wnompag = ""
        sw_ayuda = True
        ayuda_formapago.Show 1
        sw_ayuda = False
        If Len(Trim(wcodpag)) > 0 Then
            txtcodpag.Text = Trim$(wcodpag)
            pnldespag.Caption = Trim$(wnompag)
            txtcodpag_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcodpag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
End Sub

Private Sub txtcodpag_LostFocus()
    If sw_ayuda = False Then
        If Trim(txtcodpag.Text) <> "" Then
            If VALIDA_FPAGO(txtcodpag.Text) = True Then
                pnldespag.Caption = wnompag
            Else
                pnldespag.Caption = ""
                MsgBox "Código de forma de pago no existe.", vbInformation + vbDefaultButton1, "Atención"
                txtcodpag.Text = "": txtcodpag.SetFocus
            End If
        Else
            pnldespag.Caption = ""
        End If
    End If
End Sub

Private Sub txtcodpagxcobrar_DblClick()
    txtcodpagxcobrar_KeyDown 113, 0
End Sub

Private Sub txtcodpagxcobrar_GotFocus()
    txtcodpagxcobrar.SelStart = 0: txtcodpagxcobrar.SelLength = Len(txtcodpagxcobrar.Text)
End Sub

Private Sub txtcodpagxcobrar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcodpag = "": wnompag = ""
        sw_ayuda = True
        ayuda_formapago.Show 1
        sw_ayuda = False
        If Len(Trim(wcodpag)) > 0 Then
            txtcodpagxcobrar.Text = Trim$(wcodpag)
            pnldespagxcobrar.Caption = Trim$(wnompag)
            txtcodpagxcobrar_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcodpagxcobrar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtabrev.SetFocus
    End If
End Sub

Private Sub txtcodpagxcobrar_LostFocus()
    If sw_ayuda = False Then
        If Trim(txtcodpagxcobrar.Text) <> "" Then
            If VALIDA_FPAGO(txtcodpagxcobrar.Text) = True Then
                pnldespagxcobrar.Caption = wnompag
            Else
                pnldespagxcobrar.Caption = ""
                MsgBox "Código de forma de pago no existe.", vbInformation + vbDefaultButton1, "Atención"
                txtcodpagxcobrar.Text = "": txtcodpagxcobrar.SetFocus
            End If
        Else
            pnldespagxcobrar.Caption = ""
        End If
    End If
End Sub

Private Sub txtcontacto_GotFocus()
    txtcontacto.SelStart = 0: txtcontacto.SelLength = Len(txtcontacto.Text)
End Sub

Private Sub Txtcredis_GotFocus()
    Txtcredis.SelStart = 0: Txtcredis.SelLength = Len(Txtcredis.Text)
End Sub

Private Sub txtdia_LostFocus()
txtdia.Text = Format(txtdia.Text, "00")
End Sub

Private Sub Txtdircli_GotFocus()
'    txtdircli.Text = cmbtipvia.Text & " " & txtnomvia.Text & " " & txtnumvia.Text & " " & txtint
    Txtdircli.SelStart = 0: Txtdircli.SelLength = Len(Txtdircli.Text)
End Sub

Private Sub TxtDirCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtzona.SetFocus
    End If
End Sub

Private Sub TXTDIRCOB_GotFocus()
    TXTDIRCOB.SelStart = 0: TXTDIRCOB.SelLength = Len(TXTDIRCOB.Text)
End Sub

Private Sub TXTDIRCOB_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtvendedor.SetFocus
End Sub

Private Sub txtdocidentidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If OpcLibcli.Value Then
'            txtdia.SetFocus
'        Else
             ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
'        End If
    End If
End Sub

Private Sub txtemail_GotFocus()
    If Len(Trim(txtemail.Text)) = 0 Then
        txtemail.Text = "@"
    End If
    txtemail.SelStart = 0: txtemail.SelLength = Len(txtemail.Text)
End Sub

Private Sub TxtDIRcli_KeyUp(KeyCode As Integer, Shift As Integer)
    'Txtnomcli_KeyUp KeyCode, 0
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txtweb.SetFocus
    End If
End Sub

Private Sub txtemailcoti_GotFocus()
    If Len(Trim(txtemailcoti.Text)) = 0 Then
        txtemailcoti.Text = "@"
    End If
    txtemailcoti.SelStart = 0: txtemailcoti.SelLength = Len(txtemailcoti.Text)
End Sub

Private Sub txtemailcoti_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtescuela_DblClick()
    txtescuela_KeyDown 113, 0
End Sub

Private Sub txtescuela_GotFocus()
    txtescuela.SelStart = 0: txtescuela.SelLength = Len(txtescuela.Text)
End Sub

Private Sub txtescuela_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcodcli = "": wnomcli = "": wruccli = "": WDIRCLI = ""
        sw_ayuda = True
        sw_escuela = True
        ayuda_clientes.Show 1
        sw_ayuda = False
        sw_escuela = False
        If Len(Trim(wcodcli)) > 0 Then
            txtescuela.Text = Trim(wcodcli)
            pnlescuela.Caption = Trim(wnomcli)
        End If
    End If
End Sub

Private Sub Txtfaxcli_GotFocus()
    Txtfaxcli.SelStart = 0: Txtfaxcli.SelLength = Len(Txtfaxcli.Text)
End Sub

Private Sub Txtfaxcli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtemail.SetFocus
End Sub

Private Sub Txtfecing_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Opcruccli.SetFocus
    
End Sub

Private Sub txtfecnac_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmbsexo.SetFocus
End If
End Sub

Private Sub txtint_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub Txtlincre_GotFocus()
    Txtlincre.SelStart = 0
    Txtlincre.SelLength = Len(Txtlincre.Text)
End Sub

Private Sub Txtlincre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Txtcredis.SetFocus
End Sub

Private Sub Txtlincre_LostFocus()
    Txtlincre.Text = Format(Txtlincre.Text, "###,###,##0.00")
    Txtcredis.Text = Format(Val(Format(Txtlincre.Text, "0.00")) - Val(Format(txtxcobrar.Text, "0.00")), "###,###,##0.00")
End Sub

Private Sub txtmes_Change()
If Not (Val(txtmes.Text) >= 1 And Val(txtmes.Text) <= 12 Or Trim(txtmes.Text) = "") Then
    MsgBox "Mes Incorrecto", vbInformation, "Mensaje"
    txtmes.Text = ""
End If
End Sub

Private Sub txtmes_GotFocus()
txtmes.SelStart = 0
txtmes.SelLength = Len(txtmes.Text)
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmbsexo.SetFocus
Else
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtmes_LostFocus()
txtmes.Text = Format(txtmes.Text, "00")
End Sub


Private Sub Txtnomcli_GotFocus()
    Txtnomcli.SelStart = 0: Txtnomcli.SelLength = Len(Txtnomcli.Text)
End Sub

Private Sub Txtnomcli_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtnomvia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtnuevo_GotFocus()
    
    txtnuevo.SelStart = 0: txtnuevo.SelLength = Len(txtnuevo.Text)
    
End Sub

Private Sub txtnuevo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtnomcli.SetFocus
    
End Sub

Private Sub txtnuevo_LostFocus()
    If Len(Trim(txtnuevo.Text)) > 0 And Len(Trim(txtnuevo.Text)) <> 11 Then
        MsgBox "Falta ingresar el Ruc Correctamente", vbInformation + vbOKOnly, "Atención"
        txtnuevo.SetFocus
        Exit Sub
    End If
    
'    If Opcruccli.Value And left(Trim(txtnuevo.Text), 2) = "00" Then
'        MsgBox "Ruc Incorrecto", vbInformation + vbOKOnly, "Atención"
'        txtnuevo.SetFocus
'        Exit Sub
'    End If
    
    If traerCampo("EF2CLIENTES", "F2NEWRUC", "F2NEWRUC", Trim(txtnuevo.Text), "") = "" Then  'jcb
    Else
        MsgBox "El Cliente con el RUC " & txtnuevo.Text & " existe, Ingrese otro", vbInformation, "Atención"
        txtnuevo.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtnumvia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtpais_GotFocus()
    txtpais.SelStart = 0: txtpais.SelLength = Len(txtpais.Text)
End Sub

Private Sub txtpais_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
 
End Sub

Private Sub txtrecepcio_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub Txtrepleg_GotFocus()
    
    Txtrepleg.SelStart = 0: Txtrepleg.SelLength = Len(Txtrepleg.Text)
    
End Sub

Private Sub Txtrepleg_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtcarleg.SetFocus
    
End Sub

Private Sub Txttelcli_GotFocus()
    
    Txttelcli.SelStart = 0: Txttelcli.SelLength = Len(Txttelcli.Text)
    
End Sub

Private Sub Txttelcli_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtfaxcli.SetFocus
    
End Sub

Private Sub Txttercon_GotFocus()
    
    Txttercon.SelStart = 0: Txttercon.SelLength = Len(Txttercon.Text)
    
End Sub

Private Sub Txttercon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    
End Sub

Private Sub txtvendedor_DblClick()

    txtVendedor_KeyDown 113, 0
    
End Sub

Private Sub txtvendedor_GotFocus()

    txtvendedor.SelStart = 0: txtvendedor.SelLength = Len(txtvendedor.Text)
    
End Sub

Private Sub txtVendedor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodusuario = "": wnomusuario = ""
        wtipo_usuario = "V"
        ayuda_usuarios.Show 1
        sw_ayuda = False
        If Len(Trim(wcodusuario)) > 0 Then
            txtvendedor.Text = wcodusuario
            pnlvendedor.Caption = wnomusuario
            txtvendedor_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtvendedor_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
        If Len(txtvendedor.Text) > 0 Then
           'txtcobrador.SetFocus
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        Else
            MsgBox "Ingrese el Codigo del Vendedor", vbInformation, "Mensaje"
            txtvendedor.SetFocus
            txtvendedor.Text = ""
            pnlvendedor.Caption = ""
        End If
    End If

End Sub

Private Sub txtvendedor_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtvendedor.Text)) > 0 Then
            If VALIDA_VENDEDOR(txtvendedor.Text) = True Then
                pnlvendedor.Caption = wnomven
            Else
                pnlvendedor.Caption = ""
                MsgBox "Código de vendedor no existe. Verifique.", vbInformation + vbDefaultButton1, "Atención"
                txtvendedor.Text = "": txtvendedor.SetFocus
            End If
        Else
            pnlvendedor.Caption = ""
        End If
    End If

End Sub

Private Sub txtweb_GotFocus()

    Txtweb.SelStart = 0: Txtweb.SelLength = Len(Txtweb.Text)
    
End Sub

Private Sub txtweb_keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then txtcontacto.SetFocus

End Sub

Private Sub Txtzona_DblClick()
    
    Txtzona_KeyUp 113, 0
    
End Sub

Private Sub Txtzona_GotFocus()

    Txtzona.SelStart = 0: Txtzona.SelLength = Len(Txtzona.Text)
    
End Sub

Private Sub Txtzona_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(Trim(Txtzona.Text)) > 0 Then
            'SQL = "SELECT DISTRITO FROM EF2ZONAS_CLIENTE WHERE CODIGO= '" & Txtzona.Text & "'"
            sql = "SELECT F2DESZON FROM EF2ZONAS WHERE F2CODZON= '" & Txtzona.Text & "'"
            If rst.State = adStateOpen Then rst.Close
            rst.Open sql, cnn_dbbancos, adOpenDynamic
            If Not rst.EOF Then
                Txtdeszon.Caption = "" & rst.Fields("f2deszon")
                 ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
            Else
                Txtzona.Text = ""
                Txtzona.SetFocus
            End If
            rst.Close
'            txtdocidentidad.SetFocus
        Else
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
            Txtzona.Text = ""
            Txtdeszon.Caption = ""
        End If
    End If

End Sub

Private Sub Txtzona_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        ayuda_zonas.Show 1
        If Len(wcodzona) > 0 Then
            Txtzona.Text = wcodzona
            Txtzona_KeyPress 13
        End If
    End If
    
End Sub

Private Sub EVALUA_DATOS_FINANCIEROS(pcliente As String, ptc As Double)
Dim rsctadcto       As New ADODB.Recordset
Dim rsdocumentos    As New ADODB.Recordset
Dim nsoles          As Double
Dim ndolar          As Double
Dim nutilizado      As Double
Dim csql            As String
Dim nventas         As Double

    If rsctadcto.State = adStateOpen Then rsctadcto.Close
    rsctadcto.Open "SELECT SUM(SALDO) AS NSOLES FROM CTA_DCTO WHERE CLIENTE='" & pcliente & "' AND MONEDA='S'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsctadcto.EOF Then
        nsoles = Val("" & rsctadcto.Fields("NSOLES"))
    End If
    rsctadcto.Close
    
    If rsctadcto.State = adStateOpen Then rsctadcto.Close
    rsctadcto.Open "SELECT SUM(SALDO) AS NDOLARES FROM CTA_DCTO WHERE CLIENTE='" & pcliente & "' AND MONEDA='D'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsctadcto.EOF Then
        ndolar = Val("" & rsctadcto.Fields("NDOLARES"))
    End If
    rsctadcto.Close
    
    nutilizado = 0#
    If nsoles > 0# Then
        nutilizado = Val(Format(nsoles / ptc, "0.00"))
    End If
    nutilizado = nutilizado + ndolar
    
    txtxcobrar.Text = Format(nutilizado, "###,###,##0.00")
    Txtcredis.Text = Format(Val(Format(Txtlincre.Text, "0.00")) - nutilizado, "###,###,##0.00")
    
    '-----------------------------------------------------------------------------------------------
    '-------------------------- ACUMULA LAS VENTAS
    csql = "SELECT SUM(F4TOTFAC) AS NSOLES FROM TBVENTA_CAB WHERE F4TIPMON='S' AND (F4TIPODOCU='01' OR F4TIPODOCU='03' OR F4TIPODOCU='08') AND F2CODCLI='" & pcliente & "'"
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        nsoles = Val("" & rsdocumentos.Fields("NSOLES"))
    End If
    rsdocumentos.Close
    
    csql = "SELECT SUM(F4TOTFAC) AS NDOLARES FROM TBVENTA_CAB WHERE F4TIPMON='D' AND (F4TIPODOCU='01' OR F4TIPODOCU='03' OR F4TIPODOCU='08') AND F2CODCLI='" & pcliente & "'"
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        ndolar = Val("" & rsdocumentos.Fields("NDOLARES"))
    End If
    rsdocumentos.Close
    
    nventas = 0#
    If nsoles > 0# Then
        nventas = Val(Format(nsoles / ptc, "0.00"))
    End If
    nventas = nventas + ndolar
    
    '-----------------------------------------------------------------------------------------------
    '-------------------------- ACUMULA LAS NOTAS DE CREDITO
    csql = "SELECT SUM(F4TOTFAC) AS NSOLES FROM TBVENTA_CAB WHERE F4TIPMON='S' AND F4TIPODOCU='07' AND F2CODCLI='" & pcliente & "'"
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        nsoles = Val("" & rsdocumentos.Fields("NSOLES"))
    End If
    rsdocumentos.Close
    
    csql = "SELECT SUM(F4TOTFAC) AS NDOLARES FROM TBVENTA_CAB WHERE F4TIPMON='D' AND F4TIPODOCU='08' AND F2CODCLI='" & pcliente & "'"
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        ndolar = Val("" & rsdocumentos.Fields("NDOLARES"))
    End If
    rsdocumentos.Close
    
    '------- RESTA LAS N/C
    If nsoles > 0# Then
        nventas = nventas - Val(Format(nsoles / ptc, "0.00"))
    End If
    nventas = nventas - ndolar
    
    txtventasacum.Text = Format(nventas, "###,###,##0.00")
        
End Sub

Private Sub Conf_Grid2()

    With dxDBGrid2.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
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
        '.Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
End Sub

Private Sub AdicionaItem()

    DELETEREC_LOG dbtablecli, cnn_formcli
    DELETEREC_LOG dbtablecli, cnn_formcli
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Refresh
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
   
    With dxDBGrid1.Dataset
        sw_nuevo_item = True
        For i = 1 To 1
            .Append
            .FieldValues("ITEM") = i
            .FieldValues("CODIGO") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UNIDAD") = ""
            .FieldValues("AFECTO") = "*"
            .FieldValues("PRECIOUNIT") = Format(0, "0.00")
            .FieldValues("CODFAB") = ""
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

End Sub

Private Sub GrabarProductos()
Dim nnumitems       As Integer
Dim cvalores1       As String
Dim ccodsucursal    As String

    ReDim amovs_det(0 To 9) As a_grabacion
    If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
    TbDocumDet.Open "Select COUNT(CODIGO) AS NITEM from " & dbtablecli & " WHERE LEN(TRIM(CODIGO))<>0 OR NOT ISNULL(CODIGO)", cnn_formcli, adOpenDynamic, adLockOptimistic
    If Not TbDocumDet.EOF Then
        nnumitems = Val("" & TbDocumDet.Fields("NITEM"))
    Else
        nnumitems = 0
    End If
    TbDocumDet.Close
    
    ccodsucursal = Trim("" & dxDBGrid2.Columns.ColumnByFieldName("F2CODSUCURSAL").value)
    If nnumitems > 0 Then
        amovs_det(0).campo = "F2CODCLI": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
        amovs_det(1).campo = "F2CODSUCURSAL": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
        amovs_det(2).campo = "F2NOMCLI": amovs_det(2).valor = "": amovs_det(2).Tipo = "T"
        amovs_det(3).campo = "F5CODPRO": amovs_det(3).valor = "": amovs_det(3).Tipo = "T"
        amovs_det(4).campo = "F5NOMPRO": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
        amovs_det(5).campo = "F7CODMED": amovs_det(5).valor = "": amovs_det(5).Tipo = "T"
        amovs_det(6).campo = "F5AFECTO": amovs_det(6).valor = "": amovs_det(6).Tipo = "T"
        amovs_det(7).campo = "F5PREUNI": amovs_det(7).valor = "": amovs_det(7).Tipo = "N"
        amovs_det(8).campo = "F2NEWRUC": amovs_det(8).valor = "": amovs_det(8).Tipo = "T"
        amovs_det(9).campo = "F5CODFAB": amovs_det(9).valor = "": amovs_det(9).Tipo = "T"
        cvalores1 = "1111111111"
        i = 0
        If dxDBGrid1.Dataset.RecordCount > 0 Then
            If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
            TbDocumDet.Open "Select * from " & dbtablecli & "", cnn_formcli, adOpenDynamic, adLockOptimistic
            ReDim ValuesDet(9, dxDBGrid1.Dataset.RecordCount - 1)
            If Not TbDocumDet.EOF Then
                Do While Not TbDocumDet.EOF
                    If Len(Trim("" & TbDocumDet.Fields("codigo"))) > 0 Then
                        ValuesDet(0, i) = "" & Txtcodcli.Text
                        ValuesDet(1, i) = ccodsucursal
                        ValuesDet(2, i) = "" & Txtnomcli.Text
                        ValuesDet(3, i) = Trim("" & TbDocumDet.Fields("codigo"))
                        ValuesDet(4, i) = Trim("" & TbDocumDet.Fields("descripcion"))
                        ValuesDet(5, i) = Trim("" & TbDocumDet.Fields("unidad"))
                        ValuesDet(6, i) = Trim("" & TbDocumDet.Fields("afecto"))
                        ValuesDet(7, i) = Format(Val("" & TbDocumDet.Fields("preciounit")), "0.00")
                        ValuesDet(8, i) = "" & txtnuevo.Text
                        ValuesDet(9, i) = Trim("" & TbDocumDet.Fields("CODFAB"))
                        i = i + 1
                    End If
                    TbDocumDet.MoveNext
                Loop
            End If
            TbDocumDet.Close
            
            csql = ("DELETE FROM IF4PRODCLI WHERE F2CODCLI = '" & Txtcodcli.Text & "' AND F2CODSUCURSAL = '" & ccodsucursal & "'")
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos
            
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF4PRODCLI", "A", 9, cnn_dbbancos, "", ValuesDet(), nnumitems - 1, cvalores1, "", ""
        End If
    Else
    
        
        csql = ("DELETE FROM IF4PRODCLI WHERE F2CODCLI = '" & Txtcodcli.Text & "' AND F2CODSUCURSAL = '" & ccodsucursal & "'")
        cnn_dbbancos.Execute csql
        'AlmacenaQuery_sql csql, cnn_dbbancos
    
    
    End If
    
End Sub

Private Sub Actualiza_Productos(pcodcli As String, psucursal As String)
Dim RSCONSULTA      As New ADODB.Recordset
    
    sw_nuevo_item = True
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid1.Dataset.Active = True

    DELETEREC_LOG dbtablecli, cnn_formcli
    DELETEREC_LOG dbtablecli, cnn_formcli
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.Dataset.Refresh
    
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls
    sw_nuevo_doc = True
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open "Select * from IF4PRODCLI where F2CODCLI='" & pcodcli & "' AND F2CODSUCURSAL='" & psucursal & "' ORDER BY F5NOMPRO", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    i = 1
    If Not RSCONSULTA.EOF Then
        With dxDBGrid1.Dataset
            Do While Not RSCONSULTA.EOF
                .Append
                .FieldValues("ITEM") = i
                .FieldValues("CODIGO") = "" & RSCONSULTA.Fields("F5CODPRO")
                .FieldValues("DESCRIPCION") = "" & RSCONSULTA.Fields("F5NOMPRO")
                .FieldValues("UNIDAD") = "" & RSCONSULTA.Fields("F7CODMED")
                .FieldValues("AFECTO") = "" & RSCONSULTA.Fields("F5AFECTO")
                .FieldValues("PRECIOUNIT") = Format(Val("" & RSCONSULTA.Fields("F5PREUNI")), "0.00")
                .FieldValues("CODFAB") = "" & RSCONSULTA.Fields("F5CODFAB")
                RSCONSULTA.MoveNext
                i = i + 1
            Loop
            .Post
        End With
    Else
        AdicionaItem
    End If
    RSCONSULTA.Close
    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    sw_nuevo_item = False

End Sub

Private Sub AdicionaItem2()
Dim i As Integer

    DELETEREC_LOG ctablesuc, cnn_formcli
    DELETEREC_LOG ctablesuc, cnn_formcli
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Refresh
    
    dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid2.Dataset.Active = True
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open
    
    With dxDBGrid2.Dataset
        sw_nuevo_item = True
        For i = 1 To 1
            .Append
            .FieldValues("ITEM") = i
            .FieldValues("F2CODCLI") = ""
            .FieldValues("F2CODSUCURSAL") = ""
            .FieldValues("F2DIRCLI") = ""
            .FieldValues("F2CONTACTO") = ""
            .FieldValues("F2CODZONA") = ""
            .FieldValues("F2CODVENDEDOR") = ""
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open

End Sub

Private Sub Conf_Grid()

    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
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
        '.Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
    '---------------------------------------------------------------
    If RSCONSULTA2.State = adStateOpen Then RSCONSULTA2.Close
    RSCONSULTA2.Open "Select * from EF2SUCURSALES where F2CODCLI='" & Txtcodcli.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If RSCONSULTA2.EOF Then
        AdicionaItem2
    Else
        ACTUALIZA_SUCURSALES Txtcodcli.Text
    End If
    RSCONSULTA2.Close
    '---------------------------------------------------------------
    
End Sub

Private Sub ACTUALIZA_SUCURSALES(pcodcli As String)
Dim RSCONSULTA As New ADODB.Recordset
    
    sw_nuevo_item = True
    dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid2.Dataset.Active = True

    DELETEREC_LOG ctablesuc, cnn_formcli
    DELETEREC_LOG ctablesuc, cnn_formcli
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open
    dxDBGrid2.Dataset.Refresh
    
    dxDBGrid2.OptionEnabled = False
    dxDBGrid2.Dataset.DisableControls
    sw_nuevo_doc = True
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open "Select * from EF2SUCURSALES where F2CODCLI='" & pcodcli & "' ORDER BY F2CODSUCURSAL", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    i = 1
    If Not RSCONSULTA.EOF Then
        With dxDBGrid2.Dataset
            Do While Not RSCONSULTA.EOF
                .Append
                .FieldValues("ITEM") = i
                .FieldValues("F2CODCLI") = "" & RSCONSULTA.Fields("F2CODCLI")
                .FieldValues("F2CODSUCURSAL") = "" & RSCONSULTA.Fields("F2CODSUCURSAL")
                .FieldValues("F2DIRCLI") = "" & RSCONSULTA.Fields("F2DIRCLI")
                .FieldValues("F2CONTACTO") = "" & RSCONSULTA.Fields("F2CONTACTO")
                .FieldValues("F2CODZONA") = "" & RSCONSULTA.Fields("F2CODZONA")
                .FieldValues("F2CODVENDEDOR") = "" & RSCONSULTA.Fields("F2CODVENDEDOR")
                RSCONSULTA.MoveNext
                i = i + 1
            Loop
            .Post
        End With
        sw_nuevo_item = False
    Else
        AdicionaItem2
    End If
    RSCONSULTA.Close
    dxDBGrid2.Dataset.EnableControls
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open

End Sub

Private Sub Conf_Grid3()
    
    With dxDBGrid3.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
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
        '.Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With

    
    If RSCONSULTA2.State = adStateOpen Then RSCONSULTA2.Close
    RSCONSULTA2.Open "Select * from IF4PRODCLI where F2CODCLI='" & Txtcodcli & "' AND (ISNULL(F2CODSUCURSAL) OR LEN(TRIM(F2CODSUCURSAL))=0)", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If RSCONSULTA2.EOF Then
        AdicionaItem3
    Else
        Actualiza_Lista_especifica Txtcodcli.Text
    End If
    RSCONSULTA2.Close

End Sub

Private Sub AdicionaItem3()

    DELETEREC_LOG ctablelista, cnn_formcli
    DELETEREC_LOG ctablelista, cnn_formcli
    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Refresh
    
    dxDBGrid3.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid3.Dataset.Active = True

    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Open
   
    With dxDBGrid3.Dataset
        sw_nuevo_item = True
        For i = 1 To 1
            .Append
            .FieldValues("ITEM") = i
            .FieldValues("CODIGO") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UNIDAD") = ""
            .FieldValues("AFECTO") = "*"
            .FieldValues("PRECIOUNIT") = Format(0, "0.00")
            .FieldValues("CODFAB") = ""
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Open

End Sub

Private Sub GrabarSucursales()
Dim nnumitems       As Integer
Dim cvalores1       As String

    ReDim amovs_det(0 To 5) As a_grabacion
    If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
    TbDocumDet.Open "Select COUNT(F2CODSUCURSAL) AS NITEM from " & ctablesuc & " WHERE LEN(TRIM(F2CODSUCURSAL))<>0 OR NOT ISNULL(F2CODSUCURSAL)", cnn_formcli, adOpenDynamic, adLockOptimistic
    If Not TbDocumDet.EOF Then
        nnumitems = Val("" & TbDocumDet.Fields("NITEM"))
    Else
        nnumitems = 0
    End If
    TbDocumDet.Close
    
    If nnumitems > 0 Then
        amovs_det(0).campo = "F2CODCLI": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
        amovs_det(1).campo = "F2CODSUCURSAL": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
        amovs_det(2).campo = "F2DIRCLI": amovs_det(2).valor = "": amovs_det(2).Tipo = "T"
        amovs_det(3).campo = "F2CONTACTO": amovs_det(3).valor = "": amovs_det(3).Tipo = "T"
        amovs_det(4).campo = "F2CODZONA": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
        amovs_det(5).campo = "F2CODVENDEDOR": amovs_det(5).valor = "": amovs_det(5).Tipo = "T"
        cvalores1 = "111111"
    
        i = 0
        If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
        TbDocumDet.Open "Select * from " & ctablesuc & "", cnn_formcli, adOpenDynamic, adLockOptimistic
        ReDim ValuesDet(5, dxDBGrid2.Dataset.RecordCount - 1)
        If Not TbDocumDet.EOF Then
            Do While Not TbDocumDet.EOF
                If Len(Trim("" & TbDocumDet.Fields("F2CODSUCURSAL"))) > 0 Then
                    ValuesDet(0, i) = "" & Txtcodcli.Text
                    ValuesDet(1, i) = Trim("" & TbDocumDet.Fields("F2CODSUCURSAL"))
                    ValuesDet(2, i) = Trim("" & TbDocumDet.Fields("F2DIRCLI"))
                    ValuesDet(3, i) = Trim("" & TbDocumDet.Fields("F2CONTACTO"))
                    ValuesDet(4, i) = Trim("" & TbDocumDet.Fields("F2CODZONA"))
                    ValuesDet(5, i) = Trim("" & TbDocumDet.Fields("F2CODVENDEDOR"))
                    i = i + 1
                End If
                TbDocumDet.MoveNext
            Loop
        End If
        TbDocumDet.Close
        
        
        csql = ("DELETE * FROM EF2SUCURSALES WHERE F2CODCLI = '" & Txtcodcli.Text & "'")
        cnn_dbbancos.Execute csql
        'AlmacenaQuery_sql csql, cnn_dbbancos

    
      
        GRABA_REGISTRO_logistica_DET amovs_det(), "EF2SUCURSALES", "A", 5, cnn_dbbancos, "", ValuesDet(), dxDBGrid2.Dataset.RecordCount - 1, cvalores1, "", ""
    
    End If

End Sub

Private Function OBTIENE_CODSUCURSAL()
Dim rsobtiene_codigo    As New ADODB.Recordset
Dim ccodigo             As String
    
    If rsobtiene_codigo.State = adStateOpen Then rsobtiene_codigo.Close
    rsobtiene_codigo.Open "SELECT F2CODSUCURSAL FROM EF2SUCURSALES ORDER BY F2CODSUCURSAL DESC", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsobtiene_codigo.EOF Then
        ccodigo = Format(Val(rsobtiene_codigo.Fields("F2CODSUCURSAL") & "") + 1, "0000")
    Else
        ccodigo = "0001"
    End If
    rsobtiene_codigo.Close
    OBTIENE_CODSUCURSAL = ccodigo

End Function

Private Sub Actualiza_Lista_especifica(pcodcli As String)
Dim RSCONSULTA      As New ADODB.Recordset
    
    sw_nuevo_item = True
    dxDBGrid3.Dataset.ADODataset.ConnectionString = cnn_formcli
    dxDBGrid3.Dataset.Active = True

    DELETEREC_LOG ctablelista, cnn_formcli
    DELETEREC_LOG ctablelista, cnn_formcli
    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Open
    dxDBGrid3.Dataset.Refresh
    
    dxDBGrid3.OptionEnabled = False
    dxDBGrid3.Dataset.DisableControls
    sw_nuevo_doc = True
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open "Select * from IF4PRODCLI where F2CODCLI='" & pcodcli & "' AND (ISNULL(F2CODSUCURSAL) OR LEN(TRIM(F2CODSUCURSAL))=0) ORDER BY F5NOMPRO", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    i = 1
    If Not RSCONSULTA.EOF Then
        With dxDBGrid3.Dataset
            Do While Not RSCONSULTA.EOF
                .Append
                .FieldValues("ITEM") = i
                .FieldValues("CODIGO") = "" & RSCONSULTA.Fields("F5CODPRO")
                .FieldValues("DESCRIPCION") = "" & RSCONSULTA.Fields("F5NOMPRO")
                .FieldValues("UNIDAD") = "" & RSCONSULTA.Fields("F7CODMED")
                .FieldValues("AFECTO") = "" & RSCONSULTA.Fields("F5AFECTO")
                .FieldValues("PRECIOUNIT") = Format(Val("" & RSCONSULTA.Fields("F5PREUNI")), "0.00")
                .FieldValues("CODFAB") = "" & RSCONSULTA.Fields("F5CODFAB")
                RSCONSULTA.MoveNext
                i = i + 1
            Loop
            .Post
        End With
    Else
        AdicionaItem3
    End If
    RSCONSULTA.Close
    dxDBGrid3.Dataset.EnableControls
    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Open
    sw_nuevo_item = False
    
End Sub

Private Sub Grabar_Lista_especifica()
Dim nnumitems       As Integer
Dim cvalores1       As String
Dim ccodsucursal    As String

    ReDim amovs_det(0 To 9) As a_grabacion
    If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
    TbDocumDet.Open "Select COUNT(CODIGO) AS NITEM from " & ctablelista & " WHERE LEN(TRIM(CODIGO))<>0 OR NOT ISNULL(CODIGO)", cnn_formcli, adOpenDynamic, adLockOptimistic
    If Not TbDocumDet.EOF Then
        nnumitems = Val("" & TbDocumDet.Fields("NITEM"))
    Else
        nnumitems = 0
    End If
    TbDocumDet.Close
    
    If nnumitems > 0 Then
        amovs_det(0).campo = "F2CODCLI": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
        amovs_det(1).campo = "F2CODSUCURSAL": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
        amovs_det(2).campo = "F2NOMCLI": amovs_det(2).valor = "": amovs_det(2).Tipo = "T"
        amovs_det(3).campo = "F5CODPRO": amovs_det(3).valor = "": amovs_det(3).Tipo = "T"
        amovs_det(4).campo = "F5NOMPRO": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
        amovs_det(5).campo = "F7CODMED": amovs_det(5).valor = "": amovs_det(5).Tipo = "T"
        amovs_det(6).campo = "F5AFECTO": amovs_det(6).valor = "": amovs_det(6).Tipo = "T"
        amovs_det(7).campo = "F5PREUNI": amovs_det(7).valor = "": amovs_det(7).Tipo = "N"
        amovs_det(8).campo = "F2NEWRUC": amovs_det(8).valor = "": amovs_det(8).Tipo = "T"
        amovs_det(9).campo = "F5CODFAB": amovs_det(9).valor = "": amovs_det(9).Tipo = "T"
        cvalores1 = "1111111111"
    
        i = 0
        If dxDBGrid3.Dataset.RecordCount > 0 Then
            If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
            TbDocumDet.Open "Select * from " & ctablelista & "", cnn_formcli, adOpenDynamic, adLockOptimistic
            ReDim ValuesDet(9, dxDBGrid3.Dataset.RecordCount - 1)
            If Not TbDocumDet.EOF Then
                Do While Not TbDocumDet.EOF
                    If Len(Trim("" & TbDocumDet.Fields("codigo"))) > 0 Then
                        ValuesDet(0, i) = "" & Txtcodcli.Text
                        ValuesDet(1, i) = ""
                        ValuesDet(2, i) = "" & Txtnomcli.Text
                        ValuesDet(3, i) = Trim("" & TbDocumDet.Fields("codigo"))
                        ValuesDet(4, i) = Trim("" & TbDocumDet.Fields("descripcion"))
                        ValuesDet(5, i) = Trim("" & TbDocumDet.Fields("unidad"))
                        ValuesDet(6, i) = Trim("" & TbDocumDet.Fields("afecto"))
                        ValuesDet(7, i) = Format(Val("" & TbDocumDet.Fields("preciounit")), "0.00")
                        ValuesDet(8, i) = "" & txtnuevo.Text
                        ValuesDet(9, i) = Trim("" & TbDocumDet.Fields("CODFAB"))
                        i = i + 1
                    End If
                    TbDocumDet.MoveNext
                Loop
            End If
            TbDocumDet.Close
            
            
            csql = ("DELETE FROM IF4PRODCLI WHERE F2CODCLI = '" & Txtcodcli.Text & "' AND (ISNULL(F2CODSUCURSAL) OR LEN(TRIM(F2CODSUCURSAL))=0)")
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos

            
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF4PRODCLI", "A", 9, cnn_dbbancos, "", ValuesDet(), nnumitems - 1, cvalores1, "", ""
        End If
    Else
    
    
        csql = ("DELETE FROM IF4PRODCLI WHERE F2CODCLI = '" & Txtcodcli.Text & "' AND (ISNULL(F2CODSUCURSAL) OR LEN(TRIM(F2CODSUCURSAL))=0)")
        cnn_dbbancos.Execute csql
        'AlmacenaQuery_sql csql, cnn_dbbancos

        
    
    End If

End Sub





Private Sub Txtzona2_DblClick()
    
    Txtzona2_KeyUp 113, 0
    
End Sub

Private Sub Txtzona2_GotFocus()

    Txtzona2.SelStart = 0: Txtzona2.SelLength = Len(Txtzona2.Text)
    
End Sub

Private Sub Txtzona2_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(Trim(Txtzona2.Text)) > 0 Then
            sql = "SELECT DISTRITO FROM EF2ZONAS_CLIENTE WHERE CODIGO= '" & Txtzona2.Text & "'"
            If rst.State = adStateOpen Then rst.Close
            rst.Open sql, cnn_dbbancos, adOpenDynamic
            If Not rst.EOF Then
                txtdeszon2.Caption = "" & rst.Fields("DISTRITO")
                 ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
            Else
                Txtzona2.Text = ""
                Txtzona2.SetFocus
            End If
            rst.Close
'            txtdocidentidad.SetFocus
        Else
            Txtzona2.Text = ""
            txtdeszon2.Caption = ""
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"

        End If
    End If

End Sub

Private Sub Txtzona2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        ayuda_zonas.Show 1
        If Len(cod_zon) > 0 Then
            Txtzona2.Text = cod_zon
            Txtzona2_KeyPress 13
        End If
    End If
    
End Sub

