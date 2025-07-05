VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Registro_Compras 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7860
   ClientLeft      =   4215
   ClientTop       =   1695
   ClientWidth     =   11115
   Icon            =   "registro_compras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   1290
      Left            =   2175
      TabIndex        =   62
      Top             =   405
      Width           =   9090
      _Version        =   65536
      _ExtentX        =   16034
      _ExtentY        =   2275
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      Begin VB.TextBox txtmontoautorizado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   92
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txtordcompra 
         Height          =   285
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox TxtCodPrv 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtocompra 
         Height          =   285
         Left            =   6900
         TabIndex        =   28
         Top             =   900
         Width           =   1920
      End
      Begin VB.TextBox txtcodcta 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "PRO"
         Top             =   900
         Width           =   585
      End
      Begin VB.TextBox TxtTelPrv 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2580
         TabIndex        =   1
         Top             =   900
         Width           =   1155
      End
      Begin VB.TextBox TxtRucPrv 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         MaxLength       =   11
         TabIndex        =   24
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox TxtDirPrv 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Top             =   540
         Width           =   8100
      End
      Begin VB.TextBox TxtNomPrv 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3030
         TabIndex        =   25
         Top             =   180
         Width           =   5760
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Monto Autorizado"
         Height          =   195
         Left            =   3900
         TabIndex        =   91
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   90
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Contable"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   88
         Top             =   960
         Width           =   960
      End
      Begin VB.Label LblCod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   960
         Width           =   420
      End
      Begin VB.Label lblocompra 
         BackStyle       =   0  'Transparent
         Caption         =   "O. Compra"
         Height          =   195
         Left            =   5820
         TabIndex        =   86
         Top             =   600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         TabIndex        =   69
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblord 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P.O."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7320
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direcc."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   65
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   225
         Width           =   345
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prov."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   2415
      Left            =   60
      TabIndex        =   71
      Top             =   1740
      Width           =   12735
      _Version        =   65536
      _ExtentX        =   22463
      _ExtentY        =   4260
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtpordetra 
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
         Left            =   10320
         MaxLength       =   2
         TabIndex        =   101
         Top             =   1200
         Width           =   300
      End
      Begin VB.Frame Frame1 
         Caption         =   "Constancia de depósito de Detracción"
         Height          =   850
         Left            =   7800
         TabIndex        =   95
         Top             =   1480
         Width           =   3135
         Begin VB.TextBox txtdetraccion 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   99
            Top             =   480
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker txtfechadetraccion 
            Height          =   315
            Left            =   1800
            TabIndex        =   96
            Top             =   480
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   134676481
            CurrentDate     =   39623
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   480
            TabIndex        =   98
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Emisión"
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
            Left            =   1920
            TabIndex        =   97
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.ComboBox cmbTipdocref 
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
         ItemData        =   "registro_compras.frx":000C
         Left            =   4800
         List            =   "registro_compras.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   1200
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtSerGuia 
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
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox TxtNumGuia 
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
         Left            =   7140
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox CboCategoria 
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
         Left            =   8940
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   165
         Width           =   1995
      End
      Begin VB.ComboBox cmbigv 
         Height          =   315
         ItemData        =   "registro_compras.frx":002B
         Left            =   120
         List            =   "registro_compras.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   165
         Width           =   4215
      End
      Begin VB.TextBox TxtTipCam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   765
         Width           =   915
      End
      Begin VB.TextBox TxtPoliza 
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   18
         Top             =   -180
         Width           =   1920
      End
      Begin VB.ComboBox cmbfpagos 
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
         Left            =   8460
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   765
         Width           =   2520
      End
      Begin VB.ComboBox CmbTipDoc 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   765
         Width           =   1875
      End
      Begin VB.TextBox TxtRefere 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   315
         Left            =   1800
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "este campo es obligatorio"
         Top             =   1980
         Width           =   5700
      End
      Begin VB.TextBox TxtNumDoc 
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
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   5
         Top             =   765
         Width           =   1185
      End
      Begin VB.TextBox TxtSerDoc 
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
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   4
         Top             =   765
         Width           =   540
      End
      Begin VB.TextBox txtimporta 
         Height          =   285
         Left            =   8160
         TabIndex        =   72
         Top             =   -180
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtcentro 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1600
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker TxtFecVen 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
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
         CurrentDate     =   39539
      End
      Begin MSComCtl2.DTPicker TxtFechaRec 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
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
         CurrentDate     =   39539
      End
      Begin CONTROLSLibCtl.dxColorBtn Mon 
         Height          =   405
         Left            =   6780
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Soles"
         Top             =   720
         Width           =   405
         _Version        =   65536
         _cx             =   714
         _cy             =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FillColor       =   65535
         ForeColor       =   12648447
         Caption         =   "MN"
         Enabled         =   -1  'True
         CaptionStringCount=   1
         GroupIndex      =   -1
         Stuck           =   -1  'True
         PictureLayout   =   1
         Pushed          =   0   'False
      End
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   765
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   134676481
         CurrentDate     =   39623
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   15
         Left            =   7980
         OleObjectBlob   =   "registro_compras.frx":002F
         TabIndex        =   89
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         TabIndex        =   102
         Top             =   1200
         Width           =   255
      End
      Begin CONTROLSLibCtl.dxCheckBox chkdetraccion 
         Height          =   270
         Left            =   8400
         TabIndex        =   100
         Top             =   1200
         Width           =   1755
         _Version        =   65536
         _cx             =   3096
         _cy             =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Afecto a Detracción"
         Enabled         =   -1  'True
         AutoSize        =   -1  'True
         BackStyle       =   1
         BackColor       =   -2147483633
         ForeColor       =   0
         ViewStyle       =   1
         Checked         =   0   'False
         GroupIndex      =   -1
         TextLayout      =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento  Referencia"
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
         Left            =   3000
         TabIndex        =   94
         Top             =   1250
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro"
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
         Left            =   120
         TabIndex        =   84
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9060
         TabIndex        =   83
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label LblFecVen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Vencimiento"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   82
         Top             =   1250
         Width           =   1365
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Height          =   225
         Left            =   120
         TabIndex        =   81
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T/C Oficial"
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
         Left            =   7560
         TabIndex        =   80
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Documento"
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
         Left            =   4440
         TabIndex        =   79
         Top             =   525
         Width           =   1305
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie        Nro. Documento"
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
         Left            =   2280
         TabIndex        =   78
         Top             =   525
         Width           =   1890
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento"
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
         Left            =   120
         TabIndex        =   77
         Top             =   525
         Width           =   1380
      End
      Begin VB.Label Lblpol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importación"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7260
         TabIndex        =   76
         Top             =   -120
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Costo"
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
         Left            =   120
         TabIndex        =   75
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label pnlcosto 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2880
         TabIndex        =   74
         Top             =   1605
         Width           =   4635
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
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
         Height          =   195
         Left            =   5940
         TabIndex        =   73
         Top             =   840
         Width           =   690
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2790
      Left            =   60
      OleObjectBlob   =   "registro_compras.frx":009F
      TabIndex        =   16
      Top             =   4185
      Width           =   10935
   End
   Begin VB.ComboBox cmbdocum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11670
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtdocpag 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   11790
      MaxLength       =   11
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin Threed.SSPanel PnlOficial 
      Height          =   60
      Left            =   11790
      TabIndex        =   35
      Top             =   120
      Width           =   105
      _Version        =   65536
      _ExtentX        =   185
      _ExtentY        =   106
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      Begin VB.TextBox TxtOtrImp 
         Height          =   330
         Index           =   1
         Left            =   5670
         TabIndex        =   37
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox TxtIgv 
         Height          =   330
         Index           =   1
         Left            =   4410
         TabIndex        =   36
         Top             =   315
         Width           =   870
      End
      Begin Threed.SSPanel PnlBasImp 
         Height          =   375
         Index           =   1
         Left            =   810
         TabIndex        =   38
         Top             =   315
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   661
         _StockProps     =   15
         ForeColor       =   -2147483640
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
         TabIndex        =   39
         Top             =   360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
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
         TabIndex        =   40
         Top             =   315
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483640
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
      Begin VB.Label PnlSigMon 
         Caption         =   "S/."
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   42
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Datos a Transferir al Registro de Compras Oficial"
         Height          =   240
         Left            =   270
         TabIndex        =   41
         Top             =   45
         Width           =   4695
      End
   End
   Begin Threed.SSPanel PnlAyuda 
      Height          =   510
      Left            =   8640
      TabIndex        =   44
      Top             =   7200
      Visible         =   0   'False
      Width           =   465
      _Version        =   65536
      _ExtentX        =   820
      _ExtentY        =   900
      _StockProps     =   15
      Caption         =   "F2 --> Ayuda                    F4 ---> Eliminar      "
      ForeColor       =   12582912
      BackColor       =   -2147483648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      FloodColor      =   12582912
   End
   Begin Threed.SSCommand CmdGrabar 
      Height          =   15
      Left            =   1935
      TabIndex        =   45
      Top             =   6645
      Visible         =   0   'False
      Width           =   30
      _Version        =   65536
      _ExtentX        =   53
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Grabar"
      ForeColor       =   -2147483640
   End
   Begin Threed.SSCommand CmdImprimir 
      Height          =   15
      Left            =   5400
      TabIndex        =   46
      Top             =   6645
      Visible         =   0   'False
      Width           =   30
      _Version        =   65536
      _ExtentX        =   -53
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   -2147483640
   End
   Begin Threed.SSCommand CmdSalir 
      Height          =   15
      Left            =   7155
      TabIndex        =   47
      Top             =   6645
      Visible         =   0   'False
      Width           =   45
      _Version        =   65536
      _ExtentX        =   -79
      _ExtentY        =   26
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   -2147483640
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   780
      Left            =   30
      TabIndex        =   48
      Top             =   7050
      Width           =   11235
      _Version        =   65536
      _ExtentX        =   19817
      _ExtentY        =   1376
      _StockProps     =   15
      ForeColor       =   -2147483640
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
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2925
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   405
         Width           =   1140
      End
      Begin VB.TextBox TxtOtrImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4140
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   405
         Width           =   1140
      End
      Begin VB.TextBox txtredsuma 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5985
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtredresta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6750
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtdcto 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   7515
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   405
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Checkdatos 
         Height          =   240
         Left            =   5535
         TabIndex        =   32
         Top             =   405
         Width           =   240
      End
      Begin Threed.SSPanel PnlTotal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   9120
         TabIndex        =   49
         Top             =   285
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   4
      End
      Begin Threed.SSPanel PnlMonIna 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1575
         TabIndex        =   21
         Top             =   405
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      End
      Begin Threed.SSPanel PnlBasImp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   405
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imp."
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
         Left            =   630
         TabIndex        =   59
         Top             =   135
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Inafecto"
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
         Left            =   1725
         TabIndex        =   58
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label PnlImpuesto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGV"
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
         Left            =   3345
         TabIndex        =   57
         Top             =   135
         Width           =   285
      End
      Begin VB.Label PnlOtrImp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percepción"
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
         Left            =   4290
         TabIndex        =   56
         Top             =   135
         Width           =   810
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Red"
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
         Left            =   6075
         TabIndex        =   55
         Top             =   135
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-Red"
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
         Left            =   6885
         TabIndex        =   54
         Top             =   135
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dsctos."
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
         Left            =   7650
         TabIndex        =   53
         Top             =   135
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
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
         Height          =   240
         Left            =   9420
         TabIndex        =   52
         Top             =   15
         Width           =   645
      End
      Begin VB.Label PnlSigMon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   10260
         TabIndex        =   51
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Datos"
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
         Left            =   5400
         TabIndex        =   50
         Top             =   135
         Width           =   555
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1320
      Left            =   45
      TabIndex        =   60
      Top             =   405
      Width           =   2100
      _Version        =   65536
      _ExtentX        =   3704
      _ExtentY        =   2328
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      Begin VB.ComboBox CboMeses 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtmesmov 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   0
         MaxLength       =   6
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtNumMov 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   43
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nº MOVIMIENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   180
         TabIndex        =   61
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   635
      ButtonWidth     =   2090
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "kjkkj"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar   "
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar  "
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir  "
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anticipo"
            Object.ToolTipText     =   "Buscar Anticipos a Proveedores"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir        "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   8460
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":5BCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":6166
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":6700
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":6C9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":7234
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":77CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "registro_compras.frx":7D68
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "O. Compra"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   66
      Top             =   1020
      Width           =   750
   End
End
Attribute VB_Name = "Registro_Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Af As New ADOFunctions
Dim Rs As New ADODB.Recordset
Dim intOrdenes      As Integer
Dim strOrdenes      As String
Dim cconex_form     As String
Dim cnn_form        As New ADODB.Connection
Dim CnTmp           As New ADODB.Connection
Dim contawin        As ADODB.Connection
Dim tabla11         As ADODB.Recordset
Dim wrutacosto      As String
Dim TbDocumento1    As ADODB.Recordset
Dim tbparametro11   As ADODB.Recordset
Dim tbfpagos1       As ADODB.Recordset
Dim Tbproveedor1    As ADODB.Recordset
Dim Tbdetcompras    As New ADODB.Recordset
Dim tbcomtab1       As ADODB.Recordset
Dim TbMes1          As ADODB.Recordset
Dim RSDETALLE       As ADODB.Recordset
Dim Tabla1          As ADODB.Recordset
Dim rsif5pla        As ADODB.Recordset
Dim tbregisdoc      As ADODB.Recordset
Dim tbregismov      As ADODB.Recordset
Dim TbDocumDet   As New ADODB.Recordset
Dim af5cc           As Boolean
Dim sw_ayuda            As Boolean
'Dim sw_nuevo_doc    As Boolean
Dim sw_nuevo_item   As Boolean
Dim sw_detalle      As Boolean
Dim sw_cabecera     As Boolean
Dim sw_grabanse     As Boolean
Dim sw_igv1         As Boolean
Dim flag            As Boolean
Dim amovs_cab(0 To 50)  As a_grabacion
Dim amovs_det(0 To 20)   As a_grabacion
Dim xtipmonoc       As String
Dim xfecvencoc      As String
Dim xforpagoc       As String
Dim xtipdoc         As String
Dim wIMPUESTO       As Double
Dim wotrimp         As Double
Dim swtc            As Boolean
Dim vafecto         As Double
Dim vinafecto       As Double
Dim rsvalor1        As ADODB.Recordset
Dim SqlCad             As String
Dim meslet          As String
Dim xnomgasto       As String
Dim xctacont        As String
Dim chequeo         As ADODB.Recordset
Dim WNUMERO         As String
Dim jmes            As String
Dim jmoneda        As String
Dim strPeriodo As String
Dim strRegistro As String

Public Property Get registro() As String
    registro = strRegistro
End Property

Public Property Let registro(ByVal vNewValue As String)
    strRegistro = vNewValue
End Property

Public Property Get Periodo() As String
    Periodo = strPeriodo
End Property

Public Property Let Periodo(ByVal vNewValue As String)
    strPeriodo = vNewValue
End Property


Private Sub CargarDatos()
Dim monedita        As String
Dim Tbproveedor2    As ADODB.Recordset

    Set Tbproveedor2 = New ADODB.Recordset
    Set Tbproveedor1 = New ADODB.Recordset
    If Len(TxtCodPrv.Text) > 0 Then
        SqlCad = "Select * from ef2proveedores where f2codprov='" & TxtCodPrv.Text & "'"
        If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
        Tbproveedor1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        If Not Tbproveedor1.EOF Then
            TxtNomPrv.Text = "" & Tbproveedor1.Fields("F2NOMPROV")
            TxtDirPrv.Text = "" & Tbproveedor1.Fields("F2DIRPROV")
            monedita = "" & Tbproveedor1.Fields("F2TIPMON")
            TxtRucPrv.Text = "" & Tbproveedor1.Fields("F2NEWRUC")
            
            
            SqlCad = "Select * from bf9gin where moneda='" & monedita & "' and tipo='P' and left(codigo,1)='P'"
            If Tbproveedor2.State = adStateOpen Then Tbproveedor2.Close
            Tbproveedor2.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not Tbproveedor2.EOF Then
                txtcodcta.Text = "" & Tbproveedor2.Fields("codigo")
                TxtTelPrv.Text = "" & Tbproveedor2.Fields("cuenta")
                If monedita = "S" Then
                    Mon.Caption = "US"
                    Call Mon_Click
                Else
                    Mon.Caption = "MN"
                    Call Mon_Click
                End If
            Else
                TxtTelPrv.Text = ""
                txtcodcta.Text = ""
                'txtcodcta.SetFocus
            End If
            Tbproveedor2.Close
        Else
            MsgBox "Codigo de Proveedor no existe", vbInformation, "Aviso"
            TxtCodPrv.Text = ""
            TxtNomPrv.Text = ""
            TxtRucPrv.Text = ""
            TxtDirPrv.Text = ""
            TxtTelPrv.Text = ""
            txtcodcta.Text = ""
            dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = False
        End If
        If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
    Else
        TxtCodPrv.Text = ""
        TxtNomPrv.Text = ""
        TxtRucPrv.Text = ""
        TxtDirPrv.Text = ""
        TxtTelPrv.Text = ""
        txtcodcta.Text = ""
        dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = False
    End If
    
End Sub

Private Sub Conf_Grid()
    
    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        '.Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        '.Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        '.Set (egoTabThrough)
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
        '.Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        
    End With
  
    Call AdicionaItem
    
End Sub

Private Sub CboCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(CboCategoria.Text) <> "" And sw_cabecera = False Then
            sw_cabecera = True
        End If
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
End Sub

Private Sub CboMeses_Click()
wmes = Format(CboMeses.ListIndex + 1, "00")
meslet = Choose(Val(wmes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
TxtFecha.value = "01/" & wmes & "/" & wanno
End Sub

Private Sub CboMeses_LostFocus()
wmes = Format(CboMeses.ListIndex + 1, "00")
meslet = Choose(Val(wmes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
TxtFecha.value = "01/" & wmes & "/" & wanno
End Sub

Private Sub cmbdocum_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub cmbfpagos_Click()
Dim tbfpagos1 As ADODB.Recordset

    Set tbfpagos1 = New ADODB.Recordset
    SqlCad = "Select * from ef2forpag where f2forpag='" & left(right(cmbfpagos.Text, 4), 3) & "'"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        TxtFecVen.value = Format(CVDate(TxtFecha.value) + tbfpagos1.Fields("f2dias"), "dd/mm/yyyy")
    End If
    tbfpagos1.Close
End Sub

Private Sub cmbfpagos_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub cmbfpagos_LostFocus()
Dim tbfpagos1 As ADODB.Recordset

    Set tbfpagos1 = New ADODB.Recordset
    SqlCad = "Select * from ef2forpag where f2forpag='" & left(right(cmbfpagos.Text, 4), 3) & "'"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        TxtFecVen.value = Format(CVDate(TxtFecha.value) + tbfpagos1.Fields("f2dias"), "dd/mm/yyyy")
    End If
    tbfpagos1.Close

End Sub


Private Sub cmbigv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(cmbigv.Text) <> "" And sw_cabecera = False Then
            sw_cabecera = True
        End If
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub cmbTipDoc_Click()
    dxDBGrid1.Dataset.Edit
    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
        dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "H"
        TxtSerDoc.Visible = True
        TxtNumDoc.Visible = True
'        Line1.Visible = True
        Label5.Caption = "Serie/Doc."
        TxtPoliza.Visible = False
    Else
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "POL" Then
                TxtSerDoc.Visible = False
                TxtNumDoc.Visible = False
'                Line1.Visible = False
                Label5.Caption = "Poliza"
                TxtPoliza.Visible = True
            Else
                dxDBGrid1.Columns.ColumnByFieldName("f3DEBHAB").value = "D"
                TxtSerDoc.Visible = True
                TxtNumDoc.Visible = True
'                Line1.Visible = True
                Label5.Caption = "Serie/Doc."
                TxtPoliza.Visible = False
            End If
    End If
    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Or UCase(right(Trim(CmbTipDoc.Text), 3)) = "DEB" Then
        Label11.Visible = True
        cmbTipdocref.Visible = True
        txtSerGuia.Visible = True
        TxtNumGuia.Visible = True
    Else
        Label11.Visible = False
        cmbTipdocref.Visible = False
        txtSerGuia.Visible = False
        TxtNumGuia.Visible = False
    End If
End Sub

Private Sub CmbTipDoc_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub CmbTipDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(CmbTipDoc.Text) <> "" And sw_cabecera = False Then
            sw_cabecera = True
        End If
'        TxtSerDoc.SetFocus
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub CmbTipDoc_LostFocus()

    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
        PnlImpuesto.Caption = "Retencion"
        PnlOtrImp.Caption = "I.E.S"
        Checkdatos.Visible = False
        Label14.Visible = False
        txtredsuma.Visible = False
        txtredresta.Visible = False
        txtdcto.Visible = False
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        txtredsuma.Text = Format(0, "0.00")
        txtredresta.Text = Format(0, "0.00")
        txtdcto.Text = Format(0, "0.00")
        CALCULANDO
    Else
        PnlImpuesto.Caption = "IGV"
        PnlOtrImp.Caption = "Otros Impuestos"
        Checkdatos.Visible = True
        Label14.Visible = True
        CALCULANDO
    End If
    
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)

    If dxDBGrid1.Columns.FocusedIndex = 6 Then
        sw_nuevo_item = True
        If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
            dxDBGrid1.Dataset.Post
        End If
        sw_nuevo_item = False
        
    End If
    
    
    
CALCULANDO
CALCULAR_TOTALES
End Sub


Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case "F3CANTIDAD", "F3IMPORTE"
    Text = Format(Text, "###,###,##0.00")
Case "F3PREUNI"
    Text = Format(Text, "###,###,##0.0000")
End Select

End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)

If KeyCode = 113 Then
    Select Case dxDBGrid1.Columns.FocusedIndex
    Case 1
        wBase = "G"
        wgastos = ""
        Sw_AyuCodProv = False
        ayuda_gastos.TipoConcepto = "E"
        ayuda_gastos.Show 1
        If Len(Trim(wgastos)) > 0 Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = wgastos
            dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctacont
            If sw = 0 Then
                dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = TxtRefere.Text
            Else
                dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomgasto
            End If
            dxDBGrid1.Dataset.Post
            PROCESO_CUENTA
            
            If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
                dxDBGrid1.Columns.FocusedIndex = 3
            End If
        End If
    Case 3
       If dxDBGrid1.Columns.ColumnByFieldName("f5codpro").ReadOnly = False Then
            wcodproducto = ""
            ayuda_productos.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3concepto").value = wdesproducto
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 4
            End If
            
       End If
    Case 4
        If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
            wcodcosto = ""
            'Ayuda_CENTROS.SelectInto = "'999','998'"
            Ayuda_Centros.Show 1
            If Len(Trim(wcodcosto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = wcodcosto
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 5
            End If
        End If
    End Select
End If
'        If dxDBGrid1.Columns.FocusedIndex = 1 Then
''             If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value <> 0 Then
''                 dxDBGrid1.Dataset.Edit
''                 dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value = ""
''                 dxDBGrid1.Columns.FocusedIndex = 1
''             End If
'             wCodGasto = ""
'             ayuda_gastos.Show 1
'             If Len(Trim(wCodGasto)) > 0 Then
'                dxDBGrid1.Dataset.Edit
'                dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").Value = wCodGasto
'                dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").Value = wCtaGasto
'                If Sw = 0 Then
'                    dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").Value = TxtRefere.Text
'                Else
'                    dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").Value = wNomGasto
'                End If
'                PROCESO_CUENTA
''                dxDBGrid1.Columns.FocusedIndex = 3
'            End If
'        End If
        
''jcg08        If dxDBGrid1.Columns.FocusedIndex = 3 Then
''           PROCESO_CUENTA
''           If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
''                Ayuda_CENTROS.Show 1
''                dxDBGrid1.Dataset.Edit
''                dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value = wcodcosto
''                dxDBGrid1.Columns.FocusedIndex = 4
''           End If
''        End If
''   End If

End Sub
Private Sub CargaMeses()
CboMeses.Clear
For i = 1 To 12
    CboMeses.AddItem left(dev_mes(i), 3)
Next
End Sub
Private Sub CargaCategoria()
Dim z As Integer
Dim xz As Integer
Dim RsCat As New ADODB.Recordset
csql = "select IntCodCategoria,StrDesCategoria from Categoria order by strdescategoria"
Set RsCat = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
CboCategoria.Clear
z = 0: xz = 0
If RsCat.RecordCount > 0 Then
    RsCat.MoveFirst
    Do While Not RsCat.EOF
        CboCategoria.AddItem RsCat!strdescategoria & Space(299) & Format(RsCat!intCodCategoria, "00000000")
        If RsCat!strdescategoria = "Proveedores" Then xz = z
        RsCat.MoveNext
        z = z + 1
    Loop
    CboCategoria.ListIndex = xz
End If
End Sub

Private Sub Form_Load()
    
    If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
    cnn_dbbancos.Open StrConexDbBancos
    
    CargaMeses
    CargaCategoria
'        If MDIBancos.TTabDock.DockedForms("MenuPrincipal").Visible = True Then
'            Me.left = MDIBancos.TTabDock.DockedForms.ITEM("MenuPrincipal").Panel.Width
'        Else
'            Me.left = 0
'        End If
'
    Me.top = 1050  '''El oficial es 1050
    Set contawin = New ADODB.Connection
    Set TbDocumento1 = New ADODB.Recordset
    Set tbparametro11 = New ADODB.Recordset
    Set tbfpagos1 = New ADODB.Recordset
    
    If cnn_ctrcom.State = adStateOpen Then cnn_ctrcom.Close
    cnn_ctrcom.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CTRCOM.MDB" & ";Persist Security Info=False"
    
    Me.Height = 8265
    Me.Width = 11430
    

    '-----------
    
    'With contawin
    '    .Provider = "Microsoft.JET.OLEDB.4.0; Data Source=" & wrutaconta & "\db_tabla.mdb; Persist Security Info=False"
    '    .Open
    'End With
    'sw_nuevo_documento = False
    If sw_nuevo_doc = False Then
    Else
        sw_nuevo_doc = True
        sw_nuevo_item = False
        sw_detalle = False
        sw_cabecera = False
    End If
    CboMeses.Enabled = True
    CboMeses.ListIndex = Val(wmes) - 1
    
    
    gcodppp = ""
    SqlCad = "Select * from documentos"
    If TbDocumento1.State = adStateOpen Then TbDocumento1.Close
    TbDocumento1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not TbDocumento1.EOF Then
        Do While Not TbDocumento1.EOF
            CmbTipDoc.AddItem TbDocumento1.Fields("F2DESDOC") + Space(100) + TbDocumento1.Fields("F2CODDOC") + TbDocumento1.Fields("F2ABREV")
            TbDocumento1.MoveNext
        Loop
    End If
    CmbTipDoc.ListIndex = 0
    TbDocumento1.Close
    
    SqlCad = "Select * from param_com where f1codemp='" & wF1Dir & "'"
    If tbparametro11.State = adStateOpen Then tbparametro11.Close
    tbparametro11.Open SqlCad, cnn_ctrcom, adOpenDynamic, adLockOptimistic
    If Not tbparametro11.EOF Then
        wIgv = Val("" & tbparametro11.Fields("F1IGV"))
        gretenc = Val("" & tbparametro11.Fields("F1RETENC") / 100)
        If tbparametro11.Fields("f1mes" & wmes) = "*" Then
            MsgBox "Mes está cerrado. Verifique.", 48, "Compras"
            Meses.Show 1
            meslet = Choose(Val(wmes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
            Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
        End If
    End If
    tbparametro11.Close
        
    txtmesmov.Text = wanno & wmes
    meslet = Choose(Val(wmes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
'    txtmes.Text = meslet
    If Month(Format(Date, "dd/mm/yyyy")) > Val(wmes) Then
        TxtFecha.value = CDate(Format("01/" & wmes & "/" & Year(Date), "dd/mm/yyyy"))
    Else
        TxtFecha.value = Format(Date, "dd/mm/yyyy")
    End If
    TxtFecVen.value = Format(Date, "dd/mm/yyyy")
    If Val(Month(Date)) <> CboMeses.ListIndex + 1 Then
        TxtFechaRec.value = Format(CVDate("01/" & Format(CboMeses.ListIndex + 1, "00") & "/" & wanno), "dd/mm/yyyy")
    Else
        TxtFechaRec.value = Format(Date, "dd/mm/yyyy")
    End If
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
        
    swtc = False
    SqlCad = "Select * from ef2forpag order by f2despag"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        Do While Not tbfpagos1.EOF
            If Len(Trim(tbfpagos1.Fields("f2tipo") & "")) > 0 Then
                If tbfpagos1.Fields("f2tipo") & "" = "F" Or tbfpagos1.Fields("f2tipo") & "" = "C" Then
                    cmbfpagos.AddItem tbfpagos1.Fields("f2despag") & Space(100) & tbfpagos1.Fields("f2forpag") & tbfpagos1.Fields("f2tipo") & ""
                End If
            End If
            tbfpagos1.MoveNext
        Loop
    End If
    tbfpagos1.Close
    cmbfpagos.ListIndex = 0
    
    TxtTipCam = Format(0, "0.000")
    TABLA_TEMPORAL
    BASE_TEMPORAL
    '***verifica si existe campo
'    Call Crea_Campo(cconex_formp, DBTable, "F3PREUNI", "Double", False, "0")
'    Call Crea_Campo(cconex_formp, DBTable, "F4SERGUI", "String", False, "")
'    Call Crea_Campo(cconex_formp, DBTable, "F4NUMGUI", "String", False, "")
'    Call Crea_Campo(cconex_formp, DBTable, "F3ORDEN", "String", False, "")
'    Call Crea_Campo(cconex_formp, "TEMP_OC", "F3ORDEN", "String", False, "")
'    Call Crea_Campo(StrConexDbBancos, "REGISMOV", "F3PREUNI", "Double", False, "0")
'    Call Crea_Campo(StrConexDbBancos, "REGISDOC", "F4SERGUI", "String", False, "")
'    Call Crea_Campo(StrConexDbBancos, "REGISDOC", "F4NUMGUI", "String", False, "")
'    Call Crea_Campo(StrConexDbBancos, "REGISMOV", "F3SERGUI", "String", False, "")
'    Call Crea_Campo(StrConexDbBancos, "REGISMOV", "F3NUMGUI", "String", False, "")
    'abre conexion temporal
    cnn_form.Open cconex_formp
    
    DELETEREC_N DBTable, cconex_formp, ""
    dxDBGrid1.Dataset.Refresh
    Conf_Grid
    
    If sw_nuevo_doc = True Then
        nuevo
        wtipprov = "N"
        sw_ayuda_provee = True
        If SwFrm = True Then
            Me.TxtRucPrv = PagoProv.txtruc.Text
            'Me.TxtCodPrv
            'Me.TxtNomPrv
            
        Else
            sw_frm = True
            wocompra = ""
            wcodcliprov = ""
            wRucCliProv = ""
            wnomcliprov = ""
            FrmName = Me.Name
            Ayuda_Proveedores.Show 1
            Unload Ayuda_Proveedores
        End If
        sw_ayuda_provee = False
        TxtCodPrv.Text = wcodcliprov
        If Len(Trim(Me.TxtCodPrv.Text)) > 0 Then
            wcodgasto = ObtenerCampo("EF2PROVEEDORES", "F2CODGAS", "F2CODPROV", Trim(TxtCodPrv.Text), "T", cnn_dbbancos)
            wnomgasto = ObtenerCampo("BF9GIN", "NOMBRE", "CODIGO", wcodgasto, "T", cnn_dbbancos)
            wctagasto = ObtenerCampo("BF9GIN", "CUENTA", "CODIGO", wcodgasto, "T", cnn_dbbancos)
        End If
        Me.MousePointer = vbDefault
        TxtCodPrv_KeyPress 13
        sw_cabecera = False
    Else
        sw_cabecera = False
        sw_nuevo_doc = False
        actualiza
    End If
'    dxDBGrid1.Enabled = False
    Me.MousePointer = vbDefault
    
            dxDBGrid1.Enabled = True
        'dxDBGrid1.SetFocus
        'dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").ButtonColumn.EditButtonStyle = ebsDown
        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ButtonColumn.EditButtonStyle = ebsDown
        
        csql = "SELECT U.MAIL "
        csql = csql & "FROM EF2TAREAS AS T INNER JOIN (EF2TAREAUSERS AS TU INNER JOIN EF2USERS AS U ON TU.F2CODUSER = U.F2CODUSER) ON T.F2CODTAREA = TU.F2CODTAREA "
        csql = csql & "WHERE T.F2CODTAREA='0002' and TU.F2CODUSER='" & wusuario & "'"
        
        Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
        editTo = ""
        If Rs.State = 0 Then Exit Sub
        If Rs.RecordCount > 0 Then
            Toolbar.Buttons(8).Visible = True
            Toolbar.Buttons(9).Visible = True
        Else
            Toolbar.Buttons(8).Visible = False
            Toolbar.Buttons(9).Visible = False
        End If
        
     txtmontoautorizado.Text = Format(MontoOrdPago, "###,###,##0.00")
    sw_cabecera = False
End Sub

Public Sub BASE_TEMPORAL()

    cnombase = "templus.mdb" '"TEMP_COM.MDB"
    cconex_formp = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & cnombase & "; Persist Security Info=False"
    
End Sub

Public Sub TABLA_TEMPORAL()

    DBTable = "temp_det"
    Rem NSE sqlcad = "(F3ITEM Text(3),F3GASTO Text(4),F3CTACON Text(12),F3CENCOS Text(8),F3CONCEPTO Text(100),F3IMPORTE Double,F3AFECTO Text(1),F3DEBHAB Text(1))"
    Rem NSE CREATETABLE_N DBTable, CStr(sqlcad), temp

End Sub

Private Sub Form_Unload(Cancel As Integer)
''VerificaAtencionDeLaOrden (txtocompra.Text)
'If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
'Set cnn_DbBancos = Nothing

End Sub

Private Sub VerificaAtencionDeLaOrden(pNumero_de_Orden As String, FLGELI As String)
Dim Amov(0 To 10) As a_grabacion
Dim RsX As New ADODB.Recordset
    csql = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, "
    csql = csql & "ORDEN.F3CODFAB, ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL,ORDEN.F4VB1,ORDEN.F4VB2,ORDEN.f4monto, "
    csql = csql & "ORDEN.F4NUMORD+Format(ORDEN.item,'000') AS LLave, "
    csql = csql & "orden.F3TOTAL- iif(REGISMOV.F3AFECTO='*',IIf(IsNull(articulos.f3importe),0,articulos.f3importe*1." & wIgv & "),IIf(IsNull(articulos.f3importe),0,articulos.f3importe)) AS Saldo "
    csql = csql & "FROM "
    csql = csql & "(("
    csql = csql & "SELECT IF3ORDEN.F3CENCOS AS CODCENTRO, "
    csql = csql & "'[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']' AS Grupo, "
    csql = csql & "EF2PROVEEDORES.F2NEWRUC, IF3ORDEN.ITEM, IF4ORDEN.F4NUMORD, IF4ORDEN.F4LOCAL, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, IF3ORDEN.F5NOMPRO, "
    csql = csql & "IF3ORDEN.F3PREUNI, IF3ORDEN.F3CANPRO, IF3ORDEN.F3TOTAL, IF3ORDEN.F5VALVTA,IF4ORDEN.F4VB1,IF4ORDEN.F4VB2,IF4ORDEN.f4monto "
    csql = csql & "FROM (IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
    csql = csql & "INNER JOIN IF3ORDEN ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
    csql = csql & "WHERE IF4ORDEN.F4ESTADO < 5 AND (((EF2PROVEEDORES.F2NEWRUC)='" & wRucCliProv & "')) "
    csql = csql & "ORDER BY IF4ORDEN.F4CENTRO, IF4ORDEN.F4NUMORD"
    csql = csql & ") AS ORDEN "
    csql = csql & "LEFT JOIN ("
    csql = csql & "SELECT REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, sum(REGISMOV.F3IMPORTE) as F3IMPORTE, REGISMOV.F3AFECTO "
    csql = csql & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV)"
    csql = csql & "GROUP BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
    csql = csql & "ORDER BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO"
    csql = csql & ") AS Articulos ON (ORDEN.F4NUMORD = Articulos.F4OCOMPRA) AND (ORDEN.F3CODFAB = Articulos.F5CODPRO)) "
    csql = csql & "LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.F3COSTO "
    csql = csql & "Where ((([orden].[f5valvta] - IIf(IsNull([articulos].[f3importe]), 0, [articulos].[f3importe])) > 0)) "
    csql = csql & "and ORDEN.F4NUMORD='" & pNumero_de_Orden & "' "
    csql = csql & "ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo"
    Set RsX = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
'    If Len(Trim(pNumero_de_Orden)) > 0 Then
'        If RsX.RecordCount <> 0 Then
'            RsX.MoveFirst
'            If wMonto2doVb >= Val(RsX!F4MONTO & "") Then
'                If RsX!f4vb1 = True Or RsX!f4vb2 = True Then
'                    Amov(0).Campo = "F4ESTADO": Amov(0).valor = 2: Amov(0).TIPO = "N"
'                Else
'                    Amov(0).Campo = "F4ESTADO": Amov(0).valor = 1: Amov(0).TIPO = "N"
'                End If
'            Else
'                If RsX!f4vb1 = True And RsX!f4vb2 = True Then
'                    Amov(0).Campo = "F4ESTADO": Amov(0).valor = 2: Amov(0).TIPO = "N"
'                Else
'                    Amov(0).Campo = "F4ESTADO": Amov(0).valor = 1: Amov(0).TIPO = "N"
'                End If
'            End If
'        Else
'            Amov(0).Campo = "F4ESTADO": Amov(0).valor = 3: Amov(0).TIPO = "N"
'        End If
'        GRABA_REGISTRO Amov, "if4orden", "M", 0, StrConexDbBancos, "f4numord='" & pNumero_de_Orden & "' and f4local='1'"
'    End If
    If FLGELI = "0" And Len(txtocompra.Text) > 0 Then
            If Val(traerCampo("IF4ORDEN", "f4estado", "f4numord", txtocompra.Text)) = 7 Then
                Amov(0).campo = "F4ESTADO": Amov(0).valor = 6: Amov(0).Tipo = "N"
                '03/12/2010
                wcodIDOP = traerCampo("IF4ORDEN_PAGO", "top 1 IDOP", "ORDEN", txtocompra.Text, " And Estado = '1' Order By IDOP Desc ")
                cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO = 6 Where IDOP = '" & wcodIDOP & "' ")
            Else
                If Val(traerCampo("IF4ORDEN", "f4estado", "f4numord", txtocompra.Text)) = 6 Then
                    Amov(0).campo = "F4ESTADO": Amov(0).valor = 6: Amov(0).Tipo = "N"
                    '03/12/2010
                    wcodIDOP = traerCampo("IF4ORDEN_PAGO", "top 1 IDOP", "ORDEN", txtocompra.Text, " And Estado = '1' Order By IDOP Desc ")
                    cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO = 6 Where IDOP = '" & wcodIDOP & "' ")
                Else
                    Amov(0).campo = "F4ESTADO": Amov(0).valor = 3: Amov(0).Tipo = "N"
                    '03/12/2010
                    wcodIDOP = traerCampo("IF4ORDEN_PAGO", "top 1 IDOP", "ORDEN", txtocompra.Text, " And Estado = '1' Order By IDOP Desc ")
                    cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO = 3 Where IDOP = '" & wcodIDOP & "' ")
                End If
            End If
    Else
            Amov(0).campo = "F4ESTADO": Amov(0).valor = 2: Amov(0).Tipo = "N"
            '03/12/2010
            wcodIDOP = traerCampo("IF4ORDEN_PAGO", "top 1 IDOP", "ORDEN", txtocompra.Text, " And Estado = '1' Order By IDOP Desc ")
            cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO = 2 Where IDOP = '" & wcodIDOP & "' ")
    End If
    GRABA_REGISTRO Amov, "if4orden", "M", 0, StrConexDbBancos, "f4numord='" & pNumero_de_Orden & "' and f4local='1'"

End Sub

Private Sub Mon_Click()
If Mon.Caption = "US" Then
    Mon.Caption = "MN"
    Mon.ForeColor = &H80FFFF
    Mon.FillColor = &HFFFF&
    Mon.Pushed = 0
    PnlSigMon(0).Caption = "MN"
    wMon = "S"
    PnlOficial.Visible = False
    PnlBasImp(0).BackColor = &HC0FFFF
    PnlMonIna(0).BackColor = &HC0FFFF
    TxtIgv(0).BackColor = &HC0FFFF
    TxtOtrImp(0).BackColor = &HC0FFFF
    PnlTotal(0).BackColor = &HC0FFFF
Else
    Mon.Caption = "US"
    Mon.ForeColor = &HC0FFC0
    Mon.FillColor = &H8000&
    Mon.Pushed = 0
    PnlSigMon(0).Caption = "US"
    wMon = "D"
    PnlBasImp(0).BackColor = &HC0FFC0
    PnlMonIna(0).BackColor = &HC0FFC0
    TxtIgv(0).BackColor = &HC0FFC0
    TxtOtrImp(0).BackColor = &HC0FFC0
    PnlTotal(0).BackColor = &HC0FFC0
End If
End Sub



Private Sub Text1_Change()

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
Dim X   As Variant
Select Case Trim(Button.Caption)
    Case "Nuevo"
        Call PROCEDIMIENTO_NUEVO
    Case "Grabar"
        CorrelaPagDcto = 0
        Sw_Graba_Registro = False
        If Me.TxtRefere.Text = "Este campo es obligatorio" Or TxtRefere.Text = "" Then
            MsgBox "El concepto es obligatorio.", vbCritical, wnomcia
            TxtRefere.SetFocus
            Exit Sub
        End If
        
        
        If Len(Trim(TxtRucPrv.Text)) = 0 Then
         
            MsgBox "El proveedor es obligatorio.", vbCritical, wnomcia
            
            TxtRucPrv.SetFocus
            Exit Sub
        End If
        'valida codigos de gasto
        strOrdenes = ""
        intOrdenes = 0
        For i = 1 To dxDBGrid1.Count
            dxDBGrid1.Dataset.RecNo = i
            If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("f3gasto").value & "")) = 0 Then
                MsgBox "El código de gasto es obligatorio.", vbCritical, wnomcia
                dxDBGrid1.Columns.FocusedIndex = 1
                Exit Sub
            End If
            If Trim(dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").value & "") <> strOrdenes Then
                strOrdenes = dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").value & ""
                intOrdenes = intOrdenes + 1
            End If
        Next
        'valida codcta
        If Len(Trim(txtcodcta.Text)) = 0 Then
         
            MsgBox "El código de gasto del proveedor es obligatorio.", vbCritical, wnomcia
            
            txtcodcta.SetFocus
            Exit Sub
        End If
        
        If Not (Val(TxtTipCam.Text & "") > 2 And Val(TxtTipCam.Text & "") < 4) Then
            MsgBox "El Tipo de cambio no es válido.", vbExclamation, wnomcia
            TxtTipCam.SetFocus
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        '*****
        dxDBGrid1.Dataset.Edit
        If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
             dxDBGrid1.Dataset.Post
             sw_detalle = True
        End If
        If sw_cabecera = True Or sw_detalle = True Then
            If MsgBox("¿Desea grabar el documento?", 36, wnomcia) = vbYes Then
                grabar
                sw_detalle = False
                sw_cabecera = False
            End If
        End If
        'valida provisonales
'            valida_provi_reg.Show 1
        Me.MousePointer = vbDefault
        SwFrm = True
        If Sw_Graba_Registro = True Then
            If MsgBox("Registro grabado. ¿Desea validar el Ruc en SUNAT?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                valida_sunat TxtRucPrv.Text
                MsgBox "La Situación/Estado de " & ruc_rsocial & " es " & ruc_estado & "/" & ruc_situacion, vbInformation, "CONTROL Plus!"
            End If
            'GRABA_CONTABILIDAD
            'MsgBox "Registro grabado.", vbInformation, wNomCia
            'Call Transfiere_Compras_Automatico(Val(wAnno), Val(wmes), TxtNumMov.Text)
        End If
        '*********************************************
        Call Toolbar_ButtonClick(Toolbar.Buttons(8))
        '*********************************************
    Case "Anticipo"
        csql = "SELECT U.MAIL "
        csql = csql & "FROM EF2TAREAS AS T INNER JOIN (EF2TAREAUSERS AS TU INNER JOIN EF2USERS AS U ON TU.F2CODUSER = U.F2CODUSER) ON T.F2CODTAREA = TU.F2CODTAREA "
        csql = csql & "WHERE T.F2CODTAREA='0002' and TU.F2CODUSER='" & wusuario & "'"
        
        Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
        editTo = ""
        If Rs.State = 0 Then Exit Sub
        If Rs.RecordCount > 0 Then
            Set tbregisdoc = New ADODB.Recordset
            SqlCad = "Select * from regisdoc where f4nummov='" & TxtNumMov.Text & "' AND F4MESMOV='" & txtmesmov.Text & "'"
            If tbregisdoc.State = adStateOpen Then tbregisdoc.Close
            tbregisdoc.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not tbregisdoc.EOF Then
                
                CorrelaPagDcto = tbregisdoc.Fields("F4CORRELA")
                BuscaAnticipo
            End If
            
        End If
    Case "ID_Modificar"
        Call PROCEDIMIENTO_NUEVO
        'frmhlmov.Show 1 jcg urgente
        CONSULTA
        sw_nuevo_doc = False
    Case "Eliminar"
        Me.MousePointer = vbHourglass
        elimina txtmesmov.Text, TxtNumMov.Text
        If sw_elimina = True Then
            sw_nuevo_doc = True
            nuevo
            dxDBGrid1.Dataset.Close
            DELETEREC_N DBTable, cconex_formp, ""
            AdicionaItem
        End If
        Me.MousePointer = vbDefault
        Unload Me
    Case "Imprimir"
        If cnn_form.State = 0 Then cnn_form.Open cconex_formp
        'cnn_form.Execute "DELETE * FROM VOU_CAB"
        'cnn_form.Execute "DELETE * FROM VOU_DET"
        DELETEREC_N "VOU_CAB", cconex_formp, ""
        DELETEREC_N "VOU_DET", cconex_formp, ""
        LLENA_TEMPCAB
        LLENA_TEMPDET
        Me.MousePointer = vbHourglass
        With Imp_RegCompra
            .Periodo = txtmesmov.Text
            .NumeroRegistro = TxtNumMov.Text
            .PageSettings.LeftMargin = 600
            .PageSettings.RightMargin = 600
            .Zoom = 90
            .Caption = "Comprobante de Compra"
            .Label1.Caption = wnomcia
            .Label19.Caption = Format(Date, "dd/mm/yyyy")
            .Label20.Caption = worigen
            .Field1.Text = txtmesmov.Text
            .Field2.Text = TxtNumMov.Text
            .Field3.Text = TxtFecha.value
            .Field4.Text = TxtRefere.Text
            If Mon.Caption = "MN" Then
                .Field5.Text = "Soles"
            Else
               .Field5.Text = "Dolares"
            End If
            .Field6.Text = Format(TxtTipCam.Text, "0.000")
            SqlCad = "Select *, "
            If Mon.Caption = "MN" Then
                SqlCad = SqlCad & "iif(f3debhab='D',f3importe,'') as f3debe,iif(f3debhab='H',f3importe,'') as f3haber "
            Else
                SqlCad = SqlCad & "iif(f3debhab='D',f3imported,'') as f3debe,iif(f3debhab='H',f3imported,'') as f3haber "
            End If
            SqlCad = SqlCad & " from contable "
            .DataControl1.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\templus.mdb; Persist Security Info=False" '"\db_conta.mdb; Persist Security Info=False"
            .DataControl1.Source = SqlCad
            .Show 1
        End With
        Me.MousePointer = vbDefault
    Case "ID_ImportarDatos"
    '    Me.MousePointer = vbhourglass
        ImpVales.Show 1
    '    Me.MousePointer = vbdefault
    Case "Calcular"
        Me.MousePointer = vbHourglass
        X = Shell("Calc.exe", 1)
        Me.MousePointer = vbDefault
    Case "Salir"
        If dxDBGrid1.Dataset.State = dsEdit Then
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
        End If
        If sw_cabecera = True Or sw_detalle = True Then
            
            If MsgBox("¿Desea Grabar el Movimiento?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                grabar
                sw_detalle = False
            End If
        End If
        If SwFrm = False And (sw_cabecera = True Or sw_detalle) Then
            Unload Me
            Lista_RegCompras.Show
        Else
            Unload Me
        End If
        SwFrm = False
'        Unload Me
End Select

End Sub


Private Sub BuscaAnticipo()
Dim nSaldoPag As Double, nSaldoAnt As Double
Dim cMensaje As String
'        If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
'        cnn_dbbancos.Open
        nSaldoPag = Val(ObtenerCampo("pag_dcto", "saldo", "correla", Str(CorrelaPagDcto), "N", cnn_dbbancos) & "")
        If CorrelaPagDcto > 0 And Val(nSaldoPag) > 0 Then
            'csql = "select * from pag_dcto where proveedor='" & Trim(TxtCodPrv.Text) & "' and saldo>0 AND ((F4OCOMPRA='" & Trim(txtocompra.Text) & "' AND mid(nro_comp,1,3) = 'Ant') OR mid(nro_comp,1,3) = 'CRE')"
            csql = "select * from pag_dcto where proveedor='" & Trim(TxtCodPrv.Text) & "' and saldo>0 AND (mid(nro_comp,1,3) = 'Ant' OR mid(nro_comp,1,3) = 'CRE' OR mid(nro_comp,1,3) = 'Che')"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, StrConexDbBancos, 3, 1
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    If left(Rs.Fields("NRO_COMP").value, 3) = "CRE" Then
                        cMensaje = "¿Desea Aplicar la N/C " & Mid(Rs!NRO_COMP, 4, 6) & "?" & vbCrLf
                    ElseIf left(Rs.Fields("NRO_COMP").value, 3) = "Che" Then
                        cMensaje = "¿Desea Aplicar el Cheque " & Mid(Rs!NRO_COMP, 4, 6) & "?" & vbCrLf
                    Else
                        cMensaje = "¿Desea Aplicar el Anticipo " & Mid(Rs!NRO_COMP, 4, 6) & "?" & vbCrLf
                    End If
                    cMensaje = cMensaje & vbTab & "Referencia: " & vbTab & Rs!f4glosa & "" & vbCrLf
                    cMensaje = cMensaje & vbTab & "Saldo: " & vbTab & vbTab & Format(Rs!Saldo & "", "###,###,##0.00") & vbCrLf
                    cMensaje = cMensaje & vbTab & "Moneda: " & vbTab & vbTab & IIf(Rs!Moneda = "D", "Dólares", "Soles") & vbCrLf
                    If MsgBox(cMensaje, vbQuestion + vbYesNo + vbDefaultButton2, wnomcia) = vbYes Then
        
                        nSaldoAnt = Val(Rs!Saldo & "")
                        If wMon <> Rs!Moneda Then
                            nTipCam = 0
                            Do While Not nTipCam > 0
                                nTipCam = Val(InputBox("Ingrese el Tipo de Cambio", "Aplicación de Anticipo") & "")
                            Loop
                            If wMon = "S" And Rs!Moneda = "D" Then
                                nSaldoAnt = nSaldoAnt * nTipCam
                            Else
                                nSaldoAnt = nSaldoAnt / nTipCam
                            End If
                        End If
                        'actualiza ctaxpagar
                        If nSaldoPag > nSaldoAnt Then
                            pimporte = nSaldoPag - nSaldoAnt
                            nSaldoPag = nSaldoPag - nSaldoAnt
                            nSaldoAnt = 0
                        ElseIf nSaldoPag = nSaldoAnt Then
                            pimporte = nSaldoPag
                            nSaldoAnt = 0
                            nSaldoPag = 0
                        Else
                            pimporte = nSaldoPag
                            nSaldoAnt = nSaldoAnt - nSaldoPag
                            nSaldoPag = 0
                            
                        End If
                        'regresa a su moneda original el saldo ant
                        If wMon <> Rs!Moneda Then
                            
                            Do While Not nTipCam > 0
                                nTipCam = Val(InputBox("Ingrese el Tipo de Cambio", "Aplicación de Anticipo") & "")
                            Loop
                            If wMon = "S" And Rs!Moneda = "D" Then
                                nSaldoAnt = nSaldoAnt / nTipCam
                            Else
                                nSaldoAnt = nSaldoAnt * nTipCam
                            End If
                        End If
                        'actualiza estado de cuenta del documento
                        amovs_cab(0).campo = "saldo": amovs_cab(0).valor = nSaldoPag: amovs_cab(0).Tipo = "N"
                        GRABA_REGISTRO amovs_cab, "PAG_DCTO", "M", "0", StrConexDbBancos, "CORRELA=" & CorrelaPagDcto
                        'actualiza anticipo
                        amovs_cab(0).campo = "saldo": amovs_cab(0).valor = nSaldoAnt: amovs_cab(0).Tipo = "N"
                        GRABA_REGISTRO amovs_cab, "PAG_DCTO", "M", "0", StrConexDbBancos, "CORRELA=" & Val(Rs!Correla & "")
                        '***
                        Do While Not nTipCam > 0
                                nTipCam = Val(InputBox("Ingrese el Tipo de Cambio", "Aplicación de Anticipo") & "")
                        Loop
                        'graba pag_mvto
                        If wMon = "S" Then
                            nimputado = pimporte
                            nimputaso = Val(Format(pimporte / nTipCam, "0.00"))
                        Else
                            nimputado = pimporte
                            nimputaso = Val(Format(pimporte * nTipCam, "0.00"))
                        End If
                        nannorepo = Val(right(Year(Rs!fch_comp), 2))
                        nrorepo = Val((Day(Rs!fch_comp)))
                        nrorepo = nrorepo & Val((Month(Rs!fch_comp)))
                        Dim cinsert As String
                        cinsert = "INSERT INTO PAG_MVTO " & _
                            "(PROVEEDOR,CORR_COMP,CORR_DCTO,TCAMBIO,IMPUTADO,IMPUTASO," & _
                            "ANO_REPO,NRO_REPO,FCH_MVTO,FCH_REPO) " & _
                            "VALUES('" & Trim(TxtCodPrv.Text) & "'," & Val(Rs!Correla & "") & "," & CorrelaPagDcto & "," & nTipCam & _
                            "," & nimputado & "," & nimputaso & "," & _
                            nannorepo & "," & nrorepo & ",CVDATE('" & Format(Rs!fch_comp, "mm/dd/yyyy") & "'),CVDATE('" & Format(Rs!fch_comp, "mm/dd/yyyy") & "'))"
                            
                            cnn_dbbancos.Execute (cinsert)
                            Actualiza_Log cinsert, cnn_dbbancos.ConnectionString
                        If Sw_Graba_Registro = True Then
                            If left(Rs.Fields("NRO_COMP").value, 3) = "CRE" Then
                                MsgBox "La N/C" & Mid(Rs!NRO_COMP, 4, 6) & " ha sido aplicado.", vbInformation, wnomcia
                            Else
                                MsgBox "El Anticipo " & Mid(Rs!NRO_COMP, 4, 6) & " ha sido aplicado.", vbInformation, wnomcia
                            End If
                            If nSaldoPag = 0 Then Exit Do
                            'MsgBox "El Anticipo " & Mid(Rs!nro_comp, 4, 6) & " ha sido aplicado.", vbInformation, wNomCia
                        End If
                    End If
                    Rs.MoveNext
                Loop
            Else
                MsgBox "No se encontraron anticipos por aplicar para esta O/C.", vbInformation, wnomcia
            End If
        End If
End Sub

Private Sub txtcentro_Change()

If Len(Trim(txtcentro.Text)) Mod 3 = 0 Then
    wcodcosto = txtcentro.Text
    pnlcosto.Caption = ObtenerCampo("centros", "f3abrev", "f3costo", txtcentro.Text, "T", cnn_dbbancos)
    If txtcentro.Text <> "998" Then
        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Visible = False
    Else
        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Visible = True
    End If
ElseIf Len(Trim(txtcentro.Text)) = 0 Then
    wcodcosto = ""
    pnlcosto.Caption = ""
End If
End Sub

Private Sub txtcentro_DblClick()
txtcentro_KeyDown 113, 0
End Sub

Private Sub txtcentro_GotFocus()
txtcentro.SelStart = 0: txtcentro.SelLength = Len(txtcentro.Text)
End Sub

Private Sub txtcentro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wcodcosto = ""
        'Ayuda_CENTROS.SelectInto = "'999'"
        Ayuda_Centros.Show 1
        Unload Ayuda_Centros
        Set Ayuda_Centros = Nothing
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            If wcodcosto <> "998" Then
                dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Visible = False
            Else
                dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Visible = True
            End If
            txtcentro.Text = wcodcosto
            pnlcosto.Caption = wdescosto
            txtcentro_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcentro_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtcentro_LostFocus()
Dim RSCONSULTA      As New ADODB.Recordset

    If sw_ayuda = False Then
        If Len(Trim(txtcentro.Text)) > 0 Then
            strSQL = "SELECT F3ABREV,F3CODCLI FROM CENTROS WHERE F3COSTO='" & txtcentro.Text & "'"
            
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            RSCONSULTA.Open strSQL, StrConexDbBancos, adOpenStatic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                pnlcosto.Caption = "" & RSCONSULTA.Fields(0)
            Else
                pnlcosto.Caption = "": txtcentro.Text = ""
                MsgBox "Código del Centro de Costo no existe. Verifique.", vbInformation, "Atención"
                txtcentro.SetFocus
            End If
            RSCONSULTA.Close
            Set RSCONSULTA = Nothing
'        Else
        End If
    End If
End Sub

Private Sub txtcodcta_Change()

    If Trim(txtcodcta.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtcodcta_LostFocus()
txtcodcta_KeyPress 13
End Sub

Private Sub TxtCodPrv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        Call CargarDatos
    End If
    
    If KeyCode = 113 Then
        Me.MousePointer = vbHourglass
        txtcodcta.Text = ""
        gcodprov = "" & TxtCodPrv.Text
        wtipprov = "N"
        sw_ayuda_provee = True
        Ayuda_Proveedores.Show 1
        sw_ayuda_provee = False
        TxtCodPrv.Text = wcodcliprov
        Me.MousePointer = vbDefault
        TxtCodPrv_KeyPress 13
    End If
    
End Sub

Private Sub TxtCodPrv_KeyPress(KeyAscii As Integer)
Dim i   As Integer

    Set Tbproveedor1 = New ADODB.Recordset
    If KeyAscii = 13 Then
        If TxtCodPrv.Text = "" Then
        Else
            If TxtCodPrv.Text = "9999" Then
                TxtNomPrv.Text = ""
                TxtDirPrv.Text = ""
                TxtTelPrv.Text = ""
                TxtRucPrv.Text = ""
                TxtRucPrv.SetFocus
            Else
                Call CargarDatos
            End If
        End If
        If Len(TxtCodPrv.Text) = 0 Then
        Else
        End If

        
        ''''Insercion de la Orden de Compra
        If wocompra = "*" Then
            
            dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = True
            SqlCad = "Select * from ef2proveedores where f2codprov='" & Trim(TxtCodPrv.Text) & "'"
            If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
            Tbproveedor1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not Tbproveedor1.EOF Then
                If Tbproveedor1.Fields("f2orden") = True Then
                    wtipoc = ""
                    importar_ocompra.Show 1
                    If Len(Trim(strOrdenCompra)) > 0 Then
                        txtocompra.Text = strOrdenCompra
                        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
                            llena_oc
                            If xtipmonoc = "S" Then
                                Mon.Caption = "MN"
                                Call Mon_Click
                            Else
                                Mon.Caption = "US"
                                Call Mon_Click
                            End If

                            For i = 0 To cmbfpagos.ListCount - 1
                                If left(right(cmbfpagos.List(i), 4), 3) = xforpagoc Then
                                    cmbfpagos.ListIndex = i
                                End If
                            Next

                            If IsDate(xfecvencoc) Then
                                TxtFecVen.value = Format(xfecvencoc, "DD/MM/YYYY")  'MIG
                            End If
                        'End If
                    End If
                    wcodcosto = 0
                Else
                    
                    dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = False
                End If
            End If
            Tbproveedor1.Close
        End If
    
        '''Fin de Orden de Compra
    End If
    
End Sub

Private Sub llena_oc()
On Error GoTo CapturaError

Dim tbif5 As ADODB.Recordset
Dim tbf4oc      As ADODB.Recordset
Dim tbf3oc      As ADODB.Recordset
Dim tbef2serv As ADODB.Recordset
Dim tbgastos As ADODB.Recordset
Dim tbtemp As ADODB.Recordset
Dim tabla As ADODB.Recordset

    Set tbif5 = New ADODB.Recordset
    Set tbf4oc = New ADODB.Recordset
    Set tbf3oc = New ADODB.Recordset
    Set tbef2serv = New ADODB.Recordset
    Set tbgastos = New ADODB.Recordset
    Set tbtemp = New ADODB.Recordset
    'Set tabla = New ADODB.Recordset

    'SqlCad = "Select * from temp_det"
    'DELETEREC_N "temp_det", cnn_form
    'If tabla.State = adStateOpen Then tabla.Close
    'tabla.Open SqlCad, cnn_form, adOpenDynamic, adLockOptimistic
    'If Not tabla.EOF Then
        'If Left(gorden_cs, 3) = "O/C" Then
            SqlCad = "Select * from if4orden where f4numord='" & strOrdenCompra & "'"
            If tbf4oc.State = adStateOpen Then tbf4oc.Close
            tbf4oc.Open SqlCad, StrConexDbBancos, 3, 1
            
            If tbf4oc.RecordCount > 0 Then
                tbf4oc.MoveFirst
'                If Len(Trim(tbf4oc!F4CODSOLICITUD)) > 0 Then
'                    wCodCosto = ""
'                    Do While Not Len(Trim(wCodCosto)) > 0
'                        Selecciona_MultiplesCentros.CodigoPedido = "" & tbf4oc!F4CODSOLICITUD
'                        Selecciona_MultiplesCentros.Show 1
'                        Unload Selecciona_MultiplesCentros
'                        Set Selecciona_MultiplesCentros = Nothing
'                    Loop
'                    'MsgBox wMultCentros
'                    SqlCad = "SELECT '" & strordencompra & "' AS F4NUMORD,'*' AS F5AFECTO, TB_DETSOLICITUD.cod_producto as F3CODPRO, TB_DETSOLICITUD.precio*TB_DETSOLICITUD.candis AS SALDO, TB_DETSOLICITUD.candis AS F3CANPRO, TB_DETSOLICITUD.F5CODCOSTO AS F3CENCOS "
'                    SqlCad = SqlCad & "From TB_DETSOLICITUD "
'                    SqlCad = SqlCad & "WHERE (((TB_DETSOLICITUD.cod_solicitud)='" & tbf4oc!F4CODSOLICITUD & "') "
'                    SqlCad = SqlCad & "AND ((TB_DETSOLICITUD.F5CODCOSTO) in (" & wCodCosto & ")))"
'
'                Else
                Dim wi As Integer
                wi = 0
                If tbf4oc.Fields("f4tipdoc") & "" = "02" Then
                    SqlCad = "SELECT IF3ORDEN.*, IIf([comprado].[f3afecto]='*',"
                    SqlCad = SqlCad & "IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe])*1." & wi & ","
                    SqlCad = SqlCad & "IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe])) AS Compra, "
                    SqlCad = SqlCad & "Val(Format([if3oRDEN].[f3total]-IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe]*1." & wi & "),'#.00')) AS Saldo "
                    SqlCad = SqlCad & "FROM IF3ORDEN LEFT JOIN "
                    SqlCad = SqlCad & "(SELECT REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, Sum(REGISMOV.F3IMPORTE) AS F3IMPORTE, "
                    SqlCad = SqlCad & "REGISMOV.F3AFECTO "
                    SqlCad = SqlCad & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) "
                    SqlCad = SqlCad & "GROUP BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
                    SqlCad = SqlCad & "ORDER BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO) "
                    SqlCad = SqlCad & " AS Comprado ON (IF3ORDEN.F3CODFAB = Comprado.F5CODPRO) "
                    SqlCad = SqlCad & "AND (IF3ORDEN.F4NUMORD = Comprado.F4OCOMPRA) "
                    SqlCad = SqlCad & "Where (((Format([if3oRDEN].[f3total] - IIf(IsNull([comprado].[f3importe]), 0, [comprado].[f3importe] * 1." & wi & "), '#.00')) > 0))"
                    SqlCad = SqlCad & " AND ((IF3ORDEN.F4NUMORD)='" & strOrdenCompra & "')"
                Else
'                    SqlCad = "SELECT IF3ORDEN.*, IIf([comprado].[f3afecto]='*',"
'                    SqlCad = SqlCad & "IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe])*1." & wigv & ","
'                    SqlCad = SqlCad & "IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe])) AS Compra, "
'                    SqlCad = SqlCad & "Val(Format([if3oRDEN].[f3total]-IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe]*1." & wigv & "),'#.00')) AS Saldo "
'                    SqlCad = SqlCad & "FROM IF3ORDEN LEFT JOIN "
'                    SqlCad = SqlCad & "(SELECT REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, Sum(REGISMOV.F3IMPORTE) AS F3IMPORTE, "
'                    SqlCad = SqlCad & "REGISMOV.F3AFECTO "
'                    SqlCad = SqlCad & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) "
'                    SqlCad = SqlCad & "GROUP BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
'                    SqlCad = SqlCad & "ORDER BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO) "
'                    SqlCad = SqlCad & " AS Comprado ON (IF3ORDEN.F3CODFAB = Comprado.F5CODPRO) "
'                    SqlCad = SqlCad & "AND (IF3ORDEN.F4NUMORD = Comprado.F4OCOMPRA) "
'                    SqlCad = SqlCad & "Where (((Format([if3oRDEN].[f3total] - IIf(IsNull([comprado].[f3importe]), 0, [comprado].[f3importe] * 1." & wigv & "), '#.00')) > 0))"
'                    SqlCad = SqlCad & " AND ((IF3ORDEN.F4NUMORD)='" & StrOrdenCompra & "')"
                    SqlCad = "SELECT IF3ORDENES.*, IIf(comprado.f3afecto='*',IIf(IsNull(comprado.f3importe),0,comprado.f3importe)*1.18,IIf(IsNull(comprado.f3importe),0,comprado.f3importe)) AS Compra, "
                    SqlCad = SqlCad & "Val(Format(IF3ORDENES.f3total-IIf(IsNull(comprado.f3importe),0,comprado.f3importe*1." & wIgv & "),'#.00')) AS Saldo "
                    SqlCad = SqlCad & "FROM (SELECT IF3ORDEN.F4NUMORD, IF3ORDEN.F3CODFAB, IF3ORDEN.F3CODPRO, Sum(IF3ORDEN.F3CANPRO) AS F3CANPRO, First(IF3ORDEN.F3PRECOS) AS F3PRECOS, "
                    SqlCad = SqlCad & "Sum(IF3ORDEN.F3TOTAL) AS F3TOTAL, IF3ORDEN.F5AFECTO, First(IF3ORDEN.F3IGV) AS F3IGV, First(IF3ORDEN.F3PREUNI) AS F3PREUNI, IF3ORDEN.F3PORDCT, "
                    SqlCad = SqlCad & "IF3ORDEN.F3CENCOS, Sum(IF3ORDEN.F5VALVTA) AS F5VALVTA, Sum(IF3ORDEN.F3PREVTA) AS F3PREVTA, IF3ORDEN.UNIDAD, IF3ORDEN.F5NOMPRO From IF3ORDEN "
                    SqlCad = SqlCad & "GROUP BY IF3ORDEN.F4NUMORD, IF3ORDEN.F3CODFAB, IF3ORDEN.F3CODPRO, IF3ORDEN.F5AFECTO, IF3ORDEN.F3PORDCT, IF3ORDEN.F3CENCOS, IF3ORDEN.UNIDAD, "
                    SqlCad = SqlCad & "IF3ORDEN.F5NOMPRO) AS IF3ORDENES LEFT JOIN (SELECT REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, Sum(REGISMOV.F3IMPORTE) AS F3IMPORTE, REGISMOV.F3AFECTO "
                    SqlCad = SqlCad & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV=REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV=REGISMOV.F4MESMOV) "
                    SqlCad = SqlCad & "GROUP BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
                    SqlCad = SqlCad & "ORDER BY REGISDOC.F4OCOMPRA, REGISMOV.F5CODPRO "
                    SqlCad = SqlCad & ") AS Comprado ON (IF3ORDENES.F3CODFAB = Comprado.F5CODPRO) AND (IF3ORDENES.F4NUMORD = Comprado.F4OCOMPRA) "
                    SqlCad = SqlCad & "WHERE (((Format([IF3ORDENES].[f3total]-IIf(IsNull([comprado].[f3importe]),0,[comprado].[f3importe]*1." & wIgv & "), '#.00')) > 0))"
                    SqlCad = SqlCad & " AND ((IF3ORDEN.F4NUMORD)='" & strOrdenCompra & "')"
                    
                End If
'                End If
            End If
            
            If tbf3oc.State = adStateOpen Then tbf3oc.Close
            tbf3oc.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
''        Else
''            sql = "Select * from if4servicio where f4numord='" & Val(strordencompra & "") & "'"
''            If tbf4oc.State = adStateOpen Then tbf4oc.Close
''            tbf4oc.Open sql, StrConexDbBancos, adOpenDynamic, adLockOptimistic
''
''            sql = "Select * from if3servicio where f4numord='" & Val(strordencompra & "") & "'"
''            If tbf3oc.State = adStateOpen Then tbf3oc.Close
''            tbf3oc.Open sql, StrConexDbBancos, adOpenDynamic, adLockOptimistic
''        End If
        
        If Not tbf4oc.EOF Then
            xtipmonoc = tbf4oc.Fields("f4tipmon") & ""
            xforpagoc = tbf4oc.Fields("f4forpag") & ""
            
            xtipdoc = tbf4oc.Fields("f4tipdoc") & ""
            If xtipmonoc = "S" Then
                Mon.Caption = "US"
                Call Mon_Click
            Else
                Mon.Caption = "MN"
                Call Mon_Click
            End If
            cmbfpagos.ListIndex = -1
            For i = 0 To cmbfpagos.ListCount - 1
                If left(right(cmbfpagos.List(i), 4), 3) = xforpagoc Then
                    cmbfpagos.ListIndex = i
                End If
            Next
            If IsDate(xfecvencoc) Then
                TxtFecVen.value = Format(xfecvencoc, "DD/MM/YYYY")  'MIG
            End If
            
            If InStr(wcodcosto, ",") > 0 Then
                txtcentro.Text = tbf3oc.Fields("f4centro") & ""
            Else
                txtcentro.Text = tbf4oc.Fields("f4centro") & ""
            End If
            TxtRefere.Text = UCase(tbf4oc.Fields("F4OBSERVA") & "")
            
            If IsDate(tbf4oc.Fields("f4fecven")) Then
                xfecvencoc = Format(tbf4oc.Fields("f4fecven") & "", "dd/mm/yyyy")
            End If
            If dxDBGrid1.Dataset.RecordCount = 1 Then
                XITEM = dxDBGrid1.Dataset.RecordCount - 1
            Else
                XITEM = dxDBGrid1.Dataset.RecordCount
            End If
            
            If Len(Trim(xtipdoc)) > 0 Then
                wabrev = ObtenerCampoWhere("DOCUMENTOS", "F2ABREV", "F2CODDOC", xtipdoc, "T", cnn_dbbancos, "")
                SeleccionaEnComboRight wabrev, CmbTipDoc
            End If

            'cnn_form.Execute ("Delete * from temp_oc")
            DELETEREC_N "TEMP_OC", cconex_formp, ""
            If Not tbf3oc.EOF Then
                Do While "" & tbf3oc.Fields("F4NUMORD") = strOrdenCompra & "" 'And Not tbf3oc.EOF
                    xgasto = "": xnomgasto = "": xctacont = ""
                    'If Len(Trim(strordencompra)) > 0 Then
                        SqlCad = "Select * from if5pla where f5codpro='" & tbf3oc.Fields("f3codpro") & "'"
                        If tbif5.State = adStateOpen Then tbif5.Close
                        tbif5.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                        If Not tbif5.EOF Then
                            
                            If Len(Trim(tbif5.Fields("f5ctacon") & "")) = 0 Then
                                xctacont = wctagasto
                            Else
                                xctacont = (tbif5.Fields("f5ctacon") & "")
                            End If
                            xgasto = ObtenerCampo("BF9GIN", "CODIGO", "CUENTA", xctacont & "", "T", cnn_dbbancos)
                            If xctacont = "999999" Then
                                xnomgasto = (tbif5.Fields("f5nompro") & "")
                                If Len(Trim(xnomgasto)) = 0 Then
                                    xnomgasto = "NO DEFINIDO"
                                End If
                            Else
                                xnomgasto = (tbif5.Fields("f5nompro") & "")
                                If Len(Trim(xnomgasto)) = 0 Then
                                    xnomgasto = "NO ENCONTRADO"
                                End If
                            End If
                        End If
                        tbif5.Close
                    'Else
                    '    SqlCad = "Select * from ef2servicios where f2codser='" & tbf3oc.Fields("f3codpro") & "'"
                    '    If tbef2serv.State = adStateOpen Then tbef2serv.Close
                    '    tbef2serv.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                    '    If Not tbef2serv.EOF Then
                    '        xgasto = tbef2serv.Fields("f3gasto") & ""
                    '    End If
                    '    tbef2serv.Close
                    'End If
                    'xgasto = "281"
                    'sql = " Select * from bf9gin where codigo='" & xgasto & "' and base='G' "
                    'If tbgastos.State = adStateOpen Then tbgastos.Close
                    'tbgastos.Open sql, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                    'If Not tbgastos.EOF Then
                        'xnomgasto = Trim(tbgastos.Fields("nombre") & "")
                        'xctacont = Trim(tbgastos.Fields("cuenta") & "")
                    'End If
                    'tbgastos.Close
                    'Call Crea_Campo(cconex_formp, "temp_oc", "f3preuni", "Double", False, "0")
                    SqlCad = "Select * from temp_oc where f5codpro='" & tbf3oc.Fields("f3codpro") & "'"
                    If cnn_form.State = 0 Then cnn_form.Open cconex_formp
                    Set tbtemp = New ADODB.Recordset
                    If tbtemp.State = adStateOpen Then tbtemp.Close
                    tbtemp.Open SqlCad, cnn_form, adOpenDynamic, adLockOptimistic

'                    If Not tbtemp.EOF Then
'                        cnn_form.Execute ("UPDATE TEMP_OC SET F3IMPORTE=F3IMPORTE+" & tbf3oc.Fields("SALDO") & " WHERE f5codpro='" & tbf3oc.Fields("f3codpro") & "'")
'                    Else
                        XITEM = XITEM + 1
                        tbtemp.AddNew
                        tbtemp.Fields("f3item") = XITEM
                        tbtemp.Fields("f3anoorden") = wanno
                        tbtemp.Fields("f3gasto") = xgasto
                        tbtemp.Fields("f3ctacon") = xctacont
                        tbtemp.Fields("f5codpro") = tbf3oc.Fields("f3codpro") & ""
                        tbtemp.Fields("f3concepto") = left(xnomgasto, 255)
                        'If Left(gorden_cs, 3) = "O/C" Then
                        '    If Mon.Caption = "MN" Then
                        '        tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
                        '    Else
                        '        tbtemp.Fields("f3importe") = tbf3oc.Fields("f5valvta") + tbtemp.Fields("f3importe")
                        '   End If
                        'Else
                        'DBLPORDESC = Val(tbf3oc.Fields("F3PORDCT") & "") / 100
                        If xtipdoc = "02" Then
                            If tbf3oc.Fields("f5afecto") & "" = "*" Then
                                tbtemp.Fields("f3importe") = Format(Val(tbf3oc.Fields("saldo") & "") / 0.9, "0.00")
                            Else
                                tbtemp.Fields("f3importe") = Val(tbf3oc.Fields("saldo") & "") '- (Val(tbf3oc.Fields("saldo") & "") * DBLPORDESC)
                            End If
                        Else
                            If tbf3oc.Fields("f5afecto") & "" = "*" Then
                                tbtemp.Fields("f3importe") = Format(Val(tbf3oc.Fields("saldo") & "") / (1 + (wIgv / 100)), "0.00")
                            'If DBLPORDESC <> 0 Then
                            Else
                                tbtemp.Fields("f3importe") = Format(Val(tbf3oc.Fields("saldo") & ""), "0.00") '- (Val(tbf3oc.Fields("saldo") & "") * DBLPORDESC)
                            End If
                        End If
                        'Else
                        '    tbtemp.Fields("f3importe") = Val(tbf3oc.Fields("saldo") & "")
                        'End If
                        'End If
                        tbtemp.Fields("f3afecto") = tbf3oc.Fields("f5afecto") & ""
                        If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                            tbtemp.Fields("F3DebHab") = "H"
                        Else
                            tbtemp.Fields("F3DebHab") = "D"
                        End If
                        tbtemp.Fields("f3cencos") = Trim(tbf3oc.Fields("F3CENCOS") & "")
                        tbtemp.Fields("f3cantidad") = Format(Val(tbtemp.Fields("f3importe")) / Val(tbf3oc.Fields("f3precos")), "0.00") 'Val(tbf3oc.Fields("f3canpro") & ""),"0.00")
                        tbtemp.Fields("f3PREUNI") = Format(Val(tbtemp.Fields("f3importe")) / IIf(Val(tbtemp.Fields("f3cantidad") & "") > 0, Val(tbtemp.Fields("f3cantidad") & ""), 1), "0.0000")
                        tbtemp.Fields("f3ORDEN") = Trim(tbf3oc.Fields("F4NUMORD") & "")
                        tbtemp.Update
                        
                    'End If
                    tbf3oc.MoveNext
                    If tbf3oc.EOF Then Exit Do
                    If tbf3oc.Fields("F4NUMORD") <> (strOrdenCompra & "") Then Exit Do
                Loop
            End If
        End If
        SqlCad = "Select * from temp_oc"
        If cnn_form.State = 1 Then cnn_form.Close
        
        'If tbtemp.State = 1 Then tbtemp.Close
        'tbtemp.Open SqlCad, cnn_form, 3, 1
        Set tbtemp = Af.OpenSQLForwardOnly(SqlCad, cconex_formp)
        If tbtemp.RecordCount > 0 Then
            tbtemp.MoveFirst
            LimpiaDetalle
            sw_nuevo_item = True
            vinafecto = 0: vafecto = 0
            Do While Not tbtemp.EOF
                dxDBGrid1.Dataset.Append
                dxDBGrid1.Columns.ColumnByFieldName("f3item").value = tbtemp.Fields("f3item") '+ dxDBGrid1.Dataset.RecordCount
                dxDBGrid1.Columns.ColumnByFieldName("f3gasto").value = tbtemp.Fields("f3gasto") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f3ctacon").value = tbtemp.Fields("f3ctacon") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f3cencos").value = tbtemp.Fields("f3cencos") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f3concepto").value = tbtemp.Fields("f3concepto")
                dxDBGrid1.Columns.ColumnByFieldName("f3importe").value = Format(tbtemp.Fields("f3importe"), "0.00")
                dxDBGrid1.Columns.ColumnByFieldName("f3afecto").value = IIf(tbtemp.Fields("f3afecto") = "*", True, False)
                dxDBGrid1.Columns.ColumnByFieldName("afecto").value = IIf(tbtemp.Fields("f3afecto") = "*", Format(tbtemp.Fields("f3importe"), "0.00"), 0)
                dxDBGrid1.Columns.ColumnByFieldName("inafecto").value = IIf(tbtemp.Fields("f3afecto") = "*", 0, Format(tbtemp.Fields("f3importe"), "0.00"))
                dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value = tbtemp.Fields("f5codpro")
                dxDBGrid1.Columns.ColumnByFieldName("f3cantidad").value = tbtemp.Fields("f3cantidad")
                dxDBGrid1.Columns.ColumnByFieldName("f3preuni").value = tbtemp.Fields("f3preuni")
                'dxDBGrid1.Columns(7).Value = tbtemp.Fields("f3debhab") & ""
                dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = tbtemp.Fields("f3debhab") & ""
                dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").value = tbtemp.Fields("f3ORDEN") & ""
                NUEVO_CALCULO
                tbtemp.MoveNext
            Loop
            'dxDBGrid1.Dataset.ADODataset.Requery
            If tbtemp.State = 1 Then tbtemp.Close
            Set tbtemp = Nothing
            PnlBasImp(0).Caption = Format(vafecto, "0.00")
            PnlMonIna(0).Caption = Format(vinafecto, "0.00")
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                wIMPUESTO = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * gretenc), "###,##0.00")
                'wotrimp = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * gfonavi), "###,##0.00")
            Else
                wIMPUESTO = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * wIgv / 100), "###,##0.00")
                wotrimp = Format(TxtOtrImp(0).Text, "###,##0.00")
            End If
            If Val(Format(PnlBasImp(0).Caption, "0.00")) > 0# Then
                TxtIgv(0).Text = Format(wIMPUESTO, "###,##0.00")
                TxtOtrImp(0).Text = Format(wotrimp, "###,##0.00")
            Else
                TxtIgv(0).Text = Format(wIMPUESTO, "###,##0.00")
            End If
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                 dxDBGrid1.Dataset.Post
            
            End If
            sw_nuevo_item = False
        End If
    'End If
    wocompra = ""
    CALCULANDO
    CALCULAR_TOTALES
    Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Resume Next
    Exit Sub
End Sub

Public Sub NUEVO_CALCULO()

    'vinafecto = 0: vafecto = 0
    If dxDBGrid1.Columns.ColumnByFieldName("f3afecto").value = True Then
        vafecto = vafecto + dxDBGrid1.Columns.ColumnByFieldName("f3importe").value
    Else
        vinafecto = vinafecto + dxDBGrid1.Columns.ColumnByFieldName("f3importe").value
    End If

End Sub
Private Sub LimpiaDetalle()
           ' dxDBGrid1.Dataset.Edit
           ' dxDBGrid1.Dataset.Post
            
            'cnn_form.Execute ("Delete From temp_det")
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                 dxDBGrid1.Dataset.Post
            
            End If
            dxDBGrid1.Dataset.Delete
            dxDBGrid1.Dataset.ADODataset.Requery
End Sub

Private Sub TxtCodPrv_DblClick()

    TxtCodPrv_KeyUp 113, 0

End Sub

Private Sub txtcodcta_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        llampro = 1
        txtcodcta.SetFocus
        Sw_AyuCodProv = True
        wgastos = ""
        ayuda_gastos.TipoConcepto = "E"
        ayuda_gastos.Show 1
        If Len(Trim(wgastos)) > 0 Then
            txtcodcta.Text = wgastos
            txtcodcta_KeyPress 13
        End If
    End If

End Sub

Private Sub txtcodcta_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Set tbcomtab1 = New ADODB.Recordset
    If KeyAscii = 13 Then
         txtcodcta.Text = UCase(txtcodcta.Text)
        gcodppp = txtcodcta.Text
        SqlCad = "Select * from bf9gin where codigo='" & gcodppp & "' and TIPO='P' "
        If tbcomtab1.State = adStateOpen Then tbcomtab1.Close
        tbcomtab1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        If Not tbcomtab1.EOF Then
           gcueppp = tbcomtab1.Fields("cuenta") & ""
        End If
        tbcomtab1.Close
        If TxtCodPrv.Text = "9999" Then
            SqlCad = "Select * from  ef2proveedores where f2codprov='" & Trim(TxtCodPrv.Text) & "'"
            If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
            Tbproveedor1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not Tbproveedor1.EOF Then
                TxtTelPrv.Text = gcueppp & Tbproveedor1.Fields("f2codcon") & ""
            End If
            Tbproveedor1.Close
        Else
            SqlCad = "Select * from bf9gin where codigo='" & Trim(txtcodcta.Text) & "' and base='G' AND TIPO='P' "
            If tbcomtab1.State = adStateOpen Then tbcomtab1.Close
            tbcomtab1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not tbcomtab1.EOF Then
                If tbcomtab1.Fields("conta") & "" = "*" Then
                    TxtTelPrv.Text = "" & Trim(gcueppp) & Trim(gsegppp)
                Else
                    TxtTelPrv.Text = "" & Trim(gcueppp)
                End If
            Else
                If Len(txtcodcta.Text) = 0 Then
                    'If MsgBox("Debe Ingresar el Codigo de Gasto", vbInformation + vbOKCancel, wNomCia) = vbOK Then
                    '    txtcodcta.SetFocus
                    'End If
                Else
                    txtcodcta.Text = ""
                    TxtTelPrv.Text = ""
                    If MsgBox("El Codigo ingresado no existe. Vuelva a Ingresarlo ", vbInformation + vbOKCancel, wnomcia) = vbOK Then
                        txtcodcta.SetFocus
                    End If
                End If
            End If
            tbcomtab1.Close
        End If
        If Len(txtcodcta.Text) <> 0 Then
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
            'txtfecha.SetFocus
                'CmbTipDoc.SetFocus
        End If
    End If
    
End Sub

Private Sub txtcodcta_DblClick()
    
    txtcodcta_KeyDown 113, 0
    
End Sub



Private Sub txtdcto_GotFocus()
    
    txtdcto.SelStart = 0: txtdcto.SelLength = Len(txtdcto.Text)
    
End Sub

Private Sub txtdcto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdcto.Text = Format(txtdcto, "0.00")
        CALCULAR_TOTALES
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        'TxtCodPrv.SetFocus
    End If

End Sub

Private Sub txtdcto_LostFocus()
    
    CALCULAR_TOTALES
    
End Sub

Private Sub TxtDirPrv_Change()

    If Trim(TxtDirPrv.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub TxtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    PresionaBotonMoneda KeyCode
    If KeyCode = 13 Then
        'TxtTipCam.SetFocus
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
End Sub

Private Sub txtfecha_LostFocus()
Dim tbcambios1 As ADODB.Recordset
    
    Set tbcambios1 = New ADODB.Recordset
    If IsDate(TxtFecha.value) = True Then
        If Format(wmes, "00") <> Format(Mid(TxtFecha.value, 4, 2), "00") Then
            If wanno & wmes < Year(TxtFecha.value) & Format(Mid(TxtFecha.value, 4, 2), "00") Then
                swtc = True
                MsgBox "La fecha del documento no puede ser mayor al mes de proceso. ", 48, wnomcia
                TxtFecha.SetFocus
            Else
                swtc = False
                MsgBox "El mes del documento no corresponde al mes de proceso. ", 48, wnomcia
                SqlCad = "Select * from cambios where fecha=CVDATE('" & TxtFecha.value & "') "
                If tbcambios1.State = adStateOpen Then tbcambios1.Close
                tbcambios1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                If Not tbcambios1.EOF Then
                    TxtTipCam.Text = tbcambios1.Fields("cambio")
                Else
                    TxtTipCam.Text = "0.000"
                End If
                TxtTipCam.SelStart = 0
                TxtTipCam.SelLength = Len(TxtTipCam.Text)
                tbcambios1.Close
                TxtFecVen.value = TxtFecha.value
            End If
        Else
            swtc = False
            SqlCad = "Select * from cambios where fecha= CVDATE('" & TxtFecha.value & "')"
            If tbcambios1.State = adStateOpen Then tbcambios1.Close
            tbcambios1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not tbcambios1.EOF Then
                TxtTipCam.Text = Format(tbcambios1.Fields("cambio"), "#0.000")
            End If
            tbcambios1.Close
            TxtFecVen.value = TxtFecha.value
        End If
        If Month(TxtFecha.value) < 3 And Year(TxtFecha.value) < 2012 Then wIgv = 19
    Else
        swtc = True
        MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
        TxtFecha.SetFocus
    End If

End Sub


Private Sub TxtFechaRec_Change()
Dim tbfpagos1 As ADODB.Recordset
If IsDate(TxtFechaRec.value) Then
    Set tbfpagos1 = New ADODB.Recordset
    SqlCad = "Select * from ef2forpag where f2forpag='" & left(right(cmbfpagos.Text, 4), 3) & "'"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        TxtFecVen.value = Format(CVDate(TxtFecha.value) + tbfpagos1.Fields("f2dias"), "dd/mm/yyyy")
    End If
    tbfpagos1.Close
End If
End Sub

Private Sub TxtFechaRec_CloseUp()
Dim tbfpagos1 As ADODB.Recordset

    Set tbfpagos1 = New ADODB.Recordset
    SqlCad = "Select * from ef2forpag where f2forpag='" & left(right(cmbfpagos.Text, 4), 3) & "'"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        TxtFecVen.value = Format(CVDate(TxtFecha.value) + tbfpagos1.Fields("f2dias"), "dd/mm/yyyy")
    End If
    tbfpagos1.Close
End Sub

Private Sub TxtFechaRec_DropDown()
Dim tbfpagos1 As ADODB.Recordset

    Set tbfpagos1 = New ADODB.Recordset
    SqlCad = "Select * from ef2forpag where f2forpag='" & left(right(cmbfpagos.Text, 4), 3) & "'"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        TxtFecVen.value = Format(CVDate(TxtFecha.value) + tbfpagos1.Fields("f2dias"), "dd/mm/yyyy")
    End If
    tbfpagos1.Close
End Sub

Private Sub TxtFecVen_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
If KeyCode = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub TxtIgv_GotFocus(Index As Integer)
    TxtIgv(Index).SelStart = 0
    TxtIgv(Index).SelLength = Len(TxtIgv(Index).Text)
End Sub

Private Sub TxtIgv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub txtimporta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End If
End Sub

Private Sub txtmes_DblClick()
 Meses.Show 1
    TxtMesMov_KeyDown 113, 0

End Sub


Private Sub TxtNomPrv_Change()

    If Trim(TxtNomPrv.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtnumdoc_Change()

    If Trim(TxtNumDoc.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txtnumdoc_GotFocus()
TxtNumDoc.SelStart = 0: TxtNumDoc.SelLength = Len(TxtNumDoc.Text)
End Sub

Private Sub TxtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub TxtNumGuia_Change()
dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").Visible = Not (VerificaGuiaCabecera)
dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").Visible = Not (VerificaGuiaCabecera)
End Sub

Private Sub TxtNumGuia_GotFocus()
TxtNumGuia.SelStart = 0: TxtNumGuia.SelLength = Len(TxtNumGuia.Text)
End Sub

Private Sub TxtNumGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub TxtNumGuia_LostFocus()
TxtNumGuia.Text = Format(TxtNumGuia.Text, "0000000")
End Sub

Private Sub txtocompra_DblClick()

    txtocompra_KeyDown 113, 0
    
End Sub

Private Sub txtocompra_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wtipoc = ""
        wcodcosto = ""
        wRucCliProv = TxtRucPrv.Text
        importar_ocompra.Show 1
        Unload importar_ocompra
        Set importar_ocompra = Nothing
        If Len(Trim(strOrdenCompra)) > 0 Then
            txtocompra.Text = strOrdenCompra
            
            llena_oc
                                    
            
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
            wcodcosto = ""
        End If
    End If

End Sub

Private Sub TxtOCompra_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtocompra.Text & "")) > 0 Then
            If wocompra = "*" Then
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

Private Sub prov_oc()
Dim para1 As ADODB.Recordset
Dim i       As Integer


    Set para1 = New ADODB.Recordset
    SqlCad = "Select * from if4orden where f4numord='" & txtocompra.Text & "'"
    If para1.State = adStateOpen Then para1.Close
    para1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
    If Not para1.EOF Then
        For i% = 0 To cmbfpagos.ListCount - 1
            If left(right(cmbfpagos.List(i%), 4), 3) = para1.Fields("f4forpag") & "" Then
                cmbfpagos.ListIndex = i%
            End If
        Next
    End If
    para1.Close

End Sub

Private Sub txtordcompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtOtrImp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub TxtPoliza_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtRefere_Change()
If LCase(TxtRefere.Text) <> "este campo es obligatorio" Then
    TxtRefere.Font.Italic = False
    TxtRefere.ForeColor = vbBlack
    'If sw_nuevo_doc = True Then TxtRefere.Text = ""
End If

End Sub

Private Sub TxtRefere_GotFocus()
If LCase(TxtRefere.Text) = "este campo es obligatorio" Then
    TxtRefere.Text = ""
    TxtRefere.Font.Italic = False
    TxtRefere.ForeColor = vbBlack
    If sw_nuevo_doc = True Then TxtRefere.Text = ""
Else
    TxtRefere.SelStart = 0: TxtRefere.SelLength = Len(TxtRefere.Text)
End If
End Sub

Private Sub TxtRefere_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub TxtRefere_LostFocus()
On Error Resume Next
If Len(Trim(TxtRefere.Text)) > 0 And LCase(TxtRefere.Text) <> "este campo es obligatorio" Then
        dxDBGrid1.Enabled = True
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").ButtonColumn.EditButtonStyle = ebsDown
        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ButtonColumn.EditButtonStyle = ebsDown
Else
    TxtRefere.Text = "este campo es obligatorio"
    TxtRefere.Font.Italic = True
    TxtRefere.ForeColor = vbRed
End If
End Sub

Private Sub TxtRucPrv_DblClick()
    Me.MousePointer = vbHourglass
        txtcodcta.Text = ""
        wtipprov = "N"
        sw_ayuda_provee = True
        Ayuda_Proveedores.Show 1
        Unload Ayuda_Proveedores
        sw_ayuda_provee = False
        TxtCodPrv.Text = wcodcliprov
        Me.MousePointer = vbDefault
        TxtCodPrv_KeyPress 13
End Sub

Private Sub TxtRucPrv_LostFocus()
TxtRucPrv_KeyPress 13
End Sub

Private Sub TxtSerDoc_Change()

    If Trim(TxtSerDoc.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub


Private Sub TxtOtrImp_GotFocus(Index As Integer)
    
    TxtOtrImp(Index).SelStart = 0
    TxtOtrImp(Index).SelLength = Len(TxtOtrImp(Index).Text)

End Sub

Private Sub TxtOtrImp_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Checkdatos.value = True Then
            TxtOtrImp(0).Text = Format(TxtOtrImp(0).Text, "0.00")
            Checkdatos.SetFocus
            CALCULAR_TOTALES
        Else
            TxtOtrImp(0).Text = Format(TxtOtrImp(0).Text, "0.00")
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
                CALCULAR_TOTALES
            Else
                Checkdatos.SetFocus
                CALCULAR_TOTALES
            End If
        End If
        Call Toolbar_ButtonClick(Toolbar.Buttons(3))
'        If MsgBox("¿Desea grabar el documento?", 36, wNomCia) = vbYes Then
'            Sw_Graba_Registro = False
'            If Me.TxtRefere.Text = "este campo es obligatorio" Or TxtRefere.Text = "" Then
'                MsgBox "El concepto es obligatorio.", vbCritical, wNomCia
'                TxtRefere.SetFocus
'                Exit Sub
'            End If
'            If Len(Trim(TxtRucPrv.Text)) = 0 Then
'
'                MsgBox "El proveedor es obligatorio.", vbCritical, wNomCia
'
'                TxtRucPrv.SetFocus
'                Exit Sub
'            End If
'            'valida codigos de gasto
'            For i = 1 To dxDBGrid1.Count
'                dxDBGrid1.Dataset.RecNo = i
'                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("f3gasto").Value & "")) = 0 Then
'                    MsgBox "El código de gasto es obligatorio.", vbCritical, wNomCia
'                    dxDBGrid1.Columns.FocusedIndex = 1
'                    Exit Sub
'                End If
'            Next
'            Me.MousePointer = vbhourglass
'            '*****
'            dxDBGrid1.Dataset.Edit
'            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
'                 dxDBGrid1.Dataset.Post
'                 sw_detalle = True
'            End If
'            If sw_cabecera = True Or sw_detalle = True Then
'
'                 Grabar
'                 sw_detalle = False
'                 sw_cabecera = False
'                If Sw_Graba_Registro = True Then
'                    MsgBox "Registro Grabado", vbOKOnly, wNomCia
'                End If
'            End If
'            BuscaAnticipo
'            Me.MousePointer = vbdefault
'
'        End If
'
    End If
    
    
    
    
End Sub

Private Sub TxtOtrImp_LostFocus(Index As Integer)

    CALCULAR_TOTALES
    
End Sub

Private Sub txtredresta_GotFocus()
    
    txtredresta.SelStart = 0: txtredresta.SelLength = Len(txtredresta.Text)
    
End Sub

Private Sub txtredresta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtredresta = Format(txtredresta, "0.00")
        CALCULAR_TOTALES
        txtdcto.SetFocus
    End If

End Sub

Private Sub txtredresta_LostFocus()
    
    CALCULAR_TOTALES
    
End Sub

Private Sub txtredsuma_GotFocus()
    
    txtredsuma.SelStart = 0: txtredsuma.SelLength = Len(txtredsuma.Text)
    
End Sub

Private Sub txtredsuma_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtredsuma.Text = Format(txtredsuma, "0.00")
        CALCULAR_TOTALES
        txtredresta.SetFocus
    End If
    
End Sub

Private Sub txtredsuma_LostFocus()
    
    CALCULAR_TOTALES
    
End Sub

Private Sub TxtRucPrv_Change()

    If Trim(TxtRucPrv.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
    If Len(TxtRucPrv.Text) = 0 Then
        TxtCodPrv.Text = ""
        TxtNomPrv.Text = ""
        TxtDirPrv.Text = ""
        TxtRucPrv.Text = ""
        txtcodcta.Text = ""
        TxtTelPrv.Text = ""
        Call Mon_Click
    End If
    If Len(TxtRucPrv.Text) = 11 Then
        TxtRucPrv_KeyPress 13
    End If
End Sub

Private Sub TxtRucPrv_KeyPress(KeyAscii As Integer)
Dim i   As Integer
On Error Resume Next
    Set Tbproveedor1 = New ADODB.Recordset
    Set tbcomtab1 = New ADODB.Recordset
    
    If KeyAscii = 13 Then
        If TxtCodPrv.Text = "" And TxtRucPrv = "" Then
            MsgBox "Debe Ingresar el Ruc del Proveedor", vbInformation, "Atencion"
            TxtRucPrv.SetFocus
        Else
            If TxtCodPrv.Text = "9999" Then
                txtcodcta.SetFocus
            Else
                SqlCad = "Select * from ef2proveedores where f2newruc='" & Trim(TxtRucPrv.Text) & "'"
                If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
                Tbproveedor1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                If Not Tbproveedor1.EOF Then
                    Call SeleccionaEnComboTipDoc("" & Tbproveedor1.Fields("F2tipdoc"), CmbTipDoc)
                    Call SeleccionaEnComboForPag("" & Tbproveedor1.Fields("F2FORPAG"), cmbfpagos)
                    Call SeleccionaEnComboRight(Format("" & Tbproveedor1!intCodCategoria, "00000000"), CboCategoria)
                    
                    wcodgasto = "" & Tbproveedor1.Fields("F2CODGAS")
                    wnomgasto = ObtenerCampo("BF9GIN", "NOMBRE", "CODIGO", wcodgasto, "T", cnn_dbbancos)
                    wctagasto = ObtenerCampo("BF9GIN", "CUENTA", "CODIGO", wcodgasto, "T", cnn_dbbancos)
                    
                    TxtCodPrv.Text = "" & Tbproveedor1.Fields("F2CODPROV")
                    TxtNomPrv.Text = "" & Tbproveedor1.Fields("F2NOMPROV")
                    TxtDirPrv.Text = "" & Tbproveedor1.Fields("F2DIRPROV")
                    TxtRucPrv.Text = "" & Tbproveedor1.Fields("F2NEWRUC")
                    If Tbproveedor1.Fields("F2TIPMON") = "S" Then
                        Mon.Caption = "US"
                        Call Mon_Click
                    Else
                        Mon.Caption = "MN"
                        Call Mon_Click
                    End If
                    gsegppp = Tbproveedor1.Fields("F2CODCON") & ""
                    gmoneda = Tbproveedor1.Fields("F2TIPMON") & ""
            
                    SqlCad = "Select * from bf9gin where moneda='" & gmoneda & "' and tipo='P' and left(codigo,1)='P' and left(codigo,1)='P'"
                    If tbcomtab1.State = adStateOpen Then tbcomtab1.Close
                    tbcomtab1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                    If Not tbcomtab1.EOF Then
                        txtcodcta.Text = "" & tbcomtab1.Fields("codigo")
                        TxtTelPrv.Text = "" & tbcomtab1.Fields("cuenta")
                        If tbcomtab1.Fields("MONEDA") = "S" Then
                            Mon.Caption = "US"
                            Call Mon_Click
                        Else
                            Mon.Caption = "MN"
                            Call Mon_Click
                        End If
                    End If
                    If Len(TxtRucPrv.Text) <> 0 Then
                        wocompra = IIf(ObtenerCampo("EF2PROVEEDORES", "F2ORDEN", "F2NEWRUC", Trim(TxtRucPrv.Text), "T", cnn_dbbancos) = True, "*", "")
                        wRucCliProv = Trim(TxtRucPrv.Text)
                    End If
                    If wocompra = "*" And Len(Trim(txtocompra.Text)) = 0 Then
                        dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = True
                        If sw_nuevo_documento = True Then
                        SqlCad = "Select * from ef2proveedores where f2codprov='" & Trim(TxtCodPrv.Text) & "'"
                        If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
                        Tbproveedor1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
                        If Not Tbproveedor1.EOF Then
                            If Tbproveedor1.Fields("f2orden") = True Then
                                wtipoc = ""
                                strOrdenCompra = ""
                                'LimpiaDetalle
                                DELETEREC_N "temp_oc", cconex_formp, ""
                                swActOrden = True
                                Do While swActOrden = True
'                                    proceso_grid_ordenes
                                    importar_ocompra.Show 1
                                    Unload importar_ocompra
                                    Set importar_ocompra = Nothing
                                Loop
                                
                                wocompra = ""
                                If Len(Trim(strOrdenCompra)) > 0 Then
                                    txtocompra.Text = strOrdenCompra
                                    llena_oc
                                    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
                                    strOrdenCompra = ""
                                End If
                            End If
                        End If
                        Tbproveedor1.Close
                      End If
                      On Error Resume Next
                      TxtFecha.SetFocus
End If
                    '''Fin de Orden de Compra
                '''Viene el caso de que no exista el Ruc
                Else
                    If MsgBox("El Ruc no existe. ¿Desea registrarlo?", 36, "Atención") = 6 Then
                        mostrar = True
                        sw_nuevo_doc = False
                        sw_ayuda_provee = True
                        Mant_Proveedores.Show 1
                        sw_ayuda_provee = False
                        sw_nuevo_doc = True
                        TxtRucPrv.Text = ruc
                        TxtCodPrv.Text = codpro
                        TxtNomPrv.Text = nombre
                        TxtDirPrv.Text = direccion
                    End If
                End If
                'Tbproveedor1.Close
            End If
        End If
        If CmbTipDoc.Enabled = True Then
        End If
    End If

End Sub

Private Sub TxtSerDoc_GotFocus()
TxtSerDoc.SelStart = 0: TxtSerDoc.SelLength = Len(TxtSerDoc.Text)
End Sub

Private Sub TxtSerDoc_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Function PresionaBotonMoneda(CodigoBoton As Integer)
If CodigoBoton = 121 Then
    Call Mon_Click
End If
End Function


Private Sub TxtSerDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then TxtNumDoc.SetFocus
    

End Sub

Private Sub TxtSerDoc_LostFocus()
TxtSerDoc.Text = Format(TxtSerDoc.Text, "000")
End Sub

Private Sub txtSerGuia_Change()
dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").Visible = Not (VerificaGuiaCabecera)
dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").Visible = Not (VerificaGuiaCabecera)
End Sub

Private Sub txtSerGuia_GotFocus()
txtSerGuia.SelStart = 0: txtSerGuia.SelLength = Len(txtSerGuia.Text)
End Sub

Private Sub txtSerGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtSerGuia_LostFocus()
txtSerGuia.Text = Format(txtSerGuia.Text, "000")
End Sub

Private Sub TxtTelPrv_Change()

    If Trim(TxtTelPrv.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub txttipcam_GotFocus()

    TxtTipCam.SelStart = 0
    TxtTipCam.SelLength = Len(TxtTipCam.Text)
    
End Sub

Private Sub TxtTipCam_KeyDown(KeyCode As Integer, Shift As Integer)
PresionaBotonMoneda KeyCode
End Sub

Private Sub txttipcam_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If Val(TxtTipCam.Text) > 0# Then
            TxtTipCam.Text = Format(TxtTipCam.Text, "0.000")
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
        Else
            'MsgBox "Ingrese tipo de cambio", vbExclamation, "Atencion"
            If MsgBox("Ingrese tipo de cambio", vbExclamation + vbOKCancel, wnomcia) = vbOK Then
                TxtTipCam.SetFocus
            End If
        End If
    End If

End Sub

Private Sub txttipcam_LostFocus()
On Error Resume Next
    If swtc = False Then
        If Val(TxtTipCam.Text) = 0# And Len(Trim(TxtNumDoc.Text)) > 0 Then
            If MsgBox("Ingrese tipo de cambio", vbExclamation + vbOKCancel, wnomcia) = vbOK Then
                TxtTipCam.SetFocus
            End If
        End If
    End If
    TxtTipCam.Text = Format(TxtTipCam.Text, "0.000")

End Sub

Private Sub cmbfpagos_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If Trim(cmbfpagos.Text) <> "" And sw_cabecera = False Then
            sw_cabecera = True
        End If
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        'TxtFecVen.SetFocus
    End If

End Sub

Private Sub txtfecha_GotFocus()

    'TxtFecha.FocusSelect = True

End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'TxtTipCam.SetFocus
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtFecVen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If IsDate(TxtFecVen.value) = True Then
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        Else
            MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
            TxtFecVen.SetFocus
        End If
    End If

End Sub


Private Sub TxtFechaRec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If IsDate(TxtFechaRec.value) = True Then
            'cmbigv.SetFocus
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        Else
            MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
            TxtFechaRec.SetFocus
        End If
    End If
End Sub

Private Sub TxtRefere_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1.Enabled = True
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").ButtonColumn.EditButtonStyle = ebsDown
        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ButtonColumn.EditButtonStyle = ebsDown
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
    Else
        'KeyAscii = KeyAscii
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub nuevo()
    CboMeses.Enabled = True
    LblFecVen.Visible = True
    TxtFecVen.Visible = True
    TxtCodPrv.Text = Empty
    TxtNomPrv.Text = ""
    TxtRucPrv.Text = ""
    TxtDirPrv.Text = ""
    txtcodcta.Text = ""
    TxtTelPrv.Text = ""
    txtcentro.Text = ""
    dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = True
    Call Mon_Click
    PnlSigMon(0).Caption = "MN"
    TxtSerDoc.Text = ""
    TxtNumDoc.Text = ""
    TxtPoliza.Text = ""
    txtimporta.Text = ""
    TxtTipCam.Text = Format(0, "0.000")
    TxtRefere.Text = "este campo es obligatorio"
    TxtRefere.Font.Italic = True
    TxtRefere.ForeColor = vbRed
    gcodppp = ""
    txtocompra.Text = ""
    TxtFecha.value = Format(Date, "dd/mm/yyyy")
    TxtFecVen.value = Format(Date, "dd/mm/yyyy")
    If Val(Month(Date)) <> CboMeses.ListIndex + 1 Then
        TxtFechaRec.value = Format(CVDate("01/" & Format(CboMeses.ListIndex + 1, "00") & "/" & wanno), "dd/mm/yyyy")
    Else
        TxtFechaRec.value = Format(Date, "dd/mm/yyyy")
    End If
    PnlBasImp(0).Caption = "0.00"
    PnlBasImp(1).Caption = "0.00"
    PnlMonIna(0).Caption = "0.00"
    PnlMonIna(1).Caption = "0.00"
    TxtIgv(0).Text = "0.00"
    TxtIgv(1).Text = "0.00"
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
    TxtNumMov.Text = ""
    dxDBGrid1.Columns.FocusedIndex = 1
    PnlOtrImp.Caption = "IGV"
    PnlOtrImp.Caption = "Otros Impuestos"
    Label14.Visible = True
    Checkdatos.Visible = True
    Checkdatos.value = 0
    txtredsuma.Visible = False
    txtredresta.Visible = False
    txtdcto.Visible = False
    dxDBGrid1.Enabled = False
    TxtSerDoc.Locked = False
    TxtNumDoc.Locked = False
    CmbTipDoc.Locked = False
    wcodprov = ""
End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer
 
    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_doc = False Then
        DELETEREC_N DBTable, cconex_formp, ""
        dxDBGrid1.Dataset.Refresh
    End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cconex_formp
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        For i = 1 To 1
            If sw_nuevo_temp = True Then
                If sw_nuevo_doc = True Then
                    .Edit
                Else
                    .Append
                End If
                    sw_nuevo_temp = True
                Else
                    .Append
            End If
            .FieldValues("F3ITEM") = i

            .FieldValues("F3GASTO") = ""
            .FieldValues("F3CTACON") = ""
            .FieldValues("F3CENCOS") = ""
            .FieldValues("F3CONCEPTO") = ""
            .FieldValues("F3IMPORTE") = Format(0, "###,##0.00")
            .FieldValues("F3AFECTO") = dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").CheckColumn.ValueChecked
            .FieldValues("F3DEBHAB") = "D"
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "H"
            Else
                dxDBGrid1.Columns.ColumnByFieldName("f3DEBHAB").value = "D"
            End If
            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").ButtonColumn.EditButtonStyle = ebsSimple
            dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ButtonColumn.EditButtonStyle = ebsSimple
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
               
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "H"
            Else
                dxDBGrid1.Columns.ColumnByFieldName("f3DEBHAB").value = "D"
            End If
            dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
                    dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = wcodgasto
                    dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctagasto
                    dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomgasto
            dxDBGrid1.Dataset.FieldValues("F3AFECTO") = dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").CheckColumn.ValueChecked
            dxDBGrid1.Columns.FocusedIndex = 1
            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").ButtonColumn.EditButtonStyle = ebsDown
            dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ButtonColumn.EditButtonStyle = ebsDown
        End If
    End If
           
End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns(1).value & "")) = 0 Then
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

Dim gassto As String
Dim rsgasto As New ADODB.Recordset
Dim amovs(0 To 3)  As a_grabacion

    If dxDBGrid1.Columns.FocusedIndex = 4 Then
       'PROCESO_CUENTA
       If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
            'Ayuda_CENTROS.SelectInto = "'999','998'"
            Ayuda_Centros.Show 1
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = wcodcosto
            dxDBGrid1.Dataset.Post
            dxDBGrid1.Columns.FocusedIndex = 5
       End If
    ElseIf dxDBGrid1.Columns.FocusedIndex = 3 Then
       'PROCESO_CUENTA
       If dxDBGrid1.Columns.ColumnByFieldName("f5codpro").ReadOnly = False Then
            ayuda_productos.Show 1
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value = wcodproducto
            dxDBGrid1.Columns.ColumnByFieldName("f3concepto").value = wdesproducto
            
            dxDBGrid1.Dataset.Post
            dxDBGrid1.Columns.FocusedIndex = 4
       End If
    ElseIf UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) = "F3ORDEN" Then
           'If KeyCode = 113 Then
                wtipoc = ""
                strOrdenCompra = ""
                wRucCliProv = TxtRucPrv.Text
                importar_ocompra.Show 1
                Unload importar_ocompra
                Set importar_ocompra = Nothing
                If Len(Trim(strOrdenCompra)) > 0 Then
                    txtocompra.Text = strOrdenCompra
                    llena_oc
                    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
                    strOrdenCompra = ""
                End If
            'End If
    ElseIf UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) = "F3CTACON" Then
            wctacont = "": wnomctacont = ""
            Ayuda_PlanCta.Show 1
            If Len(Trim(wctacont)) > 0 Then
                gassto = ObtenerCampo("BF9GIN", "CODIGO", "CUENTA", wctacont, "T", cnn_dbbancos)
                If Len(Trim(gassto)) = 0 Then
                    csql = "SELECT TOP 1 Val(CODIGO) AS num From BF9GIN ORDER BY Val(CODIGO) DESC"
                    If rsgasto.State = 1 Then rsgasto.Close
                    rsgasto.Open csql, cnn_dbbancos, 3, 1
                    If rsgasto.RecordCount > 0 Then
                        gassto = Format(CStr(rsgasto.Fields("num").value + 1), "000")
                    End If
                    amovs(0).campo = "CODIGO": amovs(0).valor = gassto: amovs(0).Tipo = "T"
                    amovs(1).campo = "BASE": amovs(1).valor = "G": amovs(1).Tipo = "T"
                    amovs(2).campo = "NOMBRE": amovs(2).valor = wnomctacont: amovs(2).Tipo = "T"
                    amovs(3).campo = "CUENTA": amovs(3).valor = wctacont: amovs(3).Tipo = "T"
                    GRABA_REGISTRO amovs(), "BF9GIN", "A", 3, StrConexDbBancos, ""
                End If
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = gassto
                dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctacont
                dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomctacont
                dxDBGrid1.Dataset.Post
            End If
    End If
    If dxDBGrid1.Columns.FocusedIndex = 1 Then
'        If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value <> 0 Then
'            dxDBGrid1.Dataset.Edit
'            dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value = ""
'            dxDBGrid1.Columns.FocusedIndex = 1
'        End If
        wdestino = "E"
        wgastos = ""
        Sw_AyuCodProv = False
        ayuda_gastos.TipoConcepto = "E"
        ayuda_gastos.Show 1
        dxDBGrid1.Dataset.Edit
        If Len(Trim(wgastos)) > 0 Then
            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = wgastos
            dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctacont
            If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value & "")) = 0 Then
                dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomgasto
            End If
        End If
        dxDBGrid1.Dataset.Post
        PROCESO_CUENTA
        
        If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
            dxDBGrid1.Columns.FocusedIndex = 3
        End If
      
    End If
    
    If dxDBGrid1.Columns.FocusedIndex = 2 Then
        PROCESO_CUENTA
    End If

End Sub

Public Sub PROCESO_CUENTA()
Dim chequeo As ADODB.Recordset

''    Set chequeo = New ADODB.Recordset
''    sqlcad = "Select f5cc from cf5pla where f5codcta='" & gcodcon & "'"
''    If chequeo.State = adStateOpen Then chequeo.Close
''    chequeo.Open sqlcad, contawin, adOpenDynamic, adLockOptimistic
''    If Not chequeo.EOF Then
''        af5cc = "" & Trim(chequeo.Fields("f5cc"))
''    End If
''    af5cc = True
''    If af5cc = True Then
''        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False
''        dxDBGrid1.Columns.FocusedIndex = 3
''    Else
''        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = True
''        dxDBGrid1.Columns.FocusedIndex = 4
''    End If

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim chequeo1 As ADODB.Recordset

    Set rsif5pla = New ADODB.Recordset
    Set chequeo = New ADODB.Recordset
    Set chequeo1 = New ADODB.Recordset
    If sw_nuevo_item = False Then
    
'    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
    
        If dxDBGrid1.Columns.FocusedIndex = 1 Then
            sw_detalle = True
            SqlCad = "Select * from bf9gin where codigo='" & Me.dxDBGrid1.Columns(1).value & "' AND BASE='G'"
            If rsif5pla.State = adStateOpen Then rsif5pla.Close
            rsif5pla.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            dxDBGrid1.Dataset.Edit
            If Not rsif5pla.EOF Then
                dxDBGrid1.Columns(2).value = "" & Trim(rsif5pla.Fields("cuenta"))
                dxDBGrid1.Columns(4).value = "" & Trim(rsif5pla.Fields("nombre"))
                If txtocompra.Visible = False Then
                    dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = ""
                    'dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").Value = ""
                    dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").CheckColumn.ValueChecked
                Else
                    If wcodcosto = "998" Then
                        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = ""
                    Else
                        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = wcodcosto
                    End If
                End If
                dxDBGrid1.Dataset.Edit
                If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                    dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "H"
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f3DEBHAB").value = "D"
                End If
            Else
                If dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = "" Then
                    If wcodgasto <> "" Then
                        dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = wcodgasto
                        dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctagasto
                        If dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = "" Then
                            dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomgasto
                        End If
                    Else
                        'MsgBox "Debe Ingresar un Codigo de Gasto", vbInformation, "Atencion"
                        MsgBox "Debe Ingresar un Codigo de Gasto", vbCritical, wnomcia
                        dxDBGrid1.Columns.FocusedIndex = 0
                    End If
                Else
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = wcodgasto
                    dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = wctagasto
                    dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = wnomgasto
                    dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value = ""
                    dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = ""
                    dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").CheckColumn.ValueChecked
                    dxDBGrid1.Dataset.Edit
                    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                        dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "H"
                    Else
                        dxDBGrid1.Columns.ColumnByFieldName("f3DEBHAB").value = "D"
                    End If
                    'MsgBox "Codigo de Gasto no existe. Ingrese un nuevo Codigo de Gasto", vbInformation, "Atencion"
                    MsgBox "Codigo de Gasto no existe. Ingrese un nuevo Código de Gasto", vbCritical, wnomcia
                    dxDBGrid1.Columns.FocusedIndex = 0
                End If
            End If
'''''            rsif5pla.Close
'''''            gcodcon = dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").Value
'''''            sqlcad = "Select f5cc from cf5pla where f5codcta='" & gcodcon & "'"
'''''            If chequeo.State = adStateOpen Then chequeo.Close
'''''            chequeo.Open sqlcad, contawin, adOpenDynamic, adLockOptimistic
'''''            If Not chequeo.EOF Then
'''''                af5cc = "" & Trim(chequeo.Fields("f5cc"))
'''''            End If
'''''            chequeo.Close
'''''            af5cc = True
'''''            If af5cc = True Then
'''''                dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False
'''''                dxDBGrid1.Columns.FocusedIndex = 3
'''''            Else
''''''                If dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").Value <> "" Then
''''''                    dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = True
'''''                    dxDBGrid1.Columns.FocusedIndex = 4
''''''                End If
'''''            End If
        End If
'''''        If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
'''''            If dxDBGrid1.Columns.FocusedIndex = 3 Then
'''''                sqlcad = "Select * from centros where F3COSTO='" & dxDBGrid1.Columns(3).Value & "'"
'''''                If chequeo1.State = adStateOpen Then chequeo1.Close
'''''                chequeo1.Open sqlcad, cnn_Db, adOpenDynamic, adLockOptimistic
'''''                dxDBGrid1.Dataset.Edit
'''''                If Not chequeo1.EOF Then
'''''                    dxDBGrid1.Columns.FocusedIndex = 4
'''''                Else
''''''                    If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value = "" Then
''''''                        MsgBox "Debe ingresar un Centro de Costo", vbInformation, "Atencion"
''''''                        dxDBGrid1.Columns.FocusedIndex = 2
'''''''                    Else
''''''                        dxDBGrid1.Dataset.Edit
''''''                        dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").Value = ""
'''''
''''''                        MsgBox "Codigo de Centro no existe. Ingrese un nuevo Codigo de Centro", vbInformation, "Atencion"
'''''                        dxDBGrid1.Columns.FocusedIndex = 3
''''''                    End If
'''''                End If
'''''            End If
'''''        End If
        If dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").ColIndex _
        Or dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").ColIndex _
        Or dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3CANTIDAD").ColIndex _
        Or dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").ColIndex _
        Or dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").ColIndex _
        Or dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").ColIndex Then
            sw_nuevo_item = True
            dxDBGrid1.Dataset.Edit
            Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
            Case "F3IMPORTE"
                If Val(dxDBGrid1.Columns.ColumnByFieldName("F3CANTIDAD").value & "") = 0 Then
                    dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").value = 0
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").value = Val(dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value & "") / Val(dxDBGrid1.Columns.ColumnByFieldName("F3CANTIDAD").value & "")
                End If
                If dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = True Then
                    dxDBGrid1.Columns.ColumnByFieldName("afecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("inafecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                End If
        
            Case "F3PREUNI", "F3CANTIDAD"
                dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value = Val(dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").value & "") * Val(dxDBGrid1.Columns.ColumnByFieldName("F3CANTIDAD").value & "")
                If dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = True Then
                    dxDBGrid1.Columns.ColumnByFieldName("afecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("inafecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                End If
            Case "F4SERGUI"
                dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").value = Format(dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").value, "000")
            Case "F4NUMGUI"
                dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").value = Format(dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").value, "0000000")
            Case "F3AFECTO"
                    If dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = False Then
                        dxDBGrid1.Columns.ColumnByFieldName("inafecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                        dxDBGrid1.Columns.ColumnByFieldName("afecto").value = 0
                    Else
                        dxDBGrid1.Columns.ColumnByFieldName("afecto").value = dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value
                        dxDBGrid1.Columns.ColumnByFieldName("inafecto").value = 0
                    End If
                
            End Select
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = False
            
            CALCULANDO
            CALCULAR_TOTALES
        End If
        If dxDBGrid1.Columns.FocusedIndex = 7 Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Dataset.Post
            TxtIgv(0).SetFocus
        End If
    End If
 '   End If

End Sub

Private Sub Txtigv_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtIgv(0).Text = Format(TxtIgv(0), "###,##0.00")
        TxtOtrImp(0).SetFocus
    End If

End Sub

Private Sub TxtIgv_LostFocus(Index As Integer)
    
    CALCULAR_TOTALES

End Sub

Public Sub CALCULANDO()
Dim m           As Boolean
Dim i           As Integer
Dim wbasimp1    As Double
Dim wmonina2    As Double

    'Set Tabla1 = New ADODB.Recordset
    'Set tabla11 = New ADODB.Recordset
    'Set tbparametro11 = New ADODB.Recordset
    If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
             dxDBGrid1.Dataset.Post
    End If
    
    'wbasimp1 = 0: wmonina2 = 0
    'SqlCad = "SELECT SUM(F3IMPORTE) AS IMPORTE FROM TEMP_DET WHERE F3AFECTO=-1"
    ''If Tabla1.State = adStateOpen Then Tabla1.Close
    ''Tabla1.Open SqlCad, cnn_form, adOpenStatic, adLockOptimistic
    'Set Tabla1 = Af.OpenSQLForwardOnly(SqlCad, cconex_formp)
    'If Not Tabla1.EOF Then
    '    wbasimp1 = Val(Tabla1.Fields("IMPORTE") & "")
    'Else
    '    wbasimp1 = 0#
    'End If
    'Tabla1.Close
    'SqlCad = "SELECT SUM(F3IMPORTE) AS IMPORTE1 FROM TEMP_DET WHERE F3AFECTO=FALSE"
    ''If tabla11.State = adStateOpen Then tabla11.Close
    ''tabla11.Open SqlCad, cnn_form, adOpenStatic, adLockOptimistic
    'Set tabla11 = Af.OpenSQLForwardOnly(SqlCad, cconex_formp)
    'If Not tabla11.EOF Then
    '    wmonina2 = Val("" & tabla11.Fields("IMPORTE1"))
    'Else
    '    wmonina2 = 0#
    'End If
    'tabla11.Close
    'PnlBasImp(0).Caption = Format(wbasimp1, "###,##0.00")
    'PnlMonIna(0).Caption = Format(wmonina2, "###,##0.00")
    
    PnlBasImp(0).Caption = Format(dxDBGrid1.Columns.ColumnByFieldName("afecto").SummaryFooterValue, "###,##0.00")
    PnlMonIna(0).Caption = Format(dxDBGrid1.Columns.ColumnByFieldName("inafecto").SummaryFooterValue, "###,##0.00")
    
    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
        wIMPUESTO = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * gretenc), "###,##0.00")
        'wotrimp = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * gfonavi), "###,##0.00")
    Else
        If Month(TxtFecha.value) < 3 And Year(TxtFecha.value) < 2012 Then wIgv = 19
        wIMPUESTO = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * wIgv / 100), "###,##0.00")
        wotrimp = Format(TxtOtrImp(0).Text, "###,##0.00")
        
    End If
    If Val(Format(PnlBasImp(0).Caption, "0.00")) > 0# Then
        TxtIgv(0).Text = Format(wIMPUESTO, "###,##0.00")
        TxtOtrImp(0).Text = Format(wotrimp, "###,##0.00")
    Else
        TxtIgv(0).Text = Format(wIMPUESTO, "###,##0.00")
    End If
    If Checkdatos.value = 1 Then
        txtredsuma.Text = Format(Val("" & txtredsuma.Text), "0.00")
        txtredresta.Text = Format(Val("" & txtredresta.Text), "0.00")
        txtdcto.Text = Format(Val("" & txtdcto.Text), "0.00")
    Else
        txtredsuma.Text = Format(0, "0.00")
        txtredresta.Text = Format(0, "0.00")
        txtdcto.Text = Format(0, "0.00")
    End If
  
End Sub

Public Sub CALCULAR_TOTALES()

    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "HON" Then
        PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) - Val(Format(TxtOtrImp(0).Text, "0.00")) - Val(Format(TxtIgv(0).Text, "0.00")), "###,##0.00")
    Else
        PnlTotal(0).Caption = Format(Val(Format(PnlBasImp(0).Caption, "0.00")) + Val(Format(PnlMonIna(0).Caption, "0.00")) + Val(Format(TxtOtrImp(0).Text, "0.00")) + Val(Format(TxtIgv(0).Text, "0.00")) + Val(Format(txtredsuma.Text, "0.00")) - Val(Format(txtredresta.Text, "0.00")) - Val(Format(txtdcto.Text, "0.00")), "###,##0.00")
    End If

End Sub

Private Sub Checkdatos_Click()
    
    If Checkdatos.value = 1 Then
        txtredresta.Visible = True
        txtredsuma.Visible = True
        txtdcto.Visible = True
        Label19.Visible = True
        Label20.Visible = True
        Label21.Visible = True
        txtredsuma.SetFocus
    Else
        txtdcto.Visible = False
        txtredresta.Visible = False
        txtredsuma.Visible = False
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        CALCULAR_TOTALES
    End If

End Sub

Private Sub txtmesmov_DblClick()
    
    TxtMesMov_KeyDown 113, 0

End Sub

Private Sub TxtMesMov_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Meses.Show 1
        meslet = Choose(Val(wmes), "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")
        Me.Caption = "Registro de Compras" & " - " & meslet & " - " & wanno
        txtmesmov.Text = wanno & Format(Val(wmes), "00")
        txtmes.Text = meslet
        If Month(Format(Date, "dd/mm/yyyy")) > Val(wmes) Then
            TxtFecha.value = CDate(Format("01/" & wmes & "/" & Year(Date), "dd/mm/yyyy"))
        Else
            TxtFecha.value = Format(Date, "dd/mm/yyyy")
        End If
    End If

End Sub

Private Sub TxtMesMov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtCodPrv.SetFocus
    End If

End Sub

Function valida() As Boolean

Dim SW_VALIDA As Boolean
Dim rscta As New ADODB.Recordset

SW_VALIDA = False
Set tbregisdoc = New ADODB.Recordset
ncorrela = 0
    SqlCad = "Select * from regisdoc where f4nummov='" & TxtNumMov.Text & "' AND F4MESMOV='" & txtmesmov.Text & "'"
    If tbregisdoc.State = adStateOpen Then tbregisdoc.Close
    tbregisdoc.Open SqlCad, StrConexDbBancos, 3, 1
    If Not tbregisdoc.EOF Then
        SqlCad = "Select * from pag_dcto where correla=" & tbregisdoc.Fields("F4CORRELA")
        If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
        TbDocumDet.Open SqlCad, StrConexDbBancos, 3, 1
        If Not tbregisdoc.EOF Then
        tbregisdoc.MoveFirst
        ncorrela = Val("" & TbDocumDet.Fields("CORRELA"))
        End If
        TbDocumDet.Close
    End If
    tbregisdoc.Close

If ncorrela = 0 Then
'    MsgBox "No hay documentos por validar", vbInformation, "Sistema Logistica"
    Exit Function
End If

    csql = "SELECT PAG_DCTO.CORRELA, PAG_DCTO.nro_comp, PAG_DCTO.fch_comp, PAG_DCTO.fch_vcto, PAG_DCTO.total, PAG_DCTO.F4MONTO1, PAG_DCTO.PROVEEDOR, PAG_DCTO.NOMPROV, First(REGISMOV.F3GASTO) AS CODGASTO, " & _
        "PAG_DCTO.moneda, IIf(PAG_DCTO.MONEDA = 'S', 'MN', 'US') AS SIMBOLO, PAG_DCTO.REFERENCIA, IIf(IsNull(PAG_DCTO.F4CENTRO) Or Len(PAG_DCTO.F4CENTRO)=0,REGISMOV.F3CENCOS,PAG_DCTO.F4CENTRO) AS CENTRO, PAG_DCTO.F4FECHA1, " & _
        "PAG_DCTO.SALDO, EF2PROVEEDORES.F2NOMPROV " & _
        "FROM ((PAG_DCTO LEFT JOIN EF2PROVEEDORES ON PAG_DCTO.proveedor = EF2PROVEEDORES.F2CODPROV) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA) LEFT JOIN REGISMOV ON (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) AND (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) " & _
        "GROUP BY PAG_DCTO.CORRELA, PAG_DCTO.nro_comp, PAG_DCTO.fch_comp, PAG_DCTO.fch_vcto, PAG_DCTO.total, PAG_DCTO.F4MONTO1, PAG_DCTO.PROVEEDOR, PAG_DCTO.NOMPROV, " & _
        "PAG_DCTO.moneda, IIf(PAG_DCTO.MONEDA = 'S', 'MN', 'US'), PAG_DCTO.REFERENCIA, IIf(IsNull(PAG_DCTO.F4CENTRO) Or Len(PAG_DCTO.F4CENTRO)=0,REGISMOV.F3CENCOS,PAG_DCTO.F4CENTRO), PAG_DCTO.F4FECHA1, " & _
        "PAG_DCTO.SALDO, EF2PROVEEDORES.F2NOMPROV,PAG_DCTO.deb_hab " & _
        "HAVING PAG_DCTO.deb_hab='H' and PAG_DCTO.correla=" & ncorrela & " and PAG_DCTO.correla not in (select corr_comp from provi_mvto where tipo_doc='P')"
   
If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
TbDocumDet.Open csql, StrConexDbBancos, adOpenDynamic, adLockOptimistic
If Not TbDocumDet.EOF Then
    TbDocumDet.MoveFirst
    Do While Not TbDocumDet.EOF
        cgruporeal = ObtenerCampo("BF9GIN", "GRUPCOMP", "CODIGO", "" & TbDocumDet.Fields("CODGASTO"), "T", cnn_dbbancos)
        
        csql = "SELECT PROVISIONALES.*, BF9GIN.GRUPCOMP AS GRUPO,BF9GIN.PRORRATEO FROM PROVISIONALES,BF9GIN WHERE (BF9GIN.CODIGO=PROVISIONALES.CONCEPTO AND BF9GIN.BASE='G') AND (ISNULL(BF9GIN.PRORRATEO) OR BF9GIN.PRORRATEO='') " & _
        "AND (ISNULL(PROVISIONALES.REFERENCIA) OR PROVISIONALES.REFERENCIA='') AND PROVISIONALES.SALDO>0 AND PROVISIONALES.DEB_HAB='H' AND PROVISIONALES.EST_ANUL = 'N' AND PROVISIONALES.PRESUPUESTO = 'N' " & _
        "AND PROVISIONALES.CCOSTO='" & "" & TbDocumDet.Fields("CENTRO") & "' AND (PROVISIONALES.CONCEPTO='" & "" & TbDocumDet.Fields("CODGASTO") & "' OR BF9GIN.GRUPCOMP='" & cgruporeal & "') ORDER BY CCOSTO"
    
        If rscta.State = adStateOpen Then rscta.Close
        rscta.Open csql, StrConexDbBancos, adOpenStatic, adLockOptimistic

        If Not rscta.EOF Then
        SW_VALIDA = True
        Exit Do
        End If
        rscta.Close

    TbDocumDet.MoveNext
    Loop
End If
TbDocumDet.Close

valida = SW_VALIDA
End Function
Private Sub grabar()
    
    If Len(Trim(TxtNumDoc.Text)) = 0 Then
        'MsgBox "Debe especificar el Número de Documento", 48, "Atención"
        MsgBox "El Nº del documento no puede estar en blanco", vbOKOnly, wnomcia
        TxtNumDoc.SetFocus
        Exit Sub
    End If
    If wanno & wmes < Year(TxtFecha.value) & Format(Mid(TxtFecha.value, 4, 2), "00") Then
        swtc = True
        'MsgBox "La fecha del documento no puede ser mayor al mes de proceso. ", 48, "Atención"
        MsgBox "La fecha del documento no puede ser mayor al mes de proceso. ", vbOKOnly, wnomcia
        TxtFecha.SetFocus
        Exit Sub
    End If
    
    If Val(PnlTotal(0).Caption) < 0 Then
        MsgBox "El documento debe ser por un importe positivo.", vbOKOnly + vbExclamation, wnomcia
        dxDBGrid1.SetFocus
        Exit Sub
    End If
    
    
    GRABA_ING_PROVEEDOR
    
'    If valida = True Then
'        If MsgBox("¿Desea validar los provisionales del Registro de Compra?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
'            valida_provi.Show 1
'        End If
'    End If
    
    
End Sub

Private Sub GRABA_ING_PROVEEDOR()
Dim cnumdoc     As String
Dim ccampo      As String
Dim rstraslada  As ADODB.Recordset
Dim ctipo       As String
Dim nitems      As Integer
Dim nfil        As Integer
Dim cvalores    As String
    
    Set RSDETALLE = New ADODB.Recordset
    Set TbDocumento1 = New ADODB.Recordset
    Set rstraslada = New ADODB.Recordset
    If sw_nuevo_doc = True Then
        cnumdoc = GENERA_NUMMOV(wanno & wmes)
        TxtNumMov.Text = cnumdoc
        ctipo = "A"
    Else
        cnumdoc = TxtNumMov.Text
        If UCase(right(Trim(CmbTipDoc.Text), 3)) <> "HON" Then
            If Val(txtredsuma.Text) <> 0 Or Val(txtredresta.Text) <> 0 Or Val(txtdcto.Text) <> 0 Then
                Checkdatos.value = 1
            Else
                Checkdatos.value = 0
            End If
        End If
        ctipo = "M"
    End If
     
    '------------------------- ASIGNA DATOS DE LA CABECERA
    amovs_cab(0).campo = "F4MESMOV": amovs_cab(0).valor = wanno & wmes: amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "F4NUMMOV": amovs_cab(1).valor = TxtNumMov.Text: amovs_cab(1).Tipo = "T"
    amovs_cab(2).campo = "F4CODPRV": amovs_cab(2).valor = TxtCodPrv.Text: amovs_cab(2).Tipo = "T"
    amovs_cab(3).campo = "F4NOMPRV": amovs_cab(3).valor = TxtNomPrv.Text: amovs_cab(3).Tipo = "T"
    amovs_cab(4).campo = "F4DIRPRV": amovs_cab(4).valor = TxtDirPrv.Text: amovs_cab(4).Tipo = "T"
    amovs_cab(5).campo = "F4RUCPRV": amovs_cab(5).valor = TxtRucPrv.Text: amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "F4GRUPO": amovs_cab(6).valor = txtcodcta.Text: amovs_cab(6).Tipo = "T"
    amovs_cab(7).campo = "F4CTACONT": amovs_cab(7).valor = TxtTelPrv.Text: amovs_cab(7).Tipo = "T"
    If Mon.Caption = "MN" Then
        amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = "S": amovs_cab(8).Tipo = "T"
    Else
        amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = "D": amovs_cab(8).Tipo = "T"
    End If
'    If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
'        amovs_cab(9).Campo = "F4BASIMP": amovs_cab(9).valor = Format(Val(PnlBasImp(0).Caption * -1), "0.00"): amovs_cab(9).TIPO = "N"
'        amovs_cab(10).Campo = "F4MONINA": amovs_cab(10).valor = Format(Val(PnlMonIna(0).Caption * -1), "0.00"): amovs_cab(10).TIPO = "N"
'        amovs_cab(11).Campo = "F4IGV": amovs_cab(11).valor = Format(Val(TxtIgv(0).Text * -1), "0.00"): amovs_cab(11).TIPO = "N"
'        amovs_cab(12).Campo = "F4OTRIMP": amovs_cab(12).valor = Format(Val(TxtOtrImp(0).Text * -1), "0.00"): amovs_cab(12).TIPO = "N"
'        amovs_cab(13).Campo = "F4TOTAL": amovs_cab(13).valor = Format(Val(PnlTotal(0).Caption * -1), "0.00"): amovs_cab(13).TIPO = "N"
'        amovs_cab(14).Campo = "F4REDSUMA": amovs_cab(14).valor = Format(Val(txtredsuma.Text * -1), "0.00"): amovs_cab(14).TIPO = "N"
'        amovs_cab(15).Campo = "F4REDRESTA": amovs_cab(15).valor = Format(Val(txtredresta.Text * -1), "0.00"): amovs_cab(15).TIPO = "N"
'        amovs_cab(16).Campo = "F4DCTO": amovs_cab(16).valor = Format(Val(txtdcto.Text * -1), "0.00"): amovs_cab(16).TIPO = "N"
'    Else
        amovs_cab(9).campo = "F4BASIMP": amovs_cab(9).valor = Format(PnlBasImp(0).Caption, "0.00"): amovs_cab(9).Tipo = "N"
        amovs_cab(10).campo = "F4MONINA": amovs_cab(10).valor = Format(PnlMonIna(0).Caption, "0.00"): amovs_cab(10).Tipo = "N"
        If UCase(right(CmbTipDoc.Text, 3)) = "HON" Then
            amovs_cab(11).campo = "f4montoret": amovs_cab(11).valor = Format(TxtIgv(0).Text, "0.00"): amovs_cab(11).Tipo = "N"
            amovs_cab(50).campo = "F4IGV": amovs_cab(50).valor = 0: amovs_cab(50).Tipo = "N"
        Else
            amovs_cab(11).campo = "F4IGV": amovs_cab(11).valor = Format(TxtIgv(0).Text, "0.00"): amovs_cab(11).Tipo = "N"
            amovs_cab(50).campo = "f4montoret": amovs_cab(50).valor = 0: amovs_cab(50).Tipo = "N"
        End If
        amovs_cab(12).campo = "F4OTRIMP": amovs_cab(12).valor = Format(TxtOtrImp(0).Text, "0.00"): amovs_cab(12).Tipo = "N"
        amovs_cab(13).campo = "F4TOTAL": amovs_cab(13).valor = Format(PnlTotal(0).Caption, "0.00"): amovs_cab(13).Tipo = "N"
        amovs_cab(14).campo = "F4REDSUMA": amovs_cab(14).valor = Format(txtredsuma.Text, "0.00"): amovs_cab(14).Tipo = "N"
        amovs_cab(15).campo = "F4REDRESTA": amovs_cab(15).valor = Format(txtredresta.Text, "0.00"): amovs_cab(15).Tipo = "N"
        amovs_cab(16).campo = "F4DCTO": amovs_cab(16).valor = Format(txtdcto.Text, "0.00"): amovs_cab(16).Tipo = "N"
'    End If
    amovs_cab(17).campo = "F4SERDOC": amovs_cab(17).valor = TxtSerDoc.Text: amovs_cab(17).Tipo = "T"
    amovs_cab(18).campo = "F4NUMDOC": amovs_cab(18).valor = TxtNumDoc.Text: amovs_cab(18).Tipo = "T"
    amovs_cab(19).campo = "F4FECHA": amovs_cab(19).valor = TxtFecha.value: amovs_cab(19).Tipo = "F"
    amovs_cab(20).campo = "F4TIPCAM": amovs_cab(20).valor = TxtTipCam.Text: amovs_cab(20).Tipo = "N"
    amovs_cab(21).campo = "F4FECVEN": amovs_cab(21).valor = TxtFecVen.value: amovs_cab(21).Tipo = "F"
    amovs_cab(22).campo = "F4FECHAREC": amovs_cab(22).valor = TxtFechaRec.value: amovs_cab(22).Tipo = "F"
    amovs_cab(23).campo = "F4REFERE": amovs_cab(23).valor = TxtRefere.Text: amovs_cab(23).Tipo = "T"
    amovs_cab(24).campo = "F4TIPDOC": amovs_cab(24).valor = left(Trim(right(CmbTipDoc.Text, 5)), 2): amovs_cab(24).Tipo = "T"
    amovs_cab(25).campo = "F4FORPAG": amovs_cab(25).valor = left(right(cmbfpagos.Text, 4), 3): amovs_cab(25).Tipo = "T"
    amovs_cab(26).campo = "F4CODIGV": amovs_cab(26).valor = right(cmbigv.Text, 3): amovs_cab(26).Tipo = "T"
'''    If UCase(Right(Trim(CmbTipDoc.Text), 3)) <> "CRE" Then
'''        If OpcSol.Value = True Then
'''            amovs_cab(27).campo = "F4BASIMP": amovs_cab(27).valor = Format(PnlBasImp(0).Caption, "0.00"): amovs_cab(27).TIPO = "N"
'''            amovs_cab(28).campo = "F4MONINA": amovs_cab(28).valor = Format(PnlMonIna(0).Caption, "0.00"): amovs_cab(28).TIPO = "N"
'''            amovs_cab(29).campo = "F4IGV": amovs_cab(29).valor = Format(TxtIgv(0).Text, "0.00"): amovs_cab(29).TIPO = "N"
'''            amovs_cab(30).campo = "F4OTRIMP": amovs_cab(30).valor = Format(TxtOtrImp(0).Text, "0.00"): amovs_cab(30).TIPO = "N"
'''            amovs_cab(31).campo = "F4REDSUMA": amovs_cab(31).valor = Format(txtredsuma.Text, "0.00"): amovs_cab(31).TIPO = "N"
'''            amovs_cab(32).campo = "F4REDRESTA": amovs_cab(32).valor = Format(txtredresta.Text, "0.00"): amovs_cab(32).TIPO = "N"
'''            amovs_cab(33).campo = "F4DCTO": amovs_cab(33).valor = Format(txtdcto.Text, "0.00"): amovs_cab(33).TIPO = "N"
'''            amovs_cab(34).campo = "F4TOTAL": amovs_cab(34).valor = Format(PnlTotal(0).Caption, "0.00"): amovs_cab(34).TIPO = "N"
'''        Else
'''            amovs_cab(27).campo = "F4BASIMP": amovs_cab(27).valor = Format(Val(Format(PnlBasImp(0).Caption, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(27).TIPO = "N"
'''            amovs_cab(28).campo = "F4MONINA": amovs_cab(28).valor = Format(Val(Format(PnlMonIna(0).Caption, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(28).TIPO = "N"
'''            amovs_cab(29).campo = "F4IGV": amovs_cab(29).valor = Format(Val(Format(TxtIgv(0).Text, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(29).TIPO = "N"
'''            amovs_cab(30).campo = "F4OTRIMP": amovs_cab(30).valor = Format(Val(Format(TxtOtrImp(0).Text, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(30).TIPO = "N"
'''            amovs_cab(31).campo = "F4REDSUMA": amovs_cab(31).valor = Format(Val(Format(txtredsuma.Text, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(31).TIPO = "N"
'''            amovs_cab(32).campo = "F4REDRESTA": amovs_cab(32).valor = Format(Val(Format(txtredresta.Text, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(32).TIPO = "N"
'''            amovs_cab(33).campo = "F4DCTO": amovs_cab(33).valor = Format(Val(Format(txtdcto.Text, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(33).TIPO = "N"
'''            amovs_cab(34).campo = "F4TOTAL": amovs_cab(34).valor = Format(Val(Format(PnlTotal(0).Caption, "0.00") * TxtTipCam.Text), "0.00"): amovs_cab(34).TIPO = "N"
'''        End If
'''    Else
'''        If OpcSol.Value = True Then
'''            amovs_cab(27).campo = "F4BASIMP": amovs_cab(27).valor = Format(Val(PnlBasImp(0).Caption * -1), "0.00"): amovs_cab(27).TIPO = "N"
'''            amovs_cab(28).campo = "F4MONINA": amovs_cab(28).valor = Format(Val(PnlMonIna(0).Caption * -1), "0.00"): amovs_cab(28).TIPO = "N"
'''            amovs_cab(29).campo = "F4IGV": amovs_cab(29).valor = Format(Val(TxtIgv(0).Text * -1), "0.00"): amovs_cab(29).TIPO = "N"
'''            amovs_cab(30).campo = "F4OTRIMP": amovs_cab(30).valor = Format(Val(TxtOtrImp(0).Text * -1), "0.00"): amovs_cab(30).TIPO = "N"
'''            amovs_cab(31).campo = "F4REDSUMA": amovs_cab(31).valor = Format(Val(txtredsuma.Text * -1), "0.00"): amovs_cab(31).TIPO = "N"
'''            amovs_cab(32).campo = "F4REDRESTA": amovs_cab(32).valor = Format(Val(txtredresta.Text * -1), "0.00"): amovs_cab(32).TIPO = "N"
'''            amovs_cab(33).campo = "F4DCTO": amovs_cab(33).valor = Format(Val(txtdcto.Text * -1), "0.00"): amovs_cab(33).TIPO = "N"
'''            amovs_cab(34).campo = "F4TOTAL": amovs_cab(34).valor = Format(Val(PnlTotal(0).Caption * -1), "0.00"): amovs_cab(34).TIPO = "N"
'''        Else
'''            amovs_cab(27).campo = "F4BASIMP": amovs_cab(27).valor = Format(Val(Format(Val(Format(PnlBasImp(0).Caption, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(27).TIPO = "N"
'''            amovs_cab(28).campo = "F4MONINA": amovs_cab(28).valor = Format(Val(Format(Val(Format(PnlMonIna(0).Caption, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(28).TIPO = "N"
'''            amovs_cab(29).campo = "F4IGV": amovs_cab(29).valor = Format(Val(Format(Val(Format(TxtIgv(0).Text, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(29).TIPO = "N"
'''            amovs_cab(30).campo = "F4OTRIMP": amovs_cab(30).valor = Format(Val(Format(Val(Format(TxtOtrImp(0).Text, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(30).TIPO = "N"
'''            amovs_cab(31).campo = "F4REDSUMA": amovs_cab(31).valor = Format(Val(Format(Val(Format(txtredsuma.Text, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(31).TIPO = "N"
'''            amovs_cab(32).campo = "F4REDRESTA": amovs_cab(32).valor = Format(Val(Format(Val(Format(txtredresta.Text, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(32).TIPO = "N"
'''            amovs_cab(33).campo = "F4DCTO": amovs_cab(33).valor = Format(Val(Format(Val(Format(txtdcto.Text, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(33).TIPO = "N"
'''            amovs_cab(34).campo = "F4TOTAL": amovs_cab(34).valor = Format(Val(Format(Val(Format(PnlTotal(0).Caption, "0.00") * TxtTipCam.Text), "0.00") * -1), "0.00"): amovs_cab(34).TIPO = "N"
'''        End If
'''    End If
    amovs_cab(27).campo = "F4FECHING": amovs_cab(27).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(27).Tipo = "T"
    amovs_cab(28).campo = "F4USUARIOING": amovs_cab(28).valor = wusuario: amovs_cab(28).Tipo = "T"
    amovs_cab(29).campo = "F4HORAING": amovs_cab(29).valor = Time: amovs_cab(29).Tipo = "T"
    amovs_cab(30).campo = "F4FECHMOD": amovs_cab(30).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(30).Tipo = "T"
    amovs_cab(31).campo = "F4USUARIOMOD": amovs_cab(31).valor = wusuario: amovs_cab(31).Tipo = "T"
    amovs_cab(32).campo = "F4HORAMOD": amovs_cab(32).valor = Time: amovs_cab(32).Tipo = "T"
    amovs_cab(33).campo = "F4FECHAIMP": amovs_cab(33).valor = Format(Date, "dd/mm/yyyy"): amovs_cab(33).Tipo = "T"
    amovs_cab(34).campo = "F4USUARIOIMP": amovs_cab(34).valor = wusuario: amovs_cab(34).Tipo = "T"
    amovs_cab(35).campo = "F4HORAIMP": amovs_cab(35).valor = Time: amovs_cab(35).Tipo = "T"
    amovs_cab(36).campo = "F4POLIZA": amovs_cab(36).valor = "" & TxtPoliza.Text: amovs_cab(36).Tipo = "T"
    amovs_cab(37).campo = "F4IMPORTACION": amovs_cab(37).valor = "" & txtimporta.Text: amovs_cab(37).Tipo = "T"
    amovs_cab(38).campo = "F4OCOMPRA": amovs_cab(38).valor = "" & txtocompra.Text: amovs_cab(38).Tipo = "T"
    'If intOrdenes > 1 Then
    '     amovs_cab(38).valor = "+ de 1"
    'Else
    '     amovs_cab(38).valor = "" & strOrdenes
    'End If
    amovs_cab(39).campo = "F4OBRA": amovs_cab(39).valor = "" & txtcentro.Text: amovs_cab(39).Tipo = "T"
    amovs_cab(40).campo = "intcodcategoria": amovs_cab(40).valor = Val(right(CboCategoria.Text, 9)): amovs_cab(40).Tipo = "N"
    amovs_cab(41).campo = "F4SERGUI": amovs_cab(41).valor = "" & txtSerGuia.Text: amovs_cab(41).Tipo = "T"
    amovs_cab(42).campo = "F4NUMGUI": amovs_cab(42).valor = "" & TxtNumGuia.Text: amovs_cab(42).Tipo = "T"
    amovs_cab(43).campo = "TIPODOCAUX": amovs_cab(43).valor = "06": amovs_cab(43).Tipo = "T"
    amovs_cab(44).campo = "TIPODOCREF": amovs_cab(44).valor = IIf(cmbTipdocref.Text = "Factura", "01", "03"): amovs_cab(44).Tipo = "T"
    amovs_cab(45).campo = "SERDOCREF": amovs_cab(45).valor = "" & txtSerGuia.Text: amovs_cab(45).Tipo = "T"
    amovs_cab(46).campo = "NUMDOCREF": amovs_cab(46).valor = "" & TxtNumGuia.Text: amovs_cab(46).Tipo = "T"
    amovs_cab(47).campo = "NUMDETRACCION": amovs_cab(47).valor = "" & txtdetraccion.Text: amovs_cab(47).Tipo = "T"
    amovs_cab(48).campo = "FECHADETRACCION": amovs_cab(48).valor = "" & txtfechadetraccion.value: amovs_cab(48).Tipo = "F"
    amovs_cab(49).campo = "PORC_DETR": amovs_cab(49).valor = "" & txtpordetra.Text: amovs_cab(49).Tipo = "T"
    
    
    '------------------------- ASIGNA DATOS AL DETALLE
    amovs_det(0).campo = "F3ITEM": amovs_det(0).valor = "": amovs_det(0).Tipo = "N"
    amovs_det(1).campo = "F4MESMOV": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
    amovs_det(2).campo = "F4NUMMOV": amovs_det(2).valor = "": amovs_det(2).Tipo = "T"
    amovs_det(3).campo = "F3GASTO": amovs_det(3).valor = "": amovs_det(3).Tipo = "T"
    amovs_det(4).campo = "F3CENCOS": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
    amovs_det(5).campo = "F3CTACON": amovs_det(5).valor = "": amovs_det(5).Tipo = "T"
    amovs_det(6).campo = "F3CONCEPTO": amovs_det(6).valor = "": amovs_det(6).Tipo = "T"
    amovs_det(7).campo = "F3IMPORTE": amovs_det(7).valor = "": amovs_det(7).Tipo = "N"
    amovs_det(8).campo = "F3DEBHAB": amovs_det(8).valor = "": amovs_det(8).Tipo = "T"
    amovs_det(9).campo = "F3AFECTO": amovs_det(9).valor = "": amovs_det(9).Tipo = "T"
    amovs_det(10).campo = "F3ORDEN": amovs_det(10).valor = "": amovs_det(10).Tipo = "T"
    amovs_det(11).campo = "f5codpro": amovs_det(11).valor = "": amovs_det(11).Tipo = "T"
    amovs_det(12).campo = "F3CANTIDAD": amovs_det(12).valor = "": amovs_det(12).Tipo = "N"
    amovs_det(13).campo = "F3PREUNI": amovs_det(13).valor = "": amovs_det(13).Tipo = "N"
    amovs_det(14).campo = "F3SERGUI": amovs_det(14).valor = "": amovs_det(14).Tipo = "T"
    amovs_det(15).campo = "F3NUMGUI": amovs_det(15).valor = "": amovs_det(15).Tipo = "T"
    '------------------- CALCULA NUMERO DE FILAS
    nitems = 0
    SqlCad = "SELECT COUNT(F3ITEM) AS NITEM FROM temp_det WHERE LEN(TRIM(F3ITEM)) > 0 "
    Set RSDETALLE = Af.OpenSQLForwardOnly(SqlCad, cconex_formp)
    If Not RSDETALLE.EOF Then
        nitems = Val("" & RSDETALLE.Fields("NITEM"))
    End If
    RSDETALLE.Close
    ReDim Values(20, nitems)
    Set RSDETALLE = Af.OpenSQLForwardOnly("SELECT * FROM temp_det", cconex_formp)
    If Not RSDETALLE.EOF Then
         nfil = 0
         RSDETALLE.MoveFirst
         Do While Not RSDETALLE.EOF
             If Len(Trim(RSDETALLE.Fields("F3ITEM") & "")) > 0 And Len(RSDETALLE.Fields("F3GASTO")) > 0 Then
                 'Values(0, nfil) = RSDETALLE.Fields("F3ITEM") & ""
                 Values(0, nfil) = nfil + 1
                 Values(1, nfil) = wanno & wmes
                 Values(2, nfil) = TxtNumMov.Text
                 Values(3, nfil) = RSDETALLE.Fields("F3GASTO") & ""
                 If txtcentro.Text = "998" Then
                    Values(4, nfil) = RSDETALLE.Fields("F3CENCOS") & ""
                 Else
                    Values(4, nfil) = txtcentro.Text
                 End If
                 Values(5, nfil) = RSDETALLE.Fields("F3CTACON") & ""
                 Values(6, nfil) = RSDETALLE.Fields("F3CONCEPTO") & ""
'                If PnlTotal(0).Caption > CDbl(txtmontoautorizado.Text) And CodigoOrdPago <> "" And valor = True Then
'                    MsgBox "La cantidad a provisionar es mayor que el pago autorizado, favor de corregir", vbInformation, "ATENCIÓN"
'                    TxtNumMov.Text = ""
'                    Exit Sub
'                 Else
                    Values(7, nfil) = RSDETALLE.Fields("F3IMPORTE") & ""
                 'End If
                 Values(8, nfil) = RSDETALLE.Fields("F3DEBHAB") & ""
                 Values(9, nfil) = IIf(RSDETALLE.Fields("F3AFECTO") = True, "*", "") & ""
                 Values(10, nfil) = RSDETALLE.Fields("F3ORDEN") & ""
                 Values(11, nfil) = RSDETALLE.Fields("F5codpro") & ""
                 Values(12, nfil) = RSDETALLE.Fields("F3CANTIDAD") & ""
                 Values(13, nfil) = RSDETALLE.Fields("F3PREUNI") & ""
                 If VerificaGuiaCabecera = True Then
                    Values(14, nfil) = txtSerGuia.Text & ""
                    Values(15, nfil) = TxtNumGuia.Text & ""
                 Else
                    Values(14, nfil) = RSDETALLE.Fields("F4SERGUI") & ""
                    Values(15, nfil) = RSDETALLE.Fields("F4NUMGUI") & ""
                 End If
                 nfil = nfil + 1
             End If
             RSDETALLE.MoveNext
         Loop
     End If
     RSDETALLE.Close
     cvalores = "1111111111111111"
     
     
     If ctipo = "A" Then     '--- Nuevo
         '------- GRABA CABECERA
         GRABA_REGISTRO amovs_cab(), "regisdoc", ctipo, 50, StrConexDbBancos, ""
             '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "regismov", ctipo, 15, StrConexDbBancos, "", Values(), nfil - 1, cvalores, "", ""
     Else    '--- Modificación
        '------- GRABA CABECERA
        GRABA_REGISTRO amovs_cab(), "regisdoc", ctipo, 50, StrConexDbBancos, "F4MESMOV = '" & wanno & Format(CboMeses.ListIndex + 1, "00") & "' AND F4NUMMOV='" & TxtNumMov.Text & "'"
        '------- GRABA DETALLE
        cnn_dbbancos.Execute ("DELETE * FROM regismov WHERE F4MESMOV = '" & wanno & Format(CboMeses.ListIndex + 1, "00") & "' AND F4NUMMOV = '" & TxtNumMov.Text & "'")
        GRABA_REGISTRO_DET amovs_det(), "regismov", "A", 15, StrConexDbBancos, "F4MESMOV  = '" & wanno & Format(CboMeses.ListIndex + 1, "00") & "' AND F4NUMMOV = '" & TxtNumMov.Text & "'", Values(), nfil - 1, cvalores, "", ""
        
        Genera_Igv_Flujo wanno & Format(CboMeses.ListIndex + 1, "00")
    End If
    
    
    If right(cmbfpagos.Text, 1) = "F" Then
        TRANS_CTASXPAGO
    End If
    Upd_Orden_Pagos "0"
    VerificaAtencionDeLaOrden txtocompra.Text, "0"
    sw_nuevo_doc = False

End Sub

Private Sub Genera_Igv_Flujo(pPeriodo As String)
Dim ncorrela As Double
Dim FechaIgv As Variant
Dim Rs As New ADODB.Recordset
Dim nTotalIgv As Double

FechaIgv = DateSerial(Val(left(pPeriodo, 4)), Val(right(pPeriodo, 2)) + 1, 0)
ncorrela = 0
csql = "SELECT Sum(IIf(REGISDOC.F4MONEDA='D',REGISDOC.F4IGV*REGISDOC.F4TIPCAM,REGISDOC.F4IGV)) AS TOTALIGV From REGISDOC WHERE (((REGISDOC.F4MESMOV)='" & pPeriodo & "'))"
Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
nTotalIgv = 0
If Rs.RecordCount > 0 Then
    nTotalIgv = Val(Rs!TotalIGV & "")
End If

csql = "SELECT top 1 cambio from cambios where month(fecha)=" & Val(right(pPeriodo, 2)) & " order by fecha desc"
Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
ntc = 2.75
If Rs.RecordCount > 0 Then
    ntc = Val(Rs!Cambio & "")
End If

csql = "select correla,nro_comp from pag_dcto where fch_comp=#" & Format(FechaIgv, "mm/dd/yyyy") & "# and UCASE(left(nro_comp,3))='TRB'"
Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)

If Rs.RecordCount > 0 Then
    pcorrela = Val(Rs!Correla & "")
    cnro_comp = Rs!NRO_COMP & ""
Else
    ncorrela = 0
    csql = "select top 1 nro_comp from pag_dcto where UCASE(left(nro_comp,3))='TRB' order by nro_comp desc"
    Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
    If Rs.RecordCount > 0 Then
        cnro_comp = "Trb001/" & Format(Val(right(Rs!NRO_COMP & "", 7)) + 1, "0000000")
    Else
        cnro_comp = "Trb001/0000001"
    End If
End If
Cod_Prove = ObtenerCampo("ef2proveedores", "f2codprov", "f2newruc", "20131312955", "T", cnn_dbbancos)
GRABA_CTASXPAGAR (ncorrela), "S", ntc, nTotalIgv, (cnro_comp), Cod_Prove, CVDate(FechaIgv), "20131312955"
If Rs.State = 1 Then Rs.Close
Set Rs = Nothing
End Sub

Private Sub GRABA_CTASXPAGAR(pcorrela As Double, pmonecta As String, ptc As Double, pimporte As Double, pnrocomp As String, pcodigo As String, pfecha As Date, pruc As String)
Dim csql        As String
Dim rsctas      As New ADODB.Recordset
Dim nsaldonew   As Double
Dim cinsert     As String
Dim ncorrela    As Double
Dim nannorepo   As Integer
Dim nrorepo     As Long
Dim nimputado   As Double
Dim nimputaso   As Double
Dim swUpdate    As Boolean
Dim ArrReg(0 To 20) As a_grabacion
    swUpdate = True
    csql = "SELECT SALDO,MONEDA,REG_COM FROM PAG_DCTO WHERE CORRELA=" & pcorrela & ""
    Set rsctas = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
    If rsctas.RecordCount > 0 Then
        ctipo = "M"
        If Val(rsctas.Fields("SALDO") & "") = 0 Then
            swUpdate = False
        End If
    Else
        ctipo = "A"
        pcorrela = GeneraCorrelaPag_DCto
    End If
    
        If swUpdate = True Then
            nannorepo = Val(right(Format(Year(pfecha), "0000"), 2))
            nrorepo = Val(Format(Day(pfecha), "00") & Format(Month(pfecha), "00"))
                ArrReg(0).campo = "VIA_INGR": ArrReg(0).valor = "2": ArrReg(0).Tipo = "T"
                ArrReg(1).campo = "CORRELA": ArrReg(1).valor = pcorrela: ArrReg(1).Tipo = "N"
                ArrReg(2).campo = "NRO_COMP": ArrReg(2).valor = pnrocomp: ArrReg(2).Tipo = "T"
                ArrReg(3).campo = "FCH_COMP": ArrReg(3).valor = pfecha: ArrReg(3).Tipo = "F"
                ArrReg(4).campo = "PROVEEDORO": ArrReg(4).valor = pcodigo: ArrReg(4).Tipo = "T"
                ArrReg(5).campo = "MONEDAO": ArrReg(5).valor = pmonecta: ArrReg(5).Tipo = "T"
                ArrReg(6).campo = "TOTALO": ArrReg(6).valor = pimporte: ArrReg(6).Tipo = "N"
                ArrReg(7).campo = "TCAMBIOO": ArrReg(7).valor = ptc: ArrReg(7).Tipo = "N"
                ArrReg(8).campo = "PROVEEDOR": ArrReg(8).valor = pcodigo: ArrReg(8).Tipo = "T"
                ArrReg(9).campo = "MONEDA": ArrReg(9).valor = pmonecta: ArrReg(9).Tipo = "T"
                ArrReg(10).campo = "TCAMBIO": ArrReg(10).valor = ptc: ArrReg(10).Tipo = "N"
                ArrReg(11).campo = "TOTAL": ArrReg(11).valor = pimporte: ArrReg(11).Tipo = "N"
                ArrReg(12).campo = "SALDO": ArrReg(12).valor = pimporte: ArrReg(12).Tipo = "N"
                ArrReg(13).campo = "DEB_HAB": ArrReg(13).valor = "H": ArrReg(13).Tipo = "T"
                ArrReg(14).campo = "ANO_REPO": ArrReg(14).valor = nannorepo: ArrReg(14).Tipo = "N"
                ArrReg(15).campo = "NRO_REPO": ArrReg(15).valor = nrorepo: ArrReg(15).Tipo = "N"
                ArrReg(16).campo = "FCH_REPO": ArrReg(16).valor = pfecha: ArrReg(16).Tipo = "F"
                ArrReg(17).campo = "RUC": ArrReg(17).valor = pruc: ArrReg(17).Tipo = "T"
                GRABA_REGISTRO ArrReg, "PAG_DCTO", ctipo, 17, StrConexDbBancos, "correla=" & pcorrela
        End If
    
    
            
            
    
    rsctas.Close: Set rsctas = Nothing
                    
End Sub

Private Sub EnviaDatos()

End Sub
Private Sub TRANS_CTASXPAGO()
    Dim StrDestino As String
    StrDestino = IIf(ObtenerCampo("documentos", "f2debhab", "f2abrev", (right(Trim(CmbTipDoc.Text), 3)), "T", cnn_dbbancos) = "D", "H", "D")
    Set tbregisdoc = New ADODB.Recordset
    SqlCad = "Select * from regisdoc where f4nummov='" & TxtNumMov.Text & "' AND F4MESMOV='" & wanno & Format(CboMeses.ListIndex + 1, "00") & "'"
    Set tbregisdoc = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
    If Not tbregisdoc.EOF Then
        'CorrelaPagDcto = tbregisdoc.Fields("F4CORRELA")
        TRANS_CTASXPAGAR_NEW "P", "1", tbregisdoc.Fields("F4CORRELA"), UCase(right(Trim(CmbTipDoc.Text), 3)), tbregisdoc.Fields("F4SERDOC"), tbregisdoc.Fields("F4NUMDOC"), tbregisdoc.Fields("F4FECHA"), tbregisdoc.Fields("F4RUCPRV"), tbregisdoc.Fields("F4CODPRV"), tbregisdoc.Fields("F4MONEDA"), tbregisdoc.Fields("F4TIPCAM"), tbregisdoc.Fields("F4TOTAL"), StrDestino, tbregisdoc.Fields("F4REFERE"), tbregisdoc.Fields("F4FECVEN"), tbregisdoc.Fields("F4OBRA") & "", tbregisdoc.Fields("F4NOMPRV"), StrConexDbBancos, tbregisdoc.Fields("F4MESMOV") & tbregisdoc.Fields("F4NUMMOV"), wanno, tbregisdoc.Fields("F4OCOMPRA")
    End If

End Sub

Private Function GENERA_NUMMOV(pmes As String)
Dim cnumdoc As String
Dim TbMes1 As New ADODB.Recordset
    
    
    If Len(TxtNumMov.Text) = 0 Then
        SqlCad = "Select top 1 F4NUMMOV from regisdoc where f4mesmov='" & pmes & "' order by F4NUMMOV desc"
        Set TbMes1 = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
        If TbMes1.RecordCount > 0 Then
            TbMes1.MoveFirst
            WNUMERO = "" & Val(TbMes1.Fields("F4NUMMOV")) + 1
        Else
            WNUMERO = "1"
        End If
        cnumdoc = Format(WNUMERO, "0000000")
        SqlCad = "UPDATE meses SET F2NUMMOV='" & cnumdoc & "' where F2NUMMES='" & right(pmes, 2) & "'"
        cnn_dbbancos.Execute (SqlCad)
        TbMes1.Close
        Set TbMes1 = Nothing
    Else
        WNUMERO = "" & Val(TxtNumMov.Text)
    End If

    
    GENERA_NUMMOV = cnumdoc

End Function

Private Sub elimina(pmes As String, pnumero As String)
On Error GoTo ERROR_ELIMINA
ReDim amovs(0 To 0) As a_grabacion
Dim cmes            As String
Dim cnumdoc         As String
Dim tbregismov      As New ADODB.Recordset

    Set TbMes1 = New ADODB.Recordset
    Set tbregismov = New ADODB.Recordset
    
    sw_elimina = False
    If Len(Trim("" & TxtNumMov.Text)) = 0 Then
        MsgBox "El Numero de Movimiento no ha sido grabado. Verifique", vbCritical, "Atención"
        Exit Sub
    End If
    
    cad$ = "Select * from REGISDOC WHERE F4MESMOV = '" & pmes & "' AND F4NUMMOV = '" & pnumero & "'"
    If tbregismov.State = adStateOpen Then tbregismov.Close
    tbregismov.Open cad$, cnn_dbbancos
    If Not tbregismov.EOF Then
        
        If MsgBox("Está seguro(a) de eliminar el Movimiento ?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
        '     --------- ELIMINAR en cta. x pagar -------
            If RSPAG_DCTO.State = adStateOpen Then RSPAG_DCTO.Close
            RSPAG_DCTO.Open "select * from pag_dcto where correla=" & tbregismov.Fields("f4correla"), StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not RSPAG_DCTO.EOF Then
                If RSPAG_DCTO.Fields("saldo") = RSPAG_DCTO.Fields("TOTAL") Then
                    cnn_dbbancos.Execute "DELETE * FROM PAG_DCTO WHERE CORRELA= " & tbregismov.Fields("f4correla")
                Else
                    MsgBox "Documento YA ha sido pagado,no puede ser eliminado...", 16, "Atención"
                    Exit Sub
                End If
            End If
            RSPAG_DCTO.Close
        
            cnn_dbbancos.Execute ("DELETE * FROM regisdoc  WHERE F4MESMOV = '" & pmes & "' AND F4NUMMOV = '" & pnumero & "'")
            '------------Buscando en Movimientos
            cnn_dbbancos.Execute ("DELETE * FROM regismov WHERE F4MESMOV = '" & pmes & "' AND F4NUMMOV = '" & pnumero & "'")
        End If
        
        SqlCad = "Select f4nummov from regisdoc where F4MESMOV='" & pmes & "' order by f4nummov desc"
        If TbMes1.State = adStateOpen Then TbMes1.Close
        TbMes1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        If Not TbMes1.EOF Then
        TbMes1.MoveFirst
'        sqlcad = "Select * from meses where F2NUMMES='" & pmes & "'"
'        If tbmes1.State = adStateOpen Then tbmes1.Close
'        tbmes1.Open sqlcad, cnn_Db, adOpenDynamic, adLockOptimistic
'        If Not tbmes1.EOF Then
            WNUMERO = "" & Val(TbMes1.Fields("F4NUMMOV"))
        End If
        TbMes1.Close
        
        cnumdoc = Format(WNUMERO, "0000000")
        SqlCad = "UPDATE meses SET F2NUMMOV='" & cnumdoc & "' where F2NUMMES='" & pmes & "'"
        cnn_dbbancos.Execute (SqlCad)
        sw_elimina = True
            
    End If
     Upd_Orden_Pagos "1"
    VerificaAtencionDeLaOrden txtocompra.Text, "1"
    Exit Sub

ERROR_ELIMINA:
    MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atención"
    Resume Next

End Sub

Public Sub CONSULTA()
    Dim i As Integer
    Set tbregisdoc = New ADODB.Recordset
    Set tbregismov = New ADODB.Recordset
    
    If Len(Trim(gnummov)) > 0 Then
        SqlCad = "Select * from regisdoc where f4nummov='" & gnummov & "' and f4mesmov='" & wmes & "'"
        If tbregisdoc.State = adStateOpen Then tbregisdoc.Close
        tbregisdoc.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        If Not tbregisdoc.EOF Then
            jmes = "" & tbregisdoc.Fields("f4mesmov")
            TxtNumMov.Text = "" & tbregisdoc.Fields("f4nummov")
            TxtCodPrv.Text = "" & tbregisdoc.Fields("f4codprv")
            TxtNomPrv.Text = "" & tbregisdoc.Fields("f4nomprv")
            TxtDirPrv.Text = "" & tbregisdoc.Fields("f4dirprv")
            TxtRucPrv.Text = "" & tbregisdoc.Fields("f4rucprv")
            TxtSerDoc.Text = "" & tbregisdoc.Fields("f4serdoc")
            TxtNumDoc.Text = "" & tbregisdoc.Fields("f4numdoc")
            TxtFecha.value = "" & tbregisdoc.Fields("f4fecha")
            jmoneda = "" & tbregisdoc.Fields("f4moneda")
            If jmoneda = "S" Then
                Mon.Caption = "MN"
            Else
                Mon.Caption = "US"
            End If
            TxtTipCam.Text = Format(Val("" & tbregisdoc.Fields("f4tipcam")), "0.000")
            TxtRefere.Text = "" & tbregisdoc.Fields("f4refere")
            '---------------------------Tipo de documento
            For i = 0 To CmbTipDoc.ListCount - 1
                If left(Trim(right(CmbTipDoc.List(i), 5)), 2) = tbregisdoc.Fields("f4tipdoc") Then
                    CmbTipDoc.ListIndex = i
                End If
            Next
            '--------------------------
            CmbTipDoc_LostFocus
            If UCase(right(Trim(CmbTipDoc.Text), 3)) = "CRE" Then
                PnlBasImp(0).Caption = Format(Val(tbregisdoc.Fields("f4basimp") * -1), "0.00")
                PnlMonIna(0).Caption = Format(Val(tbregisdoc.Fields("f4monina") * -1), "0.00")
                TxtIgv(0).Text = Format(Val(tbregisdoc.Fields("f4igv") * -1), "0.00")
                TxtOtrImp(0).Text = Format(Val(tbregisdoc.Fields("f4otrimp") * -1), "0.00")
                txtredsuma.Text = Format(Val(tbregisdoc.Fields("f4redsuma") * -1), "0.00")
                txtredresta.Text = Format(Val(tbregisdoc.Fields("f4redresta") * -1), "0.00")
                txtdcto.Text = Format(Val(tbregisdoc.Fields("f4dcto") * -1), "0.00")
                PnlTotal(0).Caption = Format(Val(tbregisdoc.Fields("f4total") * -1), "0.00")
            Else
                PnlBasImp(0).Caption = Format(Val("" & tbregisdoc.Fields("f4basimp")), "0.00")
                PnlMonIna(0).Caption = Format(Val("" & tbregisdoc.Fields("f4monina")), "0.00")
                TxtIgv(0).Text = Format(Val("" & tbregisdoc.Fields("f4igv")), "0.00")
                TxtOtrImp(0).Text = Format(Val("" & tbregisdoc.Fields("f4otrimp")), "0.00")
                txtredsuma.Text = txtredsuma.Text
                If Val(tbregisdoc.Fields("f4redsuma")) <> 0 Then
                    txtredsuma.Text = Format(Val(tbregisdoc.Fields("f4redsuma")), "0.00")
                End If
                If Val(tbregisdoc.Fields("f4redresta")) <> 0 Then
                    txtredresta.Text = Format(Val(tbregisdoc.Fields("f4redresta")), "0.00")
                End If
                If Val(tbregisdoc.Fields("f4dcto")) <> 0 Then
                    txtdcto.Text = Format(Val(tbregisdoc.Fields("f4dcto")), "0.00")
                End If
                PnlTotal(0).Caption = Format(Val("" & tbregisdoc.Fields("f4total")), "0.00")
            End If
            TxtFecVen.value = "" & tbregisdoc.Fields("f4fecven")
            txtcodcta.Text = "" & tbregisdoc.Fields("f4grupo")
            TxtTelPrv.Text = "" & tbregisdoc.Fields("f4ctacont")
            TxtFechaRec.value = "" & tbregisdoc.Fields("f4fecharec")
            If UCase(right(Trim(CmbTipDoc.Text), 3)) <> "HON" Then
                If txtredsuma.Text <> 0 Or txtredresta.Text <> 0 Or txtdcto.Text <> 0 Then
                    Checkdatos.value = 1
                Else
                    Checkdatos.value = 0
                End If
            End If
            '--------------------------- formas de pago
            For i = 0 To cmbfpagos.ListCount - 1
                If left(right(cmbfpagos.List(i), 4), 3) = tbregisdoc.Fields("f4forpag") Then
                    cmbfpagos.ListIndex = i
                End If
            Next
            '--------------------------Combo del IGV
            For i% = 0 To cmbigv.ListCount - 1
                If right(cmbigv.List(i%), 3) = tbregisdoc.Fields("f4codigv") Then
                    cmbigv.ListIndex = i%
                End If
            Next
            '--------------------------
            SqlCad = "Select * from regismov where f4mesmov='" & jmes & "' and f4nummov='" & TxtNumMov.Text & "'"
            If tbregismov.State = adStateOpen Then tbregismov.Close
            tbregismov.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
            If Not tbregismov.EOF Then
                Do While Not tbregismov.EOF
                    dxDBGrid1.Dataset.Append
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns(0).value = "" & tbregismov.Fields("f3item")
                    dxDBGrid1.Columns(1).value = "" & tbregismov.Fields("f3gasto")
                    dxDBGrid1.Columns(2).value = "" & tbregismov.Fields("f3ctacon")
                    dxDBGrid1.Columns(3).value = "" & tbregismov.Fields("f3cencos")
                    dxDBGrid1.Columns(4).value = "" & tbregismov.Fields("f3concepto")
                    dxDBGrid1.Columns(5).value = Val("" & tbregismov.Fields("f3importe"))
                    If tbregismov.Fields("f3afecto") = "*" Then
                        dxDBGrid1.Columns(6).value = True
                    Else
                        dxDBGrid1.Columns(6).value = False
                    End If
                    dxDBGrid1.Columns(7).value = "" & tbregismov.Fields("f3debhab")
                    tbregismov.MoveNext
               Loop
            End If
            tbregismov.Close
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Dataset.Post
        End If
        tbregisdoc.Close
    End If
    ctipo = "M"

End Sub

Public Sub PROCEDIMIENTO_NUEVO()
    
    Me.MousePointer = vbHourglass
    sw_nuevo_doc = False
    sw_detalle = False
    nuevo
    AdicionaItem
    AdicionaItem
    
    wcodcliprov = ""
    wRucCliProv = ""
    wnomcliprov = ""
    
    FrmName = Me.Name
    Ayuda_Proveedores.Show 1
    Unload Ayuda_Proveedores
    
    If Len(Trim(wcodcliprov)) > 0 Then
        TxtCodPrv.Text = wcodcliprov
        TxtRucPrv.Text = wRucCliProv
    End If
    
    sw_nuevo_doc = True
    Me.MousePointer = vbDefault
    'TxtCodPrv.SetFocus
    
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
PresionaBotonMoneda KeyCode
    If KeyCode = 115 Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado? ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                CALCULANDO
            End If
            sw_nuevo_item = False
        End If
    End If
    
    If KeyCode = 46 Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado? ", vbQuestion + vbYesNo, wnomcia) = vbYes Then
            sw_nuevo_item = True
            
                dxDBGrid1.Dataset.Delete
                If dxDBGrid1.Dataset.RecordCount = 0 Then
                    AdicionaItem
                End If
                CALCULANDO
                CALCULAR_TOTALES
            
            sw_nuevo_item = False
        End If
    End If
    If KeyCode = 13 Then
'        If dxDBGrid1.Columns.FocusedColumn.FieldName = "F3DEBHAB" Then
'            dxDBGrid1.Columns.FocusedIndex = 8
'        End If
        If dxDBGrid1.Columns.FocusedColumn.FieldName = "F3DEBHAB" Then
            TxtIgv(0).SetFocus
        End If
    End If
    
    
End Sub

Private Sub TxtNumDoc_LostFocus()
On Error Resume Next
Dim TbCabRegis1 As ADODB.Recordset

Set TbCabRegis1 = New ADODB.Recordset
If sw_nuevo_doc = True Then
    If Len(Trim(TxtNumDoc.Text)) > 0 Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000")
        SqlCad = "Select * from regisdoc where f4codprv='" & Trim(TxtCodPrv.Text) & "' and f4tipdoc='" & Format(left(right(Trim(CmbTipDoc.Text), 5), 2), "00") & "' and f4serdoc='" & Trim(TxtSerDoc.Text) & "' and f4numdoc='" & Trim(TxtNumDoc.Text) & "'"
        If TbCabRegis1.State = adStateOpen Then TbCabRegis1.Close
        TbCabRegis1.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        If Not TbCabRegis1.EOF Then
'            sw_nuevo_doc = False
            MsgBox "El documento  ya  ha  sido  ingresado en  el Mes " & TbCabRegis1.Fields("F4MESMOV") & "  y  Nº " & TbCabRegis1.Fields("F4NUMMOV"), 48
            TxtNumDoc.Text = ""
            TxtNumDoc.SetFocus
        Else
'            TxtPoliza.SetFocus
'            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        End If
        TbCabRegis1.Close
    Else
        'MsgBox "El Nº del documento no puede estar en blanco", 48
        If MsgBox("El Nº del documento no puede estar en blanco", vbExclamation + vbOKCancel, wnomcia) = vbOK Then
            TxtNumDoc.SetFocus
        End If
    End If
End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    

End Sub

Private Sub LLENA_TEMPCAB()
'On Error GoTo error_llena
Dim TEMP_CABE As ADODB.Recordset
    
    Set TEMP_CABE = New ADODB.Recordset
    SqlCad = "Select * from vou_cab"
    TEMP_CABE.Open SqlCad, cnn_form, adOpenDynamic, adLockOptimistic
    TEMP_CABE.AddNew
    TEMP_CABE.Fields("F4COMPRO") = right(txtmesmov.Text, 2) & "/" & TxtNumMov.Text
    TEMP_CABE.Fields("F4TOTDEB") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4TOTHAB") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4CODORI") = left(worigen, 2)
    TEMP_CABE.Fields("F4ORIGEN") = "Registro de Compras"
    TEMP_CABE.Fields("F4FECHA") = TxtFecha.value
    TEMP_CABE.Fields("F4TIPCAMBD") = TxtTipCam.Text
    TEMP_CABE.Fields("F4PROVE") = TxtNomPrv.Text
    TEMP_CABE.Fields("F4IGV") = TxtIgv(0).Text
    TEMP_CABE.Fields("F4TOTAL") = PnlTotal(0).Caption
    TEMP_CABE.Fields("F4EMPRESA") = wF1Dir
    TEMP_CABE.Fields("F4MONEDA") = IIf(Mon.Caption = "MN", "SOLES", "DOLARES")
    TEMP_CABE.Fields("f4detalle") = TxtRefere.Text & ""
    TEMP_CABE.Fields("f4titulo") = "IMPRESION DEL COMPROBANTE"
    TEMP_CABE.Update
    Exit Sub
    
error_llena:
    Resume Next
    
End Sub

Private Sub LLENA_TEMPDET()
On Error GoTo error_llena_t
Dim RSDETALLE   As ADODB.Recordset
Dim temp_deta As ADODB.Recordset
Dim Busca       As String
Dim conta       As Integer
Dim DA As Integer

    Set RSDETALLE = New ADODB.Recordset
    Set temp_deta = New ADODB.Recordset
    SqlCad = "Select * from temp_det"
    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
    RSDETALLE.Open SqlCad, cnn_form, adOpenDynamic, adLockOptimistic
    Busca = wmes & "/" & TxtNumMov
    conta = 1
    RSDETALLE.MoveFirst
    If temp_deta.State = adStateOpen Then temp_deta.Close
    temp_deta.Open "Select * from vou_det", cnn_form, adOpenStatic, adLockOptimistic
    DA = temp_deta.RecordCount
    Do While Not RSDETALLE.EOF
        temp_deta.AddNew
        temp_deta.Fields("F3COMPRO") = "" & Busca
        temp_deta.Fields("F3ELEMEN") = conta
        temp_deta.Fields("F3DETALL") = RSDETALLE.Fields("F3CONCEPTO")
        temp_deta.Fields("F3CODGAS") = RSDETALLE.Fields("F3GASTO")
        temp_deta.Fields("F5CODCTA") = RSDETALLE.Fields("F3CTACON")
        If RSDETALLE.Fields("f3debhab") = "D" Then
            temp_deta.Fields("F3DEBE") = Format(Val("" & RSDETALLE.Fields("f3importe")), "#0.00")
        Else
            temp_deta.Fields("F3HABER") = Format(Val("" & RSDETALLE.Fields("f3importe")), "#0.00")
        End If
        temp_deta.Fields("F3TIPCAMBD") = TxtTipCam.Text
        temp_deta.Fields("F3TIPDOC") = ""   ' TIPDOC
        temp_deta.Fields("F3COSTO") = RSDETALLE.Fields("F3CENCOS")
        temp_deta.Update
        RSDETALLE.MoveNext
        conta = conta + 1
        If RSDETALLE.EOF Then Exit Do
    Loop
    
    If UCase(right(CmbTipDoc.Text, 3)) <> "HON" Then
        If Val(Format(TxtIgv(0).Text, "#0.00")) <> 0 Then
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "I.G.V."
            temp_deta.Fields("F3CODGAS") = "IGV"
            temp_deta.Fields("F5CODCTA") = wctaigv
            If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3HABER") = Abs(Val(Format(TxtIgv(0).Text, "#0.00")))
            Else
                temp_deta.Fields("F3DEBE") = Abs(Val(Format(TxtIgv(0).Text, "#0.00")))
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
            If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
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
            If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
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
            If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
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
        If Val(Format(txtdcto.Text, "0.00")) <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "DESCUENTOS"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wdcto
            If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
                temp_deta.Fields("F3DEBE") = Val(Format(txtdcto.Text, "#0.00"))
            Else
                temp_deta.Fields("F3HABER") = Val(Format(txtdcto.Text, "#0.00"))
             End If
            temp_deta.Fields("F3TIPCAMBD") = Val(Format(TxtTipCam.Text, "#0.000"))
            temp_deta.Fields("F3TIPDOC") = ""
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Update
        End If
        '---------------------------------------------------------
    Else
        If Val(Format(TxtIgv(0).Text, "#0.00")) <> 0 Then
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = "" & Busca
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = "RETENCION"
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = wctaret
            temp_deta.Fields("F3HABER") = Val(Format(TxtIgv(0).Text, "#0.00"))
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
        
        If Len(Trim(TxtTelPrv.Text) & "") = 0 Then
            If Mon.Caption = "US" Then
                temp_deta.Fields("F5CODCTA") = wCtaProvDol
                temp_deta.Fields("F3CODGAS") = "PROD"
            Else
                temp_deta.Fields("F5CODCTA") = wCtaProvSol
                temp_deta.Fields("F3CODGAS") = "PRO"
            End If
        Else
            temp_deta.Fields("F5CODCTA") = Trim(TxtTelPrv.Text)
        End If
        If UCase(right(CmbTipDoc.Text, 3)) <> "CRE" Then
            temp_deta.Fields("F3HABER") = Val(Format(PnlTotal(0).Caption, "#0.00"))
        Else
            temp_deta.Fields("F3DEBE") = Val(Format(PnlTotal(0).Caption, "#0.00"))
        End If
        temp_deta.Fields("F3TIPCAMBD") = TxtTipCam.Text
        temp_deta.Fields("F3TIPDOC") = right(CmbTipDoc.Text, 3)
        temp_deta.Fields("F3AUXILIAR") = TxtRucPrv.Text
        temp_deta.Fields("F3DOCUM") = right(CmbTipDoc.Text, 3) & TxtSerDoc.Text & "/" & TxtNumDoc.Text
        temp_deta.Update
    End If

    Exit Sub

error_llena_t:
    MsgBox Err.Description
    Resume Next

End Sub
Private Sub actualiza()
Dim TbConsulta  As Recordset
Dim i           As Integer
Dim xtipopago   As String
Dim Msql        As String

    Me.MousePointer = vbHourglass
    
    txtmesmov.Text = Periodo
    TxtNumMov.Text = registro
    CboMeses.ListIndex = Val(wmes) - 1
    CboMeses.Enabled = False
    
    'bloquea si ya tiene vb
    csql = "select * from regisdoc where f4mesmov='" & txtmesmov.Text & "' and f4nummov='" & TxtNumMov.Text & "' and f4vb=-1"
    Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
    If Rs.RecordCount > 0 Then
        Toolbar.Buttons(3).Visible = False
        Toolbar.Buttons(4).Visible = False
    Else
        Toolbar.Buttons(3).Visible = True
        Toolbar.Buttons(4).Visible = True
    End If
    

    SqlCad = "Select * from regisdoc where F4MESMOV = '" + txtmesmov.Text + "' AND F4NUMMOV = '" + TxtNumMov.Text + "'"
    If TbCabRegis_new.State = 1 Then TbCabRegis_new.Close
    
    'TbCabRegis_new.Open sqlcad, cnn_Db, adOpenDynamic, adLockOptimistic
    TbCabRegis_new.Open SqlCad, StrConexDbBancos, 3, 1
    
    If Not TbCabRegis_new.EOF Then
        TxtCodPrv.Text = "" & TbCabRegis_new.Fields("F4CODPRV")
        TxtNomPrv.Text = "" & TbCabRegis_new.Fields("F4NOMPRV")
        TxtDirPrv.Text = "" & TbCabRegis_new.Fields("F4DIRPRV")
        TxtRucPrv.Text = "" & TbCabRegis_new.Fields("F4RUCPRV")
        TxtTelPrv.Text = "" & TbCabRegis_new.Fields("F4CTACONT")
        txtcodcta.Text = "" & TbCabRegis_new.Fields("F4GRUPO")
        TxtFecha.value = "" & Format(TbCabRegis_new.Fields("F4FECHA"), "DD/MM/YYYY")
        cmbTipdocref.ListIndex = IIf(TbCabRegis_new.Fields("TIPODOCREF") = "01", 0, 1)
        txtSerGuia.Text = "" & TbCabRegis_new.Fields("F4SERGUI")
        TxtNumGuia.Text = "" & TbCabRegis_new.Fields("F4NUMGUI")
        '--------------------------- formas de pago
        For i = 0 To cmbfpagos.ListCount - 1
            If left(right(cmbfpagos.List(i), 4), 3) = "" & TbCabRegis_new.Fields("f4forpag") Then
                cmbfpagos.ListIndex = i
                Exit For
            End If
        Next
        Call SeleccionaEnComboRight(Format("" & TbCabRegis_new!intCodCategoria, "00000000"), CboCategoria)
        If IsNull(TbCabRegis_new.Fields("F4FECVEN")) Then
            TxtFecVen.value = Date
        Else
            TxtFecVen.value = "" & Format(TbCabRegis_new.Fields("F4FECVEN"), "DD/MM/YYYY")
        End If
        If IsNull(TbCabRegis_new.Fields("F4FECHAREC")) Then
            TxtFechaRec.value = Date
        Else
            TxtFechaRec.value = "" & Format(TbCabRegis_new.Fields("F4FECHAREC"), "DD/MM/YYYY")
        End If
        'txtdias.Text = CVDate(TxtFecVen.Value) - CVDate(txtfecha.Value)
        'txttipdoc.Text = "" & TbCabRegis_new.Fields("f4tipdoc")
    
        If wocompra = "*" Then
            dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = True
            If wf1tipdoc_asoc = "V" Then
                lblocompra.Caption = "Vale Ing."
            Else
                lblocompra.Caption = "O. Compra"
            End If
        Else
            dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").Visible = False
        End If
    
        For i% = 0 To CmbTipDoc.ListCount - 1
            If left(Trim(right(CmbTipDoc.List(i%), 5)), 2) = "" & TbCabRegis_new.Fields("f4tipdoc") Then
                CmbTipDoc.ListIndex = i%
                Exit For
            End If
        Next
    
        For i% = 0 To cmbigv.ListCount - 1
            If right(cmbigv.List(i%), 3) = TbCabRegis_new.Fields("f4codigv") Then
                cmbigv.ListIndex = i%
                Exit For
            End If
        Next
    
        TxtSerDoc.Text = "" & TbCabRegis_new.Fields("F4SERDOC")
        TxtNumDoc.Text = "" & TbCabRegis_new.Fields("F4NUMDOC")
        'TxtSerDoc.Locked = True: txtnumdoc.Locked = True: CmbTipDoc.Locked = True
        
        TxtTipCam.Text = TbCabRegis_new.Fields("F4TIPCAM") & ""
        TxtRefere.Text = "" & TbCabRegis_new.Fields("F4REFERE")
        txtimporta.Text = "" & TbCabRegis_new.Fields("F4IMPORTACION")
        TxtPoliza.Text = "" & TbCabRegis_new.Fields("F4POLIZA")
        txtocompra.Text = "" & TbCabRegis_new.Fields("F4OCOMPRA")
        txtcentro.Text = "" & TbCabRegis_new.Fields("F4OBRA")
        txtdetraccion.Text = "" & TbCabRegis_new.Fields("NUMDETRACCION")
        txtfechadetraccion.value = "" & TbCabRegis_new.Fields("FECHADETRACCION")
        'txtpordetra.Value = "" & TbCabRegis_new.Fields("PORC_DETR")
    
        If UCase(right(CmbTipDoc.Text, 3)) = "CRE" Then
            PnlBasImp(0).Caption = Format(IIf(Val(TbCabRegis_new.Fields("F4BASIMP") & "") < 0, Val(TbCabRegis_new.Fields("F4BASIMP") & "") * -1, Val(TbCabRegis_new.Fields("F4BASIMP") & "")), "###,##0.00")
            PnlMonIna(0).Caption = Format(IIf(Val(TbCabRegis_new.Fields("F4MONINA") & "") < 0, Val(TbCabRegis_new.Fields("F4MONINA") & "") * -1, Val(TbCabRegis_new.Fields("F4MONINA") & "")), "###,##0.00")
            If UCase(right(CmbTipDoc.Text, 3)) = "HON" Then
                PnlImpuesto.Caption = "Retención"
                TxtIgv(0).Text = Format(Val("" & TbCabRegis_new.Fields("F4MONTORET")), "###,##0.00")
            Else
                TxtIgv(0).Text = Format(IIf(Val("" & TbCabRegis_new.Fields("F4IGV")) < 0, Val("" & TbCabRegis_new.Fields("F4IGV")) * -1, Val("" & TbCabRegis_new.Fields("F4IGV"))), "###,##0.00")
                TxtOtrImp(0).Text = Format(IIf(Val("" & TbCabRegis_new.Fields("F4OTRIMP")) < 0, Val("" & TbCabRegis_new.Fields("F4OTRIMP")) * -1, Val("" & TbCabRegis_new.Fields("F4OTRIMP"))), "###,##0.00")
            End If
            txtredsuma.Text = Format(IIf(Val("" & TbCabRegis_new.Fields("f4redsuma")) < 0, Val("" & TbCabRegis_new.Fields("f4redsuma")) * -1, Val("" & TbCabRegis_new.Fields("f4redsuma"))), "###,##0.00")
            txtredresta.Text = Format(IIf(Val("" & TbCabRegis_new.Fields("f4redresta")) < 0, Val("" & TbCabRegis_new.Fields("f4redresta")) * -1, Val("" & TbCabRegis_new.Fields("f4redresta"))), "###,##0.00")
            txtdcto.Text = Format(IIf(Val("" & TbCabRegis_new.Fields("f4dcto")) < 0, Val("" & TbCabRegis_new.Fields("f4dcto")) * -1, Val("" & TbCabRegis_new.Fields("f4dcto"))), "###,##0.00")
            PnlTotal(0).Caption = Format(IIf(Val("" & TbCabRegis_new.Fields("F4TOTAL")) < 0, Val("" & TbCabRegis_new.Fields("F4TOTAL")) * -1, Val("" & TbCabRegis_new.Fields("F4TOTAL"))), "###,##0.00")
        Else
            PnlBasImp(0).Caption = Format(TbCabRegis_new!F4BASIMP, "###,##0.00")
            PnlMonIna(0).Caption = Format(TbCabRegis_new.Fields("F4MONINA"), "###,##0.00")
            If UCase(right(CmbTipDoc.Text, 3)) = "HON" Then
                PnlImpuesto.Caption = "Retención"
                TxtIgv(0).Text = Format(TbCabRegis_new.Fields("F4MONTORET"), "###,##0.00")
                TxtOtrImp(0).Text = Format(TbCabRegis_new.Fields("F4FONAVI"), "###,##0.00")
            Else
                TxtIgv(0).Text = Format(TbCabRegis_new.Fields("F4IGV"), "###,##0.00")
                TxtOtrImp(0).Text = Format(TbCabRegis_new.Fields("F4OTRIMP"), "###,##0.00")
            End If
            txtredsuma.Text = Format(TbCabRegis_new.Fields("f4redsuma"), "###,##0.00")
            txtredresta.Text = Format(TbCabRegis_new.Fields("f4redresta"), "###,##0.00")
            txtdcto.Text = Format(TbCabRegis_new.Fields("f4dcto"), "###,##0.00")
            PnlTotal(0).Caption = Format(TbCabRegis_new.Fields("F4TOTAL"), "###,##0.00")
        End If
    
        If TbCabRegis_new.Fields("F4MONEDA") = "S" Then
            Mon.Caption = "US"
           Call Mon_Click
        Else
            Mon.Caption = "MN"
           Call Mon_Click
           
'''''''''            TbOfiRegis.Index = "IDMESNUM"
'''''''''            TbOfiRegis.Seek "=", txtmesmov.Text, TxtNumMov.Text
'''''''''            If Not TbOfiRegis.NoMatch Then
'''''''''                If UCase(Right(CmbTipDoc.Text, 3)) = "CRE" Then
'''''''''                    PnlBasImp(1).Caption = Format(IIf(Val("" & TbOfiRegis.Fields("F4BASIMP")) < 0, Val("" & TbOfiRegis.Fields("F4BASIMP")) * -1, Val("" & TbOfiRegis.Fields("F4BASIMP"))), "###,##0.00")
'''''''''                    PnlMonIna(1).Caption = Format(IIf(Val("" & TbOfiRegis.Fields("F4MONINA")) < 0, Val("" & TbOfiRegis.Fields("F4MONINA")) * -1, Val("" & TbOfiRegis.Fields("F4MONINA"))), "###,##0.00")
'''''''''                    If UCase(Right(CmbTipDoc.Text, 3)) = "HON" Then
'''''''''                        TxtIgv(1).Text = Format(TbOfiRegis.Fields("F4MONTORET"), "###,##0.00")
'''''''''                        TxtOtrImp(1).Text = Format(TbOfiRegis.Fields("F4FONAVI"), "###,##0.00")
'''''''''                    Else
'''''''''                        TxtIgv(1).Text = Format(IIf(Val("" & TbOfiRegis.Fields("F4IGV")) < 0, Val("" & TbOfiRegis.Fields("F4IGV")) * -1, Val("" & TbOfiRegis.Fields("F4IGV"))), "###,##0.00")
'''''''''                        TxtOtrImp(1).Text = Format(IIf(Val("" & TbOfiRegis.Fields("F4OTRIMP")) < 0, Val("" & TbOfiRegis.Fields("F4OTRIMP")) * -1, Val("" & TbOfiRegis.Fields("F4OTRIMP"))), "###,##0.00")
'''''''''                    End If
'''''''''                    PnlTotal(1).Caption = Format(IIf(Val("" & TbOfiRegis.Fields("F4TOTAL")) < 0, Val("" & TbOfiRegis.Fields("F4TOTAL")) * -1, Val("" & TbOfiRegis.Fields("F4TOTAL"))), "###,##0.00")
'''''''''                Else
'''''''''                    PnlBasImp(1).Caption = Format(TbOfiRegis.Fields("F4BASIMP"), "###,##0.00")
'''''''''                    PnlMonIna(1).Caption = Format(TbOfiRegis.Fields("F4MONINA"), "###,##0.00")
'''''''''                    If UCase(Right(CmbTipDoc.Text, 3)) = "HON" Then
'''''''''                        TxtIgv(1).Text = Format(TbOfiRegis.Fields("F4MONTORET"), "###,##0.00")
'''''''''                    Else
'''''''''                        TxtIgv(1).Text = Format(TbOfiRegis.Fields("F4IGV"), "###,##0.00")
'''''''''                        TxtOtrImp(1).Text = Format(TbOfiRegis.Fields("F4OTRIMP"), "###,##0.00")
'''''''''                    End If
'''''''''                    PnlTotal(1).Caption = Format(TbOfiRegis.Fields("F4TOTAL"), "###,##0.00")
'''''''''                End If
'''''''''            End If
        End If
        DELETEREC_N DBTable, cconex_formp, ""
        sw_nuevo_item = True
       ' dxDBGrid1.Dataset.Refresh
        SqlCad = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmesmov.Text + "' AND F4NUMMOV = '" + TxtNumMov.Text + "' Order by F3ITEM"
        If Tbdetcompras.State = adStateOpen Then Tbdetcompras.Close
        Tbdetcompras.Open SqlCad, StrConexDbBancos, adOpenDynamic, adLockOptimistic
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        Do While Not Tbdetcompras.EOF
            dxDBGrid1.Dataset.Append
            dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").value = 0 + Format(Tbdetcompras.Fields("F3ITEM"), "00")
            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").value = "" & Tbdetcompras.Fields("F3GASTO")
            dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").value = "" & Tbdetcompras.Fields("F3CenCos")
            dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").value = "" & Tbdetcompras.Fields("F3CTACON")
            dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").value = "" & Tbdetcompras.Fields("F3CONCEPTO")
            dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value = IIf(IsNull(Tbdetcompras.Fields("F3IMPORTE")), 0#, Format(Tbdetcompras.Fields("F3IMPORTE"), "###,##0.00"))
            dxDBGrid1.Columns.ColumnByFieldName("F3AFECTO").value = IIf("" & Tbdetcompras.Fields("F3AFECTO") = "*", True, False)
            dxDBGrid1.Columns.ColumnByFieldName("F3DEBHAB").value = "" & Tbdetcompras.Fields("F3DEBHAB")
            dxDBGrid1.Columns.ColumnByFieldName("F3ORDEN").value = "" & Tbdetcompras.Fields("F3ORDEN")
            dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value = "" & Tbdetcompras.Fields("f5codpro")
            dxDBGrid1.Columns.ColumnByFieldName("F3CANTIDAD").value = "" & Tbdetcompras.Fields("F3CANTIDAD")
            dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").value = "" & Tbdetcompras.Fields("F3PREUNI")
            dxDBGrid1.Columns.ColumnByFieldName("F4SERGUI").value = "" & Tbdetcompras.Fields("F3SERGUI")
            dxDBGrid1.Columns.ColumnByFieldName("F4NUMGUI").value = "" & Tbdetcompras.Fields("F3NUMGUI")
            dxDBGrid1.Columns.ColumnByFieldName("AFECTO").value = IIf("" & Tbdetcompras.Fields("F3AFECTO") = "*", dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value, 0)
            dxDBGrid1.Columns.ColumnByFieldName("INAFECTO").value = IIf("" & Tbdetcompras.Fields("F3AFECTO") = "*", 0, dxDBGrid1.Columns.ColumnByFieldName("F3IMPORTE").value)
            Tbdetcompras.MoveNext
        Loop
        Tbdetcompras.Close
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Dataset.Post
        dxDBGrid1.Dataset.ADODataset.Requery
        Me.MousePointer = vbDefault
    End If
    sw_nuevo_item = False
    pnlcosto.Caption = ObtenerCampo("CENTROS", "F3DESCRIP", "F3COSTO", Trim(txtcentro.Text) & "", "T", cnn_dbbancos)
End Sub

Private Sub proceso_grid_ordenes()
On Error Resume Next
Dim csql     As String
    
    If CnTmp.State = 1 Then CnTmp.Close
    
    CnTmp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "templus.mdb;Persist Security Info=False" '"tmp_bancos.MDB;Persist Security Info=False"
    
        
        
    If ctipoadm_bd = "M" Then
        csql = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4OBSERVA, IF4ORDEN.F4CODPRV, EF2PROVEEDORES.F2NOMPROV"
        csql = csql + " FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC ORDER BY IF4ORDEN.F4NUMORD desc"
    Else
        If whelpoc = "S" Then
            If wtipoc = "I" Then
                csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                        "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                        "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                        "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
                        "AND F4LOCAL='0' " & _
                        "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND (F4CODPRV = '" & wRucCliProv & "') ) " & _
                        "ORDER BY IF4ORDEN.F4NUMORD desc"
            Else
                csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                        "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                        "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                        "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
                        "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND (F4CODPRV = '" & wRucCliProv & "') ) " & _
                        "ORDER BY IF4ORDEN.F4NUMORD desc"
            End If
        Else
            If wtipoc = "I" Then
                csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                       "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                       "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                       "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0))) " & _
                       "AND F4LOCAL='0' " & _
                       "ORDER BY IF4ORDEN.F4NUMORD desc"
                csql = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CODPRV, P.F2NOMPROV AS PROV, E.F2NOMPROV AS EMB, IF4ORDEN.F4OBSERVA "
                csql = csql & "FROM (IF4ORDEN LEFT JOIN EF2PROVEEDORES AS P ON IF4ORDEN.F4CODPRV = P.F2NEWRUC) "
                csql = csql & "LEFT JOIN EF2PROVEEDORES AS E ON IF4ORDEN.F4CODCLI = E.F2CODPROV "
                csql = csql & "WHERE (((IF4ORDEN.F4NUMORD) "
                csql = csql & "In (SELECT DISTINCTROW IF4ORDEN.F4NUMORD "
                csql = csql & "FROM IF4ORDEN INNER JOIN IF3ORDEN "
                csql = csql & "ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD "
                csql = csql & "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV "
                csql = csql & "HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND ((IF4ORDEN.F4LOCAL)='0')) "
                csql = csql & "ORDER BY IF4ORDEN.F4NUMORD DESC"
            Else
                csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                        "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                        "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                        "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0))) " & _
                        "ORDER BY IF4ORDEN.F4NUMORD desc"
            
                csql = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, ORDEN.F3CODFAB, "
                csql = csql & "ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL,ORDEN.F4NUMORD + format(ORDEN.item,'000') as LLave  "
                csql = csql & "FROM "
                csql = csql & "(SELECT Right(IF4ORDEN.F4NUMORD,5) AS CODCENTRO, "
                csql = csql & "'[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']' AS Grupo, "
                csql = csql & "EF2PROVEEDORES.F2NEWRUC,IF4ORDEN.item , IF4ORDEN.F4NUMORD,IF4ORDEN.F4local, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, "
                csql = csql & "IF3ORDEN.F5NOMPRO, IF3ORDEN.F3PREUNI, IF3ORDEN.F3CANPRO, IF3ORDEN.F3TOTAL "
                csql = csql & "FROM (IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
                csql = csql & "INNER JOIN IF3ORDEN ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
                csql = csql & "Where (((EF2PROVEEDORES.F2NEWRUC) = '" & wRucCliProv & "') ) "
                csql = csql & "ORDER BY Right(IF4ORDEN.F4NUMORD,5), Left(IF4ORDEN.F4NUMORD,12)) as "
                csql = csql & "ORDEN LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.CCONCAR "
                csql = csql & "ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo"
            End If
        End If
    End If
    
    DELETEREC_N "Tmp_ImportaOC", cconex_formp, ""
    
    Dim Amov(0 To 30) As a_grabacion
    Dim X As Integer
    
    X = 1
    
    If Rs.State = 1 Then Rs.Close
    
    Rs.Open csql, cnn_dbbancos, 3, 1
    
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        
        With Progreso_Avance
        .Show
        .prbavanza.value = 0
        .prbavanza.Max = Rs.RecordCount
        .LblDetalle.Caption = "Cargando Ordenes de Compra..."
        .LblPorc.Caption = "0.000 %"
        
        Do While Not Rs.EOF
            Amov(0).campo = "llave": Amov(0).valor = X: Amov(0).Tipo = "N"
            Amov(1).campo = "F3DESCRIP": Amov(1).valor = Rs!F3DESCRIP & "": Amov(1).Tipo = "N"
            Amov(2).campo = "Grupo": Amov(2).valor = Rs!Grupo & "": Amov(2).Tipo = "T"
            Amov(3).campo = "F2NEWRUC": Amov(3).valor = Rs!F2NEWRUC & "": Amov(3).Tipo = "T"
            Amov(4).campo = "F4NUMORD": Amov(4).valor = Rs!F4NUMORD & "": Amov(4).Tipo = "T"
            Amov(5).campo = "F4FECEMI": Amov(5).valor = Rs!F4FECEMI & "": Amov(5).Tipo = "F"
            Amov(6).campo = "F3CODFAB": Amov(6).valor = Rs!F3CODFAB & "": Amov(6).Tipo = "T"
            Amov(7).campo = "F5NOMPRO": Amov(7).valor = Rs!F5NOMPRO & "": Amov(7).Tipo = "T"
            Amov(8).campo = "F3PREUNI": Amov(8).valor = Rs!F3PREUNI & "": Amov(8).Tipo = "N"
            Amov(9).campo = "F3CANPRO": Amov(9).valor = Rs!F3CANPRO & "": Amov(9).Tipo = "N"
            Amov(10).campo = "F3TOTAL": Amov(10).valor = Rs!F3TOTAL & "": Amov(10).Tipo = "N"
            GRABA_REGISTRO Amov, "Tmp_ImportaOC", "A", 10, cconex_formp, ""
            .prbavanza.value = .prbavanza.value + 1
            .LblPorc.Caption = Format(.prbavanza.value * 100 / .prbavanza.Max, "0.000") & " %"
            .Refresh
            X = X + 1
            Rs.MoveNext
        Loop
        
        End With
        If CnTmp.State = 1 Then CnTmp.Close
        Set CnTmp = Nothing
        Unload Progreso_Avance
        Set Progreso_Avance = Nothing
            
        
    End If

    
 
End Sub

Private Function VerificaGuiaCabecera() As Boolean
VerificaGuiaCabecera = False
If Len(Trim(TxtNumGuia.Text)) > 0 Or Len(Trim(txtSerGuia.Text)) > 0 Then
    VerificaGuiaCabecera = True
End If
End Function

Private Sub Upd_Orden_Pagos(FLGELI As String)
Dim Correla As String
Dim IDOCompra As String

If FLGELI = "1" Then
    Correla = 0
    csql = "Update IF4ORDEN_PAGO set correladoc=" & Correla & ", F4ESTADO=5 Where IDOP = '" & CodigoOrdPago & "'"
    cnn_dbbancos.Execute csql
    
Else
    Correla = traerCampo("REGISDOC", "F4CORRELA", "F4OCOMPRA", txtocompra.Text, " And F4NUMMOV = '" & TxtNumMov.Text & "' And F4MESMOV = '" & wanno & wmes & "'")
    csql = "Update IF4ORDEN_PAGO set correladoc=" & Correla & ", F4ESTADO=3 Where IDOP = '" & CodigoOrdPago & "'"
    cnn_dbbancos.Execute csql
End If
End Sub

