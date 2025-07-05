VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form mant_productos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Productos"
   ClientHeight    =   7425
   ClientLeft      =   8295
   ClientTop       =   2190
   ClientWidth     =   9510
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7425
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      Caption         =   "Datos Contables"
      Enabled         =   0   'False
      Height          =   1035
      Left            =   -23864
      TabIndex        =   60
      Top             =   -20114
      Width           =   8745
      Begin VB.TextBox cuenta2 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   76
         Top             =   1530
         Width           =   1095
      End
      Begin VB.TextBox cuenta1 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   74
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox txtgasto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         MaxLength       =   4
         TabIndex        =   72
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox cuenta 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   70
         Top             =   420
         Width           =   1095
      End
      Begin Threed.SSPanel pnlgasto 
         Height          =   330
         Left            =   2535
         TabIndex        =   61
         Top             =   420
         Width           =   3705
         _Version        =   65536
         _ExtentX        =   6535
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
      Begin Threed.SSPanel pnlventa 
         Height          =   330
         Left            =   2565
         TabIndex        =   62
         Top             =   1035
         Width           =   3705
         _Version        =   65536
         _ExtentX        =   6535
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
      Begin Threed.SSPanel pnlinventario 
         Height          =   330
         Left            =   2565
         TabIndex        =   63
         Top             =   1530
         Width           =   3705
         _Version        =   65536
         _ExtentX        =   6535
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
      Begin VB.Label Label4 
         Caption         =   "Cod. Gasto"
         Height          =   255
         Left            =   6600
         TabIndex        =   67
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventarios"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   66
         Top             =   1580
         Width           =   795
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   65
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   64
         Top             =   480
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7080
      Left            =   -25229
      TabIndex        =   13
      Top             =   -22664
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   12488
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
      BorderWidth     =   1
      BevelInner      =   1
      Enabled         =   0   'False
      Begin VB.PictureBox fg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   5910
         Left            =   225
         ScaleHeight     =   5880
         ScaleWidth      =   9525
         TabIndex        =   14
         Top             =   630
         Width           =   9555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   5910
         Left            =   405
         Top             =   765
         Width           =   9510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4005
         TabIndex        =   15
         Top             =   90
         Width           =   1560
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   900
      Left            =   -24104
      TabIndex        =   19
      Top             =   -21779
      Visible         =   0   'False
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   1587
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
      Enabled         =   0   'False
      Begin VB.ComboBox cmbtipo 
         Height          =   330
         ItemData        =   "mant_productos.frx":0000
         Left            =   1560
         List            =   "mant_productos.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1200
         Width           =   2100
      End
      Begin VB.TextBox txttceuros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Text            =   "0.000"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtCostoEuros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1710
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtfactor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1710
         TabIndex        =   29
         Text            =   "0.000"
         Top             =   3195
         Width           =   1575
      End
      Begin VB.TextBox txtprevta 
         Alignment       =   1  'Right Justify
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
         Left            =   1710
         TabIndex        =   28
         Text            =   "0.0"
         Top             =   4455
         Width           =   1575
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   4050
         Width           =   1575
      End
      Begin VB.TextBox txtvalvta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   4050
         Width           =   1575
      End
      Begin VB.TextBox fob 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1695
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   2700
         Width           =   1575
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   330
         ItemData        =   "mant_productos.frx":0004
         Left            =   1560
         List            =   "mant_productos.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   2100
      End
      Begin VB.TextBox txtflete 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6345
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2655
         Width           =   1575
      End
      Begin VB.TextBox txtinstalacion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1710
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   3645
         Width           =   1575
      End
      Begin VB.TextBox txtmanipuleo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6345
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   3105
         Width           =   1575
      End
      Begin VB.TextBox txtotroscostos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6345
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   3555
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Costo"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   24
         Left            =   210
         TabIndex        =   46
         Top             =   1245
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   23
         Left            =   4365
         TabIndex        =   45
         Top             =   2175
         Width           =   870
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Costo Euros"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   22
         Left            =   200
         TabIndex        =   44
         Top             =   2175
         Width           =   885
      End
      Begin VB.Label lblmoneda 
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   2
         Left            =   870
         TabIndex        =   43
         Top             =   2745
         Width           =   315
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda Compra"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   21
         Left            =   210
         TabIndex        =   42
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factor"
         Height          =   210
         Left            =   200
         TabIndex        =   41
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label label04 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Precio Venta"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   200
         TabIndex        =   40
         Top             =   4545
         Width           =   930
      End
      Begin VB.Label label04 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "I.G.V."
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4365
         TabIndex        =   39
         Top             =   4140
         Width           =   405
      End
      Begin VB.Label label04 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Valor de Venta"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   200
         TabIndex        =   38
         Top             =   4095
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Costo"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   200
         TabIndex        =   37
         Top             =   2745
         Width           =   420
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   8520
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   8520
         Y1              =   5085
         Y2              =   5085
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Flete %"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   4365
         TabIndex        =   36
         Top             =   2745
         Width           =   540
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Manipuleo %"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   4365
         TabIndex        =   35
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Instalación %"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   200
         TabIndex        =   34
         Top             =   3645
         Width           =   960
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Otros costos"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   13
         Left            =   4365
         TabIndex        =   33
         Top             =   3645
         Width           =   945
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1710
      Left            =   -23639
      TabIndex        =   47
      Top             =   -17309
      Width           =   8160
      _Version        =   65536
      _ExtentX        =   14393
      _ExtentY        =   3016
      _StockProps     =   14
      ForeColor       =   0
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
      Enabled         =   0   'False
      Begin VB.TextBox txtF5TREP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6405
         TabIndex        =   68
         Text            =   "0"
         Top             =   1275
         Width           =   1170
      End
      Begin VB.TextBox TxtCodubi2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1305
         MaxLength       =   4
         TabIndex        =   52
         Top             =   840
         Width           =   510
      End
      Begin VB.TextBox ubicacion 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1305
         MaxLength       =   4
         TabIndex        =   51
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox maximo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6405
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   330
         Width           =   1170
      End
      Begin VB.TextBox minimo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6405
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   810
         Width           =   1170
      End
      Begin VB.TextBox txtpiezas 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1305
         MaxLength       =   13
         TabIndex        =   48
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Threed.SSPanel TxtUbica1 
         Height          =   315
         Left            =   1845
         TabIndex        =   53
         Top             =   360
         Width           =   2670
         _Version        =   65536
         _ExtentX        =   4710
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
      Begin Threed.SSPanel TxtUbica2 
         Height          =   315
         Left            =   1845
         TabIndex        =   54
         Top             =   840
         Width           =   2670
         _Version        =   65536
         _ExtentX        =   4710
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
      Begin VB.Label Label11 
         Caption         =   "días"
         Height          =   180
         Left            =   7635
         TabIndex        =   71
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de Reposición"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   14
         Left            =   4770
         TabIndex        =   69
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 2"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   15
         Left            =   180
         TabIndex        =   59
         Top             =   885
         Width           =   840
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 1"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   12
         Left            =   180
         TabIndex        =   58
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Stock  Mínimo"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   5370
         TabIndex        =   57
         Top             =   855
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Stock Máximo"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   5370
         TabIndex        =   56
         Top             =   375
         Width           =   990
      End
      Begin VB.Label lblpiezas 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº de Piezas"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   55
         Top             =   1800
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.PictureBox Panelcod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   360
      ScaleHeight     =   4695
      ScaleWidth      =   8160
      TabIndex        =   73
      Top             =   5160
      Visible         =   0   'False
      Width           =   8220
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   4185
         TabIndex        =   79
         Top             =   3360
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   0
         Left            =   585
         TabIndex        =   0
         Top             =   585
         Width           =   3210
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   1
         Left            =   4185
         TabIndex        =   78
         Top             =   585
         Width           =   3210
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   2
         Left            =   585
         TabIndex        =   77
         Top             =   1935
         Visible         =   0   'False
         Width           =   6810
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   3
         Left            =   585
         TabIndex        =   75
         Top             =   3360
         Visible         =   0   'False
         Width           =   3210
      End
      Begin Threed.SSCommand BtnFin 
         Height          =   375
         Left            =   3600
         TabIndex        =   80
         Top             =   4320
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Finalizar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "L1"
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
         Index           =   4
         Left            =   4200
         TabIndex        =   85
         Top             =   3105
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label LblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "L1"
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
         Height          =   240
         Index           =   3
         Left            =   615
         TabIndex        =   84
         Top             =   3105
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label LblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "L1"
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
         Index           =   2
         Left            =   630
         TabIndex        =   83
         Top             =   1710
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label LblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "L1"
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
         Index           =   1
         Left            =   4230
         TabIndex        =   82
         Top             =   315
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label LblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "L1"
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
         Index           =   0
         Left            =   630
         TabIndex        =   81
         Top             =   315
         Width           =   180
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6090
      Left            =   240
      TabIndex        =   16
      Top             =   405
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
      _ExtentY        =   10742
      _StockProps     =   14
      ForeColor       =   0
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
      Begin MSComCtl2.DTPicker txtfecing 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97452033
         CurrentDate     =   40620
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo"
         Height          =   555
         Left            =   270
         TabIndex        =   104
         Top             =   180
         Width           =   3345
         Begin Threed.SSOption opttipo 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   105
            Top             =   225
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Producto"
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
         Begin Threed.SSOption opttipo 
            Height          =   255
            Index           =   1
            Left            =   2070
            TabIndex        =   106
            Top             =   225
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Servicio"
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
      Begin VB.TextBox txtnompro2 
         Height          =   315
         Left            =   5265
         MaxLength       =   17
         TabIndex        =   10
         Top             =   3195
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.TextBox Txtcodmedc 
         Height          =   315
         Left            =   1335
         MaxLength       =   50
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   1665
         Width           =   1065
      End
      Begin VB.TextBox Txtnompro1 
         Height          =   315
         Left            =   1380
         MaxLength       =   190
         TabIndex        =   2
         Top             =   1260
         Width           =   6750
      End
      Begin VB.TextBox txcodmarca 
         Height          =   315
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2025
         Width           =   1050
      End
      Begin VB.TextBox Txtcodalm 
         Height          =   315
         Left            =   5220
         MaxLength       =   3
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   2025
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtcodpro 
         Height          =   315
         Left            =   1350
         TabIndex        =   90
         Top             =   855
         Width           =   1500
      End
      Begin VB.TextBox txtcodfab 
         Height          =   315
         Left            =   5700
         TabIndex        =   4
         Top             =   855
         Width           =   2380
      End
      Begin VB.TextBox arancelaria 
         Height          =   315
         Left            =   5265
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2655
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.TextBox txtnompro4 
         Height          =   795
         Left            =   9600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   5280
         Width           =   6615
      End
      Begin VB.TextBox txtnompro3 
         Height          =   795
         Left            =   11040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   4560
         Width           =   6615
      End
      Begin Threed.SSCheck chkafecto 
         Height          =   240
         Left            =   7200
         TabIndex        =   91
         Top             =   1740
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Afecto "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel txmarca 
         Height          =   315
         Left            =   2520
         TabIndex        =   92
         Top             =   2025
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
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
      Begin Threed.SSPanel Txtnomalm 
         Height          =   315
         Left            =   5715
         TabIndex        =   93
         Top             =   2025
         Visible         =   0   'False
         Width           =   2370
         _Version        =   65536
         _ExtentX        =   4180
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
      Begin Threed.SSPanel Txtnommedc 
         Height          =   315
         Left            =   2520
         TabIndex        =   94
         Top             =   1665
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
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
      Begin Threed.SSFrame SSFrame4 
         Height          =   645
         Left            =   315
         TabIndex        =   95
         Top             =   2430
         Width           =   2940
         _Version        =   65536
         _ExtentX        =   5186
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Origen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optfactor 
            Height          =   255
            Index           =   0
            Left            =   315
            TabIndex        =   7
            Top             =   315
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Nacional"
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
         Begin Threed.SSOption optfactor 
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   8
            Top             =   285
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Importado"
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
      Begin Threed.SSCheck chkinsumo 
         Height          =   240
         Left            =   1560
         TabIndex        =   96
         Top             =   3180
         Visible         =   0   'False
         Width           =   750
         _Version        =   65536
         _ExtentX        =   1323
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Insumo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkdescontinuado 
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   3420
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Producto Descontinuado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   1935
         Left            =   420
         OleObjectBlob   =   "mant_productos.frx":0008
         TabIndex        =   98
         Top             =   3660
         Visible         =   0   'False
         Width           =   7695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrip. Etiqueta"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   3720
         TabIndex        =   103
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   4305
         TabIndex        =   102
         Top             =   2070
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   5400
         TabIndex        =   101
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Fabricante"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4305
         TabIndex        =   100
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Partida Arancelaria"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   3690
         TabIndex        =   99
         Top             =   2700
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca Producto"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   89
         Top             =   2115
         Width           =   1140
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   765
         TabIndex        =   88
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   405
         TabIndex        =   87
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "U. Medida"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   11
         Left            =   510
         TabIndex        =   86
         Top             =   1710
         Width           =   705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Castelleno"
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   25
         Left            =   10920
         TabIndex        =   18
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Inglés"
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   20
         Left            =   10080
         TabIndex        =   17
         Top             =   6600
         Width           =   1080
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "mant_productos.frx":28C2
      ToolBars        =   "mant_productos.frx":8144
   End
   Begin VB.Menu manugraba 
      Caption         =   "&Registro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu menunuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu manugrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu menueliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu menuimprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu rayas 
         Caption         =   "-"
      End
      Begin VB.Menu menusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu menu200 
      Caption         =   "&Edición"
      Visible         =   0   'False
      Begin VB.Menu menu110 
         Caption         =   "Lista de p&roveedores"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu menu110 
         Caption         =   "&Lista de precios"
         Index           =   1
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "mant_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn             As ADODB.Connection
Dim cnTP            As New ADODB.Connection
Dim rstderecho      As New ADODB.Recordset
Dim rstparametro    As New ADODB.Recordset
Dim rstconsulta     As New ADODB.Recordset
Dim rst             As New ADODB.Recordset
Dim rsfield         As New ADODB.Recordset
Dim RsProducto      As New ADODB.Recordset
Dim rsTMV           As New Recordset
Dim wselec          As Integer
Dim wcodigo         As String * 10
Dim wlong1          As Integer
Dim wlong2          As Integer
Dim wniveles        As Integer
Dim wcod(5)         As String
Dim wdes(5)         As String
Dim wtippro         As String * 1
Dim wtipo           As String
Dim wgraba          As String
Dim sql             As String
Dim rsgastos        As New ADODB.Recordset
Dim sw_ayuda        As Boolean
Dim sw_activate     As Boolean
Dim gestpro         As String
Dim gestval         As String
Dim swmant_prod_especiales As Boolean
Dim derecho(4)      As Boolean
Dim wFob            As Double
Dim wvalvta1        As Double
Dim xigv            As Double
Dim wprevta         As Double
Dim wvalvta2        As Double
Dim sw_ayuda_um     As Boolean
Dim sw_ayuda_alma   As Boolean
Dim sw_nuevo_item   As Boolean

Private Function VERIFICA_MOV(pproducto As String)
Dim cverif      As String
Dim rsverif     As New ADODB.Recordset
Dim ctipo       As String
    
    '-----------------------------------------------------------
    ctipo = ""
    cverif = "SELECT F5CODPRO FROM TBVENTA_DET WHERE F5CODPRO='" & pproducto & "'"
    If rsverif.State = adStateOpen Then rsverif.Close
    rsverif.Open cverif, cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not rsverif.EOF Then
        ctipo = "P"
    End If
    rsverif.Close
    Set rsverif = Nothing
    
    VERIFICA_MOV = ctipo

End Function

Private Sub Actualiza_Producto(codprod)
    Dim i       As Integer
    If RsProducto.State = adStateOpen Then RsProducto.Close
    RsProducto.Open "SELECT * FROM IF5PLA WHERE F5CODPRO='" & codprod & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProducto.EOF Then
        txtcodpro.Text = "" & RsProducto.Fields("F5CODPRO")
        txtCodFab.Text = "" & RsProducto.Fields("F5CODFAB")
        wtippro = Mid(txtcodpro.Text, 7, 1)
        Txtnompro1.Text = "" & RsProducto.Fields("F5NOMPRO")
        txtnompro2.Text = "" & RsProducto.Fields("F5NOMPRO2")
        txtnompro3.Text = "" & RsProducto.Fields("F5TEXTO")
        txtnompro4.Text = "" & RsProducto.Fields("F5TEXTO_ING")
        txcodmarca.Text = "" & RsProducto.Fields("F5marca")
        chkDescontinuado.value = IIf(RsProducto("f5descontinuado") = "S", True, False)
        'txtCostoEuros.Text = Format(RsProducto("f5costoeuro"), "#,###,###0.00")
        'txttceuros.Text = Format(RsProducto("f5tceuro"), "#,###,###0.00")
                
        wTipoCosto = "" & RsProducto("f5tipocosto")
        For J = 0 To cmbtipo.ListCount - 1
            If cmbtipo.List(J) = wTipoCosto Then
                cmbtipo.ListIndex = J
                Exit For
            End If
        Next J
                
        If Len(Trim(txcodmarca.Text)) > 0 Then
            sql = "select F2DESMAR from EF2MARCAS where F2CODMAR = '" & txcodmarca.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
                txmarca.Caption = Trim("" & rsfield.Fields("F2DESMAR"))
            End If
            rsfield.Close
        End If

        If "" & RsProducto.Fields("F5insumo") = "*" Then
            chkinsumo.value = True
        Else
            chkinsumo.value = False
        End If
        If "" & RsProducto.Fields("F5TIPESTADO") = "N" Then
            optfactor(0).value = True
            optfactor(1).value = False
        Else
            optfactor(0).value = False
            optfactor(1).value = True
        End If
        txtfactor.Text = Val(RsProducto.Fields("F5FACTOR") & "")
        sql = "SELECT F2CODALM FROM IF6ALMA WHERE F5CODPRO = '" & Trim(txtcodpro.Text) & "'"
        If rsfield.State = 1 Then rsfield.Close
        rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsfield.EOF Then
            Txtcodalm.Text = "" & rsfield.Fields("F2CODALM")
            chkDescontinuado.Enabled = True
        Else
            Txtcodalm.Text = ""
            chkDescontinuado.Enabled = False
        End If
        rsfield.Close
        
        TXTFECING.value = IIf(IsNull(RsProducto.Fields("F5FECING")), Format(Now, "DD/MM/YYYY"), RsProducto.Fields("F5FECING"))
        CmbMoneda.ListIndex = IIf("" & RsProducto.Fields("F5MONEDA") = "S", 0, 1)
        fob.Text = Val("" & RsProducto.Fields("F5FOB"))
        txtflete.Text = Val("" & RsProducto.Fields("F5FLETE"))
        txtinstalacion.Text = Val("" & RsProducto.Fields("F5INSTALACION"))
        txtmanipuleo.Text = Val("" & RsProducto.Fields("F5MANIPULEO"))
        txtotroscostos.Text = Val("" & RsProducto.Fields("F5OTROSCOSTOS"))
        txtvalvta.Text = Val("" & RsProducto.Fields("F5valvta"))
        TxtIgv.Text = Val("" & RsProducto.Fields("F5IGVVTA"))
        txtprevta.Text = Val("" & RsProducto.Fields("F5PREVTA"))
        arancelaria.Text = "" & RsProducto.Fields("F5PARTARA")
        minimo.Text = Format(Val("" & RsProducto.Fields("F5stockmin")), "0.00")
        maximo.Text = Format(Val("" & RsProducto.Fields("F5stockmax")), "0.00")
        txtF5TREP.Text = Format(Val("" & IIf(IsNull(RsProducto.Fields("F5TREP")), 0, RsProducto.Fields("F5TREP"))), "0")
        TxtCodubi2.Text = Trim("" & RsProducto.Fields("F5ubica2"))
        ubicacion.Text = Trim("" & RsProducto.Fields("F5ubicacio"))
        cuenta.Text = Trim("" & RsProducto.Fields("F5ctacon"))
        Txtcodmedc.Text = "" & RsProducto.Fields("F7CODMED")
        If Len(Trim(Txtcodmedc.Text)) > 0 Then
            sql = "SELECT F7NOMMED FROM EF7MEDIDAS WHERE F7CODMED = '" & Txtcodmedc.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
                Txtnommedc.Caption = "" & rsfield.Fields("F7NOMMED")
            End If
            rsfield.Close
        End If
        sql = "SELECT F2NOMALM FROM EF2ALMACENES WHERE F2CODALM = '" & Txtcodalm.Text & "' "
        If rsfield.State = 1 Then rsfield.Close
        rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsfield.EOF Then
            Txtnomalm.Caption = rsfield.Fields("F2NOMALM")
        Else
            Txtnomalm.Caption = ""
        End If
        rsfield.Close
        If Trim("" & RsProducto.Fields("F5TIPO")) = "P" Then
            opttipo(0).value = True
            opttipo(1).value = False
        Else
            opttipo(0).value = False
            opttipo(1).value = True
        End If
        txtcodpro.Enabled = False
        TXTFECING.Enabled = False
        If Trim("" & RsProducto.Fields("F5AFECTO")) = "*" Then
            chkafecto.value = True
        Else
            chkafecto.value = False
        End If
        
        cuenta.Text = "" & RsProducto("F5ctacon")
        cuenta_LostFocus
        txtgasto.Text = Trim("" & RsProducto.Fields("F3GASTO"))
        cuenta1.Text = "" & RsProducto("F5ctacon1")
        cuenta1_LostFocus
        cuenta2.Text = "" & RsProducto("f5ctacon2")
        cuenta2_LostFocus
        
        txtpiezas.Text = Val(RsProducto.Fields("F5PIEZAS") & "")
        '-----------------------------------
        gestpro = RsProducto.Fields("F5ESTPRO") & ""
        gestval = RsProducto.Fields("F5ESTVAL") & ""
        '-----------------------------------
        
        'busco medidas de venta
        AbrirBasesTemp
        
        sql = "Delete from tmpMedVenta"
        cnTP.Execute sql
        'AlmacenaQuery_sql sql, cnTP
        
        AbrirBasesTemp
        If rsTMV.State = 1 Then rsTMV.Close
        rsTMV.Open "select * from MEDIVENTAS where F5CODPRO = '" & txtcodpro.Text & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
        
        If Not rsTMV.EOF Then
            i = 0
            Do While Not rsTMV.EOF
                i = i + 1
                sql = "insert into tmpMedVenta (ITEM,F7CODMED,FACTOR,F5PREVTA) values('" & i & " ','" & rsTMV.Fields("F7CODMED") & "'," & rsTMV.Fields("F5FACTOR") & "," & rsTMV.Fields("F5PREVTA") & ")"
                cnTP.Execute sql
                 'AlmacenaQuery_sql sql, cnTP
                rsTMV.MoveNext
            Loop
        Else
            AdicionaItem
        End If
        AbrirBasesTemp
        
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnTP
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
    Else
        txtcodpro.Enabled = True
        TXTFECING.Enabled = True
        MsgBox "Código de Producto no existe. Verifique.", 16, "Atención"
    End If
    RsProducto.Close
End Sub

Private Sub BtnFin_Click()
If Not rstparametro.EOF Then
    wcod(Index + 1) = left(List1(Index).Text, rstparametro.Fields("F1LONNIV" & Format(Index + 1, "0")))
    txtcodpro.Text = Trim(wcod(1)) & Trim(wcod(2)) & Trim(wcod(3)) & Trim(wcod(4)) & Trim(wcod(5))
    calcula_codigo
    Cierra_Codigo
    Else
        Exit Sub
End If

End Sub

Private Sub calcula_codigo()
Dim wser    As String
Dim wcod    As String
Dim wformat As String
Dim i       As Integer
    
    wser = Trim(txtcodpro.Text)
    wcod = "0"
    wformat = ""
    For i = 1 To wlong1 + wlong2 - Len(txtcodpro.Text)
        wformat = wformat & "0"
    Next i
    If swmant_prod_especiales = True Then
        sql = "SELECT F5CODPRO FROM IF5PLA_ESPECIALES WHERE F5CODPRO LIKE '" + wser & "%" + "' ORDER BY  F5CODPRO DESC"
    Else
        sql = "SELECT F5CODPRO FROM IF5PLA WHERE left(F5CODPRO," & wlong1 & ") = '" & wser & "' ORDER BY  F5CODPRO DESC"
    End If
    If rsfield.State = 1 Then rsfield.Close
    rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsfield.EOF Then
        rsfield.MoveFirst
        wcod = right(Trim(rsfield.Fields("F5codpro")), (wlong1 + wlong2 - Len(txtcodpro.Text)))
    End If
    wcod = Format(Val(wcod) + 1, wformat)
    wcodigo = txtcodpro.Text & wcod

'    If swmant_prod_especiales = False Then
'        sw = 0
'        Do While sw = 0
'            sql = "SELECT F5CODPRO FROM IF5PLA_ESPECIALES WHERE F5CODPRO LIKE '" + wser & "%" + "' ORDER BY  F5CODPRO DESC"
'            If rsfield.State = 1 Then rsfield.Close
'            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not rsfield.EOF Then
'                wcod = Right(Trim(rsfield.Fields("F5codpro")), (wlong1 + wlong2 - Len(txtcodpro.Text)))
'                wcodigo = Trim(wser) & Format(Val(wcod) + 1, wformat)
'                sw = 1
'            Else
'                sw = 0
'                If rsfield.EOF Then sw = 1
'            End If
'            rsfield.Close
'        Loop
'    End If
    
    txtcodpro.Text = Trim(wcodigo)
    txtCodFab.Text = Trim(wcodigo)
    Txtcodalm.Text = "01"
    txcodmarca.Text = "001"

End Sub

Private Sub Cierra_Codigo()
    'vaTabPro1.Enabled = True
    Panelcod.Visible = False
    List1(1).Visible = False
    List1(2).Visible = False
    List1(3).Visible = False
    List1(4).Visible = False
    wcod(1) = "": wcod(2) = "": wcod(3) = "": wcod(4) = "": wcod(5) = ""
    wdes(1) = "": wdes(2) = "": wdes(3) = "": wdes(4) = "": wdes(5) = ""
    LblNivel(1).Visible = False: LblNivel(2).Visible = False
    LblNivel(3).Visible = False: LblNivel(4).Visible = False
    Txtnompro1.SetFocus

End Sub

Private Sub chkafecto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub


Private Sub chkdescontinuado_Click(value As Integer)
If chkDescontinuado.value And Trim(txtcodpro.Text) <> "" Then
    sql = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr " _
    & "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) " _
    & "WHERE f5codpro='" & txtcodpro.Text & "' and IF4VALES.f2codalm='" & Txtcodalm.Text & "'"
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        nstock = Val("" & rst("ing")) - Val("" & rst("egr"))
    Else
        nstock = 0
    End If
    
    If nstock > 0 Then
        MsgBox "El Producto no Puede Descontinuarse. Tiene " & nstock & " Unidades en Stock", 16, "Sistema de Logistica"
        chkDescontinuado.value = False
    End If
End If
End Sub

Private Sub Cmbmoneda_Click()

    If CmbMoneda.ListIndex = 0 Then
        lblmoneda(2).Caption = "S/."
        ControlaEuro False
    ElseIf CmbMoneda.ListIndex = 1 Then
        lblmoneda(2).Caption = "US$"
        ControlaEuro False
    Else
        lblmoneda(2).Caption = "US$"
        ControlaEuro True
    End If

End Sub

Private Sub cmbmoneda_keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    
End Sub

Private Sub cuenta_Change()
    
    pnlgasto.Caption = "": txtgasto.Text = ""
    
End Sub

Private Sub cuenta_GotFocus()

    cuenta.SelStart = 0: cuenta.SelLength = Len(cuenta.Text)
    
End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"

End Sub

Private Sub Elimina_Producto()
Dim ctipomov    As String

    Beep
    If MsgBox("¿Está seguro de eliminar el Producto?", 36, "Atención") = 6 Then
        ctipomov = ""
        ctipomov = VERIFICA_MOV(txtcodpro.Text)
        If Len(ctipomov) = 0 Then
            sql = "select f5codpro from if5pla where f5codpro= '" & txtcodpro.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
            
                csql = "Delete From IF6ALMA where F5codpro='" & txtcodpro.Text & "'"
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
                
                csql = "Delete From if5pla where F5codpro='" & txtcodpro.Text & "'"
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
                
                csql = "Delete from MEDIVENTAS where F5CODPRO = '" & txtcodpro.Text & "'"
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
                
            End If
            rsfield.Close
            Nuevo_Producto
        Else
            MsgBox "El producto no se puede eliminar, ya ha sido facturado. Verifique", 16, "Sistema de Logistica"
        End If
    End If

End Sub

Private Sub cuenta1_Change()

    pnlventa.Caption = ""
    
End Sub

Private Sub cuenta1_GotFocus()

    cuenta1.SelStart = 0: cuenta1.SelLength = Len(cuenta1.Text)
    
End Sub

Private Sub cuenta1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub

Private Sub cuenta1_LostFocus()

    If rst.State = adStateOpen Then rst.Close
    sql = "select f5codcta, f5nomcta from cf5pla where f5codcta='" & cuenta1.Text & "'"
    rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        pnlventa.Caption = "" & rst("f5nomcta")
    End If
    rst.Close

End Sub

Private Sub cuenta2_Change()
    
    pnlinventario.Caption = ""
    
End Sub

Private Sub cuenta2_GotFocus()

    cuenta2.SelStart = 0: cuenta2.SelLength = Len(cuenta2.Text)
    
End Sub

Private Sub cuenta2_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cuenta.SetFocus
    End If

End Sub

Private Sub cuenta2_LostFocus()

    If rst.State = adStateOpen Then rst.Close
    sql = "select f5codcta, f5nomcta from cf5pla where f5codcta='" & cuenta2.Text & "'"
    rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        pnlinventario.Caption = "" & rst("f5nomcta")
    End If
    rst.Close

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    'If dxDBGrid1.Dataset.State = dsEdit Then
    '    dxDBGrid1.Dataset.Edit
    '    dxDBGrid1.Dataset.Post
    '    dxDBGrid1.Dataset.Refresh
    'End If
    
End Sub

Private Sub fob_GotFocus()
    
    fob.SelStart = 0: fob.SelLength = Len(fob.Text)
    
End Sub

Private Sub fob_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub fob_LostFocus()

    If Val(Format(fob.Text, "0.00")) > 0 Then
        fob.Text = Format(fob.Text, "###,###,##0.00")
        CALCULA_IMPORTES
    Else
        MsgBox "No ha Ingresado el FOB del Producto", 16, "Sistema de Logistica"
    End If

End Sub

Private Sub Form_Activate()

''''    If vaTabPro1.TabState = 2 Then
''''        If sw_activate = True Then
''''            If sw_nuevo_mant = True Then
''''                If List1(1).Visible = True And List1(1).ListCount > 0 Then
''''                    List1(1).SetFocus
''''                End If
''''            End If
''''            sw_activate = False
''''        End If
''''    End If
''''    If sw_mant_ayuda = True Then
''''        SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Visible = False
''''        SSActiveToolBars1.Tools.ITEM("ID_Salir").Visible = False
''''    End If
End Sub
Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
   
   If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
        End If
    End If
    
End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F7CODMED").value & "")) = 0 Then
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
    End If

End Sub

Private Sub CONFIGURA_GRID()

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
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
    With dxDBGrid1.Columns.ColumnByFieldName("F7CODMED").LookupColumn
        .LookupDataset.ADODataset.ConnectionString = cnn_dbbancos
        .LookupDataset.ADODataset.CommandText = "SELECT F7CODMED,F7NOMMED From EF7MEDIDAS ORDER BY F7CODMED;"
        .LookupKeyField = "F7CODMED"
        .LookupResultField = "F7NOMMED"
        .LookupDataset.Active = True
        .ListFieldIndex = 0
        .DisplaySize = 15
        .LookupCache = True
        .ListFieldName = "F7NOMMED"
        .ListWidth = 20
    End With
    
    AbrirBasesTemp
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnTP
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
End Sub

Private Sub AbrirBasesTemp()
    If cnTP.State = 1 Then cnTP.Close
    cnTP.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\templus.mdb;Persist Security Info=False"
End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
    AbrirBasesTemp
    dxDBGrid1.Dataset.Active = False
'    If sw_nuevo_documento = True Then
        DELETEREC_LOG "tmpMedVenta", cnTP
        dxDBGrid1.Dataset.Refresh
'    End If

    AbrirBasesTemp
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnTP
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
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
            .FieldValues("ITEM") = 1
            .FieldValues("F7CODMED") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("FACTOR") = Null
            .FieldValues("F5PREVTA") = Null
            
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
End Sub


Private Sub Form_Load()
Dim wcampo  As String
Dim i       As Integer

    Me.MousePointer = vbHourglass
    Me.Height = 7800
    Me.Width = 9600
    Me.left = 1605
    Me.top = 1065
    
    sw_ayuda_um = False
    sw_ayuda_alma = False
    
    Set cnn = New ADODB.Connection
    conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\db_tabla.mdb;Persist Security Info=False"
    cnn.Open conexion
    
    sw_activate = True
        
    CargarMoneda
    CargarTipo
    
    If sw_load_mant = True Then
        SSActiveToolBars1.Tools.ITEM("ID_Lista").Visible = False
        SSActiveToolBars1.Tools.ITEM("ID_Salir").Visible = True
    Else
        SSActiveToolBars1.Tools.ITEM("ID_Lista").Visible = True
        SSActiveToolBars1.Tools.ITEM("ID_Salir").Visible = False
    End If
    
        
    CONFIGURA_GRID
    
    If sw_nuevo_doc = True Then
        Nuevo_Producto
        LblNivel(0).Visible = True
        Panelcod.Visible = True
        List1(0).Visible = True
        AdicionaItem
    Else
        Actualiza_Producto lista_prod.dxDBGrid1.Columns(0).value
    End If
    'SF1PARAIN ESTA BDCONTROL
    If cnn_control.State = adStateOpen Then cnn_control.Close
    cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CONTROL.MDB" & ";Persist Security Info=False"
    cnn_control.Open cconex_control
    sql = "select * from sf1parain where F1CODEMP = '" & wempresa & "'"
    If rstparametro.State = 1 Then rstparametro.Close
    rstparametro.Open sql, cnn_control, adOpenDynamic, adLockOptimistic
    If Not rstparametro.EOF Then
        wlong1 = rstparametro.Fields("F1lonniv1") + rstparametro.Fields("F1lonniv2") + rstparametro.Fields("F1lonniv3") + rstparametro.Fields("F1lonniv4") + rstparametro.Fields("F1lonniv5")
        wlong2 = rstparametro.Fields("F1loncod")
        wniveles = rstparametro.Fields("F1niveles")
        For i = 1 To rstparametro.Fields("F1Niveles")
            LblNivel(i - 1).Caption = rstparametro.Fields("F1nivel0" & Format(i, "0")) & " (" & rstparametro.Fields("F1lonniv" & Format(i, "0")) & ")"
            wcod(i) = ""
            wdes(i) = ""
        Next i%
    End If
    
    sql = "SELECT * FROM SF7NIVEL01 ORDER BY F7CODCON"
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cnn_dbbancos.Open "Provider =Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB;Persist Security Info=False"
    If rstconsulta.State = 1 Then rstconsulta.Close
    rstconsulta.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rstconsulta.EOF Then
        Do While Not rstconsulta.EOF
            List1(0).AddItem rstconsulta.Fields("F7codcon") & " - " & rstconsulta.Fields("F7descon")
            rstconsulta.MoveNext
        Loop
        List1(0).ListIndex = 0
    End If
    Panelcod.top = 900
    Panelcod.left = 270
    
    wgraba = "0"
    If wf1uupp = "*" Then
        lblpiezas.Visible = True
        txtpiezas.Visible = True
    Else
        lblpiezas.Visible = False
        txtpiezas.Visible = False
    End If
    
    swmant_prod_especiales = False
    
    sql = "select * from ef2users_der where f2coduser='" & wusuario & "' and codigo = '005' order by codigo"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If rst.EOF Then
        MsgBox "Usted no puede acceder al Mantenimiento de Productos", 16, "Atención"
        Exit Sub
    AsignaDerechos
    End If
    Me.MousePointer = vbDefault
    dxDBGrid1.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'rstparametro.Close
    rstconsulta.Close
    cnn.Close
    
End Sub

Public Sub Graba_Producto()
Dim FlagAdd As Integer
Dim amovs(0 To 46)  As a_grabacion
Dim nstock          As Double
Dim ncostoini       As Double
Dim ncospro         As Double
Dim ncosprod        As Double
Dim ntotals         As Double
Dim ntotald         As Double
Dim ccampo1         As String
Dim ccampo2         As String
    
    If Len(Trim(Txtcodalm.Text)) = 0 Then
    '    MsgBox "Debe Ingresar el Almacén.", 16, "Atención"
    '    If vaTabPro1.ActiveTab = 0 Then
    '        Txtcodalm.SetFocus
    '    End If
    '    Exit Sub
    Txtcodalm.Text = "01"
    End If
    If Len(Trim(Txtcodmedc.Text)) = 0 Then
        MsgBox "Debe Ingresar La Unidad de Medida del producto", 16, "Atención"
'        If vaTabPro1.ActiveTab = 0 Then
           Txtcodmedc.SetFocus
           sw_mant_ayuda = False
'        End If
        Exit Sub
    End If
    'If Len(Trim(txtcodfab.Text)) = 0 Then
        'MsgBox "Debe Ingresar El Código de Fabricante.", vbInformation, "Atención"
        'If vaTabPro1.ActiveTab = 0 Then
        '    txtcodfab.SetFocus
        'End If
        'Exit Sub
    'End If
'    If Len(Trim(txcodmarca.Text)) = 0 Then
'        MsgBox "Debe Ingresar La Marca del Producto.", 16, "Atención"
''        If vaTabPro1.ActiveTab = 0 Then
''            txcodmarca.SetFocus
''        End If
'        Exit Sub
'    End If
    
'    If Val(fob.Text) = 0 Then
'        MsgBox "No ha Ingresado el Fob del Producto", vbInformation, "Atención"
'    End If
    
    '----------------------------------------------------------------------
'    If wf1mant_productos = "*" Then
'        If Val("" & fob.Text) = 0 Or Val("" & txtfactor.Text) = 0 Or Val("" & txtvalvta.Text) = 0 Or Val("" & txtigv.Text) = 0 Or Val("" & txtprevta.Text) = 0 Or Val("" & txtvalvta.Text) > (Val("" & fob.Text) * Val("" & txtfactor.Text)) Then
'            If Val("" & txtvalvta.Text) > (Val("" & fob.Text) * Val("" & txtfactor.Text)) And swmant_prod_especiales = True Then
'            Else
'                MsgBox "Los campos  Factor, FOB, ValVta, IGV ó PVenta deben ser ingresados...., Y el ValVta no debe ser mayor al FOB * Factor ... ", vbInformation + vbOKOnly, "Atención"
'                Exit Sub
'            End If
'        End If
'    End If
    '----------------------------------------------------------------------

    wcodigo = "" & Trim(txtcodpro.Text)
    sql = "SELECT F5CODPRO FROM IF5PLA WHERE F5CODPRO = '" & Trim(txtcodpro.Text) & "'"
    If rst.State = 1 Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        FlagAdd = 0
    Else
        FlagAdd = 1
    End If

    If FlagAdd = 1 Then
        'Valida Si Producto ya fue registrado
        wcodfab = txtCodFab.Text
        wcodmarca = txcodmarca.Text
'        If rst.State = adStateOpen Then rst.Close
'        sql = "select f5codfab from if5pla where f5codfab='" & wcodfab & "' and f5marca='" & wcodmarca & "'"
'        rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
'        If Not rst.EOF Then     'El Producto ya existe
'            MsgBox "Cod. Fab.: " & wcodfab & Chr(13) & "Marca: " & txmarca.Caption & Chr(13) & Chr(13) & "El Producto ya Existe. No Podrá Realizar la Actualización", vbInformation, "Sistema de Logistica"
'            rst.Close
'            Exit Sub
'        End If
'        rst.Close
    End If

    If optfactor(0).value = True Then
        wtipo = "N"
    Else
        wtipo = "I"
    End If

   
    amovs(0).campo = "F5CODPRO": amovs(0).valor = txtcodpro.Text: amovs(0).TIPO = "T"
    amovs(1).campo = "F5NOMPRO": amovs(1).valor = Txtnompro1.Text: amovs(1).TIPO = "T"
    amovs(2).campo = "F5FECING": amovs(2).valor = TXTFECING.value: amovs(2).TIPO = "F"
    amovs(3).campo = "F5CODFAB": amovs(3).valor = txtCodFab.Text: amovs(3).TIPO = "T"
    amovs(4).campo = "F7CODMED": amovs(4).valor = Txtcodmedc.Text: amovs(4).TIPO = "T"
    amovs(5).campo = "F5UBICACIO": amovs(5).valor = ubicacion.Text: amovs(5).TIPO = "T"
    amovs(6).campo = "F5MARCA": amovs(6).valor = IIf(Len("" & Trim(txcodmarca.Text)) = 0, " ", "" & Trim(txcodmarca.Text)): amovs(6).TIPO = "T"
    Select Case wtippro
        Case "A"
            amovs(7).campo = "F5TIPPRO": amovs(7).valor = "0": amovs(7).TIPO = "T"
        Case Else
            amovs(7).campo = "F5TIPPRO": amovs(7).valor = "1": amovs(7).TIPO = "T"
    End Select
    amovs(8).campo = "F5ubica2": amovs(8).valor = Trim("" & TxtCodubi2.Text): amovs(8).TIPO = "T"
    amovs(9).campo = "F5TARIFA": amovs(9).valor = 0#: amovs(9).TIPO = "N"
    amovs(10).campo = "F5insumo": amovs(10).valor = IIf(chkinsumo.value = True, "*", " "): amovs(10).TIPO = "T"
    amovs(11).campo = "F5MODELO": amovs(11).valor = Trim("" & txtCodFab.Text): amovs(11).TIPO = "T"
    amovs(12).campo = "F5STOCKLOG": amovs(12).valor = 0#: amovs(12).TIPO = "N"
    amovs(13).campo = "F5MONEDA": amovs(13).valor = IIf(CmbMoneda.ListIndex = 0, "S", "D"): amovs(13).TIPO = "T"
    amovs(14).campo = "F5VALVTA": amovs(14).valor = CDbl(txtvalvta.Text): amovs(14).TIPO = "N"
    amovs(15).campo = "F5IGVVTA": amovs(15).valor = CDbl(TxtIgv.Text): amovs(15).TIPO = "N"
    amovs(16).campo = "F5fob": amovs(16).valor = CDbl(fob.Text): amovs(16).TIPO = "N"
    amovs(17).campo = "F5PREVTA": amovs(17).valor = CDbl(txtprevta.Text): amovs(17).TIPO = "N"
    amovs(18).campo = "F5partara": amovs(18).valor = IIf(Len("" & Trim(arancelaria.Text)) = 0, "", "" & Trim(arancelaria.Text)): amovs(18).TIPO = "T"
    amovs(19).campo = "F5stockmin": amovs(19).valor = Val(Format(minimo.Text, "0.00")): amovs(19).TIPO = "N"
    amovs(20).campo = "F5stockmax": amovs(20).valor = Val(Format(maximo.Text, "0.00")): amovs(20).TIPO = "N"
    amovs(21).campo = "F5ctacon": amovs(21).valor = IIf(Len(Trim("" & cuenta.Text)) = 0, "", Trim("" & cuenta.Text)): amovs(21).TIPO = "T"
    amovs(22).campo = "F5TEXTO": amovs(22).valor = Trim("" & txtnompro3.Text): amovs(22).TIPO = "T"
    amovs(23).campo = "F5TIPESTADO": amovs(23).valor = "1": amovs(23).TIPO = "T"
    amovs(24).campo = "F5FACTOR": amovs(24).valor = txtfactor.Text: amovs(24).TIPO = "N"
    amovs(25).campo = "F5TIPO": amovs(25).valor = IIf(opttipo(0).value = True, "P", "S"): amovs(25).TIPO = "T"
    amovs(26).campo = "F5AFECTO": amovs(26).valor = IIf(chkafecto.value = True, "*", " "): amovs(26).TIPO = "T"
    amovs(27).campo = "F3GASTO": amovs(27).valor = txtgasto.Text: amovs(27).TIPO = "T"
    amovs(28).campo = "F5PIEZAS": amovs(28).valor = Val(txtpiezas.Text & ""): amovs(28).TIPO = "N"
    amovs(29).campo = "F5ESTPRO": amovs(29).valor = gestpro: amovs(29).TIPO = "T"
    amovs(30).campo = "F5ESTVAL": amovs(30).valor = gestval: amovs(30).TIPO = "T"
    amovs(31).campo = "F5NOMPRO2": amovs(31).valor = txtnompro2: amovs(31).TIPO = "T"
    amovs(32).campo = "F5TEXTO_ING": amovs(32).valor = txtnompro4.Text: amovs(32).TIPO = "T"
    
    If FlagAdd = 1 Then
        '-------------- ACTUALIZA IF5PLA
        amovs(33).campo = "F5FECMOD": amovs(33).valor = TXTFECING.value: amovs(33).TIPO = "F"
        amovs(34).campo = "F5USERMOD": amovs(34).valor = "": amovs(34).TIPO = "T"
    Else
        amovs(33).campo = "F5FECMOD": amovs(33).valor = Format(Date, "dd/mm/yyyy"): amovs(33).TIPO = "F"
        amovs(34).campo = "F5USERMOD": amovs(34).valor = wusuario: amovs(34).TIPO = "T"
    End If
    amovs(35).campo = "F5COSTOEURO": amovs(35).valor = CDbl(txtCostoEuros.Text): amovs(35).TIPO = "N"
    amovs(36).campo = "F5TCEURO": amovs(36).valor = CDbl(txttceuros.Text): amovs(36).TIPO = "N"
    amovs(37).campo = "F5TIPOCOSTO": amovs(37).valor = cmbtipo.Text: amovs(37).TIPO = "T"
    amovs(38).campo = "F5DESCONTINUADO": amovs(38).valor = IIf(chkDescontinuado.value, "S", "N"): amovs(38).TIPO = "T"
    amovs(39).campo = "F5ctacon1": amovs(39).valor = IIf(Len(Trim("" & cuenta1.Text)) = 0, "", Trim("" & cuenta1.Text)): amovs(39).TIPO = "T"
    amovs(40).campo = "F5ctacon2": amovs(40).valor = IIf(Len(Trim("" & cuenta2.Text)) = 0, "", Trim("" & cuenta2.Text)): amovs(40).TIPO = "T"
    amovs(41).campo = "F5MONEDAORI": amovs(41).valor = left(CmbMoneda.Text, 1): amovs(41).TIPO = "T"
    amovs(42).campo = "F5FLETE": amovs(42).valor = CDbl(txtflete.Text): amovs(42).TIPO = "N"
    amovs(43).campo = "F5INSTALACION": amovs(43).valor = CDbl(txtinstalacion.Text): amovs(43).TIPO = "N"
    amovs(44).campo = "F5MANIPULEO": amovs(44).valor = CDbl(txtmanipuleo.Text): amovs(44).TIPO = "N"
    amovs(45).campo = "F5OTROSCOSTOS": amovs(45).valor = CDbl(txtotroscostos.Text): amovs(45).TIPO = "N"
    amovs(46).campo = "F5TREP": amovs(46).valor = txtF5TREP.Text: amovs(46).TIPO = "N"
    
    If FlagAdd = 1 Then
        GRABA_REGISTRO_logistica amovs(), "IF5PLA", "A", 46, cnn_dbbancos, ""
        codProdtem = txtcodpro.Text
        grabarMedVentas
    Else
        GRABA_REGISTRO_logistica amovs(), "IF5PLA", "M", 46, cnn_dbbancos, "F5CODPRO = '" & txtcodpro.Text & "'"
        grabarMedVentas True
    End If

    '-------------- ACTUALIZA IF6ALMA
'    sql = "SELECT F5CODPRO FROM IF6ALMA WHERE f2codalm = '" & Trim(Txtcodalm.Text) & "' and F5CODPRO = '" & txtcodpro.Text & "' "
'    If rsfield.State = 1 Then rsfield.Close
'    rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    If rsfield.EOF Then
'        ncostoini = Val(Format(TxtCostoIni, "0.00"))
'        ccampo1 = "F5ING" & Format(Month(TXTFECING.Value), "00")
'        ccampo2 = "F5INGD" & Format(Month(TXTFECING.Value), "00")
'        If CmbMoneda.ListIndex = 0 Then
'            ncospro = Val(Format(TxtCostoIni, "0.00"))
'            ntotals = Val(Format(txtcostototal, "0.00"))
'            If Val(Format(txttc, "0.000")) > 0# Then
'                ncosprod = Val(Format(ncospro / Val(Format(txttc, "0.000")), "0.00"))
'                ntotald = Val(Format(ntotals / Val(Format(txttc, "0.000")), "0.00"))
'            Else
'                ncosprod = 0#
'                ntotald = 0#
'            End If
'        Else
'            ncosprod = Val(Format(TxtCostoIni, "0.00"))
'            ncospro = Val(Format(ncosprod * Val(Format(txttc, "0.000")), "0.00"))
'            ntotald = Val(Format(txtcostototal, "0.00"))
'            ntotals = Val(Format(ntotald * Val(Format(txttc, "0.000")), "0.00"))
'        End If
'        sql = "INSERT INTO IF6ALMA (F2CODALM,F5CODPRO,f5cospro,f5cosprod,F6STOCKACT,F5DEBM00,F5COSTOINI," & ccampo1 & "," & ccampo2 & ") " & _
'              " VALUES ('" & Trim(Txtcodalm.Text) & "','" & txtcodpro.Text & "'," & ncospro & "," & ncosprod & "," & _
'              nstock & "," & nstock & "," & ncostoini & "," & ntotals & "," & ntotald & ")"
'              cnn_dbbancos.Execute sql
              'AlmacenaQuery_sql sql, cnn_dbbancos
'    End If
    'rsfield.Close
    
    txtcodpro.Enabled = False
    TXTFECING.Enabled = False
    wgraba = "0"
    sw_mant_ayuda = True
    If sw_GRABA_REGISTRO_logistica Then
        MsgBox "El Producto se actualizó", vbInformation, "Sistema de Logistica"
    End If
    
    Exit Sub

End Sub

Private Sub grabarMedVentas(Optional indMod As Boolean)
Dim cSqlTMV     As String
Dim ff          As Long
        
    If indMod Then
        sql = "Delete from MEDIVENTAS where F5CODPRO = '" & txtcodpro.Text & "'"
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
    End If
    For ff = 0 To 9999999
    Next
    AbrirBasesTemp
    If rsTMV.State = 1 Then rsTMV.Close
    rsTMV.Open "Select * from tmpMedVenta where not isnull(F7CODMED)", cnTP, adOpenDynamic, adLockOptimistic
    If Not rsTMV.EOF Then
        'rsTMV.Requery
        rsTMV.MoveFirst
        Do While Not rsTMV.EOF
        If IsNull(rsTMV.Fields("FACTOR")) Then rsTMV.Fields("FACTOR") = 0
        If IsNull(rsTMV.Fields("F5PREVTA")) Then rsTMV.Fields("F5PREVTA") = 0
               
         cSqlTMV = "insert into MEDIVENTAS (F5CODPRO,F7CODMED, F5FACTOR,F5PREVTA) values ('" & _
                        txtcodpro.Text & "','" & rsTMV.Fields("F7CODMED") & "'," & rsTMV.Fields("FACTOR") & "," & rsTMV.Fields("F5PREVTA") & ")"
            cnn_dbbancos.Execute cSqlTMV
            'AlmacenaQuery_sql cSqlTMV, cnn_dbbancos
            
            rsTMV.MoveNext
        Loop
    End If
    
End Sub

Private Sub List1_DblClick(Index As Integer)
    
    List1_Keypress Index, 13

End Sub

Private Sub List1_Keypress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        wcod(Index + 1) = left(List1(Index).Text, rstparametro.Fields("F1LONNIV" & Format(Index + 1, "0")))
        wdes(Index + 1) = Trim(UCase(right(List1(Index).Text, Len(Trim(List1(Index).Text)) - (rstparametro.Fields("F1LONNIV" & Format(Index + 1, "0")) + 2))))
        txtcodpro.Text = Trim(wcod(1)) & Trim(wcod(2)) & Trim(wcod(3)) & Trim(wcod(4)) & Trim(wcod(5))
        If Index + 1 = wniveles Then
            calcula_codigo
            Cierra_Codigo
        Else
            Select Case Index
            Case Is = 0: sql = "Select * from SF7NIVEL02 Where F7nivel01='" + wcod(1) + "'"
            Case Is = 1: sql = "Select * from SF7NIVEL03 Where F7nivel01='" + wcod(1) + "' and F7nivel02='" + wcod(2) + "'"
            Case Is = 2: sql = "Select * from SF7NIVEL04 Where F7nivel01='" + wcod(1) + "' and F7nivel02='" + wcod(2) + "' and F7nivel03='" + wcod(3) + "'"
            Case Is = 3: sql = "Select * from SF7NIVEL05 Where F7nivel01='" + wcod(1) + "' and F7nivel02='" + wcod(2) + "' and F7nivel03='" + wcod(3) + "'and F7nivel04='" + wcod(4) + "'"
            End Select
            List1(Index + 1).Clear
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
                Do While Not rsfield.EOF
                    List1(Index + 1).AddItem rsfield.Fields("F7codcon") & " " & rsfield.Fields("F7descon")
                    rsfield.MoveNext
                Loop
                LblNivel(Index + 1).Visible = True
                List1(Index + 1).Visible = True
                List1(Index + 1).ListIndex = 0
                List1(Index + 1).SetFocus
            Else
                'MsgBox "No se registraron más niveles de descripción", 64, "Mantenimiento de Productos"
                BtnFin_Click
            End If
'            rsfield.Close
        End If
    End If
    If KeyAscii = 27 Then
        wcod(Index + 1) = ""
        wdes(Index + 1) = ""
        txtcodpro.Text = Trim(wcod(1)) & Trim(wcod(2)) & Trim(wcod(3)) & Trim(wcod(4)) & Trim(wcod(5))
        If Index = 0 Then
            Cierra_Codigo
        Else
            LblNivel(Index).Visible = False
            List1(Index).Visible = False
            List1(Index - 1).SetFocus
        End If
    End If

End Sub

Private Sub maximo_GotFocus()
    
    maximo.SelStart = 0: maximo.SelLength = Len(maximo.Text)
    
End Sub

Private Sub maximo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        minimo.SetFocus
    End If

End Sub

Private Sub minimo_GotFocus()
    
    minimo.SelStart = 0: minimo.SelLength = Len(minimo.Text)
    
End Sub

Private Sub minimo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtF5TREP.SetFocus
    End If
End Sub

Public Sub Nuevo_Producto()

    TXTFECING.value = Format(Date, "dd/mm/yyyy")
    txtcodpro.Enabled = True
    TXTFECING.Enabled = True
    txtcodpro.Text = ""
    Txtnompro1.Text = ""
    txtnompro2.Text = ""
    txcodmarca.Text = ""
    txtCodFab.Text = ""
    Txtcodmedc.Text = ""
    Txtnommedc = ""
    txtvalvta.Text = "0.00"
    TxtIgv.Text = "0.00"
    txtprevta.Text = "0.0"
    fob.Text = "0.00"
    maximo.Text = "0.00"
    minimo.Text = "0.00"
    ubicacion.Text = ""
    cuenta.Text = ""
    arancelaria.Text = ""
    txmarca.Caption = ""
    txtfactor.Text = "0.00"
    chkinsumo.value = False
    chkafecto.value = True
    txtgasto.Text = ""
    pnlgasto.Caption = ""
    txtpiezas.Text = ""
    Txtcodalm.Text = ""
    Txtnomalm.Caption = ""
    opttipo(0).value = True
    txtflete.Text = "0.00"
    txtinstalacion.Text = "0.00"
    '---------------
    gestpro = ""
    gestval = ""
    '---------------
    txtnompro3.Text = ""
    txtnompro4.Text = ""
    txtmanipuleo.Text = "0.00"
    txtotroscostos.Text = "0.00"
    AdicionaItem
    
    Frame1.Enabled = False
    
End Sub

Private Sub optfactor_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        arancelaria.SetFocus
    End If

End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        arancelaria.SetFocus
    End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

Dim msgdev As String

    Select Case Tool.Id
        Case "ID_Nuevo":
            If wgraba = "1" Then
                If MsgBox("No ha grabado el registro. Desea Grabarlo ahora", 36, "Atención") = 6 Then
                    If Len(Trim(txtcodpro.Text)) > 0 Then
                        Me.MousePointer = vbHourglass
                        Graba_Producto
                        Me.MousePointer = vbDefault
                    End If
                End If
            End If
'            If vaTabPro1.ActiveTab = 0 Then
                Nuevo_Producto
                LblNivel(0).Visible = True
                Panelcod.Visible = True
                List1(0).Visible = True
                List1(0).SetFocus
                wgraba = "1"
'            Else
'                MsgBox "Para ingresar un nuevo producto, tiene que encontrarse en Datos Generales.", vbInformation, "Atención"
'            End If
        Case "ID_Grabar":
            Me.MousePointer = vbHourglass
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                'dxDBGrid1.Dataset.Edit
                sw_nuevo_item = True
                dxDBGrid1.Dataset.Post
                sw_nuevo_item = False
                'dxDBGrid1.Dataset.Refresh
            End If
            If Trim(txtCodFab.Text) <> "" Then
                sql = "SELECT F5CODPRO FROM IF5PLA WHERE F5CODPRO = '" & Trim(txtCodFab) & "' "
                If rsfield.State = 1 Then rsfield.Close
                rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsfield.EOF Then
                    If Trim(txtcodpro.Text) <> Trim(rsfield.Fields("f5codpro")) Then
                        MsgBox "Código de Fabricante ya existe. Producto : " & rsfield.Fields("f5codpro"), 16, "Sistema de Logistica"
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If
            If Len(Trim(txtcodpro.Text)) > 0 Then
                Me.MousePointer = vbHourglass
                Graba_Producto
                If sw_mant_ayuda = True Then
                    wcodproducto = txtcodpro.Text
                    wdesproducto = Txtnompro1.Text
                    wmedida = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", Txtcodmedc.Text & "", "T", cnn_dbbancos)
                    sw_mant_ayuda = False
                    Me.MousePointer = vbDefault
                    Unload Me
                    Unload ayuda_productos
                End If

                Me.MousePointer = vbDefault
            End If
            Me.MousePointer = vbDefault
        Case "ID_Eliminar":
            Me.MousePointer = vbHourglass
            Elimina_Producto
            Me.MousePointer = vbDefault
'        Case "ID_Imprimir":
        Case "ID_Salir":
            Unload Me
        Case "ID_Lista":
            'lista_prod.Adoprod.Refresh
            Unload Me
    End Select
    
End Sub


Private Sub txcodmarca_DblClick()
    
    Txcodmarca_KeyDown 113, 0
    
End Sub

Private Sub Txcodmarca_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        wcodmar = ""
        If optfactor(0).value = True Then
            sw_ayuda_marca = "N"
        Else
            sw_ayuda_marca = "I"
        End If
        sw_ayuda = True
        'hlp_marcas.Show 1
        ayuda_marcas.Show 1
        sw_ayuda = False
        sw_ayuda_marca = " "
        If Len(Trim(wcodmar)) > 0 Then
            txcodmarca.Text = wcodmar
            txcodmarca_keypress 13
        End If
    End If

End Sub

Private Sub txcodmarca_keypress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub

Private Sub txcodmarca_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txcodmarca.Text)) > 0 Then
            sql = "select F2DESMAR from EF2MARCAS where F2CODMAR = '" & txcodmarca.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
                txmarca.Caption = Trim("" & rsfield.Fields("F2DESMAR"))
            Else
                txcodmarca.Text = ""
                txmarca.Caption = ""
                MsgBox "Código de la marca no existe. Verifique.", 16, "Atención"
                txcodmarca.SetFocus
            End If
            rsfield.Close
        End If
    End If
    
End Sub

Private Sub Txtcodalm_Change()
'txtnomalm.Caption
End Sub

Private Sub Txtcodalm_DblClick()
    
    Txtcodalm_KeyDown 113, 0
    
End Sub

Private Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Me.MousePointer = vbHourglass
        wcod_alm = "" & Txtcodalm.Text
        sw_ayuda_alma = True
        'hlp_almacenes.Show 1
        ayuda_almacen.Show 1
        sw_ayuda_alma = False
        Txtcodalm.Text = wcod_alm
        Me.MousePointer = vbDefault
        Txtcodalm_KeyPress 13
    End If

End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtcodalm_LostFocus()

    If sw_ayuda_alma = False Then
        If Len(Trim(Txtcodalm.Text)) > 0 Then
            Txtcodalm.Text = Format(Txtcodalm.Text, "00")
            sql = "select f2codalm,f2nomalm from ef2almacenes where f2codalm = '" & Txtcodalm.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsfield.EOF Then
                gcodalm = rsfield.Fields("F2codalm")
                Txtnomalm.Caption = rsfield.Fields("F2NOMalm") & ""
                chkDescontinuado.Enabled = True
            Else
                Txtnomalm.Caption = ""
                Txtcodalm.Text = ""
                chkDescontinuado.Enabled = False
                MsgBox "Código del almacén no existe. Verifique.", 16, "Atención"
                Txtcodalm.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub Txtcodfab_GotFocus()

    txtCodFab.SelStart = 0: txtCodFab.SelLength = Len(txtCodFab.Text)
    
End Sub

Private Sub txtCodFab_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    
End Sub

Private Sub txtcodfab_LostFocus()

    If Len(Trim(txtCodFab.Text)) > 0 Then
        sql = "SELECT F5CODPRO FROM IF5PLA WHERE F5CODFAB = '" & txtCodFab.Text & "' "
        If rsfield.State = 1 Then rsfield.Close
        rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsfield.EOF Then
            MsgBox "El código de fabricante ya ha sido asignado al producto " & rsfield.Fields("F5CODPRO") & "", 64, "Atención"
        End If
    End If

End Sub

Private Sub Txtcodmedc_DblClick()
    
    Txtcodmedc_KeyDown 113, 0
    
End Sub

Private Sub Txtcodmedc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        wcodmed = ""
        sw_ayuda_um = True
        'hlp_medidas.Show 1
        ayuda_unidades.Show 1
        sw_ayuda_um = False
        
        If Len(Trim(wcodmed)) > 0 Then
            Txtcodmedc.Text = wcodmed
            Txtnommedc.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7NOMMED", "EF7MEDIDAS", "F7CODMED", Trim(Txtcodmedc.Text), "T")
            
            Txtcodmedc_KeyPress 13
        End If
    End If

End Sub

Private Sub Txtcodmedc_KeyPress(KeyAscii As Integer)
'    On Error Resume Next
'
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
'    End If
    Select Case KeyAscii
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Txtcodmedc_LostFocus()

'    If sw_ayuda_um = False Then
'        If Len(Trim(Txtcodmedc.Text)) > 0 Then
'            Txtcodmedc.Text = Trim(Txtcodmedc.Text)
'            sql = "select F7NOMMED  from EF7MEDIDAS where F7CODMED = '" & Txtcodmedc.Text & "'"
'            If rsfield.State = 1 Then rsfield.Close
'            rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not rsfield.EOF Then
'                Txtnommedc.Caption = rsfield.Fields("F7NOMMED") & ""
'            Else
'                Txtcodmedc.Text = ""
'                Txtnommedc.Caption = ""
'                MsgBox "Código de medida no existe. Verifique.", 16, "Atención"
'                Txtcodmedc.SetFocus
'            End If
'        End If
'    End If
    Txtcodmedc.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODMED", "EF7MEDIDAS", "F7CODMED", Trim(Txtcodmedc.Text), "T")
    
    If Trim(Txtcodmedc.Text) <> vbNullString Then
        Txtnommedc.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7NOMMED", "EF7MEDIDAS", "F7CODMED", Trim(Txtcodmedc.Text), "T")
    Else
        Txtnommedc.Caption = vbNullString
        
        MsgBox "Código de medida no existe. Verifique.", vbInformation + vbOKOnly, App.ProductName
        
        'Txtcodmedc.SetFocus
    End If
End Sub

Private Sub Txtcodpro_Keypress(KeyAscii As Integer)
Dim i       As Integer

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        For i = 1 To wniveles
            LblNivel(i - 1).Visible = False
            List1(i - 1).Visible = False
        Next
        txtcodpro.Text = Trim(txtcodpro.Text)
        sql = "SELECT F5CODPRO FROM IF5PLA WHERE F5CODPRO = '" & Trim(txtcodpro) & " '"
        If RsProducto.State = 1 Then RsProducto.Close
        RsProducto.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProducto.EOF Then
            '--------------------------------------------------
            CodFirmaAprobacion(1) = ""
'            aprobacion.Show vbModal
            'If Len(Trim$(CodFirmaAprobacion(1))) > 0 Then
                Actualiza_Producto txtcodpro.Text
            'Else
            '    MsgBox "No se puede modificar los datos del producto.", vbCritical, "Atención"
            'End If
            '--------------------------------------------------
        Else
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        End If
        If RsProducto.State = 1 Then RsProducto.Close
    End If

End Sub

Private Sub txtcodpro_LostFocus()
    If Len(Trim(txtCodFab.Text & "")) = 0 Then
        txtCodFab.Text = txtcodpro.Text
    End If
End Sub

Private Sub TxtCodubi2_GotFocus()
    
    TxtCodubi2.SelStart = 0: TxtCodubi2.SelLength = Len(TxtCodubi2.Text)
    
End Sub

Private Sub TxtCodUbi2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtCodubi2.Text = Format(TxtCodubi2.Text, "0000")
        TxtCodubi2.Text = UCase(Trim(TxtCodubi2.Text))
        sql = "select f2desubi from ef2ubica where f2codubi= '" & TxtCodubi2.Text & "'"
        If rsfield.State = 1 Then rsfield.Close
        rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsfield.EOF Then
            TxtUbica2.Caption = "" & Mid(rsfield.Fields("F2DESUBI"), 1, 26)
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        Else
            Beep
        End If
        rsfield.Close
    End If

End Sub

Private Sub txtCostoEuros_GotFocus()
    
    txtCostoEuros.SelStart = 0: txtCostoEuros.SelLength = Len(txtCostoEuros.Text)
    
End Sub

Private Sub txtCostoEuros_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtCostoEuros_LostFocus()

    If IsNumeric(txtCostoEuros.Text) Then
        wcostoeuro = CDbl(txtCostoEuros.Text)
    Else
        wcostoeuro = 0
    End If
    
    If IsNumeric(txttceuros.Text) Then
        wtc = CDbl(txttceuros.Text)
    Else
        wtc = 0
    End If
    
    fob.Text = Format(txtCostoEuros.Text * Val(txttceuros.Text), "#,###,###0.00")

    CALCULA_IMPORTES
    
End Sub
Private Sub txtF5TREP_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    
End Sub

Private Sub txtflete_GotFocus()

    txtflete.SelStart = 0: txtflete.SelLength = Len(txtflete.Text)
    
End Sub

Private Sub txtflete_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtflete_LostFocus()

    txtflete.Text = Format(txtflete.Text, "###,###,##0.000")
    CALCULA_IMPORTES

End Sub

Private Sub txtinstalacion_GotFocus()

    txtinstalacion.SelStart = 0: txtinstalacion.SelLength = Len(txtinstalacion.Text)
    
End Sub

Private Sub txtinstalacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtinstalacion_LostFocus()

    txtinstalacion.Text = Format(txtinstalacion.Text, "###,###,##0.000")
    CALCULA_IMPORTES
    
End Sub

Private Sub txtmanipuleo_GotFocus()
    
    txtmanipuleo.SelStart = 0: txtmanipuleo.SelLength = Len(txtmanipuleo.Text)
    
End Sub

Private Sub txtmanipuleo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtmanipuleo_LostFocus()

    txtmanipuleo.Text = Format(txtmanipuleo.Text, "###,###,##0.000")
    CALCULA_IMPORTES

End Sub

Private Sub txtnompro2_GotFocus()
  
    txtnompro2.SelStart = 0: txtnompro2.SelLength = Len(txtnompro2.Text)

End Sub

Private Sub txtnompro2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub

Private Sub txtfactor_GotFocus()
    
    txtfactor.SelStart = 0: txtfactor.SelLength = Len(txtfactor.Text)
    
End Sub

Private Sub txtfactor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtfactor_LostFocus()

    txtfactor.Text = Format(txtfactor.Text, "0.000")
    If Val(Format(txtfactor.Text, "0.000")) > 0 Then
        CALCULA_IMPORTES
    End If

End Sub

Private Sub Txtfecing_GotFocus()
    
   ' txtfecing.FocusSelect = True
    
End Sub

Private Sub Txtfecing_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       Txtnompro1.SetFocus
    End If

End Sub

Private Sub txtgasto_DblClick()

    txtgasto_KeyDown 113, 0
    
End Sub

Private Sub txtgasto_GotFocus()

    txtgasto.SelStart = 0: txtgasto.SelLength = Len(txtgasto.Text)

End Sub

Private Sub txtgasto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        llampro = 0
        wcodgasto = ""
        sw_ayuda = True
        hlp_gastos.Show 1
        sw_ayuda = False
        If Len(Trim(wcodgasto)) > 0 Then
            txtgasto.Text = wcodgasto
            pnlgasto.Caption = wnomgasto
            cuenta.Text = wctagasto
        End If
    End If

End Sub

Private Sub txtgasto_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtgasto.Text)) > 0 Then
            If rsgastos.State = adStateOpen Then rsgastos.Close
            rsgastos.Open "SELECT NOMBRE,CUENTA FROM BF9GIN WHERE BASE='G' AND CODIGO='" & txtgasto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsgastos.EOF Then
                pnlgasto.Caption = Trim(rsgastos.Fields("NOMBRE") & "")
                cuenta.Text = Trim(rsgastos.Fields("CUENTA") & "")
            Else
                MsgBox "Código del gasto no existe. Verifique.", 16, "Atención"
            End If
            rsgastos.Close
        End If
    End If

End Sub

Private Sub chkinsumo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkafecto.SetFocus
    End If

End Sub

Private Sub TxtIgv_GotFocus()

    TxtIgv.SelStart = 0: TxtIgv.SelLength = Len(TxtIgv.Text)

End Sub

Private Sub Txtigv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtnompro1_GotFocus()

    Txtnompro1.SelStart = 0: Txtnompro1.SelLength = Len(Txtnompro1.Text)
    
End Sub

Private Sub txtnompro1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then txcodmarca.SetFocus

End Sub

Private Sub txtnompro3_GotFocus()

    txtnompro3.SelStart = 0: txtnompro3.SelLength = Len(txtnompro3.Text)
    
End Sub

Private Sub txtnompro3_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtnompro4_GotFocus()
    
    txtnompro4.SelStart = 0: txtnompro4.SelLength = Len(txtnompro4.Text)
    
End Sub

Private Sub txtnompro4_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtotroscostos_GotFocus()

    txtotroscostos.SelStart = 0: txtotroscostos.SelLength = Len(txtotroscostos.Text)
    
End Sub

Private Sub txtotroscostos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtotroscostos_LostFocus()

    txtotroscostos.Text = Format(txtotroscostos.Text, "###,###,##0.00")
    CALCULA_IMPORTES

End Sub

Private Sub txtprevta_GotFocus()

    txtprevta.SelStart = 0: txtprevta.SelLength = Len(txtprevta.Text)
    
End Sub

Private Sub txtprevta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtprevta_LostFocus()
    
    txtprevta.Text = Format(txtprevta.Text, "#,###,###0.0")
    txtvalvta.Text = txtprevta.Text / (1 + wIgv / 100)
    TxtIgv.Text = txtprevta.Text - txtvalvta.Text

End Sub

Private Sub txttceuros_GotFocus()

    txttceuros.SelStart = 0: txttceuros.SelLength = Len(txttceuros.Text)
    
End Sub

Private Sub txttceuros_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txttceuros_LostFocus()

    If IsNumeric(txtCostoEuros.Text) Then
        wcostoeuro = CDbl(txtCostoEuros.Text)
    Else
        wcostoeuro = 0
    End If
    
    If IsNumeric(txttceuros.Text) Then
        wtc = CDbl(txttceuros.Text)
    Else
        wtc = 0
    End If
    
    fob.Text = Format(txtCostoEuros.Text * Val(txttceuros.Text), "#,###,###0.00")
    
    CALCULA_IMPORTES
    
End Sub

Private Sub Txtvalvta_GotFocus()
    
    txtvalvta.SelStart = 0: txtvalvta.SelLength = Len(txtvalvta.Text)
    
End Sub

Private Sub Txtvalvta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub ubicacion_GotFocus()
    
    ubicacion.SelStart = 0: ubicacion.SelLength = Len(ubicacion.Text)
    
End Sub

Private Sub ubicacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ubicacion.Text = Format(ubicacion.Text, "000")
        ubicacion.Text = UCase(Trim(ubicacion.Text))
        sql = "select f2desubi from ef2ubica where f2codubi = '" & ubicacion.Text & " '"
        If rsfield.State = 1 Then rsfield.Close
        rsfield.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsfield.EOF Then
            TxtUbica1.Caption = "" & Mid(rsfield.Fields("F2DESUBI"), 1, 26)
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
        Else
            Beep
        End If
        rsfield.Close
    End If

End Sub

Public Sub CargarMoneda()

    CmbMoneda.Clear
    CmbMoneda.AddItem "Soles"
    CmbMoneda.AddItem "Dolares"
    CmbMoneda.AddItem "Euros"
    CmbMoneda.ListIndex = 1

End Sub

Public Sub CargarTipo()

    cmbtipo.AddItem "FOB"
    cmbtipo.AddItem "EXFAB"
    cmbtipo.AddItem "CIF"
    cmbtipo.AddItem "LOCAL"
    cmbtipo.ListIndex = 0

End Sub

Public Sub ControlaEuro(Estado As Boolean)
    
    txtCostoEuros.Enabled = Estado
    txttceuros.Enabled = Estado
    
End Sub

Private Sub cuenta_LostFocus()
    
    If rst.State = adStateOpen Then rst.Close
    sql = "select f5codcta, f5nomcta from cf5pla where f5codcta='" & cuenta.Text & "'"
    rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        pnlgasto.Caption = "" & rst("f5nomcta")
        sql = "select codigo from bf9gin where cuenta='" & cuenta.Text & "' and base='G'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not rst.EOF Then
            txtgasto.Text = "" & rst("codigo")
        End If
    End If
    rst.Close

End Sub

Public Sub AsignaDerechos()
    For J = 1 To 4
        derecho(J) = False
    Next J
    
    sql = "select * from ef2users_der where f2coduser='" & wusuario & "' order by codigo"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        Do While Not rst.EOF
            Select Case Val("" & rst("codigo"))
                Case 1
                    derecho(1) = True
                Case 2
                    derecho(2) = True
                Case 3
                    derecho(3) = True
                Case 4
                    derecho(4) = True
            End Select
            rst.MoveNext
        Loop
    End If
    rst.Close
    
'    For J = 1 To 4
'        If derecho(J) Then
'            vaTabPro1.Tab = J - 1
'            vaTabPro1.TabState = 0
'        Else
'            vaTabPro1.Tab = J - 1
'            vaTabPro1.TabState = 2
'        End If
'    Next J

End Sub

Private Sub CALCULA_IMPORTES()
Dim nvvta1  As Double
Dim nvvta2  As Double
Dim nvvta3  As Double
Dim nvvta4  As Double
    
    wfactor = txtfactor.Text
    wFob = CDbl(fob.Text)
    'If Val(Format(txtflete.Text, "0.00")) > 0 Then
    '    nflete = Val(Format(txtflete.Text, "0.00"))
    '    wvalvta1 = ((wfob + nflete) * wfactor) + Val(Format(txtinstalacion.Text, "0.00"))
    'Else
    '    wvalvta1 = (wfob * wfactor) + Val(Format(txtinstalacion.Text, "0.00"))
    'End If
    '-------------------------------------------------------------------------
    If CDbl(txtflete.Text) > 0 Then
        nvvta1 = CDbl(txtflete.Text)
    Else
        nvvta1 = 1
    End If
    If CDbl(txtfactor.Text) > 0 Then
        nvvta2 = CDbl(txtfactor.Text)
    Else
        nvvta2 = 1
    End If
    If CDbl(txtmanipuleo.Text) > 0 Then
        nvvta3 = CDbl(txtmanipuleo.Text)
    Else
        nvvta3 = 1
    End If
    If CDbl(txtinstalacion.Text) > 0 Then
        nvvta4 = CDbl(txtinstalacion.Text)
    Else
        nvvta4 = 1
    End If
    
    wvalvta1 = (wFob * nvvta1 * nvvta2 * nvvta3 * nvvta4) + CDbl(txtotroscostos.Text)
    '-------------------------------------------------------------------------
    
    xigv = wvalvta1 * (wIgv / 100)
    wprevta = CDbl(wvalvta1) + CDbl(xigv)
    
    txtprevta.Text = Format(wprevta, "#,###,###0.0")
    
    wprevta = CDbl(txtprevta.Text)
    wvalvta2 = wprevta / (1 + wIgv / 100)
    txtvalvta.Text = wvalvta2
    TxtIgv.Text = wprevta - wvalvta2
    
End Sub

'Private Sub vaTabPro1_TabShown(ActiveTab As Integer)
'    If Panelcod.Visible = True Then
'        If ActiveTab = 1 Then
'            vaTabPro1.ActiveTab = 0
'        End If
'    End If
'End Sub
