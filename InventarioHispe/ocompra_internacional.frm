VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ocompra_internacional 
   Caption         =   "Orden de Compra Internacional"
   ClientHeight    =   8085
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlcosto 
      Height          =   1860
      Left            =   4050
      TabIndex        =   63
      Top             =   3195
      Visible         =   0   'False
      Width           =   3705
      _Version        =   65536
      _ExtentX        =   6535
      _ExtentY        =   3281
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdopera 
         Default         =   -1  'True
         Height          =   330
         Index           =   0
         Left            =   450
         TabIndex        =   68
         Top             =   1350
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&Imprimir"
      End
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3375
         TabIndex        =   67
         Top             =   80
         Width           =   240
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Imprimir"
         Top             =   45
         Width           =   3615
      End
      Begin Threed.SSOption optcosto 
         Height          =   285
         Index           =   0
         Left            =   405
         TabIndex        =   64
         Top             =   675
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "&Con Costo"
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
      Begin Threed.SSOption optcosto 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   66
         Top             =   675
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "&Sin Costo"
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
      Begin Threed.SSCommand cmdopera 
         Cancel          =   -1  'True
         Height          =   330
         Index           =   1
         Left            =   2160
         TabIndex        =   69
         Top             =   1350
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&Salir"
      End
   End
   Begin VB.Frame Frame4 
      Height          =   900
      Left            =   11835
      TabIndex        =   53
      Top             =   6165
      Visible         =   0   'False
      Width           =   11715
      Begin VB.TextBox txtlugar_entrega 
         Height          =   300
         Left            =   1485
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   10065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Entrega"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   54
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   12000
      TabIndex        =   41
      Top             =   6405
      Visible         =   0   'False
      Width           =   5505
      Begin VB.TextBox txtbase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   495
         Width           =   1335
      End
      Begin VB.TextBox txtmonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   495
         Width           =   1200
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   495
         Width           =   1110
      End
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "B. Imponible"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Monto Inaf."
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1620
         TabIndex        =   51
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "I.G.V."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2970
         TabIndex        =   50
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label12 
         Caption         =   "Total "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4320
         TabIndex        =   49
         Top             =   270
         Width           =   435
      End
      Begin VB.Label lblmoneda 
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
         Index           =   0
         Left            =   1260
         TabIndex        =   48
         Top             =   210
         Width           =   195
      End
      Begin VB.Label lblmoneda 
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
         Index           =   1
         Left            =   2475
         TabIndex        =   47
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblmoneda 
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
         Index           =   2
         Left            =   3555
         TabIndex        =   46
         Top             =   270
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2265
      Left            =   45
      TabIndex        =   27
      Top             =   1890
      Width           =   11715
      Begin VB.TextBox txtshipto 
         Height          =   300
         Left            =   1350
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1845
         Width           =   10230
      End
      Begin VB.TextBox txtobserva 
         Height          =   315
         Left            =   7245
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   4335
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   510
         Left            =   5940
         TabIndex        =   59
         Top             =   1260
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   900
         _StockProps     =   14
         Caption         =   "Idioma"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optidioma 
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   60
            Top             =   225
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Ingles"
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
         Begin Threed.SSOption optidioma 
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   61
            Top             =   225
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Español"
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
      Begin VB.TextBox txtplazo_entrega 
         Height          =   315
         Left            =   7245
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txt_tc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10755
         TabIndex        =   6
         Text            =   "0.000"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtcodforma 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtcodsoli 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   300
         Left            =   1350
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   4380
      End
      Begin VB.ComboBox Cmbmone 
         Height          =   330
         ItemData        =   "ocompra_internacional.frx":0000
         Left            =   7245
         List            =   "ocompra_internacional.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   225
         Width           =   1410
      End
      Begin VB.ComboBox cmbtipopera 
         Height          =   330
         ItemData        =   "ocompra_internacional.frx":0004
         Left            =   9540
         List            =   "ocompra_internacional.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1410
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtcodcosto 
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   1440
         Width           =   1110
      End
      Begin VB.TextBox txtuupp 
         Height          =   315
         Left            =   4500
         MaxLength       =   2
         TabIndex        =   13
         Top             =   2610
         Width           =   1095
      End
      Begin Threed.SSPanel pnlnomsoli 
         Height          =   300
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   3585
         _Version        =   65536
         _ExtentX        =   6324
         _ExtentY        =   529
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
      Begin Threed.SSPanel pnlnomforma 
         Height          =   300
         Left            =   2160
         TabIndex        =   29
         Top             =   600
         Width           =   3585
         _Version        =   65536
         _ExtentX        =   6324
         _ExtentY        =   529
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
      Begin Threed.SSPanel pnlnomcosto 
         Height          =   300
         Left            =   2430
         TabIndex        =   30
         Top             =   1455
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   529
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
      Begin Threed.SSPanel txtdesuupp 
         Height          =   330
         Left            =   5580
         TabIndex        =   55
         Top             =   2610
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ship to"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   135
         TabIndex        =   73
         Top             =   1890
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plazo de Entrega"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5985
         TabIndex        =   40
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   39
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   38
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda "
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   5985
         TabIndex        =   37
         Top             =   330
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   17
         Left            =   9585
         TabIndex        =   36
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   35
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   5985
         TabIndex        =   34
         Top             =   1005
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Oper."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   8280
         TabIndex        =   33
         Top             =   1455
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblccosto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   32
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label lbluupp 
         AutoSize        =   -1  'True
         Caption         =   "UUPP"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3285
         TabIndex        =   31
         Top             =   2625
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   45
      TabIndex        =   18
      Top             =   540
      Width           =   11715
      Begin Threed.SSPanel pnlnumero 
         Height          =   285
         Left            =   3645
         TabIndex        =   71
         Top             =   240
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   12632256
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
      Begin VB.TextBox txtcontacto 
         Height          =   315
         Left            =   7575
         TabIndex        =   57
         Top             =   990
         Width           =   4035
      End
      Begin VB.TextBox txtusuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8460
         TabIndex        =   1
         Top             =   225
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txt_NumSolComp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6660
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox Txt_NumOC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   225
         Width           =   1125
      End
      Begin VB.TextBox Txt_Prove 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   630
         Width           =   1125
      End
      Begin Threed.SSPanel pnldireprv 
         Height          =   270
         Left            =   1530
         TabIndex        =   19
         Top             =   990
         Width           =   4725
         _Version        =   65536
         _ExtentX        =   8334
         _ExtentY        =   476
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
      Begin Threed.SSPanel pnlnomprv 
         Height          =   285
         Left            =   2700
         TabIndex        =   20
         Top             =   630
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   503
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
      Begin aBoxCtl.aBox txt_fecha 
         Height          =   315
         Left            =   10215
         TabIndex        =   2
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D01000101"
         MaxValue        =   "D99991231"
         ABoxStyle       =   2
         AlignmentVertical=   2
         HideSelection   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "24/11/2006"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         NumberFormat    =   17
         DecimalPlaces   =   0
         HotAppearance   =   2
         CalendarTrailingForeColor=   -2147483629
         BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowButton      =   1
         ButtonPicture   =   "ocompra_internacional.frx":0028
         ButtonWidth     =   21
         UpDownWidth     =   14
         NullText        =   ""
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalcBtnHotStyle =   4
         CalcBackColor   =   -2147483643
         CalcBtnBackColor=   -2147483643
         CalcBtnDigitColor=   -2147483646
         CalcBtnFuntionColor=   8388736
         CalcDisplayFrameColor=   65535
         CalcHeaderBackColor=   -2147483646
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Interno"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   2835
         TabIndex        =   70
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   6780
         TabIndex        =   58
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   9630
         TabIndex        =   26
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   7830
         TabIndex        =   25
         Top             =   270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Solicitud Suministro"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   5175
         TabIndex        =   24
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Orden Compra"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   135
         TabIndex        =   23
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   990
         Width           =   675
      End
   End
   Begin Crystal.CrystalReport Cryordcompra 
      Left            =   11430
      Top             =   7605
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton cmdFirmaAprob 
      Caption         =   "Firmas de Aprobación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6795
      TabIndex        =   16
      Top             =   8415
      Visible         =   0   'False
      Width           =   2100
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   60
      TabIndex        =   56
      Top             =   60
      Width           =   7260
      _Version        =   65536
      _ExtentX        =   12806
      _ExtentY        =   741
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
      Begin ActiveToolBars.SSActiveToolBars atbmenu 
         Left            =   45
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   7
         Tools           =   "ocompra_internacional.frx":037A
         ToolBars        =   "ocompra_internacional.frx":5BEE
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3735
      Left            =   90
      OleObjectBlob   =   "ocompra_internacional.frx":5D92
      TabIndex        =   62
      Top             =   4230
      Width           =   11670
   End
   Begin VB.Label lbldescripcion 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      Height          =   510
      Left            =   7425
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Menu mnuitems 
      Caption         =   "&Item"
      Visible         =   0   'False
      Begin VB.Menu opcdel 
         Caption         =   "&Eliminar item"
      End
      Begin VB.Menu opcinsert 
         Caption         =   "&Insertar"
      End
   End
End
Attribute VB_Name = "ocompra_internacional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOrdenCab              As ADODB.Recordset
Dim rsOrdenDet              As ADODB.Recordset
Dim rssolcab                As ADODB.Recordset
Dim rsSolDet                As ADODB.Recordset
Dim rst                     As ADODB.Recordset
Dim rstaux                  As ADODB.Recordset
Dim rsproductos             As ADODB.Recordset
Dim SWcondipago             As Integer
Dim Wnuevo                  As Boolean
Dim flagIGV                 As Boolean
Dim seleccion               As Boolean
Dim CadSql                  As String
Dim cnn_form                As New ADODB.Connection
Dim cconex_form             As String
Dim sw_nuevo_item           As Boolean
Dim ExisteOrdenCompra       As Boolean
Dim wigv                    As Single
Dim swGrabacion             As Boolean
Dim inicio                  As Boolean
Dim sw_activate             As Boolean
Dim sw_cabecera             As Boolean
Dim sw_detalle              As Boolean
Dim sw_ayuda                As Boolean
Dim wgraba                  As Integer
Dim FlagGeneraOC            As Boolean
Dim jc                      As Integer
Dim flagwin                 As Boolean
Dim FlagAcceso              As Boolean
Dim whelp_solicitud         As Boolean
Dim xnombre                 As String
Dim flag                    As Boolean
Dim wgrabar                 As Boolean
Dim xnumero                 As String
Dim wopc                    As Byte

Private Sub Imprime_Orden()
Dim SQL As String
Dim rsconsulta As New ADODB.Recordset
Set rsconsulta = New ADODB.Recordset
Dim RsPago As New ADODB.Recordset
Set RsPago = New ADODB.Recordset
Dim RsCTR_COM As New ADODB.Recordset
Set RsCTR_COM = New ADODB.Recordset

    With Acr_OrdenC_Imp
        If Cmbmone.ListIndex = 0 Then
            '.lblmoneda1.Caption = "S/."
            .lblmoneda2.Caption = "S/."
        Else
            If Cmbmone.ListIndex = 1 Then
                '.lblmoneda1.Caption = "US$"
                .lblmoneda2.Caption = "US$"
            Else
                .lblmoneda2.Caption = "Æ "
            End If
        End If
        .flddirec1.Text = wf1direc1
        .flddirec2.Text = wf1direc2
        .fldruc.Text = wrucempresa
        .fldempresa.Text = wnomcia
        '.IGV.Caption = wigv
        GOC = Txt_NumOC.Text
        If rsconsulta.State = adStateOpen Then rsconsulta.Close
        SQL = "SELECT A.F4NUMINTERNO,A.F4SHIPTO,A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA,A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA " & _
              " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
    
        rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconsulta.EOF Then
            Rem NSE .F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
            .f4numord.Text = (rsconsulta.Fields("F4NUMINTERNO") & "")
            .F4CODSOLICITUD.Text = "" & rsconsulta.Fields("F4CODSOLICITUD")
            .f2nomprov.Text = "" & rsconsulta.Fields("F2NOMPROV")
            .f2dirprov.Text = "" & rsconsulta.Fields("F2DIRPROV")
            .f2contacto.Text = "" & rsconsulta.Fields("F4CONTACTO")
            .f2telprov.Text = "" & rsconsulta.Fields("F2TELPROV")
            .f2faxprov.Text = "" & rsconsulta.Fields("F2FAXPROV")
            .f4fecemi.Text = "" & rsconsulta.Fields("F4FECEMI")
            .F4MONTO.Text = Format("" & rsconsulta.Fields("F4MONTO"), "0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & rsconsulta.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
                .f2despag.Text = "" & RsPago.Fields("F2DESPAG")
            End If
            RsPago.Close
            If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
            If Not RsCTR_COM.EOF Then
                .f4obsfecha.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OCI")
                .f4emitir.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OCI") & "                                                                   E MAIL :" & RsCTR_COM.Fields("f1email_oc")
                .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OCI")
            End If
            RsCTR_COM.Close
            Rem NSE .remitir.Text = rsconsulta.Fields("F4LUGAR_ENTREGA") & ""
            .remitir.Text = Trim(rsconsulta.Fields("F4SHIPTO") & "")
            .LBLFIRMA.Caption = Trim("" & pnlnomsoli.Caption)
            .F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
        End If
        rsconsulta.Close
        
        .DataControl1.ConnectionString = cnn_form
        .DataControl1.Source = "SELECT * FROM DETALLE"
        If optidioma(0).Value = True Then
            .F5NOMPRO.DataField = "F5NOMPRO_ING"
        Else
            .F5NOMPRO.DataField = "F5NOMPRO"
        End If
        '.F3FECEN.Text = dxDBGrid1.Columns.ColumnByFieldName("F3FENTREGA").Value
        .f3fecen.Text = txtplazo_entrega.Text
    
        wopcion = IIf(optcosto(0).Value, 1, 2)
        
        If wopcion = 1 Then
            .lblcosto.Visible = True
            .F3PREUNI.Visible = True
            .Line36.Visible = True
            .Line32.Visible = True
            .lbltotal.Visible = True
            .txttotal.Visible = True
            .Line29.Visible = True
            .Line24.Visible = True
            .lbltotal2.Visible = True
            .lblmoneda2.Visible = True
            .F4MONTO.Visible = True
        Else
            .lblcosto.Visible = False
            .F3PREUNI.Visible = False
            .Line36.Visible = False
            .Line32.Visible = False
            .lbltotal.Visible = False
            .txttotal.Visible = False
            .Line29.Visible = False
            .Line24.Visible = False
            .lbltotal2.Visible = False
            .lblmoneda2.Visible = False
            .F4MONTO.Visible = False
        End If
        .Show vbModal
    End With
    
End Sub

Private Sub Imprime_Orden2()
Dim SQL As String
Dim rsconsulta As New ADODB.Recordset
Set rsconsulta = New ADODB.Recordset
Dim RsPago As New ADODB.Recordset
Set RsPago = New ADODB.Recordset
Dim RsCTR_COM As New ADODB.Recordset
Set RsCTR_COM = New ADODB.Recordset
Dim oEXL As ActiveReportsExcelExport.ARExportExcel

    With Acr_OrdenC_Imp
        If Cmbmone.ListIndex = 0 Then
            '.lblmoneda1.Caption = "S/."
            .lblmoneda2.Caption = "S/."
        Else
            If Cmbmone.ListIndex = 1 Then
                '.lblmoneda1.Caption = "US$"
                .lblmoneda2.Caption = "US$"
            Else
                .lblmoneda2.Caption = "Æ "
            End If
        End If
        .flddirec1.Text = wf1direc1
        .flddirec2.Text = wf1direc2
        .fldruc.Text = wrucempresa
        .fldempresa.Text = wnomcia
        '.IGV.Caption = wigv
        GOC = Txt_NumOC.Text
        If rsconsulta.State = adStateOpen Then rsconsulta.Close
        SQL = "SELECT A.F4NUMINTERNO,A.F4SHIPTO,A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA,A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA " & _
              " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
    
        rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconsulta.EOF Then
            Rem NSE .F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
            .f4numord.Text = (rsconsulta.Fields("F4NUMINTERNO") & "")
            .F4CODSOLICITUD.Text = "" & rsconsulta.Fields("F4CODSOLICITUD")
            .f2nomprov.Text = "" & rsconsulta.Fields("F2NOMPROV")
            .f2dirprov.Text = "" & rsconsulta.Fields("F2DIRPROV")
            .f2contacto.Text = "" & rsconsulta.Fields("F4CONTACTO")
            .f2telprov.Text = "" & rsconsulta.Fields("F2TELPROV")
            .f2faxprov.Text = "" & rsconsulta.Fields("F2FAXPROV")
            .f4fecemi.Text = "" & rsconsulta.Fields("F4FECEMI")
            .F4MONTO.Text = Format("" & rsconsulta.Fields("F4MONTO"), "0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & rsconsulta.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
                .f2despag.Text = "" & RsPago.Fields("F2DESPAG")
            End If
            RsPago.Close
            If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
            If Not RsCTR_COM.EOF Then
                .f4obsfecha.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OCI")
                .f4emitir.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OCI") & "                                                                   E MAIL :" & RsCTR_COM.Fields("f1email_oc")
                .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OCI")
            End If
            RsCTR_COM.Close
            Rem NSE .remitir.Text = rsconsulta.Fields("F4LUGAR_ENTREGA") & ""
            .remitir.Text = Trim(rsconsulta.Fields("F4SHIPTO") & "")
            .LBLFIRMA.Caption = Trim("" & pnlnomsoli.Caption)
            .F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
        End If
        rsconsulta.Close
        
        .DataControl1.ConnectionString = cnn_form
        .DataControl1.Source = "SELECT * FROM DETALLE"
        If optidioma(0).Value = True Then
            .F5NOMPRO.DataField = "F5NOMPRO_ING"
        Else
            .F5NOMPRO.DataField = "F5NOMPRO"
        End If
        '.F3FECEN.Text = dxDBGrid1.Columns.ColumnByFieldName("F3FENTREGA").Value
        .f3fecen.Text = txtplazo_entrega.Text
        wopcion = IIf(optcosto(0).Value, 1, 2)
        If wopcion = 1 Then
            .lblcosto.Visible = True
            .F3PREUNI.Visible = True
            .Line36.Visible = True
            .Line32.Visible = True
            .lbltotal.Visible = True
            .txttotal.Visible = True
            .Line29.Visible = True
            .Line24.Visible = True
            .lbltotal2.Visible = True
            .lblmoneda2.Visible = True
            .F4MONTO.Visible = True
        Else
            .lblcosto.Visible = False
            .F3PREUNI.Visible = False
            .Line36.Visible = False
            .Line32.Visible = False
            .lbltotal.Visible = False
            .txttotal.Visible = False
            .Line29.Visible = False
            .Line24.Visible = False
            .lbltotal2.Visible = False
            .lblmoneda2.Visible = False
            .F4MONTO.Visible = False
        End If
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = "c:\mis documentos\oc" & Format(Txt_NumOC.Text, "0000000") & ".xls"
        oEXL.Export Acr_OrdenC_Imp.Pages
        .Run
    End With
    'Acr_OrdenC_Imp.Show vbModal

End Sub

Private Sub EMAIL()
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim tbcnt           As DAO.Recordset
Dim dbcnt           As DAO.Database
    
    sTo = wemail_prove  ' wemailoc
    'sTo = "nlamas@hotmail.com"
    sCC = wemailccoc
    sBCC = ""
    sSubject = wasuntooc & " " & Txt_NumOC.Text
    sBody = wtextooc
    
    ret = Shell("Start.exe " _
        & "mailto:" & """" & sTo & """" _
        & "?Subject=" & """" & sSubject & """" _
        & "&cc=" & """" & sCC & """" _
        & "&bcc=" & """" & sBCC & """" _
        & "&Body=" & """" & sBody & """" _
        & "&File=" & """" & "c:\autoexec.bat" & """" _
        & "Attach=" & """" & "c:\autoexec.bat" & """" _
        , 0)
        
End Sub

Private Sub Calcula_PvtaTot()
Dim cantidad    As Double
Dim totdcto     As Double
Dim ValVta      As Double
Dim IGV         As Double
Dim preciounit  As Double
Dim TOTAL       As Double
Dim costo       As Double

    With dxDBGrid1
        cantidad = Val(Format(.Columns(5).Value, "0.00"))
        If cantidad > 0 Then
            If cmbtipopera.ListIndex = 0 Then
                If .Columns(11).Value = "*" Then     'Afecto
                    totdcto = (Val(Format("" & .Columns(6).Value, "0.00")) * Val(Format("" & .Columns(8).Value, "0.00"))) / 100
                    .Columns(9).Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format("" & .Columns(6).Value, "0.000")) - totdcto, "0.00"))
                    .Columns(10).Value = Format$(ValVta, "###,##0.00")
                    IGV = ValVta * (wgigv / 100)
                    .Columns(12).Value = Format$(IGV, "#,##0.00")
                    Rem NSE preciounit = valvta + igv
                    preciounit = Val(Format("" & .Columns(6).Value, "0.000")) + (Val(Format("" & .Columns(6).Value, "0.000")) * (wgigv / 100))
                    .Columns(7).Value = Format$(preciounit, "###,##0.00")
                    Rem NSE total = preciounit ' * cantidad
                    TOTAL = ValVta + IGV
                    .Columns(13).Value = Format$(TOTAL, "###,##0.00")
                Else  'Inafecto
                    IGV = 0
                    .Columns(12).Value = Format$(IGV, "0.00")
                    totdcto = Val(Format("" & .Columns(6).Value, "0.000")) * Val(Format("" & .Columns(8).Value, "0.00")) / 100
                    .Columns(9).Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format(.Columns(6).Value, "0.000")) - totdcto, "0.00"))
                    .Columns(10).Value = Format$(ValVta, "###,##0.00")
                    Rem NSE preciounit = valvta
                    preciounit = Val(Format("" & .Columns(6).Value, "0.000"))
                    .Columns(7).Value = Format$(preciounit, "###,##0.00")
                    Rem NSE total = preciounit '* cantidad
                    TOTAL = ValVta + IGV
                    .Columns(13).Value = Format$(TOTAL, "###,##0.00")
                End If
            Else
                costo = Val(Format(.Columns(6).Value, "###,##0.00"))
                TOTAL = cantidad * costo                '
                .Columns(13).Value = Format$(TOTAL, "###,##0.00")
            End If
        End If
    End With
    
End Sub

Sub MostrarDatos()
Dim sw_nuevo_temp   As Boolean
Dim xnombre         As String
Dim i               As Integer
Dim entrega         As Date
Dim rs_prod         As New ADODB.Recordset
Dim wlista()        As String
    
    If rssolcab.State = adStateOpen Then rssolcab.Close
    With rssolcab
        'rssolcab.Open "select * from tb_cabsolicitud where cod_solicitud='" & num_solcomp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        rssolcab.Open "select * from tb_cabsolicitud where cod_solicitud='" & lista(1).numero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not .EOF And Not .Bof Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV,F2CONTACTO,F2EMAIL from EF2PROVEEDORES where F2newruc='" & !cs_proveedor & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                Txt_Prove.Text = "" & rst!F2NEWRUC
                pnlnomprv.Caption = rst!f2nomprov
                pnldireprv.Caption = IIf(IsNull(rst!f2dirprov), " ", rst!f2dirprov)
                txtcontacto.Text = "" & rst.Fields("F2CONTACTO")
                wemail_prove = "" & rst.Fields("F2EMAIL")
            Else
                pnlnomprv.Caption = "Ruc es menor a 11 digitos"
                pnldireprv.Caption = "No tiene "
            End If
            rst.Close
            Txt_NumSolComp = !cod_solicitud & ""
            txt_fecha.Value = !cs_fecha & ""
            
            txtcodcosto = !cs_codcosto & ""
            'wcodcosto = !cs_codcosto & ""
            
            pnlnomcosto = !cs_descosto & ""
            
            xnombre = !cs_codsolicitante
            txtobserva.Text = Trim("" & !cs_observaciones)
            If rstaux.State = adStateOpen Then rstaux.Close
            rstaux.Open "SELECT f2nomuser FROM ef2users WHERE f2coduser='" & Trim(xnombre) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rstaux.EOF Then
                txtcodsoli.Text = xnombre
                pnlnomsoli.Caption = "" & rstaux.Fields("f2nomuser")
            Else
                pnlnomsoli.Caption = ""
            End If
            rstaux.Close
            If !cs_moneda = "S" Then
                Cmbmone.ListIndex = 0
            Else
                Cmbmone.ListIndex = 1
            End If
            txtuupp.Text = .Fields("UUPP") & ""
            If VALIDA_UUPP(txtuupp.Text) = True Then
                txtdesuupp.Caption = wdeslocalidad
            End If
            txtlugar_entrega.Text = Left(Trim("" & !cs_LugEntr), 100)
            txt_tc.Text = Format(Val(.Fields("F4TIPCAM") & ""), "0.000")
        End If
        rssolcab.Close
    End With
     
    '*** detalle de solicitud de compra
    'Versión Nueva

    With dxDBGrid1
        If rsSolDet.State = adStateOpen Then rsSolDet.Close
        If lista(1).TOTAL > 1 Then
            txtcodcosto.Text = ""
            pnlnomcosto.Caption = ""
        End If
        
        If lista(1).TOTAL > 0 Then
            cad = ""
            coma = ","
            For i = 1 To lista(1).TOTAL
                cad = cad & "'" & lista(i).numero & "'" & coma
            Next i
            Mid(cad, Len(cad), 1) = " "
            cad = "(" & Trim(cad) & ")"
            
            SQL = "select * from tb_detsolicitud where cod_solicitud in " & cad & " and candis>0 ORDER BY ITEM"
        Else
            SQL = "select * from tb_detsolicitud where cod_solicitud='" & num_solcomp & "' and candis>0 ORDER BY ITEM"
        End If
        rsSolDet.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not (rsSolDet.EOF) Then
            If sw_nuevo_documento = False Then
                DELETEREC_N cnomtabla, cnn_form
                AdicionaItem
                sw_nuevo_documento = True
            End If
            .Dataset.ADODataset.ConnectionString = cnn_form
            .Dataset.Active = True
            .Dataset.Close
            .Dataset.Open
            .OptionEnabled = False
            .Dataset.DisableControls
            sw_nuevo_temp = False
            sw_nuevo_item = True
            rsSolDet.MoveFirst
            j = 0
            i = 0
            Do While Not (rsSolDet.EOF)
                j = j + 1
                If sw_nuevo_temp = False Then
                    If sw_nuevo_documento = True Then
                        dxDBGrid1.Dataset.Edit
                    Else
                        dxDBGrid1.Dataset.Append
                    End If
                    sw_nuevo_temp = True
                Else
                    dxDBGrid1.Dataset.Append
                End If
                i = i + 1
                .Dataset.FieldValues("item") = i
                .Dataset.FieldValues("f3canpro") = rsSolDet!candis
                .Dataset.FieldValues("f3canpro2") = rsSolDet!candis
                .Dataset.FieldValues("BACKORDER") = 0
                .Dataset.FieldValues("COD_SOLICITUD") = "" & rsSolDet!cod_solicitud
                .Dataset.FieldValues("f3codpro") = rsSolDet!COD_PRODUCTO & ""
                .Dataset.FieldValues("f5codfab") = Trim(rsSolDet!F5CODFAB & "")
                .Dataset.FieldValues("f5codmarca") = Trim(rsSolDet!f5codmarca & "")
                
                wcodcosto = "" & rsSolDet!f5codcosto
                If Trim(wcodcosto) <> "" Then
                    .Dataset.FieldValues("f5codcosto") = wcodcosto
                    If rst.State = adStateOpen Then rst.Close
                    SQL = "select f3abrev from centros where f3costo='" & wcodcosto & "'"
                    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
                    If Not rst.EOF Then
                        .Dataset.FieldValues("f5descosto") = "" & rst("f3abrev")
                    End If
                    rst.Close
                Else
                    .Dataset.FieldValues("f5codcosto") = ""
                    .Dataset.FieldValues("f5descosto") = ""
                End If
                
                cmarca = Trim(rsSolDet!f5codmarca & "")
                If rs_prod.State = adStateOpen Then rs_prod.Close
                rs_prod.Open "SELECT f5nompro,f5codfab,F7CODMED,f5marca,f5texto,f5texto_ing from if5pla where F5CODFAB='" & rsSolDet!F5CODFAB & "' and f5marca='" & rsSolDet!f5codmarca & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
                If Not (rs_prod.EOF) Then
                    .Dataset.FieldValues("f3medida") = Trim(rs_prod!F7CODMED & "")
                    .Dataset.FieldValues("f5nompro") = Trim(rs_prod!F5TEXTO & "")
                    .Dataset.FieldValues("f5nompro_ing") = Trim(rs_prod!F5TEXTO_ING & "")
                End If
                rs_prod.Close
                Set rs_prod = Nothing
                '------------------------------------------------------
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & cmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    .Dataset.FieldValues("f5marca") = rsmarcas.Fields("F2DESMAR")
                Else
                    .Dataset.FieldValues("f5marca") = ""
                    .Dataset.FieldValues("f5CODmarca") = ""
                End If
                rsmarcas.Close
                Set rsmarcas = Nothing
                '------------------------------------------------------
                .Dataset.FieldValues("f3precos") = Val(Format(rsSolDet!Precio, "0.000"))
                .Dataset.FieldValues("f3pordct") = Null
                .Dataset.FieldValues("f5afecto") = "*"
                If .Dataset.FieldValues("f5afecto") = "*" Then
                    .Dataset.FieldValues("f3pordct") = Null
                    .Dataset.FieldValues("f5valvta") = Null
                Else
                    .Dataset.FieldValues("f3pordct") = Null
                    .Dataset.FieldValues("f5valvta") = Null
                End If
                entrega = IIf(IsNull(rsSolDet!cs_fentrega), Format$(Date, "dd/mm/yyyy"), Format$(rsSolDet!cs_fentrega, "dd/mm/yyyy"))
                .Dataset.FieldValues("f3fentrega") = entrega
                .Dataset.FieldValues("check") = True
                .Dataset.FieldValues("cant_ant") = 0#
                rsSolDet.MoveNext
                Calcula_PvtaTot
            Loop
            .Dataset.Post
            .Dataset.EnableControls
            .Dataset.Open
            .OptionEnabled = True
            .Dataset.Refresh
            sw_nuevo_item = False
        End If
        rsSolDet.Close
        Call calcula
    End With
    
End Sub

Private Sub atbmenu_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim resp    As Integer
    
    Select Case Tool.Id
        Case "idnuevo":
            inicio = True
            Me.MousePointer = vbHourglass
            Wnuevo = True
            If swGrabacion = True Then
                resp = MsgBox("La Orden no ha sido grabada. ¿Grabar ahora?", vbYesNo + vbQuestion, "Sistema de Logística")
                If resp = vbYes Then
                    MODIFICAR_OC
                End If
            End If
            sw_nuevo_documento = False
            Limpia_Orden
            Limpiar
            AdicionaItem
            AdicionaItem
            sw_nuevo_documento = True
            Txt_NumSolComp.SetFocus
            Me.MousePointer = vbDefault
        Case "idgrabar":
            If Txt_Prove = "" Then MsgBox "Ingrese Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
            If pnlnomprv = "" Then MsgBox "Ingrese Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
            If txtcodsoli = "" Then MsgBox "Ingrese Solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
            If pnlnomsoli = "" Then MsgBox "Ingrese Solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
            If txtcodforma = "" Then MsgBox "Ingrese Forma de Pago", 48, "Sistema de Logística": txtcodforma.SetFocus: Exit Sub
            If Cmbmone.ListIndex < 0 Then MsgBox "Seleccione Moneda", 48, "Sistema de Logística": Cmbmone.SetFocus: Exit Sub
            If Val(txt_tc.Text) = 0 Then MsgBox "Ingrese Tipo de Cambio", 48, "Sistema de Logística": txt_tc.SetFocus: Exit Sub

        
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                dxDBGrid1.Dataset.Post
                sw_detalle = True
            End If
            If MsgBox("¿Desea Grabar la Orden de Compra?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
                Me.MousePointer = vbHourglass
                GrabarOC
                Me.MousePointer = vbDefault
            End If
    
        Case "idimprimir":
            If Len(Trim(Txt_NumOC.Text)) <= 0 Then
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            wopc = 1
            pnlcosto.Visible = True
            Me.MousePointer = vbDefault
        Case "idanular":
            If Trim$(Txt_NumOC.Text) = "" Then
                MsgBox "No existe Orden de Compra", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                eliminar
            End If
        Case "idemail":
            If Len(Trim(Txt_NumOC.Text)) <= 0 Then
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
                Exit Sub
            End If
            wopc = 2
            pnlcosto.Visible = True
        Case "ID_CtasxPagar"
            If Len(Trim(Txt_NumOC.Text)) > 0 Then
                If rsif4orden.State = adStateOpen Then rsif4orden.Close
                rsif4orden.Open "SELECT F4CORRELA FROM IF4ORDEN WHERE F4NUMORD=" & Txt_NumOC.Text & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsif4orden.EOF Then
                    If Val("" & rsif4orden.Fields("F4CORRELA")) > 0 Then
                        MsgBox "La orden de compra ya fue trasladada a cuentas por pagar.", vbInformation, "Atención"
                    Else
                        If MsgBox("Está seguro(a) de trasladar la Orden de Compra a Cuentas por Pagar ?", vbYesNo, "Atención") = vbYes Then
                            TRASLADA_CTASXPAGAR Txt_NumOC.Text
                        End If
                    End If
                End If
                rsif4orden.Close
            Else
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
            End If
        Case "idsalir":
            Unload Me
    End Select
    
End Sub

Private Sub cmbtipopera_Change()
    
    If cmbtipopera.ListIndex = 1 Then
       Cmbmone.ListIndex = 1
    End If
    wgraba = 0
    If Not inicio Then swGrabacion = True
    
End Sub

Private Sub cmbtipopera_Click()

    If cmbtipopera.ListIndex = 1 Then
        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Caption = "Costo"
    Else
        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Caption = "Costo Uni."
        If wf1visualiza_dctos = "*" Then
            dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = False
            dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
            dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
        End If
    End If

End Sub

Private Sub cmbtipopera_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtcodcosto.Visible = True Then
            txtcodcosto.SetFocus
        Else
            If txtuupp.Visible = True Then
                txtuupp.SetFocus
            Else
                dxDBGrid1.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub cmbtipopera_LostFocus()
    
    If cmbtipopera.ListIndex = 1 Then
        Forma_Imp
    Else
        Forma_Loc
    End If
    
End Sub

Sub Forma_Loc()
    
    Visib
    
End Sub

Sub Visib()
    
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    lblmoneda(0).Visible = True
    lblmoneda(1).Visible = True
    lblmoneda(2).Visible = True
    txtmonto.Visible = True
    txtbase.Visible = True
    txtigv.Visible = True
    txttotal.Left = 4005
    Label12.Left = 4320
    
End Sub

Sub Invisi()

    Cmbmone.ListIndex = 1
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Left = 5000
    lblmoneda(0).Visible = False
    lblmoneda(1).Visible = False
    lblmoneda(2).Left = 5600
    txtmonto.Visible = False
    txtbase.Visible = False
    txtigv.Visible = False
    txttotal.Left = 135
    Label12.Left = 135
    
End Sub

Sub Forma_Imp()

    Invisi
    
End Sub

Private Sub cmdcerrar_Click()
    
    pnlcosto.Visible = False
    
End Sub

Private Sub cmdFirmaAprob_Click()
    
'''    frmaccesocompras.Show 1
'''    txtTempo = wusuario
'''    If txtTempo <> "" Then
'''        If xcentro = "08" Then
'''            If txtAprobadoX <> "" Then MsgBox "La orden ya fue firmada por el Jefe de Compras", 48, "Sistema de Logística"
'''            txtAprobadoX = txtTempo
'''            FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''            ImgAprobadoX.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''        Else
'''            If txtAprobadoX = "" Then MsgBox "Debe firmar antes El Jefe de Compras!!!", 48, "Sistema de Logística": Exit Sub
'''            If txtAprobadoY <> "" Then MsgBox "Ya firmó el Jefe del Area!!!", 48, "Sistema de Logística": Exit Sub
'''            If txtAprobadoY = txtTempo Then
'''                FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''                ImgAprobadoY.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''            Else
'''                If txtAprobadoX = "" Then MsgBox "Debe firmar antes El Jefe de Compras!!!", 48, "Sistema de Logística": Exit Sub
'''                If txtAprobadoY = "" Then MsgBox "Debe firmar antes el Jefe del Area!!!", 48, "Sistema de Logística": Exit Sub
'''                If txtAprobadoZ <> "" Then MsgBox "Ya firmó el Gerente de Logística!!!", 48, "Sistema de Logística": Exit Sub
'''                txtAprobadoZ = txtTempo
'''                FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''                ImgAprobadoY.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''            End If
'''        End If
'''    Else
'''        Exit Sub
'''    End If

End Sub

Private Sub Cmbmone_Click()

    Select Case Cmbmone.ListIndex
        Case 0:
            lblmoneda(0).Caption = "S/. "
            lblmoneda(1).Caption = "S/. "
            lblmoneda(2).Caption = "S/. "
            WMONEDAX = "S"
        Case 1:
            lblmoneda(0).Caption = "US$ "
            lblmoneda(1).Caption = "US$ "
            lblmoneda(2).Caption = "US$ "
            WMONEDAX = "D"
        Case 2:
            lblmoneda(0).Caption = "Æ "
            lblmoneda(1).Caption = "Æ "
            lblmoneda(2).Caption = "Æ "
            WMONEDAX = "E"
    End Select
    If Not inicio Then swGrabacion = True

End Sub

Private Sub Cmbmone_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txt_tc.SetFocus
    End If
    
End Sub

Private Sub calcula()
On Error GoTo HNDERR
Dim afecto      As Double
Dim inafecto    As Double
Dim IGV         As Double
Dim SQL         As String

    If cmbtipopera.ListIndex = 0 Then
        SQL = "select sum(iif(f5afecto='*',f5valvta)) as afecto, " _
        & "sum(iif(isnull(f5afecto),f5valvta)) as inafecto, sum(f3igv) as igv from detalle"
        If rst.State = adStateOpen Then rst.Close
        
        If rst.State = adStateOpen Then rst.Close
        rst.Open SQL, cnn_form, adOpenStatic, adLockOptimistic
        If Not (rst.EOF) Then
            afecto = IIf(IsNull(rst.Fields("afecto")), 0, rst.Fields("afecto"))
            inafecto = IIf(IsNull(rst.Fields("inafecto")), 0, rst.Fields("inafecto"))
            IGV = IIf(IsNull(rst.Fields("igv")), 0, rst.Fields("igv"))
            
            txtbase.Text = Format$(afecto, "####,##0.00")
            txtmonto.Text = Format$(inafecto, "####,##0.00")
            txtigv.Text = Format(IGV, "###,###,##0.00")
            txttotal.Text = Format$(afecto + inafecto + IGV, "###,##0.00")
        End If
        rst.Close
    End If
    Exit Sub
    
HNDERR:
    Select Case Err.Number
        Case -2147217865
            Resume Next
    End Select
    
End Sub

Private Sub cmdopera_Click(Index As Integer)

    Select Case Index
        Case 0
            If wopc = 1 Then
                Imprime_Orden
            Else
                Imprime_Orden2
                EMAIL
            End If
        Case 1
            pnlcosto.Visible = False
    End Select

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    If KeyCode = 113 Then
        If dxDBGrid1.Columns.FocusedIndex = 0 Then
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(pnlnomprv.Caption)
            hlp_prov_prod.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                dxDBGrid1.Dataset.FieldValues("F3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = Format$(Date, "dd/mm/yyyy")
                If wvv_prod > 0# Then
                    dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
                End If
                If rsconsulta.State = adStateOpen Then rsconsulta.Close
                rsconsulta.Open "SELECT F5AFECTO FROM IF5PLA WHERE F5CODPRO = '" & dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsconsulta.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Value = "" & rsconsulta.Fields("F5AFECTO")
                End If
                rsconsulta.Close
                dxDBGrid1.Columns.FocusedIndex = 3
                dxDBGrid1.Dataset.Post
            End If
        End If
    End If
    
    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                If dxDBGrid1.Dataset.RecordCount = 0 Then AdicionaItem
            End If
            sw_nuevo_item = False
        End If
    End If

End Sub

Private Sub dxDBGrid1_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)

    If dxDBGrid1.Columns.FocusedIndex = 1 Then
        If Len(Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)) > 0 Then
            lbldescripcion.Visible = True
            lbldescripcion.Caption = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)
        Else
            lbldescripcion.Caption = ""
            lbldescripcion.Visible = False
        End If
    Else
        lbldescripcion.Caption = ""
        lbldescripcion.Visible = False
    End If
    
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
    'Me.Height = 8085
    'Me.Width = 12015
    'Me.Left = 1500
    'Me.Top = 1050

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_form.Close
    sw_nuevo_item = True
    dxDBGrid1.Dataset.Close
    
    Rem NSE ELIMINA_BD_N wrutatemp, cnombase
    lista_oc.dxDBGrid1.Dataset.Active = False
    lista_oc.dxDBGrid1.Dataset.Refresh
    lista_oc.dxDBGrid1.Dataset.Active = True

End Sub

Private Sub define_cabecera()

    lblmoneda(0).Left = 8580
    Label9.Left = 7575
    txtbase.Left = 7350
    
End Sub

Private Sub Form_Load()
Dim fec     As Date
    
    If wf1show_ccosto = "N" Then
        lblccosto.Visible = False
        txtcodcosto.Visible = False
        pnlnomcosto.Visible = False
    Else
        lblccosto.Visible = True
        txtcodcosto.Visible = True
        pnlnomcosto.Visible = True
    End If
    
    If wf1uupp = "*" Then
        lbluupp.Visible = True
        txtuupp.Visible = True
        txtdesuupp.Visible = True
    Else
        lbluupp.Visible = False
        txtuupp.Visible = False
        txtdesuupp.Visible = False
    End If
    
    Set rst = New ADODB.Recordset
    Set rsOrdenCab = New ADODB.Recordset
    Set rsOrdenDet = New ADODB.Recordset
    Set rsproductos = New ADODB.Recordset
    Set rssolcab = New ADODB.Recordset
    Set rsSolDet = New ADODB.Recordset
    Set rstaux = New ADODB.Recordset
    
    sw_ayuda = False
    inicio = True
    swGrabacion = False
    sw_activate = False
    
    If loc = 2 Then
        Call define_cabecera
        txtmonto.Visible = False
        txtigv.Visible = False
        txttotal.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        lblmoneda(1).Visible = False
        lblmoneda(2).Visible = False
        lblmoneda(3).Visible = False
    Else
        loc = 1
    End If
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    fec = txt_fecha.Value
    Wnuevo = True
    flagIGV = False
    SWcondipago = 0
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "select F1IGV from sf1param where f1codemp='" & UCase(wempresa) & "'", cnn_control
    If Not (rst.EOF) Then
         wgigv = rst.Fields("F1IGV")
    End If
    rst.Close
    
    Txt_Prove.Enabled = True
    If FlagGeneraOC = False Then
        Wnuevo = True
    End If
     
    jc = 0
    
    WMONEDAX = ""
    sw_nuevo_item = False
    Rem NSE cnombase = wusuario & "OCOMPRA" & Format(Time, "HH_MM_SS") & ".MDB"
    Rem NSE CREATEDATABASE_N wrutatemp & "\", cnombase
    
    cnombase = "TMPOCOMP.MDB"
    cnomtabla = "DETALLE"
    
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open cconex_form
    
    Rem NSE CadSql = "(item integer,f3codpro text(15),f5nompro text(100),f3canpro double,f3precos double," & _
    REM NSE "f3pordct double,f3totdct double,f5valvta double,f5afecto text(1),f3igv double," & _
    REM NSE "f3preuni double,f3total double,f5codfab text(20),f3fentrega date)"
    
    Rem NSE CREATETABLE_N cnomtabla, CadSql, cnn_form
    
    Call CONFIGURA_GRID
    
    Cmbmone.Clear
    Cmbmone.AddItem "Soles"
    Cmbmone.AddItem "Dólares"
    Cmbmone.AddItem "Euro"
    Cmbmone.ListIndex = 1
    
    Invisi
    If sw_nuevo_documento = True Then
        DELETEREC_N cnomtabla, cnn_form
        AdicionaItem
        Limpiar
    Else
        inicio = True
        MODIFICAR_OC
        sw_nuevo_documento = False
        inicio = False
    End If
    
    wdxcodigo = ""
    wdxcodfab = ""
    wdxdescripcion = ""
    wdxcantidad = 0
    wdxnroitems = 0
    
End Sub

Sub Limpiar()

    'dxDBGrid1.Dataset.Close
    'DELETEREC_N cnomtabla, cnn_form
    'AdicionaItem
        
    SWcondipago = 0
    Txt_NumOC = ""
    Txt_NumSolComp = ""
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    
    txtcontacto.Text = ""
    txtcodsoli = ""
    pnlnomsoli = ""
    txtcodforma = ""
    pnlnomforma = ""
           
    txt_tc.Text = "0.000"
    
    Txt_Referencia = ""
    
    txtbase = "0.00"
    txtmonto = "0.00"
    txtigv.Text = "0.00"
    txttotal = "0.00"
       
    txtuupp.Text = "": txtdesuupp.Caption = ""
       
    SWcondipago = 0
    'txtempresa.Text = UCase$(wempresa)
    
    'hc011-hc012 haro
    
    txtplazo_entrega.Text = ""
    txtlugar_entrega.Text = ""
    
    If optidioma(0).Value = True Then
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro_ing").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Visible = False
    Else
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro_ing").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Visible = True
    End If
    cmbtipopera.ListIndex = 1
    
    wdxcodigo = ""
    wdxcodfab = ""
    wdxdescripcion = ""
    wdxcantidad = 0
    wdxnroitems = 0
    txtshipto.Text = ""
    
    atbmenu.Tools("idanular").Enabled = False
    atbmenu.Tools("idemail").Enabled = False
    atbmenu.Tools("idimprimir").Enabled = False
    atbmenu.Tools("Id_CtasxPagar").Enabled = False
End Sub

Private Sub Limpia_Orden()

    pnlnomcosto.Caption = ""
    Txt_Prove.Text = ""
    pnlnomprv.Caption = ""
    txtcodsoli.Text = ""
    Txt_NumSolComp.Text = ""
    pnlnomsoli.Caption = ""
    txtcodforma.Text = ""
    pnlnomforma.Caption = ""
    txtcodcosto.Text = ""
    pnldireprv.Caption = ""
    Txt_Referencia.Text = ""
    txtobserva.Text = ""
    Txt_NumOC = ""
    pnlnumero.Caption = ""
    txt_tc.Text = "0.000"
    txttotal.Text = "0.00"
    txtigv.Text = "0.00"
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    wgraba = 1
    wdxcodigo = ""
    wdxcodfab = ""
    wdxdescripcion = ""
    wdxcantidad = 0
    wdxnroitems = 0
    
End Sub

Sub Visi()

    txtbase.Visible = True
    txtigv.Visible = True
    txttotal.Visible = True

End Sub

Sub LLENA_TEMPCAB()
Dim cnn         As ADODB.Connection
Dim tempocompra As ADODB.Recordset
Dim X           As Integer
Dim rsprod      As New ADODB.Recordset

    'Nueva Versión
    Set cnn = New ADODB.Connection
    Set tempocompra = New ADODB.Recordset
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\Tempcomp.mdb;Persist Security Info=False"
    cnn.Execute "delete * from tmpocompra"
    
    If tempocompra.State = adStateOpen Then tempocompra.Close
    tempocompra.Open "tmpocompra", cnn, adOpenStatic, adLockOptimistic
    
    With dxDBGrid1
        If .Dataset.RecordCount = 0 Then
            tempocompra.Close
            cnn.Close
            Exit Sub
        End If
        .Dataset.First
        If Not (.Dataset.EOF) Then
            .Dataset.First
            Do While Not (.Dataset.EOF)
                If Val(IIf(IsNull(.Dataset.FieldValues("f3precos")), 0, _
                .Dataset.FieldValues("f3precos"))) > 0 Then
                    tempocompra.AddNew
                    tempocompra!Orden = Format(Txt_NumOC.Text, "0000000")
                    tempocompra!PROVEEDOR = pnlnomprv.Caption
                    tempocompra!direccion = pnldireprv.Caption
                    tempocompra!ruc = Txt_Prove.Text
                    tempocompra!FECHA = txt_fecha.Value
                    tempocompra!FORPAG = pnlnomforma.Caption
                    tempocompra!Moneda = Cmbmone.Text
                    tempocompra!referencia = Txt_Referencia.Text
                    tempocompra!Centro = txtcodcosto.Text
                    tempocompra!nomcentro = pnlnomcosto.Caption
                    tempocompra!OBSERVA = txtobserva.Text
                    tempocompra!SUBTOTAL = txtbase.Text
                    tempocompra!MONTOINA = txtmonto.Text
                    tempocompra!IGV = txtigv.Text
                    tempocompra!TOTAL = dxDBGrid1.Columns.ColumnByName("F3TOTAL").SummaryFooterValue  'txttotal.Text
                    tempocompra!empresa = wnomcia
                    tempocompra!sS = Txt_NumSolComp.Text
                    tempocompra!codigo = "" & .Dataset.FieldValues("f3codpro")
                    tempocompra!Descripcion = "" & .Dataset.FieldValues("f5nompro")
                    tempocompra!cantidad = .Dataset.FieldValues("f3canpro")
                    tempocompra!costo = .Dataset.FieldValues("f3precos")
                    tempocompra!descuento = .Dataset.FieldValues("f3pordct")
                    tempocompra!Precio = .Dataset.FieldValues("f3preuni")
                    
                    If rsprod.State = adStateOpen Then rsprod.Close
                    rsprod.Open "SELECT F7CODMED from if5pla where f5codpro='" & .Dataset.FieldValues("f3codpro") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not (rsprod.EOF) Then
                        tempocompra!unidad = rsprod.Fields("F7CODMED") & ""
                    End If
                    rsprod.Close
                        
                    tempocompra.Update
                End If
                .Dataset.Next
            Loop
            .Dataset.First
        End If
        tempocompra.Close
        cnn.Close
    End With
    
End Sub

Private Sub optidioma_Click(Index As Integer, Value As Integer)
    
    If Index = 0 Then
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro_ing").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Visible = False
    Else
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro_ing").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Visible = True
    End If

End Sub

Private Sub txt_fecha_LostFocus()

    If IsDate(txt_fecha.Value) Then
        If Val(txt_tc.Text & "") = 0# Then
            If rscambios.State = adStateOpen Then rscambios.Close
            rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & txt_fecha.Value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rscambios.EOF Then
                txt_tc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
            Else
                txt_tc.Text = Format(0, "0.000")
            End If
            rscambios.Close
        End If
    Else
        MsgBox "Fecha incorrecta. Verifique.", vbCritical, "Atención"
        txt_fecha.SetFocus
    End If

End Sub

Private Sub Txt_NumOC_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_NumOC_GotFocus()
Txt_NumOC.SelStart = 0
Txt_NumOC.SelLength = Len(Txt_NumOC.Text)
End Sub

Private Sub Txt_NumSolComp_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_NumSolComp_GotFocus()
Txt_NumSolComp.SelStart = 0
Txt_NumSolComp.SelLength = Len(Txt_NumSolComp.Text)
End Sub

Private Sub Txt_Prove_Change()
    
    pnlnomprv.Caption = ""
    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_Prove_GotFocus()

    Txt_Prove.SelStart = 0: Txt_Prove.SelLength = Len(Txt_Prove)
    
End Sub

Private Sub Txt_Prove_LostFocus()
Dim rsprov  As New ADODB.Recordset

    If sw_ayuda = False Then
        If Len(Trim(Txt_Prove.Text)) > 0 Then
            If rsprov.State = adStateOpen Then rsprov.Close
            rsprov.Open "SELECT f2Email,F2NOMPROV,F2DIRPROV,F2contacto,F2TIPMON FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(Txt_Prove.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsprov.EOF Then
                pnlnomprv.Caption = "" & rsprov.Fields("F2NOMPROV")
                pnldireprv.Caption = "" & rsprov.Fields("F2DIRPROV")
                txtcontacto.Text = "" & rsprov.Fields("F2contacto")
                wemail_prove = "" & rsprov.Fields("F2email")
                GRABA_GRID Trim(Txt_Prove.Text)
                If Trim("" & rsprov.Fields("F2TIPMON")) = "D" Then
                    Cmbmone.ListIndex = 1
                ElseIf Trim("" & rsprov.Fields("F2TIPMON")) = "S" Then
                    Cmbmone.ListIndex = 0
                End If
            Else
                MsgBox "El Proveedor no Existe. Verifique.", vbInformation, "Atención"
                Txt_Prove.SetFocus
            End If
            rsprov.Close
        End If
    End If

End Sub

Private Sub Txt_Referencia_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txt_tc_Change()

    If Not inicio Then swGrabacion = True
    
    If txt_tc.Text = " .   " Then
        txt_tc.Text = "0.000"
    End If
    
End Sub

Private Sub txtcodcosto_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txtcodcosto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtuupp.Visible = True Then
            txtuupp.SetFocus
        Else
            dxDBGrid1.SetFocus
        End If
    End If
        
End Sub

Private Sub txtcodcosto_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtcodcosto.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "select f3descrip,f3direccion from centros where f3costo='" & txtcodcosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                pnlnomcosto.Caption = Trim(rst.Fields("f3descrip") & "")
            Else
                pnlnomcosto.Caption = ""
                MsgBox "Centro de costo no existe. Verifique.", vbInformation, "Atenciòn"
                txtcodcosto.SetFocus
            End If
            rst.Close
        Else
            pnlnomcosto.Caption = ""
        End If
    End If

End Sub

Private Sub txtcodforma_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Screen.MousePointer = vbHourglass
        wcodpag = ""
        sw_ayuda = True
        hlp_formapago.Show 1
'        ayu_f_p_c.Show 1
        
        sw_ayuda = False
        If Len(wcodpag) > 0 Then
            txtcodforma = wcodpag
            pnlnomforma = wnompag
            txtcodforma_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtcodforma_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtcodforma.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG='" & Trim(txtcodforma.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                pnlnomforma.Caption = Trim("" & rst!f2despag)
            Else
                pnlnomforma.Caption = ""
                MsgBox "Còdigo de forma de pago no existe. Verifique.", vbInformation, "Atenciòn"
                txtcodforma.SetFocus
            End If
            rst.Close
        End If
    End If

End Sub

Private Sub txtcodsoli_DblClick()

    txtcodsoli_KeyDown 113, 0
    
End Sub

Private Sub txtcodsoli_GotFocus()

    If Len(Trim(txtcodsoli.Text)) = 0 Then
        txtcodsoli.Text = wusuario
    End If
    txtcodsoli.SelStart = 0: txtcodsoli.SelLength = Len(txtcodsoli.Text)

End Sub

Private Sub txtcodsoli_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodusuario = ""
        'hlp_usuarios.Show 1
        ayuda_usuarios.Show 1
        
        sw_ayuda = False
        If Len(Trim(wcodusuario)) > 0 Then
            txtcodsoli.Text = wcodusuario
            pnlnomsoli.Caption = wnomusuario
            txtcodsoli_KeyPress 13
        End If
    End If

End Sub

Private Sub txtcodsoli_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtcodsoli.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT f2nomuser FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
            Else
                pnlnomsoli.Caption = ""
                MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
                txtcodsoli.SetFocus
            End If
            rst.Close
        End If
    End If

End Sub

Private Sub txtcontacto_GotFocus()
    
    txtcontacto.SelStart = 0: txtcontacto.SelLength = Len(txtcontacto)

End Sub

Private Sub txtcontacto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtcodsoli.SetFocus
    End If

End Sub

Private Sub txtobserva_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtcodcosto.Visible = True Then
            txtcodcosto.SetFocus
        Else
            If txtuupp.Visible = True Then
                txtuupp.SetFocus
            Else
                dxDBGrid1.SetFocus
            End If
        End If
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    
End Sub

Private Sub Txt_Fecha_Change()
    
    wgraba = 0
    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_Fecha_GotFocus()
    
    txt_fecha.FocusSelect = True
    
End Sub

Private Sub Txt_Fecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub Txt_NumOC_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Txt_NumOC.Text = Format(Txt_NumOC.Text, "000000")
        If Len(Txt_NumOC.Text) > 0 Then
            flagwin = True
            Wnuevo = False
            GOC = Trim(Txt_NumOC.Text)
            MODIFICAR_OC
            If ExisteOrdenCompra Then
                txt_fecha.SetFocus
            Else
                MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " no existe", vbInformation, "Sistema de Logística"
                Txt_NumOC.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub Txt_NumSolComp_DblClick()

    Call Txt_NumSolComp_KeyDown(113, 0)
    
End Sub

Private Sub Txt_NumSolComp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Screen.MousePointer = vbHourglass
        FlagAcceso = False
        flagwin = True
        whelp_solicitud = 4
        FlagAcceso = False
        'hlp_solicitudes.Show vbModal
     If Len(Trim(num_solcomp)) > 0 Then
            If lista(1).TOTAL > 0 Then
                'Txt_NumSolComp = num_solcomp
                Txt_NumSolComp = lista(1).numero
                Txt_Prove.Enabled = True
                Call MostrarDatos
                'Txt_Prove.Text = ""
                'pnlnomprv.Caption = ""
                'pnldireprv.Caption = ""
                txt_fecha.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub Txt_NumSolComp_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        num_solcomp = Txt_NumSolComp.Text
        Txt_Prove.Enabled = True
        Call MostrarDatos
        Txt_Prove.Text = ""
        pnlnomprv.Caption = ""
        pnldireprv.Caption = ""
        txt_fecha.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
End Sub

Sub MostrarDatosOC()
Dim sw_nuevo_temp   As Boolean
Dim SQL             As String
Dim i               As Integer
Dim cmarca          As String
    
    If loc = 1 Then
        With rsOrdenCab
            If Not (.EOF) Then
                'txtempresa = !F4EMPRESA & ""
                If Txt_NumOC = "" Then
                    !f4numord = " "
                Else
                    Txt_NumOC = !f4numord & ""
                End If
                Txt_NumSolComp = !F4CODSOLICITUD & ""
                pnlnumero.Caption = "" & !F4NUMINTERNO
                txt_fecha.Value = !f4fecemi
                txtobserva.Text = rsOrdenCab!F4OBSERVA & ""
                txtcontacto.Text = "" & rsOrdenCab!F4CONTACTO
                txtplazo_entrega.Text = "" & rsOrdenCab!F4PLAZO_ENTREGA
                If !F4TIPMON = "S" Then
                    Cmbmone.ListIndex = 0
                Else
                    If !F4TIPMON = "D" Then
                        Cmbmone.ListIndex = 1
                    Else
                        Cmbmone.ListIndex = 2
                    End If
                End If
                txt_tc = Format$(!F4TIPCAM, "0.000") & ""
                txtcodforma = !F4FORPAG & ""
                Txt_Referencia = !F4REFERE & ""
                txtcodsoli = !F4CODSOL & ""
                If loc = 2 Then
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                Else
                    txtigv = Format$(!F4IGV & "", "#,##0.00")
                    txtmonto = Format$(!F4MONINA & "", "#,##0.00")
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                    txttotal = Format$(!F4MONTO & "", "#,##0.00")
                End If
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV,F2EMAIL from EF2PROVEEDORES where F2newruc='" & !F4CODPRV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    Txt_Prove.Text = "" & rst!F2NEWRUC
                    pnlnomprv.Caption = rst!f2nomprov
                    pnldireprv.Caption = IIf(IsNull(rst!f2dirprov), " ", rst!f2dirprov)
                    wemail_prove = "" & rst.Fields("F2EMAIL")
                    wgraba = 0
                Else
                    pnlnomprv.Caption = "Ruc es menor a 11 digitos"
                    pnldireprv.Caption = "No tiene "
                End If
                rst.Close
                
                xnombre = rsOrdenCab!F4CODSOL
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2NOMUSER from ef2userS where f2coduser='" & UCase(Trim(xnombre)) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    txtcodsoli = UCase(xnombre)
                    pnlnomsoli.Caption = rst!F2NOMUSER & ""
                End If
                rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    pnlnomforma.Caption = "" & rst.Fields("F2DESPAG")
                    wgraba = 0
                End If
                rst.Close
                txtcodcosto.Text = !F4CENTRO
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT f3descrip from centros where f3costo='" & txtcodcosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    pnlnomcosto.Caption = "" & rst.Fields("f3descrip")
                    wgraba = 0
                End If
                rst.Close
                
                txtuupp.Text = .Fields("F4UUPP") & ""
                If VALIDA_UUPP(txtuupp.Text) = True Then
                    txtdesuupp.Caption = wdeslocalidad
                End If
                
                txtlugar_entrega.Text = Left(Trim("" & !F4LUGAR_ENTREGA), 100)
                txtshipto.Text = Trim("" & !F4SHIPTO)
        
            Else
                MsgBox "La Solicitud de Compra no existe", vbInformation, "Atención"
                Txt_NumSolComp.Enabled = True
                Txt_NumSolComp.SetFocus
                Exit Sub
            End If
        End With
    End If
          
    With rsOrdenDet
        SQL = "SELECT * from if3orden where f4numord=" & GOC
        If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
        rsOrdenDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
            
            If sw_nuevo_documento = False Then
                DELETEREC_N cnomtabla, cnn_form
                AdicionaItem
                sw_nuevo_documento = True
            End If
            
            dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
            dxDBGrid1.Dataset.Active = True
        
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
            dxDBGrid1.OptionEnabled = False
            dxDBGrid1.Dataset.DisableControls
            'dxDBGrid1.Dataset.EnableControls
            'dxDBGrid1.Dataset.Close
            'dxDBGrid1.Dataset.Open
            'dxDBGrid1.OptionEnabled = True
            
            sw_nuevo_temp = False
            sw_nuevo_item = True
            
            .MoveFirst
            Do While Not .EOF
                i = i + 1
                If loc = 1 Then
                    If rsOrdenDet.Fields("f4numord") = Txt_NumOC Then
                        If sw_nuevo_temp = False Then
                            If sw_nuevo_documento = True Then
                                dxDBGrid1.Dataset.Edit
                            Else
                                dxDBGrid1.Dataset.Append
                            End If
                            sw_nuevo_temp = True
                        Else
                            dxDBGrid1.Dataset.Append
                        End If
                
                        dxDBGrid1.Dataset.FieldValues("item") = i
                        dxDBGrid1.Dataset.FieldValues("f3codpro") = .Fields("f3codpro") & ""
                        dxDBGrid1.Dataset.FieldValues("f5nompro") = .Fields("f5nompro") & ""
                        dxDBGrid1.Dataset.FieldValues("f5nompro_ing") = .Fields("f5nompro_ing") & ""
                        dxDBGrid1.Dataset.FieldValues("f5codfab") = .Fields("f3codfab") & ""
                        If rst.State = adStateOpen Then rst.Close
                        rst.Open "SELECT f5nompro,f5codfab,F7CODMED,f5marca from if5pla where f5codfab='" & rsOrdenDet!F3CODFAB & "' and f5marca='" & rsOrdenDet!f5codmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not (rst.EOF) Then
                            dxDBGrid1.Dataset.FieldValues("f5codfab") = rst!F5CODFAB & ""
                            dxDBGrid1.Dataset.FieldValues("f3medida") = rst!F7CODMED & ""
                            cmarca = rst!F5MARCA & ""
                        End If
                        rst.Close
                        
                        If rsmarcas.State = adStateOpen Then rsmarcas.Close
                        rsmarcas.Open "SELECT F2DESMAR,F2CODMAR FROM EF2MARCAS WHERE F2CODMAR='" & cmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsmarcas.EOF Then
                            dxDBGrid1.Dataset.FieldValues("f5marca") = rsmarcas.Fields("F2DESMAR")
                            dxDBGrid1.Dataset.FieldValues("f5codmarca") = rsmarcas.Fields("F2CODMAR")
                        Else
                            dxDBGrid1.Dataset.FieldValues("f5marca") = ""
                            dxDBGrid1.Dataset.FieldValues("f5CODmarca") = ""
                        End If
                        rsmarcas.Close
                        Set rsmarcas = Nothing
                            
                        dxDBGrid1.Dataset.FieldValues("f3canpro") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("f3precos") = Format$(.Fields("f3precos"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3pordct") = Format$(.Fields("f3pordct"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3totdct") = Format$(.Fields("f3totdct"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5valvta") = Format$(.Fields("f5valvta"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5afecto") = .Fields("f5afecto")
                        dxDBGrid1.Dataset.FieldValues("f3igv") = Format$(.Fields("f3igv"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3preuni") = Format$(.Fields("f3preuni"), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3total") = Format$(.Fields("f3total"), "###,##0.00")
                        If Not (IsDate(rsOrdenDet!f3fentrega)) Then
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(txt_fecha.Value, "dd/mm/yyyy"))
                        Else
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = Format(rsOrdenDet!f3fentrega, "dd/mm/yyyy")
                        End If
                        dxDBGrid1.Dataset.FieldValues("check") = True
                        dxDBGrid1.Dataset.FieldValues("cant_ant") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("F3CANPRO2") = .Fields("F3CANPRO2")
                        dxDBGrid1.Dataset.FieldValues("BACKORDER") = .Fields("F3BACKORDER")
                        dxDBGrid1.Dataset.FieldValues("COD_SOLICITUD") = .Fields("COD_SOLICITUD")
                        dxDBGrid1.Dataset.FieldValues("F5CODCOSTO") = .Fields("F5CODCOSTO")
                        dxDBGrid1.Dataset.FieldValues("F5DESCOSTO") = .Fields("F5DESCOSTO")
                    Else
                        Exit Do
                    End If
                    If rsOrdenCab!F4LOCAL <> "1" Then
                        cmbtipopera.ListIndex = 1
                        Forma_Imp
                    Else
                        cmbtipopera.ListIndex = 0
                        Forma_Loc
                    End If
                End If
                .MoveNext
            Loop
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = False
            jc = 1
        End If
        rsOrdenDet.Close
    End With
    
    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = True
    
End Sub

Private Sub CONFIGURA_GRID()
    
'''    With dxDBGrid1.Options
'''        .Set (egoEditing)
'''        .Set (egoTabs)
'''        .Set (egoTabThrough)
'''        .Set (egoCanDelete)
'''        .Set (egoCanAppend)
'''        .Set (egoCanInsert)
'''        .Set (egoImmediateEditor)
'''        .Set (egoShowIndicator)
'''        .Set (egoCanNavigation)
'''        .Set (egoHorzThrough)
'''        .Set (egoVertThrough)
'''        .Set (egoAutoWidth)
'''        .Set (egoEnterShowEditor)
'''        .Set (egoEnterThrough)
'''        .Set (egoShowButtonAlways)
'''
'''        .Set (egoColumnSizing)
'''        .Set (egoColumnMoving)
'''        .Set (egoTabThrough)
'''        .Set (egoConfirmDelete)
'''        .Set (egoCanNavigation)
'''        .Set (egoCancelOnExit)
'''        .Set (egoLoadAllRecords)
'''        .Set (egoShowHourGlass)
'''        .Set (egoUseBookmarks)
'''        .Set (egoUseLocate)
'''        .Set (egoAutoCalcPreviewLines)
'''        .Set (egoBandSizing)
'''        .Set (egoBandMoving)
'''        .Set (egoDragScroll)
'''        .Set (egoExpandOnDblClick)
'''        .Set (egoShowFooter)
'''        .Set (egoShowGrid)
'''        .Set (egoShowButtons)
'''        .Set (egoNameCaseInsensitive)
'''        .Set (egoShowHeader)
'''        .Set (egoShowPreviewGrid)
'''        .Set (egoShowBorder)
'''        .Set (egoDynamicLoad)
'''    End With

    With dxDBGrid1.Options
        .Set (egoAutoExpandOnSearch)
        .Set (egoAutoSort)
        '.Set (egoAutoWidth)
        .Set (egoBandHeaderWidth)
        .Set (egoBandMoving)
        .Set (egoBandSizing)
        .Set (egoCanAppend)
        .Set (egoCancelOnExit)
        .Set (egoCanDelete)
        .Set (egoCanInsert)
        .Set (egoCanNavigation)
        .Set (egoColumnMoving)
        .Set (egoColumnSizing)
        .Set (egoConfirmDelete)
        .Set (egoDragScroll)
        '.Set (egoDynamicLoad)
        .Set (egoEditing)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoExactScrollBar)
        .Set (egoExpandOnDblClick)
        .Set (egoHorzThrough)
        .Set (egoImmediateEditor)
        .Set (egoLoadAllRecords)
        .Set (egoNameCaseInsensitive)
        '.Set (egoShowBands)
        .Set (egoShowBorder)
        .Set (egoShowButtonAlways)
        .Set (egoShowButtons)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowHeader)
        .Set (egoShowHourGlass)
        .Set (egoShowIndicator)
        .Set (egoShowPreviewGrid)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoTabThrough)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoVertThrough)
    End With

    dxDBGrid1.Columns(0).Visible = False
    
    If wf1visualiza_dctos = "*" Then
        dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
    End If
    
    dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO2").Visible = False
    
End Sub

Sub Nueva_orden()
Dim SQL     As String
Dim Orden   As String

    SQL = "select f4numord from if4orden where f4inicial<>'*' OR ISNULL(F4INICIAL) ORDER BY VAL(F4NUMORD) DESC"
    If rst.State = adStateOpen Then rst.Close
    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not (rst.EOF) Then
        Orden = rst.Fields("f4numord") + 1
    Else
        Orden = 1
    End If
    Txt_NumOC.Text = Format$(Orden, "000000")
    
End Sub

Sub GrabarOC()
Dim codi                As String
Dim wcantidad           As Double
Dim wcc                 As String
Dim wproducto           As String
Dim SQL                 As String
Dim ocompra             As Double
Dim Cant                As Double
Dim rsdetaoc            As New ADODB.Recordset
Dim ncant_ant           As Double
Dim amovs_cab(0 To 27)  As a_grabacion
Dim ctipo               As String
Dim csql                As String

    flag = 0
    If Trim(Txt_NumOC.Text) <> "" Then
        jc = 1
    Else
        jc = 0
    End If
    
    'Nueva Versión
    If loc = 1 Then
        Select Case jc
            Case 0
                Call Nueva_orden
        End Select
    End If
    
    SQL = "select * from detalle where check and (Not detalle.f3codpro Is Null)"

    If rst.State = adStateOpen Then rst.Close
    rst.Open SQL, cnn_form, adOpenStatic, adLockOptimistic
    If rst.EOF Then
        MsgBox "Debe Ingresar y/o Seleccionar Productos a Comprar", vbInformation, "Sistema de Logística"
        dxDBGrid1.SetFocus
        rst.Close
        Exit Sub
    End If
    rst.Close
    
    If loc = 1 Then
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL,F4NUMINTERNO from if4orden where f4numord=" & Txt_NumOC, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (rsOrdenCab.EOF) Then
            ctipo = "M"
        Else
            ctipo = "A"
            flag = 1
            sw_nuevo_documento = False
        End If
        
        amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = Txt_NumOC.Text: amovs_cab(0).TIPO = "N"
        If ctipo = "A" Then
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = "N": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = "1": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = 0: amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECGRA": amovs_cab(4).valor = Format(Date, "dd/MM/yyyy"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEGRA": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
            xprov = NumXProv(Txt_Prove.Text)
            xnumero = wanno & "-" & Txt_NumOC.Text & "-" & Format(xprov, "000")
            pnlnumero.Caption = xnumero
            amovs_cab(26).campo = "F4NUMINTERNO": amovs_cab(26).valor = xnumero: amovs_cab(26).TIPO = "T"
        Else
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = rsOrdenCab.Fields("F4ESTNUL") & "": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = rsOrdenCab.Fields("F4FALTA") & "": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = rsOrdenCab.Fields("F4ESTVAL") & "": amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECMOD": amovs_cab(4).valor = Format(Date, "dd/MM/yyyy"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEMOD": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
            amovs_cab(26).campo = "F4NUMINTERNO": amovs_cab(26).valor = rsOrdenCab.Fields("F4NUMINTERNO") & "": amovs_cab(26).TIPO = "T"
        End If
        
        amovs_cab(6).campo = "F4CODSOL": amovs_cab(6).valor = txtcodsoli.Text: amovs_cab(6).TIPO = "T"
        amovs_cab(7).campo = "F4FECEMI": amovs_cab(7).valor = Format(txt_fecha.Value, "DD/MM/YYYY"): amovs_cab(7).TIPO = "F"
        amovs_cab(8).campo = "F4CODPRV": amovs_cab(8).valor = Txt_Prove: amovs_cab(8).TIPO = "T"
        amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = txt_tc.Text: amovs_cab(9).TIPO = "N"
        amovs_cab(10).campo = "F4FORPAG": amovs_cab(10).valor = txtcodforma.Text: amovs_cab(10).TIPO = "T"
        amovs_cab(11).campo = "F4REFERE": amovs_cab(11).valor = Txt_Referencia.Text: amovs_cab(11).TIPO = "T"
        amovs_cab(12).campo = "F4OBSERVA": amovs_cab(12).valor = txtobserva.Text: amovs_cab(12).TIPO = "T"
        amovs_cab(13).campo = "F4CENTRO": amovs_cab(13).valor = txtcodcosto.Text: amovs_cab(13).TIPO = "T"
        amovs_cab(14).campo = "F4CODSOLICITUD": amovs_cab(14).valor = Trim(Txt_NumSolComp.Text): amovs_cab(14).TIPO = "T"
        amovs_cab(15).campo = "F4TIPMON": amovs_cab(15).valor = IIf(Cmbmone.ListIndex = 0, "S", "D"): amovs_cab(15).TIPO = "T"
        amovs_cab(16).campo = "F4IGV": amovs_cab(16).valor = Val(Format(txtigv.Text, "0.00")): amovs_cab(16).TIPO = "N"
        amovs_cab(17).campo = "F4MONINA": amovs_cab(17).valor = Val(Format(txtmonto.Text, "0.00")): amovs_cab(17).TIPO = "N"
        amovs_cab(18).campo = "F4BASIMP": amovs_cab(18).valor = Val(Format(txtbase.Text, "0.00")): amovs_cab(18).TIPO = "N"
        amovs_cab(19).campo = "F4MONTO": amovs_cab(19).valor = Val(Format(dxDBGrid1.Columns.ColumnByName("F3TOTAL").SummaryFooterValue, "0.00")): amovs_cab(19).TIPO = "N"
        amovs_cab(20).campo = "F4LOCAL": amovs_cab(20).valor = IIf(cmbtipopera.ListIndex = 0, 1, 0): amovs_cab(20).TIPO = "T"
        amovs_cab(21).campo = "F4EMPRESA": amovs_cab(21).valor = wnomcia: amovs_cab(21).TIPO = "T"
        amovs_cab(22).campo = "F4UUPP": amovs_cab(22).valor = txtuupp.Text: amovs_cab(22).TIPO = "T"
        amovs_cab(23).campo = "F4PLAZO_ENTREGA": amovs_cab(23).valor = txtplazo_entrega: amovs_cab(23).TIPO = "T"
        amovs_cab(24).campo = "F4LUGAR_ENTREGA": amovs_cab(24).valor = txtlugar_entrega.Text: amovs_cab(24).TIPO = "T"
        amovs_cab(25).campo = "F4CONTACTO": amovs_cab(25).valor = txtcontacto.Text: amovs_cab(25).TIPO = "T"
        amovs_cab(27).campo = "F4SHIPTO": amovs_cab(27).valor = txtshipto.Text: amovs_cab(27).TIPO = "T"
        
        rsOrdenCab.Close
        
        If ctipo = "A" Then     '--- Nuevo
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 27, cnn_dbbancos, ""
        Else
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 27, cnn_dbbancos, "F4NUMORD = " & Txt_NumOC.Text & ""
        End If
        
    End If
    
    '---------- GRABANDO EL DETALLE DE LA ORDEN DE COMPRA ----------------------'
    
    cnn_dbbancos.Execute ("delete * from if3orden where f4numord= " & Txt_NumOC.Text)
    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
    rsOrdenDet.Open "select * from if3orden", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    rsdetaoc.Open "SELECT * FROM " & cnomtabla & " where check and (Not detalle.f3codpro Is Null)", cnn_form, adOpenStatic, adLockOptimistic
    If Not rsdetaoc.EOF Then
        With rsdetaoc
            .MoveFirst
            Do While Not .EOF
                If .Fields("check") = True And Len(Trim(.Fields("f3codpro") & "")) > 0 Then
                    wgrabar = True
                Else
                    wgrabar = False
                End If
                
                If wgrabar Then
                    rsOrdenDet.AddNew
                    rsOrdenDet!f4numord = Txt_NumOC.Text
                    rsOrdenDet!F3CODFAB = .Fields("F5CODFAB") & ""
                    rsOrdenDet!F3CODPRO = .Fields("f3codpro") & ""
                    rsOrdenDet!F5NOMPRO = Trim(.Fields("F5NOMPRO") & "")
                    rsOrdenDet!F5NOMPRO_ING = Trim(.Fields("F5NOMPRO_ING") & "")
                    codi = .Fields("f3codpro") & ""
                    rsOrdenDet!F3CANPRO = .Fields("f3canpro")
                    rsOrdenDet!F5MARCA = "" & .Fields("f5marca")
                    rsOrdenDet!f5codmarca = "" & .Fields("f5codmarca")
                    
                    'Actualiza Centro de Productos
                    wcantidad = Val("" & .Fields("f3canpro"))
                    wcc = Trim$(txtcodcosto.Text)
                    wproducto = Trim$(codi)
                    
                    'SQL = "select f3presu,f3consumido,f3ocompra from centroproductos where " _
                    '& "f3costo='" & wcc & "' and f5codpro='" & wproducto & "'"
                    'If rstaux.State = adStateOpen Then rst.Close
                    'rstaux.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    'If Not (rstaux.EOF) Then
                    '    If jc = 0 Then  'Nuevo
                    '        ocompra = Val(rstaux.Fields("f3ocompra").Value)
                    '        rstaux.Fields("f3ocompra").Value = ocompra + wcantidad
                    '    Else             'Modifica
                    '        rstaux.Fields("f3ocompra").Value = wcantidad
                    '    End If
                    '    rstaux.Update
                    'End If
                    'rstaux.Close
                    
                    rsOrdenDet!f3canfal = .Fields("f3canpro")
                    rsOrdenDet!F3PREUNI = .Fields("f3preuni")
                    rsOrdenDet!f3PRECOS = .Fields("f3precos")
                    rsOrdenDet!F3PORDCT = .Fields("f3pordct")
                    rsOrdenDet!f3totdct = .Fields("f3totdct")
                    rsOrdenDet!f5valvta = .Fields("f5valvta")
                    rsOrdenDet!F5AFECTO = IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto"))
                    rsOrdenDet!F3IGV = .Fields("f3igv")
                    rsOrdenDet!F3TOTAL = .Fields("f3total")
                    rsOrdenDet!f3fentrega = Format$(.Fields("f3fentrega"), "dd/mm/yyyy")
                    rsOrdenDet!F3CANPRO2 = .Fields("f3canpro2")
                    rsOrdenDet!F3BACKORDER = .Fields("BACKORDER")
                    rsOrdenDet!cod_solicitud = "" & .Fields("COD_SOLICITUD")
                    rsOrdenDet!f5codcosto = "" & .Fields("f5codcosto")
                    rsOrdenDet!f5descosto = "" & .Fields("f5descosto")
                    rsOrdenDet.Update
                End If
                .MoveNext
            Loop
            rsOrdenDet.Close
        End With
    End If
    wcuenta = rsdetaoc.RecordCount
    rsdetaoc.Close
    
    '-------------------------------------------------------------------------------
    csql = "UPDATE IF3ORDEN INNER JOIN IF5PLA ON (IF3ORDEN.F3CODFAB = IF5PLA.F5CODFAB) AND " & _
           "(IF3ORDEN.F5CODMARCA = IF5PLA.F5MARCA) " & _
           "SET IF5PLA.F5TEXTO = IF3ORDEN.F5NOMPRO, IF5PLA.F5TEXTO_ING = IF3ORDEN.F5NOMPRO_ING " & _
           "WHERE (Len(Trim(IF3ORDEN.F5NOMPRO))>0 OR Len(Trim(IF3ORDEN.F5NOMPRO_ING))>0) AND " & _
           "(IF3ORDEN.F4NUMORD=" & Txt_NumOC.Text & ")"
    cnn_dbbancos.Execute (csql)
    '-------------------------------------------------------------------------------
    
    Call VERIFIC_PPRV
    'If Txt_NumSolComp.Text <> "" Then
    If wcuenta > 0 Then
        If Txt_NumSolComp.Text <> "" Then
        If wcuenta = 1 And Len(Trim(Txt_NumSolComp.Text)) > 0 Then
            SQL = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC.Text & "',f4cerrado='S' where cod_solicitud='" & Txt_NumSolComp.Text & "'"
        Else
            SQL = "update tb_cabsolicitud set f4cerrado='S' where cod_solicitud='" & Txt_NumSolComp.Text & "'"
        End If
        cnn_dbbancos.Execute SQL
        End If

        If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
        rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
        If Not rsdetaoc.EOF Then
            With rsdetaoc
                .MoveFirst
                Do While Not .EOF
                    If Len(Trim(.Fields("f3codpro") & "")) > 0 Then
                        codprod = .Fields("f3codpro")
                        codsolicitud = "" & .Fields("cod_solicitud")
                        wf5codmarca = "" & .Fields("f5codmarca")
                        wf5codfab = "" & .Fields("f5codfab")
                        If .Fields("check") = True Then
                            Cant = Val("" & .Fields("f3canpro"))
                            ncant_ant = Val("" & .Fields("cant_ant"))
                            cnn_dbbancos.Execute "update tb_detsolicitud set candis= candis+" & ncant_ant & "-" & _
                            Cant & " where cod_solicitud='" & _
                            codsolicitud & "' and f5codfab='" & wf5codfab & "' and f5codmarca='" & wf5codmarca & "'"
                        End If
                    End If
                                        
                    If Len(Trim(codsolicitud)) > 0 Then
                        If rst.State = adStateOpen Then rst.Close
                        rst.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & codsolicitud & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If rst!Cant <= 0 Then
                            cnn_dbbancos.Execute "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & codsolicitud & "'"
                        End If
                    End If

                    .MoveNext
                Loop
                
                'If rst.State = adStateOpen Then rst.Close
                'rst.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                'If rst!Cant <= 0 Then
                '    cnn_dbbancos.Execute "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & Txt_NumSolComp & "'"
                'End If
                'If rst.State = adStateOpen Then rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                wgraba = 1
            End With
        End If
        rsdetaoc.Close
    End If
    
    atbmenu.Tools("idanular").Enabled = True
    atbmenu.Tools("idemail").Enabled = True
    atbmenu.Tools("Id_CtasxPagar").Enabled = True
    atbmenu.Tools("idimprimir").Enabled = True
    
    MsgBox "Orden de Compra Actualizada", vbInformation, "Orden de Compra"
    swGrabacion = False
    
End Sub

Private Sub VERIFIC_PPRV()
Dim CodProv     As String
Dim NOMPROV     As String
Dim NomProd     As String
Dim rsdetaoc    As New ADODB.Recordset
Dim SQL         As String
Dim cmoneda     As String
Dim dfecha      As Date
Dim ccodfab     As String
Dim ccodmed     As String
Dim nprecos     As Double

    If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
    If Not rsdetaoc.EOF Then
        With rsdetaoc
            .MoveFirst
            Do While Not .EOF
                CodProv = Txt_Prove.Text
                NOMPROV = pnlnomprv.Caption
                codprod = .Fields("f3codpro") & ""
                NomProd = .Fields("f5nompro") & ""
                cmoneda = IIf(Cmbmone.ListIndex = 0, "S", "D")
                dfecha = Format(txt_fecha.Value, "DD/MM/YYYY")
                nprecos = Val("" & .Fields("F3PRECOS"))
                If rsproductos.State = adStateOpen Then rsproductos.Close
                rsproductos.Open "SELECT F5CODFAB,F7codmed FROM IF5PLA WHERE F5CODPRO='" & codprod & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsproductos.EOF Then
                    ccodfab = Left("" & rsproductos.Fields("F5CODFAB"), 15)
                    ccodmed = "" & rsproductos.Fields("F7codmed")
                End If
                rsproductos.Close
                    
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT * FROM EF2PROD_PROV WHERE F5CODPRO='" & codprod & "' AND " _
                & "F2CODPRV='" & CodProv & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rst.EOF Then
                    rst.AddNew
                    rst!F2CODPRV = CodProv
                    rst!F2NOMPRV = NOMPROV
                    rst!F5CODPRO = codprod
                    rst!F5NOMPRO = NomProd
                    rst!f5valvta = nprecos
                    rst.Fields("F2MONEDA") = cmoneda
                    rst.Fields("F2FECHA") = dfecha
                    rst!F5CODFAB = ccodfab
                    rst!F7CODMED = ccodmed
                    rst.Fields("F2COND_PAGO") = txtcodforma.Text
                    rst.Fields("F2FORPAG") = txtcodforma.Text
                    rst.Update
                Else
                    SQL = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA=CVDATE('" & dfecha & "') WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CodProv & "'"
                    cnn_dbbancos.Execute (SQL)
                End If
                rst.Close
                .MoveNext
            Loop
        End With
    End If
    rsdetaoc.Close
    
End Sub

Private Sub Txt_Prove_DblClick()

    Txt_Prove_KeyDown 113, 0
    
End Sub

Private Sub Txt_Prove_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Screen.MousePointer = vbHourglass
        sw_ayuda = True
        sw_ocompra = True
        hlp_proveedores.Show 1
        sw_ocompra = False
        sw_ayuda = False
        Txt_Prove.Text = wrucprov
        pnlnomprv.Caption = wnomprov
        pnldireprv.Caption = wdirprov
        txtcontacto.Text = wcontacto
        If Len(Trim(wfpagoprov)) > 0 Then
            txtcodforma.Text = wfpagoprov
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                pnlnomforma.Caption = Trim("" & rst.Fields("F2DESPAG"))
            End If
            rst.Close
        End If
        Txt_Prove_KeyPress 13
    End If
    
End Sub

Private Sub Txt_Prove_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtcontacto.SetFocus
    End If
    
End Sub

Private Sub Txt_Referencia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        txtobserva.SetFocus
    End If
    
End Sub

Private Sub Txt_TC_GotFocus()
    
    txt_tc.SelStart = 0
    txt_tc.SelLength = Len(txt_tc.Text)
    
End Sub

Private Sub Txt_TC_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        If txt_tc = "" Then
            MsgBox "Ingrese tipo de cambio", 48, "Sistema de Logística"
            txt_tc.Text = 0#
            txt_tc.SetFocus
            Exit Sub
        End If
        txt_tc = Format(txt_tc, "#0.000")
        txtcodforma.SetFocus
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub Txt_TC_LostFocus()
 
    If Not IsNumeric(txt_tc) Then
        MsgBox "Dato mal ingresado ...Verifique!", vbInformation, "Sistema de Logistica"
        txt_tc.SetFocus
    End If
    
End Sub

Private Sub txtcodcosto_DblClick()
    
    txtcodcosto_KeyDown 113, 0
    
End Sub

Private Sub txtcodcosto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodcosto = ""
        sw_ayuda = True
        'hlp_centros.Show 1
        ayuda_centros.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            txtcodcosto = wcodcosto
            pnlnomcosto = wdescosto
            txtcodcosto_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtcodforma_DblClick()
    
    txtcodforma_KeyDown 113, 0

End Sub

Private Sub txtcodforma_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtplazo_entrega.SetFocus
    End If
    
End Sub

Private Sub txtcodsoli_Change()

    If Not inicio Then swGrabacion = True
    
End Sub

Private Sub txtcodsoli_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Cmbmone.SetFocus
    End If
    
End Sub

Private Sub imprimir()
    
    LLENA_TEMPCAB
    'Cryordcompra.DataFiles(0) = wrutatemp & "\TEMPCOMP.MDB"
    'Cryordcompra.ReportFileName = wrutatemp & "\OCOMPRA_AIC.RPT"
    'Cryordcompra.Destination = crptToWindow
    'Cryordcompra.Action = 0
    acr_ocompra.Show 1

End Sub

Private Sub eliminar()
Dim gcodigo     As String
Dim gcant       As Double
    
    If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
    rsOrdenCab.Open "SELECT * from if4orden where f4numord=" & Txt_NumOC, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsOrdenCab.EOF Then
        If MsgBox("¿Está Seguro de Anular La Orden de Compra Nº " & Txt_NumOC.Text & "?", vbDefaultButton2 + vbYesNo + vbQuestion, "Sistema de Logística") = vbYes Then
            cnn_dbbancos.Execute "Update if4ORDEN set f4estnul='S' where F4NUMORD=" & Txt_NumOC.Text & ""
            With dxDBGrid1
                .Dataset.First
                If Not (.Dataset.EOF) Then
                    .Dataset.First
                    Do While Not (.Dataset.EOF)
                        gcodigo = .Dataset.FieldValues("f3codpro")
                        gcant = .Dataset.FieldValues("f3canpro")
                        If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
                        rsOrdenDet.Open "select * from tb_detsolicitud where " _
                        & "cod_solicitud='" & Txt_NumSolComp.Text & "' and cod_producto='" & _
                        gcodigo & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        
                        If Not (rsOrdenDet.EOF) Then
                            rsOrdenDet.Fields("candis") = rsOrdenDet.Fields("candis") + Val(gcant)
                            rsOrdenDet.Update
                        End If
                        .Dataset.Next
                    Loop
                    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
                    rsOrdenDet.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If rsOrdenDet(0).Value = 0 Then
                        cnn_dbbancos.Execute "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & Txt_NumSolComp & "'"
                    Else
                        cnn_dbbancos.Execute "update tb_cabsolicitud set cs_estado='P' where cod_solicitud='" & Txt_NumSolComp & "'"
                    End If
                    rsOrdenDet.Close
                    MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbInformation, "Sistema de Logística"
                    Call Visi
                    Call Limpia_Orden
                    sw_nuevo_documento = False
                    AdicionaItem
                    AdicionaItem
                    sw_nuevo_documento = True
                    Call Limpiar
                    txt_fecha.SetFocus
                End If
            End With
        End If
    End If
    rsOrdenCab.Close
    
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.FocusedIndex = 0
        End If
        If Action = daPost Then
            calcula
        End If
    End If

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim rsproduc    As New ADODB.Recordset

    sw_detalle = True
    
    If sw_nuevo_item = False Then
        If Not inicio Then swGrabacion = True
        If dxDBGrid1.Columns.FocusedIndex = 2 Or dxDBGrid1.Columns.FocusedIndex = 3 Or dxDBGrid1.Columns.FocusedIndex = 4 Or dxDBGrid1.Columns.FocusedIndex = 5 Or dxDBGrid1.Columns.FocusedIndex = 8 Then
            dxDBGrid1.Dataset.Edit
            Calcula_PvtaTot
            sw_nuevo_item = True
            dxDBGrid1.Dataset.Post
            dxDBGrid1.Dataset.Refresh
            sw_nuevo_item = False
            calcula
        Else
            If dxDBGrid1.Columns.FocusedColumn.FieldName = "F5CODFAB" Then
                If pnlnomprv.Caption = "" Then
                    MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                    Txt_Prove.SetFocus
                    Exit Sub
                End If
                
                wdxcodigo = dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value
                wdxcodfab = dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value
                wdxdescripcion = dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value
                wdxcantidad = Val("" & dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO").Value)
                Select Case dxDBGrid1.Columns.FocusedIndex
                    Case 0:
                        dxDBGrid1.Dataset.Edit
                        If rsproduc.State = adStateOpen Then rsproduc.Close
                        rsproduc.Open "SELECT B.F5CODPRO,B.F5TEXTO_ING,B.F5TEXTO,B.F5NOMPRO,B.F5TEXTO_ING,B.F5AFECTO,B.F5CODFAB,B.F5VALVTA,B.F7CODMED,B.F5MARCA,F5FOB FROM EF2PROD_PROV AS A,IF5PLA AS B WHERE A.F2CODPRV='" & wrucprov & "' AND (B.F5CODFAB='" & wdxcodfab & "') ORDER BY B.F5CODPRO", cnn_dbbancos, adOpenStatic, adLockReadOnly
                        If Not rsproduc.EOF Then
                            dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = rsproduc.Fields("F5CODPRO") & ""
                            If Len(Trim(rsproduc.Fields("F5TEXTO")) & "") > 0 Or Len(Trim(rsproduc.Fields("F5TEXTO_ING")) & "") > 0 Then
                                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5TEXTO") & ""
                                dxDBGrid1.Columns.ColumnByFieldName("f5nompro_ing").Value = rsproduc.Fields("F5TEXTO_ING") & ""
                            Else
                                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5NOMPRO") & ""
                                dxDBGrid1.Columns.ColumnByFieldName("f5TEXTO_ING").Value = rsproduc.Fields("F5TEXTO_ING") & ""
                            End If
                            dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = rsproduc.Fields("F7CODMED") & ""
                            If rsmarcas.State = adStateOpen Then rsmarcas.Close
                            rsmarcas.Open "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & rsproduc.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                            If Not rsmarcas.EOF Then
                                dxDBGrid1.Columns.ColumnByFieldName("f5codmarca").Value = rsmarcas.Fields("F2CODMAR")
                                dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                            Else
                                dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsproduc.Fields("F5MARCA") & ""
                            End If
                            dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = rsproduc.Fields("F5AFECTO") & ""
                            dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Value = Format(Val("" & rsproduc.Fields("F5FOB")), "###,##0.00")
                            dxDBGrid1.Dataset.FieldValues("f5valvta") = Val(rsproduc.Fields("F5VALVTA") & "")
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format$(Date, "DD/MM/YYYY"))
                            dxDBGrid1.Dataset.FieldValues("check") = True
                            dxDBGrid1.Dataset.Post
                        End If
                        rsproduc.Close
                        Set rsproduc = Nothing
                        dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("check").ColIndex - 1
                End Select
            End If
        End If
        If UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) = "F3PRECOS" Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
            dxDBGrid1.Dataset.Post
        End If
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
        dxDBGrid1.Columns.ColumnByFieldName("BACKORDER").Value = Val("" & dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO2").Value) - Val("" & dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO").Value)
        dxDBGrid1.Dataset.Post
    End If
    wdxcantidad = Val("" & dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO").Value)

End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid1.Dataset.Delete
        End If
    End If

End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer
    
    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_documento = False Then
        DELETEREC_N cnomtabla, cnn_form
        dxDBGrid1.Dataset.Refresh
    End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    'dxDBGrid1.OptionEnabled = False
    'dxDBGrid1.Dataset.DisableControls
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
            .FieldValues("item") = i
            .FieldValues("f3codpro") = ""
            .FieldValues("f5nompro") = ""
            .FieldValues("f5nompro_ing") = ""
            .FieldValues("f3medida") = ""
            .FieldValues("f5marca") = ""
            .FieldValues("f5codmarca") = ""
            .FieldValues("f3canpro") = Null
            .FieldValues("f3precos") = Null
            .FieldValues("f3pordct") = Null
            .FieldValues("f3totdct") = Null
            .FieldValues("f5valvta") = Null
            .FieldValues("f5afecto") = ""
            .FieldValues("f3igv") = Null
            .FieldValues("f3preuni") = Null
            .FieldValues("f3total") = Null
            .FieldValues("f5codfab") = ""
            .FieldValues("f3fentrega") = Format$(Date, "dd/mm/yyyy")
            .FieldValues("check") = False
            .FieldValues("cant_ant") = 0#
            .FieldValues("f3canpro2") = Null
        Next
        .Post
        sw_nuevo_item = False
    End With
    'dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    'dxDBGrid1.OptionEnabled = True

End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim SQL         As String
    
    Select Case dxDBGrid1.Columns.FocusedIndex
        Case 0:
            If pnlnomprv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
            
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(pnlnomprv.Caption)
            hlp_prov_prod.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3MEDIDA").Value = wmedida
                                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5codmarca").Value = rsmarcas.Fields("F2CODMAR")
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5codmarca").Value = ""
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = ""
                End If
                rsmarcas.Close
                Set rsmarcas = Nothing
                        
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = Format$(Date, "dd/mm/yyyy")
                If rsconsulta.State = adStateOpen Then rsconsulta.Close
                rsconsulta.Open "SELECT F5TEXTO_ING,F5AFECTO,F5FOB FROM IF5PLA WHERE F5CODPRO = '" & dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
                If Not rsconsulta.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Value = "" & rsconsulta.Fields("F5AFECTO")
                    dxDBGrid1.Dataset.FieldValues("F3PRECOS") = Format(Val("" & rsconsulta.Fields("F5FOB")), "###,##0.00")
                    If Val("" & rsconsulta.Fields("F5FOB")) > 0# Then
                        dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
                    End If
                Else
                    dxDBGrid1.Dataset.FieldValues("F3PRECOS") = "0.00"
                End If
                dxDBGrid1.Columns.ColumnByFieldName("F5TEXTO_ING").Value = "" & rsconsulta.Fields("F5TEXTO_ING")
                rsconsulta.Close
                Set rsconsulta = Nothing
            End If
    End Select
End Sub

Private Sub txtplazo_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Txt_Referencia.SetFocus
    End If

End Sub

Private Sub txtusuario_Change()
    
    If Not inicio Then swGrabacion = True

End Sub

Private Sub MODIFICAR_OC()

    flagwin = True
    Wnuevo = False
    Txt_NumOC.Text = GOC
    With rsOrdenCab
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        rsOrdenCab.Open "SELECT * from if4orden where f4numord=" & GOC, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
            MostrarDatosOC
            ExisteOrdenCompra = True
        Else
            ExisteOrdenCompra = False
        End If
        .Close
    End With

End Sub

Private Sub txtuupp_DblClick()

    txtuupp_KeyDown 113, 0
    
End Sub

Private Sub txtuupp_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodlocalidad = "": wdeslocalidad = ""
        hlp_uupp.Show 1
        If Len(Trim(wcodlocalidad)) > 0 Then
            txtuupp.Text = Trim(wcodlocalidad)
            txtdesuupp.Caption = Trim(wdeslocalidad)
            txtuupp_KeyPress 13
        End If
    End If

End Sub

Private Sub txtuupp_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
    End If

End Sub

Private Sub txtuupp_LostFocus()

    If Len(Trim(txtuupp.Text)) > 0 Then
        If VALIDA_UUPP(txtuupp.Text) = True Then
            txtdesuupp.Caption = wdeslocalidad
        Else
            MsgBox "Unidad de producciòn no existe", vbInformation + vbDefaultButton1, "Atención"
            txtuupp.Text = "": txtuupp.SetFocus
        End If
    End If

End Sub

Private Sub GRABA_GRID(prucprov As String)
Dim rsprodprov      As New ADODB.Recordset
Dim rstempdet       As New ADODB.Recordset
Dim csql            As String
Dim nitem           As Integer
Dim ccodprod        As String
Dim nprecos         As Double
Dim totdcto         As Double
Dim ValVta          As Double
Dim IGV             As Double
Dim preciounit      As Double
Dim TOTAL           As Double
    
    If rstempdet.State = adStateOpen Then rstempdet.Close
    rstempdet.Open "SELECT * FROM " & cnomtabla & " ORDER BY F3CODPRO", cnn_form, adOpenDynamic, adLockBatchOptimistic
    If Not rstempdet.EOF Then
        rstempdet.MoveFirst
        Do While Not rstempdet.EOF
            nitem = Val(rstempdet.Fields("ITEM") & "")
            ccodprod = Trim(rstempdet.Fields("F3CODPRO") & "")
            If rsprodprov.State = adStateOpen Then rsprodprov.Close
            rsprodprov.Open "SELECT F5VALVTA FROM EF2PROD_PROV WHERE F5CODPRO='" & ccodprod & "' AND F2CODPRV='" & prucprov & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsprodprov.EOF Then
                nprecos = Val(rsprodprov.Fields("F5VALVTA") & "")
                If Trim(rstempdet.Fields("F5AFECTO") & "") = "*" Then
                    totdcto = nprecos * Val(rstempdet.Fields("F3PORDCT") & "") / 100
                    ValVta = Val(rstempdet.Fields("F3CANPRO") & "") * nprecos - totdcto
                    IGV = ValVta * (wgigv / 100)
                    preciounit = nprecos + (nprecos * (wgigv / 100))
                    TOTAL = ValVta + IGV
                Else
                    totdcto = nprecos * Val(rstempdet.Fields("F3PORDCT") & "") / 100
                    ValVta = Val(rstempdet.Fields("F3CANPRO") & "") * nprecos - totdcto
                    IGV = 0#
                    preciounit = nprecos
                    TOTAL = ValVta + IGV
                End If
                csql = "UPDATE " & cnomtabla & " SET F3PRECOS=" & nprecos & ",F3TOTDCT=" & totdcto & _
                       ",F5VALVTA=" & ValVta & ",F3IGV=" & IGV & ",F3PREUNI=" & preciounit & ",F3TOTAL=" & TOTAL & _
                       " WHERE ITEM=" & nitem & " AND F3CODPRO='" & ccodprod & "'"
                cnn_form.Execute (csql)
            End If
            rsprodprov.Close
            rstempdet.MoveNext
        Loop
    End If
    rstempdet.Close
    dxDBGrid1.Dataset.Refresh

    calcula

End Sub

Private Sub TRASLADA_CTASXPAGAR(pnumero As String)
Dim ncorre_d            As Double
Dim amovs_cab(0 To 18)  As a_grabacion
Dim rsif4orden            As New ADODB.Recordset
Dim rsbf5pla            As New ADODB.Recordset
Dim rsproveedor         As New ADODB.Recordset
Dim ntotal              As Double
Dim ntc                 As Double
Dim cdetal              As String
Dim dfechamov           As Date
Dim ccodprov            As String
Dim cnomprov            As String
Dim cruc                As String
Dim cnro_comp           As String
Dim Moneda              As String
Dim csql                As String
Dim rspag_dcto          As New ADODB.Recordset

    If cnn_ctaspag.State = adStateOpen Then cnn_ctaspag.Close
    cconex_ctaspag = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\CTASPAG.MDB" & ";Persist Security Info=False"
    cnn_ctaspag.Open cconex_ctaspag
    
    ntotal = 0#: ntc = 0#: cdetal = ""
    If rsif4orden.State = adStateOpen Then rsif4orden.Close
    rsif4orden.Open "SELECT F4MONTO,F4FECEMI,F4TIPCAM,F4OBSERVA FROM IF4ORDEN WHERE F4NUMORD=" & pnumero & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsif4orden.EOF Then
        ntotal = Val("" & rsif4orden.Fields("F4MONTO"))
        ntc = Val("" & rsif4orden.Fields("F4TIPCAM"))
        cdetal = Left(Trim("" & rsif4orden.Fields("F4OBSERVA")), 100)
        dfechamov = Format(rsif4orden.Fields("F4FECEMI"), "DD/MM/YYYY")
    End If
    rsif4orden.Close
    
    cruc = Txt_Prove.Text
    If rsproveedor.State = adStateOpen Then rsproveedor.Close
    csql = "SELECT F2NOMPROV,F2CODPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & cruc & "'"
    rsproveedor.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsproveedor.EOF Then
        cnomprov = rsproveedor.Fields("F2NOMPROV") & ""
        ccodprov = rsproveedor.Fields("F2CODPROV") & ""
    End If
    rsproveedor.Close
    
    If rspag_dcto.State = adStateOpen Then rspag_dcto.Close
    rspag_dcto.Open "SELECT CORRELA FROM PAG_DCTO ORDER BY CORRELA DESC", cnn_ctaspag, adOpenDynamic, adLockOptimistic
    If Not rspag_dcto.EOF Then
        ncorre_d = rspag_dcto.Fields("CORRELA") + 1
    Else
        ncorre_d = 1
    End If
    rspag_dcto.Close
    
    cnro_comp = "O/c" & Format(pnumero, "0000000")
    Moneda = IIf(Cmbmone.ListIndex = 0, "S", "D")
    
    amovs_cab(0).campo = "VIA_INGR": amovs_cab(0).valor = "1": amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "CORRELA": amovs_cab(1).valor = ncorre_d: amovs_cab(1).TIPO = "N"
    amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = cnro_comp: amovs_cab(2).TIPO = "T"
    amovs_cab(3).campo = "FCH_COMP": amovs_cab(3).valor = dfechamov: amovs_cab(3).TIPO = "F"
    amovs_cab(4).campo = "PROVEEDORO": amovs_cab(4).valor = ccodprov: amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "RUC": amovs_cab(5).valor = cruc: amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "MONEDAO": amovs_cab(6).valor = Moneda: amovs_cab(6).TIPO = "T"
    amovs_cab(7).campo = "TOTALO": amovs_cab(7).valor = ntotal: amovs_cab(7).TIPO = "N"
    amovs_cab(8).campo = "TCAMBIOO": amovs_cab(8).valor = ntc: amovs_cab(8).TIPO = "N"
    amovs_cab(9).campo = "PROVEEDOR": amovs_cab(9).valor = ccodprov: amovs_cab(9).TIPO = "T"
    amovs_cab(10).campo = "MONEDA": amovs_cab(10).valor = Moneda: amovs_cab(10).TIPO = "T"
    amovs_cab(11).campo = "TCAMBIO": amovs_cab(11).valor = ntc: amovs_cab(11).TIPO = "N"
    amovs_cab(12).campo = "TOTAL": amovs_cab(12).valor = ntotal: amovs_cab(12).TIPO = "N"
    amovs_cab(13).campo = "SALDO": amovs_cab(13).valor = ntotal: amovs_cab(13).TIPO = "N"
    amovs_cab(14).campo = "DEB_HAB": amovs_cab(14).valor = "H": amovs_cab(14).TIPO = "T"
    amovs_cab(15).campo = "REFERENCIA": amovs_cab(15).valor = cdetal: amovs_cab(15).TIPO = "T"
    amovs_cab(16).campo = "NOMPROV": amovs_cab(16).valor = cnomprov: amovs_cab(16).TIPO = "T"
    amovs_cab(17).campo = "CONCEPTO": amovs_cab(17).valor = cdetal: amovs_cab(17).TIPO = "T"
    amovs_cab(18).campo = "FCH_VCTO": amovs_cab(18).valor = dfechamov: amovs_cab(18).TIPO = "F"
    
    GRABA_REGISTRO amovs_cab(), "PAG_DCTO", "A", 18, cnn_ctaspag, ""
    
    cnn_ctaspag.Close
    
    cnn_dbbancos.Execute ("UPDATE IF4ORDEN SET F4CORRELA=" & ncorre_d & " WHERE F4NUMORD=" & pnumero & "")
    
End Sub

Public Function NumXProv(cod)
    
    SQL = "select count(*) from if4orden where f4codprv='" & cod & "'"
    If rst.State = adStateOpen Then rst.Close
    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        xnum = rst.Fields(0).Value + 1
    Else
        xnum = 1
    End If
    rst.Close
    NumXProv = xnum

End Function
