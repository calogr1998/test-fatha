VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ordencompra_imp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Compra Internacional"
   ClientHeight    =   9510
   ClientLeft      =   1185
   ClientTop       =   3450
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "ordencompra_imp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtResumen 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2400
      MaxLength       =   250
      TabIndex        =   76
      Top             =   4380
      Width           =   9435
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      TabIndex        =   34
      Top             =   540
      Width           =   11715
      Begin VB.TextBox Txt_Referencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox Txt_Prove 
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
         Left            =   1710
         TabIndex        =   39
         Top             =   630
         Width           =   1125
      End
      Begin VB.TextBox Txt_NumOC 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   3180
      End
      Begin VB.TextBox Txt_NumSolComp 
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
         Height          =   315
         Left            =   8325
         MaxLength       =   4
         TabIndex        =   36
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox txtusuario 
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
         Height          =   315
         Left            =   7320
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtcontacto 
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
         Left            =   7560
         TabIndex        =   40
         Top             =   960
         Width           =   4000
      End
      Begin Threed.SSPanel pnldireprv 
         Height          =   270
         Left            =   1710
         TabIndex        =   41
         Top             =   990
         Width           =   4950
         _Version        =   65536
         _ExtentX        =   8731
         _ExtentY        =   476
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Left            =   2910
         TabIndex        =   53
         Top             =   630
         Width           =   8620
         _Version        =   65536
         _ExtentX        =   15205
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Left            =   10080
         TabIndex        =   37
         Top             =   240
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
         Text            =   "21/01/2013"
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
         ButtonPicture   =   "ordencompra_imp.frx":000C
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
         Caption         =   "Nº Orden Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   83
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   61
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Requerimiento"
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
         Height          =   210
         Index           =   10
         Left            =   6840
         TabIndex        =   57
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
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
         Height          =   210
         Index           =   7
         Left            =   6720
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Height          =   210
         Index           =   9
         Left            =   9435
         TabIndex        =   55
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
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
         Height          =   210
         Index           =   6
         Left            =   6840
         TabIndex        =   54
         Top             =   1020
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   120
      TabIndex        =   16
      Top             =   1860
      Width           =   11715
      Begin VB.TextBox Text1 
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
         Left            =   1305
         TabIndex        =   85
         Top             =   1320
         Width           =   4425
      End
      Begin VB.TextBox TxtPI 
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
         Left            =   4380
         TabIndex        =   82
         Top             =   2040
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox TxtRequest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   80
         Top             =   2100
         Width           =   1635
      End
      Begin VB.ComboBox CboTipo 
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
         ItemData        =   "ordencompra_imp.frx":035E
         Left            =   1350
         List            =   "ordencompra_imp.frx":036B
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2040
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.TextBox TxtCli 
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
         Left            =   1350
         TabIndex        =   47
         Top             =   945
         Width           =   1020
      End
      Begin VB.ComboBox cmbseguro 
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
         ItemData        =   "ordencompra_imp.frx":039F
         Left            =   1350
         List            =   "ordencompra_imp.frx":03A9
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1680
         Width           =   4380
      End
      Begin VB.ComboBox cmbvia 
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
         ItemData        =   "ordencompra_imp.frx":03DE
         Left            =   7440
         List            =   "ordencompra_imp.frx":03EB
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1350
         Width           =   1680
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "ordencompra_imp.frx":040B
         Left            =   7440
         List            =   "ordencompra_imp.frx":041E
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1740
         Width           =   1680
      End
      Begin VB.TextBox txtuupp 
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
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   19
         Top             =   3120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtcodcosto 
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
         Left            =   780
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox cmbtipopera 
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
         ItemData        =   "ordencompra_imp.frx":043B
         Left            =   2220
         List            =   "ordencompra_imp.frx":0445
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3540
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ComboBox Cmbmone 
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
         ItemData        =   "ordencompra_imp.frx":045F
         Left            =   6750
         List            =   "ordencompra_imp.frx":046C
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox txtcodsoli 
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
         Left            =   1350
         TabIndex        =   42
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txtcodforma 
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
         Left            =   1350
         TabIndex        =   45
         Top             =   585
         Width           =   1020
      End
      Begin VB.TextBox txt_tc 
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
         Height          =   315
         Left            =   10410
         TabIndex        =   44
         Text            =   "3.200"
         Top             =   225
         Width           =   1140
      End
      Begin VB.TextBox txtplazo_entrega 
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
         Left            =   7440
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   600
         Width           =   4140
      End
      Begin Threed.SSPanel pnlnomsoli 
         Height          =   330
         Left            =   2385
         TabIndex        =   20
         Top             =   240
         Width           =   3350
         _Version        =   65536
         _ExtentX        =   5909
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Height          =   330
         Left            =   2385
         TabIndex        =   21
         Top             =   600
         Width           =   3350
         _Version        =   65536
         _ExtentX        =   5909
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483630
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
      Begin aBoxCtl.aBox abofechaentrega 
         Height          =   315
         Left            =   7440
         TabIndex        =   48
         Top             =   945
         Width           =   1650
         _ExtentX        =   2910
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
         Text            =   "21/01/2013"
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
         ButtonPicture   =   "ordencompra_imp.frx":0487
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
      Begin Threed.SSPanel txtdesuupp 
         Height          =   330
         Left            =   3435
         TabIndex        =   23
         Top             =   3000
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Height          =   330
         Left            =   915
         TabIndex        =   22
         Top             =   3570
         Visible         =   0   'False
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483630
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
      Begin Threed.SSPanel LblNomCli 
         Height          =   330
         Left            =   2385
         TabIndex        =   73
         Top             =   960
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5909
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483630
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cotizacion"
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
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   1380
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label LblPI 
         AutoSize        =   -1  'True
         Caption         =   "PI"
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
         Height          =   195
         Left            =   4140
         TabIndex        =   81
         Top             =   2100
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Request"
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
         Height          =   195
         Index           =   14
         Left            =   5940
         TabIndex        =   79
         Top             =   2100
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   77
         Top             =   2040
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarcador"
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
         Height          =   195
         Left            =   135
         TabIndex        =   74
         Top             =   1005
         Width           =   855
      End
      Begin VB.Image imgbote 
         Height          =   900
         Left            =   9870
         Picture         =   "ordencompra_imp.frx":07D9
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgavion 
         Height          =   900
         Left            =   9870
         Picture         =   "ordencompra_imp.frx":101B
         Stretch         =   -1  'True
         Top             =   1065
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Imgcarro 
         Height          =   900
         Left            =   9870
         Picture         =   "ordencompra_imp.frx":1325
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "Via de Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5940
         TabIndex        =   66
         Top             =   1395
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Terminos de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5940
         TabIndex        =   65
         Top             =   1785
         Width           =   1575
      End
      Begin VB.Label lbluupp 
         AutoSize        =   -1  'True
         Caption         =   "UUPP"
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
         Height          =   210
         Left            =   3900
         TabIndex        =   33
         Top             =   3180
         Width           =   390
      End
      Begin VB.Label lblccosto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Height          =   210
         Left            =   3720
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seguro"
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
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   31
         Top             =   1740
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Referencia Interna"
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
         Height          =   375
         Index           =   4
         Left            =   150
         TabIndex        =   30
         Top             =   1260
         Visible         =   0   'False
         Width           =   1140
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
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
         Height          =   210
         Index           =   17
         Left            =   8985
         TabIndex        =   29
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda "
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
         Height          =   210
         Index           =   19
         Left            =   5940
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
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
         Height          =   210
         Left            =   135
         TabIndex        =   27
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plazo de Entrega"
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
         Height          =   210
         Index           =   1
         Left            =   5940
         TabIndex        =   25
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Entrega"
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
         Height          =   210
         Index           =   8
         Left            =   5940
         TabIndex        =   24
         Top             =   990
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   6300
      TabIndex        =   4
      Top             =   7740
      Width           =   5500
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3735
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   315
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3735
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   675
         Width           =   1575
      End
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1035
         Width           =   1575
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   480
         Width           =   510
      End
      Begin VB.TextBox txtmonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   495
         Width           =   480
      End
      Begin VB.TextBox txtbase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FLETE :"
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
         Left            =   2280
         TabIndex        =   72
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DCTO :"
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
         Left            =   2325
         TabIndex        =   70
         Top             =   690
         Width           =   645
      End
      Begin VB.Label lbltermino 
         Caption         =   "EXW"
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
         Left            =   2520
         TabIndex        =   68
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblmoneda 
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
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         Top             =   510
         Width           =   360
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
         Left            =   2400
         TabIndex        =   14
         Top             =   480
         Width           =   360
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
         Left            =   1125
         TabIndex        =   13
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label12 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2040
         TabIndex        =   12
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label10 
         Caption         =   "Monto Inaf."
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
         Height          =   210
         Left            =   1620
         TabIndex        =   10
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "B. Imponible"
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
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   120
      TabIndex        =   0
      Top             =   7740
      Width           =   6180
      Begin VB.TextBox Text2 
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
         Left            =   1425
         TabIndex        =   58
         Top             =   240
         Width           =   2800
      End
      Begin VB.TextBox txtobserva 
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
         Left            =   1425
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtempresa 
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
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtlugar_entrega 
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
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   67
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Peso Aprox."
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "ordencompra_imp.frx":1BEF
      TabIndex        =   52
      Top             =   4740
      Width           =   11715
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   11910
      _Version        =   65536
      _ExtentX        =   21008
      _ExtentY        =   741
      _StockProps     =   15
      ForeColor       =   -2147483630
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
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   9
         Tools           =   "ordencompra_imp.frx":9C0D
         ToolBars        =   "ordencompra_imp.frx":10DF3
      End
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
      Left            =   4200
      TabIndex        =   62
      Top             =   7200
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto Resumen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   11
      Left            =   120
      TabIndex        =   75
      Top             =   4380
      Width           =   2265
   End
   Begin VB.Label lbldescripcion 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7335
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   4380
   End
End
Attribute VB_Name = "ordencompra_imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsOrdenCab              As ADODB.Recordset
Dim rsOrdenDet              As New ADODB.Recordset
Dim rssolcab                As ADODB.Recordset
Dim rsSolDet                As ADODB.Recordset
Dim rst                     As ADODB.Recordset
Dim rstaux                  As ADODB.Recordset
Dim rsproductos             As ADODB.Recordset
Dim SWcondipago             As Integer
Dim Wnuevo                  As Boolean
Dim flawigv                 As Boolean
Dim seleccion               As Boolean
Dim CadSql                  As String
Dim cnn_form                As New ADODB.Connection
Dim cconex_form             As String
Dim sw_nuevo_item           As Boolean
Dim ExisteOrdenCompra       As Boolean
Dim wIgv                    As Single
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
Dim wNO As String
Dim fecha                   As String
Dim existe As Boolean
Dim RSCONSULTA As New ADODB.Recordset
Dim cDia As String, cmes As String, cAño As String

Private Sub Imprime_Orden(opcion)
Dim sql As String

'Set RSCONSULTA = New ADODB.Recordset
Dim RsPago As New ADODB.Recordset
Set RsPago = New ADODB.Recordset
Dim RsCTR_COM As New ADODB.Recordset
Set RsCTR_COM = New ADODB.Recordset
Dim REPORTE As New ActiveReport


If opcion = 1 Or opcion = 0 Then
    If opcion = 1 Then
            With Acr_OrdenCImp
            If Cmbmone.ListIndex = 0 Then
            .lblmoneda2.Caption = "S/."
            ElseIf Cmbmone.ListIndex = 1 Then

            .lblmoneda2.Caption = "US$"
            Else
            .lblmoneda2.Caption = " € "
            
            End If

            GOC = Txt_NumOC.Text
            
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        sql = "SELECT A.F4NUMORD,D.F2CODCLI,A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
              "A.F4FECEMI,A.F4FECVEN,A.F4REFERE, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA,A.F4CODPRV, " & _
              "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4PESO,A.F4MARCAS,A.F4VIATRANS,A.F4TERCOMPRA,A.F4SEGURO,A.F4FECENT,A.F4TIPMON,A.F4REFERE,C.F2NOMCLI " & _
              "FROM IF4ORDEN AS A, EF2PROVEEDORES AS B,CENTROS AS CE, EF2CLIENTES C,IF3ORDEN AS D WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
              "' AND A.F4NUMORD=D.F4NUMORD AND A.F4ESTNUL<>'S' AND D.F3CENCOS=CE.F3ABREV AND CE.F3CODCLI=C.F2CODCLI ORDER BY D.F2CODCLI;"
            
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
            .F4NUMORD.Text = Format("" & RSCONSULTA.Fields("F4NUMORD"), "0000000") & "/" & Right(Year(RSCONSULTA.Fields("F4FECEMI")), 2)
            .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
            .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
            .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
            .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
            .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
            .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
            .F4REFERE.Text = "" & RSCONSULTA.Fields("F4REFERE")
            .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
            .F4FECENT.Text = "" & RSCONSULTA.Fields("F4FECENT")
            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
            If RSCONSULTA.Fields("F4TIPMON") = "S" Then
                .F4TIPMON.Text = "SOLES"
            Else
                If RSCONSULTA.Fields("F4TIPMON") = "D" Then
                .F4TIPMON.Text = "DOLARES"
                End If
            End If
            .F4LUGAR_ENTREGA.Text = "" & RSCONSULTA.Fields("F4LUGAR_ENTREGA")
            
            Select Case RSCONSULTA.Fields("F4VIATRANS") & ""
            Case "0":
            .fldviatrans.Text = "Maritimo"
            Case "1":
            .fldviatrans.Text = "Aéreo"
            Case "2":
            .fldviatrans.Text = "Terrestre"
            End Select
            Select Case RSCONSULTA.Fields("F4SEGURO") & ""
            Case "0":
            .fldseguro.Text = "Mercaderia Asegurada"
            Case "1":
            .fldseguro.Text = "Seguro por nuestra cuenta"
            End Select
            Select Case RSCONSULTA.Fields("F4tercompra") & ""
            Case "0":
            .fldtercompra.Text = "EXW"
            .lbltermino.Caption = "EXW"
            Case "1":
            .fldtercompra.Text = "FOB"
            .lbltermino.Caption = "FOB"
            Case "2":
            .fldtercompra.Text = "CFR"
            .lbltermino.Caption = "CFR"
            Case "3":
            .fldtercompra.Text = "CIF"
            .lbltermino.Caption = "CIF"
            End Select

            
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
            .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")

            End If
            RsPago.Close
            .F4EMITIR.Text = "" & wnomcia
            .f4emitir2.Text = "" & wdireccion & " - Perú"
            .f4emitir3.Text = "Ph: " & wtelefono & "  Fax: " & wfax
            .F4OBSFECHA.Text = "" & RSCONSULTA.Fields("F4PESO") & ""
            .REMITIR.Text = RSCONSULTA.Fields("F4MARCAS") & ""
            .F4OBSGEN.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            
            End If
            
        'AGRUPACION POR CLIENTE
        '********************************************
        sql = "SELECT CE.F3CODCLI,D.ITEM, D.F5NOMPRO, D.F3CANPRO, D.F3PRECOS, D.F3TOTAL FROM IF3ORDEN AS D, CENTROS AS CE"
        sql = sql & " " & "Where D.F4LOCAL = '0'  AND CE.F3ABREV=D.F3CENCOS And D.F4NUMORD = '" & GOC & "' ORDER BY CE.F3CODCLI"


       If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, 3, 1

        Set .DataControl1.Recordset = Rs
        .GroupHeader1.DataField = "f3codcli"
        .DataControl1.ConnectionString = cnn_dbbancos
        .DataControl1.Source = sql
        .Caption = "Orden de Compra Internacional"
        RSCONSULTA.Close
        .Show 1
        
        '*********************************************
        End With
    Else
        With Acr_OrdenCImp_ingles
        If Cmbmone.ListIndex = 0 Then
            .lblmoneda1.Caption = "S/."
            .lblmoneda2.Caption = "S/."
            .lblmoneda3.Caption = "S/."
        ElseIf Cmbmone.ListIndex = 1 Then
            .lblmoneda1.Caption = "US$"
            .lblmoneda2.Caption = "US$"
            .lblmoneda3.Caption = "US$"
        Else
            .lblmoneda1.Caption = " € "
            .lblmoneda2.Caption = " € "
            .lblmoneda3.Caption = " € "

        End If

        GOC = Txt_NumOC.Text
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        sql = "SELECT A.F4NUMORD,D.F2CODCLI,A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
              "A.F4FECEMI,A.F4FECVEN,A.F4REFERE, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA,A.F4CODPRV, " & _
              "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4PESO,A.F4MARCAS,A.F4VIATRANS,A.F4TERCOMPRA,A.F4SEGURO,A.F4FECENT,A.F4TIPMON,A.F4REFERE,C.F2NOMCLI " & _
              "FROM IF4ORDEN AS A, EF2PROVEEDORES AS B,CENTROS AS CE, EF2CLIENTES C,IF3ORDEN AS D WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
              "' AND A.F4NUMORD=D.F4NUMORD AND A.F4ESTNUL<>'S' AND D.F3CENCOS=CE.F3ABREV AND CE.F3CODCLI=C.F2CODCLI ORDER BY D.F2CODCLI;"
    
        RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RSCONSULTA.EOF Then
            .LblTitle.Caption = UCase(wnomcia)
            wNO = RSCONSULTA.Fields("F4NUMORD") & ""
            .F4NUMORD.Text = "" & RSCONSULTA.Fields("F4NUMORD")
            .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
            .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
            
            .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
            .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
            .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
            .F4REFERE.Text = "" & RSCONSULTA.Fields("F4REFERE")
 
        
            If RSCONSULTA.Fields("F4TIPMON") = "S" Then
                .F4TIPMON.Text = "SOLES"
            Else
                If RSCONSULTA.Fields("F4TIPMON") = "D" Then
                    .F4TIPMON.Text = "USD"
                End If
            End If

            Select Case Day(RSCONSULTA.Fields("F4FECEMI") & "")
            Case 1, 31
                cDia = Day(RSCONSULTA.Fields("F4FECEMI") & "") & "st"
            Case 2
                cDia = Day(RSCONSULTA.Fields("F4FECEMI") & "") & "nd"
            Case 3
                cDia = Day(RSCONSULTA.Fields("F4FECEMI") & "") & "rd"
            Case 4 To 30
                cDia = Day(RSCONSULTA.Fields("F4FECEMI") & "") & "th"
            End Select
            cmes = dev_mes_ingles(Month("" & RSCONSULTA.Fields("F4FECEMI")))
            cAño = (Year("" & RSCONSULTA.Fields("F4FECEMI")))
            .F4FECEMI.Text = cmes & " " & cDia & " " & cAño
            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
            Select Case RSCONSULTA.Fields("F4VIATRANS") & ""
            Case "0":
                .fldviatrans.Text = "Sea"
            Case "1":
                .fldviatrans.Text = "Air"
            Case "2":
                .fldviatrans.Text = "Ground"
            End Select
            Select Case RSCONSULTA.Fields("F4SEGURO") & ""
            Case "0":
            '    .fldseguro.Text = "Mercaderia Asegurada"
            Case "1":
             '   .fldseguro.Text = "Seguro por nuestra cuenta"
            End Select
            Select Case RSCONSULTA.Fields("F4tercompra") & ""
            Case "0":
                '.fldtercompra.Text = "EXW"
                '.lbltermino.Caption = "EXW"
                .LblUnit.Caption = "Unit Price EXW"
                .LblTot.Caption = "Total EXW"
                .LblTotal.Caption = "Total EXW"
            Case "1":
                '.fldtercompraText = "FOB"
                '.lbltermino.Caption = "FOB"
                .LblUnit.Caption = "Unit Price FOB"
                .LblTot.Caption = "Total FOB"
                .LblTotal.Caption = "Total FOB"
            Case "2":
                '.fldtercompra.Text = "CFR"
                '.lbltermino.Caption = "CFR"
                .LblUnit.Caption = "Unit Price CFR"
                .LblTot.Caption = "Total CFR"
                .LblTotal.Caption = "Total CFR"
            Case "3":
                '.fldtercompra.Text = "CIF"
                '.lbltermino.Caption = "CIF"
                .LblUnit.Caption = "Unit Price CIF"
                .LblTot.Caption = "Total CIF"
                .LblTotal.Caption = "Total CIF"
            End Select
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAGeng FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
                .F2DESPAG.Text = Space(15) & RsPago.Fields("F2DESPAGeng")
            End If
            RsPago.Close
            .F4EMITIR.Text = UCase("" & wnomcia)
            .f4emitir2.Text = "" & wdireccion
            .f4emitir3.Text = "" & wDistrito & " - Perú"
            .f4emitir4.Text = "Ph: " & wtelefono & "  Fax: " & wfax

            
            .LblShip0.Text = UCase("" & wnomcia)
            wrucprov = traerCampo("EF2PROVEEDORES", "f2newruc", "f2codprov", RSCONSULTA.Fields("f2codcli") & "", "")
            wnomprov = traerCampo("EF2embarcadorES", "f2nomemb", "f2newruc", wrucprov & "", "")
            
                .LblShip1.Text = wnomprov
                .LblShip2.Text = traerCampo("EF2embarcadorES", "f2diremb", "f2newruc", wrucprov & "", "")
                If Len(Trim(traerCampo("EF2embarcadorES", "f2telemb", "f2newruc", wrucprov & "", ""))) > 0 Then
                    .LblShip3.Text = "Ph: " & traerCampo("EF2embarcadorES", "f2telemb", "f2newruc", wrucprov & "", "")
                Else
                    .LblShip3.Text = ""
                    .LblShip3.Visible = False
                End If
                If Len(Trim(traerCampo("EF2embarcadorES", "f2FAXemb", "f2newruc", wrucprov & "", ""))) > 0 Then
                    .LblShip3.Visible = True
                    If Len(Trim(.LblShip3.Text)) > 0 Then .LblShip3.Text = .LblShip3.Text & Space(3)
                    .LblShip3.Text = .LblShip3.Text & "Fax: " & traerCampo("EF2embarcadorES", "f2FAXemb", "f2newruc", wrucprov & "", "")
                Else
                    '.LblShip3.Visible = False
                End If
                If Len(Trim(traerCampo("EF2embarcadorES", "f2nomcont", "f2newruc", wrucprov & "", ""))) > 0 Then
                    .LblShip34.Text = "Contact: " & traerCampo("EF2embarcadorES", "f2nomcont", "f2newruc", wrucprov & "", "")
                Else
                    .LblShip34.Visible = False
                End If
                If Len(Trim(traerCampo("EF2embarcadorES", "f2mailcont", "f2newruc", wrucprov & "", ""))) > 0 Then
                    .LblShip5.Text = "E-mail: " & traerCampo("EF2embarcadorES", "f2mailcont", "f2newruc", wrucprov & "", "")
                Else
                    .LblShip5.Visible = False
                End If
                If Len(Trim(.LblShip5.Text)) = 0 Then
                If Len(Trim(traerCampo("EF2embarcadorES", "f2CUENTA", "f2newruc", wrucprov & "", ""))) > 0 Then
                    wnomprov = traerCampo("EF2proveedores", "f2nomprov", "f2newruc", wrucprov & "", "")
                    .LblShip5.Text = UCase(Trim(Left(wnomprov, InStr(1, wnomprov, " ")))) & " Account# " & Trim(traerCampo("EF2embarcadorES", "f2CUENTA", "f2newruc", wrucprov & "", ""))
                    .LblShip5.Visible = True
                Else
                    .LblShip5.Visible = False
                End If
                End If
            .LblMarks0.Text = UCase("" & wnomcia)
            .LblMarks1.Text = "P.O. " & Mid(wNO, 7, 4) & "/" & Left(wNO, 4)
            If IsDate(RSCONSULTA.Fields("f4fecent")) Then
                Select Case Day(RSCONSULTA.Fields("f4fecent") & "")
                Case 1, 31
                    cDia = Day(RSCONSULTA.Fields("f4fecent") & "") & "st"
                Case 2
                    cDia = Day(RSCONSULTA.Fields("f4fecent") & "") & "nd"
                Case 3
                    cDia = Day(RSCONSULTA.Fields("f4fecent") & "") & "rd"
                Case 4 To 30
                    cDia = Day(RSCONSULTA.Fields("f4fecent") & "") & "th"
                End Select
                cmes = dev_mes_ingles(Month("" & RSCONSULTA.Fields("f4fecent")))
                cAño = (Year("" & RSCONSULTA.Fields("f4fecent")))
                .LblRequired.Text = cmes & " " & cDia & " " & cAño
            Else
                .LblRequired.Text = "ASAP"
            End If
            .LblReference0.Caption = UCase(Trim(Left(wnomcia, InStr(1, wnomcia, " ")))) & "'s Reference:"
            .LblReference1.Text = RSCONSULTA.Fields("F4REFERE")
            .F4OBSGEN.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            If Len(Trim(RSCONSULTA.Fields("F4MARCAS") & "")) > 0 Then
                .LblFactory0.Caption = UCase(Trim(Left(.F2NOMPROV.Text, InStr(1, .F2NOMPROV.Text, " ")))) & "'s Reference:"
                .LblFactory1.Text = "" & RSCONSULTA.Fields("F4MARCAS")
            Else
                .LblFactory0.Visible = False
                .LblFactory1.Visible = False
            End If
         
            
        End If
        
        
        

        '*********************************************************
        'AGRUPACION POR CLIENTE
        sql = "SELECT CE.F3CODCLI,D.ITEM, D.F5NOMPRO, D.F3CANPRO, D.F3PRECOS, D.F3TOTAL FROM IF3ORDEN AS D, CENTROS AS CE"
        sql = sql & " " & "Where D.F4LOCAL = '0'  AND CE.F3ABREV=D.F3CENCOS And D.F4NUMORD = '" & GOC & "' ORDER BY CE.F3CODCLI"


       If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, 3, 1



        Set .DataControl1.Recordset = Rs
        .GroupHeader1.DataField = "f3codcli"
        .DataControl1.ConnectionString = cnn_dbbancos
        .DataControl1.Source = sql
        .Caption = "Orden de Compra Internacional"
        RSCONSULTA.Close
        .Show 1
        
        '**********************************************
       
    End With
    End If
    


ElseIf opcion = 2 Then
With Acr_OrdenC_Otros
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
    "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA, " & _
    "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,F4OBSERVA " & _
    " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & _
    " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"

    RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
        .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
        .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").Value
        .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
        .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
        .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
        .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
        .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
        .F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
        .F4NOTA.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
        
        If RsPago.State = adStateOpen Then RsPago.Close
        RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsPago.EOF Then
            .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
        End If
        RsPago.Close
'
        If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
        RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
        If Not RsCTR_COM.EOF Then
            .F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
            
            .F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
            '.F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
        End If
        RsCTR_COM.Close
        .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""

    End If
    RSCONSULTA.Close

    .DataControl1.ConnectionString = cnn_form
    .DataControl1.Source = "select * from tmpOrdendeCompra"

    .Show vbModal
End With
End If
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
    
    sTo = wemailoc
    sCC = wemailccoc
    sBCC = ""
    sSubject = wasuntooc
    sBody = wtextooc
    
    ret = Shell("Start.exe " _
        & "mailto:" & """" & sTo & """" _
        & "?Subject=" & """" & sSubject & """" _
        & "&cc=" & """" & sCC & """" _
        & "&bcc=" & """" & sBCC & """" _
        & "&Body=" & """" & sBody & """" _
        & "&File=" & """" & "c:\autoexec.bat" & """" _
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
        cantidad = Val(Format(.Columns.ColumnByFieldName("F3CANPRO").Value, "0.00"))
        If cantidad > 0 Then
            If cmbtipopera.ListIndex = 0 Then
                If .Columns.ColumnByFieldName("F5AFECTO").Value = "*" Then     'Afecto
                    totdcto = (Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00"))) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                    IGV = ValVta * (wwigv / 100)
                    .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "#,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.000")) + (Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.000")) * (wwigv / 100))
                    .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
                Else  'Inafecto
                    IGV = 0
                    .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "0.00")
                    totdcto = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.000")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00")) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "0.000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.000"))
                    .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
                End If
            Else
                costo = Val(.Columns.ColumnByFieldName("F3PRECOS").Value)
                TOTAL = cantidad * costo                '
                .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
            End If
        End If
    End With
    
End Sub

Sub MostrarDatos()
Dim sw_nuevo_temp   As Boolean
Dim xnombre         As String
Dim I               As Integer
Dim entrega         As Date
    
    If rssolcab.State = adStateOpen Then rssolcab.Close
    With rssolcab
        rssolcab.Open "select * from tb_cabsolicitud where cod_solicitud='" & num_solcomp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not .EOF And Not .Bof Then
            'Txt_Prove.Text="" !
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV from EF2PROVEEDORES where F2newruc='" & !cs_proveedor & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                Txt_Prove.Text = "" & rst!F2NEWRUC
                PnlNomPrv.Caption = rst!F2NOMPROV
                pnldireprv.Caption = IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
            Else
                PnlNomPrv.Caption = "Ruc es menor a 11 digitos"
                pnldireprv.Caption = "No tiene "
            End If
            rst.Close

            Txt_NumSolComp = !cod_solicitud & ""
            txt_fecha.Value = !cs_fecha & ""
            TxtCodCosto = !cs_codcosto & ""
            PnlNomCosto = !cs_descosto & ""
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
            ElseIf !cs_moneda = "D" Then
                Cmbmone.ListIndex = 1
            Else
                Cmbmone.ListIndex = 2
            End If
            txtuupp.Text = .Fields("UUPP") & ""
            If VALIDA_UUPP(txtuupp.Text) = True Then
                txtdesuupp.Caption = wdeslocalidad
            End If
            txtlugar_entrega.Text = Left(Trim("" & !cs_LugEntr), 100)
            txt_tc.Text = Format(Val(.Fields("F4TIPCAM") & ""), "2.810")
        End If
        rssolcab.Close
    End With
     
    '*** detalle de solicitud de compra
    'Versión Nueva
    With dxDBGrid1
        If rsSolDet.State = adStateOpen Then rsSolDet.Close
        rsSolDet.Open "select * from tb_DETsolicitud where cod_solicitud='" & _
        num_solcomp & "' and candis>0 ORDER BY ITEM", cnn_dbbancos, adOpenDynamic, adLockOptimistic
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
            Do While Not (rsSolDet.EOF)
                If rsSolDet!cod_solicitud = Trim(Txt_NumSolComp.Text) Then
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
                    I = I + 1
                    .Dataset.FieldValues("item") = I
                    .Dataset.FieldValues("f3canpro") = rsSolDet!candis
                    .Dataset.FieldValues("f3codpro") = rsSolDet!COD_PRODUCTO & ""
                    .Dataset.FieldValues("f5nompro") = rsSolDet!ds_descripcion & ""
                    .Dataset.FieldValues("f3medida") = rsSolDet!ds_unidmed & ""
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
                    .Dataset.FieldValues("f5codfab") = rsSolDet!f5codfab & ""
                End If
                rsSolDet.MoveNext
                Calcula_PvtaTot
            Loop
            .Dataset.Post
            .Dataset.EnableControls
            .Dataset.Open
            '.OptionEnabled = True
            sw_nuevo_item = False
        End If
        rsSolDet.Close
        Call calcula
    End With
    
End Sub

Private Sub abofechaentrega_GotFocus()

    abofechaentrega.FocusSelect = True
    
End Sub

Private Sub abofechaentrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        dxDBGrid1.SetFocus
'        dxDBGrid1.Columns.FocusedIndex = 1
        SendKeys "{tab}"
    End If

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
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                dxDBGrid1.Dataset.Post
                sw_detalle = True
            End If
            If MsgBox("¿Desea Grabar la Orden de Compra?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
                Me.MousePointer = vbHourglass
                'If TxtResumen.Text = "" Then
                 '   MsgBox "Debe ingresar el concepto"
                  '  Exit Sub
                'Else
                GrabarOC
                ActualizarNumOrd
                Me.MousePointer = vbDefault
                'End If
            End If
    
        Case "ID_Español":
           Me.MousePointer = vbHourglass
            If Len(Trim(Txt_NumOC.Text)) > 0 Then
                Imprime_Orden 1
            Else
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
            End If
            Me.MousePointer = vbDefault
        Case "ID_Inglés":
           Me.MousePointer = vbHourglass
            If Len(Trim(Txt_NumOC.Text)) > 0 Then
                Imprime_Orden 0
            Else
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
            End If
            Me.MousePointer = vbDefault
        Case "idanular":
            If Trim$(Txt_NumOC.Text) = "" Then
                MsgBox "No existe Orden de Compra", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                eliminar
            End If
        Case "idemail":
            EMAIL
            
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
Private Sub ActualizarNumOrd()
Dim prSol As New Recordset
Dim cadena As String
       sql = "Select*from TB_CABSOLICITUD where Cod_Solicitud='" & Trim(Txt_NumSolComp) & "'"
        prSol.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
        If Not (prSol.EOF Or prSol.Bof) Then
            If Trim(prSol!numorden) <> "" Then
                cadena = "" & prSol!numorden & " , " & Trim(Txt_NumOC.Text)
            Else
                cadena = Trim(Txt_NumOC.Text)
            End If
            sql = "Update TB_CABSOLICITUD set NumOrden='" & Trim(cadena) & "' where Cod_Solicitud='" & Trim(Txt_NumSolComp) & "'"
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
        End If
    prSol.Close
End Sub

Private Sub CboTipo_Click()
If CboTipo.ListIndex = 2 Then
    LblPI.Visible = True
    TxtPI.Visible = True
Else
    LblPI.Visible = False
    TxtPI.Visible = False
End If
End Sub

Private Sub cmbseguro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
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
        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("ds_unidmed").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("CANT_EMPQ").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("EMPAQUE").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F5PARTARA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Caption = "Precio Unit."
        Forma_Imp
    Else
        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("f3medida").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Caption = "Costo Unit."
        If wf1visualiza_dctos = "*" Then
            dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = False
            dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
            dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
        End If
    End If

End Sub

Private Sub cmbtipopera_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If TxtCodCosto.Visible = True Then
            TxtCodCosto.SetFocus
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
    
End Sub

Sub Invisi()

'    Cmbmone.ListIndex = 1
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
'    Label12.Left = 5000
    lblmoneda(0).Visible = False
    lblmoneda(1).Visible = False
'    lblmoneda(2).Left = 5600
    txtmonto.Visible = False
    txtbase.Visible = False
    txtigv.Visible = False
'    txttotal.Left = 4905
    
End Sub

Sub Forma_Imp()

    Invisi
    
End Sub

Private Sub cmdcerrar_Click()
'pnlcosto.Visible = False
End Sub

Private Sub cmbvia_Click()
Select Case cmbvia.ListIndex
Case 0:
    Imgcarro.Visible = False
    imgavion.Visible = False
    imgbote.Visible = True
Case 1:
    imgavion.Visible = True
    Imgcarro.Visible = False
    imgbote.Visible = False
Case 2:
    Imgcarro.Visible = True
    imgavion.Visible = False
    imgbote.Visible = False
End Select
End Sub

Private Sub cmbvia_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If


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
        Case 1:
            lblmoneda(0).Caption = "US$ "
            lblmoneda(1).Caption = "US$ "
            lblmoneda(2).Caption = "US$ "
        Case 2:
            lblmoneda(0).Caption = " € "
            lblmoneda(1).Caption = " € "
            lblmoneda(2).Caption = " € "
    End Select
    If Not inicio Then swGrabacion = True

End Sub

Private Sub Cmbmone_KeyPress(KeyAscii As Integer)
    

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If


End Sub

Private Sub calcula()
On Error GoTo HNDERR
Dim afecto      As Double
Dim inafecto    As Double
Dim IGV         As Double
Dim sql         As String

    If cmbtipopera.ListIndex = 0 Then
        sql = "select sum(iif(f5afecto='*',f5valvta)) as afecto, " _
        & "sum(iif(isnull(f5afecto),f5valvta)) as inafecto, sum(f3igv) as igv from tmpOrdendeCompra"
        If rst.State = adStateOpen Then rst.Close
        
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_form, adOpenStatic, adLockOptimistic
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
    Else
        sql = "select sum(iif(f5afecto='*',f3Total)) as afecto, " _
        & "sum(iif(isnull(f5afecto),f3total)) as inafecto, sum(f3igv) as igv from tmpOrdendeCompra"
        If rst.State = adStateOpen Then rst.Close
        
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_form, adOpenStatic, adLockOptimistic
        If Not (rst.EOF) Then
            afecto = IIf(IsNull(rst.Fields("afecto")), 0, rst.Fields("afecto"))
            inafecto = IIf(IsNull(rst.Fields("inafecto")), 0, rst.Fields("inafecto"))
            IGV = IIf(IsNull(rst.Fields("igv")), 0, rst.Fields("igv"))
            
'            txtbase.Text = Format$(afecto, "####,##0.00")
'            txtmonto.Text = Format$(inafecto, "####,##0.00")
'            txtigv.Text = Format(IGV, "###,###,##0.00")
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
'If Index = 0 Then
'Me.MousePointer = vbHourglass
'Rem NSE IMPRIMIR
'If Len(Trim(Txt_NumOC.Text)) > 0 Then
'    If optcosto(0).Value Then
'        Imprime_Orden 1
'    Else
'        Imprime_Orden 2
'    End If
'Else
'    MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
'End If
'Me.MousePointer = vbDefault
'Else
'    pnlcosto.Visible = False
'End If
End Sub

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
   Case 0:
        lbltermino.Caption = "EXW"
    Case 1:
        lbltermino.Caption = "FOB"
    Case 2:
        lbltermino.Caption = "CFR"
    Case 3:
        lbltermino.Caption = "CIF"
 End Select
    If Not inicio Then swGrabacion = True

End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case Column.Caption
   Case "Codigo", "Fab.", "Descripción", "U.M.", "Cant": Text = Format(Text, "#,###,###0.00")
   Case "Precio Unit.": Text = Format(Text, "#,###,###0.0000")
   Case "Total": Text = Format(Text, "#,###,###0.0000")
    Case "Costo Unit.": Text = Format(Text, "#,###,###0.0000")
End Select
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Dim sql         As String
    If KeyCode = 113 Then
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
        
            If PnlNomPrv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
        
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(PnlNomPrv.Caption)
            
            Con_Ayu = 3
            ayuda_productos.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
              '  dxDBGrid1.Columns.ColumnByFieldName("f5c").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = wcodfab
                dxDBGrid1.Columns.ColumnByFieldName("f5partara").Value = wpartar
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = wafecto

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Columns.FocusedIndex = 4
            End If
    End Select

    End If
    
    If KeyCode = 115 Or KeyCode = 46 Then
        If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
            sw_nuevo_item = False
        End If
    End If
End Sub

'Private Sub dxDBGrid1_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal y As Single)
'
'    If dxDBGrid1.Columns.FocusedIndex = 1 Then
'        If Len(Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)) > 0 Then
'            lbldescripcion.Visible = True
'            lbldescripcion.Caption = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)
'        Else
'            lbldescripcion.Caption = ""
'            lbldescripcion.Visible = False
'        End If
'    Else
'        lbldescripcion.Caption = ""
'        lbldescripcion.Visible = False
'    End If
'
'End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Me.Height = 8085
    'Me.Width = 12015
    'Me.Left = 30
   ' Me.Top = 300
'    Txt_NumSolComp.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)

    sw_nuevo_item = True
    dxDBGrid1.Dataset.Close
    lista_oc.dxDBGrid1.Dataset.Active = False
    lista_oc.dxDBGrid1.Dataset.Refresh
    lista_oc.dxDBGrid1.Dataset.Active = True

End Sub
Private Sub DesHabilitar()
        atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
        atbmenu.Tools.ITEM("ID_Email").Enabled = True
        atbmenu.Tools.ITEM("ID_Eliminar").Enabled = True
End Sub
Private Sub define_cabecera()

'dxDBGrid1.Columns.ColumnByFieldName("ITEM").Visible = False

'    lblmoneda(0).Left = 8580
'    Label9.Left = 7575
'    txtbase.Left = 7350
    
End Sub

Private Sub Form_Load()

Dim fec     As Date
    
    Me.MousePointer = 11
    num_solcomp = ""
    If wf1show_ccosto = "N" Then
        lblccosto.Visible = False
        TxtCodCosto.Visible = False
        PnlNomCosto.Visible = False
    Else
'        lblccosto.Visible = True
'        txtcodcosto.Visible = True
'        pnlnomcosto.Visible = True
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
    
'    If loc = 2 Then
''        Call define_cabecera
''        txtmonto.Visible = False
''        txtigv.Visible = False
''        txttotal.Visible = False
''        Label10.Visible = False
''        Label11.Visible = False
''        Label12.Visible = False
''        lblmoneda(1).Visible = False
''        lblmoneda(2).Visible = False
''        lblmoneda(3).Visible = False
'
'    Else
    loc = 1
    Combo2.ListIndex = 0
    cmbseguro.ListIndex = 0
    cmbvia.ListIndex = 0
    imgbote.Visible = True
'    End If
    
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    fec = txt_fecha.Value
    Wnuevo = True
    flawigv = False
    SWcondipago = 0
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "select F1IGV from sf1param where f1codemp='" & UCase(wempresa) & "'", cnn_control
    If Not (rst.EOF) Then
         wwigv = rst.Fields("F1IGV")
    End If
    rst.Close
    
    Txt_Prove.Enabled = True
    If FlagGeneraOC = False Then
        Wnuevo = True
    End If
     
    jc = 0
    
    sw_nuevo_item = False
    
    cnombase = "TEMPLUS.MDB"
    cnomtabla = "tmpOrdendeCompra"
    
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open cconex_form
    
    Call CONFIGURA_GRID
    
    If sw_nuevo_documento = True Then
        DELETEREC_N cnomtabla, cnn_form
        Limpiar
        AdicionaItem
    Else
        inicio = True
        MODIFICAR_OC
        sw_nuevo_documento = False
        inicio = False
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
    End If
    Me.MousePointer = 1
    CboTipo.ListIndex = 0
End Sub

Sub Limpiar()

    SWcondipago = 0
    Txt_NumOC = ""
    Txt_NumSolComp = ""
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    abofechaentrega.Value = Format(Date, "dd/MM/yyyy")
    Me.TxtCli.Text = ""
    LblNomCli.Caption = ""
    txtcontacto.Text = ""
    txtcodsoli = wusuario
    Cmbmone.ListIndex = 0
    txtcodforma = ""
    pnlnomforma = ""
           
    txt_tc.Text = "2.810"
    
    Txt_Referencia = ""
    
    txtbase = "0.00"
    txtmonto = "0.00"
    txtigv.Text = "0.00"
    txttotal = "0.00"
       
    txtuupp.Text = "": txtdesuupp.Caption = ""
       
    SWcondipago = 0
    txtempresa.Text = UCase$(wempresa)
    
    txtplazo_entrega.Text = ""
    txtlugar_entrega.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    cmbtipopera.ListIndex = 1
    'Txt_Referencia.SetFocus
End Sub

Private Sub Limpia_Orden()

    PnlNomCosto.Caption = ""
    Txt_Prove.Text = ""
    PnlNomPrv.Caption = ""
    txtcontacto.Text = ""
    txtcodsoli.Text = ""
    Txt_NumSolComp.Text = ""
    pnlnomsoli.Caption = ""
    txtcodforma.Text = ""
    pnlnomforma.Caption = ""
    TxtCodCosto.Text = ""
    pnldireprv.Caption = ""
    Txt_Referencia.Text = ""
    txtobserva.Text = ""
    Txt_NumOC = ""
    
    txt_tc.Text = "2.810"
    txttotal.Text = "0.00"
    txtigv.Text = "0.00"
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    
    wgraba = 1
    
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
    
    sql = "delete * from tmpocompra"
    cnn.Execute sql
    AlmacenaQuery_sql sql, cnn
    
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
                    tempocompra!PROVEEDOR = PnlNomPrv.Caption
                    tempocompra!direccion = pnldireprv.Caption
                    tempocompra!ruc = Txt_Prove.Text
                    'tempocompra!CONTACTO = txtcontacto.Text
                    tempocompra!fecha = txt_fecha.Value
                    tempocompra!FORPAG = pnlnomforma.Caption
                    tempocompra!Moneda = Cmbmone.Text
                    tempocompra!referencia = Txt_Referencia.Text
                    tempocompra!centro = TxtCodCosto.Text
                    tempocompra!nomcentro = PnlNomCosto.Caption
                    tempocompra!OBSERVA = txtobserva.Text
                    tempocompra!SUBTOTAL = txtbase.Text
                    tempocompra!MONTOINA = txtmonto.Text
                    tempocompra!IGV = txtigv.Text
                    tempocompra!TOTAL = txttotal.Text
                    tempocompra!empresa = txtempresa.Text
                    tempocompra!ss = Txt_NumSolComp.Text
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





Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        SendKeys "{tab}"
        txtobserva.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text3_Change()
calcula
txttotal.Text = ((Val(txttotal.Text) + Val(Text4.Text)) - Val(Text3.Text))
End Sub

Private Sub Text4_Change()
calcula
txttotal.Text = ((Val(txttotal.Text) + Val(Text4.Text)) - Val(Text3.Text))
End Sub

Private Sub Text5_Change()

End Sub

Private Sub txt_fecha_LostFocus()

    If IsDate(txt_fecha.Value) Then
        If Val(txt_tc.Text & "") = 0# Then
            If rscambios.State = adStateOpen Then rscambios.Close
            If ctipoadm_bd = "M" Then
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE FECHA='" & txt_fecha.Value & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            Else
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & txt_fecha.Value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            End If
            If Not rscambios.EOF Then
                txt_tc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
            Else
                txt_tc.Text = Format(2.81, "0.000")
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

Private Sub Txt_NumSolComp_Change()

    If Not inicio Then swGrabacion = True
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
End Sub

Private Sub Txt_Prove_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_Prove_GotFocus()

    Txt_Prove.SelStart = 0: Txt_Prove.SelLength = Len(Txt_Prove)
    
End Sub

Private Sub Txt_Prove_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(Txt_Prove.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NOMPROV,F2DIRPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(Txt_Prove.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                PnlNomPrv.Caption = "" & rst.Fields("F2NOMPROV")
                pnldireprv.Caption = "" & rst.Fields("F2DIRPROV")
                GRABA_GRID Trim(Txt_Prove.Text)
                Txt_Prove_KeyPress 13
            Else
                MsgBox "El proveedor no existe. Verifique.", vbInformation, "Atención"
                Txt_Prove.SetFocus
            End If
            If rst.State = adStateOpen Then rst.Close
        End If
    End If

End Sub

Private Sub Txt_Referencia_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txt_tc_Change()

    If Not inicio Then swGrabacion = True
    
    If txt_tc.Text = " .   " Then
        txt_tc.Text = "3.200"
    End If
    
End Sub

Private Sub TxtCli_Change()
Dim NOM As New ADODB.Recordset
If Len(TxtCli.Text) = 4 Then
sql = "SELECT F2NOMprov From EF2proveedoreS WHERE F2CODprov='" & TxtCli.Text & "'"
NOM.Open sql, cnn_dbbancos, 3, 1
If NOM.RecordCount > 0 Then LblNomCli.Caption = NOM!F2NOMPROV & ""
NOM.Close
End If
Set NOM = Nothing
End Sub

Private Sub TxtCli_DblClick()
TxtCli_KeyDown 113, 0
End Sub

Private Sub TxtCli_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodpag = ""
        sw_ayuda = True
        'hlp_clientes.Show 1
        ayuda_embarcadores.Show 1
'        ayu_f_p_c.Show 1
        sw_ayuda = False
        If Len(wcodpag) > 0 Then
            TxtCli.Text = wcodpag
            LblNomCli.Caption = wnompag
'            txtcodforma_KeyPress 13
        End If
    End If
    
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
            SendKeys "{tab}"
        End If
    End If
        
End Sub

Private Sub txtcodcosto_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(TxtCodCosto.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "select f3descrip,f3direccion from centros where f3costo='" & TxtCodCosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                PnlNomCosto.Caption = Trim(rst.Fields("f3descrip") & "")
            Else
                MsgBox "Centro de costo no existe. Verifique.", vbInformation, "Atenciòn"
                TxtCodCosto.SetFocus
            End If
            rst.Close
        End If
    End If

End Sub

Private Sub txtcodforma_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodpag = ""
        sw_ayuda = True
        ayuda_formapago.Show 1
'        ayu_f_p_c.Show 1
        sw_ayuda = False
        If Len(wcodpag) > 0 Then
            txtcodforma = wcodpag
            pnlnomforma = wnompag
'            txtcodforma_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtcodforma_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtcodforma.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG='" & Trim(txtcodforma.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                pnlnomforma.Caption = Trim("" & rst!F2DESPAG)
                txtplazo_entrega.SetFocus
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
Call txtcodsoli_KeyDown(113, 0)
End Sub

Private Sub txtcodsoli_GotFocus()

    If Len(Trim(txtcodsoli.Text)) = 0 Then
        txtcodsoli.Text = wusuario
    End If

End Sub

Private Sub txtcodsoli_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodusuario = ""
        sw_ayuda = True
        ayuda_usuarios.Show 1
'        ayu_f_p_c.Show 1
        sw_ayuda = False
        If Len(wcodusuario) > 0 Then
            txtcodsoli.Text = wcodusuario
            pnlnomsoli.Caption = wnomusuario
'            txtcodforma_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcodsoli_LostFocus()

    If Len(Trim(txtcodsoli.Text)) > 0 Then
        If rst.State = adStateOpen Then rst.Close
        rst.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rst.EOF Then
            pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
        Else
            pnlnomsoli.Caption = ""
            MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
            txtcodsoli.SetFocus
        End If
        rst.Close
    End If

End Sub

Private Sub txtcontacto_GotFocus()
    txtcontacto.SelStart = 0: txtcontacto.SelLength = Len(txtcontacto)

End Sub

Private Sub txtcontacto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtempresa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Txt_NumSolComp.SetFocus
    End If

End Sub

Private Sub txtlugar_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtempresa.SetFocus
    End If

End Sub

Private Sub txtobserva_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtCodCosto.SetFocus
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
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub Txt_NumSolComp_DblClick()

    Call Txt_NumSolComp_KeyDown(113, 0)
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
    
End Sub

Private Sub Txt_NumSolComp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        FlagAcceso = False
        flagwin = True
        whelp_solicitud = 4
        FlagAcceso = False
    
        'hlp_solicitudes.Show vbModal
        ayuda_solicitudes_OC.Show 1
        
        If Len(Trim(num_solcomp)) > 0 Then
            Txt_NumSolComp = num_solcomp
            Txt_Prove.Enabled = True
        
            Call MostrarDatos
            'Txt_Prove.Text = ""
            'pnlnomprv.Caption = ""
            'pnldireprv.Caption = ""
            dxDBGrid1.Dataset.ADODataset.Requery
            txt_fecha.SetFocus
            
        End If
    End If
    
End Sub

Private Sub Txt_NumSolComp_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        num_solcomp = Txt_NumSolComp.Text
        
        If num_solcomp <> "" Then
        Txt_Prove.Enabled = True
        Call MostrarDatos
        Txt_Prove.Text = ""
        PnlNomPrv.Caption = ""
        pnldireprv.Caption = ""
        
        End If
        txt_fecha.SetFocus
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
    End If
    
End Sub

Sub MostrarDatosOC()
Dim sw_nuevo_temp   As Boolean
Dim sql             As String
Dim I               As Integer
    
    If loc = 1 Then
        With rsOrdenCab
            If Not (.EOF) Then
                txtempresa = !F4EMPRESA & ""
                If Txt_NumOC = "" Then
                    !F4NUMORD = " "
                Else
                    Txt_NumOC = Format(!F4NUMORD & "", "0000000")
                End If
                Txt_NumSolComp = !F4CODSOLICITUD & ""
                txt_fecha.Value = !F4FECEMI
                txtobserva.Text = rsOrdenCab!F4OBSERVA & ""
                txtcontacto.Text = rsOrdenCab!F4CONTACTO & ""
                TxtCli.Text = rsOrdenCab!F4CODCLI & ""
                If !F4TIPMON = "S" Then
                    Cmbmone.ListIndex = 0
                ElseIf !F4TIPMON = "D" Then
                    Cmbmone.ListIndex = 1
                Else
                    Cmbmone.ListIndex = 2
                End If
                txt_tc = Format$(!F4TIPCAM, "0.000") & ""
                txtcodforma = !F4FORPAG & ""
                Txt_Referencia = !F4REFERE & ""
                txtcodsoli = !F4CODSOL & ""
                cmbseguro.ListIndex = Val(!F4SEGURO & "")
                cmbvia.ListIndex = Val(!F4VIATRANS & "")
                Combo2.ListIndex = Val(!F4TERCOMPRA & "")
                Text2.Text = !F4PESO & ""
                Text1.Text = !F4MARCAS & ""
                CboTipo.ListIndex = Val(!f4tipo & "")
                TxtRequest.Text = !F4REQUEST & ""
                TxtResumen.Text = !F4concepto & ""
                TxtPI.Text = !f4pi & ""
                abofechaentrega.Value = Format(!F4FECENT, "DD/MM/YYYY")
                
                If loc = 2 Then
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                Else
                    txtigv = Format$(!F4IGV & "", "#,##0.00")
                    txtmonto = Format$(!F4MONINA & "", "#,##0.00")
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                    txttotal = Format$(!F4MONTO & "", "#,##0.00")
                End If
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV from EF2PROVEEDORES where F2newruc='" & !F4CODPRV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    Txt_Prove.Text = "" & rst!F2NEWRUC
                    PnlNomPrv.Caption = rst!F2NOMPROV
                    pnldireprv.Caption = IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
                    wgraba = 0
                Else
                    PnlNomPrv.Caption = "Ruc es menor a 11 digitos"
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
                If rst.State = 1 Then rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    pnlnomforma.Caption = "" & rst.Fields("F2DESPAG")
                    wgraba = 0
                End If
                rst.Close
                TxtCodCosto.Text = !F4CENTRO & ""
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT f3descrip from centros where f3costo='" & TxtCodCosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    PnlNomCosto.Caption = "" & rst.Fields("f3descrip")
                    wgraba = 0
                End If
                rst.Close
                
                txtuupp.Text = .Fields("F4UUPP") & ""
                If VALIDA_UUPP(txtuupp.Text) = True Then
                    txtdesuupp.Caption = wdeslocalidad
                End If
                
                txtlugar_entrega.Text = Left(Trim("" & !F4LUGAR_ENTREGA), 100)
        
            Else
                MsgBox "La Solicitud de Compra no existe", vbInformation, "Atención"
                Txt_NumSolComp.Enabled = True
                Txt_Referencia.SetFocus
                Exit Sub
            End If
        End With
    Else
    End If
          
    With rsOrdenDet
        sql = "SELECT * from if3orden where f4numord='" & GOC & "' and F4local = '0'"
        If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
        rsOrdenDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
            existe = True
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
                I = I + 1
                If loc = 1 Then
                    If rsOrdenDet.Fields("f4numord") = GOC Then
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
                
                        dxDBGrid1.Dataset.FieldValues("item") = I
                        dxDBGrid1.Dataset.FieldValues("f3codpro") = .Fields("f3codpro") & ""
                        If rst.State = adStateOpen Then rst.Close
                        rst.Open "SELECT P.f5nompro,P.f5codfab,P.F7codmed,M.F2DESMAR from if5pla P, EF2MARCAS M where P.f5codpro='" & rsOrdenDet!f3codpro & "' AND P.F5MARCA=M.F2CODMAR", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not (rst.EOF) Then
                            dxDBGrid1.Dataset.FieldValues("f5nompro") = rst.Fields("f5nompro") & ""
                            dxDBGrid1.Dataset.FieldValues("f5codfab") = rst!f5codfab & ""
                            dxDBGrid1.Dataset.FieldValues("f3medida") = rst!f7codmed & ""
                            dxDBGrid1.Dataset.FieldValues("f5marca") = rst!f2desmar & ""
                        End If
                        rst.Close
                        
                        If Len(Trim(.Fields("f5nompro") & "")) > 0 Then
                            dxDBGrid1.Dataset.FieldValues("f5nompro") = Trim(.Fields("f5nompro") & "")
                        End If
                        dxDBGrid1.Dataset.FieldValues("f3cencos") = .Fields("f3cencos") & ""
                        dxDBGrid1.Dataset.FieldValues("f3canpro") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("f3precos") = Format$(Val("" & .Fields("f3precos")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3pordct") = Format$(Val("" & .Fields("f3pordct")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3totdct") = Format$(Val("" & .Fields("f3totdct")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5valvta") = Format$(Val("" & .Fields("f5valvta")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5afecto") = .Fields("f5afecto")
                        dxDBGrid1.Dataset.FieldValues("f3igv") = Format$(Val("" & .Fields("f3igv")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3preuni") = Format$(Val("" & .Fields("f3preuni")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3total") = Format$(Val("" & .Fields("f3total")), "###,##0.00")
                        If Not (IsDate(rsOrdenDet!f3fentrega)) Then
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
                        Else
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = Format(rsOrdenDet!f3fentrega, "dd/mm/yyyy")
                        End If
                        dxDBGrid1.Dataset.FieldValues("check") = True
                        dxDBGrid1.Dataset.FieldValues("cant_ant") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("cant_empaq") = .Fields("f3cantemp") & ""
                        dxDBGrid1.Dataset.FieldValues("empaque") = .Fields("f3empaque") & ""
                        dxDBGrid1.Dataset.FieldValues("f5partara") = Trim(.Fields("f5partara") & "")
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
    
    If existe = True Then
       dxDBGrid1.Dataset.EnableControls
       dxDBGrid1.Dataset.Open
       dxDBGrid1.OptionEnabled = True
       existe = False
    Else
       Exit Sub
    End If
End Sub

Private Sub CONFIGURA_GRID()
    
    With dxDBGrid1.Options
'        .Set (egoEditing)
'        .Set (egoTabs)
'        .Set (egoTabThrough)
'        .Set (egoCanDelete)
'        .Set (egoCanAppend)
'        .Set (egoCanInsert)
'        .Set (egoImmediateEditor)
'        '.Set (egoShowIndicator)
'        .Set (egoCanNavigation)
'        .Set (egoHorzThrough)
'        .Set (egoVertThrough)
'        .Set (egoAutoWidth)
'        .Set (egoEnterShowEditor)
'        .Set (egoEnterThrough)
'        .Set (egoShowButtonAlways)
'
'        .Set (egoColumnSizing)
'        .Set (egoColumnMoving)
'        .Set (egoTabThrough)
'        .Set (egoConfirmDelete)
'        .Set (egoCanNavigation)
'        .Set (egoCancelOnExit)
'        .Set (egoLoadAllRecords)
'        .Set (egoShowHourGlass)
'        .Set (egoUseBookmarks)
'        .Set (egoUseLocate)
'        .Set (egoAutoCalcPreviewLines)
'        .Set (egoBandSizing)
'        .Set (egoBandMoving)
'        .Set (egoDragScroll)
'        .Set (egoExpandOnDblClick)
'        .Set (egoShowFooter)
'        .Set (egoShowGrid)
'        .Set (egoShowButtons)
'        .Set (egoNameCaseInsensitive)
'        .Set (egoShowHeader)
'        .Set (egoShowPreviewGrid)
'        .Set (egoShowBorder)
'        .Set (egoDynamicLoad)


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
    dxDBGrid1.Columns(1).Visible = False
        'carga combo centros
'    With dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").LookupColumn
'        .LookupDataset.ADODataset.ConnectionString = cnn_dbbancos
'        .LookupDataset.ADODataset.CommandText = "Select F3COSTO,f3abrev from CENTROS WHERE F3ESTNUL='N' ORDER BY f3abrev"
'        .LookupKeyField = "F3COSTO"
'        .LookupResultField = "f3abrev"
'        .LookupDataset.Active = True
'        .ListFieldIndex = 0
'        .DisplaySize = 50
'        .LookupCache = True
'        .ListFieldName = "f3abrev"
'        .ListWidth = 50
'    End With
    
    If wf1visualiza_dctos = "*" Then
        dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
    End If
    
End Sub

Sub Nueva_orden()
'Dim sql     As String
'Dim Orden   As String
'    If CboTipo.ListIndex = 2 Then
'        sql = "select f4numord from if4orden where f4numord like '" & wanno & "%' and f4local='0' and f4tipo=2 and (f4inicial<>'%' OR F4INICIAL is null) ORDER BY F4NUMORD DESC"
'    Else
'        sql = "select f4numord from if4orden where f4numord like '" & wanno & "%' and f4local='0' "
'        sql = sql & "and (f4tipo=0 or f4tipo=1) and (f4inicial<>'%' OR F4INICIAL is null) ORDER BY F4NUMORD DESC"
'    End If
'    If rst.State = adStateOpen Then rst.Close
'    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
'    If Not (rst.EOF) Then
'        Orden = wanno & "-" & Format(Mid(rst.Fields("f4numord"), 6, 5) + 1, "00000") & "/0"
'    Else
'        Orden = wanno & "-" & Format(1, "00000") & "/0"
'    End If
'    Txt_NumOC.Text = (Orden)

Dim sql     As String
Dim Orden   As String
    

        sql = "SELECT MAX(F4NUMORD) AS MAYOR FROM IF4ORDEN WHERE F4LOCAL='0'"
     
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not (rst.EOF) Then
        'Orden = rst.Fields("f4numord") + 1
        'Orden = wanno & "-" & Format(Mid(rst.Fields("f4numord"), 6, 5) + 1, "00000") & "/0"
        wmes = Month(txt_fecha.Value)
        
        Orden = "EXW" & "-" & Format(Mid(rst.Fields("MAYOR"), 5, 5) + 1, "00000") & "-" & "0" & wmes & "-" & wanno
    Else
        'Orden = wanno & "-" & Format(1, "00000") & "/0"
        wmes = Month(txt_fecha.Value)
        Orden = "EXW" & "-" & "00001" & "-" & "0" & wmes & "-" & wanno
    End If
    Txt_NumOC.Text = (Orden)
    
   
End Sub

Sub GrabarOC()
Dim codi                As String
Dim wcantidad           As Double
Dim wcc                 As String
Dim wproducto           As String
Dim sql                 As String
Dim ocompra             As Double
Dim Cant                As Double
Dim rsdetaoc            As New ADODB.Recordset
Dim ncant_ant           As Double
Dim amovs_cab(0 To 40)  As a_grabacion
Dim ctipo               As String
Dim nf4Tipo             As Integer
Dim XF4TIPO             As Integer
    flag = 0
    If Trim(Txt_NumOC.Text) <> "" Then
        jc = 1
    Else
        jc = 0
    End If
    
    sql = "select * from tmpOrdendeCompra where check"
'     SQL = "select * from detalle where check and (Not detalle.f3codpro Is Null)"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_form, adOpenStatic, adLockOptimistic
    If rst.EOF Then
        MsgBox "Debe Ingresar y/o Seleccionar Productos a Comprar", vbInformation, "Sistema de Logística"
        dxDBGrid1.SetFocus
        rst.Close
        Exit Sub
    End If
    rst.Close
    
    If Txt_Prove = "" Then MsgBox "Ingrese Código de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    If PnlNomPrv = "" Then MsgBox "Ingrese Nombre de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    If txtcodsoli = "" Then MsgBox "Ingrese Código de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    If pnlnomsoli = "" Then MsgBox "Ingrese Nombre de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    If txtcodforma = "" Then MsgBox "Ingrese código de forma de pago", 48, "Sistema de Logística": txtcodforma.SetFocus: Exit Sub
    If Cmbmone.ListIndex < 0 Then MsgBox "Seleccione moneda", 48, "Sistema de Logística": Cmbmone.SetFocus: Exit Sub
    If TxtCli.Text = "" Then MsgBox "Ingrese código del cliente", 48, "Sistema de Logística": txtcodforma.SetFocus: Exit Sub
    If Val(txt_tc.Text) = 0 Then MsgBox "Ingrese Tipo de Cambio", 48, "Sistema de Logística": txt_tc.SetFocus: Exit Sub
    
    'Nueva Versión
    If loc = 1 Then
        Select Case jc
            Case 0
                Call Nueva_orden
        End Select
    End If
        
    
    If loc = 1 Then
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL,F4TIPO from if4orden where f4numord='" & Txt_NumOC.Text & "' AND F4LOCAL = '0'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (rsOrdenCab.EOF) Then
            ctipo = "M"
            nf4Tipo = Val(rsOrdenCab!f4tipo & "")
        Else
            ctipo = "A"
            flag = 1
            sw_nuevo_documento = False
        End If
        
        amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = Txt_NumOC.Text: amovs_cab(0).TIPO = "T"
        If ctipo = "A" Then
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = "N": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = "1": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = 0: amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECGRA": amovs_cab(4).valor = Format(Date, "dd/MM/yyyy"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEGRA": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
        Else
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = rsOrdenCab.Fields("F4ESTNUL"): amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = rsOrdenCab.Fields("F4FALTA"): amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = rsOrdenCab.Fields("F4ESTVAL"): amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECMOD": amovs_cab(4).valor = Format(Date, "dd/MM/yyyy"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEMOD": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
        End If
        
        amovs_cab(6).campo = "F4CODSOL": amovs_cab(6).valor = txtcodsoli.Text: amovs_cab(6).TIPO = "T"
        amovs_cab(7).campo = "F4FECEMI": amovs_cab(7).valor = Format(txt_fecha.Value, "DD/MM/YYYY"): amovs_cab(7).TIPO = "F"
        amovs_cab(8).campo = "F4CODPRV": amovs_cab(8).valor = Txt_Prove: amovs_cab(8).TIPO = "T"
        amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = txt_tc.Text: amovs_cab(9).TIPO = "N"
        amovs_cab(10).campo = "F4FORPAG": amovs_cab(10).valor = txtcodforma.Text: amovs_cab(10).TIPO = "T"
        amovs_cab(11).campo = "F4REFERE": amovs_cab(11).valor = Txt_Referencia.Text: amovs_cab(11).TIPO = "T"
        amovs_cab(12).campo = "F4OBSERVA": amovs_cab(12).valor = txtobserva.Text: amovs_cab(12).TIPO = "T"
        amovs_cab(13).campo = "F4CENTRO": amovs_cab(13).valor = TxtCodCosto.Text: amovs_cab(13).TIPO = "T"
        amovs_cab(14).campo = "F4CODSOLICITUD": amovs_cab(14).valor = Trim(Txt_NumSolComp.Text): amovs_cab(14).TIPO = "T"
        amovs_cab(15).campo = "F4TIPMON": amovs_cab(15).valor = IIf(Cmbmone.ListIndex = 0, "S", IIf(Cmbmone.ListIndex = 1, "D", "E")): amovs_cab(15).TIPO = "T"
        amovs_cab(16).campo = "F4IGV": amovs_cab(16).valor = Val(Format(txtigv.Text, "0.00")): amovs_cab(16).TIPO = "N"
        amovs_cab(17).campo = "F4MONINA": amovs_cab(17).valor = Val(Format(txtmonto.Text, "0.00")): amovs_cab(17).TIPO = "N"
        amovs_cab(18).campo = "F4BASIMP": amovs_cab(18).valor = Val(Format(txtbase.Text, "0.00")): amovs_cab(18).TIPO = "N"
        amovs_cab(19).campo = "F4MONTO": amovs_cab(19).valor = Val(Format(txttotal.Text, "0.00")): amovs_cab(19).TIPO = "N"
        amovs_cab(20).campo = "F4LOCAL": amovs_cab(20).valor = "0": amovs_cab(20).TIPO = "T"
        amovs_cab(21).campo = "F4EMPRESA": amovs_cab(21).valor = txtempresa.Text: amovs_cab(21).TIPO = "T"
        amovs_cab(22).campo = "F4UUPP": amovs_cab(22).valor = txtuupp.Text: amovs_cab(22).TIPO = "T"
        amovs_cab(23).campo = "F4PLAZO_ENTREGA": amovs_cab(23).valor = txtplazo_entrega: amovs_cab(23).TIPO = "T"
        amovs_cab(24).campo = "F4LUGAR_ENTREGA": amovs_cab(24).valor = txtlugar_entrega.Text: amovs_cab(24).TIPO = "T"
        amovs_cab(25).campo = "F4CONTACTO": amovs_cab(25).valor = txtcontacto.Text: amovs_cab(25).TIPO = "T"
        amovs_cab(26).campo = "F4FECENT": amovs_cab(26).valor = Format(abofechaentrega.Value, "DD/MM/YYYY"): amovs_cab(26).TIPO = "F"
        amovs_cab(27).campo = "F4SEGURO": amovs_cab(27).valor = CStr(cmbseguro.ListIndex) & "": amovs_cab(27).TIPO = "T"
        amovs_cab(28).campo = "F4TERCOMPRA": amovs_cab(28).valor = CStr(Combo2.ListIndex) & "": amovs_cab(28).TIPO = "T"
        amovs_cab(29).campo = "F4VIATRANS": amovs_cab(29).valor = CStr(cmbvia.ListIndex) & "": amovs_cab(29).TIPO = "T"
        amovs_cab(30).campo = "F4PESO": amovs_cab(30).valor = Text2.Text & "": amovs_cab(30).TIPO = "T"
        amovs_cab(31).campo = "F4MARCAS": amovs_cab(31).valor = Text1.Text & "": amovs_cab(31).TIPO = "T"
        amovs_cab(32).campo = "F4CODCLI": amovs_cab(32).valor = TxtCli.Text & "": amovs_cab(32).TIPO = "T"
        amovs_cab(33).campo = "F4concepto": amovs_cab(33).valor = TxtResumen.Text & "": amovs_cab(33).TIPO = "T"
        amovs_cab(34).campo = "F4TIPO": amovs_cab(34).valor = CboTipo.ListIndex & "": amovs_cab(34).TIPO = "N"
        'XF4TIPO = 0
        'amovs_cab(34).campo = "F4TIPO": amovs_cab(34).valor = XF4TIPO & "": amovs_cab(34).TIPO = "N"
        'MsgBox amovs_cab(34).campo
        amovs_cab(35).campo = "F4REQUEST": amovs_cab(35).valor = TxtRequest.Text & "": amovs_cab(35).TIPO = "T"
        amovs_cab(36).campo = "F4PI": amovs_cab(36).valor = TxtPI.Text & "": amovs_cab(36).TIPO = "T"
        rsOrdenCab.Close
        
        If ctipo = "A" Then     '--- Nuevo
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 36, cnn_dbbancos, ""
        Else
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 36, cnn_dbbancos, "F4NUMORD = '" & Txt_NumOC.Text & _
            "' AND F4LOCAL = '0' and f4tipo=" & nf4Tipo
        End If
        
    End If
    
    '---------- GRABANDO EL DETALLE DE LA ORDEN DE COMPRA ----------------------'
    If ctipoadm_bd = "M" Then
    csql = ("delete from if3orden where f4numord= '" & Txt_NumOC.Text & "' AND F4LOCAL = '0'")
    cnn_dbbancos.Execute csql
    AlmacenaQuery_sql csql, cnn_dbbancos
    Else
    csql = ("delete * from if3orden where f4numord= '" & Txt_NumOC.Text & "' AND F4LOCAL = '0'")
    cnn_dbbancos.Execute csql
    AlmacenaQuery_sql csql, cnn_dbbancos
    End If
    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
    rsOrdenDet.Open "select * from if3orden", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    Dim w As Integer
    w = 1
    If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
    If Not rsdetaoc.EOF Then
        With rsdetaoc
            .MoveFirst
            Do While Not .EOF
            
                Rem NSE If (Len(Trim(.Fields("f3codpro")))) = 0 Or (Val(Format(.Fields("f3canpro") & "", "0.00")) = 0) Or (Val(Format(.Fields("f3precos") & "", "0.000")) = 0) Then
                Rem NSE     wgrabar = False
                Rem NSE Else
                Rem NSE     wgrabar = True
                Rem NSE End If
                
                If .Fields("check") = True Then
                    wgrabar = True
                Else
                    wgrabar = False
                End If
                
                If wgrabar Then
                
                    codi = .Fields("f3codpro")
                    'Actualiza Centro de Productos
                    
                    wcantidad = .Fields("f3canpro")
                    wcc = Trim$(TxtCodCosto.Text)
                    wproducto = Trim$(codi)
                    
                    sql = "select f3presu,f3consumido,f3ocompra from centroproductos where " _
                    & "f3costo='" & wcc & "' and f5codpro='" & wproducto & "'"
    
                    If rstaux.State = adStateOpen Then rst.Close
                    rstaux.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not (rstaux.EOF) Then
                        If jc = 0 Then  'Nuevo
                            ocompra = Val(rstaux.Fields("f3ocompra").Value)
                            rstaux.Fields("f3ocompra").Value = ocompra + wcantidad
                        Else             'Modifica
                            rstaux.Fields("f3ocompra").Value = wcantidad
                        End If
                        rstaux.Update
                    End If
                    rstaux.Close
'
'                    SQL = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD" _
'                        & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL" _
'                        & ",F3FENTREGA,F5NOMPRO,F3CANTEMP,F3EMPAQUE,F5PARTARA, F4LOCAL) VALUES " _
'                        & "('" & (Txt_NumOC.Text) & "','" & .Fields("f3codpro") & "','" & .Fields("f5codfab") & "'," _
'                        & .Fields("f3canpro") & ",'" & .Fields("f5marca") & "','" & .Fields("f3medida") & "'," _
'                        & .Fields("f3canpro") & "," & IIf(IsNull(.Fields("f3preuni")), "0", .Fields("f3preuni")) & "," & IIf(IsNull(.Fields("f3precos")), "0", .Fields("f3precos")) & "," _
'                        & IIf(IsNull(.Fields("f3pordct")), "0", .Fields("f3pordct")) & "," & IIf(IsNull(.Fields("f3totdct")), "0", .Fields("f3totdct")) & "," _
'                        & IIf(IsNull(.Fields("f5valvta")), "0", .Fields("f5valvta")) & ",'" & IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto")) & "'," _
'                        & IIf(IsNull(.Fields("f3igv")), "0", .Fields("f3igv")) & "," & .Fields("f3total") & ",'" & .Fields("f3fentrega") & "','" _
'                        & .Fields("f5nompro") & "'," & IIf(IsNull(.Fields("f3canpro") & ""), _
'                        "0", .Fields("cant_empaq" & "")) & ",'" & .Fields("EMPAQUE") & "" & "','" & _
'                        .Fields("f5partara") & "', '0')"

                    sql = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD" _
                        & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL" _
                        & ",F3FENTREGA,F5NOMPRO, F4LOCAL,F3CENCOS,item,f4tipo) VALUES " _
                        & "('" & (Txt_NumOC.Text) & "','" & .Fields("f3codpro") & "','" & .Fields("f5codfab") & "'," _
                        & .Fields("f3canpro") & ",'" & .Fields("f5marca") & "','" & .Fields("f3medida") & "'," _
                        & .Fields("f3canpro") & "," & IIf(IsNull(.Fields("f3preuni")), "0", .Fields("f3preuni")) & "," & IIf(IsNull(.Fields("f3precos")), "0", .Fields("f3precos")) & "," _
                        & IIf(IsNull(.Fields("f3pordct")), "0", .Fields("f3pordct")) & "," & IIf(IsNull(.Fields("f3totdct")), "0", .Fields("f3totdct")) & "," _
                        & IIf(IsNull(.Fields("f5valvta")), "0", .Fields("f5valvta")) & ",'" & IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto")) & "'," _
                        & IIf(IsNull(.Fields("f3igv")), "0", .Fields("f3igv")) & "," & .Fields("f3total") & ",'" & .Fields("f3fentrega") & "','" _
                        & .Fields("f5nompro") & "','0','" & .Fields("F3CENCOS") & "'," & w & "," & CboTipo.ListIndex & ")"
                    cnn_dbbancos.Execute sql
                    AlmacenaQuery_sql sql, cnn_dbbancos
                    

'                    rsOrdenDet.AddNew
'                    rsOrdenDet!F4NUMORD = (Txt_NumOC.Text)
'                    rsOrdenDet!F3CODPRO = .Fields("f3codpro")
'
'                    rsOrdenDet!F3CODFAB = .Fields("f5codfab")
'                    rsOrdenDet!F3CANPRO = .Fields("f3canpro")
'                    rsOrdenDet!F5MARCA = .Fields("f5marca")
'                    rsOrdenDet!unidad = .Fields("f3medida")
'                    rsOrdenDet!f3canfal = .Fields("f3canpro")
'                    rsOrdenDet!F3PREUNI = .Fields("f3preuni")
'                    rsOrdenDet!f3PRECOS = .Fields("f3precos")
'                    rsOrdenDet!F3PORDCT = .Fields("f3pordct")
'                    rsOrdenDet!f3totdct = .Fields("f3totdct")
'                    rsOrdenDet!f5valvta = .Fields("f5valvta")
'                    rsOrdenDet!F5AFECTO = IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto"))
'                    rsOrdenDet!F3IGV = .Fields("f3igv")
'                    rsOrdenDet!F3TOTAL = .Fields("f3total")
'                    rsOrdenDet!f3fentrega = Format$(.Fields("f3fentrega"), "dd/mm/yyyy")
'                    rsOrdenDet!F5NOMPRO = Trim(.Fields("F5NOMPRO") & "")


'                    rsOrdenDet.Update
                End If
                w = w + 1
                .MoveNext
            Loop
            'rsOrdenDet.Close
        End With
    End If
    rsdetaoc.Close
    
    Call VERIFIC_PPRV
    
    If Txt_NumSolComp.Text <> "" Then
        sql = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & Txt_NumSolComp & "'"
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
    
        If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
        rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
        If Not rsdetaoc.EOF Then
            With rsdetaoc
                .MoveFirst
                Do While Not .EOF
                    codprod = .Fields("f3codpro")
                    Rem NSE If Val("" & .Fields("f3precos")) > 0 Then
                    If .Fields("check") = True Then
                        Cant = Val("" & .Fields("f3canpro"))
                        ncant_ant = Val("" & .Fields("cant_ant"))
                        
                        csql = "update tb_detsolicitud set candis= candis+" & ncant_ant & "-" & _
                        Cant & " where cod_solicitud='" & _
                        Txt_NumSolComp.Text & "' and cod_producto='" & codprod & "'"
                        cnn_dbbancos.Execute csql
                        AlmacenaQuery_sql csql, cnn_dbbancos
                    
                    End If
                    .MoveNext
                Loop
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rst!Cant <= 0 Then
                    sql = "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & Txt_NumSolComp & "'"
                    cnn_dbbancos.Execute sql
                    AlmacenaQuery_sql sql, cnn_dbbancos
                End If
                rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                wgraba = 1
            End With
        End If
        rsdetaoc.Close
    End If
    
    MsgBox "Orden de Compra Actualizada", vbInformation, "Sistema de Logistica"
    swGrabacion = False
    
End Sub

Private Sub VERIFIC_PPRV()
Dim CODPROV     As String
Dim NOMPROV     As String
Dim NomProd     As String
Dim rsdetaoc    As New ADODB.Recordset
Dim sql         As String
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
                CODPROV = Txt_Prove.Text
                NOMPROV = PnlNomPrv.Caption
                codprod = .Fields("f3codpro") & ""
                NomProd = .Fields("f5nompro") & ""
                cmoneda = IIf(Cmbmone.ListIndex = 0, "S", IIf(Cmbmone.ListIndex = 1, "D", "E"))
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
                & "F2CODPRV='" & CODPROV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rst.RecordCount = 0 Then
'                    rst.AddNew
'                    rst!F2CODPRV = CodProv
'                    rst!F2NOMPRV = NOMPROV
'                    rst!f5codpro = codprod
'                    rst!f5nompro = NomProd
'                    rst!f5valvta = nprecos
'                    rst.Fields("F2MONEDA") = cmoneda
'                    rst.Fields("F2FECHA") = dfecha
'                    rst!f5codfab = ccodfab
'                    rst!f7codmed = ccodmed
'                    rst.Fields("F2COND_PAGO") = txtcodforma.Text
'                    rst.Fields("F2FORPAG") = txtcodforma.Text
'                    rst.Update
                    If ctipoadm_bd = "M" Then
                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "','" & dfecha & "','" _
                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
                    Else
                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "',CVDATE('" & dfecha & "'),'" _
                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
                    End If
                Else
                    If ctipoadm_bd = "M" Then
                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA='" & dfecha & "' WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
                    Else
                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA=CVDATE('" & dfecha & "') WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
                    End If
                    cnn_dbbancos.Execute (sql)
                    AlmacenaQuery_sql sql, cnn_dbbancos
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
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
End Sub

Private Sub Txt_Prove_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        sw_ayuda = True
        sw_ocompra = False
        'hlp_proveedores.Show 1
        ayuda_proveedores_ocl.Show 1
        sw_ayuda = False
        Txt_Prove.Text = wrucprov
        PnlNomPrv.Caption = wnomprov
        pnldireprv.Caption = wdirprov
        If Len(Trim(wfpagoprov)) > 0 Then
            txtcodforma.Text = wfpagoprov
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT * from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                pnlnomforma.Caption = Trim("" & rst.Fields("F2DESPAG"))
            End If
            rst.Close
        End If
'        Txt_Prove_KeyPress 13
    End If
        atbmenu.Tools.ITEM("IDGrabar").Enabled = True
        atbmenu.Tools.ITEM("IDImprimir").Enabled = True
        atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("IDAnular").Enabled = True
End Sub

Private Sub Txt_Prove_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
'        SendKeys "{tab}"
        txtcontacto.SetFocus
    End If
    
End Sub

Private Sub Txt_Referencia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
        
        SendKeys "{tab}"



    End If
    
End Sub

Private Sub Txt_TC_LostFocus()
 
    If Not IsNumeric(txt_tc) Then
        MsgBox "Dato mal ingresado ...Verifique!", vbInformation, "Sistema de Logistica"
        txt_tc.SetFocus
    End If
'
End Sub

Private Sub txtcodcosto_DblClick()
    
    txtcodcosto_KeyDown 113, 0
    
End Sub

Private Sub txtcodcosto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodcosto = ""
        sw_ayuda = True
        'hlp_centros.Show 1
        Ayuda_CC.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            TxtCodCosto = wcodcosto
            PnlNomCosto = wdescosto
            txtcodcosto_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtcodforma_DblClick()
    
    txtcodforma_KeyDown 113, 0

End Sub

Private Sub txtcodforma_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
        'txtplazo_entrega.SetFocus
    End If
    
End Sub

Private Sub txtcodsoli_Change()
On Error Resume Next
    If Not inicio Then swGrabacion = True
    If Len(Trim(txtcodsoli.Text)) > 0 Then
        If rst.State = adStateOpen Then rst.Close
        rst.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rst.EOF Then
            pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
        Else
            pnlnomsoli.Caption = "NO EXISTE"
            'MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
            txtcodsoli.SetFocus
        End If
        rst.Close
    End If
End Sub

Private Sub txtcodsoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub imprimir()
    
    LLENA_TEMPCAB
    acr_ocompra.Show 1

End Sub

Private Sub eliminar()
Dim gcodigo     As String
Dim gcant       As Double
    
    If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
    rsOrdenCab.Open "SELECT * from if4orden where f4numord='" & Txt_NumOC & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsOrdenCab.EOF Then
        If MsgBox("¿Desea Anular La Orden de Compra?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística") = 6 Then
            
            sql = "Update if4ORDEN set f4estnul='S' where F4NUMORD='" & Txt_NumOC.Text & "'"
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            
            With dxDBGrid1
                .Dataset.First
                If Not (.Dataset.EOF) Then
                    .Dataset.First
                    If .Dataset.RecordCount > 0 Then
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
                    End If
                    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
                    rsOrdenDet.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If rsOrdenDet(0).Value = 0 Then
                        csql = "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & Txt_NumSolComp & "'"
                        cnn_dbbancos.Execute csql
                        AlmacenaQuery_sql csql, cnn_dbbancos
                    Else
                        csql = "update tb_cabsolicitud set cs_estado='P' where cod_solicitud='" & Txt_NumSolComp & "'"
                        cnn_dbbancos.Execute csql
                        AlmacenaQuery_sql csql, cnn_dbbancos
                        
                    End If
                    rsOrdenDet.Close
                    MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbDefaultButton1, "Sistema de Logistica"
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
            dxDBGrid1.Columns.FocusedIndex = 1
        End If
        If Action = daPost Then
            calcula
        End If
    End If
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim rsproduc    As New ADODB.Recordset
    If Not inicio Then swGrabacion = True
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
            If PnlNomPrv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
            
            If rsproduc.State = adStateOpen Then rsproduc.Close
            sql = "SELECT B.F5CODPRO,B.F5TEXTO,B.F5NOMPRO,B.F5AFECTO,B.F5CODFAB,B.F5VALVTA,B.F7CODMED,B.F5MARCA FROM EF2PROD_PROV AS A,IF5PLA AS B WHERE A.F2CODPRV='" & wrucprov & "' AND B.F5CODFAB='" & dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "' ORDER BY B.F5CODPRO"
            rsproduc.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
            If Not rsproduc.EOF Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = rsproduc.Fields("F5CODPRO") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = rsproduc.Fields("F5CODFAB") & ""
                If Len(Trim(rsproduc.Fields("F5TEXTO")) & "") > 0 Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5TEXTO") & ""
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5NOMPRO") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("ds_unidmed").Value = rsproduc.Fields("F7CODMED") & ""
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & rsproduc.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsproduc.Fields("F5MARCA") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = rsproduc.Fields("F5AFECTO") & ""
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Val(rsproduc.Fields("F5VALVTA") & "")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format$(abofechaentrega.Value, "DD/MM/YYYY"))
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Dataset.Post
            End If
            rsproduc.Close
            Set rsproduc = Nothing
'            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("check").ColIndex - 1
'            dxDBGrid1.Columns.FocusedIndex = 4
        Case "f3canpro", "f3precos", "f3preuni", "f3pordct":
            dxDBGrid1.Dataset.Edit
            Calcula_PvtaTot
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
            sw_nuevo_item = False
            calcula
        End Select
        If dxDBGrid1.Columns.FocusedColumn.ObjectName = "check" Then
            dxDBGrid1.Dataset.Edit
            'dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
            Calcula_PvtaTot
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
            sw_nuevo_item = False
            calcula
        End If
End Sub
Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
     If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "")) = 0 Then
                    Allow = False
                Else
                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").ColIndex
                End If
            End If
        End If
        If Action = daDelete Then
            sw_detalle = True
            dxDBGrid1.Dataset.Refresh
        End If
    End If
   
'    If sw_nuevo_item = False Then
'        If Action = daInsert Then
'            If dxDBGrid1.Dataset.RecordCount > 0 Then
'                If Len(Trim(dxDBGrid1.Columns(1).Value & "")) = 0 Then
'                    Allow = False
'                End If
'            End If
'        End If
'        If Action = daDelete Then
'            dxDBGrid1.Dataset.Delete
'        End If
'    End If

End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim I               As Integer
    
    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_documento = False Then
        DELETEREC_N cnomtabla, cnn_form
        dxDBGrid1.Dataset.Refresh
    End If
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.ADODataset.CommandText = "select * from tmpOrdendeCompra"

    
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
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
            .FieldValues("item") = I
            .FieldValues("f3codpro") = ""
            .FieldValues("f5nompro") = ""
            .FieldValues("f3cencos") = ""
            .FieldValues("f3medida") = ""
            .FieldValues("f5marca") = ""
            .FieldValues("f3canpro") = Null
            .FieldValues("f3precos") = Null
            .FieldValues("f3pordct") = Null
            .FieldValues("f3totdct") = Null
            .FieldValues("f5valvta") = Null
            .FieldValues("f5afecto") = ""
            .FieldValues("f3igv") = Format(0, "###,##0.00")
            .FieldValues("f3preuni") = Format(0, "###,##0.00")
            .FieldValues("f3total") = Format(0, "###,##0.00")
            .FieldValues("f5codfab") = ""
            .FieldValues("f3fentrega") = Format$(abofechaentrega.Value, "dd/mm/yyyy")
            .FieldValues("check") = False
            .FieldValues("cant_ant") = 0#
    
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim sql As String
    
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
        'validacion
            If PnlNomPrv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
        
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(PnlNomPrv.Caption)
            Con_Ayu = 3
            ayuda_productos.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
               
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = wcodfab
                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = wafecto
                dxDBGrid1.Columns.ColumnByFieldName("f5partara").Value = wpartar

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
                dxDBGrid1.Dataset.FieldValues("f3canpro") = Format(0, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 4
            End If
            '***************
           
    End Select
     Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3cencos":
              'If dxDBGrid1.Columns.FocusedColumn.ColumnByFieldName = "f3cencos" Then
            'MsgBox "CenCos"
                           
            wcodcosto = "": wdescosto = "": wunicosto = "":
            
            Lista_Centros.Show 1
                      
            If Len(Trim(wcodcosto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3cencos").Value = wunicosto
                dxDBGrid1.Dataset.Post
               
            End If
           ' End If
            '****************
            
            dxDBGrid1.Columns.FocusedIndex = 4
    End Select
     '****************
    If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNELIMINAR" Then
        If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Count = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
                sw_detalle = False
                atbmenu.Tools("IDGrabar").Enabled = False
            Else
                dxDBGrid1.Dataset.Delete
            End If
            calcula
            sw_nuevo_item = False
        End If
    End If

End Sub

Private Sub txtplazo_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
        sql = "SELECT * from if4orden where f4numord='" & GOC & "' ANd F4LOCAL = '0'"
        rsOrdenCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
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
'        hlp_uupp.Show 1
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
                    IGV = ValVta * (wwigv / 100)
                    preciounit = nprecos + (nprecos * (wwigv / 100))
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
                AlmacenaQuery_sql sql, cnn_form
            End If
            rsprodprov.Close
            rstempdet.MoveNext
        Loop
    End If
    rstempdet.Close
    dxDBGrid1.Dataset.Refresh

End Sub

Private Sub TRASLADA_CTASXPAGAR(pnumero As String)
Dim ncorre_d            As Double
Dim amovs_cab(0 To 18)  As a_grabacion
Dim rsif4orden            As New ADODB.Recordset
Dim rsbf5pla            As New ADODB.Recordset
Dim RsProveedor         As New ADODB.Recordset
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
Dim RSPAG_DCTO          As New ADODB.Recordset

    If cnn_ctaspag.State = adStateOpen Then cnn_ctaspag.Close
    cconex_ctaspag = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos.MDB" & ";Persist Security Info=False"
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
    If RsProveedor.State = adStateOpen Then RsProveedor.Close
    csql = "SELECT F2NOMPROV,F2CODPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & cruc & "'"
    RsProveedor.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProveedor.EOF Then
        cnomprov = RsProveedor.Fields("F2NOMPROV") & ""
        ccodprov = RsProveedor.Fields("F2CODPROV") & ""
    End If
    RsProveedor.Close
    
    If RSPAG_DCTO.State = adStateOpen Then RSPAG_DCTO.Close
    RSPAG_DCTO.Open "SELECT CORRELA FROM PAG_DCTO ORDER BY CORRELA DESC", cnn_ctaspag, adOpenDynamic, adLockOptimistic
    If Not RSPAG_DCTO.EOF Then
        ncorre_d = RSPAG_DCTO.Fields("CORRELA") + 1
    Else
        ncorre_d = 1
    End If
    RSPAG_DCTO.Close
    
    cnro_comp = "O/c" & Format(pnumero, "0000000")
    Moneda = IIf(Cmbmone.ListIndex = 0, "S", IIf(Cmbmone.ListIndex = 1, "D", "E"))
    
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
    sql = ("UPDATE IF4ORDEN SET F4CORRELA=" & ncorre_d & " WHERE F4NUMORD=" & pnumero & "")
    cnn_dbbancos.Execute sql
    AlmacenaQuery_sql sql, cnn_dbbancos
End Sub
