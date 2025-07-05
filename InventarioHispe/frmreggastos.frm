VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmreggastos 
   Caption         =   "Registro de Gastos"
   ClientHeight    =   6015
   ClientLeft      =   2295
   ClientTop       =   2250
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7185
   Begin ComctlLib.Toolbar tblbar 
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDAYUDA1 
      Caption         =   "..."
      Height          =   250
      Index           =   1
      Left            =   3030
      TabIndex        =   21
      ToolTipText     =   "Ayuda del Plan de Cuentas"
      Top             =   1440
      Width           =   300
   End
   Begin Threed.SSPanel pnlcuenta 
      Height          =   285
      Left            =   3420
      TabIndex        =   22
      Top             =   1440
      Width           =   3420
      _Version        =   65536
      _ExtentX        =   6032
      _ExtentY        =   503
      _StockProps     =   15
      BackColor       =   -2147483644
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
      Enabled         =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6750
      Left            =   15
      TabIndex        =   5
      Top             =   540
      Width           =   7035
      _Version        =   65536
      _ExtentX        =   12409
      _ExtentY        =   11906
      _StockProps     =   15
      BackColor       =   -2147483644
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
      Begin Threed.SSFrame SSFrame3 
         Height          =   1095
         Left            =   165
         TabIndex        =   29
         Top             =   5865
         Visible         =   0   'False
         Width           =   6675
         _Version        =   65536
         _ExtentX        =   11774
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   "Saldos Iniciales"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtegresosfin 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2880
            TabIndex        =   32
            Text            =   "0.00"
            Top             =   630
            Width           =   1365
         End
         Begin VB.TextBox txtegresos 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2880
            TabIndex        =   30
            Text            =   "0.00"
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Egresos Financieros"
            Height          =   195
            Left            =   1080
            TabIndex        =   33
            Top             =   675
            Width           =   1425
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Egresos"
            Height          =   195
            Left            =   1935
            TabIndex        =   31
            Top             =   315
            Width           =   570
         End
      End
      Begin VB.CommandButton cmdhelp 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   2250
         TabIndex        =   12
         ToolTipText     =   "Ayuda de Grupos de Gastos"
         Top             =   1260
         Width           =   300
      End
      Begin Threed.SSCheck chktotalizada 
         Height          =   270
         Left            =   285
         TabIndex        =   4
         Top             =   2940
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   476
         _StockProps     =   78
         Caption         =   "Totalizada            "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.TextBox txtgrupo 
         Height          =   285
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1215
         Width           =   540
      End
      Begin VB.TextBox txtcuenta 
         Height          =   285
         Left            =   1650
         MaxLength       =   12
         TabIndex        =   2
         Top             =   870
         Width           =   1260
      End
      Begin VB.TextBox txtdescrip 
         Height          =   285
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   1
         Top             =   525
         Width           =   5145
      End
      Begin VB.TextBox txtcodigo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   0
         Top             =   180
         Width           =   765
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   780
         Left            =   180
         TabIndex        =   10
         Top             =   4545
         Width           =   6660
         _Version        =   65536
         _ExtentX        =   11747
         _ExtentY        =   1376
         _StockProps     =   14
         Caption         =   " Tipo de Gasto "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   0
            Left            =   165
            TabIndex        =   17
            Top             =   360
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Cliente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   1
            Left            =   1050
            TabIndex        =   18
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Proveedor"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   2
            Left            =   2235
            TabIndex        =   19
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Empleado"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   4
            Left            =   5625
            TabIndex        =   20
            Top             =   360
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Ninguno"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   3
            Left            =   3375
            TabIndex        =   23
            Top             =   360
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Transferencia"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   5
            Left            =   4770
            TabIndex        =   34
            Top             =   360
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Otros"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   765
         Left            =   180
         TabIndex        =   11
         Top             =   3660
         Width           =   6660
         _Version        =   65536
         _ExtentX        =   11747
         _ExtentY        =   1349
         _StockProps     =   14
         Caption         =   " Referencia Contable "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optrefer 
            Height          =   255
            Index           =   0
            Left            =   750
            TabIndex        =   14
            Top             =   345
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Documento"
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
         Begin Threed.SSOption optrefer 
            Height          =   255
            Index           =   1
            Left            =   2730
            TabIndex        =   15
            Top             =   345
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Referencia"
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
         Begin Threed.SSOption optrefer 
            Height          =   255
            Index           =   2
            Left            =   4770
            TabIndex        =   16
            Top             =   345
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Ninguno"
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
      Begin Threed.SSPanel pnlgrupo 
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Top             =   1260
         Width           =   4185
         _Version        =   65536
         _ExtentX        =   7382
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483644
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
         Enabled         =   0   'False
      End
      Begin Threed.SSCheck chkanticipo 
         Height          =   330
         Left            =   270
         TabIndex        =   24
         Top             =   3255
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Anticipo                "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkletras 
         Height          =   330
         Left            =   2610
         TabIndex        =   25
         Top             =   2940
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Renovacion de Letras "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkcheque 
         Height          =   330
         Left            =   4590
         TabIndex        =   26
         Top             =   3270
         Width           =   2130
         _Version        =   65536
         _ExtentX        =   3757
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Cheques Devueltos     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkpagares 
         Height          =   330
         Left            =   2520
         TabIndex        =   27
         Top             =   3255
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Seguimiento/Pagarés  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkrenovvenc 
         Height          =   330
         Left            =   4635
         TabIndex        =   28
         Top             =   2940
         Width           =   2130
         _Version        =   65536
         _ExtentX        =   3757
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Gastos Financieros      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1200
         Left            =   165
         TabIndex        =   35
         Top             =   1635
         Width           =   6660
         _Version        =   65536
         _ExtentX        =   11747
         _ExtentY        =   2117
         _StockProps     =   14
         Caption         =   "Flujo de Caja"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtgrupoF 
            Height          =   285
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   37
            Top             =   300
            Width           =   555
         End
         Begin VB.CommandButton Cd_help 
            Caption         =   "..."
            Height          =   240
            Left            =   1950
            TabIndex        =   36
            ToolTipText     =   "Ayuda de Grupos de Flujo"
            Top             =   345
            Width           =   285
         End
         Begin Threed.SSPanel pnlgrupoF 
            Height          =   285
            Left            =   2310
            TabIndex        =   38
            Top             =   300
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   503
            _StockProps     =   15
            BackColor       =   -2147483644
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
            Enabled         =   0   'False
         End
         Begin Threed.SSCheck chkquincena 
            Height          =   270
            Left            =   735
            TabIndex        =   41
            Top             =   780
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   476
            _StockProps     =   78
            Caption         =   "Quincenal        "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSCheck chkprorrateo 
            Height          =   330
            Left            =   3600
            TabIndex        =   42
            Top             =   750
            Width           =   1950
            _Version        =   65536
            _ExtentX        =   3440
            _ExtentY        =   582
            _StockProps     =   78
            Caption         =   "Prorrateado        "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Grupo de Flujo"
            Height          =   195
            Left            =   135
            TabIndex        =   39
            Top             =   345
            Width           =   1140
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Gastos"
         Height          =   195
         Left            =   315
         TabIndex        =   40
         Top             =   1245
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   315
         TabIndex        =   9
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   315
         TabIndex        =   8
         Top             =   555
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   315
         TabIndex        =   7
         Top             =   210
         Width           =   495
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   2940
      Top             =   -150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":073E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":0A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":0D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":0E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":1196
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":14B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmreggastos.frx":17CA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmreggastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbtablas As DAO.Database
Dim grupgast As New ADODB.Recordset
Dim tbflujo As DAO.Recordset
Dim dbplan As DAO.Database
Dim tbplan As DAO.Recordset

Private Sub Cd_help_Click()
  
   txtgrupoF_KeyDown 113, 0
  
End Sub

Private Sub chktotalizada_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      If optrefer(0).value = True Then optrefer(0).SetFocus
      If optrefer(1).value = True Then optrefer(1).SetFocus
      If optrefer(2).value = True Then optrefer(2).SetFocus
   End If

End Sub

Private Sub CMDAYUDA1_Click(Index As Integer)

    wctacont = ""
    FrmHlpPlanCta.Show 1
    Unload FrmHlpPlanCta
    Set FrmHlpPlanCta = Nothing
    DoEvents
    If Len(Trim(wctacont)) > 0 Then
       txtcuenta.Text = Trim("" & wctacont)
       pnlcuenta.Caption = Trim("" & wnomctacont)
    End If

End Sub

Private Sub cmdhelp_Click(Index As Integer)

    LLAMADA = "ayuda"
    cod_grupo = ""
    des_grupo = ""
    frmselegrupos.Show 1
'    If Len(txtgrupo.Text) = 0 Then
        pnlgrupo.Caption = des_grupo
        txtgrupo.Text = cod_grupo
'    End If
End Sub

Private Sub Form_Load()

    Set dbplan = OpenDatabase(wrutaconta & "\db_tabla.mdb")
    Set tbplan = dbplan.OpenRecordset("cf5pla")
    
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cconexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    cnn_dbbancos.Open cconexion
'    Set dbbancos = OpenDatabase(wrutabanco & "\db_bancos.mdb")
'    Set grupgast = dbbancos.OpenRecordset("grupos")
    Set dbtabla = OpenDatabase(wrutabancos & "\db_tabla.mdb")
    Set tbflujo = dbtabla.OpenRecordset("FlujoCaja")
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbcodigos = dbcompras.OpenRecordset("BF9GIN")
    
    If sw_load_mant = True Then wgastos = ""
        tbcodigos.Index = "idcodigo"
        tbcodigos.Seek "=", "G", wgastos
        If Not tbcodigos.NoMatch Then
            txtcodigo.Text = tbcodigos.Fields("codigo") & ""
            txtcodigo.Enabled = False
            txtdescrip.Text = tbcodigos.Fields("nombre") & ""
            txtcuenta.Text = tbcodigos.Fields("cuenta") & ""
            tbplan.Index = "cf5pla"
            tbplan.Seek "=", txtcuenta.Text
            If Not tbplan.NoMatch Then
                pnlcuenta.Caption = tbplan.Fields("f5nomcta") & ""
            End If
            
            txtgrupo.Text = tbcodigos.Fields("grupo") & ""
            If grupgast.State = 1 Then grupgast.Close
            grupgast.Open "select nom_grup from grupos where cod_grup='" & txtgrupo.Text & "" & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not grupgast.EOF Then
                pnlgrupo.Caption = grupgast.Fields("nom_grup") & ""
            End If
            txtgrupoF.Text = tbcodigos.Fields("grupoflujo") & ""
            If grupgast.State = 1 Then grupgast.Close
            grupgast.Open "select nombre from grupos_flujo where codigo='" & txtgrupoF.Text & "" & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not grupgast.EOF Then
                pnlgrupoF.Caption = grupgast.Fields("nombre") & ""
            End If
    '        tbflujo.Index = "idcodigo"
    '        tbflujo.Seek "=", txtgrupoF.Text
    '        If Not tbflujo.NoMatch Then
    '            pnlgrupoF.Caption = tbflujo.Fields("descripcion") & ""
    '        End If
            chktotalizada.value = IIf(tbcodigos.Fields("conta") & "" = "*", True, False)
            chkanticipo.value = IIf(tbcodigos.Fields("anticipo") & "" = "*", True, False)
            
            chkpagares.value = IIf(tbcodigos.Fields("SEGUIM_PAGARES") & "" = "*", True, False)
            Select Case tbcodigos.Fields("refer") & ""
                Case "D": optrefer(0).value = True
                Case "R": optrefer(1).value = True
                Case "", " ": optrefer(2).value = True
            End Select
            Select Case tbcodigos.Fields("tipo") & ""
                Case "C": opttipo(0).value = True
                Case "P": opttipo(1).value = True
                Case "E": opttipo(2).value = True
                Case "T": opttipo(3).value = True
                Case "O": opttipo(5).value = True
                Case "", " ": opttipo(4).value = True
            End Select
            chkcheque.value = IIf(tbcodigos.Fields("CHEQUES_DEV") & "" = "*", True, False)
            chkrenovvenc.value = IIf(tbcodigos.Fields("gastos_financieros") & "" = "*", True, False)
            chkletras.value = IIf(tbcodigos.Fields("letras") & "" = "*" Or tbcodigos.Fields("RENOV_VENC") = "*", True, False)
            chkquincena.value = IIf(tbcodigos.Fields("quincenal") & "" = "*", True, False)
            chkprorrateo.value = IIf(tbcodigos.Fields("prorrateo") & "" = "*", True, False)
            txtegresos.Text = Format(Val(tbcodigos.Fields("egresos") & ""), "###,###,##0.00")
            txtegresosfin.Text = Format(Val(tbcodigos.Fields("egresosfin") & ""), "###,###,##0.00")
        End If
        
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error_bd

    tbplan.Close
    dbplan.Close
    
    grupgast.Close
    tbflujo.Close
    dbtabla.Close
    
    Exit Sub
    
error_bd:
    Resume Next

End Sub

Private Sub optrefer_KeyPress(Index As Integer, KeyAscii As Integer)

   If KeyAscii = 13 Then
      Select Case Index
         Case 0:
            optrefer(1).SetFocus
         Case 1:
            optrefer(2).SetFocus
         Case 2:
            If opttipo(0).value = True Then opttipo(0).SetFocus
            If opttipo(1).value = True Then opttipo(1).SetFocus
            If opttipo(2).value = True Then opttipo(2).SetFocus
            If opttipo(3).value = True Then opttipo(3).SetFocus
      End Select
   End If

End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)

   If KeyAscii = 13 Then
      Select Case Index
         Case 0: opttipo(1).SetFocus
         Case 1: opttipo(2).SetFocus
         Case 2: opttipo(3).SetFocus
         Case 3: opttipo(4).SetFocus
      End Select
   End If

End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)

   Select Case Button.Index
      Case 1:
         grabar
         wgastos = txtcodigo.Text
         wcodgasto = wgastos
         wnomgasto = txtdescrip.Text
         Unload Me
      Case 2:
         eliminar
         Unload Me
      Case 3:
         Unload Me
   End Select

End Sub

Private Sub grabar()
On Error GoTo ERROR_GRABAGASTO

    tbcodigos.Index = "idcodigo"
    tbcodigos.Seek "=", "G", txtcodigo.Text
    If Not tbcodigos.NoMatch Then
        tbcodigos.Edit
    Else
        tbcodigos.AddNew
        tbcodigos.Fields("base") = "G"
        tbcodigos.Fields("codigo") = txtcodigo.Text
    End If
    tbcodigos.Fields("nombre") = Trim(txtdescrip.Text)
    tbcodigos.Fields("cuenta") = Trim(txtcuenta.Text)
    tbcodigos.Fields("grupo") = Trim(txtgrupo.Text)
    tbcodigos.Fields("grupoflujo") = Trim(txtgrupoF.Text)
    tbcodigos.Fields("conta") = IIf(chktotalizada.value = True, "*", " ")
    If optrefer(0).value = True Then tbcodigos.Fields("refer") = "D"
    If optrefer(1).value = True Then tbcodigos.Fields("refer") = "R"
    If optrefer(2).value = True Then tbcodigos.Fields("refer") = " "
    
    If opttipo(0).value = True Then tbcodigos.Fields("tipo") = "C"
    If opttipo(1).value = True Then tbcodigos.Fields("tipo") = "P"
    If opttipo(2).value = True Then tbcodigos.Fields("tipo") = "E"
    If opttipo(3).value = True Then tbcodigos.Fields("tipo") = "T"
    If opttipo(4).value = True Then tbcodigos.Fields("tipo") = " "
    If opttipo(5).value = True Then tbcodigos.Fields("tipo") = "O"
    tbcodigos.Fields("anticipo") = IIf(chkanticipo.value = True, "*", " ")
    If chkletras.value = True And wf1renovacion = "1" Then
        tbcodigos.Fields("letras") = "*"
    Else
        tbcodigos.Fields("letras") = " "
    End If
    
    If chkletras.value = True And wf1renovacion = "2" Then
        tbcodigos.Fields("RENOV_VENC") = "*"
    Else
        tbcodigos.Fields("RENOV_VENC") = " "
    End If
    If chkquincena.value = True Then
        tbcodigos.Fields("quincenal") = "*"
    Else
        tbcodigos.Fields("quincenal") = " "
    End If
    If chkprorrateo.value = True Then
        tbcodigos.Fields("prorrateo") = "*"
    Else
        tbcodigos.Fields("prorrateo") = " "
    End If
    If chkrenovvenc.value = True Then
        tbcodigos.Fields("gastos_financieros") = "*"
    Else
        tbcodigos.Fields("gastos_financieros") = " "
    End If
    'tbcodigos.Fields("letras") = IIf(chkletras.value = True, "*", " ")
    tbcodigos.Fields("CHEQUES_DEV") = IIf(chkcheque.value = True, "*", " ")
    'tbcodigos.Fields("RENOV_VENC") = IIf(chkrenovvenc.value = True, "*", " ")
    
    tbcodigos.Fields("egresos") = Val(Format(txtegresos.Text, "0.00"))
    tbcodigos.Fields("egresosfin") = Val(Format(txtegresosfin.Text, "0.00"))
        
    tbcodigos.Update
    
    Exit Sub
    
ERROR_GRABAGASTO:
    Select Case Err
        Case 3186:
            MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
            For I = 1 To 10000
            Next
            Resume
        Case 3163:
            MsgBox Err.Description & "  Verifique la base de datos.", 48, "Bancos"
            Resume Next
        Case Else
            MsgBox Err.Description, 48, "Bancos"
            Resume Next
    End Select
   
End Sub

Private Sub eliminar()

   If MsgBox("Está seguro de eliminar el registro ? ", vbYesNo, "Bancos") = vbYes Then
      tbcodigos.Seek "=", "G", txtcodigo.Text
      If Not tbcodigos.NoMatch Then
         tbcodigos.Delete
      Else
         MsgBox "El registro no puede eliminarse porque aún no ha sido grabado", 48, "Bancos"
      End If
   End If

End Sub

Private Sub txtcodigo_GotFocus()

   txtcodigo.SelStart = 0
   txtcodigo.SelLength = Len(txtcodigo.Text)

End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      tbcodigos.Seek "=", "G", txtcodigo.Text
      If Not tbcodigos.NoMatch Then
         MsgBox "Código de gasto existe. Verifíque.", 48, "Bancos"
         txtcodigo.SetFocus
      Else
         txtdescrip.SetFocus
      End If
   End If

End Sub

Private Sub txtcodigo_LostFocus()

   If Len(Trim(txtcodigo.Text)) > 0 Then
      tbcodigos.Seek "=", "G", txtcodigo.Text
      If Not tbcodigos.NoMatch Then
         MsgBox "Código de gasto existe. Verifique.", 48, "Bancos"
         txtcodigo.SetFocus
      Else
         txtdescrip.SetFocus
      End If
   End If

End Sub

Private Sub txtcuenta_DblClick()

   txtcuenta_KeyDown 113, 0

End Sub

Private Sub txtcuenta_GotFocus()

   txtcuenta.SelStart = 0
   txtcuenta.SelLength = Len(txtcuenta.Text)

End Sub

Private Sub txtcuenta_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 113 Then
      wctacont = ""
      FrmHlpPlanCta.Show 1
      Unload FrmHlpPlanCta
      Set FrmHlpPlanCta = Nothing
      DoEvents
      If Len(Trim(wctacont)) > 0 Then
         txtcuenta.Text = Trim("" & wctacont)
         pnlcuenta.Caption = Trim("" & wnomctacont)
      End If
      txtcuenta_KeyPress 13
   End If
   
End Sub

Private Sub txtcuenta_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      tbplan.Index = "cf5pla"
      tbplan.Seek "=", txtcuenta.Text
      If Not tbplan.NoMatch Then
         pnlcuenta.Caption = tbplan.Fields("f5nomcta") & ""
      Else
         MsgBox "La cuenta contable no existe. Verifique. ", 48, "Bancos"
         pnlcuenta.Caption = ""
      End If
      txtgrupo.SetFocus
   End If

End Sub

Private Sub txtdescrip_GotFocus()

   txtdescrip.SelStart = 0
   txtdescrip.SelLength = Len(txtdescrip.Text)

End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      txtcuenta.SetFocus
   End If

End Sub

Private Sub txtgrupo_GotFocus()

   txtgrupo.SelStart = 0
   txtgrupo.SelLength = Len(txtgrupo.Text)

End Sub

Private Sub txtgrupo_KeyDown(KeyCode As Integer, Shift As Integer)
 
   If KeyCode = 113 Then
      LLAMADA = "ayuda"
      frmselegrupos.Show 1
      pnlgrupo.Caption = des_grupo
      txtgrupo.Text = cod_grupo
   End If
   
   If KeyCode = 13 Then
    If Len(Trim(txtgrupo.Text) & "") > 0 Then
       If grupgast.State = 1 Then grupgast.Close
       grupgast.Open "SELECT * FROM GRUPOS WHERE COD_GRUP='" & txtgrupo.Text & "' AND TIPO= 'G'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
       If Not grupgast.EOF Then
          pnlgrupo.Caption = grupgast.Fields("nom_grup") & ""
       Else
          MsgBox "El Grupo de Gasto No Existe. Verifíque. ", 48, "Bancos"
          txtgrupo.Text = ""
          pnlgrupo.Caption = ""
       End If
     Else
          txtgrupo.Text = ""
          pnlgrupo.Caption = ""
     End If
   End If

End Sub

Private Sub txtgrupo_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      txtgrupoF.SetFocus
   End If

End Sub

Private Sub txtgrupoF_DblClick()

   txtgrupoF_KeyDown 113, 0

End Sub

Private Sub txtgrupoF_GotFocus()

   txtgrupoF.SelStart = 0
   txtgrupoF.SelLength = Len(txtgrupoF.Text)
   
End Sub

Private Sub txtgrupoF_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 113 Then
      wdestino = "E"
      LLAMADA = "ayuda"
      frm_SeleFlujo.Show 1
      pnlgrupoF.Caption = des_grupo
      txtgrupoF.Text = cod_grupo
      txtgrupoF_KeyPress 13
   End If

End Sub

Private Sub txtgrupoF_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
        If Len(Trim(txtgrupoF.Text) & "") > 0 Then
            If grupgast.State = 1 Then grupgast.Close
            grupgast.Open "SELECT * FROM GRUPOS_FLUJO WHERE CODIGO='" & txtgrupoF.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not grupgast.EOF Then
               pnlgrupoF.Caption = grupgast.Fields("nombre") & ""
            Else
               MsgBox "El Grupo de Flujo No Existe. Verifique. ", 48, "Bancos"
               txtgrupoF.Text = ""
               pnlgrupoF.Caption = ""
            End If
         Else
              txtgrupoF.Text = ""
              pnlgrupoF.Caption = ""
         End If
   
'      tbflujo.Index = "idcodigo"
'      tbflujo.Seek "=", txtgrupoF.Text
'      If Not tbflujo.NoMatch Then
'         pnlgrupoF.Caption = tbflujo.Fields("descripcion") & ""
'         chktotalizada.SetFocus
'      Else
'         MsgBox "El grupo de flujo no existe. Verifique. ", 48, "Bancos"
'      End If
   End If
   
End Sub

