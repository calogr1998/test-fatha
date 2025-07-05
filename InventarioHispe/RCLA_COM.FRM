VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRepRegistro 
   Appearance      =   0  'Flat
   Caption         =   "Consulta Registro de Compras"
   ClientHeight    =   5160
   ClientLeft      =   1785
   ClientTop       =   1425
   ClientWidth     =   10065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5160
   ScaleWidth      =   10065
   Begin Threed.SSCommand CmdSalir 
      Height          =   390
      Left            =   5130
      TabIndex        =   13
      Top             =   4680
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2307
      _ExtentY        =   698
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand CmdAceptar 
      Height          =   390
      Left            =   3735
      TabIndex        =   12
      Top             =   4680
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox PnlProducto 
      Height          =   4545
      Left            =   45
      ScaleHeight     =   4485
      ScaleWidth      =   9900
      TabIndex        =   1
      Top             =   45
      Width           =   9960
      Begin MSMask.MaskEdBox txtdesde 
         Height          =   300
         Left            =   3510
         TabIndex        =   11
         Top             =   4080
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin Threed.SSCheck chkfechas 
         Height          =   255
         Left            =   285
         TabIndex        =   10
         Top             =   4095
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3154
         _ExtentY        =   444
         _StockProps     =   78
         Caption         =   "Rango de Fechas"
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
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2412
         Left            =   240
         TabIndex        =   9
         Top             =   1392
         Width           =   9468
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   0
            Left            =   672
            TabIndex        =   14
            Top             =   480
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   1
            Left            =   2736
            TabIndex        =   15
            Top             =   480
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   2
            Left            =   672
            TabIndex        =   16
            Top             =   672
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   3
            Left            =   2736
            TabIndex        =   17
            Top             =   672
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   4
            Left            =   672
            TabIndex        =   18
            Top             =   864
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   5
            Left            =   2736
            TabIndex        =   19
            Top             =   864
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   6
            Left            =   672
            TabIndex        =   20
            Top             =   1056
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   7
            Left            =   2736
            TabIndex        =   21
            Top             =   1056
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   210
            Index           =   8
            Left            =   675
            TabIndex        =   22
            Top             =   1260
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   9
            Left            =   2736
            TabIndex        =   23
            Top             =   1248
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   10
            Left            =   672
            TabIndex        =   24
            Top             =   1440
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   11
            Left            =   2736
            TabIndex        =   25
            Top             =   1440
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   12
            Left            =   672
            TabIndex        =   26
            Top             =   1632
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   13
            Left            =   2736
            TabIndex        =   27
            Top             =   1632
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   14
            Left            =   5040
            TabIndex        =   28
            Top             =   480
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   15
            Left            =   5040
            TabIndex        =   29
            Top             =   672
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   16
            Left            =   5040
            TabIndex        =   30
            Top             =   864
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   17
            Left            =   5040
            TabIndex        =   31
            Top             =   1056
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   18
            Left            =   5040
            TabIndex        =   32
            Top             =   1248
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   19
            Left            =   5040
            TabIndex        =   33
            Top             =   1440
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   20
            Left            =   5040
            TabIndex        =   34
            Top             =   1632
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   21
            Left            =   7344
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   22
            Left            =   7344
            TabIndex        =   36
            Top             =   672
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   210
            Index           =   23
            Left            =   7344
            TabIndex        =   37
            Top             =   870
            Visible         =   0   'False
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   24
            Left            =   7344
            TabIndex        =   38
            Top             =   1056
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   25
            Left            =   7344
            TabIndex        =   39
            Top             =   1248
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   26
            Left            =   7344
            TabIndex        =   40
            Top             =   1440
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   27
            Left            =   7344
            TabIndex        =   41
            Top             =   1632
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   28
            Left            =   672
            TabIndex        =   42
            Top             =   1824
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   29
            Left            =   672
            TabIndex        =   43
            Top             =   2016
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   30
            Left            =   2736
            TabIndex        =   44
            Top             =   1824
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   31
            Left            =   2736
            TabIndex        =   45
            Top             =   2016
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   32
            Left            =   5040
            TabIndex        =   46
            Top             =   1824
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   33
            Left            =   5040
            TabIndex        =   47
            Top             =   2016
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   34
            Left            =   7344
            TabIndex        =   48
            Top             =   1824
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
         Begin Threed.SSCheck ChkTipDoc 
            Height          =   204
            Index           =   35
            Left            =   7344
            TabIndex        =   49
            Top             =   2016
            Visible         =   0   'False
            Width           =   1548
            _Version        =   65536
            _ExtentX        =   2730
            _ExtentY        =   360
            _StockProps     =   78
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ordenado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   732
         Left            =   225
         TabIndex        =   5
         Top             =   576
         Width           =   9468
         Begin Threed.SSOption optorden 
            Height          =   204
            Index           =   2
            Left            =   7824
            TabIndex        =   8
            Top             =   336
            Width           =   1212
            _Version        =   65536
            _ExtentX        =   2138
            _ExtentY        =   360
            _StockProps     =   78
            Caption         =   "Fecha"
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
         Begin Threed.SSOption optorden 
            Height          =   204
            Index           =   1
            Left            =   4224
            TabIndex        =   7
            Top             =   336
            Width           =   1356
            _Version        =   65536
            _ExtentX        =   2392
            _ExtentY        =   360
            _StockProps     =   78
            Caption         =   "N° de Mov"
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
         Begin Threed.SSOption optorden 
            Height          =   204
            Index           =   0
            Left            =   528
            TabIndex        =   6
            Top             =   336
            Width           =   2988
            _Version        =   65536
            _ExtentX        =   5270
            _ExtentY        =   360
            _StockProps     =   78
            Caption         =   "Tipo de Documento y N° de Mov."
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
      End
      Begin VB.TextBox TxtMes 
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
         Height          =   285
         Left            =   4752
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   465
      End
      Begin MSMask.MaskEdBox txthasta 
         Height          =   300
         Left            =   6660
         TabIndex        =   50
         Top             =   4095
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin VB.Label lblhasta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Left            =   6045
         TabIndex        =   4
         Top             =   4125
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lbldesde 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   2835
         TabIndex        =   3
         Top             =   4125
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   4230
         TabIndex        =   2
         Top             =   285
         Width           =   300
      End
   End
End
Attribute VB_Name = "FrmRepRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkfechas_Click(Value As Integer)

    If Value = True Then
        lbldesde.Visible = True
        lblhasta.Visible = True
        TxtDesde.Visible = True
        TxtHasta.Visible = True
        TxtDesde.SetFocus
    Else
        lbldesde.Visible = False
        lblhasta.Visible = False
        TxtDesde.Visible = False
        TxtHasta.Visible = False
    End If

End Sub

Private Sub cmdaceptar_Click()

    Me.MousePointer = 11
    FrmRepReg.Show 1
    Me.MousePointer = 1
    CmdSalir.SetFocus
    
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
Dim sw_add  As Integer
Dim i       As Integer

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")

    i = 0
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.MoveFirst
    Do While Not TbDocumento.EOF
        If TbDocumento.Fields("f2saldos") & "" = "*" Then
            sw_add = True
            If wf1formato = "9" Then
                sw_add = True
            Else
                If UCase(TbDocumento.Fields("f2abrev") & "") = "HON" Then
                    sw_add = False
                Else
                    sw_add = True
                End If
            End If
            If sw_add = True Then
                If i >= 0 Then
                    ChkTipDoc(i).Visible = True
                    ChkTipDoc(i).Value = True
                End If
                ChkTipDoc(i).Caption = Left("" & TbDocumento.Fields("F2DESDOC"), 18) & Space(50) & "" & TbDocumento.Fields("F2CODDOC")
                i = i + 1
            End If
        End If
        TbDocumento.MoveNext
    Loop
    txtmes.Text = Format$(mes, "00")
    TxtDesde.Text = Format(Date, "dd/mm/yyyy")
    TxtHasta.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set dbcomtabla = Nothing
    
End Sub

Private Sub optorden_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtHasta.SetFocus
    End If

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            optorden(0).SetFocus
        End If
    End If

End Sub

