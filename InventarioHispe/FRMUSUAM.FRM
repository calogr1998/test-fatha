VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmusuam 
   Appearance      =   0  'Flat
   Caption         =   "Usuario"
   ClientHeight    =   3750
   ClientLeft      =   2835
   ClientTop       =   2175
   ClientWidth     =   6375
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   6375
   Begin Threed.SSPanel SSPanel1 
      Height          =   3660
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6270
      _Version        =   65536
      _ExtentX        =   11060
      _ExtentY        =   6456
      _StockProps     =   15
      BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlfechaing 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   900
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlusuarioing 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   495
         Width           =   4470
         _Version        =   65536
         _ExtentX        =   7885
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlhoraing 
         Height          =   285
         Left            =   4365
         TabIndex        =   15
         Top             =   900
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlusuariomod 
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Top             =   1710
         Width           =   4470
         _Version        =   65536
         _ExtentX        =   7885
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlfechamod 
         Height          =   285
         Left            =   1485
         TabIndex        =   17
         Top             =   2070
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlhoramod 
         Height          =   285
         Left            =   4410
         TabIndex        =   18
         Top             =   2070
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlusuarioimp 
         Height          =   285
         Left            =   1485
         TabIndex        =   19
         Top             =   2835
         Width           =   4470
         _Version        =   65536
         _ExtentX        =   7885
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlfechaimp 
         Height          =   285
         Left            =   1485
         TabIndex        =   20
         Top             =   3195
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
      Begin Threed.SSPanel pnlhoraimp 
         Height          =   285
         Left            =   4410
         TabIndex        =   21
         Top             =   3195
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483648
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
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "INGRESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MODIFICACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   225
         TabIndex        =   11
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   630
         TabIndex        =   10
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   630
         TabIndex        =   9
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   675
         TabIndex        =   8
         Top             =   1755
         Width           =   660
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   675
         TabIndex        =   7
         Top             =   2115
         Width           =   540
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3780
         TabIndex        =   6
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3825
         TabIndex        =   5
         Top             =   2115
         Width           =   405
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3825
         TabIndex        =   4
         Top             =   3240
         Width           =   405
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   675
         TabIndex        =   3
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   675
         TabIndex        =   2
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "IMPRESION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   225
         TabIndex        =   1
         Top             =   2565
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmusuam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    TbCabRegis.Index = "IDMESNUM"
    If wload_usuario = "I" Then
        'TbCabRegis.Seek "=", fregimp.txtmesmov.Text, fregimp.TxtNumMov.Text
    Else
        TbCabRegis.Seek "=", frmregiscom.txtmesmov.Text, frmregiscom.TxtNumMov.Text
    End If
    If Not TbCabRegis.NoMatch Then
        pnlusuarioing.Caption = TbCabRegis.Fields("f4usuarioing") & ""
        pnlfechaing.Caption = Format(TbCabRegis.Fields("f4feching") & "", "dd/mm/yyyy")
        pnlhoraing.Caption = TbCabRegis.Fields("f4horaing") & ""
        pnlusuariomod.Caption = TbCabRegis.Fields("f4usuariomod") & ""
        pnlfechamod.Caption = Format(TbCabRegis.Fields("f4fechmod") & "", "dd/mm/yyyy")
        pnlhoramod.Caption = TbCabRegis.Fields("f4horamod") & ""
        pnlusuarioimp.Caption = TbCabRegis.Fields("f4usuarioimp") & ""
        pnlfechaimp.Caption = Format(TbCabRegis.Fields("f4fechaimp") & "", "dd/mm/yyyy")
        pnlhoraimp.Caption = TbCabRegis.Fields("f4horaimp") & ""
    Else
        MsgBox "El registro aún no ha sido grabado.", 48, "Compras"
    End If

End Sub

