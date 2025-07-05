VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form cons_regcompras 
   Caption         =   "Consulta - Registro de Compras"
   ClientHeight    =   2505
   ClientLeft      =   4935
   ClientTop       =   2460
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5265
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5100
      Begin VB.Frame Frame2 
         Caption         =   " Ordenar "
         Height          =   825
         Left            =   135
         TabIndex        =   5
         Top             =   765
         Width           =   4875
         Begin VB.OptionButton optorden 
            Caption         =   "Fecha"
            Height          =   420
            Index           =   1
            Left            =   3555
            TabIndex        =   7
            Top             =   315
            Width           =   1005
         End
         Begin VB.OptionButton optorden 
            Caption         =   "Nro. Registro"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   6
            Top             =   405
            Value           =   -1  'True
            Width           =   2130
         End
      End
      Begin VB.TextBox txtmes 
         Height          =   330
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   1
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   1935
         TabIndex        =   2
         Top             =   405
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   465
      Left            =   2655
      TabIndex        =   3
      Top             =   1935
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Salir"
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
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   465
      Left            =   1305
      TabIndex        =   4
      Top             =   1935
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Aceptar"
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
      Font3D          =   3
   End
End
Attribute VB_Name = "cons_regcompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()

    wmesregcompras = Val(txtmes.Text)
    wordenfecha = IIf(optorden(0).Value = True, "P", "F")
    frmhlmov.Show 1
    
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    txtmes.Text = wmes

End Sub

Private Sub txtmes_GotFocus()

    txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub
