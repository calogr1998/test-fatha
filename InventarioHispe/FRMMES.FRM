VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FRMMES 
   Appearance      =   0  'Flat
   Caption         =   "Gastos del Mes"
   ClientHeight    =   2400
   ClientLeft      =   3225
   ClientTop       =   2745
   ClientWidth     =   4230
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   4230
   Begin Threed.SSCommand btn_Salir 
      Height          =   390
      Left            =   2160
      TabIndex        =   4
      Top             =   1890
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Salir"
      Font3D          =   3
   End
   Begin Threed.SSCommand btn_Aceptar 
      Height          =   390
      Left            =   765
      TabIndex        =   3
      Top             =   1890
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Aceptar"
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D2 
      Height          =   1680
      Left            =   45
      ScaleHeight     =   1620
      ScaleWidth      =   4050
      TabIndex        =   5
      Top             =   90
      Width           =   4110
      Begin Threed.SSFrame SSFrame1 
         Height          =   735
         Left            =   90
         TabIndex        =   7
         Top             =   45
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   1296
         _StockProps     =   14
         Caption         =   "Moneda"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin Threed.SSOption option2 
            Height          =   210
            Left            =   2475
            TabIndex        =   1
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1714
            _ExtentY        =   360
            _StockProps     =   78
            Caption         =   "Dólares"
            ForeColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.02
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption option1 
            Height          =   210
            Left            =   450
            TabIndex        =   0
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1714
            _ExtentY        =   360
            _StockProps     =   78
            Caption         =   "Soles"
            ForeColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.02
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.ComboBox text1 
         Appearance      =   0  'Flat
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
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   765
         TabIndex        =   6
         Top             =   1125
         Width           =   300
      End
   End
End
Attribute VB_Name = "FRMMES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TIPMON      As String

Private Sub btn_Aceptar_Click()

    xmes = Format(Text1.ListIndex + 1, "00")
    If option1.Value = True Then
        TIPMON = "S"
    End If
    If option2.Value = True Then
        TIPMON = "D"
    End If
    Frmgruga.Show 1
    btn_Salir.SetFocus

End Sub

Private Sub btn_Aceptar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If (Val(text1.Text) > 0 And Val(text1.Text) <= 12) Then
            xmes = Format(Text1.ListIndex + 1, "00")
            If option1.Value = True Then
                TIPMON = "S"
            End If
            If option2.Value = True Then
                TIPMON = "D"
            End If
            Frmgruga.Show 1
        'Else
        '    MsgBox " Mes no valido "
        '    text1.SetFocus
        'End If
    End If
    btn_Salir.SetFocus

End Sub

Private Sub Btn_Salir_Click()

    Unload Me

End Sub

Private Sub btn_Salir_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()

    Text1.Clear
    Text1.AddItem "Enero"
    Text1.AddItem "Febrero"
    Text1.AddItem "Marzo"
    Text1.AddItem "Abril"
    Text1.AddItem "Mayo"
    Text1.AddItem "Junio"
    Text1.AddItem "Julio"
    Text1.AddItem "Agosto"
    Text1.AddItem "Setiembre"
    Text1.AddItem "Octubre"
    Text1.AddItem "Noviembre"
    Text1.AddItem "Diciembre"
    Text1.ListIndex = Val(mes) - 1

End Sub

Private Sub option1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Text1.SetFocus

End Sub

Private Sub option2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Text1.SetFocus

End Sub

Private Sub text1_Click()
'BTN_ACEPTAR.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'If (Val(text1.Text) > 0 And Val(text1.Text) <= 12) Then
            xmes = Format(Text1.ListIndex + 1, "00")
            If option1.Value = True Then
                TIPMON = "S"
            End If
            If option2.Value = True Then
                TIPMON = "D"
            End If
            Frmgruga.Show 1
        'Else
        '    MsgBox " Mes no valido "
        '    text1.SetFocus
        'End If
    End If

End Sub

