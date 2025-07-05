VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form c_cuenta 
   Caption         =   "Consulta por Cuenta Contable"
   ClientHeight    =   3810
   ClientLeft      =   2115
   ClientTop       =   1920
   ClientWidth     =   5580
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5580
   Begin Threed.SSCommand cmdsalir 
      Height          =   375
      Left            =   2745
      TabIndex        =   4
      Top             =   3330
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3330
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3120
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      _Version        =   65536
      _ExtentX        =   9393
      _ExtentY        =   5503
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
      Begin Threed.SSFrame Frame3D2 
         Height          =   600
         Left            =   180
         TabIndex        =   13
         Top             =   2205
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   1058
         _StockProps     =   14
         Caption         =   "Moneda"
         ForeColor       =   8388608
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
         Begin VB.OptionButton opcmoneda 
            Caption         =   "Dólares"
            Height          =   285
            Index           =   1
            Left            =   3105
            TabIndex        =   15
            Top             =   225
            Width           =   1005
         End
         Begin VB.OptionButton opcmoneda 
            Caption         =   "Soles"
            Height          =   240
            Index           =   0
            Left            =   765
            TabIndex        =   14
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   645
         Left            =   180
         TabIndex        =   10
         Top             =   1395
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Modo"
         ForeColor       =   8388608
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
         Begin VB.OptionButton opcmodo 
            Caption         =   "Acumulado"
            Height          =   240
            Index           =   1
            Left            =   3105
            TabIndex        =   12
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton opcmodo 
            Caption         =   "Mensual"
            Height          =   240
            Index           =   0
            Left            =   765
            TabIndex        =   11
            Top             =   270
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin Threed.SSFrame Frame3D3 
         Height          =   735
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   1296
         _StockProps     =   14
         Caption         =   "Cuenta Contable"
         ForeColor       =   8388608
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
         Begin VB.TextBox txtctahasta 
            Height          =   285
            Left            =   3420
            TabIndex        =   9
            Top             =   315
            Width           =   1050
         End
         Begin VB.TextBox txtcuenta 
            Height          =   285
            Left            =   1035
            TabIndex        =   7
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   210
            Left            =   2745
            TabIndex        =   8
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   210
            Left            =   270
            TabIndex        =   6
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.TextBox Txtmes 
         Height          =   285
         Left            =   2610
         TabIndex        =   2
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   2070
         TabIndex        =   1
         Top             =   180
         Width           =   300
      End
   End
End
Attribute VB_Name = "c_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
    If opcmoneda(0).Value = True Then
        gmonedacta = "S"
        gmondes = "SOLES"
    Else
        gmonedacta = "D"
        gmondes = "DOLARES"
    End If
    R_CUENTA.Show 1
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  txtmes.Text = mes
End Sub

Private Sub opcmodo_Click(Index As Integer)
Select Case Index
        Case 0: opcmoneda(0).SetFocus
        Case 1: opcmoneda(0).SetFocus
    End Select
End Sub

Private Sub opcmodo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        opcmoneda(0).SetFocus
    End If
End Sub

Private Sub opcmoneda_Click(Index As Integer)
    Select Case Index
        Case 0:
            gmonedacta = "S"
            gmondes = "SOLES"
        Case 1:
            gmonedacta = "D"
            gmondes = "DOLARES"
    End Select
    cmdaceptar.SetFocus

End Sub

Private Sub opcmoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0:
                gmonedacta = "S"
                gmondes = "SOLES"
            Case 0:
                gmonedacta = "D"
                gmondes = "DOLARES"
        End Select
        cmdaceptar.SetFocus
    End If
End Sub

Private Sub txtctahasta_DblClick()
    txtctahasta_KeyDown 113, 0
End Sub

Private Sub txtctahasta_GotFocus()
  txtctahasta.SelStart = 0: txtctahasta.SelLength = Len(txtctahasta.Text)
End Sub

Private Sub txtctahasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodcon = "" & txtctahasta.Text
        FrmAyudaCon.Show 1
        txtctahasta.Text = gcodcon
        Me.MousePointer = 1
        opcmodo(0).Value = True
    End If

End Sub

Private Sub txtctahasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set dbplancta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
        Set tbplancta = dbplancta.OpenRecordset("cf5pla")
        tbplancta.Index = "cf5pla"
        tbplancta.Seek "=", txtctahasta.Text
        If Not tbplancta.NoMatch Then
            opcmodo(0).SetFocus
        Else
            MsgBox "Cuenta contable no existe. Verifique.", 48, "Compras"
            txtctahasta.SetFocus
        End If
        tbplancta.Close
        dbplancta.Close
        
    End If

End Sub

Private Sub txtcuenta_DblClick()
    txtcuenta_KeyDown 113, 0
End Sub

Private Sub txtcuenta_GotFocus()
  txtcuenta.SelStart = 0: txtcuenta.SelLength = Len(txtcuenta.Text)
End Sub

Private Sub txtcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodcon = "" & txtcuenta.Text
        FrmAyudaCon.Show 1
        txtcuenta.Text = gcodcon
        Me.MousePointer = 1
    End If

End Sub

Private Sub txtcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set dbplancta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
        Set tbplancta = dbplancta.OpenRecordset("cf5pla")
        tbplancta.Index = "cf5pla"
        tbplancta.Seek "=", txtcuenta.Text
        If Not tbplancta.NoMatch Then
            txtctahasta.Text = txtcuenta.Text
            txtctahasta.SetFocus
        Else
            MsgBox "Cuenta contable no existe. Verifique.", 48, "Compras"
            txtcuenta.SetFocus
        End If
        tbplancta.Close
        dbplancta.Close
    End If

End Sub

Private Sub txtmes_GotFocus()
  txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtmes.Text = Format(txtmes.Text, "00")
        If Val(txtmes.Text) >= 1 Or Val(txtmes.Text) <= 12 Then
            txtcuenta.SetFocus
        Else
            MsgBox "Mes incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub

Private Sub txtmes_LostFocus()
    If Len(Trim(txtmes.Text)) > 0 Then
        txtmes.Text = Format(txtmes.Text, "00")
        If Val(txtmes.Text) >= 1 Or Val(txtmes.Text) <= 12 Then
            txtcuenta.SetFocus
        Else
            MsgBox "Mes incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub
