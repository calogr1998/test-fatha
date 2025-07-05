VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frciemes 
   Appearance      =   0  'Flat
   Caption         =   "Cierre del Mes"
   ClientHeight    =   1500
   ClientLeft      =   3555
   ClientTop       =   2175
   ClientWidth     =   3735
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   3735
   Begin Threed.SSCommand cmdresp 
      Height          =   375
      Index           =   1
      Left            =   1935
      TabIndex        =   2
      Top             =   1035
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   375
      Index           =   0
      Left            =   630
      TabIndex        =   1
      Top             =   1035
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   870
      Left            =   90
      ScaleHeight     =   810
      ScaleWidth      =   3510
      TabIndex        =   3
      Top             =   45
      Width           =   3570
      Begin VB.TextBox txtmes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   0
         Top             =   225
         Width           =   420
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
         Left            =   1215
         TabIndex        =   4
         Top             =   270
         Width           =   300
      End
   End
End
Attribute VB_Name = "frciemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cierre()
Dim xopcion As String
Dim xmes    As String

    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "IDCODEMP"

    xmes = Format(txtmes.Text, "00")
    tbparametro1.Seek "=", wempresa
    If Not tbparametro1.NoMatch Then
        If tbparametro1.Fields("f1mes" & xmes) & "" = "*" Then
            If MsgBox("Está seguro de habilitar el mes ?", 36, "Compras") = 6 Then
                xopcion = " "
            Else
                xopcion = "*"
            End If
        Else
            If MsgBox("Está seguro de cerrar el mes ?", 36, "Compras") = 6 Then
                xopcion = "*"
            Else
                xopcion = " "
            End If
        End If
        tbparametro1.Edit
        tbparametro1.Fields("f1mes" & xmes) = xopcion
        tbparametro1.Update
    End If

    tbparametro1.Close
    dbcontrol.Close

End Sub

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            cierre
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

    txtmes.Text = Format(mes, "00")

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            txtmes.Text = Format(txtmes.Text, "00")
            cmdresp(0).SetFocus
        End If
    End If

End Sub

Private Sub txtmes_LostFocus()

    If Len(Trim(txtmes.Text)) > 0 Then
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            txtmes.Text = Format(txtmes.Text, "00")
            cmdresp(0).SetFocus
        Else
            MsgBox "Mes incorrecto. Verifique.", 48, "Compras"
        End If
    End If

End Sub

