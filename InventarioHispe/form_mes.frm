VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Form_Mes 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Mes "
   ClientHeight    =   2520
   ClientLeft      =   4275
   ClientTop       =   4380
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2520
   ScaleWidth      =   7500
   Begin Threed.SSPanel SSPanel1 
      Height          =   1860
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   7395
      _Version        =   65536
      _ExtentX        =   13044
      _ExtentY        =   3281
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
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Enero"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   1
         Left            =   1350
         TabIndex        =   3
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Febrero"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   2
         Left            =   2550
         TabIndex        =   4
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Marzo"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   3
         Left            =   3750
         TabIndex        =   5
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Abril"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   4
         Left            =   4950
         TabIndex        =   6
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Ma&yo"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   5
         Left            =   6150
         TabIndex        =   7
         Top             =   195
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Junio"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "J&ulio"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   7
         Left            =   1350
         TabIndex        =   9
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "A&gosto"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   8
         Left            =   2550
         TabIndex        =   10
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Setiembre"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   9
         Left            =   3750
         TabIndex        =   11
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Octubre"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   10
         Left            =   4950
         TabIndex        =   12
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Noviembre"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin Threed.SSCommand com_mes 
         Height          =   495
         Index           =   11
         Left            =   6150
         TabIndex        =   13
         Top             =   795
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Diciembre"
         ForeColor       =   8388608
         Font3D          =   3
      End
      Begin VB.Label lbl_mes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2700
         TabIndex        =   14
         Top             =   1410
         Width           =   2055
      End
   End
   Begin Threed.SSCommand CmdAceptar 
      Height          =   450
      Left            =   2745
      TabIndex        =   0
      Top             =   1980
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "Form_Mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbcnt As Database
Dim tbcnt As Recordset
Dim dbbase As Database
Dim tbmeses As Recordset

Private Sub cmdaceptar_Click()

'    If wf1cnting = "*" Then
'        If Val(mes) > 2 Then
'            mes = "02"
'            MsgBox "Esta versión solo permite ingresar información en el mes de Febrero.", 48, "Compras"
'            com_mes(0).SetFocus
'        End If
'    End If

'''jcg urgente    Set dbcnt = OpenDatabase(App.Path & "\CTRCOM.MDB")
'''    Set tbcnt = dbcnt.OpenRecordset("PARAM_COM")
'''    tbcnt.Index = "idcodemp"
'''    tbcnt.Seek "=", GCODEMP
'''    If Not tbcnt.NoMatch Then
'''        If tbcnt.Fields("f1mes" & mes) = "*" Then
'''            MsgBox "Mes está cerrado. Verifique.", 48, "Compras"
'''            Exit Sub
'''        Else
'''            tbcnt.Edit
'''            tbcnt.Fields("f1proame") = wanno & mes
'''            tbcnt.Update
'''        End If
'''        tbcnt.Close
'''        dbcnt.Close
'''        Unload Me
'''    End If
    Unload Me
End Sub

Private Sub com_mes_Click(Index As Integer)

    Select Case Index
           Case 0
                lbl_mes = "Enero"
           Case 1
                lbl_mes = "Febrero"
           Case 2
                lbl_mes = "Marzo"
           Case 3
                lbl_mes = "Abril"
           Case 4
                lbl_mes = "Mayo"
           Case 5
                lbl_mes = "Junio"
           Case 6
                lbl_mes = "Julio"
           Case 7
                lbl_mes = "Agosto"
           Case 8
                lbl_mes = "Setiembre"
           Case 9
                lbl_mes = "Octubre"
           Case 10
                lbl_mes = "Noviembre"
           Case 11
                lbl_mes = "Diciembre"
   End Select
   mes = Format(Index + 1, "00")
   CmdAceptar.SetFocus

End Sub

Private Sub Form_Activate()
    
    Dim tbmeses As New ADODB.Recordset
    
    'Set dbbase = OpenDatabase(wrutabanco & "\DB_BANCOS.mdb")
    'Set tbmeses = dbbase.OpenRecordset("MESES")
    'tbmeses.Index = "IDNUMMES"
    'tbmeses.Seek "=", mes

    tbmeses.Open "Select * from meses", cnn_dbbancos, adOpenStatic, adLockReadOnly

    Select Case mes
         Case "01"
                com_mes(0).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "02"
                com_mes(1).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "03"
                com_mes(2).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "04"
                com_mes(3).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "05"
                com_mes(4).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "06"
                com_mes(5).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "07"
                com_mes(6).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "08"
                com_mes(7).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "09"
                com_mes(8).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "10"
                com_mes(9).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "11"
                com_mes(10).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
         Case "12"
                com_mes(11).SetFocus
                lbl_mes = tbmeses.Fields("F2NOMMES")
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'tbmeses.Close
    'dbbase.Close

End Sub

