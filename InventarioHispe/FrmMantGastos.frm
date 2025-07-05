VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form MantGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Mantenimiento de Gastos :::"
   ClientHeight    =   2124
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   5016
   Icon            =   "FrmMantGastos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2124
   ScaleWidth      =   5016
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Height          =   360
      Left            =   3828
      TabIndex        =   2
      Top             =   1632
      Width           =   1104
      _Version        =   65536
      _ExtentX        =   1947
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Salir"
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   732
      Left            =   108
      TabIndex        =   0
      Top             =   24
      Width           =   4824
      _Version        =   65536
      _ExtentX        =   8509
      _ExtentY        =   1291
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtcodcta 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   948
         TabIndex        =   6
         Top             =   264
         Width           =   564
      End
      Begin Threed.SSPanel txtdescta 
         Height          =   300
         Left            =   1548
         TabIndex        =   9
         Top             =   264
         Width           =   3132
         _Version        =   65536
         _ExtentX        =   5524
         _ExtentY        =   529
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre1 :"
         Height          =   192
         Left            =   132
         TabIndex        =   4
         Top             =   288
         Width           =   744
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   756
      Left            =   108
      TabIndex        =   1
      Top             =   768
      Width           =   4824
      _Version        =   65536
      _ExtentX        =   8509
      _ExtentY        =   1333
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel PnlN2 
         Height          =   300
         Left            =   1536
         TabIndex        =   8
         Top             =   300
         Width           =   3132
         _Version        =   65536
         _ExtentX        =   5524
         _ExtentY        =   529
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   948
         TabIndex        =   7
         Top             =   300
         Width           =   564
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre2 :"
         Height          =   192
         Left            =   144
         TabIndex        =   5
         Top             =   324
         Width           =   744
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   360
      Left            =   2700
      TabIndex        =   3
      Top             =   1632
      Width           =   1104
      _Version        =   65536
      _ExtentX        =   1947
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Aceptar"
   End
End
Attribute VB_Name = "MantGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand2_Click()
sql = "Update BF9GIN Set CodGasto = " & Text2.Text & " Where Codigo = '" & _
txtcodcta.Text & "'"
cnn_dbbancos.Execute sql
MsgBox "Codigo vinculado, correctamente", vbInformation + vbSystemModal, "Mensaje: InfoPlus"
Unload Me
Me.Show 1
End Sub

Private Sub Text2_DblClick()
text2_KeyDown 113, 0
End Sub

Private Sub txtcodcta_DblClick()
txtcodcta_KeyDown 113, 0
End Sub

Private Sub txtcodcta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        llampro = 1
'        txtcodcta.SetFocus
        ayuda_gastos.Show 1
        txtcodcta.Text = wcodgasto
        txtcodcta_KeyPress 13
    End If
End Sub


Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        
'        txtcodcta.SetFocus
        ayuda_Gastos_Imp.Show 1
        Text2.Text = wcodgastimp
        PnlN2.Caption = wNombregasimo
        txtcodcta_KeyPress 13
    End If
End Sub

Private Sub txtcodcta_KeyPress(KeyAscii As Integer)
Dim tbcomtab1 As ADODB.Recordset
    Set tbcomtab1 = New ADODB.Recordset
    If KeyAscii = 13 Then
        gcodppp = txtcodcta.Text
        sql = "Select * from bf9gin where codigo='" & gcodppp & "' and base='G'"
        If tbcomtab1.State = adStateOpen Then tbcomtab1.Close
        tbcomtab1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not tbcomtab1.EOF Then
        '   gcueppp = tbcomtab1.Fields("cuenta") & ""
          ' txtcodcta.Text = gcueppp
           txtdescta.Caption = tbcomtab1.Fields("nombre").Value & ""
           
        Else
            txtcodcta.Text = ""
            txtdescta.Caption = ""
            MsgBox "El Codigo ingresado no existe. Vuelva a Ingresarlo ", vbInformation, "Atencion"
            txtcodcta.SetFocus
        End If
        tbcomtab1.Close

    End If
End Sub

Private Sub SSCommand1_Click()
Unload Me
End Sub
