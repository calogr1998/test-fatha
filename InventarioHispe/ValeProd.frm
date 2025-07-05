VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form ValeProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Actualizacion de Costos :::"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "ValeProd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Height          =   360
      Left            =   3540
      TabIndex        =   10
      Top             =   1350
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Procesar..."
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   5970
      Begin VB.TextBox txtproducto 
         Height          =   330
         Left            =   1020
         MaxLength       =   12
         TabIndex        =   3
         Top             =   675
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   225
         Width           =   1710
      End
      Begin Threed.SSPanel pnlproducto 
         Height          =   330
         Left            =   2085
         TabIndex        =   4
         Top             =   675
         Width           =   3750
         _Version        =   65536
         _ExtentX        =   6615
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483640
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   75
      TabIndex        =   6
      Top             =   1140
      Width           =   3270
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1905
         MaxLength       =   12
         TabIndex        =   9
         Top             =   255
         Width           =   1230
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "ValeProd.frx":000C
         Left            =   1005
         List            =   "ValeProd.frx":0016
         TabIndex        =   8
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   285
         Width           =   675
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   360
      Left            =   4830
      TabIndex        =   11
      Top             =   1350
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ValeProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For I = 1 To 12
    Combo1.AddItem MonthName(I)
Next I
Combo1.Text = MonthName(Month(Date))
End Sub

Private Sub SSCommand1_Click()
Dim sql As String
If Combo2.ListIndex = 0 Then
    campo = "F3Valvta = " & Val(Text1.Text) & ""
Else
    campo = "F3ValDol = " & Val(Text1.Text) & ""
End If
sql = "Update IF3VALES Set " & campo & " Where Month(F4FecVal) = " & Combo1.ListIndex + 1 & _
" And Year(F4FecVal) = " & wanno & " And Left(F4NUMVAL,1) = 'I' And F5CodPro = '" & txtproducto.Text & "'"
cnn_dbbancos.Execute sql
MsgBox "Costos Actualizados", vbInformation + vbSystemModal, "Mensaje: Sistema"
End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub
Private Sub txtproducto_DblClick()

    txtproducto_KeyDown 113, 0
    
End Sub

Private Sub txtproducto_GotFocus()

    txtproducto.SelStart = 0: txtproducto.SelLength = Len(txtproducto.Text)

End Sub

Private Sub txtproducto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        'wcod_alm = ""
        wcodproducto = ""
        sw_ayuda_prod = True
        wmarca = ""
       ' ayuda_productos.cad = ""
       Con_Ayu = 5
        ayuda_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            txtproducto.Text = wcodproducto
            pnlproducto.Caption = wdesproducto
            'txtmedida.Text = wmedida
            txtproducto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtproducto.Text = Trim(txtproducto.Text)
        sql = "SELECT F5CODPRO,F5NOMPRO, F7CODMED FROM IF5PLA WHERE F5CODPRO = '" & Trim(txtproducto) & " '"
        If rs.State = 1 Then rs.Close
        rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            pnlproducto.Caption = rs.Fields("f5nompro")
            'txtmedida.Text = rs.Fields("F7CODMED")
        ElseIf Len(Trim(txtproducto.Text)) > 0 Then
            MsgBox "Código de Producto no existe. Verifique.", 16, "Atención"
            pnlproducto.Caption = ""
        Else
            pnlproducto.Caption = ""
        End If
        'abodesde.SetFocus
    End If
    
End Sub

Private Sub txtproducto_LostFocus()
        txtproducto.Text = Trim(txtproducto.Text)
        sql = "SELECT F5CODPRO,F5NOMPRO FROM IF5PLA WHERE F5CODPRO = '" & Trim(txtproducto) & " '"
        If rs.State = 1 Then rs.Close
        rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            pnlproducto.Caption = rs.Fields("f5nompro")
        End If
End Sub


