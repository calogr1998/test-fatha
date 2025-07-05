VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form inicio 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Control de Acceso"
   ClientHeight    =   3450
   ClientLeft      =   1935
   ClientTop       =   1545
   ClientWidth     =   3615
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
   Icon            =   "inicio.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   3615
   Begin VB.TextBox resolucion 
      Height          =   315
      Left            =   540
      TabIndex        =   9
      Top             =   3465
      Width           =   1995
   End
   Begin Threed.SSPanel PanelUser 
      Height          =   825
      Left            =   100
      TabIndex        =   4
      Top             =   2500
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   1455
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox Txtpasuse 
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox Txtcoduse 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   5
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   210
         TabIndex        =   8
         Top             =   495
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   390
         TabIndex        =   7
         Top             =   135
         Width           =   645
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2175
      Left            =   100
      TabIndex        =   0
      Top             =   150
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   3836
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox Txtcodemp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   900
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "CONTROL Plus - SISTEMA DE INVENTARIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   100
         TabIndex        =   3
         Top             =   50
         Width           =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   100
         TabIndex        =   2
         Top             =   1750
         Width           =   645
      End
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cpassword   As String

Private Sub Form_Activate()
    Dim resolucionX&, resolucionY&
    Txtcoduse.SetFocus

    resolucionX = Screen.Width / Screen.TwipsPerPixelX
    resolucionY = Screen.Height / Screen.TwipsPerPixelY
    resolucion.Text = CStr(resolucionX & "x" & resolucionY)
    Txtcoduse.SetFocus
    'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    Txtcoduse.SetFocus

End Sub

Private Sub Form_Load()

    wingreso = False
    ctipoadm_bd = "A"
    Me.Height = 3450
    Me.Width = 3600
    Me.left = 5000
    Me.top = 2500
    
    Txtcodemp.Text = "SUMA_2013"

End Sub

Private Sub Txtcodemp_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Txtcoduse.SetFocus
    End If

End Sub

Private Sub Txtcodemp_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
        End
    End If

End Sub

Private Sub Txtcodemp_LostFocus()

    If Len(Trim(Txtcodemp.Text)) > 0 Then
        wempresa = Trim(Txtcodemp.Text)
        wF1Dir = wempresa
        cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RUTAS.MDB" & ";Persist Security Info=False"
        cnn_control.Open cconex_control
        rscontrol.Open "SELECT * FROM SRUTAS WHERE EMPRESA ='" & Txtcodemp.Text & "'", cnn_control
        If Not rscontrol.EOF Then
            wrutaconta = Trim(rscontrol.Fields("CONTABILIDAD") & "")
            wrutabancos = Trim(rscontrol.Fields("BANCOS") & "")
            wrutatemp = Trim(rscontrol.Fields("TEMPORALES") & "")
            PanelUser.Visible = True
        Else
            Beep
            Txtcodemp.Text = ""
            Txtcodemp.SetFocus
        End If
        rscontrol.Close
        cnn_control.Close
    Else
        Txtcodemp.SetFocus
    End If

End Sub

Private Sub Txtcoduse_GotFocus()
Txtcoduse.SelStart = 0: Txtcoduse.SelLength = Len(Txtcoduse)
End Sub

Private Sub Txtcoduse_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Txtpasuse.SetFocus
    End If
    
End Sub

Private Sub Txtcoduse_LostFocus()
Dim cconex_usuarios     As String
Dim cnn_usuarios        As New ADODB.Connection

    If Len(Trim(Txtcoduse.Text)) > 0 Then
        wusuario = Txtcoduse.Text
       ' cconex_usuarios = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
        
        If ctipoadm_bd = "M" Then
            cconex_usuarios = "PROVIDER=MySqlProv;DATA SOURCE=DSN=db_bancos_gratuito"
        Else
            cconex_usuarios = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
        End If
        cnn_usuarios.Open cconex_usuarios
        rsusuarios.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(Txtcoduse.Text) & "'", cnn_usuarios, adOpenDynamic, adLockOptimistic
        If Not rsusuarios.EOF Then
            cpassword = Trim(rsusuarios.Fields("F2PASUSER") & "")
            'wnomuser = Trim(rsusuarios.Fields("F2NOMUSER") & "")
            'wcorreouser = Trim(rsusuarios.Fields("f2CORREO"))
            'wcargo = Trim(rsusuarios.Fields("f2CARGO"))
            'wusermail = "" & Trim(rsusuarios.Fields("f2USEMAIL"))
            'wpaswmail = "" & Trim(rsusuarios.Fields("f2PASWMAIL"))
            wuseractprod = "" & rsusuarios.Fields("f2useractprod")
            wuserempresa = "" & rsusuarios.Fields("f2userempresa")
        Else
            Beep
            Txtcoduse.Text = ""
            Txtcoduse.SetFocus
        End If
        rsusuarios.Close
        cnn_usuarios.Close
    End If

End Sub

Private Sub Txtpasuse_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        End
    End If

End Sub

Private Sub TxtPasUse_KeyPress(KeyAscii As Integer)
    Dim cnn_usuarios As New ADODB.Connection
    Dim cconex_usuarios As String
    If KeyAscii = 13 Then
        wusuario = Txtcoduse.Text
        If ctipoadm_bd = "M" Then
            cconex_usuarios = "PROVIDER=MySqlProv;DATA SOURCE=DSN=db_bancos_gratuito"
        Else
            cconex_usuarios = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
        End If
        cnn_usuarios.Open cconex_usuarios
        If rsusuarios.State = 1 Then rsusuarios.Close
        rsusuarios.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(Txtcoduse.Text) & "'", cnn_usuarios, adOpenDynamic, adLockOptimistic
        If Not rsusuarios.EOF Then
            cpassword = Trim(rsusuarios.Fields("F2PASUSER") & "")
            'wnomuser = Trim(rsusuarios.Fields("F2NOMUSER") & "")
            'wcorreouser = Trim(rsusuarios.Fields("f2CORREO"))
            'wcargo = Trim(rsusuarios.Fields("f2CARGO"))
            'wusermail = "" & Trim(rsusuarios.Fields("f2USEMAIL"))
            'wpaswmail = "" & Trim(rsusuarios.Fields("f2PASWMAIL"))
            wuseractprod = "" & rsusuarios.Fields("f2useractprod")
            wuserempresa = "" & rsusuarios.Fields("f2userempresa")
'            If Not IsNull(rsusuarios.Fields("F2DIRUSER")) Then
'                wrutatemp = rsusuarios.Fields("F2DIRUSER")
'            End If
            wrutatemp = wrutatemp & Trim(Txtcoduse.Text) & "\"
        End If
        If cpassword = Trim(Txtpasuse.Text) Or cpassword = UCase(Trim(Txtpasuse.Text)) Then
            Unload Me
            Menu.Show
        Else
            Beep
            Txtpasuse.Text = ""
            Txtpasuse.SetFocus
        End If
    End If
    
End Sub
