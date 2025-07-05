VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form cons_solicitudes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Solicitudes de Materiales"
   ClientHeight    =   3855
   ClientLeft      =   2685
   ClientTop       =   2205
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7890
   Begin Threed.SSPanel SSPanel1 
      Height          =   3030
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   7710
      _Version        =   65536
      _ExtentX        =   13600
      _ExtentY        =   5345
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
      Begin Threed.SSPanel pnlobra 
         Height          =   330
         Left            =   2340
         TabIndex        =   12
         Top             =   2295
         Width           =   5100
         _Version        =   65536
         _ExtentX        =   8996
         _ExtentY        =   582
         _StockProps     =   15
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
      Begin VB.TextBox txtobra 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1100
         TabIndex        =   3
         Top             =   2295
         Width           =   1185
      End
      Begin VB.ComboBox cmbsolicitante 
         Height          =   315
         Left            =   1100
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   6225
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1140
         Left            =   225
         TabIndex        =   7
         Top             =   855
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   2011
         _StockProps     =   14
         Caption         =   " Rango de fechas "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSMask.MaskEdBox abofdesde 
            Height          =   285
            Left            =   1215
            TabIndex        =   1
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox abofhasta 
            Height          =   285
            Left            =   5310
            TabIndex        =   2
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   4725
            TabIndex        =   9
            Top             =   495
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   495
            TabIndex        =   8
            Top             =   495
            Width           =   465
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   11
         Top             =   2340
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Width           =   735
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   2565
      TabIndex        =   4
      Top             =   3285
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   3285
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
End
Attribute VB_Name = "cons_solicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub abofdesde_GotFocus()

    'abofdesde.FocusSelect = True
    abofdesde.SelStart = 0
    abofdesde.SelLength = 2
    
End Sub

Private Sub abofdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        abofhasta.SetFocus
    End If
    
End Sub

Private Sub abofdesde_LostFocus()

    If Not IsDate(abofdesde.Text) Then
        MsgBox "Fecha inicial incorrecta. Verifique.", vbCritical, "Sistema de Logística"
        abofdesde.SetFocus
    End If

End Sub

Private Sub abofhasta_GotFocus()

    'abofhasta.FocusSelect = True
    abofhasta.SelStart = 0
    abofhasta.SelLength = 2
    
End Sub

Private Sub abofhasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtobra.SetFocus
    End If
    
End Sub

Private Sub abofhasta_LostFocus()

    If Not IsDate(abofhasta.Text) Then
        MsgBox "Fecha final incorrecta. Verifique.", vbCritical, "Sistema de Logística"
        abofhasta.SetFocus
    Else
        If abofdesde.Text > abofhasta.Text Then
            MsgBox "Rango de fechas incorrecto. Verifique.", vbCritical, "Sistema de Logística"
            abofhasta.SetFocus
        End If
    End If

End Sub

Private Sub cmbsolicitante_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        abofdesde.SetFocus
    End If

End Sub

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            cons_solicitudes_view.Show 1
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub LLENA_SOLICITANTES()
    
    rsusuarios.Open "SELECT F2CODUSER,F2NOMUSER FROM EF2USERS ORDER BY F2CODUSER", cnn_dbbancos
    If Not rsusuarios.EOF Then
        rsusuarios.MoveFirst
        Do While Not rsusuarios.EOF
            cmbsolicitante.AddItem rsusuarios.Fields("F2NOMUSER") & "" & Space(150) & rsusuarios.Fields("F2CODUSER")
            rsusuarios.MoveNext
        Loop
    End If
    rsusuarios.Close

End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1050
    abofdesde.Text = Format(Date, "dd/mm/yyyy")
    abofhasta.Text = Format(Date, "dd/mm/yyyy")
    
    LLENA_SOLICITANTES
    Me.MousePointer = vbDefault
End Sub

Private Sub txtobra_DblClick()
    
    txtobra_KeyDown 113, 0
    
End Sub

Private Sub txtobra_GotFocus()

    txtobra.SelStart = 0
    txtobra.SelLength = Len(txtobra.Text)
    
End Sub

Private Sub txtobra_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ctempusuario    As String

    If KeyCode = 113 Then
        ctempusuario = wusuario
        wusuario = Trim(right(cmbsolicitante.Text, 8))
        Ayuda_Centros.Show 1
        Rem WVR If Len(Trim(CodCosto)) > 0 Then
            txtobra.Text = Trim$(wcodcosto)
            pnlobra.Caption = Trim$(wdescosto)
            txtobra_KeyPress 13
        'End If
        wusuario = ctempusuario
    End If

End Sub

Private Sub txtobra_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdresp(0).SetFocus
    End If
    
End Sub

Private Sub txtobra_LostFocus()
Dim rstobras    As New ADODB.Recordset

    If Len(Trim(txtobra.Text)) > 0 Then
        rstobras.Open "select f3descrip from centros where f3costo='" & txtobra.Text & "'", cnn_dbbancos
        If Not rstobras.EOF Then
            pnlobra.Caption = rstobras.Fields("f3descrip") & ""
        End If
        rstobras.Close
    End If
    
End Sub
