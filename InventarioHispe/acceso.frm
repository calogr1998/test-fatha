VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form acceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso a Firma de Solicitud"
   ClientHeight    =   2640
   ClientLeft      =   3960
   ClientTop       =   2880
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "acceso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4095
   Begin Threed.SSCommand cmdopera 
      Height          =   420
      Index           =   0
      Left            =   315
      TabIndex        =   3
      Top             =   2115
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      MousePointer    =   99
      MouseIcon       =   "acceso.frx":0442
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1860
      Left            =   225
      TabIndex        =   0
      Top             =   90
      Width           =   3705
      _Version        =   65536
      _ExtentX        =   6535
      _ExtentY        =   3281
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtpassword 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1260
         Width           =   1860
      End
      Begin VB.TextBox txtusuario 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label lblvalor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   855
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   210
         Left            =   225
         TabIndex        =   7
         Top             =   855
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   210
         Left            =   225
         TabIndex        =   5
         Top             =   1350
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   210
         Left            =   225
         TabIndex        =   1
         Top             =   405
         Width           =   555
      End
   End
   Begin Threed.SSCommand cmdopera 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   2430
      TabIndex        =   4
      Top             =   2115
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Cancelar"
      MousePointer    =   99
      MouseIcon       =   "acceso.frx":075C
   End
End
Attribute VB_Name = "acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstUsuario As ADODB.Recordset

Private Sub cmdopera_Click(Index As Integer)
Static intento As Byte
Dim sql         As String
Dim contraseña  As String
Dim resp        As Integer

    Select Case Index
        Case 0
            intento = intento + 1
            'If txtusuario.Text <> wusuario Then
            '    If intento = 3 Then GoTo SinAcceso
            '    MsgBox "Usuario Incorrecto", vbExclamation, "Sistema de Logística"
            '    txtusuario.SetFocus
            '    Exit Sub
            'End If
            Set rstUsuario = New ADODB.Recordset
            sql = "select us_autogenerado from ef2users where f2coduser='" & txtusuario.Text & "'"
            rstUsuario.Open sql, cnn_dbbancos
            If rstUsuario.EOF Then
                If intento = 3 Then GoTo SinAcceso
                MsgBox "Usuario Incorrecto", vbExclamation, "Atención"
                txtusuario.SetFocus
                Exit Sub
            End If
            contraseña = numero()
            If (txtpassword.Text) <> contraseña Then
                If intento = 3 Then GoTo SinAcceso
                MsgBox "Password Incorrecto", vbExclamation, "Sistema de Logística"
                lblvalor.Caption = GeneraValor()
                txtpassword.SetFocus
                Exit Sub
            End If
            resp = MsgBox("Firmará la Solicitud." & Chr$(13) & "¿Está Seguro?", vbQuestion + vbDefaultButton1 + vbYesNo, "Sistema de Logística")
            If resp = vbYes Then
                CodFirmaSolicitud(1) = UCase(Trim$(txtusuario.Text))
                Unload Me
            End If
            rstUsuario.Close
            
        Case 1
            Unload Me
    End Select
    Exit Sub

SinAcceso:
    MsgBox "Ud. no está Autorizado para realizar esta Operación", vbExclamation, "Sistema de Logística"
    intento = 0
    Unload Me

End Sub

Private Sub Form_Activate()

    If Len(Trim(txtusuario.Text)) = 0 Then
        txtusuario.SetFocus
    Else
        txtpassword.SetFocus
    End If

End Sub

Private Sub Form_Load()

    txtusuario.Text = wusuario
    lblvalor.Caption = GeneraValor()
    
End Sub

Function numero()
Dim producto    As String
Dim pass        As String
Dim I           As Integer
Dim Password    As String
Dim X           As Integer
Dim valor       As Integer
Dim res         As String
Dim signo       As String
    
    producto = ""
    pass = ""
    For I = 1 To Len(rstUsuario.Fields("us_autogenerado"))
      pass = pass + Chr(Asc(Mid(rstUsuario.Fields("us_autogenerado"), I, 1)) - 5)
    Next
    Password = pass
    X = 1
    Do While Len(Password) > 0
        valor = Val(Mid(Password, 1, 1))
        Select Case valor
            Case 1 To 6
                producto = producto & Val(Mid(lblvalor.Caption, valor, 1))
            Case 0
                If X = 1 Then
                    res = producto
                    signo = Mid(Password, 1, 1)
                    X = X + 1
                Else
                    res = operacion(Val(res), signo, Val(producto))
                    signo = Mid(Password, 1, 1)
                End If
                producto = ""
        End Select
        Password = Mid(Password, 2)
    Loop
    res = operacion(Val(res), signo, Val(producto))
    numero = res

End Function

Function operacion(num1, signo, num2)
    
    Select Case signo
        Case "+"
            operacion = num1 + num2
        Case "-"
            operacion = num1 - num2
        Case "*"
            operacion = num1 * num2
    End Select
    
End Function

Public Function GeneraValor()
Dim valor As Long

    'Devuelve una cadena de 6 números aleatorios con rango de 1 a 9
    Randomize
    valor = Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1)
    GeneraValor = valor

End Function

Private Sub txtpassword_Change()
    
    If Trim$(txtpassword.Text) = Empty Then
        cmdopera(0).Enabled = False
    Else
        cmdopera(0).Enabled = True
    End If

End Sub

Private Sub txtpassword_GotFocus()
    
    txtpassword.SelStart = 0
    txtpassword.SelLength = Len(txtpassword.Text)

End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
    
    
    
    
    If KeyAscii = vbKeyReturn Then
        If cmdopera(0).Enabled = True Then
            cmdopera(0).SetFocus
        Else
            cmdopera(1).SetFocus
        End If
    End If

End Sub

Private Sub txtusuario_GotFocus()
    
    txtusuario.SelStart = 0
    txtusuario.SelLength = Len(txtusuario.Text)

End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        txtpassword.SetFocus
    End If

End Sub
