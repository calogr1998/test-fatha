VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form CodApro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Codigo de Aprobación :::"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "FrmCodApro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand SSCommand1 
      Height          =   390
      Left            =   4260
      TabIndex        =   2
      Top             =   1395
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "Ok..."
      BevelWidth      =   3
      Font3D          =   4
      AutoSize        =   1
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   825
      TabIndex        =   0
      Top             =   990
      Width           =   4230
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3345
      TabIndex        =   1
      Top             =   1425
      Width           =   795
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario :"
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
      Left            =   135
      TabIndex        =   16
      Top             =   1035
      Width           =   645
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el número resultado de su formula :"
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
      Left            =   135
      TabIndex        =   15
      Top             =   1470
      Width           =   3150
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   105
      X2              =   5040
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   5040
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   4305
      TabIndex        =   14
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   5
      Left            =   4500
      TabIndex        =   13
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   3465
      TabIndex        =   12
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   4
      Left            =   3660
      TabIndex        =   11
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   2625
      TabIndex        =   10
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   3
      Left            =   2820
      TabIndex        =   9
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1770
      TabIndex        =   8
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   2
      Left            =   1965
      TabIndex        =   7
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   930
      TabIndex        =   6
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   1
      Left            =   1125
      TabIndex        =   5
      Top             =   105
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1º Num."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   105
      TabIndex        =   4
      Top             =   495
      Width           =   735
   End
   Begin VB.Label LbNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   105
      Width           =   180
   End
End
Attribute VB_Name = "CodApro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valores
Private Sub Form_Load()
Randomize
valores = ""
For I = 0 To LbNum.Count - 1
    LbNum(I).Caption = Int((9 * Rnd) + 1)
    valores = valores & LbNum(I).Caption
Next I
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Tag = "N": Hide
End Sub

Private Sub SSCommand1_Click()
Call Text1_KeyPress(13)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text = "" Then MsgBox "Debe ingresar un valor", vbExclamation + vbSystemModal, "Mensaje: Sistema": Text1.SetFocus: Exit Sub
    If Text2.Text = "" Then MsgBox "Debe ingresar un usuario", vbExclamation + vbSystemModal, "Mensaje: Sistema": Text2.SetFocus: Exit Sub
    If Val(Text1.Text) <> numero() Then
        MsgBox "Ingreso un password incorrecto", vbExclamation + vbSystemModal, "Mensaje: Validación": Call Form_Load
    Else
        Hide
        Me.Tag = "S"
    End If
End If
End Sub
Function numero()

Dim valor As Integer
Dim restar As Integer
Set rstaprob = New ADODB.Recordset
            sql = "select us_autogenerado from ef2users where f2coduser='" & Text2.Text & "'"
            rstaprob.Open sql, cnn_dbbancos
            If rstaprob.EOF Then
                If intento = 3 Then GoTo SinAcceso
                MsgBox "Usuario Incorrecto", vbExclamation, "Sistema de Logística"
                Call Form_Load
                Text1.Text = ""
                Text2.Text = ""
                txtusuario.SetFocus
                Exit Function
            End If
    PRODUCTOX = ""
    pass = ""
    restar = 5
    
    For I = 1 To Len(rstaprob.Fields("us_autogenerado"))
        restar = 5
        valor = Mid(rstaprob.Fields("us_autogenerado"), I, 1)
        If valor = 1 Then restar = 7
        If valor = 5 Then restar = 0
        If valor = 3 Then restar = -3
        pass = pass + Chr(Asc(valor) - restar)
    Next

    Password = pass
    X = 1
    Do While Len(Password) > 0
        valor = Val(Mid(Password, 1, 1))
        Select Case valor
            Case 1 To 6
                PRODUCTOX = PRODUCTOX & Val(Mid(valores, valor, 1))
            Case 0
                If X = 1 Then
                    res = PRODUCTOX
                    signo = Mid(Password, 1, 1)
                    X = X + 1
                Else
                    res = operacion(Val(res), signo, Val(PRODUCTOX))
                    signo = Mid(Password, 1, 1)
                End If
                PRODUCTOX = ""
        End Select
        Password = Mid(Password, 2)
    Loop
    res = operacion(Val(res), signo, Val(PRODUCTOX))
    numero = res
      Exit Function

SinAcceso:
    MsgBox "Ud. no está Autorizado para realizar esta Operación", vbExclamation, "Sistema de Logística"
    intento = 0
    Unload Me
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
