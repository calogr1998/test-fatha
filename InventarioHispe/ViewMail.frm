VERSION 5.00
Begin VB.Form Viewmail 
   Caption         =   "Datos para enviar correo de solicitud de aprobación"
   ClientHeight    =   6405
   ClientLeft      =   2685
   ClientTop       =   2745
   ClientWidth     =   12975
   KeyPreview      =   -1  'True
   LinkTopic       =   "MainForm"
   ScaleHeight     =   6405
   ScaleWidth      =   12975
   Begin VB.TextBox editCC 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   6855
   End
   Begin VB.TextBox editMessageText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1500
      Width           =   12675
   End
   Begin VB.TextBox editSubject 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1020
      Width           =   3975
   End
   Begin VB.TextBox editTo 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Copiar cada uno de los datos en OUTLOOK y enviar el correo para su aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   6945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Instrucciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "C.C"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   915
   End
   Begin VB.Label lblSubject 
      Caption         =   "Asunto"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label lblTo 
      Caption         =   "Para"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Viewmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    editTo.Text = destino
    editSubject.Text = asunto
    editCC.Text = "betania_bk@outlook.com.pe, responder.britania@gmail.com, psalas@betania.com.pe"
    editMessageText.Text = cuerpo
    
    editMessageText.SelStart = Len(editMessageText.Text) + 1
    editMessageText.SelLength = 0
    editMessageText.SelText = ""

End Sub


