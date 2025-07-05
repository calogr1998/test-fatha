VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form firmas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firmas"
   ClientHeight    =   3345
   ClientLeft      =   510
   ClientTop       =   2505
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   11040
   Begin Threed.SSFrame SSFrame2 
      Height          =   1725
      Left            =   5580
      TabIndex        =   6
      Top             =   585
      Width           =   5370
      _Version        =   65536
      _ExtentX        =   9472
      _ExtentY        =   3043
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
      Begin VB.Label lblfirma 
         Height          =   1140
         Index           =   4
         Left            =   2925
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblfirma 
         Height          =   1140
         Index           =   3
         Left            =   225
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Line Line2 
         X1              =   2655
         X2              =   2655
         Y1              =   135
         Y2              =   1665
      End
      Begin VB.Image imgfirma 
         Height          =   1320
         Index           =   2
         Left            =   135
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2400
      End
      Begin VB.Image imgfirma 
         Height          =   1320
         Index           =   3
         Left            =   2835
         Top             =   225
         Width           =   2400
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1725
      Left            =   90
      TabIndex        =   5
      Top             =   585
      Width           =   5370
      _Version        =   65536
      _ExtentX        =   9472
      _ExtentY        =   3043
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
      Begin VB.Label lblfirma 
         Height          =   1140
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   315
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label lblfirma 
         Height          =   1140
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Line Line1 
         X1              =   2655
         X2              =   2655
         Y1              =   135
         Y2              =   1665
      End
      Begin VB.Image imgfirma 
         Height          =   1320
         Index           =   0
         Left            =   90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2400
      End
      Begin VB.Image imgfirma 
         Height          =   1320
         Index           =   1
         Left            =   2835
         Top             =   225
         Width           =   2400
      End
   End
   Begin Threed.SSCommand cmdaceptar 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   4860
      TabIndex        =   0
      Top             =   2700
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "&Aceptar"
      MousePointer    =   99
      MouseIcon       =   "firmas.frx":0000
   End
   Begin VB.Label Label4 
      Caption         =   "FIRMA DE APROBACION 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Left            =   8460
      TabIndex        =   4
      Top             =   225
      Width           =   2310
   End
   Begin VB.Label Label3 
      Caption         =   "FIRMA DE APROBACION 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Left            =   5760
      TabIndex        =   3
      Top             =   225
      Width           =   2310
   End
   Begin VB.Label Label2 
      Caption         =   "FIRMA DE SOLICITUD 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Left            =   2880
      TabIndex        =   2
      Top             =   225
      Width           =   2310
   End
   Begin VB.Label Label1 
      Caption         =   "FIRMA DE SOLICITUD 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   225
      Width           =   2310
   End
End
Attribute VB_Name = "firmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
    
    Unload Me
    
End Sub

Public Sub PresentaFirmas()
On Error GoTo HndErr

    If Len(Trim$(CodFirmaSolicitud(1))) > 0 Then
        firma = 1
        FileCopy Devuelve_Path("BMP") & Trim$(CodFirmaSolicitud(1)) & ".bmp", Devuelve_Path("BMP") & "firma.bmp"
        imgfirma(0).Picture = LoadPicture(Devuelve_Path("BMP") & "firma.bmp")
    End If
    
    If Len(Trim$(CodFirmaAprobacion(1))) > 0 Then
        firma = 3
        FileCopy Devuelve_Path("BMP") & Trim$(CodFirmaAprobacion(1)) & ".bmp", Devuelve_Path("BMP") & "firma.bmp"
        imgfirma(2).Picture = LoadPicture(Devuelve_Path("BMP") & "firma.bmp")
    End If
    
    Exit Sub

HndErr:
    Select Case firma
        Case 1
            lblfirma(1).Caption = "Error Nº " & Err.Number & Chr$(13) & Err.Description
            lblfirma(1).Visible = True
        Case 3
            lblfirma(3).Caption = "Error Nº " & Err.Number & Chr$(13) & Err.Description
            lblfirma(3).Visible = True
    End Select

End Sub

Private Sub Form_Activate()
    
    Call PresentaFirmas

End Sub
