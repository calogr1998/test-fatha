VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Begin VB.Form cons_ValesSal 
   Caption         =   "Consulta de Vales de Salida"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00808080&
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   2730
      TabIndex        =   6
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00808080&
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1185
   End
   Begin VB.PictureBox SSFrame3 
      Height          =   750
      Left            =   120
      ScaleHeight     =   690
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      Begin aBoxCtl.aBox ABODESDE 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D01000101"
         MaxValue        =   "D99991231"
         ABoxStyle       =   2
         AlignmentVertical=   2
         HideSelection   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "21/07/2004"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         NumberFormat    =   17
         DecimalPlaces   =   0
         HotAppearance   =   2
         CalendarTrailingForeColor=   -2147483629
         BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowButton      =   1
         ButtonPicture   =   "cons_ValesSal.frx":0000
         ButtonWidth     =   21
         UpDownWidth     =   14
         NullText        =   ""
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalcBtnHotStyle =   4
         CalcBackColor   =   -2147483643
         CalcBtnBackColor=   -2147483643
         CalcBtnDigitColor=   -2147483646
         CalcBtnFuntionColor=   8388736
         CalcDisplayFrameColor=   65535
         CalcHeaderBackColor=   -2147483646
      End
      Begin aBoxCtl.aBox ABOHASTA 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D01000101"
         MaxValue        =   "D99991231"
         ABoxStyle       =   2
         AlignmentVertical=   2
         HideSelection   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "21/07/2004"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         NumberFormat    =   17
         DecimalPlaces   =   0
         HotAppearance   =   2
         CalendarTrailingForeColor=   -2147483629
         BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowButton      =   1
         ButtonPicture   =   "cons_ValesSal.frx":0352
         ButtonWidth     =   21
         UpDownWidth     =   14
         NullText        =   ""
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalcBtnHotStyle =   4
         CalcBackColor   =   -2147483643
         CalcBtnBackColor=   -2147483643
         CalcBtnDigitColor=   -2147483646
         CalcBtnFuntionColor=   8388736
         CalcDisplayFrameColor=   65535
         CalcHeaderBackColor=   -2147483646
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   3030
         TabIndex        =   4
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   285
         TabIndex        =   3
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "cons_ValesSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
'    Dim rsTVS       As New ADODB.Recordset
    Dim csqlVS      As String
    
'    csqlVS = "SELECT TBVENTA_CAB.F4TIPODOCU, TBVENTA_CAB.F4SERDOC, TBVENTA_CAB.F4NUMDOC, TBVENTA_CAB.F4SERGUI, TBVENTA_CAB.F4NUMGUI, TBVENTA_CAB.F4VALE, TBVENTA_CAB.F4FECEMI, TBVENTA_CAB.F4TOTFAC, TBVENTA_CAB.F4ESTNUL,DOCUMENTOS.F2DESDOC" & _
'                " FROM TBVENTA_CAB,DOCUMENTOS " & _
'                " WHERE TBVENTA_CAB.F4TIPODOCU = DOCUMENTOS.F2CODDOC AND" & _
'                " CVDATE(TBVENTA_CAB.F4FECEMI) BETWEEN CVDATE('" & ABODESDE.Value & "') AND CVDATE('" & ABOHASTA.Value & "')"

    csqlVS = "SELECT TBVENTA_CAB.F4TIPODOCU, TBVENTA_CAB.F4SERDOC, TBVENTA_CAB.F4NUMDOC, TBVENTA_CAB.F4SERGUI, TBVENTA_CAB.F4NUMGUI, TBVENTA_CAB.F4VALE, TBVENTA_CAB.F4FECEMI, TBVENTA_CAB.F4TOTFAC, TBVENTA_CAB.F4ESTNUL, TBVENTA_CAB_1.F4VALE as VALEGUIA" & _
                " FROM TBVENTA_CAB LEFT JOIN TBVENTA_CAB AS TBVENTA_CAB_1 ON (TBVENTA_CAB.F4SERGUI = TBVENTA_CAB_1.F4SERDOC) AND (TBVENTA_CAB.F4NUMGUI = TBVENTA_CAB_1.F4NUMDOC)" & _
                " WHERE (((CVDate([TBVENTA_CAB].[F4FECEMI])) Between CVDate('" & abodesde.Value & "') And CVDate('" & abohasta.Value & "'))) ORDER BY TBVENTA_CAB.F4TIPODOCU,TBVENTA_CAB.F4SERDOC, TBVENTA_CAB.F4NUMDOC;"

'    If rsTVS.State = 1 Then rsTVS.Close
'    rsTVS.Open csqlVS, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    With Acr_ConsValesSal
        .DataControl1.ConnectionString = cnn_dbbancos
        .DataControl1.Source = csqlVS
        .lblempresa.Caption = wnomcia
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .Show 1
    End With
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = 1500
    Me.Top = 980
    abodesde.Value = Format(Date, "DD/MM/YYYY")
    abohasta.Value = Format(Date, "DD/MM/YYYY")
End Sub
