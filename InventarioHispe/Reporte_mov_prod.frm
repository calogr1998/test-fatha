VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Reporte_mov_prod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reportes de Movimientos de Productos :::"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   Icon            =   "Reporte_mov_prod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reporte"
      Height          =   930
      Left            =   135
      TabIndex        =   10
      Top             =   1695
      Width           =   6015
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Left            =   810
         TabIndex        =   13
         Top             =   390
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sin Vale"
         Height          =   255
         Left            =   2340
         TabIndex        =   12
         Top             =   390
         Width           =   1005
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Con Vale"
         Height          =   255
         Left            =   4005
         TabIndex        =   11
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.TextBox txtproducto 
      Height          =   330
      Left            =   1095
      MaxLength       =   12
      TabIndex        =   0
      Top             =   345
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fecha"
      Height          =   765
      Left            =   135
      TabIndex        =   5
      Top             =   885
      Width           =   6015
      Begin aBoxCtl.aBox AboDesde 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   285
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
         Text            =   "21/05/2007"
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
         ButtonPicture   =   "Reporte_mov_prod.frx":000C
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
      Begin aBoxCtl.aBox AboHasta 
         Height          =   315
         Left            =   3900
         TabIndex        =   2
         Top             =   285
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
         Text            =   "21/05/2007"
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
         ButtonPicture   =   "Reporte_mov_prod.frx":035E
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3255
         TabIndex        =   7
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   615
         TabIndex        =   6
         Top             =   315
         Width           =   555
      End
   End
   Begin Threed.SSPanel pnlproducto 
      Height          =   330
      Left            =   2175
      TabIndex        =   8
      Top             =   345
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   345
      Left            =   1830
      TabIndex        =   3
      Top             =   2775
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "&Procesar..."
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   345
      Left            =   3195
      TabIndex        =   4
      Top             =   2775
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "&Salir"
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Producto :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   285
      TabIndex        =   9
      Top             =   375
      Width           =   735
   End
End
Attribute VB_Name = "Reporte_mov_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
AboDesde.Value = Date
AboHasta.Value = Date
End Sub

Private Sub SSCommand1_Click()
Dim sql_stk_actual As String, sql As String
sql_stk_actual = "(SELECT IF3VALES.F5CODPRO, IF3VALES.F5NOMPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I', " & _
                "IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1) FROM IF3VALES)"

sql = "SELECT IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF3VALES.F2CODALM, " & _
"IF3VALES.F4NUMVAL, IF3VALES.F3CANPRO FROM IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO " & _
"= IF5PLA.F5CODPRO WHERE " & sql_stk_actual & " > 0"



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
'            txtmedida.Text = wmedida
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
'            txtmedida.Text = rs.Fields("F7CODMED")
        ElseIf Len(Trim(txtproducto.Text)) > 0 Then
            MsgBox "Código de Producto no existe. Verifique.", 16, "Atención"
            pnlproducto.Caption = ""
        Else
            pnlproducto.Caption = ""
        End If
    
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


