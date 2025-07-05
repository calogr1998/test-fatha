VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Begin VB.Form consultaInventa 
   Caption         =   "Reporte de Movimientos por Almacen"
   ClientHeight    =   2280
   ClientLeft      =   3375
   ClientTop       =   4785
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00808080&
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   1575
      TabIndex        =   9
      Top             =   1680
      Width           =   1185
   End
   Begin VB.PictureBox SSFrame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   90
      ScaleHeight     =   600
      ScaleWidth      =   5655
      TabIndex        =   2
      Top             =   840
      Width           =   5715
      Begin VB.TextBox txtcodigoal 
         Height          =   285
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   0
         Top             =   180
         Width           =   870
      End
      Begin VB.PictureBox pnldescripcional 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   2115
         ScaleHeight     =   225
         ScaleWidth      =   3375
         TabIndex        =   3
         Top             =   180
         Width           =   3435
         Begin VB.TextBox TXTDESCRIPCION 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen:"
         Height          =   285
         Left            =   225
         TabIndex        =   8
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblalmacen 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   4
         Top             =   2790
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00808080&
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   2745
      TabIndex        =   1
      Top             =   1680
      Width           =   1185
   End
   Begin VB.PictureBox SSFrame3 
      Height          =   750
      Left            =   90
      ScaleHeight     =   690
      ScaleWidth      =   5655
      TabIndex        =   5
      Top             =   45
      Width           =   5715
      Begin aBoxCtl.aBox ABODESDE 
         Height          =   315
         Left            =   960
         TabIndex        =   10
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
         Text            =   "01/03/2003"
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
         ButtonPicture   =   "consultainventa.frx":0000
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
         TabIndex        =   11
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
         Text            =   "01/03/2003"
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
         ButtonPicture   =   "consultainventa.frx":0352
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   3030
         TabIndex        =   6
         Top             =   240
         Width           =   510
      End
   End
End
Attribute VB_Name = "consultaInventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim cnomalm  As String
Private Sub abodesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ABOHASTA.SetFocus
    End If
End Sub
Private Sub abohasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtcodigoal.SetFocus
    End If
End Sub
Private Sub cmdaceptar_Click()
Dim fechai As String
Dim fechaf As String
Dim codAl As String
Dim SQL As String

fechai = ABODESDE.Value
fechaf = ABOHASTA.Value
codAl = txtcodigoal.Text
'IF5PLA.CODFAB,
SQL = "SELECT IF3VALES.F5CODPRO,IF5PLA.CODFAB, IF5PLA.F5NOMPRO, " & _
      " IIf(IF3VALES.INGSAL='I',IF3VALES.F3CANPRO,0) AS INGRESOS, " & _
      " IIf(IF3VALES.INGSAL='S',IF3VALES.F3CANPRO,0) AS SALIDAS, " & _
      " IF3VALES.F4FECVAL, IF3VALES.INGSAL " & _
      " FROM IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO " & _
      " WHERE (((CVDATE(IF3VALES.F4FECVAL))>='" & CVDate(fechai) & "' And " & _
      " (CVDATE(IF3VALES.F4FECVAL))<='" & CVDate(fechaf) & "') AND ((IF3VALES.F2CODALM)='" & codAl & "'))"

'SQL = "SELECT IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, " & _
 '     " IIf(IF3VALES.INGSAL='I',IF3VALES.F3CANPRO,0) AS INGRESOS, " & _
  '    " IIf(IF3VALES.INGSAL='S',IF3VALES.F3CANPRO,0) AS SALIDAS, " & _
   '   " IF3VALES.F4FECVAL, IF3VALES.INGSAL " & _
    '  " FROM IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO " & _
     ' " WHERE (((CVDATE(IF3VALES.F4FECVAL))>='" & CVDate(fechai) & "' And " & _
     ' " (CVDATE(IF3VALES.F4FECVAL))<='" & CVDate(fechaf) & "') AND ((IF3VALES.F2CODALM)='" & codAl & "'))"

Coneccion
       With acrInventa
            .DataControl1.ConnectionString = cn
            .DataControl1.Source = SQL
            .lblalmacen.Caption = TXTDESCRIPCION
            .lblF1.Caption = fechai
            .LblF2.Caption = fechaf
            .lblFecha.Caption = Date
            .Show 1
     End With
End Sub
Private Sub cmdcancelar_Click()
   Unload Me
End Sub
Private Sub Form_Load()
    Me.Left = 2500
    Me.Top = 1900
    Me.Height = 2685
    Me.Width = 6000
    ABODESDE.Value = Format(Date, "dd/mm/yyyy")
    ABOHASTA.Value = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If cn.State = adStateOpen Then cn.Close
End Sub
Private Sub txtcodigoal_Change()
    TXTDESCRIPCION.Text = ""
End Sub
Private Sub txtcodigoal_DblClick()
    txtcodigoal_KeyDown 113, 0
End Sub
Private Sub txtcodigoal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        txtcodigoal.Text = ""
        wtipoayuda = "A"
        Ayudas.Top = 3800
        Ayudas.Left = 6000
        Ayudas.Show 1
        txtcodigoal.Text = wcodigos
        Unload Ayudas
        cmdaceptar.SetFocus
    End If
End Sub
Private Sub txtcodigoal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcodigoal.Text) <> "" Then
        If VALIDA_ALMACEN(txtcodigoal.Text) = True Then
            TXTDESCRIPCION.Text = cnomalm
        Else
            MsgBox "Codigo de Almacen no existe", vbInformation + vbDefaultButton1, "Atencion"
            txtcodigoal.Text = "": txtcodigoal.SetFocus
        End If
    End If
    End If
End Sub
Public Sub Coneccion()
'Set cn = New ADODB.Connection

cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\bancowin\Obc02\db_bancos.mdb;Persist Security Info=False"
cn.Open
End Sub
Private Sub txtcodigoal_LostFocus()
    If Trim(txtcodigoal.Text) <> "" Then
        If VALIDA_ALMACEN(txtcodigoal.Text) = True Then
            TXTDESCRIPCION.Text = cnomalm
        Else
            MsgBox "Codigo de Almacen no existe", vbInformation + vbDefaultButton1, "Atencion"
            txtcodigoal.Text = "": txtcodigoal.SetFocus
        End If
    End If
End Sub
Public Function VALIDA_ALMACEN(pcodialm As String)
Dim sw1 As Boolean
Dim RST1 As ADODB.Recordset
    Set RST1 = New ADODB.Recordset

    sw1 = False
    If RST1.State Then RST1.Close
    RST1.Open "Select * from ef2almacenes where f2codalm='" & Trim(pcodialm) & "'", cnn_dbbancos
    If Not RST1.EOF Then
        cnomalm = Trim(RST1!F2NOMALM & "")
        sw1 = True
    Else
        sw1 = False
    End If
    RST1.Close
    VALIDA_ALMACEN = sw1
End Function

