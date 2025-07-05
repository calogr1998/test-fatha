VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Ventas_Vales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Consistencia entre Ventas y Vales :::"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "Ventas_Vales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Reporte"
      Height          =   840
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   6150
      Begin VB.OptionButton Option3 
         Caption         =   "Con Vale"
         Height          =   255
         Left            =   4215
         TabIndex        =   9
         Top             =   360
         Width           =   1005
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sin Vale"
         Height          =   255
         Left            =   2550
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Left            =   1020
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fecha"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   990
      Width           =   6180
      Begin Threed.SSCommand SSCommand1 
         Height          =   345
         Left            =   4860
         TabIndex        =   5
         Top             =   345
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "&Procesar..."
      End
      Begin aBoxCtl.aBox ABODESDE 
         Height          =   315
         Left            =   945
         TabIndex        =   1
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D10000101"
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
         FocusSelect     =   -1  'True
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
         ButtonPicture   =   "Ventas_Vales.frx":000C
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
         Left            =   3300
         TabIndex        =   3
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D10000101"
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
         FocusSelect     =   -1  'True
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
         ButtonPicture   =   "Ventas_Vales.frx":035E
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
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   390
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   285
         TabIndex        =   2
         Top             =   390
         Width           =   555
      End
   End
End
Attribute VB_Name = "Ventas_Vales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()
Dim sql As String

Dim rt As New ADODB.Recordset

If rs.State = 1 Then rs.Close
rs.Open "Select F5CODPRO, F5NOMPRO, F7CODMED, F3CANPRO, F4TIPODOCU, F4SERDOC, F4NUMDOC, F4FECEMI, F4NUMVAL, " & _
        "F2CODALM FROM TBVENTA_DET WHERE F4FECEMI BETWEEN #" & Format(CDate(ABODESDE.Value), "dd/mm/yyyy") _
        & "# AND #" & Format(CDate(ABOHASTA.Value), "dd/mm/yyyy") & "# And F4TIPODOCU = '86' And F4ESTNUL = 'N'", cnn_dbbancos, adOpenStatic, adLockReadOnly

Dim tipodoc As String, codpro As String, fecha As String

cnn_form.Execute "DELETE FROM Venta_Vales"

Do While rs.EOF = False

    tipodoc = rs("F4TIPODOCU")
    codpro = rs("F5CODPRO")
    fecha = rs("F4FECEMI")
    
    If rt.State = 1 Then rt.Close
    rt.Open "Select * From IF4VALES Where F4NUMVAL = '" & rs("F4NUMVAL") & "' AND F2CODALM = '" & rs("F2CODALM") & _
    "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    
    If rt.EOF = False Then
        If Option1.Value = True Or Option3.Value = True Then
            sql = "Insert Into Venta_Vales Values('" & tipodoc & "', '" & rs("F4SERDOC") & "', '" & rs("F4NUMDOC") & _
            "', '" & codpro & "', '" & rs("F5NOMPRO") & "', " & rs("F3CANPRO") & ", '" & rt("F4NUMVAL") & "', " & _
            rs("F3CANPRO") & ", '" & rs("F4FECEMI") & "', '" & rt("F4FECVAL") & "')"
        End If
    Else
        If Option2.Value = True Or Option1.Value = True Then
            sql = "Insert Into Venta_Vales Values('" & tipodoc & "', '" & rs("F4SERDOC") & "', '" & rs("F4NUMDOC") & _
            "', '" & codpro & "' , '" & rs("F5NOMPRO") & "', " & rs("F3CANPRO") & ", 'NO HAY VALE', 0, '" & _
            rs("F4FECEMI") & "', '01/01/1900')"
        End If
    End If
    If sql <> "" Then cnn_form.Execute sql
    sql = ""
    
    rs.MoveNext
Loop
acr_ventas_Vales.datos.ConnectionString = cnn_form
acr_ventas_Vales.Label25.Caption = "Entre: " & ABODESDE.Value & " Hasta: " & ABODESDE.Value
acr_ventas_Vales.datos.Source = "Select * From Venta_Vales"
acr_ventas_Vales.Show 1
End Sub
