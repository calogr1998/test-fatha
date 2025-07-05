VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form RInventario 
   Caption         =   "Reporte de Inventario Por Tipo de Operación"
   ClientHeight    =   2745
   ClientLeft      =   2505
   ClientTop       =   2130
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   7740
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "RInventario.frx":0000
      ToolBars        =   "RInventario.frx":1985
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7455
      Begin VB.ComboBox CboConcepto 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox CboMov 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   855
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _Version        =   65536
         _ExtentX        =   11880
         _ExtentY        =   1508
         _StockProps     =   14
         Caption         =   " Rango de fechas "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         ShadowStyle     =   1
         Begin aBoxCtl.aBox txtdesde 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
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
            Text            =   "10/12/2002"
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
            ButtonPicture   =   "RInventario.frx":1A29
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
         Begin aBoxCtl.aBox txthasta 
            Height          =   315
            Left            =   4680
            TabIndex        =   3
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
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
            Text            =   "10/12/2002"
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
            ButtonPicture   =   "RInventario.frx":1D7B
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
         Begin VB.Label lblfecemi 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   720
            TabIndex        =   5
            Top             =   405
            Width           =   465
         End
         Begin VB.Label lblfecven 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3840
            TabIndex        =   4
            Top             =   405
            Width           =   420
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   810
      End
   End
End
Attribute VB_Name = "RInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsConcepto As New ADODB.Recordset

Private Sub CboMov_Click()

    CboConcepto.Clear
    If RsConcepto.State = adStateOpen Then RsConcepto.Close
    RsConcepto.Open "SELECT F1CODORI,F1NOMORI,F1TIPMOV FROM SF1ORIGENES WHERE F1TIPMOV = '" & Left(CboMov.Text, 1) & "'  ORDER BY F1NOMORI ASC", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsConcepto.EOF Then
        RsConcepto.MoveFirst
        Do While Not RsConcepto.EOF
            CboConcepto.AddItem RsConcepto.Fields("F1NOMORI") & Space(50) & RsConcepto.Fields("F1CODORI")
            RsConcepto.MoveNext
        Loop
    RsConcepto.Close
    End If
    CboConcepto.ListIndex = -1

End Sub

Private Sub Form_Load()
Me.Height = 3150
Me.Left = 2445
Me.Top = 1785
Me.Width = 7860

CboMov.AddItem "INGRESO"
CboMov.AddItem "SALIDA"
CboMov.ListIndex = -1

If RsConcepto.State = adStateOpen Then RsConcepto.Close
RsConcepto.Open "SELECT F1CODORI,F1NOMORI FROM SF1ORIGENES ORDER BY F1NOMORI ASC", cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsConcepto.EOF Then
    RsConcepto.MoveFirst
    Do While Not RsConcepto.EOF
        CboConcepto.AddItem RsConcepto.Fields("F1NOMORI") & Space(50) & RsConcepto.Fields("F1CODORI")
        RsConcepto.MoveNext
    Loop
    RsConcepto.Close
End If
CboConcepto.ListIndex = -1

txtdesde.Value = Format(Date, "DD/MM/YYYY")
txthasta.Value = Format(Date, "DD/MM/YYYY")

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Imprimir"
        Imprimir
    Case "ID_Salir"
        Unload Me
End Select
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txthasta.SetFocus
End Sub

Private Sub txtdesde_LostFocus()
If Not IsDate(txtdesde.Value) Then
    MsgBox "Error en la Fecha. Verifique", vbExclamation, "AVISO"
    txtdesde.Value = Format(Date, "DD/MM/YYYY")
    txtdesde.SetFocus
    Exit Sub
End If
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboMov.SetFocus
End Sub

Private Sub txthasta_LostFocus()
If Not IsDate(txthasta.Value) Then
    MsgBox "Error en la Fecha. Verifique", vbExclamation, "AVISO"
    txthasta.Value = Format(Date, "DD/MM/YYYY")
    txthasta.SetFocus
    Exit Sub
End If
End Sub

Private Sub Imprimir()
    If CboMov.ListIndex = -1 Or CboConcepto.ListIndex = -1 Then
        MsgBox "Ingrese Tipo de Operación", vbExclamation, "AVISO"
        Exit Sub
    End If

    With Acr_Operaciones
        .DataControl1.ConnectionString = cnn_dbbancos
        SQL = "SELECT A.F4NUMVAL, A.F2CODALM, A.F4FECVAL, B.F5CODPRO,C.F5NOMPRO, B.F3CANPRO, [B]![F3TOTITE]/[B]![F3CANPRO] AS PREUNI, B.F3TOTITE " & _
              " FROM (IF3VALES AS B INNER JOIN IF4VALES AS A ON (B.F2CODALM = A.F2CODALM) AND (B.F4NUMVAL = A.F4NUMVAL)) INNER JOIN IF5PLA AS C ON B.F5CODPRO = C.F5CODPRO WHERE (LEFT(A.F4NUMVAL,1) = '" & Left(CboMov.Text, 1) & "' " & _
              " AND A.F1CODORI = '" & Right(CboConcepto.Text, 3) & "' AND CVDATE(A.F4FECVAL)>= '" & CVDate(txtdesde.Value) & "' AND CVDATE(A.F4FECVAL) <= '" & CVDate(txthasta.Value) & "') ORDER BY A.F4NUMVAL, A.F2CODALM;"
        .DataControl1.Source = SQL
        .fldtitulo = "DEL " & txtdesde.Value & Space(3) & " HASTA EL " & txthasta.Value
        If rsconsulta.State = adStateOpen Then rsconsulta.Close
        rsconsulta.Open "Select F1CODORI,F1NOMORI FROM SF1ORIGENES", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconsulta.EOF Then
            .lbltitulo.Caption = "" & rsconsulta.Fields("F1CODORI") & "-" & rsconsulta.Fields("F1NOMORI")
        Else
            .lbltitulo.Caption = "INVENTARIO X TIPO DE OPERACIÓN"
        End If
        rsconsulta.Close
        .fldempresa.Text = wnomcia
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .Show vbModal
    End With

End Sub
