VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form consulta_partedealmacen 
   Caption         =   "Parte de Almacén"
   ClientHeight    =   4140
   ClientLeft      =   2730
   ClientTop       =   2160
   ClientWidth     =   7350
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
   LockControls    =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   7350
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   7170
      Begin VB.TextBox txtalmacen 
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   2
         Top             =   2295
         Width           =   690
      End
      Begin VB.TextBox txtconcepto 
         Height          =   315
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2745
         Visible         =   0   'False
         Width           =   690
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   870
         Left            =   225
         TabIndex        =   7
         Top             =   315
         Width           =   6720
         _Version        =   65536
         _ExtentX        =   11853
         _ExtentY        =   1535
         _StockProps     =   14
         Caption         =   " Rango de Fechas"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         Begin aBoxCtl.aBox abohasta 
            Height          =   315
            Left            =   4845
            TabIndex        =   1
            Top             =   405
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            ABoxType        =   ""
            MinValue        =   "D01000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   1
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
            Text            =   "15/06/2004"
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
            ButtonPicture   =   "consulta_partedealmacen.frx":0000
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
         Begin aBoxCtl.aBox abodesde 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   405
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            ABoxType        =   ""
            MinValue        =   "D01000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   1
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
            Text            =   "15/06/2004"
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
            ButtonPicture   =   "consulta_partedealmacen.frx":0352
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   210
            Left            =   585
            TabIndex        =   9
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   210
            Left            =   4260
            TabIndex        =   8
            Top             =   450
            Width           =   420
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   825
         Left            =   225
         TabIndex        =   10
         Top             =   1260
         Width           =   6720
         _Version        =   65536
         _ExtentX        =   11853
         _ExtentY        =   1455
         _StockProps     =   14
         Caption         =   " Tipo "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   0
            Left            =   1080
            TabIndex        =   15
            Top             =   360
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "&General"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   1
            Left            =   4725
            TabIndex        =   16
            Top             =   360
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "&Detallado"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSPanel pnlconcepto 
         Height          =   330
         Left            =   1845
         TabIndex        =   11
         Top             =   2745
         Visible         =   0   'False
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
      Begin Threed.SSPanel pnlalmacen 
         Height          =   330
         Left            =   1845
         TabIndex        =   13
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   210
         Left            =   270
         TabIndex        =   14
         Top             =   2340
         Width           =   630
      End
      Begin VB.Label lblconcepto 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   210
         Left            =   270
         TabIndex        =   12
         Top             =   2790
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   510
      Left            =   3735
      TabIndex        =   5
      Top             =   3465
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   510
      Left            =   2295
      TabIndex        =   4
      Top             =   3465
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
Attribute VB_Name = "consulta_partedealmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sw_ayuda    As Boolean

Private Function VALIDA_CONCEPTO_INV(pconcepto As String)
Dim sw_e    As Boolean

    If rsconcepto_inv.State = adStateOpen Then rsconcepto_inv.Close
    rsconcepto_inv.Open "SELECT F1PRECIO,F1NOMORI,F1PARTIDA FROM SF1ORIGENES WHERE F1CODORI='" & pconcepto & "'", cnn_dbbancos
    If Not rsconcepto_inv.EOF Then
        wnomconcepto = Trim(rsconcepto_inv.Fields("F1NOMORI") & "")
        Rem NSE wpartida = Trim(rsconcepto_inv.Fields("F1PARTIDA") & "")
        Rem NSE wprecio = Trim(rsconcepto_inv.Fields("F1PRECIO") & "")
        sw_e = True
    Else
        sw_e = False
    End If
    rsconcepto_inv.Close
    VALIDA_CONCEPTO_INV = sw_e

End Function

Private Sub abodesde_GotFocus()

    abodesde.FocusSelect = True

End Sub

Private Sub abodesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        abohasta.SetFocus
    End If

End Sub

Private Sub abohasta_GotFocus()
    
    abohasta.FocusSelect = True
    
End Sub

Private Sub abohasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        opttipo(0).SetFocus
    End If
    
End Sub

Private Sub cmdaceptar_Click()
If Not IsDate(abodesde.Value) Then
    MsgBox "Debe Ingresar Fecha de Inicio de Rango"
    abodesde.SetFocus
    Exit Sub
End If
If Not IsDate(abohasta.Value) Then
    MsgBox "Debe Ingresar Fecha de Fin de Rango"
    abohasta.SetFocus
    Exit Sub
End If

Me.MousePointer = vbHourglass
PROCESA_CONSULTA
Me.MousePointer = vbDefault
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub PROCESA_CONSULTA()
Dim csql            As String
Dim calmacen        As String
Dim dfdesde         As Date
Dim dfhasta         As Date
Dim cwhere_almacen  As String
Dim cconcepto       As String
        
    calmacen = Trim(txtalmacen.Text)
    dfdesde = Format(abodesde.Value, "DD/MM/YYYY")
    dfhasta = Format(abohasta.Value, "DD/MM/YYYY")
    cconcepto = Trim(txtconcepto.Text)
    cwhere_almacen = ""
    If Len(calmacen) > 0 Then
        cwhere_almacen = " AND ((IF4VALES.F2CODALM) = '" & calmacen & "') "
    Else
        cwhere_almacen = ""
    End If
    
    '---------- CONSULTA GENERAL
    If opttipo(0).Value = True Then
        csql = "SELECT IF4VALES.F1CODORI AS CODIGO, SF1ORIGENES.F1NOMORI AS NOMBRE, " & _
               "Sum(IF3VALES.F3CANPRO*IF3VALES.F3VALVTA) AS TOTAL, First(SF1ORIGENES.F1TIPMOV) AS TIPMOV," & _
               "Sum(IF3VALES.F3CANPRO) AS CANTIDAD " & _
               "FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) " & _
               "INNER JOIN ((IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) " & _
               "INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) ON " & _
               "(IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) " & _
               "WHERE ((IF3VALES.F4FECVAL) >= cvdate('" & dfdesde & "') AND " & _
               "(IF3VALES.F4FECVAL) <= cvdate('" & dfhasta & "')) " & _
               cwhere_almacen & _
               "GROUP BY IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, SF1ORIGENES.F1TIPMOV " & _
               "ORDER BY First(SF1ORIGENES.F1TIPMOV),IF4VALES.F1CODORI;"
        
        With acr_partedealmacen_general
            .Caption = "Parte de Almacén"
            .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cnn_dbbancos & ""
            .lbltitulo2.Caption = "Del " & Format(abodesde.Value, "DD/MM/YYYY") & " Al " & Format(abohasta.Value, "DD/MM/YYYY")
            .DataControl1.Source = csql
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .lblempresa.Caption = wnomcia
            If Len(Trim(pnlalmacen.Caption)) = 0 Then
                .lbltitulo1.Caption = "(TODOS LOS ALMACENES)"
            Else
                .lbltitulo1.Caption = pnlalmacen.Caption
            End If
            .Show vbModal
        End With
    Else
        cad2 = ""
        If Trim(pnlconcepto.Caption) <> "" Then
            cad2 = " and IF4VALES.F1CODORI='" & Trim(txtconcepto.Text) & "'"
        End If
        
        csql = "SELECT First(SF1ORIGENES.F1NOMORI) AS ORIGEN, IF4VALES.F1CODORI, First(EF2MARCAS.F2DESMAR) AS MARCA, First(IF5PLA.F5CODFAB) AS CODFAB, IF3VALES.F5CODPRO, First(IF5PLA.F5NOMPRO) AS NOMPRO, Sum(IF3VALES.F3CANPRO) AS CANTIDAD, First(EF7MEDIDAS.F7NOMMED) AS UM, Sum(IF3VALES.F3CANPRO*IF3VALES.F3VALVTA) AS TOTAL, SF1ORIGENES.F1TIPMOV,First(IF4VALES.F4NUMVAL) AS NUMVAL, FIRST(IF4VALES.F4FECVAL) AS FECVAL " _
        & "FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (((IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) " _
        & "WHERE ((IF3VALES.F4FECVAL) >= cvdate('" & dfdesde & "') AND (IF3VALES.F4FECVAL) <= cvdate('" & dfhasta & "')) " & _
        cwhere_almacen & cad2 _
        & "GROUP BY SF1ORIGENES.F1TIPMOV,IF4VALES.F1CODORI, IF3VALES.F5CODPRO " _
        & "ORDER BY F1TIPMOV,IF4VALES.F1CODORI,FIRST(IF4VALES.F4FECVAL), First(IF4VALES.F4NUMVAL),First(IF5PLA.F5CODFAB)"
        
        With acr_partedealmacen_detallado
            .Caption = "Parte de Almacén"
            .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cnn_dbbancos & ""
            .lbltitulo2.Caption = "Del " & Format(abodesde.Value, "DD/MM/YYYY") & " Al " & Format(abohasta.Value, "DD/MM/YYYY")
            .DataControl1.Source = csql
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .lblempresa.Caption = wnomcia
            If Len(Trim(pnlalmacen.Caption)) = 0 Then
                .lbltitulo1.Caption = "(TODOS LOS ALMACENES)"
            Else
                .lbltitulo1.Caption = pnlalmacen.Caption
            End If
            .Show vbModal
        End With
    End If

End Sub

Private Sub Form_Load()

    abodesde.Value = Format(Date, "DD/MM/YYYY")
    abohasta.Value = Format(Date, "DD/MM/YYYY")

End Sub



Private Sub opttipo_Click(Index As Integer, Value As Integer)
    If opttipo(0).Value = True Then
        lblconcepto.Visible = False
        txtconcepto.Visible = False
        pnlconcepto.Visible = False
    Else
        lblconcepto.Visible = True
        txtconcepto.Visible = True
        pnlconcepto.Visible = True
    End If

End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtalmacen.SetFocus
    End If

End Sub

Private Sub txtalmacen_DblClick()

    txtalmacen_KeyDown 113, 0
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtalmacen.SelStart = 0: txtalmacen.SelLength = Len(txtalmacen.Text)

End Sub

Private Sub txtalmacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        hlp_almacenes.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            txtalmacen.Text = wcod_alm
            pnlalmacen.Caption = wnomalmacen
            txtalmacen_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtalmacen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtconcepto.Visible = True Then
            txtconcepto.SetFocus
        Else
            cmdaceptar.SetFocus
        End If
    End If

End Sub

Private Sub txtalmacen_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            If VALIDA_ALMACEN(txtalmacen.Text) = True Then
                pnlalmacen.Caption = wnomalmacen
            Else
                MsgBox "El código del almacén no existe. Verifique.", vbCritical, "Atención"
                txtalmacen.SetFocus
            End If
        End If
    End If

End Sub

Private Sub txtconcepto_Change()
pnlconcepto.Caption = ""
End Sub

Private Sub txtconcepto_DblClick()

    txtconcepto_KeyDown 113, 0
    
End Sub

Private Sub txtconcepto_GotFocus()

    txtconcepto.SelStart = 0: txtconcepto.SelLength = Len(txtconcepto.Text)
    
End Sub

Private Sub txtconcepto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wconcepto = ""
        wtipmov = "": wnomconcepto = ""
        hlp_conceptos_inv.Show 1
        sw_ayuda = False
        If Len(Trim(wconcepto)) > 0 Then
            txtconcepto.Text = wconcepto
            pnlconcepto.Caption = wnomconcepto
            txtconcepto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtconcepto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub

Private Sub txtconcepto_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            If Len(Trim(txtconcepto.Text)) > 0 Then
                If VALIDA_CONCEPTO_INV(txtconcepto.Text) = True Then
                    pnlconcepto.Caption = wnomconcepto
                Else
                    MsgBox "Código de concepto no existe. Verifique", vbCritical, "Atención"
                    txtconcepto.SetFocus
                End If
            Else
                MsgBox "Falta ingresar el código del concepto. Verifique", vbCritical, "Atención"
                txtconcepto.SetFocus
            End If
        End If
    End If
    
End Sub
