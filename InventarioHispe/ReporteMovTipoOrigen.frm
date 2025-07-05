VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form ReporteMovTipoOrigen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reporte por tipo de Origen :::"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "ReporteMovTipoOrigen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   75
      TabIndex        =   5
      Top             =   675
      Width           =   5100
      Begin VB.TextBox TxtAlmacen 
         Height          =   330
         Left            =   1125
         MaxLength       =   3
         TabIndex        =   6
         Top             =   210
         Width           =   495
      End
      Begin Threed.SSPanel PnlAlmacen 
         Height          =   330
         Left            =   1620
         TabIndex        =   7
         Top             =   210
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen :"
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
         Left            =   345
         TabIndex        =   8
         Top             =   240
         Width           =   705
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   372
      Left            =   3852
      TabIndex        =   4
      Top             =   2820
      Width           =   1308
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Procesar..."
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   5100
      Begin VB.TextBox Txtcodori 
         Height          =   330
         Left            =   1125
         MaxLength       =   3
         TabIndex        =   2
         Top             =   210
         Width           =   495
      End
      Begin Threed.SSPanel PnlNomOri 
         Height          =   330
         Left            =   1620
         TabIndex        =   3
         Top             =   210
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conceptos :"
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
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   870
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   72
      TabIndex        =   9
      Top             =   1992
      Width           =   5100
      _Version        =   65536
      _ExtentX        =   8996
      _ExtentY        =   1164
      _StockProps     =   14
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
      Alignment       =   2
      Font3D          =   3
      Begin aBoxCtl.aBox aboDesde 
         Height          =   315
         Left            =   1125
         TabIndex        =   10
         Top             =   210
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
         Text            =   "24/07/2007"
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
         ButtonPicture   =   "ReporteMovTipoOrigen.frx":000C
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
      Begin aBoxCtl.aBox aboHasta 
         Height          =   315
         Left            =   3285
         TabIndex        =   12
         Top             =   210
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
         Text            =   "24/07/2007"
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
         ButtonPicture   =   "ReporteMovTipoOrigen.frx":035E
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
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Top             =   240
         Width           =   555
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   600
      Left            =   75
      TabIndex        =   14
      Top             =   1368
      Width           =   5100
      _Version        =   65536
      _ExtentX        =   8996
      _ExtentY        =   1058
      _StockProps     =   14
      Caption         =   "Tipo de Reporte"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   3
      Begin Threed.SSOption Opval 
         Height          =   195
         Left            =   3555
         TabIndex        =   15
         Top             =   285
         Width           =   1050
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Valorizado"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption Opnval 
         Height          =   192
         Left            =   228
         TabIndex        =   16
         Top             =   288
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   339
         _StockProps     =   78
         Caption         =   "No valorizado"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
End
Attribute VB_Name = "ReporteMovTipoOrigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
AboDesde.Value = Format(Date, "DD/MM/YYYY")
AboHasta.Value = Format(Date, "DD/MM/YYYY")
Call txtalmacen_LostFocus
Call Txtcodori_Change
End Sub

Private Sub SSCommand1_Click()
Dim StrSql As String
    If Txtcodori.Text = "" And txtalmacen.Text = "" Then
        StrSql = "SELECT IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4NUMVAL, " & _
        " IF4VALES.F4TIPOVALE, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, " & _
        " IF3VALES.F3CANPRO, IF5PLA.F7CODMED, IF3VALES.F3VALVTA,IF3VALES.f2codalm, [F3CANPRO]*[F3VALVTA] AS Importe, " & _
        " EF2PROVEEDORES.F2NOMPROV FROM ((IF4VALES INNER JOIN SF1ORIGENES ON " & _
        " IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES INNER " & _
        " JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL " & _
        " = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) INNER JOIN " & _
        " EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC WHERE " & _
        " IF4VALES.F4FECVAL Between #" & Format(AboDesde.Value, "MM/DD/YYYY") & "# And #" & Format(AboHasta.Value, "MM/DD/YYYY") & "# ORDER BY IF4VALES.F4FECVAL;"
    ElseIf Txtcodori.Text = "" And txtalmacen.Text <> "" Then
        StrSql = "SELECT IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4NUMVAL, " & _
        " IF4VALES.F4TIPOVALE, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, " & _
        " IF3VALES.F3CANPRO, IF5PLA.F7CODMED, IF3VALES.F3VALVTA, [F3CANPRO]*[F3VALVTA] AS Importe, " & _
        " EF2PROVEEDORES.F2NOMPROV FROM ((IF4VALES INNER JOIN SF1ORIGENES ON " & _
        " IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES INNER " & _
        " JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL " & _
        " = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) INNER JOIN " & _
        " EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC" _
        & " WHERE IF3VALES.F2CODALM = '" & txtalmacen.Text & _
        "' And IF4VALES.F4FECVAL Between #" & Format(AboDesde.Value, "MM/DD/YYYY") & "# And #" & Format(AboHasta.Value, "MM/DD/YYYY") & "#  ORDER BY IF4VALES.F4FECVAL;"
    ElseIf Txtcodori.Text <> "" And txtalmacen.Text = "" Then
        StrSql = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F2CODALM & IF4VALES.F4NUMVAL as UNIDO, IF4VALES.F4FECVAL, IF4VALES.F2CODPROV, EF2PROVEEDORES.F2NOMPROV, " & _
                    "IF4VALES.F4TIPDOC, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F1CODORI, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, " & _
                    "IF5PLA.F7CODMED, IF3VALES.F3CANPRO, IF3VALES.F3VALVTA AS SOLES, IF3VALES.F3VALDOL,[F3CANPRO]*[F3VALVTA] AS Importe,[F3CANPRO]*[F3VALDOL] AS ImporteD " & _
                    "FROM (IF4VALES LEFT JOIN EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC) INNER JOIN (IF3VALES INNER " & _
                    "JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = " & _
                    "IF3VALES.F2CODALM) " & _
                    " WHERE IF4VALES.F1CODORI = '" & Txtcodori.Text & "' And IF4VALES.F4FECVAL Between #" & Format(AboDesde.Value, "MM/DD/YYYY") & "# And #" & Format(AboHasta.Value, "MM/DD/YYYY") & "#  ORDER BY IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL;"
    ElseIf Txtcodori.Text <> "" And txtalmacen.Text <> "" Then
        StrSql = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F2CODALM & IF4VALES.F4NUMVAL as UNIDO, IF4VALES.F4FECVAL, IF4VALES.F2CODPROV, EF2PROVEEDORES.F2NOMPROV, " & _
                    "IF4VALES.F4TIPDOC, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F1CODORI, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, " & _
                    "IF5PLA.F7CODMED, IF3VALES.F3CANPRO, IF3VALES.F3VALVTA AS SOLES, IF3VALES.F3VALDOL,[F3CANPRO]*[F3VALVTA] AS Importe,[F3CANPRO]*[F3VALDOL] AS ImporteD " & _
                    "FROM (IF4VALES LEFT JOIN EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC) INNER JOIN (IF3VALES INNER " & _
                    "JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = " & _
                    "IF3VALES.F2CODALM) " & _
                    " WHERE IF4VALES.F1CODORI = '" & Txtcodori.Text & "' AND IF3VALES.F2CODALM = '" & txtalmacen.Text & "'And IF4VALES.F4FECVAL Between #" & Format(AboDesde.Value, "MM/DD/YYYY") & "# And #" & Format(AboHasta.Value, "MM/DD/YYYY") & "#  ORDER BY IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL;"
    End If
    
Screen.MousePointer = 11
    If Txtcodori.Text <> "" Then
         'acr_Mov_TipoOrigen_detallado.Field11.Visible = Opval.Value
       ' acr_Mov_TipoOrigen_detallado.Field13.Visible = Opval.Value
       ' acr_Mov_TipoOrigen_detallado.Field14.Visible = Opval.Value
       ' acr_Mov_TipoOrigen_detallado.Field15.Visible = Opval.Value
        'acr_Mov_TipoOrigen_detallado.Field20.Visible = Opval.Value
        'acr_Mov_TipoOrigen_detallado.Label8.Visible = Opval.Value
       ' acr_Mov_TipoOrigen_detallado.Label17.Visible = Opval.Value
        acr_Mov_TipoOrigen_detallado.Labelmov.Caption = acr_Mov_TipoOrigen.Labelmov.Caption & " " & PnlNomOri.Caption
        acr_Mov_TipoOrigen_detallado.lblempresa.Caption = wempresa
        acr_Mov_TipoOrigen_detallado.fldfecha.Text = Format(Date, "DD/MM/YYYY")
        acr_Mov_TipoOrigen_detallado.datos.ConnectionString = cnn_dbbancos
        acr_Mov_TipoOrigen_detallado.datos.Source = StrSql
        acr_Mov_TipoOrigen_detallado.Show 1
    Else
        acr_Mov_TipoOrigen.Field11.Visible = Opval.Value
        acr_Mov_TipoOrigen.Field13.Visible = Opval.Value
        acr_Mov_TipoOrigen.Field14.Visible = Opval.Value
        acr_Mov_TipoOrigen.Field15.Visible = Opval.Value
        acr_Mov_TipoOrigen.Field20.Visible = Not Opval.Value
        acr_Mov_TipoOrigen.Label8.Visible = Opval.Value
        acr_Mov_TipoOrigen.Label17.Visible = Opval.Value
        acr_Mov_TipoOrigen.Labelmov.Caption = acr_Mov_TipoOrigen.Labelmov.Caption & " " & PnlNomOri.Caption
        acr_Mov_TipoOrigen.lblempresa.Caption = wempresa
        acr_Mov_TipoOrigen.fldfecha.Text = Format(Date, "DD/MM/YYYY")
        acr_Mov_TipoOrigen.datos.ConnectionString = cnn_dbbancos
        acr_Mov_TipoOrigen.datos.Source = StrSql
        acr_Mov_TipoOrigen.Show 1

    End If


End Sub

Private Sub txtalmacen_Change()
    
    pnlalmacen.Caption = ""
    
End Sub

Private Sub txtalmacen_DblClick()

    txtalmacen_KeyDown 113, 0
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtalmacen.SelStart = 0
    txtalmacen.SelLength = Len(txtalmacen.Text)

End Sub

Private Sub txtalmacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            txtalmacen.Text = wcod_alm
            pnlalmacen.Caption = wnomalmacen
            
        End If
    End If
    
End Sub

Private Sub txtalmacen_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            wnomalmacen = ""
            If VALIDA_ALMACEN(txtalmacen.Text) = True Then
                pnlalmacen.Caption = wnomalmacen
            Else
                MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                txtalmacen.SetFocus
            End If
        Else
            pnlalmacen.Caption = "TODOS LOS ALMACENES"
        End If
    End If

End Sub

Private Sub Txtcodori_Change()

If Txtcodori.Text = "" Then PnlNomOri.Caption = "TODOS LOS MOVIMIENTOS"

End Sub

Private Sub Txtcodori_DblClick()
    
    Txtcodori_KeyDown 113, 0
    
End Sub

Private Sub Txtcodori_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        sw_ayudaO = True
        wcodori = ""
        wtipmov = ""
        ayuda_conceptos.Show 1
        sw_ayudaO = False
        If Len(Trim(wconcepto)) > 0 Then
            Txtcodori.Text = wconcepto
            PnlNomOri.Caption = wnomconcepto
            
        End If
    End If
   
End Sub

