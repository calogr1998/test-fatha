VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form cons_consolidadoxconcepto 
   Caption         =   "Movimiento Consolidado por Concepto"
   ClientHeight    =   6540
   ClientLeft      =   1515
   ClientTop       =   2385
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10560
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Width           =   10290
      Begin Threed.SSPanel pnlconcepto 
         Height          =   360
         Left            =   2010
         TabIndex        =   9
         Top             =   255
         Width           =   8130
         _Version        =   65536
         _ExtentX        =   14340
         _ExtentY        =   635
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.TextBox txtconcepto 
         Height          =   330
         Left            =   1095
         TabIndex        =   8
         Top             =   255
         Width           =   825
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   645
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   9990
         _Version        =   65536
         _ExtentX        =   17621
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Rango de Fechas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin aBoxCtl.aBox abodesde 
            Height          =   315
            Left            =   2580
            TabIndex        =   3
            Top             =   180
            Width           =   1245
            _ExtentX        =   2196
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
            ButtonPicture   =   "cons_consolidadoxconcepto.frx":0000
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
         Begin aBoxCtl.aBox abohasta 
            Height          =   315
            Left            =   6945
            TabIndex        =   4
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
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
            ButtonPicture   =   "cons_consolidadoxconcepto.frx":0352
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   6165
            TabIndex        =   6
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   195
            Left            =   1725
            TabIndex        =   5
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   690
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   315
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "cons_consolidadoxconcepto.frx":06A4
      ToolBars        =   "cons_consolidadoxconcepto.frx":5F40
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4380
      Left            =   105
      OleObjectBlob   =   "cons_consolidadoxconcepto.frx":6032
      TabIndex        =   0
      Top             =   1695
      Width           =   10290
   End
End
Attribute VB_Name = "cons_consolidadoxconcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL     As String

Private Sub Imprimir()
'''Dim Reg As New ADODB.Recordset
'''
'''    With Acr_ResumenVentas
'''        .DataControl1.ConnectionString = cnn_dbbancos
'''
'''        If SOFab.Value = True Then
'''            SQL = "SELECT DISTINCTROW IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, Sum(TBVENTA_DET.F3CANPRO) AS CANTIDAD, Sum([TBVENTA_DET]![F3PREVTA]-[TBVENTA_DET]![F3IGV]) AS VVTA, Sum(TBVENTA_DET.F3IGV) AS IGV, Sum(TBVENTA_DET.F3PREVTA) AS PRECIO FROM IF5PLA INNER JOIN (TBVENTA_DET INNER JOIN TBVENTA_CAB ON (TBVENTA_DET.F4SERDOC = TBVENTA_CAB.F4SERDOC) AND (TBVENTA_DET.F4NUMDOC = TBVENTA_CAB.F4NUMDOC)) ON IF5PLA.F5CODPRO = TBVENTA_DET.F5CODPRO " & _
'''            "WHERE (((TBVENTA_CAB.F4ESTNUL)='N') And (TBVENTA_CAB.F4FECEMI >= CVDate( '" & abodesde.Value & "') And TBVENTA_CAB.F4FECEMI <= CVDate( '" & abohasta.Value & "'))) GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO ORDER BY IF5PLA.F5CODPRO;"
'''        Else
'''            SQL = "SELECT DISTINCTROW IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, Sum(TBVENTA_DET.F3CANPRO) AS CANTIDAD, Sum([TBVENTA_DET]![F3PREVTA]-[TBVENTA_DET]![F3IGV]) AS VVTA, Sum(TBVENTA_DET.F3IGV) AS IGV, Sum(TBVENTA_DET.F3PREVTA) AS PRECIO FROM IF5PLA INNER JOIN (TBVENTA_DET INNER JOIN TBVENTA_CAB ON (TBVENTA_DET.F4SERDOC = TBVENTA_CAB.F4SERDOC) AND (TBVENTA_DET.F4NUMDOC = TBVENTA_CAB.F4NUMDOC)) ON IF5PLA.F5CODPRO = TBVENTA_DET.F5CODPRO " & _
'''            "WHERE (((TBVENTA_CAB.F4ESTNUL)='N') And (TBVENTA_CAB.F4FECEMI >= CVDate( '" & abodesde.Value & "') And TBVENTA_CAB.F4FECEMI <= CVDate( '" & abohasta.Value & "'))) GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO ORDER BY 6 DESC;"
'''
'''         End If
'''        .DataControl1.Source = SQL
'''        If Reg.State = adStateOpen Then Reg.Close
'''        Reg.Open "SELECT * FROM SF1PARAM WHERE F1CODEMP = '" & wempresa & "'", cnn_control, adOpenKeyset, adLockOptimistic
'''        If Not Reg.EOF Then
'''           .FldEmpresa.Text = Reg.Fields("F1NOMEMP")
'''           Reg.Close
'''        Else
'''           .FldEmpresa.Text = wempresa
'''           Reg.Close
'''        End If
'''        .fldfecha.Text = Format(Now, "dd/mm/yyyy")
'''        .fldtitulo.Text = "Del  " & Format(abodesde.Value, "dd/mm/yyyy") & "  al  " & Format(abohasta.Value, "dd/mm/yyyy")
'''        .Label10.Caption = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").SummaryFooterValue
'''        .Label11.Caption = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTA").SummaryFooterValue, "###,##0.00")
'''        .Label12.Caption = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "###,##0.00")
'''        .Label13.Caption = Format(dxDBGrid1.Columns.ColumnByFieldName("PRECIO").SummaryFooterValue, "###,##0.00")
'''        .Show vbModal
'''    End With

End Sub

Private Sub FILL()

    SQL = "SELECT DISTINCTROW IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, Sum(IF3VALES.F3CANPRO) AS CANTIDAD, Sum([TBVENTA_DET]![F3PREVTA]-[TBVENTA_DET]![F3IGV]) AS VVTA, Sum(TBVENTA_DET.F3IGV) AS IGV, Sum(TBVENTA_DET.F3PREVTA) AS PRECIO FROM IF5PLA INNER JOIN (TBVENTA_DET INNER JOIN TBVENTA_CAB ON (TBVENTA_DET.F4SERDOC = TBVENTA_CAB.F4SERDOC) AND (TBVENTA_DET.F4NUMDOC = TBVENTA_CAB.F4NUMDOC)) ON IF5PLA.F5CODPRO = TBVENTA_DET.F5CODPRO " & _
          "WHERE (((TBVENTA_CAB.F4ESTNUL)='N') And (TBVENTA_CAB.F4FECEMI >= CVDate( '" & abodesde.Value & "') And TBVENTA_CAB.F4FECEMI <= CVDate( '" & abohasta.Value & "'))) GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO ORDER BY IF5PLA.F5CODPRO;"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = SQL
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F5CODPRO"
    CABECERA
    dxDBGrid1.Top = 1800
    dxDBGrid1.Height = 4725

End Sub

Private Sub CONECTAR()
    
    With dxDBGrid1
         .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With

End Sub

Private Sub CABECERA()
    
    With dxDBGrid1
         .Columns(0).Caption = "Codigo": .Columns(0).Width = 50
         .Columns(1).Caption = "Producto": .Columns(1).Width = 145
         .Columns(2).Caption = "Cantidad": .Columns(2).Width = 30: .Columns(2).Alignment = taRightJustify
         .Columns(3).Caption = "V.Venta": .Columns(3).Width = 40: .Columns(3).DecimalPlaces = 2: .Columns(3).Alignment = taRightJustify
         .Columns(4).Caption = "IGV": .Columns(4).Width = 40: .Columns(4).DecimalPlaces = 2: .Columns(4).Alignment = taRightJustify
         .Columns(5).Caption = "Precio": .Columns(5).Width = 40: .Columns(5).DecimalPlaces = 2: .Columns(5).Alignment = taRightJustify
    End With

End Sub

Private Sub Form_Load()

    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    
    abodesde.Value = Format(Now, "dd/mm/yyyy")
    abohasta.Value = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Aceptar"
            Me.MousePointer = 11
            CONECTAR
            FILL
            Me.MousePointer = 1
        Case "ID_Imprimir"
            Me.MousePointer = 11
            Imprimir
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    End Select

End Sub

Private Sub txtconcepto_DblClick()

    txtconcepto_KeyDown 113, 0
    
End Sub

Private Sub txtconcepto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wconcepto = ""
        wtipmov = "S": wnomconcepto = ""
        hlp_conceptos_inv.Show 1
        If Len(Trim(wconcepto)) > 0 Then
            txtconcepto.Text = wconcepto
            pnlconcepto.Caption = wnomconcepto
            txtconcepto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtconcepto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        abodesde.SetFocus
    End If

End Sub
