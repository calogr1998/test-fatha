VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F7E69521-3C28-11D2-B3E7-00AA00B42B7C}#3.1#0"; "fpTab30.ocx"
Begin VB.Form Pedido_Sugerido 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7140
   ClientLeft      =   -3285
   ClientTop       =   1830
   ClientWidth     =   11775
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11775
   Begin TabproADOLib.fpTabProADO fpTabProADO1 
      Height          =   6045
      Left            =   90
      TabIndex        =   5
      Top             =   915
      Width           =   11670
      _Version        =   196609
      _ExtentX        =   20585
      _ExtentY        =   10663
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   2
      TabCount        =   2
      AlignTextH      =   1
      AlignTextV      =   1
      ThreeD          =   0   'False
      TabShape        =   3
      ApplyTo         =   2
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      DrawFocusRect   =   2
      DataField       =   ""
      DataMember      =   ""
      TabCaption      =   "Pedido_Sugerido.frx":0000
      PageEarMarkPictureNext=   "Pedido_Sugerido.frx":02D3
      PageEarMarkPicturePrev=   "Pedido_Sugerido.frx":02EF
      EarMarkPictureNext=   "Pedido_Sugerido.frx":030B
      EarMarkPicturePrev=   "Pedido_Sugerido.frx":0327
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Todos"
         Height          =   195
         Left            =   8685
         TabIndex        =   12
         Top             =   450
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Llenar Requerimiento"
         Height          =   330
         Left            =   6660
         TabIndex        =   11
         Top             =   405
         Width           =   1815
      End
      Begin VB.CheckBox CheckFiltro3 
         Caption         =   "Activar Filtro"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -16679
         TabIndex        =   9
         Top             =   -15689
         Width           =   1455
      End
      Begin VB.CheckBox Checkagrupar 
         Caption         =   "Agrupar columnas"
         Height          =   255
         Left            =   1710
         TabIndex        =   7
         Top             =   435
         Width           =   2055
      End
      Begin VB.CheckBox CheckFiltro 
         Caption         =   "Activar Filtro"
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   435
         Width           =   1455
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5025
         Left            =   60
         OleObjectBlob   =   "Pedido_Sugerido.frx":0343
         TabIndex        =   8
         Top             =   795
         Width           =   11550
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid3 
         Height          =   5025
         Left            =   -26609
         OleObjectBlob   =   "Pedido_Sugerido.frx":3F19
         TabIndex        =   10
         Top             =   -20819
         Width           =   11550
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "Pedido_Sugerido.frx":4B75
      ToolBars        =   "Pedido_Sugerido.frx":CA65
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
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
         Left            =   4050
         TabIndex        =   1
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
         Text            =   "13/06/2012"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         ThousandSeparator=   "."
         DecimalSeparator=   ","
         CurrencySymbol  =   "$"
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
         ButtonPicture   =   "Pedido_Sugerido.frx":CC08
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
         Left            =   7425
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
         Text            =   "13/06/2012"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         ThousandSeparator=   "."
         DecimalSeparator=   ","
         CurrencySymbol  =   "$"
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
         ButtonPicture   =   "Pedido_Sugerido.frx":CF5A
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
      Begin VB.Label lblfecven 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   6585
         TabIndex        =   4
         Top             =   405
         Width           =   420
      End
      Begin VB.Label lblfecemi 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   3225
         TabIndex        =   3
         Top             =   405
         Width           =   465
      End
   End
End
Attribute VB_Name = "Pedido_Sugerido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dias As Integer
'Private Opc As Boolean

Private Sub Check1_Click()
    dxDBGrid1.Dataset.Close
    
    dxDBGrid1.Dataset.Active = False
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.KeyField = "ITEM"

If Check1.Value = 0 Then
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "PEDSUGE <> 0 "
Else
    dxDBGrid1.Dataset.Filtered = False
    dxDBGrid1.Dataset.Filter = ""
End If
dxDBGrid1.Dataset.Refresh
'If Opc = True Then
    dxDBGrid1.Dataset.Active = True
'Else
 '   MsgBox "No hay registro", vbCritical
' '   Command1.SetFocus
'End If
End Sub

Private Sub Command1_Click()

    If dxDBGrid1.Dataset.State = dsEdit Then dxDBGrid1.Dataset.Post

    GENERAR_REQUERIMIENTO
    
End Sub

Private Sub Form_Activate()
 sw_load_mant = False
    
    
    '-----------------------------------------------
    dxDBGrid3.Dataset.ADODataset.Requery
    
    dxDBGrid3.Options.Unset (egoShowGroupPanel)
    dxDBGrid3.Filter.FilterActive = False
    dxDBGrid3.Dataset.Close
    dxDBGrid3.Dataset.Open
    '-----------------------------------------------

End Sub

Private Sub Form_Load()
    Me.MousePointer = 11
    Me.left = 100
    Me.top = 1000
    
    cnombase = "TEMPLUS.mdb"
    cnomtabla = "tmpPedSugerido"
    
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    dxDBGrid3.Dataset.ADODataset.ConnectionString = cnn_form

    TxtDesde.Value = Format(Date, "dd/mm/yyyy")
    TxtHasta.Value = Format(Date, "dd/mm/yyyy")
    dias = 1
    
    FILL
    Me.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
cnn_form.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim n       As Byte
Dim IValue  As Byte

    Select Case Tool.Id
        Case "ID_Actualizar"
            Me.MousePointer = 11
            FILL
            Me.MousePointer = 1
'        Case "ID_Preliminar"
 '           IValue = SSActiveToolBars1.Tools.ITEM("ID_Preliminar").UseMaskColor
  '          GridNum = 1: OldValue = 1
   '         GridInit IValue, OldValue
    '        OldValue = IValue
'        Case "ID_ExportExcell"
 '           IValue = SSActiveToolBars1.Tools.ITEM("ID_ExportExcell").UseMaskColor
  '          GridNum = 1: OldValue = 1
   '         GridInit IValue - 10, OldValue
    '        OldValue = IValue
        Case "ID_Imprimir"
        
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub FILL()
Dim csql As String
Dim I As Integer
Dim rsOri As New ADODB.Recordset


sql = "delete * from " & cnomtabla
cnn_form.Execute sql
AlmacenaQuery_sql sql, cnn_form

If ctipoadm_bd = "M" Then
    csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F5STOCKMIN,IF5PLA.F5STOCKMAX, (select " & _
           "Sum(a.F6STOCKACT) from if6alma a where a.f5codpro = IF5PLA.F5CODPRO) AS STOCK, (SELECT " & _
           "sum(v1.F3CANPRO) * 1/(" & dias & ") FROM IF4VALES v2 INNER JOIN IF3VALES v1 ON (v2.F4NUMVAL = " & _
           "v1.F4NUMVAL) AND (v2.F2CODALM = v1.F2CODALM) WHERE (((v1.F5CODPRO)=if5pla.f5codpro) AND " & _
           "((v2.F1CODORI)='XV1'))) AS VENTA, If(([IF5PLA].[F5TREP]*([VENTA]))+[IF5PLA].[F5STOCKMIN]-" & _
           "[STOCK]<=0,0,[IF5PLA].[F5STOCKMAX]-[STOCK]) AS PEDIDO, IF5PLA.F5PRECOS, [PEDIDO]*[IF5PLA].[F5PRECOS] AS TOTAL " & _
           "FROM IF4VALES INNER JOIN ((IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO) " & _
           "INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = " & _
           "IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)" & _
           "WHERE (((IF3VALES.F4FECVAL)<=#" & TxtHasta.Text & "# And (IF3VALES.F4FECVAL)>=#" & TxtDesde.Text & "#) AND ((IF4VALES.F1CODORI)='XV1'))" & _
           "GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F5STOCKMIN, " & _
           "IF5PLA.F5PRECOS, IF5PLA.F5TREP,IF5PLA.F5STOCKMAX;"
Else
    csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F5STOCKMIN,IF5PLA.F5STOCKMAX, (select " & _
           "Sum(a.F6STOCKACT) from if6alma a where a.f5codpro = IF5PLA.F5CODPRO) AS STOCK, (SELECT " & _
           "sum(v1.F3CANPRO) * 1/(" & dias & ") FROM IF4VALES v2 INNER JOIN IF3VALES v1 ON (v2.F4NUMVAL = " & _
           "v1.F4NUMVAL) AND (v2.F2CODALM = v1.F2CODALM) WHERE (((v1.F5CODPRO)=if5pla.f5codpro) AND " & _
           "((v2.F1CODORI)='XV1'))) AS VENTA, IIf(([IF5PLA].[F5TREP]*([VENTA]))+[IF5PLA].[F5STOCKMIN]-" & _
           "[STOCK]<=0,0,[IF5PLA].[F5STOCKMAX]-[STOCK]) AS PEDIDO, IF5PLA.F5PRECOS, [PEDIDO]*[IF5PLA].[F5PRECOS] AS TOTAL " & _
           "FROM IF4VALES INNER JOIN ((IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO) " & _
           "INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = " & _
           "IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)" & _
           "WHERE (((IF3VALES.F4FECVAL)<=#" & TxtHasta.Text & "# And (IF3VALES.F4FECVAL)>=#" & TxtDesde.Text & "#) AND ((IF4VALES.F1CODORI)='XV1'))" & _
           "GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F5STOCKMIN, " & _
           "IF5PLA.F5PRECOS, IF5PLA.F5TREP,IF5PLA.F5STOCKMAX;"
End If
    
    If rsOri.State = adStateOpen Then rsOri.Close
    If ctipoadm_bd = "M" Then
        rsOri.Open csql, cnn_dbbancos
    Else
        rsOri.Open csql, cnn_dbbancos ' cnn_form
    End If
    If Not rsOri.EOF Then
    
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    
        rsOri.MoveFirst
            I = 0
            Do While Not rsOri.EOF
                I = I + 1
                csql = "insert into " & cnomtabla & " (ITEM,CODPRO,CODFAB,NOMPRO,STOCKMIN,STOCKMAX,STOCKACT,VENTA,PEDSUGE,VALORCOM,VALORTOT) values " & _
                       "('" & I & "','" & rsOri.Fields("F5CODPRO").Value & "','" & rsOri.Fields("F5CODFAB").Value & "','" & rsOri.Fields("F5NOMPRO") & "', " & rsOri.Fields("F5STOCKMIN").Value & _
                       "," & rsOri.Fields("F5STOCKMax").Value & "," & rsOri.Fields("STOCK").Value & "," & rsOri.Fields("VENTA") & ", " & rsOri.Fields("PEDIDO").Value & ", " & IIf(IsNull(rsOri.Fields("F5PRECOS").Value), 0, rsOri.Fields("F5PRECOS").Value) & ", " & IIf(IsNull(rsOri.Fields("TOTAL").Value), 0, rsOri.Fields("TOTAL").Value) & ")"
                cnn_form.Execute csql
                AlmacenaQuery_sql csql, cnn_form
                rsOri.MoveNext
            Loop
    End If
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.ADODataset.CommandText = "select * from " & cnomtabla
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "ITEM"
    
    rsOri.Close: Set rsOri = Nothing
    
    If dxDBGrid1.Dataset.RecordCount > 0 Then
        dxDBGrid1.Dataset.ADODataset.Requery
        
        dxDBGrid1.Options.Unset (egoShowGroupPanel)
        dxDBGrid1.Filter.FilterActive = False
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
    End If


    '-----------------------------------------------------------------------
    Dim X As Integer
    
    csql = ""
    csql = csql & "TRANSFORM Sum(IF3VALES.F3CANPRO) AS SumaDeF3CANPRO "
    csql = csql & "SELECT IF4VALES.F4FECVAL AS Fecha "
    csql = csql & "FROM IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) "
    csql = csql & "WHERE (((IF4VALES.F4FECVAL) Between #" & TxtDesde.Value & "# And #" & TxtHasta.Value & "#) AND ((IF4VALES.F1CODORI)='XV1')) "
    csql = csql & "GROUP BY IF4VALES.F4FECVAL "
    csql = csql & "ORDER BY IF4VALES.F4FECVAL, [IF3VALES].[F5CODPRO] & ' - ' & [IF5PLA].[F5NOMPRO] "
    csql = csql & "PIVOT [IF3VALES].[F5CODPRO] & ' - ' & [IF5PLA].[F5NOMPRO];"
    
    dxDBGrid3.Dataset.Active = False
    dxDBGrid3.Dataset.ADODataset.CommandText = csql
    dxDBGrid3.DefaultFields = True
    dxDBGrid3.KeyField = ""
    dxDBGrid3.Dataset.Active = True
    
    If dxDBGrid3.Dataset.RecordCount <> 0 Then
        dxDBGrid3.KeyField = "Fecha"
        For X = 0 To dxDBGrid3.Columns.Count - 1
            dxDBGrid3.Columns(X).DisableEditor = True
            dxDBGrid3.Columns(X).Alignment = taCenter
            dxDBGrid3.Columns(X).HeaderAlignment = taCenter
            dxDBGrid3.Columns(X).Visible = True
            dxDBGrid3.Columns(X).FontColor = &H0&
            dxDBGrid3.Columns(X).Color = &HC0FFFF
        Next
    End If
    
    '-----------------------------------------------------------------------
    
End Sub

Private Sub txtdesde_GotFocus()
    TxtDesde.FocusSelect = True
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtHasta.SetFocus
        calc_dias
    End If

End Sub

Private Sub txtdesde_LostFocus()
    calc_dias
End Sub

Private Sub txthasta_GotFocus()

    TxtHasta.FocusSelect = True
    
End Sub
Private Sub txthasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        calc_dias
    End If

End Sub
Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.Value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub CheckFiltro3_Click()
    If CheckFiltro3.Value = 1 Then
      dxDBGrid3.Filter.FilterActive = True
    Else
      dxDBGrid3.Filter.FilterActive = False
    End If
End Sub
Private Sub txthasta_LostFocus()
    calc_dias
End Sub

Private Sub calc_dias()
    dias = DateDiff("d", TxtDesde.Text, TxtHasta.Text) + 1
    FILL
End Sub
Private Sub GENERAR_REQUERIMIENTO()
Dim csql        As String
Dim rssol       As New ADODB.Recordset
Dim cnumsol     As String
Dim csqlcab     As String
Dim csqldet     As String
Dim RsInsumos   As New ADODB.Recordset
Dim rsif5pla    As New ADODB.Recordset
Dim sql         As String
    
    '----------------------
    If RsInsumos.State = adStateOpen Then RsInsumos.Close  'WHERE F3MARCA=False
    
    sql = ""
    sql = sql & "SELECT * FROM tmpPedSugerido "
    If Me.Check1.Value = 0 Then sql = sql & "where PEDSUGE <> 0"
    
    RsInsumos.Open sql, cnn_form, adOpenStatic, adLockOptimistic
    If RsInsumos.EOF Then
        'Opc = False
        Exit Sub
    Else
        If MsgBox("Desea generar el requerimiento automáticamente.", vbQuestion + vbYesNo, "Atención") = vbNo Then
            Exit Sub
        End If
    End If
    RsInsumos.Close
    '------- OBTENER NUMERO DE SOLICITUD
    csql = "SELECT COD_SOLICITUD FROM TB_CABSOLICITUD ORDER BY COD_SOLICITUD DESC"
    If rssol.State = adStateOpen Then rssol.Close
    rssol.Open csql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rssol.EOF Then
        cnumsol = Format(Val(rssol.Fields("COD_SOLICITUD") & "") + 1, "0000")
    Else
        cnumsol = "0001"
        'Opc = False
    End If
    rssol.Close
    Set rssol = Nothing
    '----------------------
    If ctipoadm_bd = "M" Then
    csqlcab = "INSERT INTO TB_CABSOLICITUD " & _
           "(COD_SOLICITUD,CS_FECHA,CS_CODSOLICITANTE,CS_ESTADO,CS_PRIORIDAD,CS_USUARIO,CS_MONEDA) " & _
           "VALUES ('" & cnumsol & "','" & Date & "','" & wusuario & "','P','1','" & wusuario & "','S')"
    
    Else
        csqlcab = "INSERT INTO TB_CABSOLICITUD " & _
               "(COD_SOLICITUD,CS_FECHA,CS_CODSOLICITANTE,CS_ESTADO,CS_PRIORIDAD,CS_USUARIO,CS_MONEDA) " & _
               "VALUES ('" & cnumsol & "',CVDATE('" & Date & "'),'" & wusuario & "','P','1','" & wusuario & "','S')"
    End If
    cnn_dbbancos.Execute (csqlcab)
    AlmacenaQuery_sql csqlcab, cnn_dbbancos
    
    If RsInsumos.State = adStateOpen Then RsInsumos.Close ' WHERE F3MARCA=False
    
    sql = ""
    sql = sql & "SELECT * FROM tmpPedSugerido "
    If Me.Check1.Value = 0 Then sql = sql & "where PEDSUGE <> 0"

    RsInsumos.Open sql, cnn_form, adOpenStatic, adLockOptimistic
    If Not RsInsumos.EOF Then
        RsInsumos.MoveFirst
        Dim nstockminimo As Integer
        Dim npedidominimo As Integer
        Dim ncantidad As Integer
        Do While Not RsInsumos.EOF
            nstockminimo = 0#: npedidominimo = 0#
            If rsif5pla.State = adStateOpen Then rsif5pla.Close
            rsif5pla.Open "SELECT F5STOCKMIN,F5PEDIDO_MINIMO FROM IF5PLA WHERE F5CODPRO='" & RsInsumos.Fields("CODPRO") & "'", cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rsif5pla.EOF Then
                nstockminimo = Val(rsif5pla.Fields("F5STOCKMIN") & "")
                npedidominimo = Val(rsif5pla.Fields("F5PEDIDO_MINIMO") & "")
            End If
            rsif5pla.Close: Set rsif5pla = Nothing
            
            'ncantidad = nstockminimo - (Val(RsInsumos.Fields("STOCKACT") & "") - Val(RsInsumos.Fields("PEDSUGE") & ""))
            'ncantidad = nstockminimo - (Val(RsInsumos.Fields("STOCKACT") & "") - Val(RsInsumos.Fields("PEDSUGE") & ""))
            ncantidad = Val(RsInsumos.Fields("PEDSUGE") & "")
            ncantidad = IIf(ncantidad < npedidominimo, npedidominimo, ncantidad)
            
            If ctipoadm_bd = "M" Then
                csqldet = "INSERT INTO TB_DETSOLICITUD " & _
                          "(COD_SOLICITUD,ITEM,DS_CANTIDAD,COD_PRODUCTO,DS_UNIDMED," & _
                          "DS_DESCRIPCION,CS_FENTREGA,CANDIS,CS_AFECTO) " & _
                          "VALUES ('" & cnumsol & "'," & RsInsumos.Fields("ITEM") & "," & _
                          ncantidad & ",'" & _
                          RsInsumos.Fields("CODPRO") & "','" & _
                          "UNI" & "','" & _
                          RsInsumos.Fields("NOMPRO") & "','" & Date & "'," & RsInsumos.Fields("PEDSUGE") & ",'*' ) "
            Else
                csqldet = "INSERT INTO TB_DETSOLICITUD " & _
                          "(COD_SOLICITUD,ITEM,F5CodFab,DS_CANTIDAD,COD_PRODUCTO,DS_UNIDMED," & _
                          "DS_DESCRIPCION,CS_FENTREGA,CANDIS,CS_AFECTO) " & _
                          "VALUES ('" & cnumsol & "'," & RsInsumos.Fields("ITEM") & ",'" & RsInsumos.Fields("CodFab") & "'," & _
                          ncantidad & ",'" & _
                          RsInsumos.Fields("CODPRO") & "','" & _
                          "UNI" & "','" & _
                          RsInsumos.Fields("NOMPRO") & "',CVDATE('" & Date & "')," & RsInsumos.Fields("PEDSUGE") & ",'*' ) "
            End If
            cnn_dbbancos.Execute (csqldet)
           
            AlmacenaQuery_sql csqldet, cnn_dbbancos
            RsInsumos.MoveNext
        Loop
       ' Opc = True
    End If
    RsInsumos.Close
    Set RsInsumos = Nothing

    MsgBox "Se ha generado la solicitud Nº " & cnumsol, vbInformation, "Atención"

End Sub


