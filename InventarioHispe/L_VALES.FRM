VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form FrmAyudaVale 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Vales"
   ClientHeight    =   4485
   ClientLeft      =   2430
   ClientTop       =   1605
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   4845
   Begin Threed.SSPanel SSPanel1 
      Height          =   4380
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4740
      _Version        =   65536
      _ExtentX        =   8361
      _ExtentY        =   7726
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin MSAdodcLib.Adodc DataAyuda 
         Height          =   375
         Left            =   1560
         Top             =   2640
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtope 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Simplifique la consulta digitando el Simbolo (*)"
         Top             =   3915
         Visible         =   0   'False
         Width           =   2445
      End
      Begin TrueOleDBGrid70.TDBGrid DataGrid 
         Bindings        =   "L_vales.frx":0000
         Height          =   3165
         Left            =   135
         TabIndex        =   1
         Top             =   90
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   5583
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "F2CODALM"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Vale"
         Columns(1).DataField=   "F4NUMVAL"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2487"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2408"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=74000"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=4868"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4789"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=74000"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=780,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(23)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(26)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(29)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(36)  =   ":id=28,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(41)  =   ":id=32,.locked=-1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(45)  =   "Named:id=33:Normal"
         _StyleDefs(46)  =   ":id=33,.parent=0"
         _StyleDefs(47)  =   "Named:id=34:Heading"
         _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   ":id=34,.wraptext=-1"
         _StyleDefs(50)  =   "Named:id=35:Footing"
         _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   "Named:id=36:Selected"
         _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(58)  =   "Named:id=39:EvenRow"
         _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(60)  =   "Named:id=40:OddRow"
         _StyleDefs(61)  =   ":id=40,.parent=33"
         _StyleDefs(62)  =   "Named:id=41:RecordSelector"
         _StyleDefs(63)  =   ":id=41,.parent=34"
         _StyleDefs(64)  =   "Named:id=42:FilterBar"
         _StyleDefs(65)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "[F3] --> Filtro Avanzado : Vale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   3420
         Width           =   2085
      End
      Begin VB.Label lblbusca 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   3690
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label1 
         Caption         =   "Ayuda: Hacer Click Derecho en la Columna que desee buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   2970
         TabIndex        =   3
         Top             =   3555
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Menu MNUPRI 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnufiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu mnufiltroavanz 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu mnuordasc 
         Caption         =   "Ord. Asc."
      End
      Begin VB.Menu mnuorddesc 
         Caption         =   "Ord. Desc."
      End
      Begin VB.Menu mnutodos 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "FrmAyudaVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Private Sub datagrid_DblClick()
    datagrid_KeyDown 13, 0
End Sub

Private Sub datagrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            wcod_alm = DataAyuda.Recordset.Fields("F2CODALM")
            wnumval = DataAyuda.Recordset.Fields("F4NUMVAL")
            Unload Me
        Case 27:
            wcod_alm = ""
            wnumval = ""
        Case 114:
            DataGrid.col = 1
            mnufiltroavanz_Click
End Select

End Sub


Private Sub datagrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.Visible = True
 
End Sub

Private Sub datagrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If txtope.Visible = True Then txtope.Visible = False
    If lblbusca.Visible = True Then lblbusca.Visible = False
    mnufiltro.Caption = "Filtrar [" + DataGrid.Columns(DataGrid.col).Text + "]"
    Select Case Button
        Case 2
            PopupMenu MNUPRI
    End Select
End Sub

Private Sub Form_Activate()
    
    DataGrid.OddRowStyle.BackColor = &HC0FFFF
    DataGrid.EvenRowStyle.BackColor = &HFFFFFF
    DataGrid.HighlightRowStyle.BackColor = vbActiveTitleBar
    DataGrid.HighlightRowStyle.ForeColor = vbWhite
    DataGrid.AlternatingRowStyle = True
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
      Case 27:

        wcod_alm = ""
        wnumval = ""
        Unload Me
  End Select
    
End Sub

Private Sub Form_Load()
DataAyuda.ConnectionString = cnn_dbbancos
    
    Select Case Gtipval
    Case Is = "0": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='I' AND F1CODORI='XC0' ORDER BY F2CODALM,F4NUMVAL" 'ingreso por compras
    Case Is = "1": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='I' ORDER BY F2CODALM,F4NUMVAL" 'ingresos general
    Case Is = "2": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='S' ORDER BY F2CODALM,F4NUMVAL" 'salidas general
    Case Is = "3": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='S' AND LEFT(F4NUMDOC,1)='I' ORDER BY F2CODALM,F4NUMVAL" 'salidas general

    End Select
    DataAyuda.Refresh



End Sub

Private Sub mnufiltro_Click()
        Set rs = DataAyuda.Recordset
  Select Case DataAyuda.Recordset.Fields(DataGrid.Columns(DataGrid.col).DataField).Type
         Case 10
              rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = '" + Trim("" & DataGrid.Columns(DataGrid.col).Text) + "'"
         Case 4
              rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " = " + DataGrid.Columns(DataGrid.col).Text
         Case 8
              If IsDate(DataGrid.Columns(DataGrid.col).Text) Then
                 rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + "=#" + DataGrid.Columns(DataGrid.col).Text + "#"
              Else
                 MsgBox "Ingrese una Fecha Valida..!", 32, "Advertencia"
                 Exit Sub
              End If
  End Select
  Set DataAyuda.Recordset = rs.DataSource
  Set rs = Nothing
End Sub

Private Sub mnufiltroavanz_Click()
Select Case DataGrid.col
        Case 0
                    lblbusca.Visible = True
                    lblbusca.Caption = Trim(DataGrid.Columns(0).Caption)
                    txtope.Visible = True
                    txtope.Text = ""
                    txtope.SetFocus
                    
        Case 1
                    lblbusca.Visible = True
                    lblbusca.Caption = Trim(DataGrid.Columns(1).Caption)
                    txtope.Visible = True
                    txtope.Text = ""
                    txtope.SetFocus
       
       
        End Select
End Sub

Private Sub mnuordasc_Click()
  Set rs = DataAyuda.Recordset
  rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Asc"
  Set DataAyuda.Recordset = rs.DataSource
  Set rs = Nothing
End Sub

Private Sub mnuorddesc_Click()
        Set rs = DataAyuda.Recordset
  rs.Sort = "[" + DataGrid.Columns(DataGrid.col).DataField + "] Desc"
  Set DataAyuda.Recordset = rs.DataSource
  Set rs = Nothing

End Sub

Private Sub mnutodos_Click()
    Select Case Gtipval
    Case Is = "0": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='I' AND F1CODORI='XC0' ORDER BY F2CODALM,F4NUMVAL" 'ingreso por compras
    Case Is = "1": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='I' ORDER BY F2CODALM,F4NUMVAL" 'ingresos general
    Case Is = "2": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='S' ORDER BY F2CODALM,F4NUMVAL" 'salidas general
    Case Is = "3": DataAyuda.RecordSource = "Select * FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='S' AND LEFT(F4NUMDOC,1)='I' ORDER BY F2CODALM,F4NUMVAL" 'salidas general

    End Select
    DataAyuda.Refresh

End Sub

Private Sub SSPanel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False
End Sub

Private Sub txtope_KeyPress(KeyAscii As Integer)
Dim SQL     As String
Dim DATO    As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(txtope.Text) = 0 Then txtope.SetFocus: Exit Sub
        txtope.Text = "*" & txtope.Text
        If InStr(txtope, "*") = Len(Trim(txtope)) Then
            DATO = Left(txtope, Len(Trim(txtope)) - 1)
        Else
            DATO = txtope.Text
        End If
        Set rs = DataAyuda.Recordset
        rs.Filter = "[" + DataGrid.Columns(DataGrid.col).DataField + "]" + " Like  '" + DATO + "*'"
        If rs.EOF Then txtope.SetFocus: Exit Sub Else txtope.Visible = False: lblbusca.Visible = False
        Set DataAyuda.Recordset = rs.DataSource
        Set rs = Nothing
        DataGrid.SetFocus
    End If
End Sub
