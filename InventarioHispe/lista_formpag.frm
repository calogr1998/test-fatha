VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form lista_formpag 
   Caption         =   "Listado de Formas de Pago"
   ClientHeight    =   4080
   ClientLeft      =   765
   ClientTop       =   2175
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4935
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4770
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   3285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   315
         TabIndex        =   3
         Top             =   405
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc adoctasctes 
      Height          =   330
      Left            =   2025
      Top             =   6795
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "adoctasctes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid tdbformpag 
      Bindings        =   "lista_formpag.frx":0000
      Height          =   750
      Left            =   180
      TabIndex        =   0
      Top             =   4725
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1323
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "f2forpag"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion del Pago"
      Columns(1).DataField=   "f2despag"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2540"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2461"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=260"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=13705"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=13626"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=260"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12648447
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=0"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   6975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "lista_formpag.frx":001A
      ToolBars        =   "lista_formpag.frx":58AA
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2895
      Left            =   45
      OleObjectBlob   =   "lista_formpag.frx":5A0E
      TabIndex        =   4
      Top             =   990
      Width           =   4770
   End
End
Attribute VB_Name = "lista_formpag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col As TrueOleDBGrid70.Column
Dim cols As TrueOleDBGrid70.Columns

Private Sub dxDBGrid1_OnDblClick()
sw_nuevo_mant = False
mant_formpag.Show 1
End Sub

Private Sub Form_Load()
    
    'Me.Height = 7890
    'Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    
    sw_load_mant = False
    
'    tdbformpag.EvenRowStyle.BackColor = &HFFFFFF
'    tdbformpag.OddRowStyle.BackColor = &HC0FFFF
'    tdbformpag.HighlightRowStyle.BackColor = vbActiveTitleBar
'    tdbformpag.HighlightRowStyle.ForeColor = vbWhite
'    tdbformpag.AlternatingRowStyle = True
'
'    adoctasctes.ConnectionString = cnn_dbbancos
'    adoctasctes.RecordSource = "Select * FROM EF2FORPAG ORDER BY F2FORPAG"
'    adoctasctes.Refresh
'

With dxDBGrid1
        .DefaultFields = True
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
FILL
End Sub

Private Sub FILL()

    csql = "Select F2FORPAG,F2DESPAG FROM EF2FORPAG ORDER BY F2FORPAG"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2FORPAG"
    CABECERA

    dxDBGrid1.Columns(0).Color = &HC0FFFF
    dxDBGrid1.Columns(1).Color = &HC0FFFF
       
End Sub
Private Sub CABECERA()

    With dxDBGrid1
        .Columns(0).Caption = "Codigo": .Columns(0).Width = 45: .Columns(0).DisableEditor = True
        .Columns(1).Caption = "Descripcion.": .Columns(1).Width = 70: .Columns(1).DisableEditor = True
    End With
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
        sw_nuevo_mant = True
        mant_formpag.Show 1
    Case "ID_Imprimir":
        With Acr_formpag
            .DataControl1.ConnectionString = cnn_dbbancos
            .DataControl1.Source = "Select * FROM EF2FORPAG ORDER BY F2FORPAG"
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .lblempresa.Caption = wnomcia
            .Show 1
        End With
    Case "ID_Salir"
        Unload Me
        
End Select
End Sub

'
'Private Sub tdbformpag_FilterChange()
'On Error GoTo errhandler
'Set cols = tdbformpag.Columns
'Dim c As Integer
'    c = tdbformpag.col
'    tdbformpag.HoldFields
'    adoctasctes.Recordset.Filter = getFilter()
'    tdbformpag.col = c
'    tdbformpag.EditActive = True
'    Exit Sub
'errhandler:
'
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description
'
'
'  For Each col In tdbformpag.Columns
'        col.FilterText = ""
'    Next col
'        adoctasctes.Recordset.Filter = adFilterNone
'
'End Sub


'Private Function getFilter() As String
'Dim cadena As String
'Dim n As Integer
'    For Each col In cols
'        If Trim(col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                cadena = cadena & " AND "
'            End If
'            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
'        End If
'    Next col
'    getFilter = cadena
'
'End Function



Private Sub txtbusqueda_Change()
     
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "f2forpag LIKE '*" & txtbusqueda.Text & "*' OR " & " f2despag LIKE '*" & txtbusqueda.Text & "*' "
    
    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "f2forpag LIKE '*" & txtbusqueda.Text & "*' OR " & " f2despag LIKE '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If

End Sub
