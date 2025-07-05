VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form cons_ordenes_compra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ordenes de Compra"
   ClientHeight    =   8955
   ClientLeft      =   1035
   ClientTop       =   1845
   ClientWidth     =   19980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   19980
   WindowState     =   2  'Maximized
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   7455
      Left            =   120
      OleObjectBlob   =   "cons_ordenes_compra.frx":0000
      TabIndex        =   6
      Top             =   1080
      Width           =   19695
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   13995
      _Version        =   65536
      _ExtentX        =   24686
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
         Left            =   3105
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
         Text            =   "08/03/2011"
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
         ButtonPicture   =   "cons_ordenes_compra.frx":531A
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
         Left            =   6480
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
         Text            =   "08/03/2011"
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
         ButtonPicture   =   "cons_ordenes_compra.frx":566C
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
         Left            =   2280
         TabIndex        =   4
         Top             =   405
         Width           =   465
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
         Left            =   5640
         TabIndex        =   3
         Top             =   405
         Width           =   420
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7740
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   19755
      _Version        =   65536
      _ExtentX        =   34846
      _ExtentY        =   13652
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
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "cons_ordenes_compra.frx":59BE
      ToolBars        =   "cons_ordenes_compra.frx":D8AE
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   810
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
End
Attribute VB_Name = "cons_ordenes_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnombase            As String
Dim cnomtabla           As String
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection
Dim GridNum             As Byte
Dim OldValue            As Byte
Dim sw_nuevo_item       As Boolean

Public Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim I As Byte
    
    If Ind > 199 Then
        SaveTo (Ind)
        Exit Sub
    End If
    
End Sub

Public Sub SaveTo(Index)
On Error GoTo errhandler
Dim FileName As String

    If GridNum <> 0 Then
        With cmdSave
            .CancelError = True
            .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
            '.DialogTitle = menu.dxSideBar1.StuckLink.Item.Caption
            .DialogTitle = "Ordenes de Compra"
            Select Case Index
                Case 204
                    .Filter = "Text Files (*.txt)|*.txt"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    If GetGridByActive().Ex.SelectedCount = 0 Then
                        GetGridByActive().m.SaveAllToTextFile (FileName)
                    Else
                        GetGridByActive().m.SaveSelectedToTextFile (FileName)
                    End If
                Case 245
                    .Filter = "Excel Files (*.xls)|*.xls"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToXLS FileName
                Case 202
                    .Filter = "HTML Files (*.htm)|*.htm"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToHTML FileName
                Case 205
                    .Filter = "XML Files (*.xml)|*.xml"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToXML FileName
                Case 201
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Sistema de Logistica") = vbYes Then _
                        GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
    End If
    
errhandler:
    
    Exit Sub
 
End Sub

Public Function GetGridByActive() As dxDBGrid
    
    Set GetGridByActive = dxDBGrid2
    
End Function

Private Sub Checkagrupar_Click()

End Sub

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

Select Case UCase(Column.FieldName)
   Case "TOTAL"
    Text = Format(Text, "#,###,###0.00")
   Case "F3TOTAL", "F3PREUNI"
    Text = Format(Text, "#,###,###0.00")
End Select

End Sub


Private Sub dxDBGrid2_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

Select Case UCase(Column.FieldName)
   Case "TOTAL"
        Text = Format(Text, "#,###,###0.00")
        Color = &HC0FFFF
   Case "F4TOTAL"
        Text = Format(Text, "#,###,###0.00")
        Color = &HC0FFC0
End Select
End Sub

Private Sub Form_Load()
Dim CadSql      As String
    
    Me.MousePointer = 11
    
    Me.left = 0
    Me.top = 0

    Me.MousePointer = 11

    TxtDesde.Value = Format(Date, "dd/mm/yyyy")
    TxtHasta.Value = Format(Date, "dd/mm/yyyy")
    CONFIGURA_GRID
    LLENA_TEMPORAL
    
    Me.MousePointer = 1
    Me.MousePointer = 1
End Sub

Private Sub LLENA_TEMPORAL()
Dim X       As Integer

        csql = "SELECT CENTROS.F3ABREV, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODPRO, IF5PLA.F5NOMPRO, IF3ORDEN.F3CANPRO, EF7MEDIDAS.F7SIGMED, REGISDOC.F4NOMPRV, IF3ORDEN.F3PREUNI, IF3ORDEN.F3TOTAL, REGISDOC.F4FECHA, REGISDOC.F4TOTAL, PAG_DCTO.nro_comp, PAG_DCTO.total, PAG_DCTO.saldo " & _
        "FROM ((REGISDOC LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) RIGHT JOIN (IF4ORDEN INNER JOIN CENTROS ON IF4ORDEN.F4CENTRO = CENTROS.F3COSTO) ON REGISDOC.F4OCOMPRA = IF4ORDEN.F4NUMORD) INNER JOIN ((IF3ORDEN INNER JOIN IF5PLA ON IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) " & _
        "WHERE IF4ORDEN.F4ESTNUL='N'"
        
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
    End With
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '5cnn_form.Close
    
    dxDBGrid2.Dataset.Close
    'ELIMINA_BD_N wrutatemp & "\", cnombase
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim n       As Byte
Dim IValue  As Byte
Dim o_Excel As Excel.Application
Dim FileName As String


    Select Case Tool.Id
        Case "ID_Actualizar"
            Me.MousePointer = 11
            LLENA_TEMPORAL
            Me.MousePointer = 1
        Case "ID_Preliminar"
            IValue = SSActiveToolBars1.Tools.ITEM("ID_Preliminar").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue, OldValue
            OldValue = IValue
        Case "ID_ExportExcell"
            With cmdSave
                .CancelError = True
                .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
                .DialogTitle = wnomcia
                .Filter = "Excel Files (*.xls)|*.xls"
                .FileName = ""
                .ShowSave
                FileName = .FileName
                GetGridByActive().m.ExportToXLS FileName
                Set o_Excel = CreateObject("Excel.application")
                o_Excel.Workbooks.Open FileName:=.FileName
                o_Excel.Visible = True
                If Not o_Excel Is Nothing Then
                    Set o_Excel = Nothing
                End If
            End With
        Case "ID_Imprimir"
        
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub txtdesde_GotFocus()

    TxtDesde.FocusSelect = True

End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtHasta.SetFocus
    End If

End Sub

Private Sub txthasta_GotFocus()

    TxtHasta.FocusSelect = True
    
End Sub
Private Sub CONFIGURA_GRID()
    
    With dxDBGrid2.Options
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoCanInsert)
        .Set (egoDynamicLoad)
        .Set (egoEditing)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoImmediateEditor)
        .Set (egoBandHeaderWidth)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoVertThrough)
        .Set (egoEnterShowEditor)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoShowCellTip)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowIndicator)
        .Set (egoShowBorder)
    End With


End Sub
