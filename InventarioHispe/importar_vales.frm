VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form importar_vales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Vales"
   ClientHeight    =   7290
   ClientLeft      =   660
   ClientTop       =   1215
   ClientWidth     =   10215
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10215
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      Top             =   105
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   105
      Width           =   1455
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   405
      Top             =   -90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "importar_vales.frx":0000
      ToolBars        =   "importar_vales.frx":32A8
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4530
      Left            =   120
      OleObjectBlob   =   "importar_vales.frx":3364
      TabIndex        =   0
      Top             =   405
      Width           =   10065
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1650
      Left            =   120
      OleObjectBlob   =   "importar_vales.frx":6B9E
      TabIndex        =   4
      Top             =   5040
      Width           =   10005
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
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
      Height          =   210
      Left            =   5400
      TabIndex        =   1
      Top             =   90
      Width           =   4650
   End
End
Attribute VB_Name = "importar_vales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditLookUp  As Boolean
Dim i           As Byte

Dim X As Integer, Y As Integer
Dim IsClipRgnExists As Boolean
Dim PrevClipRgn As Long, Rgn As Long
Dim R As Rect, REdge As Rect
Dim DBName As String

Const TRANSPARENT = 1
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BDR_OUTER = &H3
Const BDR_INNER = &HC
Const COLOR_BTNFACE = 15
Const SRCCOPY = &HCC0020
Const DT_CENTER = &H1
Const DT_RIGHT = &H2
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10
Const DT_SINGLELINE = &H20
Const DT_NOPREFIX = &H800

Private Type Rect
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wformat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nindex As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal HBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Private Sub Checkagrupar_Click()
    If Checkagrupar.value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If

End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxDBGrid1_OnBackgroundDraw(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, Done As Boolean)
    Dim X As Integer, Y As Integer
    Dim IsClipRgnExists As Boolean
    Dim PrevClipRgn As Long, Rgn As Long
    Dim s, OldFont As Long
    Dim Font1 As IFont

    If dxDBGrid1.Ex.GroupColumnCount < 1 Then
     s = "Arrastre una columna aquí para agrupar información"
     SetBkMode hDC, TRANSPARENT
     Set Font1 = dxDBGrid1.Columns.HeaderFont
     OldFont = SelectObject(hDC, Font1.hFont)
     DrawText hDC, s, Len(s), R, DT_SINGLELINE + DT_VCENTER
     Call SelectObject(hDC, OldFont)
    End If
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    dxDBGrid1_RowColChange
End Sub

Sub dxDBGrid1_RowColChange()

    proceso2 dxDBGrid1.Columns.ColumnByFieldName("f4numval").value, dxDBGrid1.Columns.ColumnByFieldName("f2codalm").value

End Sub

Private Sub dxDBGrid1_OnClick()
dxDBGrid1_RowColChange

End Sub

Private Sub dxDBGrid1_OnDblClick()
    sw_nuevo_documento = False
    Me.MousePointer = vbHourglass
    If wtipoguia = "I" Then
        vale_ingreso.Show 1
    Else
        vale_salida.Show 1
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    dxDBGrid1.Filter.FilterActive = False
   
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.AutoRedraw = False
    
    Me.left = 1600
    Me.top = 1050
    
    sw_nuevo_documento = True
    If wtipoguia = "S" Then
        Me.Caption = "Vale de Salida"
    Else
        Me.Caption = "Vale de Ingreso"
    End If
    Me.AutoRedraw = True
    proceso
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    dxDBGrid1_RowColChange
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
wtipoguia = ""
Set lista_vales = Nothing
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = vbHourglass
            If wtipoguia = "I" Then
                vale_ingreso.Show 1
            Else
                vale_salida.Show 1
            End If
            Me.MousePointer = vbDefault
        Case "ID_Salir"
            Unload Me
    End Select

End Sub

Public Sub proceso()
Dim sql     As String
    
    With dxDBGrid1
    
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        If wtipoguia = "S" Then
           ' SQL = "SELECT A.F2CODALM & A.F4NUMVAL AS ALM_VALE,A.F2CODALM,A.F4NUMVAL,A.F4SERGUIA, A.F4NUMGUIA, A.F4SERDOC, A.F4NUMDOC,A.F4FECVAL,B.F1NOMORI " & _
           ' "FROM IF4VALES AS A, SF1ORIGENES AS B WHERE MID(F4NUMVAL,1,1) = 'S' AND " & _
           ' "A.F1CODORI=B.F1CODORI ORDER BY A.F2CODALM,A.F4NUMVAL DESC"
           sql = "SELECT A.F2CODALM & A.F4NUMVAL AS ALM_VALE,A.F2CODALM,A.F4NUMVAL,A.F4FECVAL,B.F1NOMORI " & _
            "FROM IF4VALES AS A, SF1ORIGENES AS B WHERE MID(F4NUMVAL,1,1) = 'S' AND " & _
            "A.F1CODORI=B.F1CODORI ORDER BY A.F2CODALM,A.F4FECVAL DESC"

        Else
            sql = "SELECT A.F2CODALM & A.F4NUMVAL AS ALM_VALE,A.F2CODALM,A.F4NUMVAL,A.F4SERGUIA, A.F4NUMGUIA,A.F4SERDOC,A.F4NUMDOC,A.F4FECVAL,B.F1NOMORI " & _
            "FROM IF4VALES AS A, SF1ORIGENES AS B WHERE MID(F4NUMVAL,1,1) = 'I' AND " & _
            "A.F1CODORI=B.F1CODORI ORDER BY A.F4FECVAL DESC,A.F2CODALM,A.F4NUMVAL DESC"
         End If

        .Columns("0").Visible = False

        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "f4numval"
    End With
 
End Sub

Public Sub proceso2(ByVal Codigo As String, almacen As String)
    Dim sql As String
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        sql = "select a.f5codpro,b.f5nompro,b.f7codmed,a.f3canpro,a.f3punit,a.f3valvta "
        sql = sql & "from if3vales a inner join if5pla b on a.f5codpro=b.f5codpro "
        sql = sql & "where a.f4numval='" & Codigo & "' AND a.f2codalm='" & almacen & "'"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "f5codpro"
    End With
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With
End Sub
