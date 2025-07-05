VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Begin VB.Form hlp_ocompra 
   Caption         =   "Ayuda de Ordenes de Compra"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   5250
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1710
      TabIndex        =   2
      Top             =   135
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1770
      Left            =   90
      OleObjectBlob   =   "hlp_ocompra.frx":0000
      TabIndex        =   0
      Top             =   3285
      Width           =   10500
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2685
      Left            =   90
      OleObjectBlob   =   "hlp_ocompra.frx":2445
      TabIndex        =   3
      Top             =   495
      Width           =   10500
   End
End
Attribute VB_Name = "hlp_ocompra"
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
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wformat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal Rgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal Rgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal HBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

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

Private Sub dxDBGrid1_OnBackgroundDraw(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, Done As Boolean)
    Dim X As Integer, Y As Integer
    Dim IsClipRgnExists As Boolean
    Dim PrevClipRgn As Long, Rgn As Long
    Dim s, OldFont As Long
    Dim Font1 As IFont
    
    If dxDBGrid1.Ex.GroupColumnCount < 1 Then
     s = "Arrastre una columna aquí para agrupar información"
     SetBkMode hdc, TRANSPARENT
     Set Font1 = dxDBGrid1.Columns.HeaderFont
     OldFont = SelectObject(hdc, Font1.hFont)
     DrawText hdc, s, Len(s), R, DT_SINGLELINE + DT_VCENTER
     Call SelectObject(hdc, OldFont)
    End If
End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
    Dim valor As String
    valor = Val("" & Str(dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value))
    proceso2 (valor)
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodcosto = dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value
    wdescosto = dxDBGrid1.Columns.ColumnByFieldName("f4fecemi").Value
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.AutoRedraw = False
    Me.Left = 1600
    Me.Top = 1050
    
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    proceso
    proceso2 Val("" & Str(dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value))
    With dxDBGrid1
        .Options.Unset (egoShowGroupPanel)
        .Filter.FilterActive = False
    End With
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = 11
            If wtipoguia = "I" Then
                vale_ingreso.Show 1
            Else
                vale_salida.Show 1
            End If
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Public Sub proceso()
Dim csql     As String
    
    With dxDBGrid1
    
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos

    If whelpoc = "S" Then
        csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
               "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
               "WHERE (((IF4ORDEN.F4NUMORD) In " & _
               "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
               "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND (F4CODPRV = '" & wrucprov & "') ) " & _
               "ORDER BY IF4ORDEN.F4NUMORD;"
    Else
        csql = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
               "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
               "WHERE (((IF4ORDEN.F4NUMORD) In " & _
               "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0))) " & _
               "ORDER BY IF4ORDEN.F4NUMORD;"
    End If
    
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "f4numord"
    End With
 
End Sub

Public Sub proceso2(ByVal codigo As String)
    Dim SQL As String
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        SQL = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro FROM if3orden a left join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= " & codigo & " order by a.f3codpro"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SQL
        .Dataset.Active = True
        .KeyField = "f3codpro"
    End With
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.M.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
End With
End Sub


