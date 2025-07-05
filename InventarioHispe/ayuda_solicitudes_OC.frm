VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_solicitudes_OC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Requerimientos"
   ClientHeight    =   6645
   ClientLeft      =   2460
   ClientTop       =   2910
   ClientWidth     =   16485
   Icon            =   "ayuda_solicitudes_OC.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   16485
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "ayuda_solicitudes_OC.frx":058A
      TabIndex        =   1
      Top             =   3000
      Width           =   16245
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2745
      Left            =   60
      OleObjectBlob   =   "ayuda_solicitudes_OC.frx":4611
      TabIndex        =   0
      Top             =   120
      Width           =   16215
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "ayuda_solicitudes_OC.frx":81E1
      ToolBars        =   "ayuda_solicitudes_OC.frx":F36B
   End
End
Attribute VB_Name = "ayuda_solicitudes_OC"
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
Dim r As Rect, REdge As Rect
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

Private strTipoDocumento As String

'Propiedad Tipo de Documento
Public Property Let TipoDocumento(ByVal value As String)
    strTipoDocumento = value
End Property

Public Property Get TipoDocumento() As String
    TipoDocumento = strTipoDocumento
End Property

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
     DrawText hDC, s, Len(s), r, DT_SINGLELINE + DT_VCENTER
     Call SelectObject(hDC, OldFont)
    End If
End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
    dxDBGrid1_OnClick
End Sub

Private Sub dxDBGrid1_OnClick()
    listarSolicitudDetalle Trim(dxDBGrid1.Columns.ColumnByFieldName("cs_documento").value & ""), _
                            Trim(dxDBGrid1.Columns.ColumnByFieldName("cod_solicitud").value & "")
                                
End Sub

Private Sub dxDBGrid1_OnDblClick()
    If dxDBGrid2.Dataset.RecordCount > 0 Then
'        TOC = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
'        num_solcomp = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value
'        item_solcomp = 0

        With objAyudaSolicitud
            .inicializarEntidades
            
            .TipoDocumento = Trim(dxDBGrid1.Columns.ColumnByFieldName("CS_DOCUMENTO").value & "")
            .Codigo = Trim(dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
            
            .CodProducto = vbNullString
        End With
        
        Me.Hide
    End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dxDBGrid1_OnDblClick
    End Select
End Sub

Private Sub dxDBGrid2_OnDblClick()
    'TOC = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
    'num_solcomp = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value
    'item_solcomp = dxDBGrid2.Columns.ColumnByFieldName("ITEM").Value
    If dxDBGrid1.Dataset.RecordCount > 0 Then
        With objAyudaSolicitud
            .inicializarEntidades
            
            .TipoDocumento = Trim(dxDBGrid1.Columns.ColumnByFieldName("CS_DOCUMENTO").value & "")
            .Codigo = Trim(dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
            
            .CodProducto = Trim(dxDBGrid2.Columns.ColumnByFieldName("COD_PRODUCTO").value & "")
        End With
        
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.AutoRedraw = False
    Me.left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 2
    
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    
    listarSolicitudesPendientes
    
    With dxDBGrid1
        .Options.Unset (egoShowGroupPanel)
        .Filter.FilterActive = False
    End With
    
    Me.MousePointer = vbDefault
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
          Case "ID_Filtrar"
            If Tool.State = ssChecked Then
                dxDBGrid1.Filter.FilterActive = True
            Else
                dxDBGrid1.Filter.FilterActive = False
            End If
        Case "ID_Agrupar"
            If Tool.State = ssChecked Then
                dxDBGrid1.Options.Set (egoShowGroupPanel)
            Else
                dxDBGrid1.Options.Unset (egoShowGroupPanel)
            End If
    End Select
End Sub

Public Sub listarSolicitudesPendientes()
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        Rem SK ADD:
        sql = vbNullString
        sql = sql & "SELECT "
        sql = sql & "CAB.CS_DOCUMENTO, "
        sql = sql & "CAB.COD_SOLICITUD, "
        sql = sql & "CAB.CS_FECHA, "
        sql = sql & "USU.F2NOMUSER AS CS_CODSOLICITANTE, "
        sql = sql & "CAB.NUMORDEN, "
        sql = sql & "CAB.CS_OBSERVACIONES AS OBSERVA "
        sql = sql & "FROM "
        sql = sql & "(TB_CABSOLICITUD AS CAB "
        sql = sql & "LEFT JOIN EF2USERS AS USU ON USU.F2CODUSER = CAB.CS_CODSOLICITANTE) "
        sql = sql & "RIGHT JOIN "
        sql = sql & "("
        sql = sql & "SELECT "
        sql = sql & "DET.CS_DOCUMENTO, DET.COD_SOLICITUD, "
        sql = sql & "SUM(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS SALDO "
        sql = sql & "FROM "
        sql = sql & "TB_DETSOLICITUD AS DET "
        sql = sql & "LEFT JOIN "
        sql = sql & "(SELECT "
        sql = sql & "DET.COD_SOLICITUD, "
        sql = sql & "DET.F3CODPRO, "
        sql = sql & "SUM(DET.F3CANPRO) AS CANTIDAD "
        sql = sql & "FROM "
        sql = sql & "IF3ORDEN AS DET "
        sql = sql & "WHERE "
        sql = sql & "DET.COD_SOLICITUD <> '' "
        sql = sql & "GROUP BY "
        sql = sql & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
        sql = sql & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND  MOVPROD.F3CODPRO = DET.COD_PRODUCTO "
        sql = sql & "GROUP BY "
        sql = sql & "DET.CS_DOCUMENTO, DET.COD_SOLICITUD) AS MOVORD "
        sql = sql & "ON MOVORD.CS_DOCUMENTO = CAB.CS_DOCUMENTO AND MOVORD.COD_SOLICITUD = CAB.COD_SOLICITUD "
        sql = sql & "WHERE "
        sql = sql & "CAB.CS_ESTADO NOT IN ('0', '1', '5') AND "
        sql = sql & "MOVORD.SALDO > 0 "
        sql = sql & "ORDER BY "
        sql = sql & "CAB.CS_FECHA DESC, CAB.COD_SOLICITUD DESC"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        
        .KeyField = "COD_SOLICITUD"
    End With
    
    sql = vbNullString
End Sub

Public Sub listarSolicitudDetalle(ByVal strTipoSolicitud As String, ByVal strCodSolicitud As String)
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        Rem SK ADD:
        sql = vbNullString
        sql = sql & "SELECT "
        sql = sql & "DET.CS_DOCUMENTO, "
        sql = sql & "DET.COD_SOLICITUD, "
        sql = sql & "DET.ITEM, "
        sql = sql & "DET.COD_PRODUCTO, "
        sql = sql & "DET.DS_DESCRIPCION, "
        sql = sql & "MED.F7SIGMED, "
        sql = sql & "DET.DS_CANTIDAD, "
        sql = sql & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS SALDO "
        sql = sql & "FROM "
        sql = sql & "(TB_DETSOLICITUD AS DET "
        sql = sql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.DS_UNIDMED) "
        sql = sql & "LEFT JOIN "
        sql = sql & "("
        sql = sql & "SELECT "
        sql = sql & "DET.COD_SOLICITUD, "
        sql = sql & "DET.F3CODPRO, "
        sql = sql & "SUM(DET.F3CANPRO) AS CANTIDAD "
        sql = sql & "FROM "
        sql = sql & "IF3ORDEN AS DET "
        sql = sql & "WHERE "
        sql = sql & "DET.COD_SOLICITUD <> '' "
        sql = sql & "GROUP BY "
        sql = sql & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
        sql = sql & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND  MOVPROD.F3CODPRO = DET.COD_PRODUCTO "
        sql = sql & "WHERE "
        sql = sql & "DET.CS_DOCUMENTO = '" & strTipoSolicitud & "' AND "
        sql = sql & "DET.COD_SOLICITUD = '" & strCodSolicitud & "' AND "
        sql = sql & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 "
        sql = sql & "ORDER BY "
        sql = sql & "DET.DS_DESCRIPCION"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        
        .KeyField = "ITEM"
    End With
    
    sql = vbNullString
End Sub
