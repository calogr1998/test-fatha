VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form cons_solicitudes_view 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Suministros de Materiales"
   ClientHeight    =   7005
   ClientLeft      =   210
   ClientTop       =   1185
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10860
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "cons_solicitudes_view.frx":0000
      TabIndex        =   4
      Top             =   4080
      Width           =   10695
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   135
      Width           =   1455
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   135
      Width           =   2055
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "cons_solicitudes_view.frx":3694
      TabIndex        =   0
      Top             =   480
      Width           =   10695
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   2700
      Top             =   6570
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      Tools           =   "cons_solicitudes_view.frx":5F59
      ToolBars        =   "cons_solicitudes_view.frx":856E
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   4650
   End
End
Attribute VB_Name = "cons_solicitudes_view"
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
Private Sub PROCESA()
Dim csql                As String
Dim cobra               As String
Dim sql     As String
    
    If Len(Trim(cons_solicitudes.txtobra.Text)) > 0 Then
        cobra = " AND CS_CODCOSTO = '" & cons_solicitudes.txtobra.Text & "'"
    Else
        cobra = ""
    End If
    csql = "SELECT COD_SOLICITUD,CS_FECHA,CS_CODCOSTO,CS_DESCOSTO,CS_LUGENTR,CS_OBSERVACIONES FROM TB_CABSOLICITUD WHERE CS_CODSOLICITANTE = '" & Trim(right(cons_solicitudes.cmbsolicitante.Text, 8)) & "' AND CS_FECHA >= CVDATE('" & cons_solicitudes.abofdesde.Text & "') AND CS_FECHA <= CVDATE('" & cons_solicitudes.abofhasta.Text & "')" & cobra & " ORDER BY COD_SOLICITUD"
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "ALM_VALE"
    End With
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim valor As String
    valor = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value
    proceso2 (valor)
End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
Dim valor As String
    valor = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").value
    proceso2 (valor)
End Sub


Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1050
    PROCESA
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.MousePointer = vbDefault
End Sub
Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Salir":
            Unload Me
        Case "ID_Resumen":
            imprime_resumen
        Case "ID_Detallado":
            imprime_detalle
    End Select

End Sub

Private Sub imprime_resumen()
Dim dbtempo             As DAO.Database
Dim tbtempres           As DAO.Recordset
Dim csql                As String
Dim cobra               As String
Dim rssolcab            As New ADODB.Recordset

    If Len(Trim(cons_solicitudes.txtobra.Text)) > 0 Then
        cobra = " AND CS_CODCOSTO = '" & cons_solicitudes.txtobra.Text & "'"
    Else
        cobra = ""
    End If
    csql = "SELECT COD_SOLICITUD,CS_FECHA,CS_CODCOSTO,CS_DESCOSTO,CS_LUGENTR,CS_OBSERVACIONES FROM TB_CABSOLICITUD WHERE CS_CODSOLICITANTE = '" & Trim(right(cons_solicitudes.cmbsolicitante.Text, 8)) & "' AND CS_FECHA >= CVDATE('" & cons_solicitudes.abofdesde.Text & "') AND CS_FECHA <= CVDATE('" & cons_solicitudes.abofhasta.Text & "')" & cobra & " ORDER BY COD_SOLICITUD"
    rssolcab.Open csql, cnn_dbbancos
    
    If Not rssolcab.EOF Then
                
        Set dbtempo = OpenDatabase(wrutatemp & "\TEMPLUS.MDB")
        csql = ("DELETE * FROM tmpConsultaSolicitudes")
        dbtempo.Execute csql
        'AlmacenaQuery_sql sql, dbtempo
         
        Set tbtempres = dbtempo.OpenRecordset("tmpConsultaSolicitudes")
    
        rssolcab.MoveFirst
        Do While Not rssolcab.EOF
            tbtempres.AddNew
            tbtempres.Fields("EMPRESA") = wnomcia
            tbtempres.Fields("TITULO") = "REPORTE RESUMIDO DE SOLICITUDES DE MATERIALES"
            tbtempres.Fields("NUMSOL") = rssolcab.Fields("COD_SOLICITUD")
            tbtempres.Fields("FECHA") = rssolcab.Fields("CS_FECHA")
            tbtempres.Fields("CODOBRA") = rssolcab.Fields("CS_CODCOSTO")
            tbtempres.Fields("DESOBRA") = Trim(left(rssolcab.Fields("CS_DESCOSTO"), 50))
            tbtempres.Fields("LUGAR") = Trim(left(rssolcab.Fields("CS_LUGENTR"), 50))
            tbtempres.Fields("OBSERVACIONES") = Trim(left(rssolcab.Fields("CS_OBSERVACIONES"), 100))
            tbtempres.Fields("TITULO2") = "DEL " & Format(cons_solicitudes.abofdesde.Text, "DD/MM/YYYY") & " AL " & Format(cons_solicitudes.abofhasta.Text, "DD/MM/YYYY")
            tbtempres.Fields("SOLICITANTE") = Trim(left(cons_solicitudes.cmbsolicitante.Text, 50))
            tbtempres.Update
            rssolcab.MoveNext
        Loop
    
        tbtempres.Close
        dbtempo.Close
    End If
    rssolcab.Close
    
    'cryreporte.DataFiles(0) = wrutatemp & "\tempcomp.mdb"
    'cryreporte.ReportFileName = wrutatemp & "\sol_resumen.rpt"
    'cryreporte.Action = 1 hacer active jcg

End Sub

Private Sub imprime_detalle()
Dim dbtempo             As DAO.Database
Dim tbtempres           As DAO.Recordset
Dim csql                As String
Dim cobra               As String
Dim rssolcab            As New ADODB.Recordset

    If Len(Trim(cons_solicitudes.txtobra.Text)) > 0 Then
        cobra = " AND CS_CODCOSTO = '" & cons_solicitudes.txtobra.Text & "'"
    Else
        cobra = ""
    End If
    
    csql = "SELECT A.CS_DESCOSTO,A.CS_CODCOSTO,A.COD_SOLICITUD,A.CS_CODSOLICITANTE,A.CS_FECHA,B.ITEM,B.COD_PRODUCTO,B.DS_DESCRIPCION,B.DS_UNIDMED,B.DS_CANTIDAD,B.PRECIO,B.PRESUG,B.CS_FENTREGA FROM TB_CABSOLICITUD AS A,TB_DETSOLICITUD AS B WHERE A.COD_SOLICITUD=B.COD_SOLICITUD AND A.CS_CODSOLICITANTE = '" & Trim(right(cons_solicitudes.cmbsolicitante.Text, 8)) & "' AND A.CS_FECHA >= CVDATE('" & cons_solicitudes.abofdesde.Text & "') AND A.CS_FECHA <= CVDATE('" & cons_solicitudes.abofhasta.Text & "')" & cobra & " ORDER BY B.COD_SOLICITUD,B.ITEM"
    
    rssolcab.Open csql, cnn_dbbancos
    
    If Not rssolcab.EOF Then
                
        Set dbtempo = OpenDatabase(wrutatemp & "\TEMPLUS.MDB")
        
        csql = ("DELETE * FROM tmpConsultaSolicitudes")
        dbtempo.Execute csql
        'AlmacenaQuery_sql csql, dbtempo
        
        Set tbtempres = dbtempo.OpenRecordset("tmpConsultaSolicitudes")
    
        rssolcab.MoveFirst
        Do While Not rssolcab.EOF
            tbtempres.AddNew
            tbtempres.Fields("EMPRESA") = wnomcia
            tbtempres.Fields("TITULO") = "REPORTE DETALLADO DE SOLICITUDES DE MATERIALES"
            tbtempres.Fields("NUMSOL") = rssolcab.Fields("COD_SOLICITUD")
            tbtempres.Fields("CODOBRA") = rssolcab.Fields("CS_CODCOSTO")
            tbtempres.Fields("DESOBRA") = Trim(left(rssolcab.Fields("CS_DESCOSTO"), 50))
            
            tbtempres.Fields("ITEM") = rssolcab.Fields("ITEM")
            tbtempres.Fields("CODPRO") = rssolcab.Fields("COD_PRODUCTO")
            tbtempres.Fields("DESPRO") = Trim(left(rssolcab.Fields("DS_DESCRIPCION"), 50))
            tbtempres.Fields("UMED") = rssolcab.Fields("DS_UNIDMED")
            tbtempres.Fields("CANT") = rssolcab.Fields("DS_CANTIDAD")
            tbtempres.Fields("PCOSTO") = rssolcab.Fields("PRECIO")
            tbtempres.Fields("PSUGERIDO") = rssolcab.Fields("PRESUG")
            tbtempres.Fields("FENTREGA") = rssolcab.Fields("CS_FENTREGA")
            
            
            tbtempres.Fields("TITULO2") = "DEL " & Format(cons_solicitudes.abofdesde.Text, "DD/MM/YYYY") & " AL " & Format(cons_solicitudes.abofhasta.Text, "DD/MM/YYYY")
            tbtempres.Fields("SOLICITANTE") = Trim(left(cons_solicitudes.cmbsolicitante.Text, 50))
            tbtempres.Update
            rssolcab.MoveNext
        Loop
    
        tbtempres.Close
        dbtempo.Close
    End If
    rssolcab.Close
    
  '  cryreporte.DataFiles(0) = wrutatemp & "\tempcomp.mdb"
  '  cryreporte.ReportFileName = wrutatemp & "\sol_detallado.rpt"
  '  cryreporte.Action = 1 'hacer active jcg

End Sub
Private Sub CheckFiltro_Click()
    If CheckFiltro.value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Public Sub proceso2(ByVal Codigo As String)
    
    Dim sql As String
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        sql = "SELECT COD_SOLICITUD,ITEM,COD_PRODUCTO,DS_DESCRIPCION,DS_UNIDMED,DS_CANTIDAD,PRECIO,PRESUG,CS_FENTREGA FROM TB_DETSOLICITUD WHERE COD_SOLICITUD= '" & dxDBGrid1.Columns(0).value & "' ORDER BY COD_SOLICITUD,ITEM"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "f5codpro"
    End With
    
    
End Sub
