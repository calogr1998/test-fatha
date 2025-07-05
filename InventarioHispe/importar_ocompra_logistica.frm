VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form importar_ocompra_logistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Ordenes de Compra"
   ClientHeight    =   6870
   ClientLeft      =   2115
   ClientTop       =   2235
   ClientWidth     =   11820
   Icon            =   "importar_ocompra_logistica.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
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
      Height          =   3210
      Left            =   60
      OleObjectBlob   =   "importar_ocompra_logistica.frx":000C
      TabIndex        =   0
      Top             =   3600
      Width           =   11700
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3045
      Left            =   60
      OleObjectBlob   =   "importar_ocompra_logistica.frx":2922
      TabIndex        =   3
      Top             =   495
      Width           =   11685
   End
End
Attribute VB_Name = "importar_ocompra_logistica"
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

Private Sub dxDBGrid1_OnChangeNodeEx()
    'Dim valor As String
    'valor = ("" & (dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value))
    proceso2 Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodcosto = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
    wdescosto = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4FECEMI").value & "")
    
    Me.Hide
End Sub

Private Sub Form_Load()
    'Me.AutoRedraw = False
    Me.left = 1600
    Me.top = 1150
    
    sw_nuevo_documento = True
    'Me.AutoRedraw = True
    proceso
    
    dxDBGrid1.Dataset.ADODataset.Requery
    'proceso2 (dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value)
    With dxDBGrid1
        .Options.Unset (egoShowGroupPanel)
        .Filter.FilterActive = False
    End With
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
    'Dim SqlCad     As String
    
    With dxDBGrid1
    
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        If ctipoadm_bd = "M" Then
            SqlCad = vbNullString
            SqlCad = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4OBSERVA, IF4ORDEN.F4CODPRV, EF2PROVEEDORES.F2NOMPROV"
            SqlCad = SqlCad & " FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC ORDER BY IF4ORDEN.F4NUMORD desc"
        Else
            If whelpoc = "S" Then
                If wtipoc = "I" Then
                    SqlCad = vbNullString
                    SqlCad = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                            "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                            "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                            "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
                            "AND F4LOCAL='0' " & _
                            "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND (F4CODPRV = '" & wrucprov & "') ) " & _
                            "ORDER BY IF4ORDEN.F4NUMORD desc"
                Else
                'SqlCad = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                '       "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                '       "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                '       "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
                '       "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND (F4CODPRV = '" & wrucprov & "') ) " & _
                '       "ORDER BY IF4ORDEN.F4NUMORD desc"
'                    SqlCad = "SELECT IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4CODPRV, IF3ORDEN.F4NUMORD, First(IF4ORDEN.F4OBSERVA) AS F4OBSERVA "
'                    SqlCad = SqlCad & "FROM IF4ORDEN INNER JOIN (IF3ORDEN LEFT JOIN (SELECT IF4VALES.NUMORDEN, IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO, IF3VALES.F3CANPRO "
'                    SqlCad = SqlCad & "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM))as cantidadxorden ON (IF3ORDEN.F4NUMORD = cantidadxorden.NUMORDEN) AND (IF3ORDEN.F3CODPRO = cantidadxorden.F5CODPRO)) ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
'                    SqlCad = SqlCad & "GROUP BY IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4CODPRV, IF3ORDEN.F4NUMORD "
'                    SqlCad = SqlCad & "HAVING (((Sum([IF3ORDEN].[F3CANPRO]-Val([cantidadxorden].[F3CANPRO] & '')))>0) AND (F4CODPRV = '" & wrucprov & "'))"
                    
                    Rem SK ADD:
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "CAB.F4FECEMI, "
                    SqlCad = SqlCad & "CAB.F4CENTRO, "
                    SqlCad = SqlCad & "CAB.F4CODPRV, "
                    SqlCad = SqlCad & "CAB.F4NUMORD, "
                    SqlCad = SqlCad & "CAB.F4OBSERVA "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF4ORDEN AS CAB "
                    SqlCad = SqlCad & "LEFT JOIN "
                    SqlCad = SqlCad & "(SELECT "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "COUNT(DET.F3CODPRO) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "(((IF3ORDEN AS DET "
                    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
                    SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
                    SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
                    SqlCad = SqlCad & "LEFT JOIN "
                    SqlCad = SqlCad & "(SELECT "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
                    SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF3VALES AS DET "
                    SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
                    SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') <> '' "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "DET.F4NUMORD, "
                    SqlCad = SqlCad & "DET.COD_SOLICITUD, "
                    SqlCad = SqlCad & "DET.F5CODPROORIGINAL"
                    SqlCad = SqlCad & ") AS INGRESOS "
                    SqlCad = SqlCad & "ON "
                    SqlCad = SqlCad & "INGRESOS.F4NUMORD = DET.F4NUMORD AND "
                    SqlCad = SqlCad & "INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND "
                    SqlCad = SqlCad & "INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
                    SqlCad = SqlCad & "(  (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)  )  ) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)  ) - VAL(INGRESOS.CANTIDAD & '') > 0 "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "DET.F4NUMORD) AS DETALLE ON DETALLE.F4NUMORD = CAB.F4NUMORD "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "CAB.F4LOCAL = 'OC' AND "
                    SqlCad = SqlCad & "CAB.F4CODPRV = '" & wrucprov & "' AND "
                    SqlCad = SqlCad & "DETALLE.CANTIDAD > 0"
                End If
            Else
                If wtipoc = "I" Then
                    SqlCad = vbNullString
                    SqlCad = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                           "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                           "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                           "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0))) " & _
                           "AND F4LOCAL='0' " & _
                           "ORDER BY IF4ORDEN.F4NUMORD desc"
                           
                    SqlCad = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CODPRV, P.F2NOMPROV AS PROV, E.F2NOMPROV AS EMB, IF4ORDEN.F4OBSERVA "
                    SqlCad = SqlCad & "FROM (IF4ORDEN LEFT JOIN EF2PROVEEDORES AS P ON IF4ORDEN.F4CODPRV = P.F2NEWRUC) "
                    SqlCad = SqlCad & "LEFT JOIN EF2PROVEEDORES AS E ON IF4ORDEN.F4CODCLI = E.F2CODPROV "
                    SqlCad = SqlCad & "WHERE (((IF4ORDEN.F4NUMORD) "
                    SqlCad = SqlCad & "In (SELECT DISTINCTROW IF4ORDEN.F4NUMORD "
                    SqlCad = SqlCad & "FROM IF4ORDEN INNER JOIN IF3ORDEN "
                    SqlCad = SqlCad & "ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD "
                    SqlCad = SqlCad & "GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV "
                    SqlCad = SqlCad & "HAVING Sum(IF3ORDEN.F3CANFAL)>0)) AND ((IF4ORDEN.F4LOCAL)='0')) "
                    SqlCad = SqlCad & "ORDER BY IF4ORDEN.F4NUMORD DESC"
                Else
                    SqlCad = vbNullString
                    SqlCad = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4FECEMI,IF4ORDEN.F4CENTRO,IF4ORDEN.F4OBSERVA,IF4ORDEN.F4CODPRV,EF2PROVEEDORES.F2NOMPROV " & _
                            "FROM IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC " & _
                            "WHERE (((IF4ORDEN.F4NUMORD) In " & _
                            "(SELECT DISTINCTROW IF4ORDEN.F4NUMORD FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD GROUP BY IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV HAVING Sum(IF3ORDEN.F3CANFAL)>0))) " & _
                            "ORDER BY IF4ORDEN.F4NUMORD desc"
                End If
            End If
        End If
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "F4NUMORD"
    End With
End Sub

Public Sub proceso2(ByVal strNumeroOrden As String)
    'Dim sql As String
    
    With dxDBGrid2
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        'sql = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro, a.f3canfal FROM if3orden a left join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= '" & codigo & "' order by a.f3codpro"
'        sql = vbNullString
'        sql = sql & "SELECT IF3ORDEN.F3CODPRO, IF3ORDEN.F5NOMPRO, IF3ORDEN.UNIDAD as f7codmed, IF3ORDEN.F3CANPRO, IF3ORDEN.F3CANPRO-Val(cantidadxorden.F3CANPRO & '') AS f3canfal"
'        sql = sql & " FROM IF3ORDEN LEFT JOIN (SELECT IF4VALES.NUMORDEN, IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO, IF3VALES.F3CANPRO"
'        sql = sql & " FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
'        sql = sql & " ) as cantidadxorden ON (IF3ORDEN.F3CODPRO = cantidadxorden.F5CODPRO) AND (IF3ORDEN.F4NUMORD = cantidadxorden.NUMORDEN)"
'        sql = sql & " WHERE (((IF3ORDEN.F4NUMORD)='" & Codigo & "') AND (([IF3ORDEN].[F3CANPRO]-Val([cantidadxorden].[F3CANPRO] & ''))>0)) order by IF3ORDEN.f3codpro"
        
        Rem SK ADD:
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        'SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "IIF(TRIM(DET.COD_SOLICITUD & '') <> '', DET.COD_SOLICITUD, 'STOCK LIBRE') AS COD_SOLICITUD,"
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "DET.F5NOMPRO, "
        SqlCad = SqlCad & "MED.F7SIGMED, "
        SqlCad = SqlCad & "((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD, "
        SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS SALDO "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
        SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') <> '' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL"
        SqlCad = SqlCad & ") AS INGRESOS "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = 'OC' AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
        SqlCad = SqlCad & "(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) > 0 "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "DET.F3CODPRO"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "f3codpro"
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    wcodcosto = vbNullString
    
    Unload Me
End Sub

