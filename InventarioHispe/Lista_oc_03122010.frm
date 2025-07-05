VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Begin VB.Form lista_oc 
   Caption         =   "Lista de Ordenes de Compra"
   ClientHeight    =   9270
   ClientLeft      =   420
   ClientTop       =   1650
   ClientWidth     =   14175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lista_oc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14175
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CmdSave 
      Left            =   10920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   10005
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   9600
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "Lista_oc.frx":000C
      ToolBars        =   "Lista_oc.frx":8AB5
   End
   Begin TabDock.TTabDock TTabDock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5235
      Left            =   240
      OleObjectBlob   =   "Lista_oc.frx":8C69
      TabIndex        =   2
      Top             =   1440
      Width           =   14775
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar 
      Height          =   675
      Left            =   240
      OleObjectBlob   =   "Lista_oc.frx":11708
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "lista_oc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Af As New ADOFunctions
Dim sql         As String
Dim EditLookUp  As Boolean


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

Dim Fexit As Boolean
Dim ChangechNeed As Boolean

Const PicW = 16
Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type


Private Type Rect
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal HBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wformat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const DT_CALCRECT = &H400


Public Sub LLENADO()
Dim strCodCentros As String
    strCodCentros = VerificaAutorizaciones("OCN", wusuario)
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        If ctipoadm_bd = "M" Then
            sql = "SELECT A.F4NUMORD,A.F4CODSOLICITUD,A.F4ESTNUL,B.F2NOMPROV,A.F4FECEMI,IF(A.F4TIPMON = 'S','Soles','Dólares') AS F4TIPMON,A.F4MONTO,A.F4FECENT FROM IF4ORDEN A, " & _
                  "EF2PROVEEDORES B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4LOCAL='" & wtipo_orden & "' ORDER BY A.F4NUMORD DESC"
             'SQL = "select * from if4orden"
        Else
            sql = "SELECT A.F4NUMORD,A.F4CODSOLICITUD,A.F4ESTNUL,B.F2NOMPROV,A.F4FECEMI,A.F4TIPMON,iif(A.F4TIPMON='S',A.F4MONTO,'') as F4SOLES,iif(A.F4TIPMON='D',A.F4MONTO,'') as F4DOLARES,A.F4FECENT FROM IF4ORDEN AS A, " & _
                  "EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND F4LOCAL='" & wtipo_orden & "' ORDER BY A.F4NUMORD DESC"
                  
            sql = "SELECT CENTROS.F3ABREV, A.F4NUMORD, A.F4CODSOLICITUD, A.F4ESTNUL, B.F2NOMPROV, A.F4FECEMI, A.F4TIPMON, A.F4COLOCADA, A.F4COLOCADAUSER, A.F4COLOCADAFECHA , A.F4ATENDIDA, A.F4ATENDIDAUSER, A.F4ATENDIDAFECHA, "
            sql = sql & "IIf(A.F4TIPMON='S',A.F4MONTO,'') AS F4SOLES, "
            sql = sql & "IIf(A.F4TIPMON='D',A.F4MONTO,'') AS F4DOLARES, "
            sql = sql & "A.F4MONTO,A.F4TIPCAM,A.F4FECENT,A.F4ESTADO,A.F4VB1,A.F4VBUSER1,A.F4VBFECHA1,A.F4VB2,A.F4VBUSER2,A.F4VBFECHA2 "
            sql = sql & "FROM (IF4ORDEN AS A LEFT JOIN EF2PROVEEDORES AS B ON A.F4CODPRV = B.F2NEWRUC) LEFT JOIN CENTROS ON A.F4CENTRO = CENTROS.F3COSTO "
            sql = sql & "WHERE A.F4LOCAL='" & wtipo_orden & "' "
            
            If strCodCentros <> "'999'" Then
                sql = sql & "AND A.F4CENTRO IN (" & strCodCentros & ")"
            End If
                sql = sql & "ORDER BY A.F4NUMORD DESC"
            End If
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "F4NUMORD"
        .Columns.ColumnByFieldName("f4estado").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
    End With

End Sub


Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Numeord = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value

End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
    Dim valor As String
    valor = (dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value)
    proceso2 (valor)
End Sub

Public Sub proceso2(ByVal Codigo As String)
    Dim sql As String
    With Lista_Oc_Detalle.dxDBGrid2
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        'If ctipoadm_bd = "M" Then
        '    sql = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro, a.f3canfal FROM if3orden a lEFT join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= '" & Codigo & "' and f4local = '0' order by a.f3codpro"
        'Else
            If wtipo_orden = "1" Then
                'sql = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro, a.f3canfal FROM if3orden a left join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= '" & Codigo & "' and f4local = '1' order by a.f3codpro"
                sql = "SELECT a.F3CODPRO, a.F3CODFAB, a.F5NOMPRO, EF7MEDIDAS.F7SIGMED as f7codmed, a.F3CANPRO, a.F3CANFAL "
                sql = sql & "FROM if3orden AS a LEFT JOIN EF7MEDIDAS ON a.UNIDAD = EF7MEDIDAS.F7CODMED "
                sql = sql & "WHERE (((a.F4NUMORD)='" & Codigo & "') AND ((a.F4LOCAL)='1')) "
                sql = sql & "ORDER BY a.F3CODPRO"
            Else
                sql = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro, a.f3canfal FROM if3orden a left join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= '" & Codigo & "' and f4local = '0'  order by a.f3codpro"
                
            End If
        'End If
        
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "f3codpro"
    End With
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim RsS As New ADODB.Recordset
Dim StrPregunta As String
Dim SwAprobacion As Boolean
Dim intIndex As Integer
Dim dblMonto As Double
Dim intEstado As Integer
Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
Case "F4VB1", "F4VB2"
    dblMonto = Val(dxDBGrid1.Columns.ColumnByFieldName("f4monto").Value & "")
    If dxDBGrid1.Columns.ColumnByFieldName("f4tipmon").Value & "" = "D" Then
        dblMonto = Val(dxDBGrid1.Columns.ColumnByFieldName("f4monto").Value & "") * Val(dxDBGrid1.Columns.ColumnByFieldName("f4tipcam").Value & "")
    End If
    intIndex = right(dxDBGrid1.Columns.FocusedColumn.FieldName, 1)
    If dxDBGrid1.Columns.FocusedColumn.Value = False Then
            StrPregunta = "¿ Desea aprobar la orden de compra ?"
            SwAprobacion = False
    Else
            StrPregunta = "¿ Desea quitar la aprobación de la orden de compra ?"
            SwAprobacion = True
    End If
    If Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 0 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 1 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 2 Then
        dxDBGrid1.Dataset.Edit
        If MsgBox(StrPregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
            dxDBGrid1.Columns.FocusedColumn.Value = Not SwAprobacion
        Else
            dxDBGrid1.Columns.FocusedColumn.Value = SwAprobacion
        End If
          dxDBGrid1.Dataset.Post
        If dxDBGrid1.Columns.ColumnByFieldName("f4vb" & intIndex).Value = True Then
            If wMonto2doVb >= dblMonto Then
                csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                intEstado = 2
            Else
                If intIndex = 1 Then
                    If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").Value = True Then
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 2
                    Else
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 1
                    End If
                ElseIf intIndex = 2 Then
                    If dxDBGrid1.Columns.ColumnByFieldName("f4vb1").Value = True Then
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 2
                    Else
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 1
                    End If
                End If
            End If
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("f4vbuser" & intIndex).Value = wusuario
            dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = intEstado
            dxDBGrid1.Columns.ColumnByFieldName("f4vbfecha" & intIndex).Value = Now
            dxDBGrid1.Dataset.Post
            
        Else
            If wMonto2doVb >= dblMonto Then
                'csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                'intEstado = 2
                If intIndex = 1 Then
                    If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").Value = True Then
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 2
                    Else
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 1
                    End If
                ElseIf intIndex = 2 Then
                    If dxDBGrid1.Columns.ColumnByFieldName("f4vb1").Value = True Then
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 2
                    Else
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 1
                    End If
                End If
            Else
                'If intIndex = 1 Then
                '    If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").Value = True Then
                '        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                '    Else
                        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                        intEstado = 1
                '    End If
                'ElseIf intIndex = 2 Then
                '    If dxDBGrid1.Columns.ColumnByFieldName("f4vb1").Value = True Then
                '        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                '    Else
                '        csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "' AND F4LOCAL='1'"
                '    End If
                'End If
            End If
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("f4vbuser" & intIndex).Value = ""
            dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = intEstado
            dxDBGrid1.Columns.ColumnByFieldName("f4vbfecha" & intIndex).Value = Null
            dxDBGrid1.Dataset.Post
        End If
        'cnn_dbbancos.Execute csql
        csql = "select cs_estado from tb_cabsolicitud where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
        
            Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
            If RsS.RecordCount > 0 Then
                RsS.MoveFirst
                If intEstado = 2 Then
                    If RsS!CS_ESTADO & "" = "3" Then
                        csql = "update tb_cabsolicitud set cs_estado='4' where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
                        cnn_dbbancos.Execute csql
                    End If
                ElseIf intEstado = 1 Then
                    If RsS!CS_ESTADO & "" = "4" Then
                        csql = "update tb_cabsolicitud set cs_estado='3' where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
                        cnn_dbbancos.Execute csql
                    End If
                End If
            End If
    Else
        If dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value = 3 Then
            MsgBox "La orden de compra ya fue atendida.", vbInformation, App.Title
        ElseIf dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value = 4 Then
            MsgBox "La orden de compra ha sido anulada.", vbInformation, App.Title
        End If
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.FocusedColumn.Value = SwAprobacion
        dxDBGrid1.Dataset.Post
    End If
Case "F4COLOCADA", "F4ATENDIDA"
    'dxDBGrid1.Dataset.Edit
    'dxDBGrid1.Columns.FocusedColumn.Value = Not dxDBGrid1.Columns.FocusedColumn.Value
    'd'xDBGrid1.Dataset.Post
    
    '*******************
    

    Select Case UCase(Column.FieldName)
    Case "F4COLOCADA"
        dxDBGrid1.Dataset.Edit
        If dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Value = True Then
            dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = 2
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Value = False
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAFECHA").Value = Null
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAUSER").Value = ""
        Else
            dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = 4
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Value = True
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAFECHA").Value = Now
            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAUSER").Value = wusuario
        End If
        dxDBGrid1.Dataset.Post
    Case "F4ATENDIDA"
        dxDBGrid1.Dataset.Edit
        If dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").Value = True Then
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").Value = False
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAFECHA").Value = Null
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAUSER").Value = ""
        Else
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").Value = True
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAFECHA").Value = Now
            dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAUSER").Value = wusuario
        End If
        dxDBGrid1.Dataset.Post
    End Select
    
End Select

End Sub

Private Sub dxDBGrid1_OnClick()

Numeord = dxDBGrid1.Columns(0).Value
'Text1.Text = Numeord
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
'   Case "No.O/C", "No.Solicitud", "Proveedor", "Fecha", "Moneda"
   Case "F4SOLES", "F4DOLARES": Text = Format(Text, "#,###,###0.00")
End Select

End Sub

Private Sub dxDBGrid1_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
'   Case "No.O/C", "No.Solicitud", "Proveedor", "Fecha", "Moneda"
   Case "F4SOLES"
        Text = Format(Text, "#,###,###0.00")
        Color = &HC0FFFF
   Case "F4DOLARES"
        Text = Format(Text, "#,###,###0.00")
        Color = &HC0FFC0
End Select

End Sub

Private Sub dxDBGrid1_OnDblClick()

    If dxDBGrid1.Dataset.RecordCount > 0 And UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4VB1" And UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4VB2" And UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4COLOCADA" And UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4ATENDIDA" Then
        sw_nuevo_documento = False
        GOC = dxDBGrid1.Columns(0).Value
        Me.MousePointer = vbHourglass
        If wtipo_orden = "1" Then
            ordendecompra.Show 1
            Unload ordendecompra
            Set ordendecompra = Nothing
        Else
'            loc = 2
            ordencompra_imp.Show 1
        End If
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub dxDBGrid1_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, R As Single, b As Single, NeedShowTip As Boolean)
Dim opValue As Byte
Dim hDC0 As Long, Old_hFont As Long
Dim nc As NONCLIENTMETRICS
Dim rgnR As Rect
Dim StrEstado
Select Case UCase(Column.FieldName)
Case "F4ESTADO"
    NeedShowTip = True
    
    rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
    hDC0 = GetDC(0)
    nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
    Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
    Select Case UCase(Column.FieldName)
    Case "F4ESTADO"
        Select Case Val(TipText & "")
        Case 1: TipText = "Registrando"
        Case 2: TipText = "Aprobado"
        Case 3: TipText = "Atendido"
        Case 4: TipText = "Cerrado"
        Case 5: TipText = "Anulado"
        Case Else: TipText = "No definido"
        End Select
    End Select
    DrawText hDC0, TipText, Len(Trim(TipText)), rgnR, DT_CALCRECT + DT_WORDBREAK

    SelectObject hDC0, Old_hFont
    DeleteObject Old_hFont
    ReleaseDC hwnd, hDC0
    b = t + rgnR.bottom + 6
    R = l + rgnR.right + PicW * 2 + 4
End Select
End Sub

Private Sub Form_Activate()

    'If wtipo_orden = "1" Then
    '    SSActiveToolBars1.Tools.item("ID_NuevoO.CompraMúltiple").Visible = True
    'Else
        dxDBGrid1.Filter.FilterActive = False
    'End If

End Sub

Private Sub Form_Load()
    dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Visible = (VerificaPermiso("0006", wusuario))
    dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Visible = (VerificaPermiso("0007", wusuario))
    dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Visible = (VerificaPermiso("0008", wusuario))
    dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").Visible = (VerificaPermiso("0009", wusuario))
    
    Me.MousePointer = 11
    Me.AutoRedraw = False
'    Me.Height = 7380
    'Me.Width = 10530
    Me.left = 0
    Me.top = 0
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4TIPDOC", "String", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VB1", "YESNO", False, "False"
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBUSER1", "STRING", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBFECHA1", "DATE", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VB2", "YESNO", False, "False"
    
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBUSER2", "STRING", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBFECHA2", "DATE", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ESTADO", "INTEGER", True, 1
    
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADA", "YESNO", False, "False"
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADAUSER", "STRING", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADAFECHA", "DATE", True, ""
    
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDA", "YESNO", False, "False"
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDAUSER", "STRING", True, ""
    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDAFECHA", "DATE", True, ""
    
    LLENADO
    
    'dxDBGrid1.Options.Unset (egoShowGroupPanel)
    'dxDBGrid1.Filter.FilterActive = False
    
    TTabDock1.AddForm Lista_Oc_Detalle, tdDocked, tdAlignBottom, "Lista_Oc_Detalle"
    TTabDock1.DockedForms.ITEM("Lista_Oc_Detalle").Panel.Height = 2500
    TTabDock1.FormShow "Lista_Oc_Detalle"
    Me.MousePointer = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    FraBusqueda.left = 0
    FraBusqueda.top = 0
    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + TTabDock1.DockedForms.ITEM("Lista_Oc_Detalle").Panel.Height)
     
    FraBusqueda.Width = dxDBGrid1.Width
    TxtBusqueda.Width = dxDBGrid1.Width - 350
    
End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo"
            Screen.MousePointer = vbHourglass
            sw_nuevo_documento = True
            If wtipo_orden = "1" Then
                
                ordendecompra.Show 1
                Unload ordendecompra
                Set ordendecompra = Nothing
            Else
'                loc = 2
                ordencompra_imp.Show 1
            End If
        Case "ID_Movimiento":
            imprimir_movoc
            'sw_nuevo_documento = True
            'ocompra_multiples.Show 1
        
       Case "ID_kardex":
        
       'imprimir_Kardex
     
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
        Case "ID_Excel"
            Me.MousePointer = 11
            GENERA_EXCEL
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    End Select

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

Private Sub TTabDock1_PanelResize(ByVal Panel As TabDock.TTabDockHost)
Form_Resize
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2NOMPROV LIKE '*" & TxtBusqueda.Text & "*' OR " & " f4numord LIKE '*" & TxtBusqueda.Text & "*' OR f3abrev LIKE '*" & TxtBusqueda.Text & "*'  OR F4codsolicitud LIKE '*" & TxtBusqueda.Text & "*' "
    
    If Len(Trim(TxtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    dxDBGrid1.SetFocus
    End If
End Sub




Private Sub GENERA_EXCEL()

    IValue = 255
    GridNum = 1: OldValue = 1
    GridInit IValue - 10, OldValue
    OldValue = IValue
            
End Sub

Private Function GetGridByActive() As dxDBGrid

    Set GetGridByActive = dxDBGrid1
    
End Function


Private Sub SaveTo(Index)
On Error GoTo errhandler
Dim FileName As String
Dim o_Excel As Excel.Application

'    If GridNum <> 0 Then
        With cmdSave
            .CancelError = True
            .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
            .DialogTitle = wnomcia
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
                    Set o_Excel = CreateObject("Excel.application")
                    o_Excel.Workbooks.Open FileName:=.FileName
                    o_Excel.Visible = True
                    If Not o_Excel Is Nothing Then
                        Set o_Excel = Nothing
                    End If
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
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "CONTROL Plus!") = vbYes Then _
                    GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
'    End If
    Exit Sub
    
errhandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
    
End Sub

Private Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim I As Byte

    If Ind > 199 Then   'ind=245
        SaveTo (Ind)
        Exit Sub
    End If

End Sub

Private Sub imprimir_movoc()
'Dim rsrep       As New ADODB.Recordset
Dim csql        As String
'Dim Cad         As String
Dim I           As Integer

    acr_movoc.fldempresa.Text = wnomcia
    acr_movoc.fldfecha.Text = Format(Date, "dd/mm/yyyy")
    acr_movoc.Field14.Visible = True
    acr_movoc.lblnumorden.Visible = True
    acr_movoc.fldnumorden.Visible = True
    acr_movoc.fldnomproveedores.Visible = True
    acr_movoc.lblproveedores.Visible = True
'
'    Cad = "SELECT IF4ORDEN.F4NUMORD FROM ((IF4ORDEN INNER JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) INNER JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC WHERE ((PAG_DCTO.NRO_COMP='" & Grid.Columns(2).Value & "') AND (([if4orden].[f4monto]-[regisdoc].[f4total])<=-0.02 Or ([if4orden].[f4monto]-[regisdoc].[f4total])>=0.02)) AND ((PAG_DCTO.CORRELA)=" & Grid.Columns(0).Value & ")"
'    Set rsrep = Af.OpenSQLForwardOnly(Cad, StrConexDbBancos)
'    If rsrep.RecordCount = 0 Then
'        MsgBox "No hay registros para Mostrar.", vbExclamation, "Sistema de Bancos"
'        rsrep.Close
'        Set rsrep = Nothing
'        Exit Sub
'    End If
    ''I = 0
''''''''''    csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA as FH, iif(REGISDOC.F4MONEDA='S', 'S/.', '$') as F4MONEDA, (REGISDOC.F4TOTAL)* -1 as F4TOTAL, PAG_DCTO.saldo , ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, IF4ORDEN.F4MONTO - PAG_DCTO.TOTAL AS TOTAL  "          ', IF4ORDEN.F4MONTO - Sum(REGISDOC.F4TOTAL) as Saldo2 "
''''''''''    'csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA as FH, iif(REGISDOC.F4MONEDA='S', 'S/.', '$') as F4MONEDA, (REGISDOC.F4TOTAL)* -1 as F4TOTAL, PAG_DCTO.saldo , ((REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, IF4ORDEN.F4MONTO - PAG_DCTO.TOTAL AS TOTAL  "          ', IF4ORDEN.F4MONTO - Sum(REGISDOC.F4TOTAL) as Saldo2 "
''''''''''    csql = csql + "FROM ((IF4ORDEN INNER JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) INNER JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC "
''''''''''    csql = csql + "WHERE ((IF4ORDEN.F4NUMORD) = '" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "')"
''''''''''''    csql = csql + " AND (([if4orden].[f4monto]-[regisdoc].[f4total])<=-0.02 Or ([if4orden].[f4monto]-[regisdoc].[f4total])>=0.02) "
''''''''''    'csql = csql + " Group By IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA, REGISDOC.F4MONEDA, REGISDOC.F4TOTAL, PAG_DCTO.saldo, (REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO * -1 "

    csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, "
    csql = csql + "IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA AS FH, IIf(REGISDOC.F4MONEDA='S','S/.','$') AS F4MONEDA, (REGISDOC.F4TOTAL)*-1  "
    csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, (PAG_DCTO.TOTAL - PAG_DCTO.saldo)*-1 AS TOTAL, 'Factura' as Concepto  "
    'csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, IF4ORDEN.F4MONTO-PAG_DCTO.TOTAL AS TOTAL "
    csql = csql + " FROM ((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) "
    csql = csql + " LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC "
    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "')) "
    csql = csql + " union ALL"
    csql = csql + " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV,  "
    csql = csql + " IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, PAG_DCTO.FCH_COMP AS FH, IIf(PAG_DCTO.MONEDA='S','S/.','$') AS F4MONEDA, (IF4ORDEN.F4MONTO-REGISDOC.F4TOTAL)*-1 AS F4TOTAL, "
    csql = csql + " PAG_DCTO.saldo, ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, PAG_DCTO.TOTAL*-1, 'Pago a Cuenta' as Concepto "
    csql = csql + " FROM ((IF4ORDEN LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) LEFT JOIN PAG_DCTO ON  "
    csql = csql + " IF4ORDEN.F4NUMORD = PAG_DCTO.F4OCOMPRA) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA "
    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "') AND ((Mid([pag_dcto].[nro_comp],1,3))='Ant')) "
    acr_movoc.datconexion.ConnectionString = cnn_dbbancos
    acr_movoc.datconexion.Source = csql
    Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    If RsS.RecordCount > 0 Then
        acr_movoc.Show vbModal
    Else
        MsgBox "No hay registros"
    End If
''    rsrep.Close
''    Set rsrep = Nothing


'Dim csql        As String
'Dim cad         As String
'Dim I           As Integer
'
'    acr_movoc.fldempresa.Text = wnomcia
'    acr_movoc.fldfecha.Text = Format(Date, "dd/mm/yyyy")
'    acr_movoc.lblalmacen.Visible = True
'    acr_movoc.fldcodalmacen.Visible = True
'    acr_movoc.fldnomalmacen.Visible = True
'
'    csql = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4NUMDOC, REGISDOC.F4FECHA, REGISDOC.F4MONEDA, REGISDOC.F4TOTAL, PAG_DCTO.saldo "
'    csql = csql + "FROM ((IF4ORDEN INNER JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) INNER JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC "
'    csql = csql + "WHERE ((IF4ORDEN.F4NUMORD) = '" & Trim(Numeord) & "')"
'    'csql = csql + " AND (([if4orden].[f4monto]-[regisdoc].[f4total])<=-0.02 Or ([if4orden].[f4monto]-[regisdoc].[f4total])>=0.02)"
'    acr_movoc.datconexion.ConnectionString = cnn_dbbancos
'    acr_movoc.datconexion.Source = csql
'    Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
'    If RsS.RecordCount > 0 Then
'        acr_movoc.Show vbModal
'    Else
'        MsgBox "No hay registros"
'    End If
      
End Sub



