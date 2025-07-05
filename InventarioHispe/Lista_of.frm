VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form lista_of 
   Caption         =   "Lista de Ordenes de Compra"
   ClientHeight    =   6480
   ClientLeft      =   540
   ClientTop       =   1860
   ClientWidth     =   15285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lista_of.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlgOrden 
      Left            =   10920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10005
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   9600
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   11
      ShowShortcutsInToolTips=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "Lista_of.frx":000C
      ToolBars        =   "Lista_of.frx":A411
   End
   Begin TabDock.TTabDock TTabDock1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5235
      Left            =   120
      OleObjectBlob   =   "Lista_of.frx":A645
      TabIndex        =   2
      Top             =   1080
      Width           =   14775
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar 
      Height          =   555
      Left            =   0
      OleObjectBlob   =   "Lista_of.frx":152DA
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   840
   End
   Begin MSComctlLib.ImageList imgLstEstado 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1AA94
            Key             =   "Estado 1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1B02E
            Key             =   "Estado 2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1B5C8
            Key             =   "Estado 3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1BB62
            Key             =   "Estado 4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1C0FC
            Key             =   "Estado 5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1C696
            Key             =   "Estado 6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1CC30
            Key             =   "Estado 7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lista_of.frx":1D1CA
            Key             =   "Estado 8"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "lista_of"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cconex_form As String
Dim RsListaOC As New ADODB.Recordset
Dim amovs_cab(0 To 4)  As a_grabacion
'Option Explicit
Dim Af As New ADOFunctions
Dim sql         As String

Dim EditLookUp  As Boolean


'        lfItalic As Byte
'        lfUnderline As Byte
'        lfStrikeOut As Byte
'        lfCharSet As Byte
'        lfOutPrecision As Byte
'        lfClipPrecision As Byte
'        lfQuality As Byte
'        lfPitchAndFamily As Byte
'        lfFaceName(1 To LF_FACESIZE) As Byte
'End Type
'
'Private Type NONCLIENTMETRICS
'        cbSize As Long
'        iBorderWidth As Long
'        iScrollWidth As Long
'        iScrollHeight As Long
'        iCaptionWidth As Long
'        iCaptionHeight As Long
'        lfCaptionFont As LOGFONT
'        iSMCaptionWidth As Long
'        iSMCaptionHeight As Long
'        lfSMCaptionFont As LOGFONT
'        iMenuWidth As Long
'        iMenuHeight As Long
'        lfMenuFont As LOGFONT
'        lfStatusFont As LOGFONT
'        lfMessageFont As LOGFONT
'End Type
'
'
'Private Type Rect
'        left As Long
'        top As Long
'        right As Long
'        bottom As Long
'End Type
'
'Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal HBrush As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wformat As Long) As Long
'Private Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long
'
'Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
'Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
'Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
'
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'
'
'Private Const SPI_GETNONCLIENTMETRICS = 41
'Private Const DT_CALCRECT = &H400


Public Sub LLENADO()
    Dim strCodCentros As String
    
    strCodCentros = VerificaAutorizaciones("OCN", wusuario)
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        If ctipoadm_bd = "M" Then
            sql = "SELECT A.F4NUMORD,A.F4CODSOLICITUD,A.F4ESTNUL,B.F2NOMPROV,A.F4FECEMI,IF(A.F4TIPMON = 'S','Soles','Dólares') AS F4TIPMON,A.F4MONTO,A.F4FECENT FROM IF4ORDEN A, " & _
                  "EF2PROVEEDORES B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4LOCAL='" & wtipo_orden & "' ORDER BY A.F4NUMORD DESC"
        Else
            'sql = "SELECT A.F4LOCAL,A.F4NUMORD,A.F4CODSOLICITUD,A.F4ESTNUL,B.F2NOMPROV,A.F4FECEMI,A.F4TIPMON,iif(A.F4TIPMON='S',A.F4MONTO,'') as F4SOLES,iif(A.F4TIPMON='D',A.F4MONTO,'') as F4DOLARES,A.F4FECENT FROM IF4ORDEN AS A, " & _
                  "EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND F4LOCAL='" & wtipo_orden & "' ORDER BY A.F4NUMORD DESC"
            
            sql = vbNullString
            sql = sql & "SELECT "
            sql = sql & "A.F4LOCAL, "
            sql = sql & "CENTROS.F3ABREV, "
            sql = sql & "A.F4NUMORD, "
            sql = sql & "A.F4CODSOLICITUD, "
            sql = sql & "A.F4OBSERVA AS F3DESCRIP, "
            sql = sql & "A.F4ESTNUL, "
            sql = sql & "B.F2NOMPROV, "
            sql = sql & "A.F4FECEMI, "
            sql = sql & "A.F4TIPMON, "
            sql = sql & "A.F4COLOCADA, "
            sql = sql & "A.F4COLOCADAUSER, "
            sql = sql & "A.F4COLOCADAFECHA , "
            sql = sql & "A.F4ATENDIDA, "
            sql = sql & "A.F4ATENDIDAUSER, "
            sql = sql & "A.F4ATENDIDAFECHA, "
            sql = sql & "IIf(A.F4TIPMON='S',A.F4MONTO,'') AS F4SOLES, "
            sql = sql & "IIf(A.F4TIPMON='D',A.F4MONTO,'') AS F4DOLARES, "
            sql = sql & "A.F4MONTO, "
            sql = sql & "A.F4TIPCAM, "
            sql = sql & "A.F4FECENT, "
            sql = sql & "A.F4ESTADO, "
            sql = sql & "A.F4VB1, "
            sql = sql & "A.F4VBUSER1, "
            sql = sql & "A.F4VBFECHA1, "
            sql = sql & "A.F4VB2, "
            sql = sql & "A.F4VBUSER2, "
            sql = sql & "A.F4VBFECHA2 "
            'sql = sql & "IF4VALES.F2CODALM & IF4VALES.F4NUMVAL AS ALMACEN, "
            'sql = sql & "REGISDOC.F4NUMDOC AS FACTURA "
            sql = sql & "FROM "
            sql = sql & "(IF4ORDEN AS A "
            sql = sql & "LEFT JOIN EF2PROVEEDORES AS B ON A.F4CODPRV = B.F2NEWRUC) "
            sql = sql & "LEFT JOIN CENTROS ON A.F4CENTRO = CENTROS.F3COSTO "
            'sql = sql & "LEFT JOIN IF4VALES ON A.F4NUMORD = IF4VALES.NUMORDEN) "
            'sql = sql & "LEFT JOIN REGISDOC ON A.F4NUMORD = REGISDOC.F4OCOMPRA "
        End If
        
        If Len(Trim(wObra)) > 0 Then
            sql = sql & "where A.F4CENTRO IN ('" & wObra & "')"
        End If
        sql = sql & "ORDER BY A.F4FECEMI DESC"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "F4NUMORD"
        .Columns.ColumnByFieldName("f4estado").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
    End With
End Sub


Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    'Numeord = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value
    'proceso2 Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").Value & ""), _
                Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "")
End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
'    Dim valor As String, tipoor As String
'
'    If dxDBGrid1.Dataset.RecordCount > 1 Then 'filtrox = 0 And
'        valor = (dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value)
'        tipoor = dxDBGrid1.Columns.ColumnByFieldName("f4local").Value
'        Call proceso2(valor, tipoor)
'    End If
'    filtrox = 1
    If dxDBGrid1.Dataset.RecordCount > 1 Then
        proceso2 Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value & ""), _
                    Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
    End If
    
    filtrox = 1
End Sub

Public Sub proceso2(ByVal strTipoDocumento As String, ByVal strNumOrden As String)
    'Dim sql As String
    
    With Lista_Oc_Detalle.dxDBGrid2
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        sql = vbNullString
        
        If wtipo_orden = "1" Then
            
            sql = sql & "SELECT "
            sql = sql & "IIF(DET.COD_SOLICITUD <> '', DET.COD_SOLICITUD, 'STOCK LIBRE') AS COD_SOLICITUD, "
            sql = sql & "DET.F3CODPRO, "
            sql = sql & "DET.F3CODFAB, "
            sql = sql & "DET.F5NOMPRO_ING AS F5NOMPRO, "
            sql = sql & "MED.F7SIGMED, "
            sql = sql & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) AS F3CANPRO, "
            sql = sql & "(DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(INGRESOS.CANTIDAD & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS F3CANFAL, "
            sql = sql & "DET.F3PRENETO AS F3PRECOS "
            sql = sql & "FROM "
            sql = sql & "((IF3ORDEN AS DET "
            sql = sql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
            sql = sql & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
            sql = sql & "LEFT JOIN "
            sql = sql & "(SELECT "
            sql = sql & "TRIM(DET.F4NUMORD & '') AS F4NUMORD, "
            sql = sql & "DET.COD_SOLICITUD AS COD_SOLICITUD, "
            sql = sql & "TRIM(DET.F5CODPROORIGINAL & '') AS F5CODPROORIGINAL, "
            
            sql = sql & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
            'sql = sql & "(DET.F3CANPRO / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) AS CANTIDAD "
            
            sql = sql & "FROM "
            'sql = sql & "((IF3VALES AS DET "
            sql = sql & "IF3VALES AS DET "
            sql = sql & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
            'sql = sql & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPROORIGINAL) "
            'sql = sql & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = PROD.F5CODPRO AND MEDALTER.F7CODMED = PROD.F7CODMED "
            sql = sql & "WHERE "
            sql = sql & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
            sql = sql & "TRIM(DET.F4NUMORD & '') = '" & strNumOrden & "' "
            sql = sql & "GROUP BY "
            sql = sql & "TRIM(DET.F4NUMORD & ''), "
            sql = sql & "DET.COD_SOLICITUD, "
            sql = sql & "TRIM(DET.F5CODPROORIGINAL & '')) AS INGRESOS "
            
            sql = sql & "ON TRIM(INGRESOS.F4NUMORD & '') = TRIM(DET.F4NUMORD & '') AND TRIM(INGRESOS.COD_SOLICITUD & '') = DET.COD_SOLICITUD AND TRIM(INGRESOS.F5CODPROORIGINAL & '') = TRIM(DET.F3CODPRO & '') "
            sql = sql & "WHERE "
            sql = sql & "DET.F4LOCAL = '" & strTipoDocumento & "' AND "
            sql = sql & "DET.F4NUMORD = '" & strNumOrden & "' "
            sql = sql & "ORDER BY "
            sql = sql & "DET.F5NOMPRO_ING"
        Else
            sql = sql & "SELECT "
            sql = sql & "A.COD_SOLICITUD, "
            sql = sql & "DET.F3CODPRO, "
            sql = sql & "DET.F3CODFAB, "
            sql = sql & "DET.F5NOMPRO, "
            sql = sql & "MED.F7SIGMED, "
            sql = sql & "DET.F3CANPRO, "
            sql = sql & "DET.F3CANFAL, "
            sql = sql & "DET.F3PRENETO AS F3PRECOS "
            sql = sql & "FROM "
            sql = sql & "(IF3ORDEN AS DET "
            sql = sql & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F3CODPRO) "
            sql = sql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED "
            sql = sql & "WHERE "
            sql = sql & "DET.F4LOCAL = '" & TOC & "' AND "
            sql = sql & "DET.F4NUMORD = '" & strNumOrden & "' "
            sql = sql & "ORDER BY "
            sql = sql & "DET.F3CODPRO"
        End If
        
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "f3codpro"
    End With
    
    sql = vbNullString
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Dim RsS As New ADODB.Recordset
    Dim StrPregunta As String
    Dim SwAprobacion As Boolean
    Dim intIndex As Integer
    Dim dblMonto As Double
    Dim intEstado As Integer
    Dim variable As String
    Dim A2 As String
    Dim A1 As Boolean
    
    Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
        Case "F4VB1", "F4VB2"
            '    est1 = traerCampo("IF4ORDEN_PAGO", "TOP 1 EST_AUT", "ORDEN", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
            '    If Val(est1) >= 1 Then
            ''        dxDBGrid1.Dataset.Edit
            ''        dxDBGrid1.Columns.ColumnByFieldName("f4vb1").Value = False
            ''        dxDBGrid1.Columns.ColumnByFieldName("f4vb2").Value = False
            ''        dxDBGrid1.Dataset.Post
            '
            '        If State = cbsChecked Then
            '            A2 = traerCampo("IF4ORDEN_PAGO", "Top 1 Estado", "ORDEN", dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value, " Order By IDOP DESC")
            '            If A2 = "0" Then
            '                    sw_est_orden_pago = False
            '                    sw_ordendepago = True
            '                    frmOrdPago.Show 1
            '            Else
            '                sw_est_orden_pago = False
            '                sw_ordendepago = False
            '                frmOrdPago.Show 1
            '            End If
            '        End If
            '
            '    Else
                
''            dblMonto = Val(dxDBGrid1.Columns.ColumnByFieldName("f4monto").Value & "")
''
''            If dxDBGrid1.Columns.ColumnByFieldName("f4tipmon").Value & "" = "D" Then
''                dblMonto = Val(dxDBGrid1.Columns.ColumnByFieldName("f4monto").Value & "") * Val(dxDBGrid1.Columns.ColumnByFieldName("f4tipcam").Value & "")
''            End If
''
            intIndex = right(dxDBGrid1.Columns.FocusedColumn.FieldName, 1)
            
            If dxDBGrid1.Columns.FocusedColumn.value = False Then
                StrPregunta = "¿ Desea aprobar la orden de compra ?"
                SwAprobacion = False
            Else
                StrPregunta = "¿ Desea quitar la aprobación de la orden de compra ?"
                SwAprobacion = True
            End If
            
            If Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value & "") = 0 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value & "") = 1 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value & "") = 2 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value & "") = 3 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value & "") = 6 Then
            'If Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 0 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 1 Or Val(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Value & "") = 2 Then
                dxDBGrid1.Dataset.Edit
                
                If dxDBGrid1.Columns.ColumnByFieldName("F4VB2").value = False Then
                    'dxDBGrid1.Columns.FocusedColumn.Value = Not SwAprobacion
                Else
                    If MsgBox(StrPregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
            '            sw_ordendepago = False
            '            If State = cbsChecked Then
            '                frmOrdPago.Show
            '            End If
                        dxDBGrid1.Columns.FocusedColumn.value = Not SwAprobacion
                    Else
                        dxDBGrid1.Columns.FocusedColumn.value = SwAprobacion
                        If SwAprobacion = False Then Exit Sub
                    End If
                End If
                
''                If State = cbsChecked Then
''                    A2 = traerCampo("IF4ORDEN_PAGO", "Top 1 Estado", "ORDEN", dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value, " Order By IDOP DESC")
''
''                    If A2 = "0" Then
''                            sw_e_ordenpago = False 'Cuando quiero Eliminar
''                            sw_est_orden_pago = False
''                            sw_ordendepago = True
''                            frmOrdPago.Show 1
''                    Else
''                        sw_e_ordenpago = False 'Cuando quiero Eliminar
''                        sw_est_orden_pago = False
''                        sw_ordendepago = False
''                        frmOrdPago.Show 1
''                    End If
''                End If
                'If A2 = "0" Then
                dxDBGrid1.Dataset.Post
                'End If
                
                If dxDBGrid1.Columns.ColumnByFieldName("f4vb" & intIndex).value = True Then
                    If intIndex = 1 Then
                        variable = 2
                    Else
                        variable = 1
                    End If
                  
                    A1 = traerCampo("IF4ORDEN", "f4vb" & variable, "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value)
                    
                    If A1 = True Then
                        If wMonto2doVb >= dblMonto Then
                            csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND F4LOCAL='1'"
                            intEstado = 2
                        Else
                            If intIndex = 1 Then
                                If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").value = True Then
                                    csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND F4LOCAL='1'"
                                    intEstado = 2
                                Else
                                    csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "' AND F4LOCAL='1'"
                                    intEstado = 1
                                End If
                            ElseIf intIndex = 2 Then
                                If dxDBGrid1.Columns.ColumnByFieldName("f4vb1").value = True Then
                                    csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND F4LOCAL='1'"
                                    intEstado = 2
                                Else
                                    csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "' AND F4LOCAL='1'"
                                    intEstado = 1
                                End If
                            End If
                        End If
                                                    
                        dxDBGrid1.Dataset.Edit
                        dxDBGrid1.Columns.ColumnByFieldName("f4vbuser" & intIndex).value = wusuario
                        dxDBGrid1.Columns.ColumnByFieldName("f4estado").value = intEstado
                        dxDBGrid1.Columns.ColumnByFieldName("f4vbfecha" & intIndex).value = Now
                        dxDBGrid1.Dataset.Post
                    End If
                Else
                    If wMonto2doVb >= dblMonto Then
                        'csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        'intEstado = 2
                        If intIndex = 1 Then
                            If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").value = True And SwAprobacion = False Then
                                csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND F4LOCAL='1'"
                                intEstado = 2
                            Else
                                csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "' AND F4LOCAL='1'"
                                intEstado = 1
                            End If
                        ElseIf intIndex = 2 And SwAprobacion = False Then
                            If dxDBGrid1.Columns.ColumnByFieldName("f4vb1").value = True Then
                                csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND F4LOCAL='1'"
                                intEstado = 2
                            Else
                                csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "' AND F4LOCAL='1'"
                                intEstado = 1
                            End If
                        End If
                    Else
                        'If intIndex = 1 Then
                        '    If dxDBGrid1.Columns.ColumnByFieldName("f4vb2").Value = True Then
                        '        csql = "UPDATE IF4ORDEN SET F4ESTADO=2 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").Value & "' AND F4LOCAL='1'"
                        '    Else
                                csql = "UPDATE IF4ORDEN SET F4ESTADO=1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "' AND F4LOCAL='1'"
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
                    
                    If intEstado = 0 Then intEstado = 1
                    
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("f4vbuser" & intIndex).value = ""
                    dxDBGrid1.Columns.ColumnByFieldName("f4estado").value = intEstado
                    dxDBGrid1.Columns.ColumnByFieldName("f4vbfecha" & intIndex).value = Null
                    dxDBGrid1.Dataset.Post
                End If
                'cnn_dbbancos.Execute csql
                '        csql = "select cs_estado from tb_cabsolicitud where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
                '
                '            Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
                '            If RsS.RecordCount > 0 Then
                '                RsS.MoveFirst
                '                If intEstado = 2 Then
                '                    If RsS!CS_ESTADO & "" = "3" Then
                '                        csql = "update tb_cabsolicitud set cs_estado='4' where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
                '                        cnn_dbbancos.Execute csql
                '                    End If
                '                ElseIf intEstado = 1 Then
                '                    If RsS!CS_ESTADO & "" = "4" Then
                '                        csql = "update tb_cabsolicitud set cs_estado='3' where cod_solicitud='" & dxDBGrid1.Columns.ColumnByFieldName("f4codsolicitud").Value & "'"
                '                        cnn_dbbancos.Execute csql
                '                    End If
                '                End If
                '            End If
            Else
                If dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value = 3 Then
                    MsgBox "La orden de compra ya fue atendida.", vbInformation, App.Title
                ElseIf dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").value = 4 Then
                    MsgBox "La orden de compra ha sido anulada.", vbInformation, App.Title
                End If
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.FocusedColumn.value = SwAprobacion
                dxDBGrid1.Dataset.Post
            End If
            'End If
        Case "F4COLOCADA", "F4ATENDIDA"
            'dxDBGrid1.Dataset.Edit
            'dxDBGrid1.Columns.FocusedColumn.Value = Not dxDBGrid1.Columns.FocusedColumn.Value
            'd'xDBGrid1.Dataset.Post
            
            '*******************
                    
            Select Case UCase(Column.FieldName)
                Case "F4COLOCADA"
                    dxDBGrid1.Dataset.Edit
                    If dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").value = True Then
                        If VerificaPermiso("0011", wusuario) Then
                            If MsgBox("¿Desea eliminar la Autorización de Registro/Pago para la Orden de Compra Nº " & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value, 4 + 32, "ATENCIÓN") = vbYes Then
                                csql = "DELETE * FROM IF4ORDEN_PAGO WHERE ORDEN ='" & dxDBGrid1.Columns.ColumnByFieldName("f4NUMORD").value & "' AND correladoc = 0 AND correlaanticipo = 0"
                                cnn_dbbancos.Execute csql
                                dxDBGrid1.Columns.ColumnByFieldName("f4estado").value = 1
                                dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").value = False
                                dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAFECHA").value = Null
                                dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAUSER").value = ""
                            Else
                                Exit Sub
                            End If
                        Else
                            MsgBox "No cuenta con autorización para revertir la aprobación de la Orden de Compra", vbCritical, "ATENCIÓN"
                        End If
                    Else
                        'If dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = 2 Then
                            dxDBGrid1.Columns.ColumnByFieldName("f4estado").value = 4
                            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").value = True
                            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAFECHA").value = Now
                            dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADAUSER").value = wusuario
                            frmOrdPago.Show 1
                        'Else
                        '    MsgBox "La Orden de Compra no está Aprobada"
                        '    dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Value = False
                        '    dxDBGrid1.Dataset.Post
                        '    Exit Sub
                        'End If
                    End If
                    
                    dxDBGrid1.Dataset.Post
                Case "F4ATENDIDA"
                    dxDBGrid1.Dataset.Edit
                    
                    If dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").value = True Then
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").value = False
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAFECHA").value = Null
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAUSER").value = ""
                    Else
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").value = True
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAFECHA").value = Now
                        dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDAUSER").value = wusuario
                    End If
                    
                    filtrox = 1
                    dxDBGrid1.Dataset.Post
                    filtrox = 0
            End Select
    End Select
End Sub

Private Sub Llena_Orden_de_Pago()
sw_est_orden_pago = True
frmOrdPago.Show
End Sub

Private Sub dxDBGrid1_OnClick()
    'Dim valor As String, tipoor As String
    'Numeord = dxDBGrid1.Columns(0).Value
    filtrox = 0
    'valor = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "")
    'tipoor = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").Value & "")
    
    Me.MousePointer = vbHourglass
    
    proceso2 Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value & ""), _
                Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
                
    Me.MousePointer = vbDefault
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
    If dxDBGrid1.Dataset.RecordCount > 0 And UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4VB1" And _
        UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4VB2" And _
        UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4COLOCADA" And _
        UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) <> "F4ATENDIDA" Then
        
        sw_nuevo_documento = False
        
        GOC = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value
        TOC = dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value
        GSOL = dxDBGrid1.Columns.ColumnByFieldName("F4CODSOLICITUD").value
        
        Me.MousePointer = vbHourglass
        
        'If wtipo_orden = "1" Then
        '    ordendecompra.Show 1
        '    Unload ordendecompra
        '    Set ordendecompra = Nothing
        'Else
'       '     loc = 2
        '    ordencompra_imp.Show 1
        'End If
        
        With ordendecompra
            .TipoOrden = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value & "")
            .NumeroOrden = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
            
            .Show 1
        End With
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dxDBGrid1_OnDblClick
    End Select
End Sub

Private Sub dxDBGrid1_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, R As Single, b As Single, NeedShowTip As Boolean)
    Dim opValue As Byte
    Dim hDC0 As Long, Old_hFont As Long
    'Dim nc As NONCLIENTMETRICS
    Dim rgnR As Rect
    Dim StrEstado
    
    Select Case UCase(Column.FieldName)
        Case "F4NUMORD"
'            NeedShowTip = True
'
'            rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
'            hDC0 = GetDC(0)
'            nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
'            SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
'            Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
'
'            TipText = dxDBGrid1.Columns.ColumnByFieldName("F4OBSERVA").value
'
'            DrawText hDC0, TipText, Len(Trim(TipText)), rgnR, DT_CALCRECT + DT_WORDBREAK
'
'            SelectObject hDC0, Old_hFont
'            DeleteObject Old_hFont
'            ReleaseDC HWnd, hDC0
'            b = t + rgnR.bottom + 6
'            R = l + rgnR.right + PicW * 2 + 4
        
        Case "F4ESTADO"
            NeedShowTip = True
'
'            rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
'            hDC0 = GetDC(0)
'            nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
'            SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
'            Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
'            Select Case UCase(Column.FieldName)
'                Case "F4ESTADO"
'                    Select Case Val(TipText & "")
'                        Case 1: TipText = "Registrando"
'                        Case 2: TipText = "Aprobado"
'                        Case 3: TipText = "Atendido"
'                        Case 4: TipText = "Cerrado"
'                        Case 5: TipText = "Anulado"
'                        Case 6: TipText = "Pagado"
'                        Case 7: TipText = "Anticipado"
'                        Case Else: TipText = "No definido"
'                    End Select
'            End Select
'            DrawText hDC0, TipText, Len(Trim(TipText)), rgnR, DT_CALCRECT + DT_WORDBREAK
'
'            SelectObject hDC0, Old_hFont
'            DeleteObject Old_hFont
'            ReleaseDC HWnd, hDC0
'            b = t + rgnR.bottom + 6
'            R = l + rgnR.right + PicW * 2 + 4

                    Select Case Val(TipText & "")
                        Case 1
                            TipText = "Orden en Edición"
                        Case 2
                            TipText = "Orden Aprobada"
                        Case 3
                            TipText = "Orden Enviada"
                        Case 4
                            TipText = "Orden Recepcionada"
                        Case 5
                            TipText = "Atención Parcial"
                        Case 6
                            TipText = "Atención Total"
                        Case 7
                            TipText = "Orden Cerrada"
                        Case 8
                            TipText = "Orden Anulada"
                    End Select
    End Select
End Sub

Private Sub Form_Activate()

    'If wtipo_orden = "1" Then
    '    SSActiveToolBars1.Tools.item("ID_NuevoO.CompraMúltiple").Visible = True
    'Else
        dxDBGrid1.Filter.FilterActive = False
        filtrox = 0
    'End If
    dxDBGrid1_OnChangeNodeEx
End Sub

Private Sub Form_Load()
    dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Visible = (VerificaPermiso("0006", wusuario))
    dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Visible = False '(VerificaPermiso("0007", wusuario))
    'dxDBGrid1.Columns.ColumnByFieldName("F4COLOCADA").Visible = (VerificaPermiso("0008", wusuario))
    'dxDBGrid1.Columns.ColumnByFieldName("F4ATENDIDA").Visible = False '(VerificaPermiso("0009", wusuario))
''''    cnombase = "TEMPLUS.MDB"
''''    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
''''    cnn_form.Open cconex_form

    wTipoReq = 0: TOC = ""
    Me.MousePointer = vbHourglass
    Me.AutoRedraw = False
'    Me.Height = 7380
    'Me.Width = 10530
    Me.left = 0
    Me.top = 0
    sw_ordendepago = False
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    
    
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4TIPDOC", "String", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VB1", "YESNO", False, "False"
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBUSER1", "STRING", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBFECHA1", "DATE", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VB2", "YESNO", False, "False"
'
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBUSER2", "STRING", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4VBFECHA2", "DATE", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ESTADO", "INTEGER", True, 1
'
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADA", "YESNO", False, "False"
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADAUSER", "STRING", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4COLOCADAFECHA", "DATE", True, ""
'
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDA", "YESNO", False, "False"
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDAUSER", "STRING", True, ""
'    Crea_Campo cconex_dbbancos, "IF4ORDEN", "F4ATENDIDAFECHA", "DATE", True, ""
    
    LLENADO
    
    dxDBGrid1_OnClick
    
    'dxDBGrid1.Options.Unset (egoShowGroupPanel)
    'dxDBGrid1.Filter.FilterActive = False
    
    
    
    TTabDock1.AddForm Lista_Oc_Detalle, tdDocked, tdAlignBottom, "Lista_Oc_Detalle"
    TTabDock1.DockedForms.ITEM("Lista_Oc_Detalle").Panel.Height = 2500
    TTabDock1.FormShow "Lista_Oc_Detalle"
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FraBusqueda.left = 0
    FraBusqueda.top = 0
    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + TTabDock1.DockedForms.ITEM("Lista_Oc_Detalle").Panel.Height)
     
    FraBusqueda.Width = dxDBGrid1.Width
    txtBusqueda.Width = dxDBGrid1.Width - 350
    
End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim EstOC As String
    Select Case Tool.Id
        Case "ID_Nuevo"
            'Screen.MousePointer = vbHourglass
            
            sw_nuevo_documento = True
            wTipoOC = 1
            TOC = "OC"
            
            If ModUtilitario.validarFormAbierto("ordendecompra") Then
                Unload ordendecompra
            End If
            
            With ordendecompra
                .TipoOrden = "OC"
                .NumeroOrden = vbNullString
                
                .Show 1
            End With
        Case "ID_NuevaOS"
            Screen.MousePointer = vbHourglass
            
            sw_nuevo_documento = True
            
            wTipoOC = 2
            TOC = "OS"
            
            If ModUtilitario.validarFormAbierto("ordendecompra") Then
                Unload ordendecompra
            End If
            
            With ordendecompra
                .TipoOrden = "OS"
                .NumeroOrden = vbNullString
                
                .Show 1
            End With
        Case "ID_AutorizarPago"
            MsgBox "Opción no disponible.", vbInformation + vbOKOnly, App.ProductName
            
'            If ValidaSaldoOrden > 0 Then
'                EstOC = traerCampo("IF4ORDEN", "F4ESTADO", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
'                If EstOC = "2" Or EstOC = "1" Or EstOC = "4" Then
'                      MsgBox "No se puede realizar una nueva autorización porque la Orden de Compra aun no ha sido atendida", vbCritical, "ATENCIÓN"
'                      est1 = "0"
'                      Exit Sub
'                Else
'                   If MsgBox("¿Desea realizar una siguiente Autorización sobre un Saldo de " & ValidaSaldoOrden & " para la Orden de Compra Nº " & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value, 4 + 32, "ATENCIÓN") = vbYes Then
'                        'intIndex = right(dxDBGrid1.Columns.FocusedColumn.FieldName, 1)
'                        'csql = "Update IF4ORDEN SET F4ESTADO = 1 WHERE F4NUMORD='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "'"
'                        'intEstado = 1
'                        'dxDBGrid1.Dataset.Edit
'                        'dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = False
'                        'dxDBGrid1.Columns.ColumnByFieldName("F4colocada").Value = False
'                        'dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Value = False
'                        'dxDBGrid1.Columns.ColumnByFieldName("f4vbuser" & intIndex).Value = wusuario
'                        'dxDBGrid1.Columns.ColumnByFieldName("f4estado").Value = intEstado
'                        'dxDBGrid1.Columns.ColumnByFieldName("f4vbfecha" & intIndex).Value = Now
'                        'dxDBGrid1.Dataset.Post
'                        'est1 = "1"
'                        frmOrdPago.Show 1
'                    Else
'                        Exit Sub
'                    End If
'                End If
'            ElseIf ValidaSaldoOrden = 0 Then
'                MsgBox "No se puede autorizar más pagos porque el saldo de la Orden de Compra es Cero", vbInformation, "ATENCIÓN"
'                Exit Sub
'            Else
'                MsgBox "La Orden de Compra no registra pagos parciales", vbInformation, "ATENCIÓN"
'                Exit Sub
'            End If
'
'
''            If dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = True Then
''                Llena_Orden_de_Pago
''            Else
''                frmOrdPago.Show 1
''            End If
        Case "ID_Movimiento":
            'MsgBox "Opción no disponible.", vbInformation + vbOKOnly, App.ProductName
            
'            imprimir_movoc
            'sw_nuevo_documento = True
            'ocompra_multiples.Show 1
            
            If Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value & "") <> "OC" Then
                MsgBox "Opción disponible solo para Ordenes de Compra.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Dim rpt As New rptOrdenMovimiento
            
            With rpt
                .TipoOrden = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4LOCAL").value & "")
                .NumeroOrden = Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")
                
                .Show
            End With
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
            Me.MousePointer = vbHourglass
            GENERA_EXCEL
            Me.MousePointer = vbDefault
        Case "ID_Salir"
            Unload Me
    End Select

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

Private Sub TTabDock1_PanelResize(ByVal Panel As TabDock.TTabDockHost)
Form_Resize
End Sub

Private Sub txtbusqueda_Change()
    filtrox = 1
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2NOMPROV LIKE '*" & txtBusqueda.Text & "*' OR " & " f4numord LIKE '*" & txtBusqueda.Text & "*' OR f3abrev LIKE '*" & txtBusqueda.Text & "*'  OR F4codsolicitud LIKE '*" & txtBusqueda.Text & "*'  OR F3DESCRIP LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
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
        With cmdsave
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
                    GetGridByActive().M.SaveAllToTextFile (FileName)
                    Else
                    GetGridByActive().M.SaveSelectedToTextFile (FileName)
                    End If
                Case 245
                    .Filter = "Excel Files (*.xls)|*.xls"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().M.ExportToXLS FileName
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
                    GetGridByActive().M.ExportToHTML FileName
                Case 205
                    .Filter = "XML Files (*.xml)|*.xml"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().M.ExportToXML FileName
                Case 201
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "CONTROL Plus!") = vbYes Then _
                    GetGridByActive().M.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().M.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
'    End If
    Exit Sub
    
errhandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    
    Exit Sub
End Sub

Private Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim i As Byte

    If Ind > 199 Then   'ind=245
        SaveTo (Ind)
        Exit Sub
    End If

End Sub

Private Sub imprimir_movoc()
'Dim rsrep       As New ADODB.Recordset
Dim csql        As String
'Dim Cad         As String
Dim i           As Integer

    acr_movoc2.FldEmpresa.Text = wnomcia
    acr_movoc2.fldFecha.Text = Format(Date, "dd/mm/yyyy")
    acr_movoc2.Label43.Visible = True
    '''''acr_movoc.Field14.Visible = True
    acr_movoc2.lblnumorden.Visible = True
    acr_movoc2.fldnumorden.Visible = True
    acr_movoc2.fldnomproveedores.Visible = True
    acr_movoc2.lblproveedores.Visible = True
    acr_movoc2.Field11.Text = Format("" & traerCampo("IF4ORDEN", "F4MONTO", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value), "#,###,###0.00")
    acr_movoc2.fldnomproveedores.Text = "" & dxDBGrid1.Columns.ColumnByFieldName("F2NOMPROV").value
    acr_movoc2.Field40.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    'acr_movoc2.Field39.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value) = "S", "S/", "$")
    acr_movoc2.Field38.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    acr_movoc2.Field37.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    acr_movoc2.Field16.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    'acr_movoc2.Field28.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value) = "S", "S/", "$")
    acr_movoc2.Field30.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    acr_movoc2.Field15.Text = "" & IIf(traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value) = "S", "S/", "$")
    
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
''''''''''    csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA as FH, iif(REGISDOC.F4MONEDA='S', 'S/', '$') as F4MONEDA, (REGISDOC.F4TOTAL)* -1 as F4TOTAL, PAG_DCTO.saldo , ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, IF4ORDEN.F4MONTO - PAG_DCTO.TOTAL AS TOTAL  "          ', IF4ORDEN.F4MONTO - Sum(REGISDOC.F4TOTAL) as Saldo2 "
''''''''''    'csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA as FH, iif(REGISDOC.F4MONEDA='S', 'S/', '$') as F4MONEDA, (REGISDOC.F4TOTAL)* -1 as F4TOTAL, PAG_DCTO.saldo , ((REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, IF4ORDEN.F4MONTO - PAG_DCTO.TOTAL AS TOTAL  "          ', IF4ORDEN.F4MONTO - Sum(REGISDOC.F4TOTAL) as Saldo2 "
''''''''''    csql = csql + "FROM ((IF4ORDEN INNER JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) INNER JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC "
''''''''''    csql = csql + "WHERE ((IF4ORDEN.F4NUMORD) = '" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "')"
''''''''''''    csql = csql + " AND (([if4orden].[f4monto]-[regisdoc].[f4total])<=-0.02 Or ([if4orden].[f4monto]-[regisdoc].[f4total])>=0.02) "
''''''''''    'csql = csql + " Group By IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA, REGISDOC.F4MONEDA, REGISDOC.F4TOTAL, PAG_DCTO.saldo, (REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO * -1 "

''    csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, "
''    csql = csql + "IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA AS FH, IIf(REGISDOC.F4MONEDA='S','S/','$') AS F4MONEDA, (REGISDOC.F4TOTAL)*-1  "
''    csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, (PAG_DCTO.TOTAL - PAG_DCTO.saldo)*-1 AS TOTAL, 'Factura' as Concepto  "
''    'csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, IF4ORDEN.F4MONTO-PAG_DCTO.TOTAL AS TOTAL "
''    csql = csql + " FROM ((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) "
''    csql = csql + " LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC "
''    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "')) "
''    csql = csql + " union ALL"
''    csql = csql + " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV,  "
''    csql = csql + " IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, PAG_DCTO.FCH_COMP AS FH, IIf(PAG_DCTO.MONEDA='S','S/','$') AS F4MONEDA, (IF4ORDEN.F4MONTO-REGISDOC.F4TOTAL)*-1 AS F4TOTAL, "
''    csql = csql + " PAG_DCTO.saldo, ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO)*-1 AS PORCENTAJE, PAG_DCTO.TOTAL*-1, 'Pago a Cuenta' as Concepto "
''    csql = csql + " FROM ((IF4ORDEN LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) LEFT JOIN PAG_DCTO ON  "
''    csql = csql + " IF4ORDEN.F4NUMORD = PAG_DCTO.F4OCOMPRA) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA "
''    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "') AND ((Mid([pag_dcto].[nro_comp],1,3))='Ant')) "




    csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, "
    csql = csql + "IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA AS FH, IIf(REGISDOC.F4MONEDA='S','S/','$') AS F4MONEDA, (REGISDOC.F4TOTAL)  "
    csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO)  AS PORCENTAJE, (PAG_DCTO.TOTAL - PAG_DCTO.saldo) AS TOTAL,IF4ORDEN_PAGO.observacion AS Concepto, IF4ORDEN_PAGO.IMPORTE, ((REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO) AS PDOCU, IF4ORDEN_PAGO.FECHA, PAG_DCTO.FCH_COMP  "
    'csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, (PAG_DCTO.TOTAL - PAG_DCTO.saldo) AS TOTAL, 'Factura' as Concepto, IF4ORDEN_PAGO.IMPORTE, ((REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO) AS PDOCU, IF4ORDEN_PAGO.FECHA, PAG_DCTO.FCH_COMP  "
    csql = csql + " FROM (((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) "
    csql = csql + " LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
    csql = csql + " LEFT JOIN IF4ORDEN_PAGO ON (REGISDOC.F4CORRELA = IF4ORDEN_PAGO.correladoc) AND (REGISDOC.F4OCOMPRA = IF4ORDEN_PAGO.ORDEN) "
    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "')) And RegisDoc.F4TOTAL > 0 And IF4ORDEN_PAGO.ESTADO = '1' "
    csql = csql + " union ALL"
    csql = csql + " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV,  "
    csql = csql + " IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA AS FH, IIf(PAG_DCTO.MONEDA='S','S/','$') AS F4MONEDA, (IF4ORDEN.F4MONTO-REGISDOC.F4TOTAL) AS F4TOTAL, "
    ''csql = csql + " IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, PAG_DCTO.FCH_COMP AS FH, IIf(PAG_DCTO.MONEDA='S','S/','$') AS F4MONEDA, (IF4ORDEN.F4MONTO-REGISDOC.F4TOTAL) AS F4TOTAL, "
    csql = csql + " PAG_DCTO.saldo, ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO) AS PORCENTAJE, PAG_DCTO.TOTAL,  IF4ORDEN_PAGO.observacion AS concepto, IF4ORDEN_PAGO.IMPORTE, '0.00' AS PDOCU, IF4ORDEN_PAGO.FECHA, PAG_DCTO.FCH_COMP   "
    csql = csql + " FROM (((IF4ORDEN LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) LEFT JOIN PAG_DCTO ON  "
    csql = csql + " IF4ORDEN.F4NUMORD = PAG_DCTO.F4OCOMPRA) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA) "
    csql = csql + " LEFT JOIN IF4ORDEN_PAGO ON PAG_DCTO.correla = IF4ORDEN_PAGO.correlaanticipo "
    csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "') AND ((Mid([pag_dcto].[nro_comp],1,3))='Ant'))  And IF4ORDEN_PAGO.ESTADO = '1' "
    csql = csql + " union ALL"
    csql = csql + " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN_PAGO.ORDEN AS F4NUMORD, IF4ORDEN_PAGO.FECHA AS F4FECEMI, IF4ORDEN_PAGO.IdOp AS nro_comp, '' AS F2NOMPROV, "
    csql = csql + " IF4ORDEN_PAGO.IMPORTE AS F4MONTO, IF4ORDEN_PAGO.MONEDA AS F4TIPMON, "
    csql = csql + " '' AS FH, '' AS F4MONEDA, '' AS F4TOTAL, '' AS SALDO, '0.00' AS PORCENTAJE, '' AS TOTAL, IF4ORDEN_PAGO.observacion AS Concepto, IF4ORDEN_PAGO.IMPORTE, '0.00' AS PDOCU, IF4ORDEN_PAGO.FECHA, '' AS FCH_COMP "
    csql = csql + " FROM IF4ORDEN INNER JOIN IF4ORDEN_PAGO ON IF4ORDEN.F4NUMORD = IF4ORDEN_PAGO.ORDEN "
    csql = csql + " WHERE (((IF4ORDEN_PAGO.ORDEN)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "') AND ((IF4ORDEN_PAGO.correladoc) =0) AND ((IF4ORDEN_PAGO.correlaanticipo) =0))   And IF4ORDEN_PAGO.ESTADO = '1' "
    csql = csql + " ORDER BY FECHA"

    acr_movoc2.datconexion.ConnectionString = cnn_dbbancos
    acr_movoc2.datconexion.Source = csql
    Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    If RsS.RecordCount > 0 Then
        acr_movoc2.Show vbModal
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
Private Sub Mostrar_Reporte()
''''Dim rs1 As New ADODB.Recordset
''''Dim rs2 As New ADODB.Recordset
''''Dim rs3 As New ADODB.Recordset
''''
''''csql = ""
''''csql = "SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV, "
''''csql = csql + "IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, REGISDOC.F4FECHA AS FH, IIf(REGISDOC.F4MONEDA='S','S/','$') AS F4MONEDA, (REGISDOC.F4TOTAL)  "
''''csql = csql + "AS F4TOTAL, PAG_DCTO.saldo, '0.00' AS PORCENTAJE, (PAG_DCTO.TOTAL - PAG_DCTO.saldo)*-1 AS TOTAL, 'Factura' as Concepto, IF4ORDEN_PAGO.IMPORTE, ((REGISDOC.F4TOTAL*100)/IF4ORDEN.F4MONTO) AS PDOCU  "
''''csql = csql + " FROM (((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) "
''''csql = csql + " LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
''''csql = csql + " LEFT JOIN IF4ORDEN_PAGO ON (REGISDOC.F4CORRELA = IF4ORDEN_PAGO.correladoc) AND (REGISDOC.F4OCOMPRA = IF4ORDEN_PAGO.ORDEN) "
''''csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "')) "
''''
''''If rs1.State = adStateOpen Then rs1.Close
''''rs1.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
''''
''''csql = ""
''''csql = " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, PAG_DCTO.NRO_COMP, EF2PROVEEDORES.F2NOMPROV,  "
''''csql = csql + " IF4ORDEN.F4MONTO, IF4ORDEN.F4TIPMON, PAG_DCTO.FCH_COMP AS FH, IIf(PAG_DCTO.MONEDA='S','S/','$') AS F4MONEDA, (IF4ORDEN.F4MONTO-REGISDOC.F4TOTAL) AS F4TOTAL, "
''''csql = csql + " PAG_DCTO.saldo, ((PAG_DCTO.TOTAL*100)/IF4ORDEN.F4MONTO) AS PORCENTAJE, PAG_DCTO.TOTAL, 'Pago a Cuenta' as Concepto, IF4ORDEN_PAGO.IMPORTE, '0.00' AS PDOCU "
''''csql = csql + " FROM (((IF4ORDEN LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) LEFT JOIN PAG_DCTO ON  "
''''csql = csql + " IF4ORDEN.F4NUMORD = PAG_DCTO.F4OCOMPRA) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA) "
''''csql = csql + " LEFT JOIN IF4ORDEN_PAGO ON PAG_DCTO.correla = IF4ORDEN_PAGO.correlaanticipo "
''''csql = csql + " WHERE (((IF4ORDEN.F4NUMORD)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "') AND ((Mid([pag_dcto].[nro_comp],1,3))='Ant')) "
''''
''''If rs2.State = adStateOpen Then rs2.Close
''''rs2.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
''''
''''csql = ""
''''csql = " SELECT IF4ORDEN.F4CODSOLICITUD, IF4ORDEN_PAGO.ORDEN AS F4NUMORD, IF4ORDEN_PAGO.FECHA AS F4FECEMI, IF4ORDEN_PAGO.IdOp AS nro_comp, '' AS F2NOMPROV, "
''''csql = csql + " IF4ORDEN_PAGO.IMPORTE AS F4MONTO, IF4ORDEN_PAGO.MONEDA AS F4TIPMON, "
''''csql = csql + " IF4ORDEN_PAGO.FECHA AS FH, '' AS F4MONEDA, '' AS F4TOTAL, '' AS SALDO, '0.00' AS PORCENTAJE, '' AS TOTAL, IF4ORDEN_PAGO.observacion AS Concepto, IF4ORDEN_PAGO.IMPORTE, '0.00' AS PDOCU "
''''csql = csql + " FROM IF4ORDEN INNER JOIN IF4ORDEN_PAGO ON IF4ORDEN.F4NUMORD = IF4ORDEN_PAGO.ORDEN "
''''csql = csql + " WHERE (((IF4ORDEN_PAGO.ORDEN)='" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value & "') AND ((IF4ORDEN_PAGO.correladoc) =0) AND ((IF4ORDEN_PAGO.correlaanticipo) =0))"
''''
''''If rs3.State = adStateOpen Then rs3.Close
''''rs3.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
''''
''''
''''
End Sub

Private Sub SALDOS()
Dim rsAutPagos As New ADODB.Recordset
Dim J As Integer
Dim CountOrdenes As Integer
Dim i As Integer
Dim SumImportes As Double
Dim SaldosOCs  As Double

'If dxDBGrid1.Dataset.RecordCount > 0 Then
'    For I = 1 To dxDBGrid1.Dataset.RecordCount
'        dxDBGrid1.Dataset.RecNo = I
'        SumImportes = 0
'        CountOrdenes = traerCampo("IF4ORDEN_PAGO", "COUNT(*)", "ORDEN", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
'        For J = 1 To CountOrdenes
'            dxDBGrid1.Dataset.Edit
'            SumImportes = SumImportes + IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("IMPORTE").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("IMPORTE").Value) 'traerCampo("IF4ORDEN_PAGO", "IMPORTE", "F4NUMORD", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
'            dxDBGrid1.Columns.ColumnByFieldName("F4SALDOOC").Value = dxDBGrid1.Columns.ColumnByFieldName("F4MONTO").Value - SumSaldos
'            cnn_dbbancos.Execute ("Update IF4ORDEN")
'            dxDBGrid1.Dataset.Post
'        Next J
'    Next I
'End If
ctipo = "A"
If RsListaOC.RecordCount > 0 Then
    RsListaOC.MoveFirst
    Do While Not RsListaOC.EOF
        SaldosOCs = 0
        SumImportes = 0
        CountOrdenes = traerCampo("IF4ORDEN_PAGO", "COUNT(*)", "ORDEN", RsListaOC.Fields("F4NUMORD").value)
        
        sql = "SELECT PAGOS.IDOP, CENTROS.F3ABREV, A.F4NUMORD, A.F4CODSOLICITUD, A.F4ESTNUL, B.F2NOMPROV, A.F4FECEMI, A.F4TIPMON, A.F4COLOCADA, A.F4COLOCADAUSER, A.F4COLOCADAFECHA, A.F4ATENDIDA, A.F4ATENDIDAUSER, A.F4ATENDIDAFECHA, IIf(A.F4TIPMON='S',A.F4MONTO,'') AS F4SOLES, IIf(A.F4TIPMON='D',A.F4MONTO,'') AS F4DOLARES, A.F4MONTO, A.F4TIPCAM, A.F4FECENT, A.F4ESTADO, A.F4VB1, A.F4VBUSER1, A.F4VBFECHA1, A.F4VB2, A.F4VBUSER2, A.F4VBFECHA2, IF4ORDEN_PAGO.IMPORTE, A.F4MONTO - SUM(PAGOS.IMPORTE) AS SALDO_OC"
        sql = sql & " FROM ((IF4ORDEN AS A LEFT JOIN EF2PROVEEDORES AS B ON A.F4CODPRV = B.F2NEWRUC) LEFT JOIN CENTROS ON A.F4CENTRO = CENTROS.F3COSTO) LEFT JOIN (SELECT IF4ORDEN_PAGO.ORDEN, IF4ORDEN_PAGO.IMPORTE, IF4ORDEN_PAGO.ESTADO, IF4ORDEN_PAGO.IDOP FROM IF4ORDEN_PAGO WHERE IF4ORDEN_PAGO.ESTADO='1') AS PAGOS ON A.F4NUMORD = PAGOS.ORDEN"
        sql = sql & " WHERE (((A.F4LOCAL)='1')) And A.F4NUMORD = '" & RsListaOC.Fields("f4numord").value & "'  "
        
        If strCodCentros <> "'999'" Then
            sql = sql & "AND A.F4CENTRO IN (" & strCodCentros & ")"
        End If
            ''''
        sql = sql & " Group By PAGOS.IDOP, CENTROS.F3ABREV, A.F4NUMORD, A.F4CODSOLICITUD, A.F4ESTNUL, B.F2NOMPROV, A.F4FECEMI, A.F4TIPMON, A.F4COLOCADA, A.F4COLOCADAUSER, A.F4COLOCADAFECHA, A.F4ATENDIDA, A.F4ATENDIDAUSER, A.F4ATENDIDAFECHA, IIf(A.F4TIPMON='S',A.F4MONTO,''), IIf(A.F4TIPMON='D',A.F4MONTO,''), A.F4MONTO, A.F4TIPCAM, A.F4FECENT, A.F4ESTADO, A.F4VB1, A.F4VBUSER1, A.F4VBFECHA1, A.F4VB2, A.F4VBUSER2, A.F4VBFECHA2,  IF4ORDEN_PAGO.IMPORTE "
        sql = sql & "ORDER BY A.F4NUMORD DESC"

        If rsAutPagos.State = adStateOpen Then rsAutPagos.Close
        rsAutPagos.Open csql, cnn_bancos, adOpenKeyset, adLockOptimistic
        
        'For J = 1 To CountOrdenes
        Do While Not rsAutPagos.EOF
            SumImportes = SumImportes + IIf(IsNull(RsListaOC.Fields("IMPORTE").value), 0, RsListaOC.Fields("IMPORTE").value)
            SaldosOCs = RsListaOC.Fields("F4MONTO").value - SumImportes
            
''''            amovs_cab(0).campo = "IDOP": amovs_cab(0).valor = "" & RsListaOC(0): amovs_cab(0).TIPO = "T"
''''            amovs_cab(1).campo = "F3ABREV": amovs_cab(1).valor = "" & RsListaOC(1): amovs_cab(1).TIPO = "T"
''''            amovs_cab(2).campo = "F4NUMORD": amovs_cab(2).valor = "" & RsListaOC(2): amovs_cab(2).TIPO = "T"
''''            amovs_cab(3).campo = "F4CODSOLICITUD": amovs_cab(3).valor = "" & RsListaOC(3): amovs_cab(3).TIPO = "T"
''''            amovs_cab(4).campo = "F4ESTNUL": amovs_cab(4).valor = "" & RsListaOC(4): amovs_cab(4).TIPO = "T"
''''            amovs_cab(5).campo = "F2NOMPROV": amovs_cab(5).valor = "" & RsListaOC(5): amovs_cab(5).TIPO = "T"
''''            amovs_cab(6).campo = "F4FECEMI": amovs_cab(6).valor = "" & RsListaOC(6): amovs_cab(6).TIPO = "F"
''''            amovs_cab(7).campo = "F4TIPMON": amovs_cab(7).valor = "" & RsListaOC(7): amovs_cab(7).TIPO = "T"
''''            amovs_cab(8).campo = "F4COLOCADA": amovs_cab(8).valor = "" & RsListaOC(8): amovs_cab(8).TIPO = "N"
''''            amovs_cab(9).campo = "F4COLOCADAUSER": amovs_cab(9).valor = "" & RsListaOC(9): amovs_cab(9).TIPO = "T"
''''            amovs_cab(10).campo = "F4COLOCADAFECHA": amovs_cab(10).valor = "" & RsListaOC(10): amovs_cab(10).TIPO = "F"
''''            amovs_cab(11).campo = "F4ATENDIDA": amovs_cab(11).valor = "" & RsListaOC(11): amovs_cab(11).TIPO = "N"
''''            amovs_cab(12).campo = "F4ATENDIDAUSER": amovs_cab(12).valor = "" & RsListaOC(12): amovs_cab(12).TIPO = "T"
''''            amovs_cab(13).campo = "F4ATENDIDAFECHA": amovs_cab(13).valor = "" & RsListaOC(13): amovs_cab(13).TIPO = "F"
''''            amovs_cab(14).campo = "F4SOLES": amovs_cab(14).valor = "" & RsListaOC(14): amovs_cab(14).TIPO = "T"
''''            amovs_cab(15).campo = "F4DOLARES": amovs_cab(15).valor = "" & RsListaOC(15): amovs_cab(15).TIPO = "T"
''''            amovs_cab(16).campo = "F4MONTO": amovs_cab(16).valor = "" & RsListaOC(16): amovs_cab(16).TIPO = "N"
''''            amovs_cab(17).campo = "F4TIPCAM": amovs_cab(17).valor = "" & RsListaOC(17): amovs_cab(17).TIPO = "N"
''''            amovs_cab(18).campo = "F4FECENT": amovs_cab(18).valor = "" & RsListaOC(18): amovs_cab(18).TIPO = "F"
''''            amovs_cab(19).campo = "F4ESTADO": amovs_cab(19).valor = "" & RsListaOC(19): amovs_cab(19).TIPO = "N"
''''            amovs_cab(20).campo = "F4VB1": amovs_cab(20).valor = "" & RsListaOC(20): amovs_cab(20).TIPO = "N"
''''            amovs_cab(21).campo = "F4VBUSER1": amovs_cab(21).valor = "" & RsListaOC(21): amovs_cab(21).TIPO = "T"
''''            amovs_cab(22).campo = "F4VBFECHA1": amovs_cab(22).valor = "" & RsListaOC(22): amovs_cab(22).TIPO = "F"
''''            amovs_cab(23).campo = "F4VB2": amovs_cab(23).valor = "" & RsListaOC(23): amovs_cab(23).TIPO = "N"
''''            amovs_cab(24).campo = "F4VBUSER2": amovs_cab(24).valor = "" & RsListaOC(24): amovs_cab(24).TIPO = "T"
''''            amovs_cab(25).campo = "F4VBFECHA2": amovs_cab(25).valor = "" & RsListaOC(25): amovs_cab(25).TIPO = "F"
''''            amovs_cab(26).campo = "IMPORTE": amovs_cab(26).valor = "" & RsListaOC(26): amovs_cab(26).TIPO = "N"
''''            amovs_cab(27).campo = "SALDO_OC": amovs_cab(27).valor = "" & SaldosOCs: amovs_cab(27).TIPO = "N"
            rsAutPagos.MoveNext
        Loop
            

        'Next J
        
        amovs_cab(0).campo = "IMPORTE": amovs_cab(0).valor = "" & IIf(IsNull(RsListaOC(26)), 0#, RsListaOC(26)): amovs_cab(0).Tipo = "N"
        amovs_cab(1).campo = "SUMIMPORTES": amovs_cab(1).valor = "" & Format(SumImportes, "0.00"): amovs_cab(1).Tipo = "N"
        amovs_cab(2).campo = "MONTO": amovs_cab(2).valor = "" & IIf(IsNull(RsListaOC(16)), 0#, RsListaOC(16)): amovs_cab(2).Tipo = "N"
        amovs_cab(3).campo = "ORDEN": amovs_cab(3).valor = "" & IIf(IsNull(RsListaOC(2)), 0#, RsListaOC(2)): amovs_cab(3).Tipo = "T"
        amovs_cab(4).campo = "COD": amovs_cab(4).valor = "" & IIf(IsNull(RsListaOC(0)), "", RsListaOC(0)): amovs_cab(4).Tipo = "T"
        If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
                 GRABA_REGISTRO_logistica amovs_cab(), "TMP_SALDOSOCS", ctipo, 4, cnn_form, ""
        Else    '--- Modificación
        '------- GRABA CABECERA
                 GRABA_REGISTRO_logistica amovs_cab(), "TMP_SALDOSOCS", ctipo, 4, cnn_form, ""
        '        'AlmacenaQuery_sql csql, cnn_dbbancos
        End If

        
        RsListaOC.MoveNext
    Loop
End If
End Sub

Private Function ValidaSaldoOrden() As Double
    Dim MontoOCs As Double
    Dim MontoAutPag As Double
    
    MontoAutPag = "" & traerCampo("IF4ORDEN_PAGO", "IIF(Sum(Importe) is Null, 0, Sum(Importe))", "ORDEN", dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value, " And Estado = '1'")
    MontoOCs = Val("" & traerCampo("IF4ORDEN", "F4MONTO", "F4NUMORD", Trim(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "")))
    
    If MontoAutPag > 0 Then
        ValidaSaldoOrden = Format(MontoOCs - MontoAutPag, "#,###,###0.00")
    Else
        ValidaSaldoOrden = -1
    End If
End Function
