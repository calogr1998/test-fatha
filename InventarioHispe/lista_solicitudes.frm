VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Begin VB.Form lista_solicitudes 
   Caption         =   "Listado de Requerimientos"
   ClientHeight    =   7590
   ClientLeft      =   960
   ClientTop       =   1800
   ClientWidth     =   16125
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   16125
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlgSolicitud 
      Left            =   0
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   735
      Left            =   10200
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   3480
      Top             =   5880
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   10005
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   8040
      End
      Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
         Height          =   270
         Left            =   8520
         TabIndex        =   4
         Top             =   360
         Width           =   870
         _Version        =   65536
         _cx             =   1535
         _cy             =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Fechas"
         Enabled         =   -1  'True
         AutoSize        =   -1  'True
         BackStyle       =   1
         BackColor       =   15790320
         ForeColor       =   0
         ViewStyle       =   1
         Checked         =   0   'False
         GroupIndex      =   -1
         TextLayout      =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4035
      Left            =   120
      OleObjectBlob   =   "lista_solicitudes.frx":0000
      TabIndex        =   0
      Top             =   1080
      Width           =   10215
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "lista_solicitudes.frx":5FEE
      ToolBars        =   "lista_solicitudes.frx":EB24
   End
   Begin TabDock.TTabDock TTabDock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar 
      Height          =   675
      Left            =   600
      OleObjectBlob   =   "lista_solicitudes.frx":ED4C
      TabIndex        =   1
      Top             =   5460
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "lista_solicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTimer As Integer
Dim rslista     As ADODB.Recordset
Dim I           As Byte
Dim DBName      As String
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
Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal HBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wformat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal Rgn As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const DT_CALCRECT = &H400


Private Sub dxCheckBox1_Click()
    If dxCheckBox1.Checked = 1 Then
        dxDBGrid1.Columns.ColumnByFieldName("Fecha_graba").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("Fecha_modifica").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("cs_motivos").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("VBFECHA").Visible = True
    Else
        dxDBGrid1.Columns.ColumnByFieldName("Fecha_graba").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("Fecha_modifica").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("cs_motivos").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("VBFECHA").Visible = False
    End If
End Sub

Private Sub dxDBGrid1_OnChangeNodeEx()
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName
        Case "cs_documento", "cod_solicitud"
            Dim valor As String
            Dim Tipo As String
            
            valor = dxDBGrid1.Columns.ColumnByFieldName("cod_solicitud").Value
            Tipo = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                Call proceso2Sql(valor, Tipo)
            Else
                Call proceso2(valor, Tipo)
            End If
    End Select
End Sub

Public Sub proceso2(ByVal Codigo As String, Tipo As String)
    'Dim SqlCad As String
    
    With lista_solicitudes_detalle.dxDBGrid2
        If .Dataset.State = 1 Then .Dataset.Close
        
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DS.CS_DOCUMENTO, "
        SqlCad = SqlCad & "DS.COD_SOLICITUD, "
        SqlCad = SqlCad & "DS.ITEM, "
        SqlCad = SqlCad & "DS.COD_PRODUCTO, "
        SqlCad = SqlCad & "DS.F5CODFAB, "
        SqlCad = SqlCad & "TRIM(DS.DS_DESCRIPCION & '') AS DS_DESCRIPCION, "
        'SqlCad = SqlCad & "ds.ds_unidmed, ds.ds_cantidad, ds.proveedor, ds.centro,ds.observa "
        SqlCad = SqlCad & "MED.F7SIGMED AS DS_UNIDMED, "
        SqlCad = SqlCad & "VAL(DS.DS_CANTIDAD & '') AS DS_CANTIDAD, "
        'SqlCad = SqlCad & "VAL(DS.DS_CANTIDAD & '') - VAL(RQENORDEN.CANTIDAD & '') AS SALDO, "
        SqlCad = SqlCad & "VAL(COMPROMISO.CANTIDAD & '') AS CANTCOMPROMETIDA, "
        SqlCad = SqlCad & "VAL(PORLLEGAR.CANTIDAD & '') AS CANTPORLLEGAR, "
        SqlCad = SqlCad & "VAL(DS.DS_CANTIDAD & '') - (VAL(COMPROMISO.CANTIDAD & '') + VAL(PORLLEGAR.CANTIDAD & '')) AS SALDO, "
        SqlCad = SqlCad & "DS.PROVEEDOR, "
        SqlCad = SqlCad & "DS.CENTRO, "
        'SqlCad = SqlCad & "DS.OBSERVA "
        SqlCad = SqlCad & "IIF(VAL(SALDO & '') = 0, 'ATENDIDO TOTALMENTE', IIF(VAL(SALDO & '') = VAL(DS.DS_CANTIDAD & ''), 'POR ATENDER', IIF(VAL(SALDO & '') > 0 ,'ATENDIDO PARCIALMENTE', 'ATENCIÓN EXCEDE A LO REQUERIDO'))) AS OBSERVA "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "((TB_DETSOLICITUD AS DS "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DS.DS_UNIDMED) "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "("
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.COD_SOLICITUD AS NROPEDIDO, "
        SqlCad = SqlCad & "DET.F5CODPRO, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F4TIPOVALE = 'I' AND "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XCS') AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD = '" & Codigo & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPRO"
        SqlCad = SqlCad & ") AS COMPROMISO "
        SqlCad = SqlCad & "ON "
        SqlCad = SqlCad & "TRIM(COMPROMISO.NROPEDIDO & '') = TRIM(DS.COD_SOLICITUD & '') AND "
        SqlCad = SqlCad & "TRIM(COMPROMISO.F5CODPRO & '') = TRIM(DS.COD_PRODUCTO & '')) "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "("
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        'SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
        SqlCad = SqlCad & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) - VAL(INGRESOS.CANTIDAD & '')) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') AS F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD AS COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
        SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') <> '' AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' " 'AND "
        'SqlCad = SqlCad & "TRIM(DET.F5CODPROORIGINAL & '') = '" & strCodProducto & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "TRIM(DET.F4NUMORD & ''), "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        SqlCad = SqlCad & "ON TRIM(INGRESOS.F4NUMORD & '') = TRIM(DET.F4NUMORD & '') AND TRIM(INGRESOS.COD_SOLICITUD & '') = DET.COD_SOLICITUD AND TRIM(INGRESOS.F5CODPROORIGINAL & '') = TRIM(DET.F3CODPRO & '') "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.COD_SOLICITUD = '" & Codigo & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO"
        SqlCad = SqlCad & ") AS PORLLEGAR "
        SqlCad = SqlCad & "ON "
        SqlCad = SqlCad & "TRIM(PORLLEGAR.COD_SOLICITUD & '') = TRIM(DS.COD_SOLICITUD & '') AND "
        SqlCad = SqlCad & "TRIM(PORLLEGAR.F3CODPRO & '') = TRIM(DS.COD_PRODUCTO & '') "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DS.COD_SOLICITUD = '" & Codigo & "' AND "
        SqlCad = SqlCad & "DS.CS_DOCUMENTO = '" & Tipo & "' "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "DS.ITEM, DS.COD_SOLICITUD"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        
        .KeyField = "ITEM"
    End With
    
    SqlCad = vbNullString
End Sub

Public Sub proceso2Sql(ByVal Codigo As String, Tipo As String)
    'Dim SqlCad As String
    
    With lista_solicitudes_detalle.dxDBGrid2
        If .Dataset.State = 1 Then .Dataset.Close
'
'        .Dataset.ADODataset.ConnectionString = cnBdCPlus
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DS.CS_DOCUMENTO, "
        SqlCad = SqlCad & "DS.COD_SOLICITUD, "
        SqlCad = SqlCad & "DS.ITEM, "
        SqlCad = SqlCad & "DS.COD_PRODUCTO, "
        SqlCad = SqlCad & "DS.F5CODFAB, "
        SqlCad = SqlCad & "RTRIM(LTRIM(DS.DS_DESCRIPCION)) AS DS_DESCRIPCION, "
        'SqlCad = SqlCad & "ds.ds_unidmed, ds.ds_cantidad, ds.proveedor, ds.centro,ds.observa "
        SqlCad = SqlCad & "MED.F7SIGMED AS DS_UNIDMED, "
        SqlCad = SqlCad & "CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD) AS DS_CANTIDAD, "
        'SqlCad = SqlCad & "VAL(DS.DS_CANTIDAD & '') - VAL(RQENORDEN.CANTIDAD & '') AS SALDO, "
        SqlCad = SqlCad & "ISNULL(COMPROMISO.CANTIDAD, 0) AS CANTCOMPROMETIDA, "
        SqlCad = SqlCad & "ISNULL(PORLLEGAR.CANTIDAD, 0) AS CANTPORLLEGAR, "
        SqlCad = SqlCad & "CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD) - (ISNULL(COMPROMISO.CANTIDAD, 0) + ISNULL(PORLLEGAR.CANTIDAD, 0)) AS SALDO, "
        SqlCad = SqlCad & "DS.PROVEEDOR, "
        SqlCad = SqlCad & "DS.CENTRO, "
        'SqlCad = SqlCad & "DS.OBSERVA "
        SqlCad = SqlCad & "IIF(CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD) - (ISNULL(COMPROMISO.CANTIDAD, 0) + ISNULL(PORLLEGAR.CANTIDAD, 0)) = 0, 'ATENDIDO TOTALMENTE', IIF(CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD) - (ISNULL(COMPROMISO.CANTIDAD, 0) + ISNULL(PORLLEGAR.CANTIDAD, 0)) = CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD), 'POR ATENDER', IIF(CONVERT(DECIMAL(10, 2), DS.DS_CANTIDAD) - (ISNULL(COMPROMISO.CANTIDAD, 0) + ISNULL(PORLLEGAR.CANTIDAD, 0)) > 0 ,'ATENDIDO PARCIALMENTE', 'ATENCIÓN EXCEDE A LO REQUERIDO'))) AS OBSERVA "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.TB_DETSOLICITUD AS DS "
        SqlCad = SqlCad & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DS.DS_UNIDMED "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "("
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.COD_SOLICITUD AS NROPEDIDO, "
        SqlCad = SqlCad & "DET.F5CODPRO, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F4TIPOVALE = 'I' AND "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XCS') AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD = '" & Codigo & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPRO"
        SqlCad = SqlCad & ") AS COMPROMISO "
        SqlCad = SqlCad & "ON "
        SqlCad = SqlCad & "COMPROMISO.NROPEDIDO = DS.COD_SOLICITUD AND "
        SqlCad = SqlCad & "COMPROMISO.F5CODPRO = DS.COD_PRODUCTO "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "("
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        'SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
        SqlCad = SqlCad & "SUM(((DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) * ISNULL(MEDALTER.F5FACTOR, 1)) - ISNULL(INGRESOS.CANTIDAD, 0)) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "RTRIM(LTRIM(DET.F4NUMORD)) AS F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD AS COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(DET.TIPO = 'S', -1, 1)) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0', 'XNC') AND "
        SqlCad = SqlCad & "RTRIM(LTRIM(DET.F4NUMORD)) <> '' AND "
        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' " 'AND "
        'SqlCad = SqlCad & "TRIM(DET.F5CODPROORIGINAL & '') = '" & strCodProducto & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "RTRIM(LTRIM(DET.F4NUMORD)), "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.COD_SOLICITUD = '" & Codigo & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO"
        SqlCad = SqlCad & ") AS PORLLEGAR "
        SqlCad = SqlCad & "ON "
        SqlCad = SqlCad & "PORLLEGAR.COD_SOLICITUD = DS.COD_SOLICITUD AND "
        SqlCad = SqlCad & "PORLLEGAR.F3CODPRO = DS.COD_PRODUCTO "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DS.COD_SOLICITUD = '" & Codigo & "' AND "
        SqlCad = SqlCad & "DS.CS_DOCUMENTO = '" & Tipo & "' "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "DS.ITEM, DS.COD_SOLICITUD"
        
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = SqlCad
'        .Dataset.Active = True
'
'        .KeyField = "ITEM"
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "ITEM"
    End With
    
    SqlCad = vbNullString
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Dim StrPregunta As String
    Dim SwAprobacion As Boolean
    Dim StrEstado As String
    
    Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
        Case "VBJEFECC"
            If dxDBGrid1.Columns.ColumnByFieldName("VBJEFECC").Value = False Then
                StrPregunta = "¿ Desea aprobar el Requerimiento ?"
                
                SwAprobacion = False
            Else
                StrPregunta = "¿Desea quitar la aprobación del Requerimiento?"
                
                SwAprobacion = True
            End If
            
            If dxDBGrid1.Columns.ColumnByFieldName("CS_ESTADO").Value <= "2" Then
                dxDBGrid1.Dataset.Edit
                
                If MsgBox(StrPregunta, vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
                    dxDBGrid1.Columns.FocusedColumn.Value = Not SwAprobacion
                Else
                    dxDBGrid1.Columns.FocusedColumn.Value = SwAprobacion
                End If
                
                If dxDBGrid1.Columns.ColumnByFieldName("VBJEFECC").Value = True Then
                    StrEstado = "2"
                Else
                    If dxDBGrid1.Columns.ColumnByFieldName("CS_ESTADO").Value = "0" Then
                        StrEstado = "0"
                    Else
                        StrEstado = "1"
                    End If
                End If
                
                dxDBGrid1.Columns.ColumnByFieldName("CS_ESTADO").Value = StrEstado
                
                If StrEstado = "2" Then
                    dxDBGrid1.Columns.ColumnByFieldName("CS_APROBADOX").Value = wusuario
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("CS_APROBADOX").Value = ""
                End If
                
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Dataset.ADODataset.Requery
            Else
                Select Case dxDBGrid1.Columns.ColumnByFieldName("CS_ESTADO").Value & ""
                    Case "3"
                        MsgBox "El requerimiento ya fue atendido.", vbInformation, App.Title
                    Case "4"
                        MsgBox "El requerimiento ya fue cerrado.", vbInformation, App.Title
                    Case "5"
                        MsgBox "El requerimiento ha sido anulado.", vbInformation, App.Title
                End Select
                
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.FocusedColumn.Value = SwAprobacion
                
                dxDBGrid1.Dataset.Post
                
                dxDBGrid1.Dataset.ADODataset.Requery
            End If
    End Select
End Sub

Private Sub dxDBGrid1_OnClick()
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName
        Case "cs_documento", "cod_solicitud"
            Dim valor As String
            Dim Tipo As String
            
            valor = dxDBGrid1.Columns.ColumnByFieldName("cod_solicitud").Value
            Tipo = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                Call proceso2Sql(valor, Tipo)
            Else
                Call proceso2(valor, Tipo)
            End If
    End Select
End Sub

Private Sub dxDBGrid1_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, R As Single, b As Single, NeedShowTip As Boolean)
    Dim opValue As Byte
    Dim hDC0 As Long, Old_hFont As Long
    Dim nc As NONCLIENTMETRICS
    Dim rgnR As Rect
    Dim StrEstado
    Select Case UCase(Column.FieldName)
        Case "CS_ESTADO", "CS_PRIORIDAD"
            NeedShowTip = True
            
            rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
            hDC0 = GetDC(0)
            nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
            SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
            
            Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
            
            Select Case UCase(Column.FieldName)
                Case "CS_ESTADO"
                    Select Case Val(TipText & "")
                        Case 1: TipText = "Registrando"
                        Case 2: TipText = "Aprobado"
                        Case 3: TipText = "Atendido"
                        Case 4: TipText = "Cerrado"
                        Case 5: TipText = "Anulado"
                        Case Else: TipText = "No definido"
                    End Select
                Case "CS_PRIORIDAD"
                    Select Case Val(TipText & "")
                        Case 1: TipText = "Normal"
                        Case 2: TipText = "Alta"
                        Case 0: TipText = "Baja"
                    End Select
            End Select
            DrawText hDC0, TipText, Len(Trim(TipText)), rgnR, DT_CALCRECT + DT_WORDBREAK
        
            SelectObject hDC0, Old_hFont
            DeleteObject Old_hFont
            ReleaseDC HWnd, hDC0
            b = t + rgnR.bottom + 6
            R = l + rgnR.right + PicW * 2 + 4
    End Select
End Sub

Private Sub Form_Load()
    
'   Crea_Campo cconex_dbbancos, "TB_CABSOLICITUD", "VBJEFECC", "Boolean"
    '****NUEVO CAMBIO ****************
    'Crea_Campo cconex_dbbancos, "TB_CABSOLICITUD", "VBJEFECC", "YESNO", False, "False"
    'Crea_Campo cconex_dbbancos, "TB_CABSOLICITUD", "VBFECHA", "DATE", True, ""
    'Crea_Campo cconex_dbbancos, "TB_CABSOLICITUD", "VBUSER", "STRING", True, ""
    
    dxDBGrid1.Columns.ColumnByFieldName("VBJEFECC").Visible = (VerificaPermiso("0005", wusuario))
    Me.MousePointer = vbHourglass
    'Me.AutoRedraw = False
    Me.Height = 7500
    Me.Width = 10400
    Me.left = 0
    Me.top = 50
    sw_nuevo_documento = True
    'Me.AutoRedraw = True
    'verifica_mysql
    proceso
    
    TTabDock1.AddForm lista_solicitudes_detalle, tdDocked, tdAlignBottom, "lista_solicitudes_detalle"
    TTabDock1.DockedForms.ITEM("lista_solicitudes_detalle").Panel.Height = 2500
    TTabDock1.FormShow "lista_solicitudes_detalle"
    oTipoRequerimiento = ""
    Me.MousePointer = vbDefault
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.ID
        Case "ID_NuevaOC"
            oTipoRequerimiento = ""
            sw_nuevo_documento = True
            wTipoReq = 1
            Me.MousePointer = vbHourglass
            solicitud.Show 1
            Me.MousePointer = vbDefault
        Case "ID_Nueva/O.S."
            oTipoRequerimiento = Tool.ID
            sw_nuevo_documento = True
            wTipoReq = 2
            Me.MousePointer = vbHourglass
            solicitud.Show 1
            '2102
            Me.MousePointer = vbDefault
        Case "ID_Actualizar"
            Me.MousePointer = vbHourglass
            
            'verifica_mysql
            'actualizarRequerimientos
            dxDBGrid1.Dataset.Close
            
            lista_solicitudes_detalle.dxDBGrid2.Dataset.Close
            
            ModMilano.importarRequerimientosServidorExterno fraProceso, pgbProceso, Trim(SSActiveToolBars1.Tools("ID_NroPedidoT").Edit.Text)
            
            'Me.dxDBGrid1.Dataset.ADODataset.Requery
            'Me.dxDBGrid1.Dataset.Refresh
            
            SSActiveToolBars1.Tools("ID_NroPedidoT").Edit.Text = vbNullString
            
            proceso
            
            dxDBGrid1_OnChangeNodeEx
            
            Me.MousePointer = vbDefault
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
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgSolicitud
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dxDBGrid1.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "ID_Salir"
            Unload Me
        Case Else
            MsgBox "Opciones Inhabilitadas Temporalmente.", vbInformation + vbOKOnly, App.ProductName
    End Select
    
End Sub

Private Sub dxDBGrid1_OnDblClick()
    If dxDBGrid1.Dataset.RecordCount > 0 Then
        sw_nuevo_documento = False
        Me.MousePointer = vbHourglass
        solicitud.Show 1
        Unload solicitud
        Set solicitud = Nothing
        Me.MousePointer = vbDefault
    End If
End Sub

Public Sub proceso()
Dim csql        As String
Dim strCodCentros As String
Dim numobra As String
    strCodCentros = VerificaAutorizaciones("LRI", wusuario)
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        'If ctipoadm_bd = "M" Then
        '    csql = "select cod_solicitud,cs_fecha,cs_codsolicitante,if(cs_moneda = 'S', 'SOLES','DOLARES') AS CS_MONEDA,cs_codcosto,cs_total,anulado " & _
        '          " from tb_cabsolicitud WHERE anulado ='N' order by cod_solicitud DESC"
        'Else
        'csql = "select cod_solicitud,cs_prioridad,cs_fecha,cs_codsolicitante,iif(cs_moneda = 'S', 'SOLES','DOLARES') AS CS_MONEDA,"
        'csql = csql & "cs_codcosto,cs_total,anulado,NumOrden "
        'csql = csql & " from tb_cabsolicitud WHERE anulado ='N' order by cod_solicitud DESC"
        'End If
        
        csql = "SELECT csol.cs_documento,csol.cod_solicitud, UCase(csol.cs_observaciones) AS glosa, csol.cs_fecha, csol.cs_fentrega, csol.cs_prioridad, csol.cs_estado, "
        csql = csql & "FORMAT(csol.VBFECHA, 'dd/mm/yyyy') AS VBFECHA, FORMAT(csol.Fecha_graba, 'dd/mm/yyyy') AS FECHA_GRABA, FORMAT(csol.Fecha_modifica, 'dd/mm/yyyy') AS FECHA_MODIFICA, csol.cs_motivos,EF2USERS.F2NOMUSER, CENTROS.F3ABREV, csol.numorden,csol.VBJEFECC,csol.cs_aprobadox "
        csql = csql & "FROM (tb_cabsolicitud AS csol LEFT JOIN EF2USERS ON csol.cs_codsolicitante = EF2USERS.F2CODUSER) "
        csql = csql & "LEFT JOIN CENTROS ON csol.cs_codcosto = CENTROS.F3COSTO "
        
        If strCodCentros <> "'999'" Then
            csql = csql & "WHERE CS_CODCOSTO IN (" & strCodCentros & ")"
            If Len(Trim(wObra)) > 0 Then
                csql = csql & "AND CS_CODCOSTO IN ('" & wObra & "')"
            End If
        Else
            If Len(Trim(wObra)) > 0 Then
                'numobra = ObtenerCampo("CENTROS", "F3COSTO", "CODEXT2", wObra, "T", cnn_dbbancos)
                csql = csql & "WHERE CS_CODCOSTO ='" & wObra & "'"
            End If
        End If
        
        csql = csql & "ORDER BY csol.cs_fentrega DESC, csol.cs_fecha desc,csol.cod_solicitud DESC"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "cod_solicitud"
        
        .Columns.ColumnByFieldName("cs_prioridad").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageList")
        .Columns.ColumnByFieldName("cs_estado").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
    End With

End Sub



'''Private Sub cmdColor_Click(Index As Integer)
'''
'''    With cdColor
'''        .CancelError = True
'''        On Error GoTo ErrHandler
'''        .Flags = cdlCCRGBInit
'''        Select Case Index
'''            Case 0: .Color = dxDBGrid1.AutoSearchColor
'''            Case 1: .Color = dxDBGrid1.AutoSearchTextColor
'''        End Select
'''        .ShowColor
'''        Select Case Index
'''            Case 0:
'''                dxDBGrid1.AutoSearchColor = .Color
'''                lblresult.BackColor = .Color
'''            Case 1:
'''                dxDBGrid1.AutoSearchTextColor = .Color
'''                lblresult.ForeColor = .Color
'''        End Select
'''    End With
'''
'''ErrHandler:
'''    Exit Sub
'''
'''End Sub

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

Private Sub Form_Resize()
    On Error Resume Next
    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + TTabDock1.DockedForms.ITEM("lista_solicitudes_detalle").Panel.Height)
    FraBusqueda.left = 0
    FraBusqueda.top = 0
'    FraBusqueda.Width = dxDBGrid1.Width - 2000
'    txtbusqueda.Width = dxDBGrid1.Width - 350
    fraProceso.left = FraBusqueda.Width + 100
    fraProceso.top = FraBusqueda.top
    fraProceso.Width = (dxDBGrid1.Width - FraBusqueda.Width) - 100
    fraProceso.Height = FraBusqueda.Height
    pgbProceso.Width = fraProceso.Width - 1000
    pgbProceso.left = (fraProceso.Width - pgbProceso.Width) / 2
End Sub

Private Sub Timer1_Timer()

'If nTimer = 1000 Then
'    'If Trim(leeTxt) = "1" Then
'        Me.dxDBGrid1.Dataset.ADODataset.Requery
'        Me.dxDBGrid1.Dataset.Refresh
'        Open wrutabancos & "\llego correo.txt" For Output As #1
'        Write #1, 0
'        Close #1
'    'End If
'        nTimer = 0
'Else
'        nTimer = nTimer + 1
'End If

End Sub
Function leeTxt() As String
    Dim linea, Texto As String
    Open wrutabancos & "\llego correo.txt" For Input As #1
    While Not EOF(1)
       Line Input #1, linea
       If linea = vbNullString Then linea = " "
       Texto = linea
    Wend
    Close
    leeTxt = Texto
End Function

Private Sub TTabDock1_PanelResize(ByVal Panel As TabDock.TTabDockHost)
    Form_Resize
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "COD_SOLICITUD LIKE '*" & txtbusqueda.Text & "*' OR " & _
                                        "GLOSA LIKE '*" & txtbusqueda.Text & "*' OR " & _
                                        "F2NOMUSER LIKE '*" & txtbusqueda.Text & "*' OR " & _
                                        "F3ABREV LIKE '*" & txtbusqueda.Text & "*' OR " & _
                                        "NUMORDEN LIKE '*" & txtbusqueda.Text & "*'"
            
            If Len(Trim(txtbusqueda.Text)) = 0 Then
                    dxDBGrid1.Dataset.Filtered = False
            End If
        Case vbKeyDown
            dxDBGrid1.Columns.FocusedIndex = 1
            
            dxDBGrid1.SetFocus
    End Select
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        dxDBGrid1.SetFocus
'    End If
End Sub

Public Sub verifica_mysql()

Dim jj As Integer
Dim JJE As String
Dim EstRecibir As Boolean
Dim cnn_Recibe_Mysql As New ADODB.Connection
Dim cnn_Recibe As New ADODB.Connection
Dim RsM As New ADODB.Recordset
Dim xcad As String
Dim exito As Boolean
Dim leyo As Integer

    On Error GoTo Errores
    leyo = 0
    EstRecibir = False
    If cnn_Recibe_Mysql.State = 1 Then cnn_Recibe_Mysql.Close
    
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cnn_dbbancos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB;Persist Security Info=False"
    If Len(Trim(strODBC)) > 0 Then
    cnn_Recibe_Mysql.Open strODBC
    Else
        Exit Sub
    End If
    
    If cnn_Recibe.State = 1 Then cnn_Recibe.Close
    cnn_Recibe.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_Recibe.MDB;Persist Security Info=False"
    
    sql = "SELECT * FROM " & strTabla & " WHERE ISNULL(SUC_1) ORDER BY ITEM"
    If RsM.State = 1 Then RsM.Close
    RsM.Open sql, cnn_Recibe_Mysql, 3, 1
    Do While Not RsM.EOF
        leyo = 1
        exito = True
        cnn_Recibe.Execute "insert into querys (item,wquery) values ('" & RsM.Fields("item") & "','" & RsM.Fields("wquery") & "')"
        If exito = True Then
            sql = "SELECT * FROM QUERYS WHERE  item= " & Val(RsM.Fields("item"))
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, cnn_Recibe, 3, 1
            If Rs.RecordCount = 1 Then
                cnn_Recibe_Mysql.Execute "update " & strTabla & " SET SUC_1=-1 where item = " & RsM.Fields("item")
            Else
                Exit Sub
            End If
        End If
        EstRecibir = True
    RsM.MoveNext
    Loop
    
'    If EstRecibir = True Then
        sql = "SELECT * FROM QUERYS WHERE  APLICADO= 0 ORDER BY ITEM"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_Recibe, 3, 1
        Do While Not Rs.EOF
            jj = 0
            xcad = Replace(Rs!WQUERY, "|", "'")
            cnn_dbbancos.Execute xcad
            If jj = 0 Then
                cnn_Recibe.Execute "UPDATE QUERYS SET APLICADO=-1 WHERE ITEM =" & Rs!ITEM
            Else
                cnn_Recibe.Execute "UPDATE QUERYS SET ErrDiscrip = '" & JJE & "' WHERE ITEM =" & Rs!ITEM
            End If
            Rs.MoveNext
        Loop
    'End If
    EstRecibir = False
    Exit Sub
Errores:
    If leyo = 0 Then
        Exit Sub
    Else
        jj = 1
        JJE = Err.Description
        exito = False
        Resume Next
    End If
End Sub
