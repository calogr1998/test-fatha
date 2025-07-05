VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReportePedidos 
   ClientHeight    =   9840
   ClientLeft      =   750
   ClientTop       =   1020
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   14925
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   14520
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "FrmReportePedidos.frx":0000
      TabIndex        =   0
      Top             =   1440
      Width           =   17775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      ButtonWidth     =   2275
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Object.ToolTipText     =   "Exportar a un Archivo Excel"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro"
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "R. Parciales"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8460
         Top             =   0
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
               Picture         =   "FrmReportePedidos.frx":8335
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":88CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":8E69
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":9403
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":999D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":9F37
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":A2D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmReportePedidos.frx":A86B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar 
      Height          =   195
      Left            =   12360
      OleObjectBlob   =   "FrmReportePedidos.frx":AC05
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   15000
      Top             =   240
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2595
      Left            =   135
      OleObjectBlob   =   "FrmReportePedidos.frx":103BF
      TabIndex        =   5
      Top             =   7965
      Width           =   14895
   End
End
Attribute VB_Name = "FrmReportePedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rsj                 As New ADODB.Recordset
Dim RsDetalles          As New ADODB.Recordset
Dim Numeord             As String
Dim GridNum             As Byte
Dim OldValue            As Byte
Dim Af As New ADOFunctions


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

Private Sub dxDBGrid1_OnDblClick()
    X = 1
End Sub

Private Sub dxDBGrid2_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Numeord = dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value
End Sub

Private Sub dxDBGrid2_OnChangeNodeEx()
Dim valor As String
valor = (dxDBGrid2.Columns.ColumnByFieldName("Codigo_Solicitud").value)
proceso2 (valor)

End Sub

Private Sub dxDBGrid2_OnClick()
Call dxDBGrid2_OnChangeNodeEx
End Sub

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Monto OC" Or Column.Caption = "Monto" Or Column.Caption = "Total" Or Column.Caption = "Saldo" Then
    Text = Format(Text, "#,##0.00;(#,##0.00)")
    If Val(Format(Text, "#0.00")) < 0 Then
        FontColor = vbRed
    Else
        FontColor = RGB(0, 10, 150) 'RGB(10, 10, 100)
    End If
End If
If Column.FieldName = "Moneda_OC" Or Column.FieldName = "Moneda" Or Column.FieldName = "Moneda_Compr" Then
        FontColor = RGB(10, 10, 100)
End If
If Column.Caption = "Estado" Then
    If Node.Values(dxDBGrid2.Columns.ColumnByFieldName("Anticipos").Index) = "A" Then
        Color = &HFFFFC0
    End If
End If
If Column.FieldName = "Moneda_OC" Or Column.FieldName = "Moneda" Or Column.FieldName = "Moneda_Compr" Then
    If Text = "S/." Then
        Color = &HC0FFFF
    ElseIf Text = "$" Then
        Color = &HC0FFC0
    End If
End If

If Column.FieldName = "No_Orden" Then
    If Node.Values(dxDBGrid2.Columns.ColumnByFieldName("SALDO_OC").Index) <> "" And Node.Values(dxDBGrid2.Columns.ColumnByFieldName("SALDO_OC").Index) > 0 Then
        Color = RGB(255, 90, 90)
        End If
End If
    


If Column.Caption = "Saldo" Then
    If Val(Text) = 0 Then
        FontColor = vbRed
    End If
End If

If Column.Caption = "Est Solicitud" Then
    If Text = "Anulada" Then
        FontColor = vbRed
    End If
End If
If Column.Caption = "Estado OC" Then
    If Text = "Anulada" Then
        FontColor = vbRed
    End If
End If

End Sub

Private Sub dxDBGrid2_OnCustomDrawColumnHeader(ByVal hDC As Long, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Sorted As DXDBGRIDLibCtl.ExTreeColumnSort, Done As Boolean)
    Font.Bold = True
End Sub

Private Sub dxDBGrid2_OnDblClick()
Dim ncorrelativo As Double
    
    ncorrelativo = dxDBGrid2.Columns.ColumnByFieldName("correla").value
    csql = "SELECT PAG_DCTO.correla, PAG_DCTO_1.nro_comp, PAG_DCTO_1.fch_comp, PAG_DCTO_1.moneda, PAG_MVTO.imputado, PAG_DCTO_1.CTABANC, PAG_DCTO_1.MOVBANC "
    csql = csql & "FROM (PAG_DCTO INNER JOIN PAG_MVTO ON PAG_DCTO.correla = PAG_MVTO.corr_dcto) INNER JOIN PAG_DCTO AS PAG_DCTO_1 ON PAG_MVTO.corr_comp = PAG_DCTO_1.correla"

End Sub

Private Sub dxDBGrid2_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, r As Single, b As Single, NeedShowTip As Boolean)
Dim opValue As Byte
Dim hDC0 As Long, Old_hFont As Long
Dim nc As NONCLIENTMETRICS
Dim rgnR As Rect
Dim StrEstado
Select Case Column.FieldName
Case "cs_estado", "F4ESTADO", "Estados"
    NeedShowTip = True
    
    rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
    hDC0 = GetDC(0)
    nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
    Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
    Select Case Column.FieldName
    Case "cs_estado"
        Select Case Val(TipText & "")
        Case 1: TipText = "G.Req. Emitida"
        Case 2: TipText = "G.Req. Aprobada"
        Case 3: TipText = "G.Req. Atendida"
        Case 4: TipText = "G.Req. Cerrada"
        Case 5: TipText = "G.Req. Anulada"
        Case Else: TipText = "No definido"
        End Select
    'Case "Estado_OC"
    Case "F4ESTADO"
        Select Case Val(TipText & "")
'        Case 1: TipText = "Normal"
'        Case 2: TipText = "Alta"
'        Case 0: TipText = "Baja"
        Case 1: TipText = "O.C. Emitida"
        Case 2: TipText = "O.C. Aprobada"
        Case 4: TipText = "O.C. Colocada"
        Case 3: TipText = "O.C. Atendida"
        Case 5: TipText = "O.C. Anulada"
        Case 6: TipText = "O.C. Pagada"
        Case 7: TipText = "O.C. Anticipada"
        End Select
    Case "Estados"
        Select Case Val(TipText & "")
            Case 0: TipText = "O.C. Anulada"
            Case 1: TipText = "Req. Emitido"
            Case 2: TipText = "Req. Aprobado"
            Case 4: TipText = "O.C. Aprobada"
            Case 3: TipText = "O.C. Emitida"
            Case 5: TipText = "O.C. Colocada"
            Case 6: TipText = "O.C. Atendida"
            Case 7: TipText = "O.C. Pagada"
        End Select

'    Case "F4COLOCADA"
'        Select Case Val(TipText & "")
'        Case -1: TipText = "O.C. Colocada"
'        End Select
    End Select
    DrawText hDC0, TipText, Len(Trim(TipText)), rgnR, DT_CALCRECT + DT_WORDBREAK

    SelectObject hDC0, Old_hFont
    DeleteObject Old_hFont
    ReleaseDC hwnd, hDC0
    b = t + rgnR.bottom + 6
    r = l + rgnR.right + PicW * 2 + 4
End Select

End Sub

Private Sub Form_Activate()
    'dxDBGrid2.Options.Unset (egoShowGroupPanel)
    dxDBGrid2.Filter.FilterActive = False
    txtBusqueda.SetFocus
End Sub

Private Sub Form_Load()

Me.left = 0
Me.top = 0
''''csql = "SELECT TB_CABSOLICITUD.cod_solicitud as Codigo_Solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))) AS Est_Solicitud, TB_CABSOLICITUD.cs_fecha as Fecha_Solicitud, TB_CABSOLICITUD.cs_codsolicitante as Solicitante, TB_CABSOLICITUD.cs_descosto as Centro_de_Costo, IF4ORDEN.F4NUMORD as No_Orden, EF2PROVEEDORES.F2NOMPROV as Nombre_Proveedor, "
''''csql = csql + " IF4ORDEN.F4FECEMI as Fecha_Emision, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda_OC, IF4ORDEN.F4MONTO as Monto_OC, iif(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada', IIF(IF4ORDEN.F4ESTADO=5, 'Anulada',IIF(IF4ORDEN.F4ESTADO=6, 'Pagada' ,IIF(IF4ORDEN.F4ESTADO=7, 'Anticipada', ''))))))) AS Estado_OC, PAG_DCTO.nro_comp as No_Comprobante, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda, IF4ORDEN.F4MONTO as Monto, PAG_DCTO.fch_comp as Fecha, iif(REGISDOC.F4MONEDA='S', 'S/.', IIF(REGISDOC.F4MONEDA='D','$')) as Moneda_Compr, REGISDOC.F4TOTAL as Total, PAG_DCTO.saldo as Saldo, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, "
'''''csql = csql + " IF4ORDEN.F4FECEMI as Fecha_Emision, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda_OC, IF4ORDEN.F4MONTO as Monto_OC, iif(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada','Anulada')))) AS Estado_OC, PAG_DCTO.nro_comp as No_Comprobante, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda, IF4ORDEN.F4MONTO as Monto, PAG_DCTO.fch_comp as Fecha, iif(REGISDOC.F4MONEDA='S', 'S/.', IIF(REGISDOC.F4MONEDA='D','$')) as Moneda_Compr, REGISDOC.F4TOTAL as Total, PAG_DCTO.saldo as Saldo, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, "
''''csql = csql + " iif(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5', iif(IF4ORDEN.F4ESTADO=6,'7', iif(IF4ORDEN.F4ESTADO=7,'7', IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 and IF4ORDEN.F4ESTADO<>2 and IF4ORDEN.F4ESTADO<>3 and IF4ORDEN.F4ESTADO<> 4 and IF4ORDEN.F4ESTADO<> 5  and IF4ORDEN.F4ESTADO<> 6 and IF4ORDEN.F4ESTADO<> 7, '2',iif(tb_cabsolicitud.cs_estado='2','2','1'))))))))) as Estados "
'''''csql = csql + " iif(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5', IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 and IF4ORDEN.F4ESTADO<>2 and IF4ORDEN.F4ESTADO<>3 and IF4ORDEN.F4ESTADO<> 4 and IF4ORDEN.F4ESTADO<> 5, '2',iif(tb_cabsolicitud.cs_estado='2','2','1'))))))) as Estados "
'''''csql = "SELECT TB_CABSOLICITUD.cod_solicitud as Codigo_Solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))) AS Est_Solicitud, TB_CABSOLICITUD.cs_fecha as Fecha_Solicitud, TB_CABSOLICITUD.cs_codsolicitante as Solicitante, TB_CABSOLICITUD.cs_descosto as Centro_de_Costo, IF4ORDEN.F4NUMORD as No_Orden, EF2PROVEEDORES.F2NOMPROV as Nombre_Proveedor, IF4ORDEN.F4FECEMI as Fecha_Emision, iif(IF4ORDEN.F4TIPMON='S', 'S/.', '$') as Moneda_OC, IF4ORDEN.F4MONTO as Monto_OC, IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) AS Estado_OC, PAG_DCTO.nro_comp as No_Comprobante, iif(IF4ORDEN.F4TIPMON='S', 'S/.', '$') as Moneda, IF4ORDEN.F4MONTO as Monto, PAG_DCTO.fch_comp as Fecha, iif(REGISDOC.F4MONEDA='S', 'S/.', '$') as Moneda_Compr, REGISDOC.F4TOTAL as Total, PAG_DCTO.saldo as Saldo, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, '' as Estado "
''''csql = csql + "FROM TB_CABSOLICITUD LEFT JOIN (((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) ON TB_CABSOLICITUD.cod_solicitud = IF4ORDEN.F4CODSOLICITUD"
''''



'---------------------------
csql = "SELECT TB_CABSOLICITUD.cod_solicitud as Codigo_Solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))) AS Est_Solicitud, TB_CABSOLICITUD.cs_fecha as Fecha_Solicitud, TB_CABSOLICITUD.cs_codsolicitante as Solicitante, TB_CABSOLICITUD.cs_descosto as Centro_de_Costo,TB_CABSOLICITUD.cs_codcosto, IF4ORDEN.F4NUMORD as No_Orden, EF2PROVEEDORES.F2NOMPROV as Nombre_Proveedor, "
csql = csql + " IF4ORDEN.F4FECEMI as Fecha_Emision, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda_OC, IF4ORDEN.F4MONTO as Monto_OC, iif(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada', IIF(IF4ORDEN.F4ESTADO=5, 'Anulada',IIF(IF4ORDEN.F4ESTADO=6, 'Pagada' ,IIF(IF4ORDEN.F4ESTADO=7, 'Anticipada', ''))))))) AS Estado_OC, PAG_DCTO.nro_comp as No_Comprobante, iif(IF4ORDEN.F4TIPMON='S', 'S/.', IIF(IF4ORDEN.F4TIPMON='D','$')) as Moneda, IF4ORDEN.F4MONTO as Monto, PAG_DCTO.fch_comp as Fecha, iif(REGISDOC.F4MONEDA='S', 'S/.', IIF(REGISDOC.F4MONEDA='D','$')) as Moneda_Compr, REGISDOC.F4TOTAL as Total, PAG_DCTO.saldo as Saldo, PAG_DCTO.correla, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, "
csql = csql + " iif(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5', iif(IF4ORDEN.F4ESTADO=6,'7', iif(IF4ORDEN.F4ESTADO=7,'7', IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 and IF4ORDEN.F4ESTADO<>2 and IF4ORDEN.F4ESTADO<>3 and IF4ORDEN.F4ESTADO<> 4 and IF4ORDEN.F4ESTADO<> 5  and IF4ORDEN.F4ESTADO<> 6 and IF4ORDEN.F4ESTADO<> 7, '2',iif(tb_cabsolicitud.cs_estado='2','2','1'))))))))) as Estados, IF4ORDEN.F4MONTO-Sum(IF4ORDEN_PAGO.IMPORTE) AS SALDO_OC, iif(mid(pag_dcto.nro_comp, 1, 3)= 'Ant', 'A', '') as Anticipos "
csql = csql + " FROM (TB_CABSOLICITUD LEFT JOIN (((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) ON TB_CABSOLICITUD.cod_solicitud = IF4ORDEN.F4CODSOLICITUD) LEFT JOIN IF4ORDEN_PAGO ON IF4ORDEN.F4NUMORD = IF4ORDEN_PAGO.ORDEN "
csql = csql + " GROUP BY TB_CABSOLICITUD.cod_solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))), TB_CABSOLICITUD.cs_fecha, TB_CABSOLICITUD.cs_codsolicitante, TB_CABSOLICITUD.cs_descosto, TB_CABSOLICITUD.cs_codcosto,IF4ORDEN.F4NUMORD, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4FECEMI, IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')), IF4ORDEN.F4MONTO, IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada',IIf(IF4ORDEN.F4ESTADO=5,'Anulada',IIf(IF4ORDEN.F4ESTADO=6,'Pagada',IIf(IF4ORDEN.F4ESTADO=7,'Anticipada',''))))))), "
csql = csql + " PAG_DCTO.nro_comp, IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')), IF4ORDEN.F4MONTO, PAG_DCTO.fch_comp, IIf(REGISDOC.F4MONEDA='S','S/.',IIf(REGISDOC.F4MONEDA='D','$')), REGISDOC.F4TOTAL, PAG_DCTO.saldo,PAG_DCTO.correla, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, IIf(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5',IIf(IF4ORDEN.F4ESTADO=6,'7',IIf(IF4ORDEN.F4ESTADO=7,'7',IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 And IF4ORDEN.F4ESTADO<>2 And IF4ORDEN.F4ESTADO<>3 And IF4ORDEN.F4ESTADO<>4 And IF4ORDEN.F4ESTADO<>5 And IF4ORDEN.F4ESTADO<>6 And IF4ORDEN.F4ESTADO<>7,'2',IIf(tb_cabsolicitud.cs_estado='2','2','1'))))))))) "
If wObra <> "" Then
    csql = csql + "HAVING (((TB_CABSOLICITUD.cs_codcosto)='" & wObra & "'))"
End If
csql = csql + " UNION ALL "

csql = csql + " SELECT TB_CABSOLICITUD.cod_solicitud AS Codigo_Solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))) AS Est_Solicitud, TB_CABSOLICITUD.cs_fecha AS Fecha_Solicitud, TB_CABSOLICITUD.cs_codsolicitante AS Solicitante, TB_CABSOLICITUD.cs_descosto AS Centro_de_Costo, TB_CABSOLICITUD.cs_codcosto,IF4ORDEN.F4NUMORD AS No_Orden, EF2PROVEEDORES.F2NOMPROV AS Nombre_Proveedor, IF4ORDEN.F4FECEMI AS Fecha_Emision, "
csql = csql + " IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')) AS Moneda_OC, IF4ORDEN.F4MONTO AS Monto_OC, IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada',IIf(IF4ORDEN.F4ESTADO=5,'Anulada',IIf(IF4ORDEN.F4ESTADO=6,'Pagada',IIf(IF4ORDEN.F4ESTADO=7,'Anticipada',''))))))) AS Estado_OC, PAG_DCTO.nro_comp AS No_Comprobante, IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')) AS Moneda, IF4ORDEN.F4MONTO AS Monto, PAG_DCTO.fch_comp AS Fecha, "
csql = csql + " '' AS Moneda_Compr, '' as Total, PAG_DCTO.saldo AS Saldo,PAG_DCTO.correla, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, IIf(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5',IIf(IF4ORDEN.F4ESTADO=6,'7',IIf(IF4ORDEN.F4ESTADO=7,'7',IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 And IF4ORDEN.F4ESTADO<>2 And IF4ORDEN.F4ESTADO<>3 And IF4ORDEN.F4ESTADO<>4 And IF4ORDEN.F4ESTADO<>5 And IF4ORDEN.F4ESTADO<>6 And IF4ORDEN.F4ESTADO<>7,'2',IIf(tb_cabsolicitud.cs_estado='2','2','1'))))))))) AS Estados, IF4ORDEN.F4MONTO-Sum(IF4ORDEN_PAGO.IMPORTE) AS SALDO_OC, IIf(Mid(pag_dcto.nro_comp,1,3)='Ant','A','') AS Anticipos "
csql = csql + " FROM PAG_DCTO RIGHT JOIN ((TB_CABSOLICITUD LEFT JOIN (IF4ORDEN LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) ON TB_CABSOLICITUD.cod_solicitud = IF4ORDEN.F4CODSOLICITUD) LEFT JOIN IF4ORDEN_PAGO ON IF4ORDEN.F4NUMORD = IF4ORDEN_PAGO.ORDEN) ON PAG_DCTO.F4OCOMPRA = IF4ORDEN.F4NUMORD "
csql = csql + " Where mid(Pag_DCTO.nro_comp, 1, 3) = 'Ant'"
csql = csql + " GROUP BY TB_CABSOLICITUD.cod_solicitud, IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Cerrada'))), TB_CABSOLICITUD.cs_fecha, TB_CABSOLICITUD.cs_codsolicitante, TB_CABSOLICITUD.cs_descosto, TB_CABSOLICITUD.cs_codcosto,IF4ORDEN.F4NUMORD, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4FECEMI, IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')), IF4ORDEN.F4MONTO, IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada', "
csql = csql + " IIf(IF4ORDEN.F4ESTADO=5,'Anulada',IIf(IF4ORDEN.F4ESTADO=6,'Pagada',IIf(IF4ORDEN.F4ESTADO=7,'Anticipada',''))))))), PAG_DCTO.nro_comp, IIf(IF4ORDEN.F4TIPMON='S','S/.',IIf(IF4ORDEN.F4TIPMON='D','$')), IF4ORDEN.F4MONTO, PAG_DCTO.fch_comp, PAG_DCTO.saldo,PAG_DCTO.correla, TB_CABSOLICITUD.cs_estado, IF4ORDEN.F4ESTADO, IIf(IF4ORDEN.F4ESTADO=5,'0',IIf(IF4ORDEN.F4ESTADO=3,'6',IIf(IF4ORDEN.F4ESTADO=4,'5',IIf(IF4ORDEN.F4ESTADO=6,'7',IIf(IF4ORDEN.F4ESTADO=7,'7',IIf(IF4ORDEN.F4ESTADO=2,'4',IIf(IF4ORDEN.F4ESTADO=1,'3',IIf(IF4ORDEN.F4ESTADO<>1 And IF4ORDEN.F4ESTADO<>2 And IF4ORDEN.F4ESTADO<>3 And IF4ORDEN.F4ESTADO<>4 And IF4ORDEN.F4ESTADO<>5 And IF4ORDEN.F4ESTADO<>6 And IF4ORDEN.F4ESTADO<>7,'2',IIf(tb_cabsolicitud.cs_estado='2','2','1'))))))))) "
If wObra <> "" Then
    csql = csql + "HAVING (((TB_CABSOLICITUD.cs_codcosto)='" & wObra & "'))"
End If
'csql = csql + " Order By iif(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada', IIF(IF4ORDEN.F4ESTADO=5, 'Anulada', ''))))) DESC "

'csql = csql + " Order By iif(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida',IIf(IF4ORDEN.F4ESTADO=4,'Colocada','Anulada')))) DESC"
'Set Rsj = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
'Do While Not Rsj.EOF
'    If Rsj.Fields("SALDO_OC") > 0 Then
'        a = Rsj.Fields("SALDO_OC")
'    End If
'    Rsj.MoveNext
'Loop
FILL

End Sub
Private Sub FILL()
dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_dbbancos
dxDBGrid2.Dataset.Active = False
dxDBGrid2.Dataset.ADODataset.CommandText = csql
dxDBGrid2.Dataset.Active = True
dxDBGrid2.KeyField = "Codigo_Solicitud"
dxDBGrid2.Dataset.Close
dxDBGrid2.Dataset.Open
dxDBGrid2.Columns.ColumnByFieldName("F4ESTADO").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
'dxDBGrid2.Columns.ColumnByFieldName("Estado_OC").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
dxDBGrid2.Columns.ColumnByFieldName("cs_estado").ImageColumn.Images = dxSideBar.GetImageListByName("dxImageEstado")
End Sub

Private Sub Form_Resize()
'On Error Resume Next
'dxDBGrid2.Move 0, Toolbar1.Height, Me.ScaleWidth, Me.ScaleHeight - (Toolbar1.Height + dxDBGrid2.Height)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim IValue As Byte
IValue = 255
Select Case Trim(Button.Caption)
        Case "Exportar":
'''''''''''''''''''''            csql = "SELECT TB_CABSOLICITUD.cod_solicitud as [Código Solicitud], IIf(TB_CABSOLICITUD.cs_estado='1','Emitida',IIf(TB_CABSOLICITUD.cs_estado='2','Aprobada',IIf(TB_CABSOLICITUD.cs_estado='3','Atendida','Anulada'))) AS [Est Solicitud], TB_CABSOLICITUD.cs_fecha as [Fecha Solicitud], TB_CABSOLICITUD.cs_codsolicitante as Solicitante, TB_CABSOLICITUD.cs_descosto as [Centro de Costo], IF4ORDEN.F4NUMORD as [N° Orden], EF2PROVEEDORES.F2NOMPROV as [Nombre Proveedor], IF4ORDEN.F4FECEMI as [Fecha Emisión], IF4ORDEN.F4TIPMON as [Moneda OC], IF4ORDEN.F4MONTO as [Monto OC], IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) AS [Estado OC], PAG_DCTO.nro_comp as [N° Comprobante], IF4ORDEN.F4TIPMON as Moneda, IF4ORDEN.F4MONTO as Monto, PAG_DCTO.fch_comp as Fecha, REGISDOC.F4MONEDA as [Moneda Compr], REGISDOC.F4TOTAL as Total, PAG_DCTO.saldo as Saldo "
'''''''''''''''''''''            csql = csql + "FROM TB_CABSOLICITUD LEFT JOIN (((IF4ORDEN LEFT JOIN REGISDOC ON IF4ORDEN.F4NUMORD = REGISDOC.F4OCOMPRA) LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) LEFT JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) ON TB_CABSOLICITUD.cod_solicitud = IF4ORDEN.F4CODSOLICITUD"
'''''''''''''''''''''            If chkemitido.Value = 1 And chkaprobado.Value = 0 And chkatendido.Value = 0 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 1 And chkatendido.Value = 0 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Aprobada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 1 And chkatendido.Value = 1 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Aprobada', 'Atendida')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 1 And chkatendido.Value = 0 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Aprobada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 1 And chkatendido.Value = 1 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Aprobada', 'Atendida')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 1 And chkatendido.Value = 1 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Aprobada', 'Atendida', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 0 And chkatendido.Value = 1 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Atendida')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 0 And chkatendido.Value = 1 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Atendida', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 0 And chkatendido.Value = 1 And chkanulado.Value = 0 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Atendida')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 0 And chkatendido.Value = 1 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Atendida', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 0 And chkatendido.Value = 0 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 0 And chkaprobado.Value = 1 And chkatendido.Value = 0 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Aprobada', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 0 And chkatendido.Value = 0 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 1 And chkatendido.Value = 0 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Aprobada', 'Anulada')"
'''''''''''''''''''''            ElseIf chkemitido.Value = 1 And chkaprobado.Value = 1 And chkatendido.Value = 1 And chkanulado.Value = 1 Then
'''''''''''''''''''''                csql = csql + " Where IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) in ('Emitida', 'Aprobada', 'Atendida', 'Anulada')"
'''''''''''''''''''''            End If
'''''''''''''''''''''            csql = csql + " Order By IIf(IF4ORDEN.F4ESTADO=1,'Emitida',IIf(IF4ORDEN.F4ESTADO=2,'Aprobada',IIf(IF4ORDEN.F4ESTADO=3,'Atendida','Anulada'))) DESC"
            If Exportar_Excel(csql, "c:\libro.xLS") Then
               '''''''''''MsgBox "Exportación finalizada", vbInformation
            End If
        Case "R. Parciales"
            'Dim csql        As String
            
            acr_ocsparciales.FldEmpresa.Text = wnomcia
            acr_ocsparciales.fldFecha.Text = Format(Date, "dd/mm/yyyy")
                        
            csql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, IF3ORDEN.F4NUMORD, TB_DETSOLICITUD.ds_cantidad, IF3ORDEN.F3CANPRO, [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') AS saldo FROM TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) AND (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD)"
            csql = csql + "Where [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') <> 0 And f3canpro > 0 And f3canpro < ds_cantidad"
            acr_ocsparciales.DataControl1.ConnectionString = cnn_dbbancos
            acr_ocsparciales.DataControl1.Source = csql
            Set RsS = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
            If RsS.RecordCount > 0 Then
                acr_ocsparciales.Show vbModal
            Else
                MsgBox "No hay registros"
            End If

        Case "Imprimir"
        'Case "ID_ExportExcell"
            'IValue = SSActiveToolBars1.Tools.Item("ID_ExportExcell").UseMaskColor
            dxDBGrid2.Columns.ColumnByFieldName("Est_Solicitud").Visible = True
            dxDBGrid2.Columns.ColumnByFieldName("Estado_OC").Visible = True
            dxDBGrid2.Columns.ColumnByFieldName("cs_estado").Visible = False
            dxDBGrid2.Columns.ColumnByFieldName("F4ESTADO").Visible = False
            
            GridNum = 1: OldValue = 1
            GridInit IValue - 21, OldValue
            OldValue = IValue
            dxDBGrid2.Columns.ColumnByFieldName("Est_Solicitud").Visible = False
            dxDBGrid2.Columns.ColumnByFieldName("Estado_OC").Visible = False
            dxDBGrid2.Columns.ColumnByFieldName("cs_estado").Visible = True
            dxDBGrid2.Columns.ColumnByFieldName("F4ESTADO").Visible = True

        Case "Filtro"
            If dxDBGrid2.Filter.FilterActive = True Then
                dxDBGrid2.Filter.FilterActive = False
                Me.Toolbar1.Buttons.ITEM(6).Image = 3
                Me.Toolbar1.Buttons.ITEM(6).ToolTipText = "Activar Filtro"
            Else
                dxDBGrid2.Filter.FilterActive = True
                Me.Toolbar1.Buttons.ITEM(6).Image = 5
                Me.Toolbar1.Buttons.ITEM(6).ToolTipText = "Desactivar Filtro"
            End If


        Case "Salir": Unload Me
End Select

End Sub
  
Private Function Exportar_Excel(sql As String, sOutputPathXLS As String) As Boolean
       
    On Error GoTo errSub
       
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Variant
    Dim Libro       As Variant
    Dim Hoja        As Variant
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
       
    Me.Enabled = False
       
    Set rec = Af.OpenSQLForwardOnly(sql, cconex_dbbancos)
    If rec.RecordCount = 0 Then
        MsgBox "No hay registros para Exportar.", vbExclamation, "Sistema de Bancos"
        Exit Function
    End If
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
       
    Set Hoja = Libro.Worksheets(1)
       
    Excel.Visible = True: Excel.UserControl = True
    
    iCol = rec.Fields.Count
    
    With Excel.Cells(2, 1)
        Excel.Range(Chr(0 + 65) & 2 & ":" & Chr(17 + 65) & 2).Select
        .value = "Seguimiento de Pedidos"
        With .Font
            .Size = 12
            .Bold = True
        End With
        Excel.Selection.Merge
        With Excel.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            ''''.IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
    End With
    
    
    For iCol = 1 To rec.Fields.Count
        Hoja.Cells(4, iCol).value = rec.Fields(iCol - 1).Name
        With Excel.Cells(4, iCol)
            With .Font
                .Size = 10 ''''''''''''''
                .Bold = True
            End With
            Excel.Range(Chr(iCol + 64) & 4 & "").Select
            With Excel.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
            End With
            With Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With

        End With
    Next
       
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(5, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
           
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
                ElseIf rec.Fields(iCol).Name = "Código Solicitud" Then
                    arrData(iCol, iRow) = "'" & arrData(iCol, iRow)
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
               
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(5, 1).Resize(iRec, rec.Fields.Count).value = GetData(arrData)
    End If
    'Orientacion de la pagina a imprimir en horizontal
    Excel.ActiveSheet.PageSetup.Orientation = xlLandscape
    'Poner escala de impresion en 75%
    Excel.ActiveSheet.PageSetup.Zoom = 75
    'Poner la vista preliminar a forma centrada de manera horizontal
    Excel.ActiveSheet.PageSetup.CenterHorizontally = True

    
    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit
  
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
       
    Set rec = Nothing
    ' -- guardar el libro
    'Libro.SaveAs sOutputPathXLS
    'Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    'Excel.Quit
    Set Excel = Nothing
       
    Exportar_Excel = True
    Me.Enabled = True
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_Excel = False
    Me.Enabled = True
End Function
  
Private Function GetData(vValue As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, t As Variant
       
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
       
    ReDim t(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            t(X, Y) = vValue(Y, X)
        Next Y
    Next X
       
    GetData = t
End Function


Private Sub txtbusqueda_Change()
    dxDBGrid2.Dataset.Filtered = True
    dxDBGrid2.Dataset.Filter = "Codigo_Solicitud LIKE '*" & txtBusqueda.Text & "*' OR " & " No_Orden LIKE '*" & txtBusqueda.Text & "*' "
    'OR f3abrev LIKE '*" & txtbusqueda.Text & "*'  OR numorden LIKE '*" & txtbusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid2.Dataset.Filtered = False
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
            .DialogTitle = "Seguimiento de Pedidos"
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
                Case 234
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
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Sistema de Bancos") = vbYes Then _
                        GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
    End If
    
errhandler:
    
    Exit Sub
 
End Sub

Public Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim i As Byte
    
    If Ind > 199 Then
        SaveTo (Ind)
        Exit Sub
    End If
    
End Sub

Public Function GetGridByActive() As dxDBGrid
    
    Set GetGridByActive = dxDBGrid2
    
End Function

Public Sub proceso2(ByVal Codigo As String)
    Dim sql As String
    'sql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, IF4ORDEN.F4NUMORD, TB_DETSOLICITUD.ds_cantidad, IF3ORDEN.F3CANPRO, [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') AS saldo FROM TB_DETSOLICITUD RIGHT JOIN (IF3ORDEN LEFT JOIN IF4ORDEN ON (IF3ORDEN.F4NUMORD = IF4ORDEN.F4NUMORD) AND (IF3ORDEN.F4LOCAL = IF4ORDEN.F4LOCAL)) ON TB_DETSOLICITUD.cod_solicitud = IF4ORDEN.F4CODSOLICITUD "
    'sql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, IF3ORDEN.F4NUMORD, TB_DETSOLICITUD.ds_cantidad, IF3ORDEN.F3CANPRO, [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') AS saldo FROM TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) AND (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD)"
    
    
'    sql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, IF3ORDEN.F4NUMORD, TB_DETSOLICITUD.ds_cantidad, IF3ORDEN.F3CANPRO, [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') AS saldo FROM (TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) AND (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD)) LEFT JOIN IF4ORDEN ON (IF3ORDEN.F4NUMORD = IF4ORDEN.F4NUMORD) AND (IF3ORDEN.F4LOCAL = IF4ORDEN.F4LOCAL)"
'    sql = sql & " Where TB_DETSOLICITUD.cod_solicitud = '" & Codigo & "' And IF3ORDEN.F4NUMORD = '" & dxDBGrid2.Columns.ColumnByFieldName("No_Orden").Value & "'"
'    Set RsDetalles = Af.OpenSQLForwardOnly(sql, StrConexDbBancos)
'    If Not RsDetalles.EOF Then
'        With dxDBGrid1
'            .Dataset.ADODataset.ConnectionString = cnn_DbBancos
'            .Dataset.Active = False
'            .Dataset.ADODataset.CommandText = sql
'            .Dataset.Active = True
'            .KeyField = "cod_solicitud"
'        End With
'    Else
        sql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, IF3ORDEN.F4NUMORD,IF3ORDEN.F3PRECOS, TB_DETSOLICITUD.ds_cantidad, IF3ORDEN.F3CANPRO, [tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '') AS saldo FROM TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) AND (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD)"
        sql = sql & " Where TB_DETSOLICITUD.cod_solicitud = '" & Codigo & "' " ''And IF3ORDEN.F4NUMORD = '" & dxDBGrid2.Columns.ColumnByFieldName("No_Orden").Value & "'"
        With dxDBGrid1
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = sql
            .Dataset.Active = True
            .KeyField = "cod_solicitud"
        End With
    'End If
End Sub

