VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Lista_RegCompras 
   Caption         =   "                Case 245"
   ClientHeight    =   5055
   ClientLeft      =   4185
   ClientTop       =   4200
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lista_regcompras.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   13200
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CmdSave 
      Left            =   10800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ChkSaldos 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldos por Moneda"
      Height          =   210
      Left            =   8280
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin TabDock.TTabDock TTabDock1 
      Left            =   4320
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo  "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro     "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel    "
            Object.ToolTipText     =   "Exportar a *.xls"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Oficial   "
            ImageIndex      =   9
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox CboMes 
         Height          =   330
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
      Begin VB.ComboBox CboAnno 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   12600
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":1658
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":1BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":218C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":2726
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":2CC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":325A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_regcompras.frx":37F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3390
      Left            =   60
      OleObjectBlob   =   "lista_regcompras.frx":3D8E
      TabIndex        =   0
      Top             =   1320
      Width           =   10005
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   4
      Top             =   420
      Width           =   10005
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mostrar Columnas"
         Height          =   615
         Left            =   7560
         TabIndex        =   7
         Top             =   200
         Width           =   2295
         Begin CONTROLSLibCtl.dxCheckBox dxCheckBox2 
            Height          =   270
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   1050
            _Version        =   65536
            _cx             =   1852
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
            Caption         =   "O.Compra"
            Enabled         =   -1  'True
            AutoSize        =   -1  'True
            BackStyle       =   1
            BackColor       =   16777152
            ForeColor       =   0
            ViewStyle       =   1
            Checked         =   0   'False
            GroupIndex      =   -1
            TextLayout      =   1
            UseMaskColor    =   -1  'True
            MaskColor       =   12632256
         End
         Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
            Height          =   270
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   615
            _Version        =   65536
            _cx             =   1085
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
            Caption         =   "Ruc"
            Enabled         =   -1  'True
            AutoSize        =   -1  'True
            BackStyle       =   1
            BackColor       =   16777152
            ForeColor       =   0
            ViewStyle       =   1
            Checked         =   0   'False
            GroupIndex      =   -1
            TextLayout      =   1
            UseMaskColor    =   -1  'True
            MaskColor       =   12632256
         End
      End
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   7200
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu vc 
         Caption         =   "Vincular a Cheque"
      End
   End
End
Attribute VB_Name = "Lista_RegCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SwLoad As Boolean
Dim StrCn As String
Dim nSaveRecNo As Integer
Dim strNomTemporalCab  As String
Dim strNomTemporalDet  As String
Dim SwProcGridPrincipal As Boolean
Dim Af As New ADOFunctions
Dim Rs As New ADODB.Recordset
Dim EditLookUp  As Boolean
Dim i           As Byte


Dim wdireccion As String, wDistrito As String, wtelefono As String, wfax As String, wPais As String
Dim wwigv As Double, wFob As Double, wDesaduana As Double, wAdela As Double

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

Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const DT_CALCRECT = &H400
Private cImgInfo As cImageInfo

Public Function GetGridByActive() As dxDBGrid

    Set GetGridByActive = dxDBGrid1
    
End Function


Private Sub CboAnno_Click()
If SwProcGridPrincipal = True Then
    proceso
End If
End Sub


Private Sub CboMes_Click()
If SwProcGridPrincipal = True Then
    wmes = Format(CboMes.ListIndex, "00")
    'Call UpdateCaptionMDI(Val(wmes), Val(wanno))
    proceso
End If
End Sub


Private Sub ChkSaldos_Click()
With dxDBGrid1.Columns
    If ChkSaldos.value = 1 Then
        .ColumnByFieldName("SALDOSOL").Visible = True
        .ColumnByFieldName("SALDODOL").Visible = True
        .ColumnByFieldName("SALDO").Visible = False
    Else
        .ColumnByFieldName("SALDOSOL").Visible = False
        .ColumnByFieldName("SALDODOL").Visible = False
        .ColumnByFieldName("SALDO").Visible = True
    End If
End With
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

Private Sub dxDBGrid1_OnKeyPress(Key As Integer)
Dim ss As Integer
Dim UpdReg(0 To 5) As a_grabacion
Dim strRegistro As String
Dim strPeriodo As String

    If Key = 32 Then
        With dxDBGrid1
            If .Columns.ColumnByFieldName("F4VB").Visible = True Then
                .Dataset.Edit
                If .Columns.ColumnByFieldName("F4VB").value = False Then
                    If MsgBox("¿Desea DAR su Vº Bº al registro seleccionado?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                        .Columns.ColumnByFieldName("F4VB").value = True
                        .Columns.ColumnByFieldName("F4VBUSER").value = UCase(wusuario)
                    Else
                        .Columns.ColumnByFieldName("F4VB").value = False
                        .Columns.ColumnByFieldName("F4VBUSER").value = ""
                    End If
                Else
                    If MsgBox("¿Desea QUITAR su Vº Bº al registro seleccionado?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                        .Columns.ColumnByFieldName("F4VB").value = False
                        .Columns.ColumnByFieldName("F4VBUSER").value = ""
                    Else
                        .Columns.ColumnByFieldName("F4VB").value = True
                        '.Columns.ColumnByFieldName("F4VBUSER").Value = wUsuario
                    End If
                End If
                Sw_Act = True
                .Dataset.Post
                Sw_Act = False
                strPeriodo = (Mid(.Columns.ColumnByFieldName("llave").value & "", 1, 6))
                strRegistro = (Mid(.Columns.ColumnByFieldName("llave").value & "", 7))
                UpdReg(0).campo = "F4VB": UpdReg(0).valor = IIf(.Columns.ColumnByFieldName("F4VB").value = True, -1, 0): UpdReg(0).TIPO = "N"
                UpdReg(1).campo = "F4VBUSER": UpdReg(1).valor = .Columns.ColumnByFieldName("F4VBUSER").value & "": UpdReg(1).TIPO = "T"
                GRABA_REGISTRO UpdReg, "REGISDOC", "M", 1, StrConexDbBancos, "F4MESMOV='" & strPeriodo & "' AND F4NUMMOV='" & strRegistro & "'"
                dxDBGrid1.SetFocus
            End If
        End With
    End If
End Sub

Private Sub dxDBGrid1_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, R As Single, b As Single, NeedShowTip As Boolean)
'Dim opValue As Byte
'Dim hDC0 As Long, Old_hFont As Long
'Dim nc As NONCLIENTMETRICS
'Dim rgnR As Rect
'
'If UCase(Column.FieldName) = "F4VB" Then
'    NeedShowTip = True
'
'    rgnR.right = Screen.Width / Screen.TwipsPerPixelX / 4
'    hDC0 = GetDC(0)
'    nc.cbSize = 340 'sizeof(NONCLIENTMETRICS)
'    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, nc, 0
'    Old_hFont = SelectObject(hDC0, CreateFontIndirect(nc.lfStatusFont))
'    TipText = dxDBGrid1.Columns.ColumnByFieldName("F4VBUSER").Value & ""
'    DrawText hDC0, TipText, Len(TipText), rgnR, DT_CALCRECT + DT_WORDBREAK
'
'    SelectObject hDC0, Old_hFont
'    DeleteObject Old_hFont
'    ReleaseDC hwnd, hDC0
'    b = t + rgnR.bottom + 6
'    R = l + rgnR.right + PicW * 2 + 4
'End If
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If Sw_Act = False Then
        dxDBGrid1_RowColChange
    End If
End Sub

Sub dxDBGrid1_RowColChange()
If dxDBGrid1.Dataset.RecordCount > 0 Then
    If dxDBGrid1.Columns.ColumnByFieldName("F4OCOMPRA").value & "" = "+ de 1" Then
        Lista_RegComprasDetalle.dxDBGrid2.Columns.ColumnByFieldName("F3ORDEN").Visible = True
    Else
        Lista_RegComprasDetalle.dxDBGrid2.Columns.ColumnByFieldName("F3ORDEN").Visible = False
    End If
    
    proceso2 dxDBGrid1.Columns.ColumnByFieldName("llave").value & ""
    
End If
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim UpdReg(0 To 5) As a_grabacion
Dim strRegistro As String
Dim strPeriodo As String

    Select Case UCase(Column.FieldName)
    Case "F4VB"
        With dxDBGrid1
            .Dataset.Edit
            If .Columns.ColumnByFieldName("F4VB").value = False Then
                If MsgBox("¿Desea DAR su Vº Bº al registro seleccionado?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                    .Columns.ColumnByFieldName("F4VB").value = True
                    .Columns.ColumnByFieldName("F4VBUSER").value = UCase(wusuario)
                Else
                    .Columns.ColumnByFieldName("F4VB").value = False
                    .Columns.ColumnByFieldName("F4VBUSER").value = ""
                End If
            Else
                If MsgBox("¿Desea QUITAR su Vº Bº al registro seleccionado?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                    .Columns.ColumnByFieldName("F4VB").value = False
                    .Columns.ColumnByFieldName("F4VBUSER").value = ""
                Else
                    .Columns.ColumnByFieldName("F4VB").value = True
                    '.Columns.ColumnByFieldName("F4VBUSER").Value = wUsuario
                End If
            End If
            .Dataset.Post
            strPeriodo = (Mid(.Columns.ColumnByFieldName("llave").value & "", 1, 6))
            strRegistro = (Mid(.Columns.ColumnByFieldName("llave").value & "", 7))
            UpdReg(0).campo = "F4VB": UpdReg(0).valor = IIf(.Columns.ColumnByFieldName("F4VB").value = True, -1, 0): UpdReg(0).TIPO = "N"
            UpdReg(1).campo = "F4VBUSER": UpdReg(1).valor = .Columns.ColumnByFieldName("F4VBUSER").value & "": UpdReg(1).TIPO = "T"
            GRABA_REGISTRO UpdReg, "REGISDOC", "M", 1, StrConexDbBancos, "F4MESMOV='" & strPeriodo & "' AND F4NUMMOV='" & strRegistro & "'"
        End With
    End Select
    
End Sub

Private Sub dxDBGrid1_OnClick()
    dxDBGrid1_RowColChange
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case "SOLES"
    If Val(Text) = 0 Then
        Text = ""
        Color = vbWhite
    Else
        If Val(Text) > 0 Then
            FontColor = RGB(0, 0, 255)
        Else
            FontColor = RGB(255, 0, 0)
        End If
        Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
        Color = &HC0FFFF
    End If
Case "DOLARES"
    If Val(Text) = 0 Then
        Text = ""
        Color = vbWhite
    Else
        If Val(Text) > 0 Then
            FontColor = RGB(0, 0, 255)
        Else
            FontColor = RGB(255, 0, 0)
        End If
        Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
        Color = &HC0FFC0
    End If
Case "SALDO", "SALDOSOL", "SALDODOL"

    Select Case UCase(Column.FieldName)
    Case "SALDOSOL"
        If Node.Values(dxDBGrid1.Columns.ColumnByFieldName("F4MONEDA").Index) = "S" Then
            If Val(Text) = 0 Then
                Color = vbWhite
                FontColor = vbGreen
            Else
                If Val(Text) > 0 Then
                    FontColor = RGB(0, 0, 255)
                Else
                    FontColor = RGB(255, 0, 0)
                End If
                Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
                Color = &HC0FFFF
            End If
            Color = vbYellow
        Else
            Text = ""
            Color = vbWhite
        End If
    Case "SALDODOL"
        If Node.Values(dxDBGrid1.Columns.ColumnByFieldName("F4MONEDA").Index) = "D" Then
            If Val(Text) = 0 Then
                Color = vbWhite
                FontColor = vbYellow
            Else
                If Val(Text) > 0 Then
                    FontColor = RGB(0, 0, 255)
                Else
                    FontColor = RGB(255, 0, 0)
                End If
                Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
                Color = &HC0FFC0
                
            End If
            Color = vbGreen
        Else
            Text = ""
            Color = vbWhite
        End If
    Case "SALDO"
        If Val(Text) = 0 Then
            Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
            FontColor = RGB(0, 255, 0)
        Else
            If Val(Text) > 0 Then
                FontColor = RGB(0, 0, 255)
            Else
                FontColor = RGB(255, 0, 0)
            End If
            Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
            Color = &HFFFF80
            
        End If
    End Select
Case "F4OCOMPRA"
    If Text = "+ de 1" Then
        FontColor = vbGrayed
        Font.Underline = False
    Else
        FontColor = vbBlue
        Font.Underline = True
    End If
End Select
End Sub

Private Sub dxDBGrid1_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case "SOLES", "SALDOSOL"
    If Val(Text) = 0 Then
        'Text = ""
        FontColor = RGB(0, 255, 0)
        Color = vbWhite
    Else
        If Val(Text) > 0 Then
            FontColor = RGB(0, 0, 255)
        Else
            FontColor = RGB(255, 0, 0)
        End If
        Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
        If UCase(Column.FieldName) = "SOLES" Then
            Color = &HC0FFFF
        Else
            Color = vbYellow
        End If
    End If
Case "DOLARES", "SALDODOL"
    If Val(Text) = 0 Then
        'Text = ""
        FontColor = RGB(0, 255, 0)
        Color = vbWhite
    Else
        If Val(Text) > 0 Then
            FontColor = RGB(0, 0, 255)
        Else
            FontColor = RGB(255, 0, 0)
        End If
        Text = Format(Text, "###,###,##0.00;(###,###,##0.00)")
        If UCase(Column.FieldName) = "DOLARES" Then
            Color = &HC0FFC0
        Else
            Color = vbGreen
        End If
    End If
End Select
End Sub

Private Sub dxDBGrid1_OnDblClick()
Dim d As Integer
If dxDBGrid1.Dataset.RecordCount > 0 And dxDBGrid1.Columns.ColumnByFieldName("f4ocompra").value & "" <> "+ de 1" Then
    If UCase(dxDBGrid1.Columns.FocusedColumn.FieldName) = "F4OCOMPRA" Then
        CargaOrdenDeCompra dxDBGrid1.Columns.ColumnByFieldName("f4ocompra").value & ""
    Else
        For d = 0 To 25
            nSaveRecNo = dxDBGrid1.Dataset.RecNo
        Next
        If CboMes.ListIndex = 0 Then
            wmes = right(dxDBGrid1.Columns.ColumnByFieldName("f4mesmov").value & "", 2)
        End If
        sw_nuevo_documento = False
        Me.MousePointer = vbHourglass
        sw_nuevo_doc = False
        FrmName = Me.Name
        Registro_Compras.Periodo = "" & dxDBGrid1.Columns.ColumnByFieldName("F4MESMOV").value
        Registro_Compras.registro = "" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMMOV").value
        dxDBGrid1.Dataset.Close
        Lista_RegComprasDetalle.dxDBGrid2.Dataset.Close
        Me.Hide
        Registro_Compras.Show 1
        Unload Registro_Compras
        Set Registro_Compras = Nothing
        proceso
        Me.Show
        If dxDBGrid1.Dataset.RecordCount >= nSaveRecNo Then
            dxDBGrid1.Dataset.RecNo = nSaveRecNo
        End If
        Me.MousePointer = vbDefault
    End If
End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
Case 13
    dxDBGrid1_OnDblClick
Case 116
    proceso
End Select
End Sub



Private Sub Form_Activate()
    
    If dxDBGrid1.Filter.FilterActive = True Then dxDBGrid1.Filter.FilterActive = False

'    proceso
'    dxDBGrid1.Options.Unset (egoShowGroupPanel)
'    dxDBGrid1.Filter.FilterActive = False
'    dxDBGrid1_RowColChange
    
End Sub
Private Sub CargaComboAnnos()
SqlCad = "SELECT Left(F4MESMOV,4) AS Anno "
SqlCad = SqlCad & "From REGISDOC "
SqlCad = SqlCad & "GROUP BY Left(F4MESMOV,4)"
SqlCad = SqlCad & "order BY Left(F4MESMOV,4)"
Set Rs = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
CboAnno.Clear
If Rs.RecordCount > 0 Then Rs.MoveFirst
CboAnno.AddItem Trim("Todos")
Do While Not Rs.EOF
    CboAnno.AddItem Trim(Rs!anno)
    Rs.MoveNext
Loop
Rs.Filter = ""
Rs.Filter = "anno='" & wanno & "'"
If Rs.RecordCount = 0 Then
    CboAnno.AddItem Trim(wanno)
End If
If Rs.State = 1 Then Rs.Close
Set Rs = Nothing
Call SeleccionaEnComboLeft(wanno, CboAnno)
End Sub
Private Sub CargaComboMeses()
CboMes.Clear
For i = 0 To 12
    CboMes.AddItem dev_mes(i) & Space(100) & Format(i, "00")
Next

End Sub

Private Sub Form_GotFocus()
Form_Activate
End Sub



Private Sub Form_Load()

''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4TIPDOC", "String", False, ""
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VB1", "YESNO", False, "False"
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VBUSER1", "STRING", False, ""
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VBFECHA1", "DATE", False, ""
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VB2", "YESNO", False, "False"
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VBUSER2", "STRING", False, ""
''    Crea_Campo StrConexDbBancos, "IF4ORDEN", "F4VBFECHA2", "DATE", False, ""
    
    Sw_Act = False
    StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "templus.mdb;Persist Security Info=False" '"Tmp_Bancos.mdb;Persist Security Info=False"
    strNomTemporalCab = ""
    strNomTemporalDet = ""
    SwProcGridPrincipal = False
    CargaComboAnnos
    CargaComboMeses
    SwProcGridPrincipal = True
    CboMes.ListIndex = Month(Date)
    
    'verifica permiso de dar vºbº
    dxDBGrid1.Columns.ColumnByFieldName("f4vb").Visible = (VerificaPermiso("0003", wusuario))
    
    Me.MousePointer = vbHourglass
    Me.AutoRedraw = False
    
    
    'wmes = Format(Month(Date), "00")
    
    Me.left = 0
    Me.top = 0
    
    '-----------
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    'Call UpdateCaptionMDI(Val(wmes), Val(wanno))
    TTabDock1.AddForm Lista_RegComprasDetalle, tdDocked, tdAlignBottom, "Lista_RegComprasDetalle"
    TTabDock1.DockedForms.ITEM("Lista_RegComprasDetalle").Panel.Height = 2500
    TTabDock1.FormShow "Lista_RegComprasDetalle"
    Me.MousePointer = vbDefault
End Sub



Private Sub Form_Resize()
On Error Resume Next
    fraBusqueda.Move 0, Toolbar.Height, Me.ScaleWidth, 870
    'txtbusqueda.Width = FraBusqueda.Width - 350
    dxDBGrid1.Move 0, Toolbar.Height + fraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (Toolbar.Height + fraBusqueda.Height + TTabDock1.DockedForms.ITEM("Lista_RegComprasDetalle").Panel.Height)
    ChkSaldos.Alignment = 0
    ChkSaldos.top = dxDBGrid1.Height + dxDBGrid1.top - 250
    ChkSaldos.left = 250

End Sub

Private Sub Form_Unload(Cancel As Integer)
wtipoguia = ""
If Len(Trim(strNomTemporalCab)) > 0 Then
    dxDBGrid1.Dataset.Close
    csql = "drop table " & strNomTemporalCab
    EJECUTA_SENTENCIA csql, StrCn
    Lista_RegComprasDetalle.dxDBGrid2.Dataset.Close
    csql = "drop table " & strNomTemporalDet
    EJECUTA_SENTENCIA csql, StrCn
End If
'If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close

End Sub



Public Sub proceso()
Dim SqlCad     As String
MousePointer = 11
'elimina anterior
If Len(Trim(strNomTemporalCab)) > 0 Then
    dxDBGrid1.Dataset.Close
    csql = "drop table " & strNomTemporalCab
    EJECUTA_SENTENCIA csql, StrCn
    Lista_RegComprasDetalle.dxDBGrid2.Dataset.Close
    csql = "drop table " & strNomTemporalDet
    EJECUTA_SENTENCIA csql, StrCn
End If
'crea temporal
strNomTemporalCab = "RC_CAB_" & Format(Now, "ddmmyyyy") & "_" & Format(Time, "hhmmss")
strNomTemporalDet = "RC_DET_" & Format(Now, "ddmmyyyy") & "_" & Format(Time, "hhmmss")
SqlCad = "SELECT PAG_DCTO.deb_hab,REGISDOC.F4MESMOV,CENTROS.F3ABREV, REGISDOC.F4FECHA, REGISDOC.F4NUMMOV, REGISDOC.F4CODPRV, "
SqlCad = SqlCad & "REGISDOC.F4RUCPRV, REGISDOC.F4FECHAREC,REGISDOC.F4MONEDA, REGISDOC.F4TIPCAM, EF2PROVEEDORES.F2NOMPROV as F4NOMPRV, DOCUMENTOS.F2ABREV AS TIPDOC, "
SqlCad = SqlCad & "REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, "
SqlCad = SqlCad & "IIf(REGISDOC.F4MONEDA='S',iif(PAG_DCTO.deb_hab='H',REGISDOC.F4TOTAL,REGISDOC.F4TOTAL*-1),0) AS SOLES, "
SqlCad = SqlCad & "IIf(REGISDOC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='H',REGISDOC.F4TOTAL,REGISDOC.F4TOTAL*-1),0) AS DOLARES, "
SqlCad = SqlCad & "iif(PAG_DCTO.deb_hab='H',PAG_DCTO.saldo,PAG_DCTO.saldo*-1) AS SALDO, "
SqlCad = SqlCad & "IIf(REGISDOC.F4MONEDA='S',iif(PAG_DCTO.deb_hab='H',PAG_DCTO.saldo,PAG_DCTO.saldo*-1),0) AS SALDOSOL, "
SqlCad = SqlCad & "IIf(REGISDOC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='H',PAG_DCTO.saldo,PAG_DCTO.saldo*-1),0) AS SALDODOL, "
SqlCad = SqlCad & "REGISDOC.F4BASIMP, REGISDOC.F4MONINA, "
SqlCad = SqlCad & "REGISDOC.F4IGV,REGISDOC.F4OCOMPRA,REGISDOC.F4VB,REGISDOC.F4VBUSER,REGISDOC.F4MESMOV+REGISDOC.F4NUMMOV AS LLAVE "
SqlCad = SqlCad & "INTO " & strNomTemporalCab & " IN '" & wrutatemp & "templus.mdb" & "' " '"tmp_bancos.mdb" & "' "
SqlCad = SqlCad & "FROM (((REGISDOC LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) LEFT JOIN DOCUMENTOS ON REGISDOC.F4TIPDOC = DOCUMENTOS.F2CODDOC) LEFT JOIN CENTROS ON REGISDOC.F4OBRA = CENTROS.F3COSTO) LEFT JOIN EF2PROVEEDORES ON REGISDOC.F4CODPRV = EF2PROVEEDORES.F2CODPROV "
If CboMes.ListIndex <> 0 Or CboAnno.ListIndex <> 0 Then
    If CboMes.ListIndex = 0 And CboAnno.ListIndex <> 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV LIKE '" & CboAnno.Text & "%' "
        If Toolbar.Buttons(14).value = tbrPressed Then
            SqlCad = SqlCad & " AND DOCUMENTOS.F2OFICIAL=-1 "
        End If
    ElseIf CboMes.ListIndex <> 0 And CboAnno.ListIndex = 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV LIKE '%" & wmes & "' "
        If Toolbar.Buttons(14).value = tbrPressed Then
            SqlCad = SqlCad & " AND DOCUMENTOS.F2OFICIAL=-1 "
        End If
    ElseIf CboMes.ListIndex <> 0 And CboAnno.ListIndex <> 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV = '" & CboAnno.Text & wmes & "' "
        If Toolbar.Buttons(14).value = tbrPressed Then
            SqlCad = SqlCad & " AND DOCUMENTOS.F2OFICIAL=-1 "
        End If
    End If
End If
EJECUTA_SENTENCIA SqlCad, StrConexDbBancos
'*****************************************
SqlCad = "SELECT * FROM " & strNomTemporalCab & " ORDER BY LLAVE DESC"
With dxDBGrid1
    .Dataset.Active = False
    .Dataset.ADODataset.ConnectionString = StrCn
    .Dataset.ADODataset.CommandText = SqlCad
    .Dataset.Active = True
    .KeyField = "LLAVE"
End With
'VERIFICA COLUMNAS DE SALDOS
ChkSaldos_Click
'crea temporal 2
SqlCad = "SELECT f4mesmov+f4nummov+format(f3item,'000') as llave,F4MESMOV,f4nummov,F3ITEM, f5codpro,F3GASTO, F3CTACON, F3CONCEPTO, F3IMPORTE, F3ORDEN "
SqlCad = SqlCad & "INTO " & strNomTemporalDet & " IN '" & wrutatemp & "templus.mdb" & "' " '"tmp_bancos.mdb" & "' "
SqlCad = SqlCad & "from REGISMOV "
If CboMes.ListIndex <> 0 Or CboAnno.ListIndex <> 0 Then
    If CboMes.ListIndex = 0 And CboAnno.ListIndex <> 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV LIKE '" & CboAnno.Text & "%' "
    ElseIf CboMes.ListIndex <> 0 And CboAnno.ListIndex = 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV LIKE '%" & wmes & "' "
    ElseIf CboMes.ListIndex <> 0 And CboAnno.ListIndex <> 0 Then
        SqlCad = SqlCad & "WHERE F4MESMOV = '" & CboAnno.Text & wmes & "' "
    End If
End If

EJECUTA_SENTENCIA SqlCad, StrConexDbBancos
'*****************************************
SqlCad = "SELECT * FROM " & strNomTemporalDet & " ORDER BY LLAVE"
With Lista_RegComprasDetalle.dxDBGrid2
    .Dataset.Active = False
    .Dataset.ADODataset.ConnectionString = StrCn
    .Dataset.ADODataset.CommandText = SqlCad
    .Dataset.Active = True
    .KeyField = "llave"
End With
'******verifica permiso
dxDBGrid1.Columns.ColumnByFieldName("f4vb").Visible = VerificaPermiso("0004", wusuario)
'**************
MousePointer = 0
    
dxDBGrid1_RowColChange
'dxDBGrid1.SetFocus
End Sub

Public Sub proceso2(ByVal pRegistro As String)

    Dim SqlCad As String

MousePointer = 11
    
    Lista_RegComprasDetalle.dxDBGrid2.Dataset.Filtered = True
    Lista_RegComprasDetalle.dxDBGrid2.Dataset.Filter = "llave like '" & pRegistro & "*'"
    
MousePointer = 0

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

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim o_Excel As Excel.Application
Dim FileName As String

Select Case Trim(Button.Caption)
    Case "Nuevo"
        sw_nuevo_documento = True
        sw_nuevo_doc = True
        Me.MousePointer = vbHourglass
        Me.Hide
        Registro_Compras.Show 1
        Unload Registro_Compras
        Set Registro_Compras = Nothing
        proceso
        Me.Show
        Me.MousePointer = vbDefault
    Case "Filtro"
        If Button.value = tbrUnpressed Then
            dxDBGrid1.Filter.FilterActive = False
            Me.Toolbar.Buttons.ITEM(9).Image = 3
            Me.Toolbar.Buttons.ITEM(9).ToolTipText = "Activar Filtro"
        Else
            dxDBGrid1.Filter.FilterActive = True
            Me.Toolbar.Buttons.ITEM(9).Image = 6
            Me.Toolbar.Buttons.ITEM(9).ToolTipText = "Desactivar Filtro"
        End If

    Case "Agrupar"
        If Button.value = tbrUnpressed Then
            dxDBGrid1.Options.Unset (egoShowGroupPanel)
            Me.Toolbar.Buttons.ITEM(10).Image = 4
            Me.Toolbar.Buttons.ITEM(10).ToolTipText = "Agrupar Columnas"
        Else
            dxDBGrid1.Options.Set (egoShowGroupPanel)
            Me.Toolbar.Buttons.ITEM(10).Image = 7
            Me.Toolbar.Buttons.ITEM(10).ToolTipText = "Desagrupar Columnas"

        End If
    Case "Imprimir"
            'GENERA_TEMP
            Imprime_RegCompras.Show 1
    Case "Excel"
    
        If Len(Trim(txtBusqueda.Text)) > 0 Then
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
    
        Else
            RutaReporte.TipoFile = 1
            Load RutaReporte
            strFilePath = RutaReporte.Ruta
            Unload RutaReporte
            Set RutaReporte = Nothing
            SqlCad = ""
            Me.MousePointer = vbHourglass
            If Val(wmes) = 0 Then
                SqlCad = ""
             
                        SqlCad = SqlCad & "SELECT REGISDOC.F4MESMOV As Periodo, REGISDOC.F4NUMMOV AS Registro, REGISDOC.F4RUCPRV as Ruc, REGISDOC.F4TIPDOC as Doc_Tipo, "
                        SqlCad = SqlCad & "REGISDOC.F4SERDOC as Doc_Serie, REGISDOC.F4NUMDOC as Doc_Numero, REGISDOC.F4FECHA as Fecha, EF2PROVEEDORES.F2NOMPROV as Proveedor, "
                        SqlCad = SqlCad & "REGISDOC.F4BASIMP as Base_Imponible,  REGISDOC.F4MONINA as Monto_Inafecto, REGISDOC.F4IGV as IGV, REGISDOC.F4TOTAL as Total, REGISDOC.F4TIPCAM as Tipo_Cambio,REGISDOC.F4OBRA AS Centro_Costo, REGISDOC.F4REFERE as Referencia "
                        SqlCad = SqlCad & "FROM (REGISDOC LEFT JOIN EF2PROVEEDORES ON REGISDOC.F4CODPRV = EF2PROVEEDORES.F2CODPROV) "
                        SqlCad = SqlCad & "LEFT JOIN CENTROS ON REGISDOC.F4OBRA = CENTROS.F3COSTO "
                        SqlCad = SqlCad & "WHERE REGISDOC.F4MESMOV like '" & CboAnno.Text & "%'"
                        SqlCad = SqlCad & "ORDER BY REGISDOC.F4MESMOV DESC, REGISDOC.F4NUMMOV DESC "
                    
                
            Else
                SqlCad = ""
                SqlCad = SqlCad & "SELECT REGISDOC.F4MESMOV AS Periodo, REGISDOC.F4NUMMOV AS Registro, REGISDOC.F4RUCPRV as Ruc, REGISDOC.F4TIPDOC as Doc_Tipo, "
                SqlCad = SqlCad & "REGISDOC.F4SERDOC as Doc_Serie, REGISDOC.F4NUMDOC as Doc_Numero, REGISDOC.F4FECHA as Fecha, EF2PROVEEDORES.F2NOMPROV as Proveedor, "
                SqlCad = SqlCad & "REGISDOC.F4BASIMP as Base_Imponible,  REGISDOC.F4MONINA as Monto_Inafecto, REGISDOC.F4IGV as IGV, REGISDOC.F4TOTAL as Total, REGISDOC.F4TIPCAM as Tipo_Cambio,REGISDOC.F4OBRA AS Centro_Costo, REGISDOC.F4REFERE as Referencia "
                SqlCad = SqlCad & "FROM (REGISDOC LEFT JOIN EF2PROVEEDORES ON REGISDOC.F4CODPRV = EF2PROVEEDORES.F2CODPROV) "
                SqlCad = SqlCad & "LEFT JOIN CENTROS ON REGISDOC.F4OBRA = CENTROS.F3COSTO "
                SqlCad = SqlCad & "WHERE (((REGISDOC.F4MESMOV)='" & CboAnno.Text & wmes & "')) "
                SqlCad = SqlCad & "ORDER BY REGISDOC.F4MESMOV DESC, REGISDOC.F4NUMMOV DESC "
            End If
            Set Rs = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
            If Trim(strFilePath) <> "" Then
                ExportaRecordsetToExcel Rs, strFilePath
            End If
            Me.MousePointer = vbDefault
        End If
    Case "Oficial"
        wmes = Format(CboMes.ListIndex, "00")
        'Call UpdateCaptionMDI(Val(wmes), Val(wanno))
        proceso
    Case "Salir"
        Unload Me
End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Me.Toolbar.Buttons.ITEM(2).Caption = ButtonMenu.Text
wmes = Format(ButtonMenu.Index, "00")

Call proceso
End Sub

Private Sub TTabDock1_PanelResize(ByVal Panel As TabDock.TTabDockHost)
Form_Resize
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    'dxDBGrid1.Dataset.Filter = "F4FECHA LIKE '*" & txtbusqueda.Text & "*' OR " & " F4NUMMOV LIKE '*" & txtbusqueda.Text & "*' OR " & " F4CODPRV LIKE '*" & txtbusqueda.Text & "*' OR " & " F4RUCPRV LIKE '*" & txtbusqueda.Text & "*' OR " & " F4NOMPRV LIKE '*" & txtbusqueda.Text & "*' OR " & " F4TIPDOC LIKE '*" & txtbusqueda.Text & "*' OR " & "F4SERDOC LIKE '*" & txtbusqueda.Text & "*' OR " & "F4NUMDOC LIKE '*" & txtbusqueda.Text & "*'"
    dxDBGrid1.Dataset.Filter = "F4NOMPRV LIKE '*" & txtBusqueda.Text & "*' OR " & "F4NUMDOC LIKE '*" & txtBusqueda.Text & "*'"
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub
Private Sub GENERA_TEMP()
Dim dbtempo As DAO.Database
Dim tbtempo As DAO.Recordset
Dim sw_cambio_igv   As Boolean

    If dxDBGrid1.Dataset.RecordCount > 0 Then
        'ctitulo = "Mes de " & wmes & " de " & wAnno
        Set dbtempo = OpenDatabase(wrutatemp & "\templus.mdb") '"\temp_com.mdb")
        dbtempo.Execute ("delete * from temp_gg")
        Set tbtempo = dbtempo.OpenRecordset("temp_gg")
        
        dxDBGrid1.Dataset.First
        Do While Not dxDBGrid1.Dataset.EOF
            tbtempo.AddNew
            tbtempo.Fields("temp_mov") = dxDBGrid1.Dataset.FieldValues("f4nummov")
            tbtempo.Fields("temp_fecha") = dxDBGrid1.Dataset.FieldValues("f4fecha")
            tbtempo.Fields("TEMP_TDOC") = left(Format(dxDBGrid1.Dataset.FieldValues("TIPDOC"), "00"), 2)
            tbtempo.Fields("temp_serie") = Format(dxDBGrid1.Dataset.FieldValues("f4serdoc"), "000")
            tbtempo.Fields("temp_docum") = Format(dxDBGrid1.Dataset.FieldValues("f4numdoc"), "0000000")
            tbtempo.Fields("temp_prov") = dxDBGrid1.Dataset.FieldValues("f4nomprv") & ""
            'tbtempo.Fields("temp_detal") = dxDBGrid1.Dataset.FieldValues("f4refere") & ""
            If Format(dxDBGrid1.Dataset.FieldValues("TIPDOC"), "00") = "02" Then
                tbtempo.Fields("TEMP_BIMP") = 0#
                tbtempo.Fields("TEMP_EXON") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP")) + Val("" & dxDBGrid1.Dataset.FieldValues("F4MONINA"))
                tbtempo.Fields("temp_totals") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP")) + Val("" & dxDBGrid1.Dataset.FieldValues("F4MONINA"))
            Else
                'sw_cambio_igv = False
                'If wmesregcompras >= wf1mescambio_igv Then
                '    If Month(dxDBGrid1.Dataset.FieldVALUES("f4fecha")) < Val(wf1mescambio_igv) Then
                '        sw_cambio_igv = True
                '    Else
                '        sw_cambio_igv = False
                '    End If
                'Else
                '    sw_cambio_igv = False
                'End If
                
                'If wmesregcompras >= wf1mescambio_igv Then
                '    If Val(dxDBGrid1.Dataset.FieldVALUES("F4PORC_IGV") & "") <> gigv Then
                '        sw_cambio_igv = True
                '    Else
                '        sw_cambio_igv = False
                '    End If
                'Else
                '    sw_cambio_igv = False
                'End If
                
                'If sw_cambio_igv = False Then
                    'Select Case dxDBGrid1.Dataset.FieldValues("F4CODIGV") & ""
                    '    Case "001":
                            tbtempo.Fields("TEMP_BIMP") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP"))
                            tbtempo.Fields("temp_igvs") = Val("" & dxDBGrid1.Dataset.FieldValues("f4igv"))
                            tbtempo.Fields("TEMP_EXON") = Val("" & dxDBGrid1.Dataset.FieldValues("F4MONINA"))
                    '    Case "002":
                    '        tbtempo.Fields("TEMP_BIMP_GYNG") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP"))
                    '        tbtempo.Fields("TEMP_IGVS_GYNG") = Val("" & dxDBGrid1.Dataset.FieldValues("f4igv"))
                    '    Case "003":
                    '        tbtempo.Fields("TEMP_BIMP_SIN") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP"))
                    '        tbtempo.Fields("TEMP_IGVS_SIN") = Val("" & dxDBGrid1.Dataset.FieldValues("f4igv"))
                    'End Select
                'Else
                    'tbtempo.Fields("TEMP_BIMP_OTRO") = Val("" & dxDBGrid1.Dataset.FieldValues("F4BASIMP"))
                    'tbtempo.Fields("TEMP_IGVS_OTRO") = Val("" & dxDBGrid1.Dataset.FieldValues("f4igv"))
                'End If
                
                'tbtempo.Fields("TEMP_EXON") = Val("" & dxDBGrid1.Dataset.FieldValues("F4OTRIMP")) + Val("" & dxDBGrid1.Dataset.FieldValues("F4MONINA")) + Val("" & dxDBGrid1.Dataset.FieldValues("F4REDSUMA")) - Val("" & dxDBGrid1.Dataset.FieldValues("F4REDRESTA")) - Val("" & dxDBGrid1.Dataset.FieldValues("F4DCTO"))
                tbtempo.Fields("temp_totals") = Val("" & dxDBGrid1.Dataset.FieldValues("SOLES"))
                tbtempo.Fields("temp_totalD") = Val("" & dxDBGrid1.Dataset.FieldValues("DOLARES"))
                tbtempo.Fields("temp_IGVD") = Val("" & dxDBGrid1.Dataset.FieldValues("F4TIPCAM"))
            End If
            'tbtempo.Fields("TEMP_RUC") = traerCampo("EF2PROVEEDORES", "F2NEWRUC", "F2TIPPROV", dxDBGrid1.Dataset.FieldValues("F4RUCPRV") & "")
            tbtempo.Fields("empresa") = wempresa
            'tbtempo.Fields("temp_moneda") = IIf(dxDBGrid1.Dataset.FieldValues("f4moneda") & "" = "S", "S/.", "US$")
            tbtempo.Fields("TEMP_RUC") = dxDBGrid1.Dataset.FieldValues("F4RUCPRV") & ""
            tbtempo.Fields("MES") = "Mes de " & wmes & " de " & wanno
            'tbtempo.Fields("TEMP_CODIGV") = dxDBGrid1.Dataset.FieldValues("F4CODIGV") & ""
            'tbtempo.Fields("TEMP_NUMDEPOSITO") = Trim(dxDBGrid1.Dataset.FieldValues("F4NUMDEPOSITO") & "")
            'tbtempo.Fields("TEMP_FECHADEPOSITO") = dxDBGrid1.Dataset.FieldValues("F4FECHADEPOSITO")
            tbtempo.Update
            dxDBGrid1.Dataset.Next
        Loop
        tbtempo.Close
        dbtempo.Close
    Else
        MsgBox "No hay registros.", 48, "Atención"
    End If

End Sub

