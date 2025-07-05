VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Begin VB.Form lista_solicitudes 
   Caption         =   "Listado de Requerimientos"
   ClientHeight    =   9135
   ClientLeft      =   10170
   ClientTop       =   1770
   ClientWidth     =   10185
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
   ScaleHeight     =   9135
   ScaleWidth      =   10185
   WindowState     =   2  'Maximized
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
      ToolsCount      =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "lista_solicitudes.frx":5B3F
      ToolBars        =   "lista_solicitudes.frx":D950
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
      OleObjectBlob   =   "lista_solicitudes.frx":DAEE
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
    Dim valor As String
    Dim TIPO As String
    valor = dxDBGrid1.Columns.ColumnByFieldName("cod_solicitud").Value
    TIPO = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
    Call proceso2(valor, TIPO)
End Sub

Public Sub proceso2(ByVal Codigo As String, TIPO As String)
    Dim sql As String
    With lista_solicitudes_detalle.dxDBGrid2
        If .Dataset.State = 1 Then .Dataset.Close
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        sql = "SELECT ds.item, ds.cod_producto, ds.f5codfab, ds.ds_descripcion, "
        sql = sql & "ds.ds_unidmed, ds.ds_cantidad, ds.proveedor, ds.centro,ds.observa "
        sql = sql & "FROM tb_detsolicitud AS ds "
        sql = sql & "WHERE ds.cod_solicitud='" & Codigo & "' AND ds.cs_documento='" & TIPO & "' "
        sql = sql & "ORDER BY ds.item, ds.cod_solicitud"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .KeyField = "item"
    End With
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim StrPregunta As String
Dim SwAprobacion As Boolean
Dim StrEstado As String
Select Case UCase(dxDBGrid1.Columns.FocusedColumn.FieldName)
Case "VBJEFECC"
    If dxDBGrid1.Columns.ColumnByFieldName("vbjefecc").Value = False Then
            StrPregunta = "¿ Desea aprobar el Requerimiento ?"
            SwAprobacion = False
    Else
            StrPregunta = "¿Desea quitar la aprobación del Requerimiento?"
            SwAprobacion = True
    End If
    If dxDBGrid1.Columns.ColumnByFieldName("CS_ESTado").Value <= "2" Then
    'Or dxDBGrid1.Columns.ColumnByFieldName("CS_ESTado").Value = "2" Then
        
        dxDBGrid1.Dataset.Edit
        If MsgBox(StrPregunta, vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            dxDBGrid1.Columns.FocusedColumn.Value = Not SwAprobacion
        Else
            dxDBGrid1.Columns.FocusedColumn.Value = SwAprobacion
        End If
          
        If dxDBGrid1.Columns.ColumnByFieldName("VBJEFECC").Value = True Then
            'csql = "UPDATE TB_CABSOLICITUD SET CS_ESTADO='2' WHERE COD_SOLICITUD='" & dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value & "'"
            StrEstado = "2"
        Else
            'csql = "UPDATE TB_CABSOLICITUD SET CS_ESTADO='1' WHERE COD_SOLICITUD='" & dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value & "'"
            If dxDBGrid1.Columns.ColumnByFieldName("CS_ESTado").Value = "0" Then
                StrEstado = "0"
            Else
                StrEstado = "1"
            End If
        End If
        dxDBGrid1.Columns.ColumnByFieldName("CS_ESTado").Value = StrEstado
        If StrEstado = "2" Then
            dxDBGrid1.Columns.ColumnByFieldName("CS_APROBADOX").Value = wusuario
        Else
            dxDBGrid1.Columns.ColumnByFieldName("CS_APROBADOX").Value = ""
        End If
        dxDBGrid1.Dataset.Post
        dxDBGrid1.Dataset.ADODataset.Requery
        'cnn_dbbancos.Execute csql
    Else
        Select Case dxDBGrid1.Columns.ColumnByFieldName("CS_ESTado").Value & ""
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
Dim valor As String
    Dim TIPO As String
    valor = dxDBGrid1.Columns.ColumnByFieldName("cod_solicitud").Value
    TIPO = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
    Call proceso2(valor, TIPO)
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
    ReleaseDC hwnd, hDC0
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
    Me.MousePointer = 11
    'Me.AutoRedraw = False
    Me.Height = 7500
    Me.Width = 10400
    Me.left = 0
    Me.top = 50
    sw_nuevo_documento = True
    'Me.AutoRedraw = True
    verifica_mysql
    proceso
    
    TTabDock1.AddForm lista_solicitudes_detalle, tdDocked, tdAlignBottom, "lista_solicitudes_detalle"
    TTabDock1.DockedForms.ITEM("lista_solicitudes_detalle").Panel.Height = 2500
    TTabDock1.FormShow "lista_solicitudes_detalle"

    Me.MousePointer = 1
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_NuevaOC"
            sw_nuevo_documento = True
            wTipoReq = 1
            Me.MousePointer = 11
            solicitud.Show 1
            Me.MousePointer = 1
        Case "ID_Nueva/O.S."
            sw_nuevo_documento = True
            wTipoReq = 2
            Me.MousePointer = 11
            solicitud.Show 1
            '2102
            Me.MousePointer = 1
        Case "ID_Actualizar"
            Me.MousePointer = 11
            'verifica_mysql
            Me.dxDBGrid1.Dataset.ADODataset.Requery
            Me.dxDBGrid1.Dataset.Refresh
            Me.MousePointer = 1
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

        Case "ID_Salir"
            Unload Me
    
    End Select
    
End Sub

Private Sub dxDBGrid1_OnDblClick()
    If dxDBGrid1.Dataset.RecordCount > 0 Then
        sw_nuevo_documento = False
        Me.MousePointer = 11
        solicitud.Show 1
        Unload solicitud
        Set solicitud = Nothing
        Me.MousePointer = 1
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
        
        csql = "SELECT csol.cs_documento,csol.cod_solicitud, UCase(csol.cs_observaciones) AS glosa, csol.cs_fecha, csol.cs_prioridad, csol.cs_estado, "
        csql = csql & "csol.VBFECHA,csol.Fecha_graba, csol.Fecha_modifica, csol.cs_motivos,EF2USERS.F2NOMUSER, CENTROS.F3ABREV, csol.numorden,csol.VBJEFECC,csol.cs_aprobadox "
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
        
        csql = csql & "ORDER BY csol.cs_fecha desc,csol.cod_solicitud DESC"
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

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "cod_solicitud LIKE '*" & txtbusqueda.Text & "*' OR " & " glosa LIKE '*" & txtbusqueda.Text & "*' OR " & " f2nomuser LIKE '*" & txtbusqueda.Text & "*' OR f3abrev LIKE '*" & txtbusqueda.Text & "*'  OR numorden LIKE '*" & txtbusqueda.Text & "*' "
    
    If Len(Trim(txtbusqueda.Text)) = 0 Then
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
    
    sql = "SELECT * FROM " & strTABLA & " WHERE ISNULL(SUC_1) ORDER BY ITEM"
    If RsM.State = 1 Then RsM.Close
    RsM.Open sql, cnn_Recibe_Mysql, 3, 1
    Do While Not RsM.EOF
        leyo = 1
        exito = True
        cnn_Recibe.Execute "insert into querys (item,wquery) values ('" & RsM.Fields("item") & "','" & RsM.Fields("wquery") & "')"
        If exito = True Then
            cnn_Recibe_Mysql.Execute "update " & strTABLA & " SET SUC_1=-1 where item = " & RsM.Fields("item")
        End If
        EstRecibir = True
    RsM.MoveNext
    Loop
    
'    If EstRecibir = True Then
        sql = "SELECT * FROM QUERYS WHERE  APLICADO= 0 ORDER BY ITEM"
        If rs.State = 1 Then rs.Close
        rs.Open sql, cnn_Recibe, 3, 1
        Do While Not rs.EOF
            jj = 0
            xcad = Replace(rs!WQUERY, "|", "'")
            cnn_dbbancos.Execute xcad
            If jj = 0 Then
                cnn_Recibe.Execute "UPDATE QUERYS SET APLICADO=-1 WHERE ITEM =" & rs!ITEM
            Else
                cnn_Recibe.Execute "UPDATE QUERYS SET ErrDiscrip = '" & JJE & "' WHERE ITEM =" & rs!ITEM
            End If
            rs.MoveNext
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

