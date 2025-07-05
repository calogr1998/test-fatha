VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form importar_ocompra 
   Caption         =   "Importar Ordenes de Compra"
   ClientHeight    =   7740
   ClientLeft      =   4335
   ClientTop       =   2520
   ClientWidth     =   10560
   Icon            =   "importar_ocompra.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda por Número de Orden"
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   10470
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   10110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presionar ENTER para realizar la búsqueda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   660
         Width           =   8475
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3045
      Left            =   60
      OleObjectBlob   =   "importar_ocompra.frx":000C
      TabIndex        =   1
      Top             =   1260
      Width           =   10425
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   635
      ButtonWidth     =   2223
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Configurar"
            Object.ToolTipText     =   "Configurar Servidor S10"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro     "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
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
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":4E08
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":53A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":593C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":5ED6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":6470
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":6A0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":6FA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":753E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":7AD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "importar_ocompra.frx":8072
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnu 
         Caption         =   "Actualizar ..."
         Index           =   0
      End
      Begin VB.Menu Mnu 
         Caption         =   "Actualizar Otra ORDEN DE COMPRA"
         Index           =   1
      End
   End
End
Attribute VB_Name = "importar_ocompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Af As New ADOFunctions
Dim wOrden As String
Dim cTipoProd As String
Dim cTipoPrv As String
Dim CnSql As New ADODB.Connection
Dim CnTmp As New ADODB.Connection
Dim Rs As New ADODB.Recordset
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
    Dim valor As String
    valor = ("" & (dxDBGrid1.Columns.ColumnByFieldName("f4numord").value))
    
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case "F3PREUNI", "F3CANPRO", "F3TOTAL", "SALDO"
    Text = Format(Text, "###,###,##0.00")
End Select

If Node.Values(dxDBGrid1.Columns.ColumnByFieldName("F4ESTADO").Index) >= "2" Then
    Color = &HFFC0C0
Else
    Color = &H8000000F
End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
If Val(dxDBGrid1.Columns.ColumnByFieldName("f4estado").value & "") >= 2 Then
    StrOrdenCompra = dxDBGrid1.Columns.ColumnByFieldName("f4numord").value & ""
    DatOrdenCompra = dxDBGrid1.Columns.ColumnByFieldName("f4fecemi").value & ""
    swActOrden = False
    Me.Hide
Else
    MsgBox "La orden de compra no está Aprobada.", vbExclamation, wnomcia
End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 13 Then
    dxDBGrid1_OnDblClick
End If
End Sub

Private Sub dxDBGrid1_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)
wCodConcar = right(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "", 5)
wdescosto = ObtenerCampo("centros", "f3abrev", "cconcar", wCodConcar, "T", cnn_dbbancos)
Mnu(0).Caption = "Actualizar Orden N° " & left(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "", LongOrdS10) & " [" & wdescosto & "]"
If Shift = 0 And Button = vbRightButton Then
 PopupMenu Menu
End If

End Sub

Private Sub Form_Load()
 
    'Me.AutoRedraw = False
    If CnTmp.State = 1 Then CnTmp.Close
    CnTmp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "templus.mdb;Persist Security Info=False" '"tmp_bancos.MDB;Persist Security Info=False"
    
    sw_nuevo_documento = True
    'Me.AutoRedraw = True
    FILL
    'dxDBGrid1.Dataset.ADODataset.Requery
    'proceso2 (dxDBGrid1.Columns.ColumnByFieldName("f4numord").Value)
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = vbHourglass
        
            Me.MousePointer = vbDefault
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

Private Sub FILL()
If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
cnn_dbbancos.Open StrConexDbBancos
With dxDBGrid1
'    'csql = "SELECT * FROM tmp_importaOC order by llave"
'            csql = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, ORDEN.F3CODFAB, "
'            csql = csql & "ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL,ORDEN.F4NUMORD + format(ORDEN.item,'000') as LLave  "
'            csql = csql & "FROM "
'            csql = csql & "(SELECT Right(IF4ORDEN.F4NUMORD,5) AS CODCENTRO, "
'            csql = csql & "'[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']' AS Grupo, "
'            csql = csql & "EF2PROVEEDORES.F2NEWRUC,IF3ORDEN.item , IF4ORDEN.F4NUMORD,IF4ORDEN.F4local, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, "
'            csql = csql & "IF3ORDEN.F5NOMPRO, IF3ORDEN.F3PREUNI, IF3ORDEN.F3CANPRO, IF3ORDEN.F3TOTAL "
'            csql = csql & "FROM (IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
'            csql = csql & "INNER JOIN IF3ORDEN ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
'            csql = csql & "Where (((EF2PROVEEDORES.F2NEWRUC) = '" & wRucCliProv & "') ) "
'            'csql = csql & "and IF4ORDEN.f4numord not in (select f4ocompra from regisdoc where len(trim(f4ocompra))>0) "
'            csql = csql & "ORDER BY Right(IF4ORDEN.F4NUMORD,5), Left(IF4ORDEN.F4NUMORD," & LongOrdS10 & ")) as "
'            csql = csql & "ORDEN LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.CCONCAR "
'            csql = csql & "ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo"
    
'    csql = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, "
'    csql = csql & "ORDEN.F3CODFAB, ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL,orden.f4estado, "
'    csql = csql & "ORDEN.F4NUMORD+Format(ORDEN.item,'000') AS LLave, "
'    'csql = csql & "IIF(ORDEN.F3PORDCT>0,orden.F3TOTAL*ORDEN.F3PORDCT/100, orden.F3TOTAL) "
'    csql = csql & "orden.F3TOTAL "
'    csql = csql & "- iif(REGISMOV.F3AFECTO='*',IIf(IsNull(articulos.f3importe),0,articulos.f3importe*1." & wigv & "),IIf(IsNull(articulos.f3importe),0,articulos.f3importe)) AS Saldo "
'    csql = csql & "FROM "
'    csql = csql & "(("
'    csql = csql & "SELECT IF3ORDEN.F3CENCOS AS CODCENTRO, "
'    csql = csql & "'[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']' AS Grupo, "
'    csql = csql & "EF2PROVEEDORES.F2NEWRUC, IF3ORDEN.ITEM, IF4ORDEN.F4NUMORD, IF4ORDEN.F4LOCAL, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, IF3ORDEN.F5NOMPRO, "
'    csql = csql & "IF3ORDEN.F3PREUNI, IF3ORDEN.F3CANPRO,IF3ORDEN.F3PORDCT, IF3ORDEN.F3TOTAL, IF3ORDEN.F5VALVTA,IF4ORDEN.f4estado "
'    csql = csql & "FROM (IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) "
'    csql = csql & "INNER JOIN IF3ORDEN ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
'    csql = csql & "WHERE (((EF2PROVEEDORES.F2NEWRUC)='" & wRucCliProv & "')) "
'    csql = csql & "ORDER BY IF4ORDEN.F4CENTRO, IF4ORDEN.F4NUMORD"
'    csql = csql & ") AS ORDEN "
'    csql = csql & "LEFT JOIN ("
'    csql = csql & "SELECT REGISMOV.F3ORDEN AS F4OCOMPRA, REGISMOV.F5CODPRO, sum(REGISMOV.F3IMPORTE) as F3IMPORTE, REGISMOV.F3AFECTO "
'    csql = csql & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV)"
'    csql = csql & "GROUP BY REGISMOV.F3ORDEN, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
'    csql = csql & "ORDER BY REGISMOV.F3ORDEN, REGISMOV.F5CODPRO"
'    csql = csql & ") AS Articulos ON (ORDEN.F4NUMORD = Articulos.F4OCOMPRA) AND (ORDEN.F3CODFAB = Articulos.F5CODPRO)) "
'    csql = csql & "LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.F3COSTO "
'    csql = csql & "Where ((("
'    'csql = csql & "orden.f5valvta - IIf(IsNull(articulos.f3importe), 0, articulos.f3importe)"
'    csql = csql & "VAL(FORMAT(orden.F3TOTAL,'0.00')) "
'    csql = csql & "- VAL(FORMAT(iif(REGISMOV.F3AFECTO='*',IIf(IsNull(articulos.f3importe),0,articulos.f3importe*1." & wigv & "),IIf(IsNull(articulos.f3importe),0,articulos.f3importe)),'0.00')) "
'    csql = csql & ") > 0)) "
'    csql = csql & "ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo,val(ORDEN.item&'')"
    
    SqlCad = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, ORDEN.F3CODFAB, ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL, ORDEN.f4estado, ORDEN.F4NUMORD+Format(ORDEN.item,'000') AS LLave, orden.F3TOTAL-IIf(REGISMOV.F3AFECTO='*',IIf(IsNull(articulos.f3importe),0,articulos.f3importe*1.18),IIf(IsNull(articulos.f3importe),0,articulos.f3importe)) AS Saldo, Val(Format([orden].[F3TOTAL],'0.00'))-Val(Format(IIf([REGISMOV].[F3AFECTO]='*',IIf(IsNull([articulos].[f3importe]),0,[articulos].[f3importe]*1.18),IIf(IsNull([articulos].[f3importe]),0,[articulos].[f3importe])),'0.00')) "
    SqlCad = SqlCad & "FROM ((SELECT IF3ORDEN.F3CENCOS AS CODCENTRO, '[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']' AS Grupo, EF2PROVEEDORES.F2NEWRUC, IF4ORDEN.F4NUMORD, IF4ORDEN.F4LOCAL, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, First(IF3ORDEN.ITEM) AS ITEM, IF3ORDEN.F5NOMPRO, Avg(IF3ORDEN.F3PREUNI) AS F3PREUNI, Sum(IF3ORDEN.F3CANPRO) AS F3CANPRO, Avg(IF3ORDEN.F3PORDCT) AS F3PORDCT, Sum(IF3ORDEN.F3TOTAL) AS F3TOTAL, Sum(IF3ORDEN.F5VALVTA) AS F5VALVTA, IF4ORDEN.F4ESTADO "
    SqlCad = SqlCad & "FROM (IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC) INNER JOIN IF3ORDEN ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
    SqlCad = SqlCad & "GROUP BY IF3ORDEN.F3CENCOS, '[N° Orden: '+Left(IF4ORDEN.F4NUMORD,12)+']; [Proveedor: '+EF2PROVEEDORES.F2NOMPROV+']; [Total: '+Format(IF4ORDEN.F4MONTO,'#,##0.00')+']', EF2PROVEEDORES.F2NEWRUC, IF4ORDEN.F4NUMORD, IF4ORDEN.F4LOCAL, IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODFAB, IF3ORDEN.F5NOMPRO, IF4ORDEN.F4ESTADO, IF4ORDEN.F4NUMORD "
    SqlCad = SqlCad & "HAVING (((EF2PROVEEDORES.F2NEWRUC)='" & wRucCliProv & "')) "
    SqlCad = SqlCad & "ORDER BY First(IF4ORDEN.F4CENTRO), IF4ORDEN.F4NUMORD "
    SqlCad = SqlCad & ") AS ORDEN LEFT JOIN (SELECT REGISMOV.F3ORDEN AS F4OCOMPRA, REGISMOV.F5CODPRO, sum(REGISMOV.F3IMPORTE) AS F3IMPORTE, REGISMOV.F3AFECTO "
    SqlCad = SqlCad & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV=REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV=REGISMOV.F4MESMOV) "
    SqlCad = SqlCad & "GROUP BY REGISMOV.F3ORDEN, REGISMOV.F5CODPRO, REGISMOV.F3AFECTO "
    SqlCad = SqlCad & "ORDER BY REGISMOV.F3ORDEN, REGISMOV.F5CODPRO) AS Articulos ON (ORDEN.F4NUMORD = Articulos.F4OCOMPRA) AND (ORDEN.F3CODFAB = Articulos.F5CODPRO)) LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.F3COSTO "
    SqlCad = SqlCad & "WHERE (((Val(Format([orden].[F3TOTAL],'0.00'))-Val(Format(IIf([REGISMOV].[F3AFECTO]='*',IIf(IsNull([articulos].[f3importe]),0,[articulos].[f3importe]*1." & wIgv & "),IIf(IsNull([articulos].[f3importe]),0,[articulos].[f3importe])),'0.00')))>0)) "
    SqlCad = SqlCad & "ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo, Val(ORDEN.item & '')"

    
    
    'csql = "SELECT CENTROS.F3DESCRIP, ORDEN.Grupo, ORDEN.F2NEWRUC, ORDEN.F4NUMORD, ORDEN.F4FECEMI, ORDEN.F3CODFAB, ORDEN.F5NOMPRO, ORDEN.F3PREUNI, ORDEN.F3CANPRO, ORDEN.F3TOTAL, ORDEN.F4NUMORD+Format(ORDEN.item,'000') AS LLave, orden.F3TOTAL- iif(REGISMOV.F3AFECTO='*',IIf(IsNull(articulos.f3importe),0,articulos.f3importe*1.19),IIf(IsNull(articulos.f3importe),0,articulos.f3importe)) AS Saldo FROM (ORDEN LEFT JOIN  Articulos ON (ORDEN.F4NUMORD = Articulos.F4OCOMPRA) AND (ORDEN.F3CODFAB = Articulos.F5CODPRO)) LEFT JOIN CENTROS ON ORDEN.CODCENTRO = CENTROS.F3COSTO Where ((([orden].[f5valvta] - IIf(IsNull([articulos].[f3importe]), 0, [articulos].[f3importe])) > 0)) ORDER BY CENTROS.F3DESCRIP, ORDEN.Grupo"
    
        .Dataset.ADODataset.ConnectionString = StrConexDbBancos  ' CnTmp
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .KeyField = "llave"
        .m.FullExpand
End With
        
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Unload Me
'swActOrden = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
fraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtBusqueda.Width = fraBusqueda.Width - 400
dxDBGrid1.Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)

End Sub

Private Sub Mnu_Click(Index As Integer)
Dim CadSql As String
Select Case Index
Case 0
'On Error GoTo ErrConex
MousePointer = 11
'carga parámetros del S10
    csql = "select servername,basedatos,username,password from servers10"
    If Rs.State = 1 Then Rs.Close
    Rs.Open csql, cnn_dbbancos, 3, 1
    If Rs.RecordCount > 0 Then
        CadSql = "Provider=SQLOLEDB.1;"
        CadSql = CadSql & "Password=" & Rs!Password
        CadSql = CadSql & ";Persist Security Info=True;"
        CadSql = CadSql & "User ID=" & Rs!UserName
        CadSql = CadSql & ";"
        CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
        CadSql = CadSql & ";"
        CadSql = CadSql & "Data Source=" & Rs!ServerName
    Else
        ConfigurarServerS10.Show 1
        Rs.Requery
        If Rs.RecordCount > 0 Then
            CadSql = "Provider=SQLOLEDB.1;"
            CadSql = CadSql & "Password=" & Rs!Password
            CadSql = CadSql & ";Persist Security Info=True;"
            CadSql = CadSql & "User ID=" & Rs!UserName
            CadSql = CadSql & ";"
            CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
            CadSql = CadSql & ";"
            CadSql = CadSql & "Data Source=" & Rs!ServerName
        Else
            Exit Sub
        End If
    End If
    If CnSql.State = 1 Then CnSql.Close
    CnSql.Open CadSql
wCodConcar = right(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "", 5)
CargaOrden left(dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").value & "", LongOrdS10), wCodConcar, ""

FILL
MousePointer = 0
Case 1

'On Error GoTo ErrConex
MousePointer = 11
'carga parámetros del S10
    csql = "select servername,basedatos,username,password from servers10"
    If Rs.State = 1 Then Rs.Close
    Rs.Open csql, cnn_dbbancos, 3, 1
    If Rs.RecordCount > 0 Then
        CadSql = "Provider=SQLOLEDB.1;"
        CadSql = CadSql & "Password=" & Rs!Password
        CadSql = CadSql & ";Persist Security Info=True;"
        CadSql = CadSql & "User ID=" & Rs!UserName
        CadSql = CadSql & ";"
        CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
        CadSql = CadSql & ";"
        CadSql = CadSql & "Data Source=" & Rs!ServerName
    Else
        ConfigurarServerS10.Show 1
        Rs.Requery
        If Rs.RecordCount > 0 Then
            CadSql = "Provider=SQLOLEDB.1;"
            CadSql = CadSql & "Password=" & Rs!Password
            CadSql = CadSql & ";Persist Security Info=True;"
            CadSql = CadSql & "User ID=" & Rs!UserName
            CadSql = CadSql & ";"
            CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
            CadSql = CadSql & ";"
            CadSql = CadSql & "Data Source=" & Rs!ServerName
        Else
            Exit Sub
        End If
    End If
    If CnSql.State = 1 Then CnSql.Close
    CnSql.Open CadSql
'    Dim I As Integer
'    Set rs = Af.OpenSQLForwardOnly("select * from proyecto", CadSql)
'    Do While Not rs.EOF
'        For I = 0 To (rs.Fields.Count - 1)
'            xcad = xcad & rs.Fields(I) & vbTab
'        Next
'        xcad = xcad & vbCrLf
'        rs.MoveNext
'    Loop
'    MsgBox xcad

wOrden = InputBox("Ingrese el Número de Orden", "Importa Orden de Compra S10")
If IsNumeric(wOrden) Then wOrden = Format(wOrden, Repetir(LongOrdS10, "0"))
If Len(Trim(wOrden)) > 0 Then
    wcodcosto = ""
    'Ayuda_CENTROS.SelectInto = "'999','998'"
    Ayuda_Centros.Show 1
    wCodConcar = ObtenerCampo("centros", "cconcar", "f3costo", wcodcosto, "T", cnn_dbbancos)
    If Len(Trim(wCodConcar)) > 0 Then
        
        CargaOrden wOrden, wCodConcar, ""

        FILL
    End If
End If
MousePointer = 0

End Select

Exit Sub
ErrConex:
    'MsgBox Err.Description
    MsgBox Err.Description, vbCritical, wnomcia
    
    Resume Next
    Exit Sub
End Sub

Private Sub BuscarOrdenS10(pNumeroDeOrden As String, pRucProveedor As String)
On Error GoTo ErrConex
MousePointer = 11
'carga parámetros del S10
    csql = "select servername,basedatos,username,password from servers10"
    If Rs.State = 1 Then Rs.Close
    Rs.Open csql, cnn_dbbancos, 3, 1
    If Rs.RecordCount > 0 Then
        CadSql = "Provider=SQLOLEDB.1;"
        CadSql = CadSql & "Password=" & Rs!Password
        CadSql = CadSql & ";Persist Security Info=True;"
        CadSql = CadSql & "User ID=" & Rs!UserName
        CadSql = CadSql & ";"
        CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
        CadSql = CadSql & ";"
        CadSql = CadSql & "Data Source=" & Rs!ServerName
    Else
        ConfigurarServerS10.Show 1
        Rs.Requery
        If Rs.RecordCount > 0 Then
            CadSql = "Provider=SQLOLEDB.1;"
            CadSql = CadSql & "Password=" & Rs!Password
            CadSql = CadSql & ";Persist Security Info=True;"
            CadSql = CadSql & "User ID=" & Rs!UserName
            CadSql = CadSql & ";"
            CadSql = CadSql & "Initial Catalog=" & Rs!basedatos
            CadSql = CadSql & ";"
            CadSql = CadSql & "Data Source=" & Rs!ServerName
        Else
            Exit Sub
        End If
    End If
    If CnSql.State = 1 Then CnSql.Close
    CnSql.Open CadSql
    

'wOrden = InputBox("Ingrese el Número de Orden", "Importa Orden de Compra S10")
If IsNumeric(pNumeroDeOrden) Then
    pNumeroDeOrden = Format(pNumeroDeOrden, Repetir(LongOrdS10, "0"))
    If Len(Trim(pNumeroDeOrden)) > 0 Then
        'wCodCosto = ""
        'Ayuda_CENTROS.Show 1
        wCodConcar = "" ' ObtenerCampo("centros", "cconcar", "f3costo", wCodCosto, "T", cnn_DbBancos)
        If Len(Trim(wRucCliProv)) > 0 Then
            
            CargaOrden pNumeroDeOrden, wCodConcar, wRucCliProv
    
            FILL
        End If
    End If
Else
    MsgBox "El número de orden no es válido", vbExclamation, wnomcia
End If
MousePointer = 0
Exit Sub
ErrConex:
    'MsgBox Err.Description
    MsgBox Err.Description, vbCritical, wnomcia
    
    Resume Next
    Exit Sub
End Sub


Private Sub CargaOrden(cNumeroOrden As String, cCentroCosto As String, cRucProveedor As String)

Dim UltFecha As Date, SwFecha As Boolean
   
    csql = "select oc.nroorden,oc.fecha,oc.codProveedor,it.descripcion,it.ruc,oc.observacion,oc.* "
    csql = csql & "from OrdenDeCompra oc,identificador it "
    csql = csql & "where oc.codproveedor=it.codidentificador "
    csql = csql & "and oc.CODORDEN='" & cNumeroOrden & "' "
    If Len(Trim(cCentroCosto)) > 0 Then
        csql = csql & "and OC.CODPROYECTO ='" & cCentroCosto & "'"
    ElseIf Len(Trim(cRucProveedor)) > 0 Then
        csql = csql & "and it.ruc ='" & cRucProveedor & "'"
    End If
    If Rs.State = 1 Then Rs.Close
    Rs.Open csql, CnSql, 3, 1
    i = 0
    If Rs.RecordCount > 0 Then
        
        Do While Not Rs.EOF
        
            Call GrabarOC
            Me.Refresh
            Rs.MoveNext
        Loop
        
        'MsgBox "Actualización Finalizada", vbInformation, wNomCia
        'Unload Me
    Else
        MsgBox "No Hay Registros para Actualizar", vbExclamation, wnomcia
        'Unload Me
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Trim(Button.Caption))
Case "CONFIGURAR"
    csql = "select servername,basedatos,username,password from servers10"
    
    Set Rs = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
    If Rs.RecordCount > 0 Then
        With ConfigurarServerS10
            .CmbDatabases.Text = "" & Rs!basedatos
            .cmbServers.Text = "" & Rs!ServerName
            .TxtUse.Text = "" & Rs!UserName
            .TxtPas.Text = "" & Rs!Password
        End With
    Else
        With ConfigurarServerS10
            .CmbDatabases.Text = ""
            .cmbServers.Text = ""
            .TxtUse.Text = ""
            .TxtPas.Text = ""
        End With
    End If
    ConfigurarServerS10.Show 1
Case "SALIR"
If cnn_form.State = 1 Then cnn_form.Close
Set cnn_form = Nothing
    swActOrden = False
    Me.Hide
End Select
End Sub

Private Sub txtbusqueda_Change()
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "f4numord LIKE '*" & txtBusqueda.Text & "*'"
        
        If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
        End If
        dxDBGrid1.m.FullExpand
        
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 2
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "f4numord LIKE '*" & txtBusqueda.Text & "*'"
        
        If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
        End If
        dxDBGrid1.m.FullExpand
        
        'If dxDBGrid1.Dataset.RecordCount = 0 Then
            DoEvents
            Call BuscarOrdenS10(txtBusqueda.Text, wRucCliProv)
        'End If
        Call txtbusqueda_KeyDown(40, 0)
    End If
End Sub

Private Sub GrabarOC()
Dim codi                As String
Dim wcantidad           As Double
Dim wcc                 As String
Dim wproducto           As String
Dim SqlCad                 As String
Dim ocompra             As Double
Dim Cant                As Double
Dim rsdetaoc            As New ADODB.Recordset
Dim ncant_ant           As Double
Dim amovs_cab(0 To 60)  As a_grabacion
Dim ctipo               As String
Dim wNumOc              As String

Dim rsOrdenCab As New ADODB.Recordset
Dim RsOC As New ADODB.Recordset
Dim rsOrdenDet As New ADODB.Recordset
Dim coddetalle  As String
Dim J As Integer
    loc = 1
    'flag = 0
    'If Trim(Rs!nroorden & "") <> "" Then
    '    jc = 1
    'Else
    '    jc = 0
    'End If
    wNumOc = Rs!CODORDEN & "/" & Rs!CODPROYECTO
    
    'If Txt_Prove = "" Then MsgBox "Ingrese Código de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    'If PnlNomPrv = "" Then MsgBox "Ingrese Nombre de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    'If txtcodsoli = "" Then MsgBox "Ingrese Código de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    'If pnlnomsoli = "" Then MsgBox "Ingrese Nombre de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    'If txtcodforma = "" Then MsgBox "Ingrese código de forma de pago", 48, "Sistema de Logística": txtcodforma.SetFocus: Exit Sub
    'If Cmbmone.ListIndex < 0 Then MsgBox "Seleccione moneda", 48, "Sistema de Logística": Cmbmone.SetFocus: Exit Sub
    'If Val(txt_tc.Text) = 0 Then MsgBox "Ingrese Tipo de Cambio", 48, "Sistema de Logística": txt_tc.SetFocus: Exit Sub
    
    'Nueva Versión
  '  If loc = 1 Then
  '      Select Case jc
  '          Case 0
  '              'Call Nueva_orden
  '      End Select
  '
  '  End If
        
    
    If loc = 1 Then
        Set rsOrdenCab = New ADODB.Recordset
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        'If SwRenovar = True Then
        '    wNumOc = Mid(Txt_NumOC.Text, 1, 11) & Val(Mid(Txt_NumOC.Text, 12, 1)) + 1
        '    rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL from if4orden where f4numord='" & wNumOc & "' AND F4LOCAL = '1'", Cnn_DbBancos, adOpenDynamic, adLockOptimistic
        'Else
            rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL from if4orden where f4numord='" & wNumOc & "' AND F4LOCAL = '" & TOC & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsOC.Filter = ""
            RsOC.Filter = "f4numord='" & wNumOc & "' AND F4LOCAL = '" & TOC & "'"
            'If RsOC.RecordCount > 0 Then
           '     If IsDate(RsOC!f4fecmod) Then
           '         If Format(CVDate(Rs!actualizacionfecha), "DD/MM/YYYY HH:MM:SS") <= Format(CVDate(RsOC!f4fecmod), "DD/MM/YYYY HH:MM:SS") Then
          '          '   Exit Sub
           '         End If
           '     End If
          '  End If
        'End If
        If Not (rsOrdenCab.EOF) Then
            ctipo = "M"
        Else
            ctipo = "A"
            'flag = 1
            sw_nuevo_documento = False
        End If
        'If SwRenovar = True Then
        '    amovs_cab(0).Campo = "F4NUMORD": amovs_cab(0).valor = wNumOc: amovs_cab(0).TIPO = "T"
        'Else
            amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = wNumOc: amovs_cab(0).TIPO = "T"
        'End If
        
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = "N": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = "1": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = 0: amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECGRA": amovs_cab(4).valor = Format(Rs!creacionfecha & "", "dd/MM/yyyy HH:MM:SS"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEGRA": amovs_cab(5).valor = Rs!CREACIONUSUARIO & "": amovs_cab(5).TIPO = "T"
            amovs_cab(28).campo = "F4FECMOD": amovs_cab(28).valor = Format(Rs!actualizacionfecha & "", "dd/MM/yyyy HH:MM:SS"): amovs_cab(28).TIPO = "F"
            amovs_cab(29).campo = "F4USEMOD": amovs_cab(29).valor = Rs!actualizacionUsuario & "": amovs_cab(29).TIPO = "T"
        
        
        amovs_cab(6).campo = "F4CODSOL": amovs_cab(6).valor = "": amovs_cab(6).TIPO = "T"
        amovs_cab(7).campo = "F4FECEMI": amovs_cab(7).valor = Format(Rs!Fecha & "", "DD/MM/YYYY"): amovs_cab(7).TIPO = "F"
        'valida proveedor
        wRucCliProv = Trim(ObtenerCampo("identificador", "RUC", "CodIdentificador", Rs!CodProveedor & "", "T", CnSql))
        wcodcliprov = Trim(ObtenerCampo("ef2proveedores", "f2codprov", "f2newruc", left(wRucCliProv & "", 11), "T", cnn_dbbancos))
        If Len(Trim(wcodcliprov)) = 0 Then
            cTipoPrv = "A"
        Else
            cTipoPrv = "M"
        End If
        ValidaProveedor (Rs!CodProveedor)
        '****
        If (Len(Trim(wRucCliProv))) = 0 Then MsgBox ("ruc vacio")
        amovs_cab(8).campo = "F4CODPRV": amovs_cab(8).valor = wRucCliProv: amovs_cab(8).TIPO = "T"
        amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = Rs!TipoDeCambio & "": amovs_cab(9).TIPO = "N"
        wforpag = Trim(ObtenerCampo("EF2FORPAG", "F2FORPAG", "F2CODS10", Rs!CODFORMADEPAGO & "", "T", cnn_dbbancos))
        amovs_cab(10).campo = "F4FORPAG": amovs_cab(10).valor = wforpag: amovs_cab(10).TIPO = "T"
        amovs_cab(11).campo = "F4REFERE": amovs_cab(11).valor = "": amovs_cab(11).TIPO = "T"
        amovs_cab(12).campo = "F4OBSERVA": amovs_cab(12).valor = Replace(Replace(Rs!Observacion, Chr(13), " "), "'", "´"): amovs_cab(12).TIPO = "T"
        wcodcosto = Trim(ObtenerCampo("CENTROS", "F3COSTO", "CCONCAR", Rs!CODPROYECTO & "", "T", cnn_dbbancos))
        amovs_cab(13).campo = "F4CENTRO": amovs_cab(13).valor = wcodcosto & "": amovs_cab(13).TIPO = "T"
        amovs_cab(14).campo = "F4CODSOLICITUD": amovs_cab(14).valor = "": amovs_cab(14).TIPO = "T"
        amovs_cab(15).campo = "F4TIPMON": amovs_cab(15).valor = IIf(Rs!CodMoneda & "" = "02", "D", "S"): amovs_cab(15).TIPO = "T"
        amovs_cab(16).campo = "F4IGV": amovs_cab(16).valor = Val(Format(Rs!VALORIGV & "", "0.00")): amovs_cab(16).TIPO = "N"
        amovs_cab(17).campo = "F4MONINA": amovs_cab(17).valor = 0: amovs_cab(17).TIPO = "N"
        amovs_cab(18).campo = "F4BASIMP": amovs_cab(18).valor = Val(Format(Rs!VALORNETO & "", "0.00")): amovs_cab(18).TIPO = "N"
        amovs_cab(19).campo = "F4MONTO": amovs_cab(19).valor = Val(Format(Rs!VALORTOTAL, "0.00")): amovs_cab(19).TIPO = "N"
        amovs_cab(20).campo = "F4LOCAL": amovs_cab(20).valor = "1": amovs_cab(20).TIPO = "T"
        amovs_cab(21).campo = "F4EMPRESA": amovs_cab(21).valor = wnomcia: amovs_cab(21).TIPO = "T"
        amovs_cab(22).campo = "F4UUPP": amovs_cab(22).valor = "": amovs_cab(22).TIPO = "T"
        amovs_cab(23).campo = "F4PLAZO_ENTREGA": amovs_cab(23).valor = "": amovs_cab(23).TIPO = "T"
        amovs_cab(24).campo = "F4LUGAR_ENTREGA": amovs_cab(24).valor = (Rs!LUGARDEENTREGA & ""): amovs_cab(24).TIPO = "T"
        amovs_cab(25).campo = "F4CONTACTO": amovs_cab(25).valor = Rs!CODCONTACTO & "": amovs_cab(25).TIPO = "T"
        amovs_cab(26).campo = "F4FECENT": amovs_cab(26).valor = Format(Rs!FechaEntrega & "", "DD/MM/YYYY HH:MM:SS"): amovs_cab(26).TIPO = "F"
        amovs_cab(27).campo = "F4RND": amovs_cab(27).valor = Val(0): amovs_cab(27).TIPO = "N"
        amovs_cab(30).campo = "F4numcotiza": amovs_cab(30).valor = Rs!nrofacturaboleta & "": amovs_cab(30).TIPO = "T"
        'aprobaciones automaticas si viene del s10
        amovs_cab(31).campo = "F4VB1": amovs_cab(31).valor = -1: amovs_cab(31).TIPO = "N"
        amovs_cab(32).campo = "F4VBUSER1": amovs_cab(32).valor = "S10": amovs_cab(32).TIPO = "T"
        amovs_cab(33).campo = "F4VBFECHA1": amovs_cab(33).valor = Format(Rs!creacionfecha & "", "dd/MM/yyyy HH:MM:SS"): amovs_cab(33).TIPO = "T"
        amovs_cab(34).campo = "F4VB2": amovs_cab(34).valor = -1: amovs_cab(34).TIPO = "N"
        amovs_cab(35).campo = "F4VBUSER2": amovs_cab(35).valor = "S10": amovs_cab(35).TIPO = "T"
        amovs_cab(36).campo = "F4VBFECHA2": amovs_cab(36).valor = Format(Rs!creacionfecha & "", "dd/MM/yyyy HH:MM:SS"): amovs_cab(36).TIPO = "T"
        amovs_cab(37).campo = "F4ESTADO": amovs_cab(37).valor = 2: amovs_cab(37).TIPO = "N"
        
        rsOrdenCab.Close
        
        If ctipo = "A" Then     '--- Nuevo
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 37, StrConexDbBancos, ""
        Else
            GRABA_REGISTRO amovs_cab(), "IF4ORDEN", ctipo, 37, StrConexDbBancos, "F4NUMORD = '" & wNumOc & _
            "' AND F4LOCAL = '" & TOC & "'"
        End If
        
    End If
    
    '---------- GRABANDO EL DETALLE DE LA ORDEN DE COMPRA ----------------------'
'    If cTipoAdm_Bd = "M" Then
'        If SwRenovar = True Then
'            Cnn_DbBancos.Execute ("delete from if3orden where f4numord= '" & wNumOc & "'  AND F4LOCAL = '1'")
'        Else
'            Cnn_DbBancos.Execute ("delete from if3orden where f4numord= '" & Txt_NumOC.Text & "'  AND F4LOCAL = '1'")
'        End If
'    Else
'        If SwRenovar = True Then
'            Cnn_DbBancos.Execute ("delete * from if3orden where f4numord= '" & wNumOc & "'  AND F4LOCAL = '1'")
'        Else
            'If wnumoc = "000000000335/02003" Then MsgBox "obs"
            cnn_dbbancos.Execute ("delete * from if3orden where f4numord= '" & wNumOc & "'  AND F4LOCAL = '" & TOC & "'")
 '       End If
'    End If
    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
    'rsOrdenDet.Open "select * from if3orden", Cnn_DbBancos, adOpenDynamic, adLockOptimistic
    coddetalle = ObtenerCampoWhere("ordendecompra", "nroorden", "codorden", left(wNumOc, LongOrdS10), "T", CnSql, " and codproyecto='" & right(wNumOc, 5) & "'")
    If Len(Trim(coddetalle)) = 0 Then Exit Sub
    csql = "select * from OrdenDeCompraDetalle where nroorden='" & coddetalle & "'"
    If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    rsdetaoc.Open csql, CnSql, 3, 1
    J = 1
    If rsdetaoc.RecordCount > 0 Then
        rsdetaoc.MoveFirst
        With rsdetaoc
            Do While Not .EOF
            
                Rem NSE If (Len(Trim(.Fields("f3codpro")))) = 0 Or (Val(Format(.Fields("f3canpro") & "", "0.00")) = 0) Or (Val(Format(.Fields("f3precos") & "", "0.000")) = 0) Then
                Rem NSE     wgrabar = False
                Rem NSE Else
                Rem NSE     wgrabar = True
                Rem NSE End If
                
'                If wgrabar Then
                    WNOMPROD = ""
                    codi = .Fields("codinsumo")
                    WNOMPROD = ObtenerCampo("INSUMO", "DESCRIPCION", "CODINSUMO", codi, "T", CnSql)
                    WNOMPROD = Replace(WNOMPROD, "'", "´")
                    WCODPROD = ObtenerCampo("if5pla", "f5codpro", "f5codpro", codi, "T", cnn_dbbancos)
                    cTipoProd = "A"
                    If Len(Trim(WCODPROD)) > 0 Then
                        cTipoProd = "M"
                    End If
                    'Actualiza Centro de Productos
                    ValidaProductos (codi)
                    wcantidad = .Fields("cantidad")
                    wcc = wcodcosto
                    wproducto = Trim$(codi)
                    
                    SqlCad = "select f3presu,f3consumido,f3ocompra from centroproductos where " _
                    & "f3costo='" & wcc & "' and f5codpro='" & wproducto & "'"
                    Set rstaux = New ADODB.Recordset
                    If rstaux.State = adStateOpen Then rst.Close
                    rstaux.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not (rstaux.EOF) Then
                        If jc = 0 Then  'Nuevo
                            ocompra = Val(rstaux.Fields("f3ocompra").value)
                            rstaux.Fields("f3ocompra").value = ocompra + wcantidad
                        Else             'Modifica
                            rstaux.Fields("f3ocompra").value = wcantidad
                        End If
                        rstaux.Update
                    End If
                    rstaux.Close
                    
                    If SwRenovar = True Then
                        SqlCad = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD" _
                        & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL" _
                        & ",F3FENTREGA,F5NOMPRO, F4LOCAL, f3cencos,item,f3observa) VALUES " _
                        & "('" & wNumOc & "','" & .Fields("f3codpro") & "','" & .Fields("f5codfab") & "'," _
                        & .Fields("f3canpro") & ",'" & .Fields("f5marca") & "','" & .Fields("f3medida") & "'," _
                        & .Fields("f3canpro") & "," & .Fields("f3preuni") & "," & .Fields("f3precos") & "," _
                        & IIf(IsNull(.Fields("f3pordct")), "0", .Fields("f3pordct")) & "," & .Fields("f3totdct") & "," _
                        & .Fields("f5valvta") & ",'" & IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto")) & "'," _
                        & .Fields("f3igv") & "," & .Fields("f3total") & ",'" & .Fields("f3fentrega") & "','" _
                        & .Fields("f5nompro") & "', '1','" & wcodcosto & "'," & J & ",'" & left(.Fields("Observacion") & "", 255) & "')"
                    Else
                        nprecio = Val(.Fields("PRECIO") & "")
                        swafecto = IIf(.Fields("igv") = 0, " ", "*")
                        If swafecto = "*" Then
                            'nprecioconigv = nprecio * 1.19
                            nprecioconigv = nprecio * (1 + wIgv / 100)
                            npreciosinigv = nprecio
                        Else
                            nprecioconigv = nprecio
                            npreciosinigv = nprecio
                        End If
                        
                
                        obs = verificaSintaxsi(left(.Fields("Observacion") & "", 255)) ''modifico jaime 07/10/10

                        WNOMPROD = verificaSintaxsi(WNOMPROD)
               
                    
                        SqlCad = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD" _
                        & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL" _
                        & ",F3FENTREGA,F5NOMPRO, F4LOCAL,f3cencos,item,f3observa) VALUES " _
                        & "('" & wNumOc & "','" & .Fields("CODINSUMO") & "','" & .Fields("CODINSUMO") & "'," _
                        & .Fields("CANTIDAD") & ",'" & "" & "','" & ObtenerCampo("INSUMO", "codunidad", "CODINSUMO", codi, "T", CnSql) & "'," _
                        & .Fields("CANTIDAD") & "," & Val(nprecioconigv & "") & "," & Val(nprecio & "") & "," _
                        & IIf(IsNull(.Fields("DESCUENTO")), "0", .Fields("DESCUENTO")) & "," & .Fields("DESCUENTOTOTAL") & "," _
                        & .Fields("VALORPARCIAL") & ",'" & swafecto & "'," _
                        & .Fields("VALORigv") & "," & Val(.Fields("ParcialConIgv") & "") & ",'" & Rs!FechaEntrega & "" & "','" _
                        & WNOMPROD & "', '1','" & wcodcosto & "'," & .Fields("Linea") & ",'" & obs & "')"
                    End If
                        
                    cnn_dbbancos.Execute SqlCad

'                    rsOrdenDet.AddNew
'                    rsOrdenDet!F4NUMORD = Val(Txt_NumOC.Text)
'                    rsOrdenDet!F3CODPRO = .Fields("f3codpro")
'
'                    rsOrdenDet!F3CODFAB = .Fields("f5codfab")
'                    rsOrdenDet!F3CANPRO = .Fields("f3canpro")
'                    rsOrdenDet!F5MARCA = .Fields("f5marca")
'                    rsOrdenDet!unidad = .Fields("f3medida")
'                    rsOrdenDet!f3canfal = .Fields("f3canpro")
'                    rsOrdenDet!F3PREUNI = .Fields("f3preuni")
'                    rsOrdenDet!f3PRECOS = .Fields("f3precos")
'                    rsOrdenDet!F3PORDCT = .Fields("f3pordct")
'                    rsOrdenDet!f3totdct = .Fields("f3totdct")
'                    rsOrdenDet!f5valvta = .Fields("f5valvta")
'                    rsOrdenDet!F5AFECTO = IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto"))
'                    rsOrdenDet!F3IGV = .Fields("f3igv")
'                    rsOrdenDet!F3TOTAL = .Fields("f3total")
'                    rsOrdenDet!f3fentrega = Format$(.Fields("f3fentrega"), "dd/mm/yyyy")
'                    rsOrdenDet!f5nompro = Trim(.Fields("F5NOMPRO") & "")
'
'
'                    rsOrdenDet.Update
'                End If
                J = J + 1
                .MoveNext
            Loop
            'rsOrdenDet.Close
        End With
    End If
    rsdetaoc.Close
    
'    Call VERIFIC_PPRV
    
    If Txt_NumSolComp <> "" Then
        SqlCad = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & _
        Txt_NumSolComp & "'"
        
        cnn_dbbancos.Execute SqlCad
    
        If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
        rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
        If Not rsdetaoc.EOF Then
            With rsdetaoc
                .MoveFirst
                Do While Not .EOF
                    codprod = .Fields("f3codpro") & ""
                    Rem NSE If Val("" & .Fields("f3precos")) > 0 Then
                    If .Fields("check") = True Then
                        Cant = Val("" & .Fields("f3canpro"))
                        ncant_ant = Val("" & .Fields("cant_ant"))
                        cnn_dbbancos.Execute "update tb_detsolicitud set candis= candis+" & ncant_ant & "-" & _
                        Cant & " where cod_solicitud='" & _
                        Txt_NumSolComp.Text & "' and cod_producto='" & codprod & "'"
                    End If
                    .MoveNext
                Loop
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rst!Cant <= 0 Then
                    cnn_dbbancos.Execute "update tb_cabsolicitud set cs_estado='A' where cod_solicitud='" & Txt_NumSolComp & "'"
                End If
                rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                wgraba = 1
            End With
        End If
        rsdetaoc.Close
    End If
    If SwRenovar = True Then
        MsgBox "Orden de Compra Renovada" & Chr(13) & Txt_NumOC.Text & " --> " & wNumOc, vbInformation, "Sistema de Logistica"
    Else
'        MsgBox "Orden de Compra Actualizada", vbInformation, "Sistema de Logistica"
    End If
    swGrabacion = False
Me.Refresh
'Me.SetFocus
DoEvents
End Sub
' verifica que no tenga sintaxsi de sql '
Function verificaSintaxsi(ByVal cad As String)
w = ""
    Dim l As Integer
    l = Len(cad)
    
    Dim Y As String
    For X = 1 To l
    Y = Mid(cad, X, 1)
     
    If Y = "´" Or Y = "'" Then
    Y = " "
    End If
    w = w & Y
    Next
    verificaSintaxsi = w
End Function

Private Sub ValidaProveedor(CodigoExterno As String)
Dim RsP As New ADODB.Recordset
Dim Amov_Prv(0 To 20) As a_grabacion
csql = "select * from identificador where codidentificador='" & CodigoExterno & "" & "'"
If RsP.State = 1 Then RsP.Close
RsP.Open csql, CnSql, 3, 1
If RsP.RecordCount > 0 Then
    RsP.MoveFirst
    If cTipoPrv = "A" Then
        Amov_Prv(0).valor = ObtieneCorrelaProveedor & ""
    Else
        Amov_Prv(0).valor = wcodcliprov & ""
    End If
    Amov_Prv(0).campo = "f2codprov": Amov_Prv(0).TIPO = "T"
    Amov_Prv(1).campo = "f2nomprov": Amov_Prv(1).valor = left(UCase(Replace(RsP!Descripcion & "", "'", "´")), 70): Amov_Prv(1).TIPO = "T"
    Amov_Prv(2).campo = "f2newruc": Amov_Prv(2).valor = left(RsP!ruc & "", 11): Amov_Prv(2).TIPO = "T"
    Amov_Prv(3).campo = "f2nomabrev": Amov_Prv(3).valor = RsP!Abreviatura & "": Amov_Prv(3).TIPO = "T"
    Amov_Prv(4).campo = "f2dirprov": Amov_Prv(4).valor = UCase(left(RsP!direccion & "", 100)): Amov_Prv(4).TIPO = "T"
    Amov_Prv(5).campo = "f7codpos": Amov_Prv(5).valor = left(RsP!codpostal & "", 2): Amov_Prv(5).TIPO = "T"
    Amov_Prv(6).campo = "f2telprov"
    Amov_Prv(6).valor = RsP!telefono1 & IIf(Len(Trim(RsP!telefono2 & "")) > 0, " / " & RsP!telefono2, "") & IIf(Len(Trim(RsP!telefono3 & "")) > 0, " / " & RsP!telefono3, "")
    Amov_Prv(6).TIPO = "T"
    Amov_Prv(7).campo = "f2faxprov": Amov_Prv(7).valor = RsP!fax & "": Amov_Prv(7).TIPO = "T"
    Amov_Prv(8).campo = "f2email": Amov_Prv(8).valor = RsP!EMAIL & "": Amov_Prv(8).TIPO = "T"
    Amov_Prv(9).campo = "f2tipprov": Amov_Prv(9).valor = IIf(RsP!naturaljuridica = 0, "N", "J") & "": Amov_Prv(9).TIPO = "T"
    Amov_Prv(10).campo = "usuario_crea": Amov_Prv(10).valor = RsP!CREACIONUSUARIO & "": Amov_Prv(10).TIPO = "T"
    Amov_Prv(11).campo = "fecha_creacion": Amov_Prv(11).valor = Format(RsP!creacionfecha, "dd/mm/yyyy hh:mm:ss") & "": Amov_Prv(11).TIPO = "F"
    Amov_Prv(12).campo = "usuario_mod": Amov_Prv(12).valor = RsP!actualizacionUsuario & "": Amov_Prv(12).TIPO = "T"
    Amov_Prv(13).campo = "fecha_modificacion": Amov_Prv(13).valor = Format(RsP!actualizacionfecha, "dd/mm/yyyy hh:mm:ss") & "": Amov_Prv(13).TIPO = "F"
    GRABA_REGISTRO Amov_Prv, "EF2PROVEEDORES", cTipoPrv, 13, StrConexDbBancos, "F2CODPROV='" & wcodcliprov & "'"
End If
End Sub

Private Function ObtieneCorrelaProveedor()
ObtieneCorrelaProveedor = "0001"
Dim RsC As New ADODB.Recordset
csql = "select top 1 f2codprov from ef2proveedores order by f2codprov desc"
Set RsC = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
If RsC.RecordCount > 0 Then
    ObtieneCorrelaProveedor = Format(Val(RsC!F2CODPROV & "") + 1, "0000")
End If
If RsC.State = 1 Then RsC.Close
Set RsC = Nothing
End Function

Private Sub ValidaProductos(CodigoExterno As String)
Dim RsP As New ADODB.Recordset
Dim Amov_Prv(0 To 20) As a_grabacion
csql = "select * from insumo where codinsumo='" & CodigoExterno & "" & "'"
If RsP.State = 1 Then RsP.Close
RsP.Open csql, CnSql, 3, 1
If RsP.RecordCount > 0 Then
    RsP.MoveFirst
    
    Amov_Prv(0).campo = "f5codpro": Amov_Prv(0).valor = CodigoExterno & "": Amov_Prv(0).TIPO = "T"
    Amov_Prv(1).campo = "f5codfab": Amov_Prv(1).valor = CodigoExterno & "": Amov_Prv(1).TIPO = "T"
    Amov_Prv(2).campo = "f5tipo": Amov_Prv(2).valor = "P": Amov_Prv(2).TIPO = "T"
    Amov_Prv(3).campo = "F5NOMPRO": Amov_Prv(3).valor = UCase(Replace(RsP!Descripcion & "", "'", "´")): Amov_Prv(3).TIPO = "T"
    Amov_Prv(4).campo = "F5MARCA": Amov_Prv(4).valor = "001": Amov_Prv(4).TIPO = "T"
    Amov_Prv(5).campo = "F5AFECTO": Amov_Prv(5).valor = IIf(RsP!AFECTOIGV = True, "*", ""): Amov_Prv(5).TIPO = "T"
    Amov_Prv(6).campo = "f7codmed": Amov_Prv(6).valor = RsP!codunidad: Amov_Prv(6).TIPO = "T"
    Amov_Prv(7).campo = "f5moneda": Amov_Prv(7).valor = "S": Amov_Prv(7).TIPO = "T"
    Amov_Prv(8).campo = "f5fecing": Amov_Prv(8).valor = Format(RsP!creacionfecha, "dd/mm/yyyy hh:mm:ss") & "": Amov_Prv(8).TIPO = "F"
    Amov_Prv(9).campo = "f5fecmod": Amov_Prv(9).valor = Format(RsP!actualizacionfecha, "dd/mm/yyyy hh:mm:ss") & "": Amov_Prv(9).TIPO = "F"
    Amov_Prv(10).campo = "f5usermod": Amov_Prv(10).valor = RsP!actualizacionUsuario & "": Amov_Prv(10).TIPO = "T"
    Amov_Prv(11).campo = "F5USERING": Amov_Prv(11).valor = RsP!CREACIONUSUARIO & "": Amov_Prv(11).TIPO = "T"
    Amov_Prv(12).campo = "F5CTACON": Amov_Prv(12).valor = RsP!codinsumoanterior & "": Amov_Prv(12).TIPO = "T"
    GRABA_REGISTRO Amov_Prv, "if5pla", cTipoProd, 12, StrConexDbBancos, "F5CODPRO='" & CodigoExterno & "'"
End If

End Sub
