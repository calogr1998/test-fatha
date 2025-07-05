VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form lista_importa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestar Importación"
   ClientHeight    =   6735
   ClientLeft      =   1845
   ClientTop       =   1785
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lista_Importa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6195
      Left            =   60
      OleObjectBlob   =   "Lista_Importa.frx":000C
      TabIndex        =   0
      Top             =   420
      Width           =   10170
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo    "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro      "
            Object.ToolTipText     =   "Activadr Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir       "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   5340
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Importa.frx":28AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Importa.frx":2E45
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Importa.frx":33DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Importa.frx":3979
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5670
      TabIndex        =   1
      Top             =   90
      Width           =   4650
   End
End
Attribute VB_Name = "lista_importa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql         As String
Dim EditLookUp  As Boolean

Public Sub LLENADO()
    
With dxDBGrid1
    .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    sql = "SELECT IMPORT_CAB.F4NumImp,iif(len(trim(IMPORT_CAB.F4NumOri))>0,'S','N') as TC, EF2PROVEEDORES.F2NOMPROV, EF2CLIENTES.F2NOMCLI, "
    sql = sql & "IMPORT_CAB.F4Fecha, IIf(IMPORT_CAB.F4CERRADO='S','S','N') AS f4cerrado "
    sql = sql & "FROM (IMPORT_CAB LEFT JOIN EF2PROVEEDORES "
    sql = sql & "ON IMPORT_CAB.F4Importador = EF2PROVEEDORES.F2CODPROV) "
    sql = sql & "LEFT JOIN EF2CLIENTES ON IMPORT_CAB.F4Cliente = EF2CLIENTES.F2CODCLI "
    sql = sql & "ORDER BY Val(F4NUMIMP) DESC"

    .Dataset.Active = False
    .Dataset.ADODataset.CommandText = sql
    .Dataset.Active = True
    .KeyField = "F4NUMIMP"
    .M.FullExpand
End With

End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Trab. Comp." And Text = "S" Then
    Color = RGB(255, 128, 128)
End If
End Sub

Private Sub dxDBGrid1_OnDblClick()

If dxDBGrid1.Dataset.RecordCount > 0 Then
    sw_nuevo_documento = False
    GOC = dxDBGrid1.Columns(0).Value
    Me.MousePointer = vbHourglass
    
    
    'New_Importaciones.Show vbModal
    Me.MousePointer = vbDefault
End If
End Sub

Private Sub Form_Activate()
LLENADO
End Sub

Private Sub Form_Load()
    
    Me.AutoRedraw = False
    'Me.Height = 8040
    'Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            'Importaciones.wnumimp = "0"
'            Importaciones.Show vbModal
            
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

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Nuevo"
    sw_nuevo_documento = True
    'New_Importaciones.wnumimp = "0"
    
'    New_Importaciones.Show 1
Case "Filtro"
        If dxDBGrid1.Filter.FilterActive = True Then
            dxDBGrid1.Filter.FilterActive = False
            Me.Toolbar.Buttons.ITEM(4).Image = 2
            Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Activar Filtro"
        Else
            dxDBGrid1.Filter.FilterActive = True
            Me.Toolbar.Buttons.ITEM(4).Image = 4
            Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Desactivar Filtro"
        End If
Case "Salir"
    Unload Me
End Select
End Sub
