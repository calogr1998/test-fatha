VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ayuda_Proveedores 
   Caption         =   "Ayuda de Proveedores"
   ClientHeight    =   5910
   ClientLeft      =   2940
   ClientTop       =   1890
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_proveedores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9540
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4665
      Left            =   120
      OleObjectBlob   =   "ayuda_proveedores.frx":058A
      TabIndex        =   1
      Top             =   1320
      Width           =   9450
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   9450
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   9090
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      ButtonWidth     =   1879
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
            Enabled         =   0   'False
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro     "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            Object.ToolTipText     =   "Agrupar Columnas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
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
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":2EC9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":3463
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":39FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":3F97
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":4531
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":4ACB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_proveedores.frx":5065
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Ayuda_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim csql        As String
Dim sw_limpia   As Boolean

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
If UCase(Column.FieldName) = "F2ORDEN" Then
    dxDBGrid1.Dataset.Edit
    If Column.value = True Then
        dxDBGrid1.Columns.ColumnByFieldName("F2ORDEN").value = False
    Else
        dxDBGrid1.Columns.ColumnByFieldName("F2ORDEN").value = True
    End If
    dxDBGrid1.Dataset.Post
End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodcliprov = "" & dxDBGrid1.Columns.ColumnByFieldName("F2CODPROV").value
    wRucCliProv = "" & dxDBGrid1.Columns.ColumnByFieldName("F2NEWRUC").value
    wnomcliprov = "" & dxDBGrid1.Columns.ColumnByFieldName("F2NOMPROV").value
    wocompra = IIf("" & dxDBGrid1.Columns.ColumnByFieldName("F2orden").value = True, "*", "")
    sw_limpia = True
    txtbusqueda.Text = ""
    sw_limpia = False
    
    Me.Hide
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

Private Sub FILL()
    csql = "SELECT F2CODPROV,F2NEWRUC,F2NOMPROV, f2orden " & _
           "FROM EF2PROVEEDORES " & _
           "ORDER BY F2NOMPROV"

    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = StrConexDbBancos
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODPROV"
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            wcodcliprov = ""
            wRucCliProv = ""
            wnomcliprov = ""
            sw_limpia = True
            txtbusqueda.Text = ""
            sw_limpia = False
            Me.Hide
        Case vbKeyUp
            'If dxDBGrid1.Dataset.RecNo = 1 Then
                txtbusqueda.SetFocus
            'End If
    End Select
End Sub

Private Sub Form_Activate()
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    dxDBGrid1.Columns.FocusedIndex = 2
'    dxDBGrid1.SetFocus
    dxDBGrid1.OptionEnabled = 1
'    TxtBusqueda.SetFocus
FILL
End Sub

Private Sub Form_Load()

    sw_limpia = False
    
    FILL
End Sub

Private Sub Form_Resize()
On Error Resume Next
fraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtbusqueda.Width = fraBusqueda.Width - 400
dxDBGrid1.Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Nuevo"
    sw_nuevo_doc = True
    sw_nuevo_documento = True
    sw_mant_ayuda = True
    Mant_Proveedores.Show 1
                  
    If sw_mant_ayuda = False Or sw_frm = True Then
    Unload Me
    End If
Case "Filtro"
    If dxDBGrid1.Filter.FilterActive = True Then
        dxDBGrid1.Filter.FilterActive = False
        Me.Toolbar.Buttons.ITEM(4).Image = 2
        Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Activar Filtro"
    Else
        dxDBGrid1.Filter.FilterActive = True
        Me.Toolbar.Buttons.ITEM(4).Image = 5
        Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Desactivar Filtro"
    End If
Case "Agrupar"
    If Button.ToolTipText = "Desagrupar Columnas" Then
        dxDBGrid1.Options.Unset (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(5).Image = 3
        Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Agrupar Columnas"
    Else
        dxDBGrid1.Options.Set (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(5).Image = 6
        Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Desagrupar Columnas"
    End If
Case "Salir"
    Me.Hide
End Select
End Sub

Private Sub txtbusqueda_Change()
Dim cGrupo  As String

    If sw_limpia = False Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "F2NEWRUC LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NOMPROV LIKE '*" & txtbusqueda.Text & "*'"
        
        If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 2
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F2NEWRUC LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NOMPROV LIKE '*" & txtbusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
        Call txtbusqueda_KeyDown(40, 0)
    End If
End Sub
