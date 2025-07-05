VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ayuda_grupoflujo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Grupos de Flujo"
   ClientHeight    =   5100
   ClientLeft      =   3135
   ClientTop       =   2325
   ClientWidth     =   5145
   Icon            =   "ayuda_grupoflujo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3855
      Left            =   60
      OleObjectBlob   =   "ayuda_grupoflujo.frx":000C
      TabIndex        =   2
      Top             =   1260
      Width           =   5010
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   635
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "ene"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "feb"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
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
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_grupoflujo.frx":2335
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_grupoflujo.frx":28CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_grupoflujo.frx":2E69
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_grupoflujo.frx":3403
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_grupoflujo.frx":399D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   5010
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   4620
      End
   End
End
Attribute VB_Name = "ayuda_grupoflujo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql        As String


Private Sub dxDBGrid1_OnDblClick()
    wcodgrupo = dxDBGrid1.Columns.ColumnByFieldName("codigo").value & ""
    wdesgrupo = UCase(dxDBGrid1.Columns.ColumnByFieldName("nombre").value & "")
'    wcodclicosto = dxDBGrid1.Columns.ColumnByFieldName("F3CODCLI").Value & ""
'    wrescosto = dxDBGrid1.Columns.ColumnByFieldName("F3RESPONSABLE").Value & ""
    Me.Hide
    
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
    If cnn_DbBancos.State = 1 Then cnn_DbBancos.Close
    cnn_DbBancos.Open
    csql = "select CODIGO,NOMBRE,INDICE from GRUPOS_FLUJO WHERE left(codigo,1)='" & wDestino & "' order by CODIGO"
    With dxDBGrid1
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = cnn_DbBancos
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "CODIGO"
    End With
'    CABECERA

'    dxDBGrid1.Columns(0).Color = &HC0FFFF
'    dxDBGrid1.Columns(1).Color = &HC0FFFF
'    dxDBGrid1.Columns(2).Color = &HC0FFFF
       
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            wcodgrupo = ""
            wdesgrupo = ""
            sw_limpia = True
            txtbusqueda.Text = ""
            sw_limpia = False
            Me.Hide
'        Case vbKeyInsert:
'            sw_load_mant = True
'            sw_nuevo_mant = True
'            addCliFac = False
'            If gtipodocu = "F" Then addCliFac = True
'            mant_clientes.Show 1
'            addCliFac = False
'            dxDBGrid1.Dataset.Refresh
'            Me.Hide

    End Select
    
    
End Sub

Private Sub Form_Activate()

FILL
dxDBGrid1.Filter.FilterActive = False
    
txtbusqueda.Text = ""
txtbusqueda.SetFocus

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyInsert Then
'        sw_load_mant = True
'        sw_nuevo_mant = True
'        mant_clientes.Show 1
'        dxDBGrid1.Dataset.Refresh
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close


End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'Select Case Tool.Id
'    Case "ID_Nuevo"
''        Screen.MousePointer = vbhourglass
'        sw_load_mant = True
'        sw_nuevo_mant = True
'        addCliFac = False
''        me.MousePointer = vbdefault
'        mant_clientes.Show 1
'        addCliFac = False
'        dxDBGrid1.Dataset.Refresh
'        Unload Me
'    Case "ID_Salir"
'        Unload Me
'End Select

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Nuevo"
    Me.MousePointer = vbHourglass
    sw_nuevo_doc = True
    cGrupo = "P"
    Mant_GrupoFlujo.Show 1
    Me.MousePointer = vbDefault
Case "Salir"
    Unload Me
End Select
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "CODIGO LIKE '*" & txtbusqueda.Text & "*' OR " & " NOMBRE LIKE '*" & txtbusqueda.Text & "*' "
    
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
''        If Len(Trim(txtbusqueda.Text)) > 0 Then
''            dxDBGrid1.Dataset.Filtered = True
''            dxDBGrid1.Dataset.Filter = "f2codcli LIKE '*" & txtbusqueda.Text & "*' OR " & " f2nomcli LIKE '*" & txtbusqueda.Text & "*' OR f2newruc LIKE '*" & txtbusqueda.Text & "*' "
''        Else
''            dxDBGrid1.Dataset.Filtered = False
''        End If
    End If
    
End Sub






