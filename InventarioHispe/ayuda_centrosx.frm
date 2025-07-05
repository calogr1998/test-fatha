VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_centrosx 
   Caption         =   "Ayuda de Centros de Costos"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   4890
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   735
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ayuda_centrosx.frx":0000
      ToolBars        =   "ayuda_centrosx.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "ayuda_centrosx.frx":663F
      TabIndex        =   3
      Top             =   960
      Width           =   4890
   End
End
Attribute VB_Name = "ayuda_centrosx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String


Private Sub dxDBGrid1_OnDblClick()
    wcodcosto = dxDBGrid1.Columns.ColumnByFieldName("F3COSTO").value & ""
    wdescosto = dxDBGrid1.Columns.ColumnByFieldName("F3DESCRIP").value & ""
    wunicosto = dxDBGrid1.Columns.ColumnByFieldName("F3ABREV").value & ""
    wclicosto = dxDBGrid1.Columns.ColumnByFieldName("F3CODCLI").value & ""
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
    
    csql = "select F3COSTO,F3ABREV,F3DESCRIP,PO,F3CODCLI from CENTROS order by F3COSTO"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F3COSTO"
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
            wcodcosto = ""
            wdescosto = ""
            wunicosto = ""
            wclicosto = ""
            sw_limpia = True
            txtBusqueda.Text = ""
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
    
txtBusqueda.Text = ""
txtBusqueda.SetFocus

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyInsert Then
'        sw_load_mant = True
'        sw_nuevo_mant = True
'        mant_clientes.Show 1
'        dxDBGrid1.Dataset.Refresh
'    End If
End Sub

Private Sub Form_Load()
    If cnn_mov.State = adStateOpen Then cnn_mov.Close
    cnn_mov.ConnectionString = cnn_dbbancos
    cnn_mov.Open cconexion
            
    With dxDBGrid1
'        .DefaultFields = True
        .Dataset.ADODataset.ConnectionString = cnn_mov
    End With
    FILL
    dxDBGrid1.Filter.FilterActive = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
    cnn_mov.Close

End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
'        Screen.MousePointer = vbhourglass
        sw_load_mant = True
        sw_nuevo_doc = True
'        addCliFac = False
'        Screen.MousePointer = 0
        frmcentros.Show 1
        sw_load_mant = False
'        addCliFac = False
'        dxDBGrid1.Dataset.Refresh
        Unload Me
    Case "ID_Salir"
        Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "po LIKE '*" & txtBusqueda.Text & "*' OR f3COSTO LIKE '*" & txtBusqueda.Text & "*' OR " & " f3DESCRIP LIKE '*" & txtBusqueda.Text & "*' OR " & " f3ABREV LIKE '*" & txtBusqueda.Text & "*' "
    
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

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
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





