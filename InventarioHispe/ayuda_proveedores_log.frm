VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_proveedores_log 
   Caption         =   "Ayuda de Proveedores"
   ClientHeight    =   5745
   ClientLeft      =   1290
   ClientTop       =   1995
   ClientWidth     =   8535
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   8175
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   6300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   315
         TabIndex        =   1
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
      Tools           =   "ayuda_proveedores_log.frx":0000
      ToolBars        =   "ayuda_proveedores_log.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4245
      Left            =   120
      OleObjectBlob   =   "ayuda_proveedores_log.frx":663F
      TabIndex        =   3
      Top             =   960
      Width           =   8235
   End
End
Attribute VB_Name = "ayuda_proveedores_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String

Private Sub dxDBGrid1_OnDblClick()
    wcodprov = "" & dxDBGrid1.Columns.ColumnByFieldName("f2codprov").Value
    wrucprov = "" & dxDBGrid1.Columns.ColumnByFieldName("f2newruc").Value
    wnomprov = "" & dxDBGrid1.Columns.ColumnByFieldName("f2nomprov").Value
    wfpagoprov = "" & dxDBGrid1.Columns.ColumnByFieldName("f2forpag").Value
    wcontacto = "" & dxDBGrid1.Columns.ColumnByFieldName("f2contacto").Value
    wmoneda_productos = "" & dxDBGrid1.Columns.ColumnByFieldName("f2tipmon").Value
    
    sw_limpia = True
    txtbusqueda.Text = ""
    sw_limpia = False
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
If ctipoadm_bd = "M" Then
    csql = "select F2CODPROV,F2NEWRUC,F2NOMPROV,F2FORPAG,F2CONTACTO,f2tipprov, F2TIPMON,F2TELPROV from EF2PROVEEDORES order by f2codprov"
Else
    csql = "select F2CODPROV,F2NEWRUC,F2NOMPROV,F2FORPAG,F2CONTACTO,iif(F2TIPPROV='E','Extranjero','Nacional') as f2tipprov, F2TIPMON,F2TELPROV from EF2PROVEEDORES order by f2codprov"
End If

    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODPROV"
           
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 13 Then
        dxDBGrid1_OnDblClick
    End If
End Sub

Private Sub dxDBGrid1_OnKeyPress(Key As Integer)
    'If Key = 13 Then
    '    dxDBGrid1_OnDblClick
    'End If
End Sub

Private Sub Form_Activate()
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    dxDBGrid1.Dataset.Active = True
    FILL
    dxDBGrid1.Columns.FocusedIndex = 1
'    dxDBGrid1.SetFocus
    txtbusqueda.SetFocus
    dxDBGrid1.OptionEnabled = 1
End Sub

Private Sub Form_Load()
'    dxDBGrid1.Options.Unset (egoShowGroupPanel)
'    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 1600
    Me.top = 1050
    
    sw_limpia = False
        
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
                
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close


End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            Mant_Proveedores.Show 1
                          
            If sw_mant_ayuda = False Then
            Unload Me
'            Else
'                dxDBGrid1.Filter.FilterActive = False
'                dxDBGrid1.Dataset.ADODataset.Requery
'                dxDBGrid1.Dataset.Active = True
'                dxDBGrid1.KeyField = "F2CODPROV"
'                dxDBGrid1.Dataset.Refresh
            End If
        Case "ID_Salir":
            Unload Me
    End Select

End Sub

Private Sub txtbusqueda_Change()

    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NEWRUC LIKE '*" & txtbusqueda.Text & "*' OR " & " f2nomprov like '*" & txtbusqueda.Text & "*' or " & " f2forpag like '*" & txtbusqueda.Text & "*' or " & " f2contacto like '*" & txtbusqueda.Text & "*' or " & " f2tipprov like '*" & txtbusqueda.Text & "*' "
        
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
        'If Len(Trim(txtbusqueda.Text)) > 0 Then
        '    dxDBGrid1.Dataset.Filtered = True
        '    dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NEWRUC LIKE '*" & txtbusqueda.Text & "*' OR " & " f2nomprov like '*" & txtbusqueda.Text & "*' or " & " f2forpag like '*" & txtbusqueda.Text & "*' or " & " f2contacto like '*" & txtbusqueda.Text & "*' or " & " f2tipprov like '*" & txtbusqueda.Text & "*' "
        'Else
        '    dxDBGrid1.Dataset.Filtered = False
        'End If
        dxDBGrid1.SetFocus
    End If
    
End Sub




