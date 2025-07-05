VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_formapago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Forma de Pago"
   ClientHeight    =   5805
   ClientLeft      =   2985
   ClientTop       =   1695
   ClientWidth     =   5025
   Icon            =   "ayuda_formapago.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleMode       =   0  'User
   ScaleWidth      =   5025
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   4770
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1080
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
      Tools           =   "ayuda_formapago.frx":058A
      ToolBars        =   "ayuda_formapago.frx":6AB6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "ayuda_formapago.frx":6B84
      TabIndex        =   3
      Top             =   945
      Width           =   4770
   End
End
Attribute VB_Name = "ayuda_formapago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim csql        As String
Dim sw_limpia   As Boolean

Private Sub dxDBGrid1_OnDblClick()
    wcodpag = "" & dxDBGrid1.Columns.ColumnByFieldName("F2FORPAG").value
    xCodigo = "" & dxDBGrid1.Columns.ColumnByFieldName("F2FORPAG").value
    wnompag = "" & dxDBGrid1.Columns.ColumnByFieldName("F2DESPAG").value
    
    sw_limpia = True
    txtBusqueda.Text = ""
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

    csql = "SELECT F2FORPAG,F2DESPAG FROM EF2FORPAG ORDER BY F2FORPAG"
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2FORPAG"
    
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            wcodpag = ""
            wnompag = ""
            xCodigo = ""
            sw_limpia = True
            txtBusqueda.Text = ""
            sw_limpia = False
            Me.Hide
    End Select
    
End Sub

Private Sub Form_Activate()
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    
    dxDBGrid1.OptionEnabled = 1
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 3600
    Me.top = 1050
    
    sw_limpia = False
        
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
'        Case "ID_Nuevo":
'            sw_nuevo_doc = True
'            sw_mant_ayuda = True
'            mant_marcas.Show 1
'            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Salir":
            Unload Me
    End Select
'
End Sub

Private Sub txtbusqueda_Change()
Dim cGrupo  As String

    If sw_limpia = False Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "F2FORPAG LIKE '*" & txtBusqueda.Text & "*' OR " & " F2DESPAG LIKE '*" & txtBusqueda.Text & "*'"
        
        If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F2FORPAG LIKE '*" & txtBusqueda.Text & "*' OR " & " F2DESPAG LIKE '*" & txtBusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub
