VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_conceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Conceptos"
   ClientHeight    =   5625
   ClientLeft      =   5910
   ClientTop       =   1965
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_conceptos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
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
         TabIndex        =   2
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "B�squeda"
         Height          =   210
         Left            =   240
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
      Tools           =   "ayuda_conceptos.frx":058A
      ToolBars        =   "ayuda_conceptos.frx":6AB6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "ayuda_conceptos.frx":6BC9
      TabIndex        =   3
      Top             =   945
      Width           =   4770
   End
End
Attribute VB_Name = "ayuda_conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql        As String

Private Sub dxDBGrid1_OnDblClick()
    wconcepto = dxDBGrid1.Columns.ColumnByFieldName("f1codori").value
    wnomconcepto = dxDBGrid1.Columns.ColumnByFieldName("f1nomori").value
    txtBusqueda.Text = ""
    Me.Hide
    'Unload Me
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
    csql = ""
    If wtipmov = "I" Then csql = "Select F1CODORI,F1NOMORI FROM SF1ORIGENES WHERE F1TIPMOV='I' ORDER BY F1CODORI"
    If wtipmov = "S" Then
        csql = "SELECT SF1ORIGENES.F1CODORI, SF1ORIGENES.F1NOMORI " & _
                "FROM ALMACEN_CONCEPTO INNER JOIN SF1ORIGENES ON ALMACEN_CONCEPTO.F1CODORI = " & _
                "SF1ORIGENES.F1CODORI where ALMACEN_CONCEPTO.f2codalm = '" & wcod_alm & "';"
    End If
    If csql = "" Then
        csql = "Select F1CODORI,F1NOMORI FROM SF1ORIGENES ORDER BY F1CODORI"
    End If
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F1CODORI"
       
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
        If KeyCode = 13 Then
            dxDBGrid1_OnDblClick
        End If
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
    
    Me.left = 3600
    Me.top = 1050
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
    
    Set ayuda_conceptos = Nothing
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            mant_conceptos.Show 1
            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Salir":
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F1CODORI LIKE '*" & txtBusqueda.Text & "*' OR " & " F1NOMORI LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F1CODORI LIKE '*" & txtBusqueda.Text & "*' OR " & " F1NOMORI LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub
