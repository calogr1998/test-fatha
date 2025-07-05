VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ListaProveDetallada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Detallada de Proveedores"
   ClientHeight    =   7065
   ClientLeft      =   1050
   ClientTop       =   2340
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10485
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   255
      TabIndex        =   3
      Top             =   150
      Width           =   1455
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1695
      TabIndex        =   2
      Top             =   150
      Width           =   2055
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "ListaProveDetallada.frx":0000
      ToolBars        =   "ListaProveDetallada.frx":4CB3
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1650
      Left            =   120
      OleObjectBlob   =   "ListaProveDetallada.frx":4D78
      TabIndex        =   0
      Top             =   5280
      Width           =   10185
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4650
      Left            =   90
      OleObjectBlob   =   "ListaProveDetallada.frx":7EAF
      TabIndex        =   1
      Top             =   480
      Width           =   10170
   End
End
Attribute VB_Name = "ListaProveDetallada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs              As New ADODB.Recordset
Dim rs1             As New ADODB.Recordset
Dim Num             As Long
Dim VSize           As Long
Dim Nueva_Columna   As dxGridColumn
Dim Values()        As Variant
Dim contador        As Long
Dim CODDOC          As String
Dim csql            As String
Dim ctipodocu       As String
Dim ccodprov         As String
Dim crucprov         As String

Private Sub Checkagrupar_Click()
    If Checkagrupar.value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If

End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxDBGrid1_OnClick()

        'ccodprov = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F2CODPROV").Value)
        'crucprov = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F2NEWRUC").Value)
        'csql = "SELECT * FROM TBVENTA_DET AS A,EF7MEDIDAS AS B WHERE F4TIPODOCU='" & ctipodocu & "' AND F4SERDOC='" & cserdoc & "' AND F4NUMDOC='" & cnumdoc & "' AND (A.F7CODMED=B.F7CODMED OR A.F7CODMED=B.F7SIGMED)"
        With dxDBGrid2
            '.Dataset.Active = False
            '.Dataset.ADODataset.CommandText = csql
            '.Dataset.Active = True
            '.KeyField = "F5CODPRO"
        End With

End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)

    'cserdoc = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4SERDOC").Value)
    'cnumdoc = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4NUMDOC").Value)
    'csql = "SELECT A.F5CODPRO,A.F5NOMPRO,B.F7SIGMED,A.F3CANPRO2,A.F3CANPRO,A.F3PREUNI,A.F3VALVTA FROM TBVENTA_DET AS A,EF7MEDIDAS AS B WHERE F4TIPODOCU='" & ctipodocu & "' AND F4SERDOC='" & cserdoc & "' AND F4NUMDOC='" & cnumdoc & "' AND (A.F7CODMED=B.F7CODMED OR A.F7CODMED=B.F7SIGMED)"
    'With dxDBGrid2
    '    .Dataset.Active = False
    '    .Dataset.ADODataset.CommandText = csql
    '    .Dataset.Active = True
    '    .KeyField = "F5CODPRO"
    'End With

End Sub

Private Sub dxDBGrid1_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)
    
    'lblTitle.Visible = True
    
End Sub

Private Sub Form_Activate()

    CONECTAR
    FILL
    
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.Height = 7890
    Me.Width = 10400
    Me.left = 1600
    Me.top = 1050
    
    dxDBGrid2.Visible = True
    dxDBGrid1.Height = 4485
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'lblTitle.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    dxDBGrid2.Dataset.Close
    
End Sub

Private Sub SSActiveToolBars1_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
        
    Call FILL
        
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
         Case "ID_Salir"
            Unload Me
    End Select

End Sub

Private Sub CONECTAR()

    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
    
    'If wf1visualiza_det_lista = "*" Then
        dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_dbbancos
        'If Len(Trim(wf1uupp)) = 0 Then
            'dxDBGrid2.Columns.ColumnByFieldName("F3CANPRO2").Visible = False
            'dxDBGrid2.Columns.ColumnByFieldName("F3CANPRO").Caption = "Cantidad"
        'End If
    'End If
    
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
Dim cIdtTip As String
'Dim nValMes As Single
'Dim nValAno As Single
    '
    If nValMes = 0 Then
        sql = "Select F2CODPROV ," & _
                     "F2NOMPROV," & _
                     "F2NEWRUC," & _
                     "F2TIPPROV " & _
                "FROM EF2PROVEEDORES " & _
              "ORDER BY F2CODPROV "
    End If
    '
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = sql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODPROV"
End Sub


