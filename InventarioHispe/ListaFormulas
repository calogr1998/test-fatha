VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form ListaFormulas 
   Caption         =   "Listado de Formulas"
   ClientHeight    =   6885
   ClientLeft      =   1335
   ClientTop       =   1245
   ClientWidth     =   10380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10380
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "ListaFormulas.frx":0000
      ToolBars        =   "ListaFormulas.frx":3F48
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1650
      Left            =   90
      OleObjectBlob   =   "ListaFormulas.frx":4054
      TabIndex        =   1
      Top             =   5085
      Width           =   10185
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4395
      Left            =   135
      OleObjectBlob   =   "ListaFormulas.frx":6D7D
      TabIndex        =   2
      Top             =   495
      Width           =   10170
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Anterior"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   4650
   End
End
Attribute VB_Name = "ListaFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS              As New ADODB.Recordset
Dim rs1             As New ADODB.Recordset
Dim Num             As Long
Dim VSize           As Long
Dim Nueva_Columna   As dxGridColumn
Dim Values()        As Variant
Dim contador        As Long
Dim CODDOC          As String
Dim csql            As String
Dim ctipodocu       As String
'Dim fgrupo         As String
Dim fcodigo         As String
Private Sub dxDBGrid1_OnClick()

        ggrupo = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4GRUPO").Value)
        fcodigo = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4CODPRO").Value)
        csql = "SELECT DISTINCTROW F3GRUPOINS, F3CODPROINS, F3NOMPRO, F3CANTIDAD, F3UNIDAD From IF3FORMULA " & _
                  " Where ((F3GRUPO = '" & ggrupo & "') And (F3CODPRO = '" & fcodigo & "')) ORDER BY F3GRUPOINS, F3CODPROINS;"
        
        With dxDBGrid2
            .Dataset.Active = False
            .Dataset.ADODataset.CommandText = csql
            .Dataset.Active = True
            .KeyField = "F3CODPROINS"
        End With
 
End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    dxDBGrid2.Dataset.Close
    
    sw_nuevo_doc = False
    Me.MousePointer = 11
    frmformula.Show 1
    Me.MousePointer = 1

End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)

        If KeyCode = 38 Or KeyCode = 40 Then
            ggrupo = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4GRUPO").Value)
            fcodigo = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("F4CODPRO").Value)
            csql = "SELECT DISTINCTROW F3GRUPOINS, F3CODPROINS, F3NOMPRO, F3CANTIDAD, F3UNIDAD From IF3FORMULA " & _
                      " Where ((F3GRUPO = '" & ggrupo & "') And (F3CODPRO = '" & fcodigo & "')) ORDER BY F3GRUPOINS, F3CODPROINS;"
            
            With dxDBGrid2
                .Dataset.Active = False
                .Dataset.ADODataset.CommandText = csql
                .Dataset.Active = True
                .KeyField = "F3CODPROINS"
            End With
        End If

End Sub

Private Sub dxDBGrid1_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)
    
    lblTitle.Visible = True
    
End Sub

Private Sub Form_Load()

    Me.AutoRedraw = False
    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    Me.AutoRedraw = True
    FILL
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblTitle.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
dxDBGrid1.Dataset.Close
dxDBGrid2.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            
            dxDBGrid2.Dataset.Close
            sw_nuevo_doc = True
            Me.MousePointer = 11
            frmformula.Show 1
            Me.MousePointer = 1
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

Private Sub FILL()
    
    dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    With dxDBGrid1
    
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
            SQL = "SELECT DISTINCTROW F4GRUPO, F4CODPRO, F4NOMPRO, F4MEDIDA, F4TIEMPOPREP From IF4FORMULA " & _
                      " Where ((F4NOMPRO <> '')) ORDER BY F4GRUPO, F4CODPRO;"

        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SQL
        .Dataset.Active = True
        .KeyField = "F4CODPRO"
    End With
    dxDBGrid1.Columns(0).Color = &HC0FFFF
    dxDBGrid1.Columns(1).Color = &HC0FFFF
    dxDBGrid1.Columns(2).Color = &HC0FFFF
    dxDBGrid1.Columns(3).Color = &HC0FFFF
    dxDBGrid1.Columns(4).Color = &HC0FFFF
    
    dxDBGrid1_OnClick
End Sub
