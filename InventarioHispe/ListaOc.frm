VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form Lista_Oc111 
   Caption         =   "Orden de Compra"
   ClientHeight    =   6810
   ClientLeft      =   1935
   ClientTop       =   2115
   ClientWidth     =   10455
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
   ScaleHeight     =   6810
   ScaleWidth      =   10455
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   1080
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "ListaOc.frx":0000
      ToolBars        =   "ListaOc.frx":32A8
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6195
      Left            =   90
      OleObjectBlob   =   "ListaOc.frx":336C
      TabIndex        =   0
      Top             =   495
      Width           =   10170
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
Attribute VB_Name = "Lista_Oc111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL         As String
Dim EditLookUp  As Boolean

Public Sub LLENADO()
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_compras
        SQL = "SELECT A.F4NUMORD,A.F4CODSOLICITUD,B.F2NOMPROV,A.F4FECEMI,A.F4TIPMON,A.F4MONTO FROM IF4ORDEN AS A, " & _
              "EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4ESTNUL<>'S' ORDER BY A.F4NUMORD DESC"
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SQL
        .Dataset.Active = True
        .KeyField = "F4NUMORD"
    End With

End Sub

Private Sub dxDBGrid1_OnDblClick()

    sw_nuevo_documento = False
    GOC = dxDBGrid1.Columns(0).Value
    Me.MousePointer = vbHourglass
    movocl.Show vbModal
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    
    Me.AutoRedraw = False
    Me.Height = 8040
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    sw_nuevo_documento = True
    Me.AutoRedraw = True
    
    LLENADO
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo"
                sw_nuevo_documento = True
                Me.MousePointer = 11
                movocl.Show 1
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
