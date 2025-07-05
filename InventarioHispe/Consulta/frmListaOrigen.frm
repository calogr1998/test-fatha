VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmListaOrigen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Conceptos de Movimiento de Almacen"
   ClientHeight    =   7950
   ClientLeft      =   3690
   ClientTop       =   2250
   ClientWidth     =   7320
   Icon            =   "frmListaOrigen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   7320
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   " Buscar "
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbOrigen 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaOrigen.frx":058A
      ToolBars        =   "frmListaOrigen.frx":2BAD
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgOrigen 
      Height          =   6615
      Left            =   120
      OleObjectBlob   =   "frmListaOrigen.frx":2C40
      TabIndex        =   0
      Top             =   840
      Width           =   7095
   End
End
Attribute VB_Name = "frmListaOrigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda        As Boolean

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Private Sub dbgOrigen_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaOrigen.Codigo = Trim(dbgOrigen.Columns.ColumnByFieldName("F1CODORI").Value & "")
        objAyudaOrigen.Descripcion = Trim(dbgOrigen.Columns.ColumnByFieldName("F1NOMORI").Value & "")
        
        Me.Hide
        'Unload Me
    Else
        With frmMantOrigen
            .Ayuda = bolAyuda
            .Codigo = Trim(dbgOrigen.Columns.ColumnByFieldName("F1CODORI").Value & "")
            
            .Show 1
            
            'Me.Hide
            
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                listarOrigenSQL
            Else
                listarOrigen
            End If
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgOrigen_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgOrigen_OnDblClick
        Case vbKeyUp
            If dbgOrigen.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    End If
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With dbgOrigen
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
        End With
        
        listarOrigenSQL
    Else
        With dbgOrigen
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        End With
        
        listarOrigen
    End If
    
    
End Sub

Public Sub listarOrigen()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "IIF(SF1ORIGENES.F1TIPMOV = 'I', 'Conceptos de Ingreso', 'Conceptos de Salida') AS TIPOMOVIMIENTO, "
    strSQL = strSQL & "SF1ORIGENES.* "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "SF1ORIGENES"
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "ESTADO = TRUE"
    End If
    
    With dbgOrigen
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F1CODORI"
        
        .m.FullExpand
    End With
End Sub

Public Sub listarOrigenSQL()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "IIF(MAESTROS.SF1ORIGENES.F1TIPMOV = 'I', 'Conceptos de Ingreso', 'Conceptos de Salida') AS TIPOMOVIMIENTO, "
    strSQL = strSQL & "SF1ORIGENES.* "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "MAESTROS.SF1ORIGENES"
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "ESTADO = TRUE"
    End If
    
    With dbgOrigen
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F1CODORI"
        
        .m.FullExpand
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbOrigen_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            With frmMantOrigen
                .Ayuda = bolAyuda
                .Codigo = vbNullString
                
                .Show 1
                
                'Me.Hide
                
                If Not bolAyuda Then
                    listarOrigen
                Else
                    Unload Me
                End If
            End With
        Case "Salir"
            objAyudaOrigen.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgOrigen
        .Dataset.Filtered = True
        .Dataset.Filter = "F1CODORI LIKE '*" & txtBusqueda.Text & "*' " & _
                            "OR F1NOMORI LIKE '*" & txtBusqueda.Text & "*'"
        
        If Trim(txtBusqueda.Text) = vbNullString Then
            .Dataset.Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgOrigen.SetFocus
    End Select
End Sub

