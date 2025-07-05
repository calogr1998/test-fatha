VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmListaUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de U.M."
   ClientHeight    =   6255
   ClientLeft      =   4590
   ClientTop       =   1710
   ClientWidth     =   7095
   Icon            =   "frmListaUM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7095
   Begin VB.Frame fraBusqueda 
      Caption         =   " Buscar "
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbUM 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaUM.frx":058A
      ToolBars        =   "frmListaUM.frx":2BAD
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgUM 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmListaUM.frx":2C68
      TabIndex        =   4
      Top             =   840
      Width           =   6855
   End
End
Attribute VB_Name = "frmListaUM"
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

Private Sub dbgUM_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaUM.Codigo = Trim(dbgUM.Columns.ColumnByFieldName("F7CODMED").Value & "")
        objAyudaUM.Descripcion = Trim(dbgUM.Columns.ColumnByFieldName("F7NOMMED").Value & "")
        
        'Me.Hide
        Unload Me
    Else
'        With frmMantUM
'            .Ayuda = bolAyuda
'            .Codigo = Trim(dbgUM.Columns.ColumnByFieldName("Codigo").value & "")
'
'            .Show 1
'
'            'Me.Hide
'
'            listarUM
'        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgUM_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgUM_OnDblClick
        Case vbKeyUp
            If dbgUM.Dataset.RecNo = 1 Then
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
        With dbgUM
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
        End With
        
        listarUMSQL
    Else
        With dbgUM
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        End With
        
        listarUM
    End If
    
End Sub

Public Sub listarUMSQL()
    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM MAESTROS.EF7MEDIDAS "
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE ESTADO = TRUE"
    End If
    
    With dbgUM
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F7CODMED"
    End With
End Sub

Public Sub listarUM()
    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM EF7MEDIDAS "
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE ESTADO = TRUE"
    End If
    
    With dbgUM
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F7CODMED"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbUM_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
'            With frmMantUM
'                .Ayuda = bolAyuda
'                .Codigo = vbNullString
'
'                .Show 1
'
'                'Me.Hide
'
'                If Not bolAyuda Then
'                    listarUM
'                Else
'                    Unload Me
'                End If
'            End With
        Case "Importar"
            dbgUM.Dataset.Close
            
            listarUM
        Case "Salir"
            objAyudaUM.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgUM
        .Dataset.Filtered = True
        .Dataset.Filter = "F7CODMED LIKE '*" & txtBusqueda.Text & "*' " & _
                            "OR F7NOMMED LIKE '*" & txtBusqueda.Text & "*'"
        
        If Trim(txtBusqueda.Text) = vbNullString Then
            .Dataset.Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgUM.SetFocus
    End Select
End Sub


