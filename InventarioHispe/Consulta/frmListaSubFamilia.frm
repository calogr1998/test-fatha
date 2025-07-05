VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmListaSubFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Sub-Familias"
   ClientHeight    =   7830
   ClientLeft      =   2460
   ClientTop       =   1725
   ClientWidth     =   7200
   Icon            =   "frmListaSubFamilia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7200
   Begin VB.Frame fraBusqueda 
      Caption         =   " Buscar "
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   5280
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
   Begin ActiveToolBars.SSActiveToolBars tlbSubFamilia 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaSubFamilia.frx":058A
      ToolBars        =   "frmListaSubFamilia.frx":2BAD
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgSubFamilia 
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "frmListaSubFamilia.frx":2C68
      TabIndex        =   4
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmListaSubFamilia"
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

Private Sub dbgSubFamilia_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaSubFamilia.Codigo = Trim(dbgSubFamilia.Columns.ColumnByFieldName("F7CODCON").Value & "")
        objAyudaSubFamilia.Descripcion = Trim(dbgSubFamilia.Columns.ColumnByFieldName("F7DESCON").Value & "")
        
        'Me.Hide
        Unload Me
    Else
'        With frmMantBienColor
'            .Ayuda = bolAyuda
'            .Codigo = Trim(dbgSubFamilia.Columns.ColumnByFieldName("Codigo").value & "")
'
'            .Show 1
'
'            'Me.Hide
'
'            listarSubFamilia
'        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgSubFamilia_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgSubFamilia_OnDblClick
        Case vbKeyUp
            If dbgSubFamilia.Dataset.RecNo = 1 Then
                txtbusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    End If
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With dbgSubFamilia
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
        End With
        
        listarSubFamiliaSQL
    Else
        With dbgSubFamilia
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        End With
        
        listarSubFamilia
    End If
    
End Sub

Public Sub listarSubFamiliaSQL()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM MAESTROS.SF7NIVEL02 "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CODEXTERNO <> ''"
        
    If bolAyuda Then
        strSQL = strSQL & " AND ESTADO = TRUE"
    End If
    
    With dbgSubFamilia
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F7CODCON"
    End With
End Sub

Public Sub listarSubFamilia()
    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM SF7NIVEL02 "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TRIM(CODEXTERNO & '') <> ''"
        
    If bolAyuda Then
        strSQL = strSQL & " AND ESTADO = TRUE"
    End If
    
    With dbgSubFamilia
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F7CODCON"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

'Private Sub tlbSubFamilia_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'    Select Case Tool.Id
'        Case "Nuevo"
'            'With frmMantBienColor
'            With frmMantSubFamilia
'                .Ayuda = bolAyuda
'                .Codigo = vbNullString
'
'                .Show 1
'
'                'Me.Hide
'
'                If Not bolAyuda Then
'                    listarSubFamilia
'                Else
'                    Unload Me
'                End If
'            End With
'        Case "Importar"
'            dbgSubFamilia.Dataset.Close
'
'            ModMilano.importarSubFamiliaServidorExterno fraProceso, pgbProceso
'
'            listarSubFamilia
'        Case "Salir"
'            objAyudaSubFamilia.inicializarEntidades
'
'            Unload Me
'    End Select
'End Sub

Private Sub txtbusqueda_Change()
    With dbgSubFamilia
        .Dataset.Filtered = True
        .Dataset.Filter = "F7CODCON LIKE '*" & txtbusqueda.Text & "*' " & _
                            "OR F7DESCON LIKE '*" & txtbusqueda.Text & "*'"
        
        If Trim(txtbusqueda.Text) = vbNullString Then
            .Dataset.Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgSubFamilia.SetFocus
    End Select
End Sub

