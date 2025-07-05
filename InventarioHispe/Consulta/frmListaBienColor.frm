VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmListaBienColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Colores"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   1710
   ClientWidth     =   5655
   Icon            =   "frmListaBienColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5655
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      Width           =   3495
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbBienColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaBienColor.frx":058A
      ToolBars        =   "frmListaBienColor.frx":2BAD
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgBienColor 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmListaBienColor.frx":2C68
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "frmListaBienColor"
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

Private Sub dbgBienColor_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaBienColor.Codigo = Trim(dbgBienColor.Columns.ColumnByFieldName("Codigo").Value & "")
        objAyudaBienColor.Descripcion = Trim(dbgBienColor.Columns.ColumnByFieldName("Descripcion").Value & "")
        
        'Me.Hide
        Unload Me
    Else
        With frmMantBienColor
            .Ayuda = bolAyuda
            .Codigo = Trim(dbgBienColor.Columns.ColumnByFieldName("Codigo").Value & "")
            
            .Show 1
            
            'Me.Hide
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                listarBienColorSQL
            Else
                listarBienColor
            End If
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgBienColor_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgBienColor_OnDblClick
        Case vbKeyUp
            If dbgBienColor.Dataset.RecNo = 1 Then
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
        With dbgBienColor
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
        End With
        
        listarBienColorSQL
    Else
        With dbgBienColor
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        End With
        
        listarBienColor
    End If

End Sub

Public Sub listarBienColor()
    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM EF2BIENCOLOR"
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE ESTADO = TRUE"
    End If
    
    With dbgBienColor
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODIGO"
    End With
End Sub

Public Sub listarBienColorSQL()
    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM MAESTROS.EF2BIENCOLOR"
        
    If bolAyuda Then
        strSQL = strSQL & " WHERE ESTADO = TRUE"
    End If
    
    With dbgBienColor
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODIGO"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbBienColor_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            With frmMantBienColor
                .Ayuda = bolAyuda
                .Codigo = vbNullString
                
                .Show 1
                
                'Me.Hide
                
                If Not bolAyuda Then
                    listarBienColor
                Else
                    Unload Me
                End If
            End With
        Case "Importar"
            dbgBienColor.Dataset.Close
            
            ModMilano.importarColorServidorExterno fraProceso, pgbProceso
            
            listarBienColor
        Case "Salir"
            objAyudaBienColor.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgBienColor
        .Dataset.Filtered = True
        .Dataset.Filter = "CODIGO LIKE '*" & txtbusqueda.Text & "*' " & _
                            "OR DESCRIPCION LIKE '*" & txtbusqueda.Text & "*'"
        
        If Trim(txtbusqueda.Text) = vbNullString Then
            .Dataset.Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgBienColor.SetFocus
    End Select
End Sub

