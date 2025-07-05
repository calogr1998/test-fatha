VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmListaBienAlterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Alternativas"
   ClientHeight    =   6255
   ClientLeft      =   3045
   ClientTop       =   2910
   ClientWidth     =   8655
   Icon            =   "frmListaBienAlterno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8655
   Begin VB.Frame fraBusqueda 
      Caption         =   " Buscar "
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese cadena a buscar"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbBienAlterno 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaBienAlterno.frx":058A
      ToolBars        =   "frmListaBienAlterno.frx":1EF6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgBienAlterno 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmListaBienAlterno.frx":1F87
      TabIndex        =   0
      Top             =   840
      Width           =   8415
   End
End
Attribute VB_Name = "frmListaBienAlterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda        As Boolean
Private strCodigoBien   As String

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let CodigoBien(ByVal Value As String)
    strCodigoBien = Value
End Property

Public Property Get CodigoBien() As String
    CodigoBien = strCodigoBien
End Property

Private Sub dbgBienAlterno_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaBienAlterno.CodigoBien = Trim(dbgBienAlterno.Columns.ColumnByFieldName("F5CODPRO").Value & "")
        objAyudaBienAlterno.CodigoBienAlterno = Trim(dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").Value & "")
        
        Unload Me
    Else
        With frmMantBienAlterno
            .Ayuda = bolAyuda
            .Codigo = Trim(dbgBienAlterno.Columns.ColumnByFieldName("F5CODPRO").Value & "")
            .CodigoAlterno = vbNullString
            
            .Show 1
            
            If Not .Ayuda Then
                listarBienAlterno
            Else
                'Me.Hide
                Unload Me
            End If
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgBienAlterno_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgBienAlterno_OnDblClick
        Case vbKeyUp
            If dbgBienAlterno.Dataset.RecNo = 1 Then
                txtbusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    End If
    
    With dbgBienAlterno
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
    
    listarBienAlterno
End Sub

Public Sub listarBienAlterno()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT  "
    strSQL = strSQL & "(ALTERNO.F5CODPRO & ' - ' & BIEN1.F5NOMPRO) AS PRINCIPAL, "
    strSQL = strSQL & "ALTERNO.F5CODPRO, "
    strSQL = strSQL & "ALTERNO.F5CODPROALTERNO AS CODIGO, "
    strSQL = strSQL & "BIEN2.F5NOMPRO AS DESCRIPCION "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(EF2BIENALTERNO AS ALTERNO "
    strSQL = strSQL & "LEFT JOIN IF5PLA AS BIEN1 ON BIEN1.F5CODPRO = ALTERNO.F5CODPRO) "
    strSQL = strSQL & "LEFT JOIN IF5PLA AS BIEN2 ON BIEN2.F5CODPRO = ALTERNO.F5CODPROALTERNO "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "NOT ISNULL(ALTERNO.F5CODPRO) "
    
    If bolAyuda Then
        strSQL = strSQL & "AND ALTERNO.ESTADO = TRUE "
    End If
    
    If strCodigoBien <> vbNullString Then
        strSQL = strSQL & "AND ALTERNO.F5CODPRO = '" & strCodigoBien & "' "
    End If
    
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "(ALTERNO.F5CODPRO & ' - ' & BIEN1.F5NOMPRO), ALTERNO.F5CODPROALTERNO"
    
    With dbgBienAlterno
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "CODIGO"
        
        .m.FullExpand
    End With
    
    strSQL = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
    strCodigoBien = vbNullString
End Sub

Private Sub tlbBienAlterno_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            With frmMantBienAlterno
                .Ayuda = bolAyuda
                
                If Not bolAyuda Then
                    .Codigo = vbNullString
                Else
                    .Codigo = strCodigoBien
                End If

                .Show 1

                If Not bolAyuda Then
                    listarBienAlterno
                Else
                    Unload Me
                End If
            End With
        Case "Salir"
            objAyudaBienAlterno.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgBienAlterno
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
            dbgBienAlterno.SetFocus
    End Select
End Sub

