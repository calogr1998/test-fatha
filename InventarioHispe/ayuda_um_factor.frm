VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form ayuda_um_factor 
   Caption         =   "Ayuda de Unidades de Medida"
   ClientHeight    =   4875
   ClientLeft      =   4110
   ClientTop       =   2625
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_um_factor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4770
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   315
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3900
      Left            =   45
      OleObjectBlob   =   "ayuda_um_factor.frx":058A
      TabIndex        =   0
      Top             =   900
      Width           =   4770
   End
End
Attribute VB_Name = "ayuda_um_factor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql        As String

Private Sub CABECERA()

    With dxDBGrid1
        .Columns(0).Caption = "Codigo": .Columns(0).Width = 45: .Columns(0).DisableEditor = True
        .Columns(1).Caption = "Descripcion": .Columns(1).Width = 70: .Columns(1).DisableEditor = True
        .Columns(2).Caption = "Factor": .Columns(2).Width = 35: .Columns(2).DisableEditor = True: .Columns(2).DecimalPlaces = 2
        .Columns(3).Caption = "Pre.Vta.": .Columns(3).Width = 40: .Columns(3).DisableEditor = True: .Columns(3).DecimalPlaces = 2: .Columns(3).Visible = False
    End With
End Sub

Private Sub dxDBGrid1_OnDblClick()
    If Val(dxDBGrid1.Columns(2).value & "") = 0 Then
        MsgBox "Factor de Conversion no puede ser cero, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    wcodmed = dxDBGrid1.Columns(0).value
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
    csql = vbNullString
    csql = csql & "SELECT "
    csql = csql & "B.F7CODMED, "
    csql = csql & "A.F7NOMMED, "
    csql = csql & "B.F5FACTOR, "
    csql = csql & "B.F5PREVTA "
    csql = csql & "FROM "
    csql = csql & "(MEDIVENTAS B "
    csql = csql & "LEFT JOIN EF7MEDIDAS A ON A.F7CODMED = B.F7CODMED) "
    csql = csql & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = B.F5CODPRO "
    
    csql = csql & "WHERE "
    csql = csql & "PROD.F5DESCONTINUADO = 'N' AND "
    csql = csql & "B.F5CODPRO = '" & wprodfactor & " ' "
    csql = csql & "ORDER BY "
    csql = csql & "B.F7CODMED"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F7CODMED"
    CABECERA

    dxDBGrid1.Columns(0).Color = &HC0FFFF
    dxDBGrid1.Columns(1).Color = &HC0FFFF
    dxDBGrid1.Columns(2).Color = &HC0FFFF
    dxDBGrid1.Columns(3).Color = &HC0FFFF
    
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            wcodmed = ""
            Me.Hide
End Select
End Sub

Private Sub Form_Activate()
    If Cnn_Mov.State = adStateOpen Then Cnn_Mov.Close
    Cnn_Mov.ConnectionString = cnn_dbbancos
    Cnn_Mov.Open cconexion
    With dxDBGrid1
        .DefaultFields = True
        .Dataset.ADODataset.ConnectionString = Cnn_Mov
    End With
    FILL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
    Cnn_Mov.Close
End Sub

Private Sub txtbusqueda_Change()
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "F7CODMED LIKE '*" & txtBusqueda.Text & "*' OR " & " F7NOMMED LIKE '*" & txtBusqueda.Text & "*' "
        
        If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
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
            dxDBGrid1.Dataset.Filter = "F7CODMED LIKE '*" & txtBusqueda.Text & "*' OR " & " F7NOMMED LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub
