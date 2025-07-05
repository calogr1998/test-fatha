VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Begin VB.Form ayuda_Gastos_Imp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Gastos de Importacion :::"
   ClientHeight    =   6372
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   5136
   Icon            =   "ayuda_Gastos_Imp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6372
   ScaleWidth      =   5136
   StartUpPosition =   2  'CenterScreen
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4428
      Left            =   120
      OleObjectBlob   =   "ayuda_Gastos_Imp.frx":000C
      TabIndex        =   3
      Top             =   1032
      Width           =   4932
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   132
      TabIndex        =   0
      Top             =   120
      Width           =   4932
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   735
      End
   End
End
Attribute VB_Name = "ayuda_Gastos_Imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String

Private Sub dxDBGrid1_OnDblClick()
    wcodgastimp = dxDBGrid1.Columns.ColumnByFieldName("F2CODIGO").Value
    wNombregasimo = dxDBGrid1.Columns.ColumnByFieldName("F2DESCRIPCION").Value
    txtbusqueda.Text = ""
    Me.Hide
    'Unload Me
End Sub

Private Sub FILL()
    csql = "Select F2CODIGO, F2DESCRIPCION From Tb_COSTOSIMP"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODIGO"
       
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
    Me.MousePointer = 11
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    
    Me.Left = 3600
    Me.Top = 1050
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    Me.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
    
    Set ayuda_Gastos_Imp = Nothing
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2DESCRIPCION LIKE '*" & txtbusqueda.Text & "*' OR " & " F2CODIGO LIKE '*" & txtbusqueda.Text & "*' "
    
    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F1CODORI LIKE '*" & txtbusqueda.Text & "*' OR " & " F1NOMORI LIKE '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub


