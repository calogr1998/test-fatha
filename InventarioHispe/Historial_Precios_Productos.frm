VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Historial_Precios_Productos 
   BorderStyle     =   0  'None
   ClientHeight    =   7995
   ClientLeft      =   7080
   ClientTop       =   4395
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraProductos 
      Caption         =   "Búsqueda"
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.TextBox TxtBusqueda 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4035
      Left            =   840
      OleObjectBlob   =   "Historial_Precios_Productos.frx":0000
      TabIndex        =   2
      Top             =   2460
      Width           =   5595
   End
End
Attribute VB_Name = "Historial_Precios_Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dxDBGrid_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim StrCodigo1 As String
Dim StrCodigo2 As String
StrCodigo1 = "'" & dxDBGrid.Columns.ColumnByFieldName("F5CODPRO").Value & "'"
StrCodigo2 = ",'" & dxDBGrid.Columns.ColumnByFieldName("F5CODPRO").Value & "'"
    
dxDBGrid.Dataset.Edit
dxDBGrid.Columns.ColumnByFieldName("GRAFICA").Value = Not dxDBGrid.Columns.ColumnByFieldName("GRAFICA").Value
dxDBGrid.Dataset.Post

'**********************
If dxDBGrid.Columns.ColumnByFieldName("GRAFICA").Value = True Then
    If Len(Trim(Historial_Precios.MSPrecios.Title)) = 0 Then
        Historial_Precios.MSPrecios.Title = Historial_Precios.MSPrecios.Title & StrCodigo1
    Else
        Historial_Precios.MSPrecios.Title = Historial_Precios.MSPrecios.Title & StrCodigo2
    End If
Else
    Historial_Precios.MSPrecios.Title = Replace(Historial_Precios.MSPrecios.Title, StrCodigo2, "")
    Historial_Precios.MSPrecios.Title = Replace(Historial_Precios.MSPrecios.Title, StrCodigo1, "")
    If InStr(Historial_Precios.MSPrecios.Title, "'") = 2 Then
        Historial_Precios.MSPrecios.Title = Replace(Historial_Precios.MSPrecios.Title, ",", "")
    End If
End If

Historial_Precios.Carga_Grafica Historial_Precios.MSPrecios.Title, Historial_Precios.aboDesde.Value, Historial_Precios.abohasta.Value
End Sub

Private Sub Form_Resize()
On Error Resume Next

FraProductos.Move 0, 0, Me.ScaleWidth, 870
txtbusqueda.Move 200, 350, FraProductos.Width - 400, 310
dxDBGrid.Move 0, FraProductos.Height, Me.ScaleWidth, Me.ScaleHeight - FraProductos.Height

End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid.Dataset.Filtered = True
    dxDBGrid.Dataset.Filter = "F5CODPRO LIKE '" & txtbusqueda.Text & "*' OR " & " F5NOMPRO LIKE '*" & txtbusqueda.Text & "*'"
    
    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid.Columns.FocusedIndex = 1
        dxDBGrid.SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dxDBGrid.SetFocus
    End If
End Sub
