VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Lista_RegComprasDetalle 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de Compras"
   ClientHeight    =   3105
   ClientLeft      =   4110
   ClientTop       =   7815
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1950
      Left            =   0
      OleObjectBlob   =   "Lista_RegComprasDetalle.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   10005
   End
End
Attribute VB_Name = "Lista_RegComprasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
    dxDBGrid2.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If UCase(Column.FieldName) = "F3IMPORTE" Then
    Text = Format(Text, "###,###,##0.00")
End If
End Sub
