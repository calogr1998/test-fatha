VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Lista_Oc_Detalle 
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   1320
   ClientTop       =   5490
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   2250
      Left            =   60
      OleObjectBlob   =   "Lista_Oc_Detalle.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   13665
   End
End
Attribute VB_Name = "Lista_Oc_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "F3CANFAL"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbRed
                Color = vbYellow
            End If
    End Select
End Sub

Private Sub Form_Resize()
    dxDBGrid2.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
