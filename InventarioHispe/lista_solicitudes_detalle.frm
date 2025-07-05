VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form lista_solicitudes_detalle 
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   1170
   ClientTop       =   5685
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   2130
      Left            =   0
      OleObjectBlob   =   "lista_solicitudes_detalle.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   16005
   End
End
Attribute VB_Name = "lista_solicitudes_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dxDBGrid2_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "SALDO"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbRed
                Color = vbYellow
            ElseIf Val(Text) < 0 Then
                Font.Bold = True
                FontColor = vbWhite
                Color = vbRed
            End If
            
            Text = Format(Val(Text), "#,0.0000;(#,0.0000)")
    End Select
End Sub

Private Sub dxDBGrid2_OnDblClick()
    Select Case UCase(dxDBGrid2.Columns.FocusedColumn.FieldName)
        Case "COD_PRODUCTO"
            With frmUtilDetalleConsolidadoPedido
                .NroPedido = Trim(dxDBGrid2.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
                .CodigoProducto = Trim(dxDBGrid2.Columns.ColumnByFieldName("COD_PRODUCTO").value & "")
                
                .Show 1
            End With
        Case "CANTCOMPROMETIDA"
'            If Val(dxDBGrid2.Columns.ColumnByFieldName("CANTCOMPROMETIDA").value & "") = 0 Then
'                MsgBox "Cantidad Comprometida en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "F"
                .NroPedido = Trim(dxDBGrid2.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
                .CodigoProducto = Trim(dxDBGrid2.Columns.ColumnByFieldName("COD_PRODUCTO").value & "")
                
                .Show vbModal
            End With
        Case "CANTPORLLEGAR"
            If Val(dxDBGrid2.Columns.ColumnByFieldName("CANTPORLLEGAR").value & "") = 0 Then
                MsgBox "Cantidad por Llegar en Cero, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With frmUtilDetalleMovimientoPedido
                .TipoCompromisoForV = "V"
                .NroPedido = dxDBGrid2.Columns.ColumnByFieldName("COD_SOLICITUD").value
                .CodigoProducto = dxDBGrid2.Columns.ColumnByFieldName("COD_PRODUCTO").value
                
                .Show vbModal
            End With
    End Select
End Sub

Private Sub Form_Resize()
dxDBGrid2.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
    
