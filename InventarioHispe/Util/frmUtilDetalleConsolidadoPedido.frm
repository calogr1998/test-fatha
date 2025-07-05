VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmUtilDetalleConsolidadoPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidado de Pedido por Insumo del"
   ClientHeight    =   5700
   ClientLeft      =   480
   ClientTop       =   2610
   ClientWidth     =   12735
   Icon            =   "frmUtilDetalleConsolidadoPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12735
   Begin VB.Timer tmrTemporizador 
      Left            =   0
      Top             =   4920
   End
   Begin ActiveToolBars.SSActiveToolBars tlbDetalle 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      Tools           =   "frmUtilDetalleConsolidadoPedido.frx":058A
      ToolBars        =   "frmUtilDetalleConsolidadoPedido.frx":2BB2
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDetalle 
      Height          =   4815
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "frmUtilDetalleConsolidadoPedido.frx":2C7B
      TabIndex        =   0
      Tag             =   "Pedido"
      Top             =   120
      Width           =   12525
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDetalle 
      Height          =   4815
      Index           =   1
      Left            =   240
      OleObjectBlob   =   "frmUtilDetalleConsolidadoPedido.frx":5B07
      TabIndex        =   1
      Tag             =   "Orden(es) de Producción"
      Top             =   360
      Width           =   12525
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgDetalle 
      Height          =   4815
      Index           =   2
      Left            =   360
      OleObjectBlob   =   "frmUtilDetalleConsolidadoPedido.frx":8993
      TabIndex        =   2
      Tag             =   "Detalle de Orden de Producción"
      Top             =   600
      Width           =   12525
   End
End
Attribute VB_Name = "frmUtilDetalleConsolidadoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNroPedido As String
Private strCodProducto As String

Rem Variables Adicionales para Control de Anterior/Siguiente
Dim dblFactorAncho As Double
Dim intIndiceGrilla As Integer
Dim intIndiceVisible As Integer
Dim intIndiceOculto As Integer
Dim bolRetroceso As Boolean


'Propiedad Numero de Pedido
Public Property Let NroPedido(ByVal value As String)
    strNroPedido = value
End Property

Public Property Get NroPedido() As String
    NroPedido = strNroPedido
End Property

'Propiedad Codigo de Producto
Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property



Private Sub listarDetalle()
    Screen.MousePointer = vbHourglass
    
    ModMilano.visualizarDetalleConsolidadoPedido dbgDetalle(0), strNroPedido, strCodProducto
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub dbgDetalle_OnCustomDrawCell(Index As Integer, ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
'    Select Case Index
'        Case 0
'
'    End Select
    
    Select Case UCase(Column.FieldName)
        Case "CANTTOTAL", "CANTIDADTOTAL", "CANTIDAD"
            Text = Format(Text, "#0.00")
        Case "ESTADO"
            If UCase(Text) = "ANULADO" Then
                Color = vbRed
                FontColor = vbWhite
            Else
                Color = vbWhite
                FontColor = vbBlue
            End If
            
            Font.Bold = True
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    listarDetalle
    
    'Para Control de Anterior/Siguiente
    tmrTemporizador.Enabled = False
    tmrTemporizador.Interval = 0
    tlbDetalle.Tools("Anterior").Enabled = False
    tlbDetalle.Tools("Siguiente").Enabled = True
    
    intIndiceGrilla = 0
    intIndiceVisible = 0
    intIndiceOculto = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgDetalle(0).Dataset.Close
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgDetalle(0).Move 0, dbgDetalle(0).top, Me.ScaleWidth, dbgDetalle(0).Height
    dbgDetalle(1).Move dbgDetalle(0).Width, dbgDetalle(0).top, dbgDetalle(0).Width, dbgDetalle(0).Height
    dbgDetalle(2).Move dbgDetalle(0).Width, dbgDetalle(0).top, dbgDetalle(0).Width, dbgDetalle(0).Height
    
    dblFactorAncho = dbgDetalle(intIndiceGrilla).Width / 40
End Sub

Private Sub tlbDetalle_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "Salir"
            Unload Me
        Case "Anterior"
            If tmrTemporizador.Enabled Then
                Exit Sub
            End If
            
            If intIndiceOculto = 0 Then
                Exit Sub
            End If
        
            bolRetroceso = True
            
            intIndiceVisible = intIndiceGrilla
            
            intIndiceGrilla = intIndiceGrilla - 1
            
            intIndiceOculto = intIndiceGrilla
            
            tmrTemporizador.Interval = 0
            tmrTemporizador.Enabled = True
            
            tmrTemporizador_Timer
        Case "Siguiente"
            If tmrTemporizador.Enabled Then
                Exit Sub
            End If
            
            If intIndiceOculto = (dbgDetalle.Count - 1) Then
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            
            dbgDetalle(intIndiceOculto + 1).Dataset.Close
            
            Select Case (intIndiceOculto + 1)
                Case 1
                    ModMilano.visualizarOPdeRequerimiento dbgDetalle(intIndiceOculto + 1), Trim(dbgDetalle(intIndiceOculto).Columns.ColumnByFieldName("IDREQUERIMIENTO").value & "")
                Case 2
                    ModMilano.visualizarOPDetalle dbgDetalle(intIndiceOculto + 1), Trim(dbgDetalle(intIndiceOculto).Columns.ColumnByFieldName("IDORDENPRODUCCION").value & "")
            End Select
            
            Screen.MousePointer = vbDefault
            
            If dbgDetalle(intIndiceOculto + 1).Dataset.RecordCount = 0 Then
                MsgBox dbgDetalle(intIndiceOculto + 1).Tag & " no disponible(s) para vista.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            bolRetroceso = False
            
            intIndiceVisible = intIndiceGrilla
            
            intIndiceGrilla = intIndiceGrilla + 1
            
            intIndiceOculto = intIndiceGrilla
            
            tmrTemporizador.Interval = 0
            tmrTemporizador.Enabled = True
            
            tmrTemporizador_Timer
    End Select
End Sub

Private Sub tmrTemporizador_Timer()
    If tmrTemporizador.Interval = 40 Then
        tmrTemporizador.Enabled = False
        
        Select Case intIndiceOculto
            Case Is = 0
                tlbDetalle.Tools("Anterior").Enabled = False
                tlbDetalle.Tools("Siguiente").Enabled = True
            Case Is = (dbgDetalle.Count - 1)
                tlbDetalle.Tools("Anterior").Enabled = True
                tlbDetalle.Tools("Siguiente").Enabled = False
            Case Else
                tlbDetalle.Tools("Anterior").Enabled = True
                tlbDetalle.Tools("Siguiente").Enabled = True
        End Select
        
        dbgDetalle(intIndiceOculto).SetFocus
    Else
        tmrTemporizador.Interval = tmrTemporizador.Interval + 1
        
        dbgDetalle(intIndiceVisible).left = dbgDetalle(intIndiceVisible).left + (dblFactorAncho * IIf(bolRetroceso, 1, -1))
        dbgDetalle(intIndiceOculto).left = dbgDetalle(intIndiceOculto).left + (dblFactorAncho * IIf(bolRetroceso, 1, -1))
    End If
End Sub
