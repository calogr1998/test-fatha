VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_orden_trab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Partida de Presupuesto"
   ClientHeight    =   8625
   ClientLeft      =   825
   ClientTop       =   1395
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12435
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   8970
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   7440
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
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ayuda_orden_trab.frx":0000
      ToolBars        =   "ayuda_orden_trab.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   7380
      Left            =   120
      OleObjectBlob   =   "ayuda_orden_trab.frx":65FA
      TabIndex        =   3
      Top             =   945
      Width           =   12090
   End
End
Attribute VB_Name = "ayuda_orden_trab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String
Dim sw_limpia   As Boolean

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
   Select Case UCase(Column.FieldName)
    Case "CANTIDAD", "PRECIO", "AUTORIZADO"
        Text = Format(Text, "###,###,##0.00")
    End Select
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wnumordentrab = dxDBGrid1.Columns.ColumnByFieldName("numorden").value
    wobservacion = dxDBGrid1.Columns.ColumnByFieldName("observacion").value
    sw_limpia = True
    txtBusqueda.Text = ""
    sw_limpia = False
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
    Dim obra_s10 As String
    Dim obra As String
    If wllamada = 0 Then
        If Len(Trim(vale_salida.txtccosto.Text)) > 0 Then
        
        obra = vale_salida.txtccosto.Text
        csql = "Select codpartida as numorden, Descripcion as OBSERVACION FROM dbo_Partida where CodPresupuesto = '" & obra & "' and PropioPartida = '01' ORDER BY codpartida"
        
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.ADODataset.CommandText = csql
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "numorden"
        Else
            MsgBox "No ha elegido una obra. Verifique.", vbCritical, "Atención"
            Exit Sub
        End If
    ElseIf wllamada = 1 Then
        If Len(Trim(solicitud.txtuupp.Text)) > 0 Then
        obra = solicitud.txtuupp.Text
        'csql = "Select codpartida as numorden, Descripcion as OBSERVACION, codUnidad as MEDIDA,material1 as cantidad, TotalMaterial1 AS PRECIO FROM dbo_Partida where CodPresupuesto = '" & obra & "' and PropioPartida = '01' ORDER BY codpartida"
        
        csql = "SELECT dbo_Partida.CodPartida AS numorden, dbo_Partida.Descripcion AS OBSERVACION, dbo_Partida.CodUnidad AS MEDIDA, dbo_Partida.Material1 AS cantidad, dbo_Partida.TotalMaterial1 AS PRECIO, Consulta43.SumaDeMETRADO AS autorizado "
        csql = csql & "FROM dbo_Partida INNER JOIN (SELECT PRESUPUESTO, PARTIDA, Sum(METRADO) AS SumaDeMETRADO FROM CRONOGRAMA GROUP BY PRESUPUESTO, PARTIDA) as Consulta43 ON (dbo_Partida.CodPartida = Consulta43.PARTIDA) AND (dbo_Partida.CodPresupuesto = Consulta43.PRESUPUESTO) "
        csql = csql & "where CodPresupuesto = '" & obra & "' AND dbo_Partida.PropioPartida='01' "
        csql = csql & "ORDER BY dbo_Partida.CodPartida"
        
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.ADODataset.CommandText = csql
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "numorden"
        Else
            MsgBox "No ha elegido una obra. Verifique.", vbCritical, "Atención"
            Exit Sub
        End If
    End If
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
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 3600
    Me.top = 1050
    
    sw_limpia = False
        
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
        Case "ID_Salir":
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()

    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "numorden LIKE '*" & txtBusqueda.Text & "*' OR " & " observacion LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
    
End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "numorden LIKE '*" & txtBusqueda.Text & "*' OR " & " observacion LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub





