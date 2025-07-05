VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form FrmLisAutorizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Autorizaciones"
   ClientHeight    =   6045
   ClientLeft      =   840
   ClientTop       =   1350
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9600
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9450
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   8160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   120
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
      ToolsCount      =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "FrmLisAutorizaciones.frx":0000
      ToolBars        =   "FrmLisAutorizaciones.frx":9751
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4935
      Left            =   60
      OleObjectBlob   =   "FrmLisAutorizaciones.frx":9942
      TabIndex        =   3
      Top             =   1020
      Width           =   9435
   End
End
Attribute VB_Name = "FrmLisAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dxDBGrid1_OnDblClick()
If dxDBGrid1.Dataset.RecordCount > 0 Then
    sw_e_ordenpago = True
    frmOrdPago.Show 1
End If
End Sub

Private Sub Form_Activate()
    'dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    'dxDBGrid1.Dataset.Close
    'dxDBGrid1.Dataset.Open
    txtBusqueda.SetFocus
End Sub

Private Sub Form_Load()
Me.MousePointer = vbHourglass
FILL
Me.MousePointer = vbDefault
End Sub

Private Sub FILL()
csql = "SELECT IF4ORDEN_PAGO.ORDEN, IF4ORDEN_PAGO.FECHA, IF4ORDEN_PAGO.USUARIO, IF4ORDEN_PAGO.MONEDA, " & _
       "IF4ORDEN_PAGO.IMPORTE, IF4ORDEN_PAGO.IDOP, IF4ORDEN_PAGO.correladoc, IF4ORDEN_PAGO.correlaanticipo, " & _
       "IF4ORDEN_PAGO.observacion, IF4ORDEN_PAGO.ESTADO, IF4ORDEN_PAGO.EST_AUT, SALDOSOPS.SALDOS " & _
       "FROM IF4ORDEN_PAGO INNER JOIN (SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4MONTO, SUM(IF4ORDEN_PAGO.IMPORTE) AS SUMIMPORTES, " & _
       "IF4ORDEN.F4MONTO - SUM(IF4ORDEN_PAGO.IMPORTE) AS SALDOS FROM IF4ORDEN LEFT JOIN IF4ORDEN_PAGO ON IF4ORDEN.F4NUMORD = IF4ORDEN_PAGO.ORDEN " & _
       "WHERE IF4ORDEN_PAGO.ESTADO='1' GROUP BY IF4ORDEN.F4NUMORD, IF4ORDEN.F4MONTO HAVING (IF4ORDEN.F4MONTO - SUM(IF4ORDEN_PAGO.IMPORTE)) <> 0) as SALDOSOPS " & _
       "ON IF4ORDEN_PAGO.ORDEN = SALDOSOPS.F4NUMORD Where IF4ORDEN_PAGO.ESTADO = '1'"
With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
'        .KeyField = "IDOP"
End With
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
        Case "ID_Nuevo"
     
       Case "ID_Filtrar"
            If Tool.State = ssChecked Then
                dxDBGrid1.Filter.FilterActive = True
            Else
                dxDBGrid1.Filter.FilterActive = False
            End If
        Case "ID_Agrupar"
            If Tool.State = ssChecked Then
                dxDBGrid1.Options.Set (egoShowGroupPanel)
            Else
                dxDBGrid1.Options.Unset (egoShowGroupPanel)
            End If
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
dxDBGrid1.Dataset.Filtered = True
dxDBGrid1.Dataset.Filter = "ORDEN LIKE '*" & txtBusqueda.Text & "*' OR " & " FECHA LIKE '*" & txtBusqueda.Text & "*' OR USUARIO LIKE '*" & txtBusqueda.Text & "*' "
    
If Len(Trim(txtBusqueda.Text)) = 0 Then
        dxDBGrid1.Dataset.Filtered = False
End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dxDBGrid1.SetFocus
End If

End Sub
