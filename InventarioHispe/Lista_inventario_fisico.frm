VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Lista_inventario_fisico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Listado de Inventarios :::"
   ClientHeight    =   3090
   ClientLeft      =   825
   ClientTop       =   2715
   ClientWidth     =   5070
   Icon            =   "Lista_inventario_fisico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2205
      Left            =   90
      OleObjectBlob   =   "Lista_inventario_fisico.frx":000C
      TabIndex        =   0
      Top             =   255
      Width           =   4800
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "Lista_inventario_fisico.frx":0C87
      ToolBars        =   "Lista_inventario_fisico.frx":71B3
   End
End
Attribute VB_Name = "Lista_inventario_fisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim col As TrueOleDBGrid70.Column
'Dim cols As TrueOleDBGrid70.Columns

Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
wultinv = dxDBGrid1.Columns(1).Value
wcod_alm = dxDBGrid1.Columns(0).Value
InventarioFisico.Actualiza_Grid
InventarioFisico.Txtcodalm.Text = wcod_alm
InventarioFisico.llenar_inv
InventarioFisico.Show 1
End Sub

Private Sub Form_Activate()
    
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    dxDBGrid1.OptionEnabled = 1
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    sw_mant_ayuda = False
End Sub

Private Sub Form_Load()
    Me.MousePointer = 11
    Me.Left = 1600
    Me.Top = 1050
    sw_limpia = False
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.MousePointer = 1
End Sub
Private Sub FILL()
    Dim sql As String
sql = "Select F2CODALM, F4FECTOM, F4CIERRE FROM H4TOMAINV"

dxDBGrid1.Dataset.Active = False

dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
dxDBGrid1.Dataset.ADODataset.CommandType = cmdText
dxDBGrid1.Dataset.ADODataset.CommandText = sql

dxDBGrid1.DefaultFields = True
dxDBGrid1.Dataset.Active = True

dxDBGrid1.Columns(0).Width = 50
dxDBGrid1.Columns(1).Width = 190
dxDBGrid1.Columns(2).Width = 190
dxDBGrid1.Columns(0).Caption = "ALMACEN"
dxDBGrid1.Columns(1).Caption = "FECHA DE TOMA"
dxDBGrid1.Columns(2).Caption = "FECHA DE CIERRE"

dxDBGrid1.KeyField = "F2CODALM"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.ID
    Case "ID_Nuevo"
        sw_nuevo_doc = True
        sw_mant_ayuda = False
        InventarioFisico.Show 1
    Case "ID_Imprimir":
        With Acr_Marcas
            .DataControl1.ConnectionString = cnn_dbbancos
            .DataControl1.Source = "Select * From ef2marcas order by f2codmar"
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .lblempresa.Caption = wnomcia
            .Show 1
        End With
    Case "ID_Salir"
        Unload Me
End Select
End Sub

Private Sub tdbmarcas_DblClick()
sw_nuevo_doc = False
mant_marcas.Show 1
End Sub

Private Sub txtbusqueda_Change()
   dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "f2codmar LIKE '*" & txtbusqueda.Text & "*' OR " & " f2desmar LIKE '*" & txtbusqueda.Text & "*' "
    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "f2codmar LIKE '*" & txtbusqueda.Text & "*' OR " & " f2desmar LIKE '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub
