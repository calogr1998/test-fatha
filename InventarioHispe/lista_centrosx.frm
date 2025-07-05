VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form lista_centrosx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Centros de Costo"
   ClientHeight    =   6120
   ClientLeft      =   3345
   ClientTop       =   2100
   ClientWidth     =   8955
   Icon            =   "lista_centrosx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   8745
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   300
         Width           =   6660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1695
      TabIndex        =   1
      Top             =   975
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   255
      TabIndex        =   0
      Top             =   975
      Visible         =   0   'False
      Width           =   1455
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4605
      Left            =   75
      OleObjectBlob   =   "lista_centrosx.frx":000C
      TabIndex        =   5
      Top             =   1035
      Width           =   8790
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "lista_centrosx.frx":2E12
      ToolBars        =   "lista_centrosx.frx":933E
   End
End
Attribute VB_Name = "lista_centrosx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rstconsulta As New ADODB.Recordset
Dim wgraba          As String
Dim wtippro         As String

Private Sub dxDBGrid1_OnDblClick()
'Me.MousePointer = vbhourglass
'sw_nuevo_doc = False
'frmcentros.Show 1

Me.MousePointer = 0
'**********
wcodcosto = dxDBGrid1.Columns.ColumnByFieldName("f3costo").value
wunicosto = dxDBGrid1.Columns.ColumnByFieldName("f3abrev").value
    
    txtBusqueda.Text = ""
    Me.Hide
End Sub

Private Sub Form_Activate()
    
dxDBGrid1.Dataset.ADODataset.Requery
wgraba = "0"
dxDBGrid1.Filter.FilterActive = False
dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
End Sub

Private Sub Form_Load()
Dim csql    As String
Dim i       As Integer

Me.MousePointer = vbHourglass
Me.left = 1300
Me.top = 980

With dxDBGrid1
    .Dataset.ADODataset.ConnectionString = cnn_dbbancos
End With
FILL

Me.MousePointer = vbDefault
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim csql As String

Select Case Tool.Id
    Case "ID_Nuevo":
        Me.MousePointer = vbHourglass
        sw_nuevo_doc = True
        frmcentros.Show 1
        Me.MousePointer = 0
    Case "ID_Imprimir"
        Me.MousePointer = vbHourglass
        RptCentroSxFecha.Show 1
        Me.MousePointer = 0
    Case "ID_Salir"
        Unload Me
End Select
    
End Sub

Private Sub FILL()
Dim csql As String

csql = "SELECT F3COSTO,F3ABREV,F3DESCRIP,PO, F3FECGRA From CENTROS where f3costo<>'999' ORDER BY F3COSTO DESC;"
'csql = "select CENTROS.F3COSTO,CENTROS.F3ABREV,CENTROS.F3DESCRIP,EF2CLIENTES.F2NOMCLI from CENTROS LEFT JOIN EF2CLIENTES ON CENTROS.F3CODCLI=EF2CLIENTES.F2CODCLI order by F3COSTO DESC"

dxDBGrid1.Dataset.Active = False
dxDBGrid1.Dataset.ADODataset.CommandText = csql
dxDBGrid1.Dataset.Active = True
dxDBGrid1.KeyField = "F3COSTO"
End Sub

Private Sub txtbusqueda_Change()
dxDBGrid1.Dataset.Filtered = True
dxDBGrid1.Dataset.Filter = "F3COSTO LIKE '*" & txtBusqueda.Text & "*' " & _
                        "OR " & " F3DESCRIP LIKE '*" & txtBusqueda.Text & "*' " & _
                        "OR " & " F3ABREV LIKE '*" & txtBusqueda.Text & "*' OR " & " PO LIKE '*" & txtBusqueda.Text & "*' "


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
    If Len(Trim(txtBusqueda.Text)) > 0 Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "F3COSTO LIKE '*" & txtBusqueda.Text & "*' " & _
                        "OR " & " F3DESCRIP LIKE '*" & txtBusqueda.Text & "*' " & _
                        "OR " & " F3ABREV LIKE '*" & txtBusqueda.Text & "*' "
'                        "OR " & " F2NOMCLI LIKE '*" & txtbusqueda.Text & "*' "

    Else
        dxDBGrid1.Dataset.Filtered = False
    End If
End If
End Sub




