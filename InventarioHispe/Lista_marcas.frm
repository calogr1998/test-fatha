VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form Lista_marcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marcas"
   ClientHeight    =   5610
   ClientLeft      =   825
   ClientTop       =   2715
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4770
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   3600
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
      Left            =   120
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "Lista_marcas.frx":0000
      ToolBars        =   "Lista_marcas.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4380
      Left            =   120
      OleObjectBlob   =   "Lista_marcas.frx":668D
      TabIndex        =   3
      Top             =   990
      Width           =   4770
   End
End
Attribute VB_Name = "Lista_marcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim col As TrueOleDBGrid70.Column
'Dim cols As TrueOleDBGrid70.Columns

Private Sub Checkagrupar_Click()
    If Checkagrupar.value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
    sw_nuevo_doc = False
    mant_marcas.Show 1
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
    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1050
    sw_limpia = False
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.MousePointer = vbDefault
End Sub
Private Sub FILL()
    csql = "Select f2codmar,F2DESMAR From ef2marcas order by f2codmar"
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "f2codmar"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
        sw_nuevo_doc = True
        sw_mant_ayuda = False
        mant_marcas.Show 1
    Case "ID_Imprimir":
        With Acr_Marcas
            .DataControl1.ConnectionString = cnn_dbbancos
            .DataControl1.Source = "Select * From ef2marcas order by f2codmar"
            .fldFecha.Text = Format(Date, "DD/MM/YYYY")
            .lblEmpresa.Caption = wnomcia
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
    dxDBGrid1.Dataset.Filter = "f2codmar LIKE '*" & txtBusqueda.Text & "*' OR " & " f2desmar LIKE '*" & txtBusqueda.Text & "*' "
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "f2codmar LIKE '*" & txtBusqueda.Text & "*' OR " & " f2desmar LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub
