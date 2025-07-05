VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_clientes 
   Caption         =   "Ayuda de Clientes"
   ClientHeight    =   5100
   ClientLeft      =   3480
   ClientTop       =   2865
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7635
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   7410
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   360
         Width           =   5400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   315
         TabIndex        =   3
         Top             =   405
         Width           =   735
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3735
      Left            =   45
      OleObjectBlob   =   "ayuda_clientes.frx":058A
      TabIndex        =   1
      Top             =   990
      Width           =   7410
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ayuda_clientes.frx":4E99
      ToolBars        =   "ayuda_clientes.frx":B3C5
   End
End
Attribute VB_Name = "ayuda_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql        As String

Private Sub CABECERA()

    With dxDBGrid1
        .Columns(0).Caption = "Codigo": .Columns(0).Width = 40: .Columns(0).DisableEditor = True
        .Columns(1).Caption = "Nombre": .Columns(1).Width = 160: .Columns(1).DisableEditor = True
        .Columns(2).Caption = "R.U.C.": .Columns(2).Width = 70: .Columns(2).DisableEditor = True
        .Columns(3).Caption = "Direccion": .Columns(3).Width = 30: .Columns(3).DisableEditor = True: .Columns(3).Visible = False
        .Columns(4).Caption = "F.Pago": .Columns(4).Width = 45: .Columns(4).DisableEditor = True: .Columns(4).Visible = False
        .Columns(5).Caption = "Cotiza": .Columns(5).Width = 50: .Columns(5).DisableEditor = True: .Columns(5).Visible = False
        .Columns(6).Caption = "Cotiza Obs": .Columns(6).Width = 40: .Columns(6).DisableEditor = True: .Columns(6).Visible = False
        .Columns(7).Caption = "Tipo Doc": .Columns(7).Width = 30: .Columns(7).DisableEditor = True: .Columns(7).Visible = False
        .Columns(8).Caption = "D.N.I": .Columns(8).Width = 70: .Columns(8).DisableEditor = True: .Columns(8).Visible = False
        .Columns(9).Caption = "Contacto": .Columns(9).Width = 70: .Columns(9).DisableEditor = True: .Columns(9).Visible = False
        .Columns(10).Caption = "Extranjero": .Columns(10).Width = 70: .Columns(10).DisableEditor = True: .Columns(10).Visible = False
        '.Columns(1).Font.
        'If gtipodocu = "P" Then
        '    .Columns(5).Visible = True
        '    .Columns(6).Visible = True
        'Else
'            .Columns(5).Visible = False
'            .Columns(6).Visible = False
'            .Columns(7).Visible = False
'            .Columns(8).Visible = False
        'End If

    End With
    
End Sub

Private Sub dxDBGrid1_OnDblClick()

    wcodcliprov = dxDBGrid1.Columns.ColumnByFieldName("f2codcli").Value & ""
    wnomcliprov = dxDBGrid1.Columns.ColumnByFieldName("f2nomcli").Value & ""
    wcodpag = dxDBGrid1.Columns.ColumnByFieldName("f2codcli").Value & ""
    wnompag = dxDBGrid1.Columns.ColumnByFieldName("f2nomcli").Value & ""
    wruccli = dxDBGrid1.Columns.ColumnByFieldName("f2newruc").Value & ""
    WDIRCLI = dxDBGrid1.Columns.ColumnByFieldName("f2dircli").Value & ""
    wforpag = "" & dxDBGrid1.Columns.ColumnByFieldName("f2forpag").Value & ""
    wtipocli = dxDBGrid1.Columns.ColumnByFieldName("F2TIPDOC").Value & ""
    wcontacto = dxDBGrid1.Columns.ColumnByFieldName("F2CONTACTO").Value & ""
    wdocidentidad = dxDBGrid1.Columns.ColumnByFieldName("F2DOCIDENTIDAD").Value & ""
    wcliext = dxDBGrid1.Columns.ColumnByFieldName("F2TIPOCLI").Value & ""
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

    If sw_escuela = True Then
        csql = "select f2codcli,f2nomcli,f2newruc,f2dircli,f2forpag,f2cotiza,F2COTIZA_OBS,F2TIPDOC,F2DOCIDENTIDAD,F2CONTACTO,F2TIPOCLI FROM EF2CLIENTES WHERE F2ESCUELA='*' ORDER BY F2CODCLI"
    Else
        csql = "select f2codcli,f2nomcli,f2newruc,f2dircli,f2forpag,f2cotiza,F2COTIZA_OBS,F2TIPDOC,F2DOCIDENTIDAD,F2CONTACTO,F2TIPOCLI from EF2CLIENTES order by F2CODCLI"
    End If
    
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "f2codcli"
'    CABECERA

'    dxDBGrid1.Columns(0).Color = &HC0FFFF
'    dxDBGrid1.Columns(1).Color = &HC0FFFF
'    dxDBGrid1.Columns(2).Color = &HC0FFFF
'    dxDBGrid1.Columns(3).Color = &HC0FFFF
'    dxDBGrid1.Columns(4).Color = &HC0FFFF
'    dxDBGrid1.Columns(5).Color = &HC0FFFF
'    dxDBGrid1.Columns(6).Color = &HC0FFFF
'    dxDBGrid1.Columns(7).Color = &HC0FFFF
'    dxDBGrid1.Columns(8).Color = &HC0FFFF
'    dxDBGrid1.Columns(9).Color = &HC0FFFF
       
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            wcodcliprov = ""
            wnomcliprov = ""
            wruccli = ""
            WDIRCLI = ""
            wforpag = ""
            wtipocli = ""
            wcontacto = ""
            wdocidentidad = ""
            wcliext = ""
            sw_limpia = True
            txtBusqueda.Text = ""
            sw_limpia = False
            Me.Hide
'        Case vbKeyInsert:
'            sw_load_mant = True
'            sw_nuevo_mant = True
'            addCliFac = False
'            If gtipodocu = "F" Then addCliFac = True
'            mant_clientes.Show 1
'            addCliFac = False
'            dxDBGrid1.Dataset.Refresh
'            Me.Hide

    End Select
    
    
End Sub

Private Sub Form_Activate()
    
txtBusqueda.Text = ""
txtBusqueda.SetFocus

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyInsert Then
'        sw_load_mant = True
'        sw_nuevo_mant = True
'        mant_clientes.Show 1
'        dxDBGrid1.Dataset.Refresh
'    End If
End Sub

Private Sub Form_Load()
    If Cnn_Mov.State = adStateOpen Then Cnn_Mov.Close
    Cnn_Mov.ConnectionString = cnn_dbbancos
    Cnn_Mov.Open cconexion
            
    With dxDBGrid1
'        .DefaultFields = True
        .Dataset.ADODataset.ConnectionString = Cnn_Mov
    End With
    FILL
    dxDBGrid1.Filter.FilterActive = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
    Cnn_Mov.Close

End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.ID
    Case "ID_Nuevo"
'        Screen.MousePointer = vbhourglass
        sw_load_mant = True
        sw_nuevo_mant = True
        addCliFac = False
'        Screen.MousePointer = 0
        mant_clientes.Show 1
        addCliFac = False
        dxDBGrid1.Dataset.Refresh
        Unload Me
    Case "ID_Salir"
        Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "f2codcli LIKE '*" & txtBusqueda.Text & "*' OR " & " f2nomcli LIKE '*" & txtBusqueda.Text & "*' OR f2newruc LIKE '*" & txtBusqueda.Text & "*' "
    
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

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    dxDBGrid1.SetFocus
''        If Len(Trim(txtbusqueda.Text)) > 0 Then
''            dxDBGrid1.Dataset.Filtered = True
''            dxDBGrid1.Dataset.Filter = "f2codcli LIKE '*" & txtbusqueda.Text & "*' OR " & " f2nomcli LIKE '*" & txtbusqueda.Text & "*' OR f2newruc LIKE '*" & txtbusqueda.Text & "*' "
''        Else
''            dxDBGrid1.Dataset.Filtered = False
''        End If
    End If
    
End Sub




