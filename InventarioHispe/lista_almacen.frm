VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form lista_almacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Almacenes"
   ClientHeight    =   5100
   ClientLeft      =   1740
   ClientTop       =   2040
   ClientWidth     =   6975
   Icon            =   "lista_almacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6975
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   6750
      Begin VB.TextBox txtbusqueda 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F3F2&
         ForeColor       =   &H00700112&
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   255
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc adoctasctes 
      Height          =   330
      Left            =   1980
      Top             =   6795
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoctasctes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3795
      Left            =   60
      OleObjectBlob   =   "lista_almacen.frx":058A
      TabIndex        =   3
      Top             =   795
      Width           =   6780
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   900
      Top             =   4290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "lista_almacen.frx":1BBC
      ToolBars        =   "lista_almacen.frx":8D99
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   195
      Left            =   3465
      TabIndex        =   4
      Top             =   3630
      Width           =   810
   End
End
Attribute VB_Name = "lista_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim col As TrueOleDBGrid70.Column
'Dim cols As TrueOleDBGrid70.Columns

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub dxDBGrid1_OnDblClick()
    If dxDBGrid1.Columns(0).value <> "" Then
        sw_nuevo_doc = False
        wcod_alm = dxDBGrid1.Columns.ColumnByFieldName("f2codalm").value
        mant_almacen.Show 1
    Else
        MsgBox "Debera de elegir un almacen", vbInformation, "Aviso"
        txtbusqueda_Change
        txtBusqueda.Text = ""
        txtBusqueda.SetFocus
    End If
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
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    
    Me.left = 1600
    Me.top = 1050


    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With dxDBGrid1
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus
            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
        End With
        
        listarAlmacenSql
    Else
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
        FILL
    End If
    
    
    Me.MousePointer = vbDefault
End Sub

Public Sub listarAlmacenSql()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "F2CODALM, "
    strSQL = strSQL & "F2NOMALM "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "MAESTROS.EF2ALMACENES"
    
    With dxDBGrid1
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F2CODALM"
        
        .m.FullExpand
    End With
End Sub

Private Sub FILL()
    Dim csql As String
    csql = "Select F2CODALM,F2NOMALM From EF2ALMACENES Order By F2CODALM"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODALM"
    CABECERA

    dxDBGrid1.Columns(0).Color = &HC0FFFF
    dxDBGrid1.Columns(1).Color = &HC0FFFF
       
End Sub
Private Sub CABECERA()
    With dxDBGrid1
        .Columns(0).Caption = "Código": .Columns(0).Width = 50: .Columns(0).DisableEditor = True
        .Columns(1).Caption = "Descripción.": .Columns(1).Width = 250: .Columns(1).DisableEditor = True
    End With
End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            mant_almacen.Show 1
        Case "ID_Imprimir":
            With Acr_Almacen
                .DataControl1.ConnectionString = cnn_dbbancos
                .DataControl1.Source = "select * from ef2almacenes order by f2codalm"
                .fldFecha.Text = Format(Date, "DD/MM/YYYY")
                .lblempresa.Caption = wnomcia
                .Show 1
            End With
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub tdblista_DblClick()
'
'    sw_nuevo_doc = False
'    mant_almacen.Show 1
    
End Sub

Private Sub txtbusqueda_Change()

  dxDBGrid1.Dataset.Filtered = True
  dxDBGrid1.Dataset.Filter = "F2CODALM LIKE '*" & txtBusqueda.Text & "*' OR " & " F2NOMALM LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F2CODALM LIKE '*" & txtBusqueda.Text & "*' OR " & " F2NOMALM LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If

End Sub
