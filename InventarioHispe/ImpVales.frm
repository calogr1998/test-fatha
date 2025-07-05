VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ImpVales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación de Datos"
   ClientHeight    =   7455
   ClientLeft      =   600
   ClientTop       =   1935
   ClientWidth     =   12285
   Icon            =   "ImpVales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12285
   Begin Threed.SSPanel SSPanel2 
      Height          =   1020
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   12075
      _Version        =   65536
      _ExtentX        =   21299
      _ExtentY        =   1799
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox Txtcodcli 
         BackColor       =   &H00F4F3F2&
         ForeColor       =   &H00700112&
         Height          =   285
         Left            =   915
         MaxLength       =   4
         TabIndex        =   0
         Top             =   135
         Width           =   960
      End
      Begin VB.ComboBox Cmbtipref 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F3F2&
         ForeColor       =   &H00700112&
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   555
         Width           =   2805
      End
      Begin Threed.SSPanel pnlnomcli 
         Height          =   285
         Left            =   1905
         TabIndex        =   5
         Top             =   135
         Width           =   6360
         _Version        =   65536
         _ExtentX        =   11218
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label lblcliente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblrefere 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Importar Datos de un(a) ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   570
         Width           =   1890
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6225
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Width           =   12105
      _Version        =   65536
      _ExtentX        =   21352
      _ExtentY        =   10980
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin DXDBGRIDLibCtl.dxDBGrid dxImpDatos 
         Height          =   3090
         Left            =   0
         OleObjectBlob   =   "ImpVales.frx":058A
         TabIndex        =   8
         Top             =   15
         Width           =   12075
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxImpDatosDet 
         Height          =   3090
         Left            =   15
         OleObjectBlob   =   "ImpVales.frx":11FE
         TabIndex        =   7
         Top             =   3120
         Width           =   12075
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "ImpVales.frx":5B98
      ToolBars        =   "ImpVales.frx":B430
   End
End
Attribute VB_Name = "ImpVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnTemp  As New ADODB.Connection
Dim sw_conectar As Boolean
Dim Ot_tipo     As String
Dim sw_primero As Boolean

Public Sub CONECTAR()

    With dxImpDatos
        .DefaultFields = True
        .Dataset.ADODataset.ConnectionString = cnn_form
    End With
    
End Sub

Public Sub CABECERA()

            With dxImpDatos
                .Columns(0).Caption = "Alm.": .Columns(0).Width = 30: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Nº Vale": .Columns(1).Width = 60: .Columns(1).DisableEditor = True: .Columns(1).Color = &HC0FFFF
                .Columns(2).Caption = "RUC": .Columns(2).Width = 60: .Columns(2).DisableEditor = True: .Columns(2).Color = &HC0FFFF: .Columns(2).Visible = False
                .Columns(3).Caption = "Proveedor": .Columns(3).Width = 180: .Columns(3).DisableEditor = True
                .Columns(4).Caption = "Fecha": .Columns(4).Width = 80: .Columns(4).DisableEditor = True
                .Columns(5).Caption = "Moneda": .Columns(5).Width = 40: .Columns(5).DisableEditor = True
                .Columns(6).Caption = "Soles": .Columns(6).Width = 60: .Columns(6).DisableEditor = True: .Columns(6).DecimalPlaces = 2
                .Columns(7).Caption = "Dólares": .Columns(7).Width = 60: .Columns(7).DisableEditor = True: .Columns(7).DecimalPlaces = 2
                .Columns(8).Caption = "Marcar": .Columns(8).Width = 50: .Columns(8).DisableEditor = False: .Columns(8).Visible = True: .Columns(8).ColumnType = gedCheckEdit
                .Columns(9).Caption = "Reg.Compras": .Columns(9).Width = 50: .Columns(9).DisableEditor = True: .Columns(9).Visible = False
            End With
 
End Sub

Private Sub DETALLE()

    Select Case Ot_tipo
         Case "R"
            With dxDBGrid1
                .Columns(0).Caption = "Nº Serie": .Columns(0).Width = 55: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Nº Documento": .Columns(1).Width = 80: .Columns(1).DisableEditor = True: .Columns(1).Color = &HC0FFFF
                .Columns(2).Caption = "Cliente": .Columns(2).Width = 200: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "F.Emision": .Columns(3).Width = 100: .Columns(3).DisableEditor = True
                .Columns(4).Caption = "Total": .Columns(4).Width = 70: .Columns(4).DisableEditor = True: .Columns(4).DecimalPlaces = 2
                .Columns(5).Caption = "Estado": .Columns(5).Width = 50: .Columns(5).DisableEditor = True
                .Columns(6).Caption = "OBSERVACIONES": .Columns(6).Width = 200: .Columns(6).DisableEditor = True: .Columns(6).Visible = False
            End With
     End Select

End Sub
Private Sub Configurar1()
    '********* esto es para refrescar el grid *********
    'dxImpDatos.Dataset.Close
    'DELETEREC_LOG "TEMP_DOC", cnn_form
    ' enviar un select vacio, por ke sale un error
    'dxImpDatos.Dataset.ADODataset.CommandText = "SELECT * FROM TEMP_DOC ORDER BY F4NUMDOC"
    'dxImpDatos.Dataset.Open
    ' nuevamente para asegurar el refresh,si no siempre deja un(o)s registro(s)
    'dxImpDatos.Dataset.Close
    DELETEREC_LOG "TempOrden", cnn_form
    'MsgBox dxImpDatos.Dataset.ADODataset.ConnectionString
    'dxImpDatos.Dataset.Open
    '**************************************************
    Select Case Ot_tipo
        Case "V"
            'CONECTAR ' falta acondicionar esto 10/09/2005
            'SQL = "select F4SERDOC,F4NUMDOC, F2NOMCLI,F4FECEMI,  F4TOTFAC, F4ESTNUL, F4CHECK FROM TBVENTA_CAB WHERE F4TIPODOCU='90' AND F4ESTNUL='N' AND F2CODCLI='" & Txtcodcli.Text & "' ORDER BY F4NUMDOC"
            SQL1 = "INSERT INTO TEMPORDEN(F4NUMORD,F4ESTNUL,F4FECEMI,F2NOMPROV, F4TIPCAM, F4MONTO, F2CODALM," & _
                    "F4NUMVAL,F4TIPOVALE,F2CODPROV,F4NUMDOC,F4FECVAL) IN '" & wrutatemp & "\TEMPLUS.MDB' SELECT 0,' ','23/02/2007',EF2PROVEEDORES.F2NOMPROV,0,0,IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4TIPOVALE, IF4VALES.F2CODPROV, '   ', IF4VALES.F4FECVAL FROM EF2PROVEEDORES " & _
                "INNER JOIN IF4VALES ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV WHERE F4REGCOM Is Null;"
            
            sql = "SELECT F2CODALM, F4NUMVAL, F4TIPOVALE, F4FECEMI,F2CODPROV, F2NOMPROV, F4NUMDOC, F4FECVAL, F4CHECK FROM TEMPORDEN ORDER BY F4NUMDOC"
        Case "C"
            'CONECTAR ' falta acondicionar esto 10/09/2005
            'SQL = "select F4SERDOC,F4NUMDOC, F2NOMCLI,F4FECEMI,  F4TOTFAC, F4ESTNUL, F4CHECK FROM TBVENTA_CAB WHERE F4TIPODOCU='90' AND F4ESTNUL='N' AND F2CODCLI='" & Txtcodcli.Text & "' ORDER BY F4NUMDOC"
            SQL1 = "  INSERT INTO TEMPORDEN(F4NUMORD,F4ESTNUL,F4FECEMI,F2NOMPROV, F4TIPCAM, F4MONTO, F2CODALM," & _
                    " F4NUMVAL,F4TIPOVALE,F2CODPROV,F4NUMDOC,F4FECVAL) IN '" & wrutatemp & "\TEMPLUS.MDB' SELECT IF4ORDEN.F4NUMORD, " & _
                    " IF4ORDEN.F4ESTNUL, IF4ORDEN.F4FECEMI, EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4TIPCAM, IF4ORDEN.F4MONTO, " & _
                    " ' ',' ', ' ',' ',' ','23/01/2007' FROM EF2PROVEEDORES INNER JOIN IF4ORDEN ON EF2PROVEEDORES.F2NEWRUC = " & _
                    " IF4ORDEN.F4CODPRV WHERE " & _
                    " WHERE (((EF2PROVEEDORES.F2CODPROV) LIKE '" & Txtcodcli.Text & "%') AND ((IF4ORDEN.F4REGCOM) Is Null))"
                    
            SQL1 = "  INSERT INTO TEMPORDEN(F4NUMORD,F4ESTNUL,F4FECEMI,F2NOMPROV, F4TIPCAM, F4MONTO, F2CODALM, "
            SQL1 = SQL1 & "F4NUMVAL,F4TIPOVALE,F2CODPROV,F4NUMDOC,F4FECVAL) IN '" & wrutatemp & "\TEMPLUS.MDB' "
            SQL1 = SQL1 & "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4ESTNUL, IF4ORDEN.F4FECEMI, "
            SQL1 = SQL1 & "EF2PROVEEDORES.F2NOMPROV, IF4ORDEN.F4TIPCAM, IF4ORDEN.F4MONTO, ' ', ' ', ' ', ' ', ' ', "
            SQL1 = SQL1 & "'23/01/2007' "
            SQL1 = SQL1 & "FROM EF2PROVEEDORES INNER JOIN IF4ORDEN "
            SQL1 = SQL1 & "ON EF2PROVEEDORES.F2NEWRUC = IF4ORDEN.F4CODPRV "
            SQL1 = SQL1 & "WHERE (((EF2PROVEEDORES.F2CODPROV) LIKE '" & Txtcodcli.Text & "%') AND ((IF4ORDEN.F4REGCOM) Is Null));"
            
            sql = "SELECT F4NUMORD, F4ESTNUL,F4FECEMI, F2NOMPROV, F4TIPCAM, F4MONTO, F4CHECK FROM TEMPORDEN ORDER BY F4NUMORD"
    End Select
'        SQL1 = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL,IF4VALES.F2CODPROV, IF4VALES.F4REFERE, IF4VALES.F4FECVAL, " & _
'           " IF4VALES.F4MONEDA, Sum(IF3VALES.F3TOTITE) AS SOLES, Sum(IF3VALES.F3TOTDOL) AS DOLARES, IF4VALES.F4CHECK,Sum(IF3VALES.F3SALDOC) AS SALDO " & _
'           "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) WHERE IF4VALES.F2CODPROV = '" & wrucprov & "'" & _
'           "GROUP BY IF4VALES.F2CODALM, IF4VALES.F4NUMVAL,IF4VALES.F2CODPROV, IF4VALES.F4REFERE, IF4VALES.F4CHECK,IF4VALES.F4FECVAL, IF4VALES.F4MONEDA, IF4VALES.F1CODORI " & _
'           "HAVING IF4VALES.F1CODORI='XC0' " & _
'           "ORDER BY IF4VALES.F4FECVAL DESC; "
            'MsgBox SQL
            
            cnn_dbbancos.Execute SQL1
            'AlmacenaQuery_sql SQL1, cnn_dbbancos
        
            dxImpDatos.Dataset.Active = False
            dxImpDatos.Dataset.ADODataset.CommandType = cmdText
            dxImpDatos.Dataset.ADODataset.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Bancowin\Templus.mdb;Persist Security Info=False"
            dxImpDatos.Dataset.ADODataset.CommandText = sql
            dxImpDatos.Dataset.Active = True
            dxImpDatos.Dataset.ADODataset.Requery
            dxImpDatos.KeyField = "F4NUMDOC"
            'CABECERA

    Select Case Ot_tipo
        Case "V"
            dxImpDatos.Columns(0).Caption = "Cod."
            dxImpDatos.Columns(0).Width = 30
            
            dxImpDatos.Columns(1).Caption = "Nº Vale"
            dxImpDatos.Columns(1).Width = 60
            
            dxImpDatos.Columns(2).Caption = "Tip. Vale"
            dxImpDatos.Columns(2).Width = 60
            
            dxImpDatos.Columns(3).Caption = "Fecha"
            dxImpDatos.Columns(3).Width = 70
            
            dxImpDatos.Columns(4).Caption = "Cod. Prov"
            dxImpDatos.Columns(4).Width = 50
            dxImpDatos.Columns(5).Caption = "Proveedor"
            dxImpDatos.Columns(5).Width = 250
            dxImpDatos.Columns(6).Caption = "Nº Doc.": dxImpDatos.Columns(6).Width = 120
            dxImpDatos.Columns(7).Caption = "Fec. Vale": dxImpDatos.Columns(7).Width = 70
            dxImpDatos.Columns(8).Caption = "Marca"
            
            'SQL = "SELECT F2CODALM, F4NUMVAL, F4TIPOVALE, F2CODPROV, F4NUMDOC, F4FECVAL, F4CHECK FROM TEMPORDEN ORDER BY F4NUMDOC"
        Case "C"
            dxImpDatos.Columns(0).Caption = "Nº Orden"
            dxImpDatos.Columns(0).Width = 90
            dxImpDatos.Columns(1).Caption = "Est. Anul."
            dxImpDatos.Columns(1).Width = 50
            dxImpDatos.Columns(2).Caption = "Fec. Emision"
            dxImpDatos.Columns(2).Width = 90
            dxImpDatos.Columns(3).Caption = "Proveedor"
            dxImpDatos.Columns(3).Width = 250
            dxImpDatos.Columns(4).Caption = "Tip. Cambio"
            dxImpDatos.Columns(4).Width = 100
            dxImpDatos.Columns(5).Caption = "Cod."
            dxImpDatos.Columns(5).Width = 90
            dxImpDatos.Columns(6).Caption = "Marcar"
            dxImpDatos.Columns(6).Width = 60
            
            'SQL = "SELECT F4NUMORD, F4ESTNUL,F4FECEMI, F2NOMPROV, F4TIPCAM, F4MONTO, F4CHECK FROM TEMPORDEN ORDER BY F4NUMORD"
    End Select
    dxImpDatos.Columns(0).Color = &HC0FFFF


End Sub

'Private Sub Configurar()
'On Error Resume Next
'    '********* esto es para refrescar el grid *********
'    'dxDBGrid1.Dataset.Close
'   'DELETEREC_LOG "TEMP_DOC", cnTemp
'    ' enviar un select vacio, por ke sale un error
'    'dxImpDatos.Dataset.ADODataset.CommandText = "SELECT * FROM TEMP_DOC ORDER BY F4NUMDOC"
'    'dxImpDatos.Dataset.Open
'    ' nuevamente para asegurar el refresh,si no siempre deja un(o)s registro(s)
'    'dxImpDatos.Dataset.Close
'
'    If sw_primero = True Then
'        sw_primero = False
'    Else
'        cnn_form.Execute "DROP TABLE TEMP_DOC"
''        DELETEREC_LOG "TEMP_DOC", cnTemp
''        DELETEREC_LOG "TEMP_DOC", cnTemp
'        dxImpDatos.Dataset.Refresh
'        dxImpDatos.Dataset.Close
'        dxImpDatos.Dataset.Open
'    End If
'
''    DELETEREC_LOG "TEMP_DOC", cnTemp
'
''    cnTemp.Execute "DROP TABLE TEMP_DOC"
'    'cnTemp.Close
'
'    'dxImpDatos.Dataset.Open
'    '**************************************************
'
'    Select Case Ot_tipo
'        Case "V"
'            'SQL = "INSERT INTO TEMP_DOC IN '" & wrutatemp & "\templus.mdb' SELECT TBVENTA_CAB.F4SERDOC, TBVENTA_CAB.F4NUMDOC, TBVENTA_CAB.F2NOMCLI, TBVENTA_CAB.F4FECEMI, TBVENTA_CAB.F4TOTFAC, TBVENTA_CAB.F4ESTNUL,TBVENTA_CAB.F4CHECK, TBVENTA_CAB.F4ESTFAC " & _
'            '" FROM TBVENTA_CAB INNER JOIN TBVENTA_DET ON (TBVENTA_CAB.F4NUMDOC = TBVENTA_DET.F4NUMDOC) AND (TBVENTA_CAB.F4SERDOC = TBVENTA_DET.F4SERDOC) AND (TBVENTA_CAB.F4TIPODOCU = TBVENTA_DET.F4TIPODOCU)" & _
'            '" WHERE TBVENTA_CAB.F4ESTNUL='N' AND (TBVENTA_CAB.F4ESTFAC='N' Or TBVENTA_CAB.F4ESTFAC Is Null Or TBVENTA_CAB.F4ESTFAC='') AND ((TBVENTA_CAB.F4TIPODOCU)='01') AND ((TBVENTA_CAB.F2CODCLI)='" & Txtcodcli.Text & "')  and (tbventa_det.f3fact<>'0' or tbventa_det.f3fact is null) GROUP BY TBVENTA_CAB.F4SERDOC, TBVENTA_CAB.F4NUMDOC, TBVENTA_CAB.F2NOMCLI, TBVENTA_CAB.F4FECEMI, TBVENTA_CAB.F4TOTFAC, TBVENTA_CAB.F4ESTNUL,TBVENTA_CAB.F4CHECK, TBVENTA_CAB.F4ESTFAC " & _
'            '" ORDER BY TBVENTA_CAB.F4NUMDOC "
'            'cnn_dbbancos.Execute SQL
'            SQL = ""
'
'            'SQL = "SELECT * FROM TEMP_DOC ORDER BY F4NUMDOC"
'            SQL = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL,IF4VALES.F2CODPROV, IF4VALES.F4REFERE, IF4VALES.F4FECVAL, " & _
'                    " IF4VALES.F4MONEDA, Sum(IF3VALES.F3TOTITE) AS SOLES, Sum(IF3VALES.F3TOTDOL) AS DOLARES, IF4VALES.F4CHECK,Sum(IF3VALES.F3SALDOC) AS SALDO INTO TEMP_DOC IN '" & wrutatemp & "templus.mdb' " & _
'                    "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) WHERE IF4VALES.F2CODPROV = '" & wrucprov & "'" & _
'                    "GROUP BY IF4VALES.F2CODALM, IF4VALES.F4NUMVAL,IF4VALES.F2CODPROV, IF4VALES.F4REFERE, IF4VALES.F4CHECK,IF4VALES.F4FECVAL, IF4VALES.F4MONEDA, IF4VALES.F1CODORI " & _
'                    "HAVING IF4VALES.F1CODORI='XC0' " & _
'                    "ORDER BY IF4VALES.F4FECVAL DESC; "
'
'           campo = "F4NUMVAL"
'        Case "C"
'            SQL = ""
'
'            SQL = "SELECT IF4ORDEN.F4NUMORD,IF4ORDEN.F4CODPRV, IF4ORDEN.F4ESTNUL, IF4ORDEN.F4FECEMI, EF2PROVEEDORES.F2NOMPROV, " & _
'            " IF4ORDEN.F4TIPCAM, IF4ORDEN.F4MONTO INTO TEMP_DOC IN '" & wrutatemp & "templus.mdb'" & " FROM " & _
'            " IF4ORDEN INNER JOIN EF2PROVEEDORES ON IF4ORDEN.F4CODPRV = EF2PROVEEDORES.F2NEWRUC; "
'
'            campo = "IF4ORDEN.F4NUMORD"
''           ' MsgBox cnn_dbbancos.ConnectionString
''            cnn_dbbancos.Execute SQL
''
''            'cnn_dbbancos.Execute SQL
''            'Debug.Print SQL
''            dxImpDatos.Dataset.Active = False
''            'MsgBox dxImpDatos.Dataset.ADODataset.ConnectionString
''            dxImpDatos.Dataset.ADODataset.CommandText = "Select * From TEMP_DOC"
''
''            dxImpDatos.Dataset.Active = True
''
''
''            dxImpDatos.KeyField = "IF4ORDEN.F4NUMORD"
'    End Select
'        'MsgBox cnn_dbbancos
'        cnn_dbbancos.Execute SQL
'            dxImpDatos.Dataset.Active = False
'            'MsgBox dxImpDatos.Dataset.ADODataset.ConnectionString
'            dxImpDatos.Dataset.ADODataset.ConnectionString = cnn_form
'            dxImpDatos.Dataset.ADODataset.CommandText = "SELECT * FROM TEMP_DOC"
'            dxImpDatos.Dataset.Active = True
'            dxImpDatos.Dataset.Open
'            dxImpDatos.OptionEnabled = False
''            dxImpDatos.Dataset.DisableControls
'
'
'            dxImpDatos.Dataset.Active = True
''              MsgBox dxImpDatos.Dataset.RecordCount
'            dxImpDatos.KeyField = campo
'            dxImpDatos.Dataset.Refresh
'            dxImpDatos.Dataset.Close
'            dxImpDatos.Dataset.Open
'
'
'
'            If Ot_tipo = "V" Then
'            End If 'CABECERA
'        dxImpDatos.Columns(0).Color = &HC0FFFF
'
'
'End Sub

Private Sub cmbtipref_Click()
    Me.MousePointer = vbHourglass
'            MsgBox cnn_dbbancos.ConnectionString
           ' cnn_dbbancos.Execute "UPDATE IF4VALES SET F4CHECK = FALSE WHERE F4CHECK = TRUE"
            If Trim(Txtcodcli.Text) <> "" Then
                Select Case Cmbtipref.ListIndex
                    Case 1:
                        Ot_tipo = "V"
                        ECS_otipo = "V"
                        Configurar1
                        sw_conectar = True
                    Case 2:
                        Ot_tipo = "C"
                        ECS_otipo = "C"
                        Configurar1
                        sw_conectar = True
                End Select
            Else
                'MsgBox "Debe seleccionar un proveedor", vbInformation + vbDefaultButton1, "Atención"
                Txtcodcli_KeyDown 113, 0
            End If
            
    Me.MousePointer = vbDefault
    If dxImpDatosDet.Dataset.RecordCount > 0 Then
        dxImpDatosDet.Dataset.Close
        DELETEREC_LOG "DATOSDET", cnn_form
        dxImpDatosDet.Dataset.Open
    End If
    
End Sub

Private Sub dxImpDatos_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)

        dxImpDatosDet.Dataset.Close
        'DELETEREC_LOG DBTable3, cnn_form
        If dxImpDatos.Dataset.State = 2 Then dxImpDatos.Dataset.Post
        AdicionaItemImp
    
End Sub


Private Sub Form_Load()
    Me.top = 135
    Me.left = 1080
    Cmbtipref.Clear
            Cmbtipref.AddItem "Ninguna"
            Cmbtipref.AddItem "Vale de Ingreso"
            Cmbtipref.AddItem "Orden de Compra"
    If cnTemp.State = 1 Then cnTemp.Close
    cnTemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\templus.MDB;Persist Security Info=False"
    CONECTAR
    Ot_tipo = "V"
    Cmbtipref.ListIndex = 1
    DBTable3 = "DATOSDET"
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\templus.mdb" & ";Persist Security Info=False"
    cnn_form.Open cconex_form
'    MsgBox cnn_form
    DELETEREC_LOG DBTable3, cnn_form
        sw_primero = True
    'CONECTAR
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cnTemp.State <> 0 Then cnTemp.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Aceptar"
            windex = Cmbtipref.ListIndex
            If dxImpDatosDet.Dataset.RecordCount > 0 Then
'                dxImpDatosDet.Dataset.Edit
'                dxImpDatosDet.Dataset.Post
'                wHayDatosImportacion = True
'                Me.MousePointer = vbhourglass
'                  VerDatos
'                Me.MousePointer = vbdefault
                With Registro_Compras
                    .TxtCodPrv.Text = Txtcodcli.Text
'                    Call .TxtCodPrv_KeyUp(8, 0)
                    Select Case Ot_tipo
                    Case "V"
                        
                        sql = "SELECT IF5PLA.F5CTACON, IF5PLA.F3GASTO, IF5PLA.F5AFECTO, Sum((IF3VALES.F3CANPRO * IF3VALES.F3VALVTA) + IF3VALES.F3IGV) as Importe" & _
                        " FROM IF4VALES INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) " & _
                        " WHERE IF4VALES.F2CODALM in (Select F2CODALM From Temporden in '" & wrutatemp & "\TEMPLUS.MDB' WHERE " & _
                        " F4Check = TRUE) AND IF4VALES.F4NUMVAL  in (Select F4NUMVAL From Temporden IN '" & wrutatemp & "\TEMPLUS.MDB'" & _
                        " Where F4Check = TRUE) group by IF5PLA.F5CTACON, IF5PLA.F3GASTO, IF5PLA.F5AFECTO;"
                    
                    'rs.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    'rs.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    Case "C"
                        sql = " SELECT IF5PLA.F5CTACON, IF5PLA.F3GASTO, IF3ORDEN.F5AFECTO, sum(IF3ORDEN.F3CANPRO * IF3ORDEN.F3PRECOS) as Importe" & _
                        " FROM IF3ORDEN, IF5PLA Where IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO " & _
                        " And IF3ORDEN.F4NUMORD in (Select F4NUMORD From Temporden in '" & wrutatemp & "\TEMPLUS.MDB' Where F4Check = TRUE) group by IF5PLA.F5CTACON, IF5PLA.F3GASTO, IF3ORDEN.F5AFECTO;"
                    End Select
                    
                    If Rs.State = 1 Then Rs.Close
                    
                    Rs.Open sql, cnn_dbbancos, adOpenKeyset, adLockOptimistic
                    'MsgBox rs(3)
                    Dim Nw_Node As DXDBGRIDLibCtl.IdxGridNode
                    Do While Rs.EOF = False
                        i = i + 1
                        If i > 1 Then .dxDBGrid1.Dataset.Append Else .dxDBGrid1.Dataset.Edit
                        
                        .dxDBGrid1.Columns(0).value = i
                        .dxDBGrid1.Columns(1).value = Rs(1)
                        
                        wcodgasto = Rs(1) & ""
'                        Call .dxDBGrid1_OnKeyDown(113, 5)
                        
'                        Call .dxDBGrid1_OnEdited(Nw_Node)
                        .dxDBGrid1.Columns(5).value = Rs(3)
                        .dxDBGrid1.Columns(6).value = IIf(Rs(2) = "*", 1, 0)
                        .dxDBGrid1.Dataset.Post
                        Call .CALCULANDO
                        Call .CALCULAR_TOTALES
                        Rs.MoveNext
                    Loop
                End With
            Else
                wHayDatosImportacion = False
            End If
            
            Unload Me
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Private Sub VerDatos()

    sw_control_items = True
    
    With dxImpDatos
        For J = 1 To dxImpDatos.Dataset.RecordCount
            dxImpDatos.Dataset.RecNo = J
            If dxImpDatos.Columns.ColumnByFieldName("F4CHECK").value = True Then
                   Registro_Compras.TxtRucPrv.Text = dxImpDatos.Columns.ColumnByFieldName("F2CODPROV").value
                'LLENANDO EL DETALLE CON LOS DATOS
                If Rs.State = adStateOpen Then Rs.Close
                For X = 1 To ImpDatos.dxImpDatos.Dataset.RecordCount
                    Registro_Compras.dxDBGrid1.Dataset.Edit
                    Registro_Compras.dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = X * 5
                    Registro_Compras.dxDBGrid1.Dataset.Post
                Next X
            End If
        Next
    End With
    sw_control_items = False
    
End Sub

Private Sub Txtcodcli_DblClick()

    Txtcodcli_KeyDown 113, 0

End Sub

Private Sub Txtcodcli_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodcli = "": wnomcli = "": wruccli = "": WDIRCLI = ""
        ayuda_proveedores_log.Show 1
        If Len(Trim(wcodprov)) > 0 Then
            Txtcodcli.Text = Trim$(wcodprov)
            pnlnomcli.Caption = Trim$(wnomprov)
        End If
    End If

End Sub

Private Sub Txtcodcli_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub Txtcodcli_LostFocus()

    If Trim(Txtcodcli.Text) <> "" Then
        If VALIDA_PROVEEDOR(Txtcodcli.Text) = True Then
            Txtcodcli.Text = wcodprov
            pnlnomcli.Caption = wnomprov
            'Cmbtipref.ListIndex = 0
            Call Configurar1
        Else
            MsgBox "Codigo de Proveedor no existe", vbInformation + vbDefaultButton1, "Atención"
            Txtcodcli.Text = "": Txtcodcli.SetFocus
        End If
    End If

End Sub
Public Sub AdicionaItemImp()
Dim sw_nuevo_temp   As Boolean
    
    
    dxImpDatosDet.Dataset.Active = False

    'If sw_nuevo_doc = False Or Cmbtipref.ListIndex >= 0 Then
    '    DELETEREC_LOG DBTable3, cnn_form
    'End If
    
    dxImpDatosDet.Dataset.ADODataset.ConnectionString = cnn_form
    dxImpDatosDet.Dataset.Active = True
    Call DELETEREC_LOG("DATOSDET", cnn_form)
    dxImpDatosDet.Dataset.Close
    dxImpDatosDet.Dataset.Open
    'adiciono la primera fila
    Dim Codigo As String, nv As String
    With dxImpDatosDet.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
'        For i = 1 To dxImpDatos.Dataset.RecordCount
'          dxImpDatos.Dataset.RecNo = i
'          If dxImpDatos.Columns.ColumnByFieldName("F4CHECK").Value = True Then
'            If rs.State = adStateOpen Then rs.Close
            Select Case Ot_tipo
                Case "V"
                    
                    sql = "SELECT IF3VALES.F2CODALM, IF3VALES.F4NUMVAL, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED,IF3VALES.F3CANPRO, iif(IF4VALES.F4MONEDA= 'S', IF3VALES.F3VALVTA,IF3VALES.F3VALDOL) AS IMPORTE, iif(IF4VALES.F4MONEDA= 'S', IF3VALES.F3TOTITE,IF3VALES.F3TOTDOL) AS TOTAL " & _
                          "FROM IF4VALES INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) " & _
                          "WHERE IF4VALES.F2CODALM in (Select F2CODALM From Temporden in '" & wrutatemp & "\TEMPLUS.MDB' Where F4Check = TRUE) AND IF4VALES.F4NUMVAL  in (Select F4NUMVAL From Temporden IN '" & wrutatemp & "\TEMPLUS.MDB' Where F4Check = TRUE);"
                    
                    'rs.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    'rs.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                Case "C"
                    sql = " SELECT IF3ORDEN.F4NUMORD, IF3ORDEN.F3PRECOS as IMPORTE, IF3ORDEN.F3TOTAL as TOTAL, IF3ORDEN.F3CANPRO, IF5PLA.F5CODPRO, IF5PLA.F7CODMED, IF5PLA.F5NOMPRO, IF3ORDEN.F5AFECTO, IF3ORDEN.F3IGV" & _
                          " FROM IF3ORDEN, IF5PLA Where IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO " & _
                          " And IF3ORDEN.F4NUMORD in (Select F4NUMORD From Temporden in '" & wrutatemp & "\TEMPLUS.MDB' Where F4Check = TRUE);"
            End Select
            Dim Ind As Byte
            Ind = 0
            
            
            dxImpDatos.Dataset.Refresh
            'rs.MoveNext
            'MsgBox rs("F4CHECK")
            'in (Select F4NUMORD From " & _
                          " Temporden in '" & wrutatemp & "\TEMPLUS.MDB' Where F4Check = TRUE);
                          
            
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, cnn_dbbancos, adOpenKeyset, adLockOptimistic
            Do While Not Rs.EOF
            Ind = 1
              If sw_nuevo_temp = False Then
                If sw_nuevo_doc = True Then
                  .Edit
                Else
                  .Append
                End If
                sw_nuevo_temp = True
              Else
                .Append
              End If
              .FieldValues("ITEM") = i
              .FieldValues("CODIGO") = "" & Rs.Fields("F5CODPRO")
              .FieldValues("DESCRIPCION") = "" & Rs.Fields("F5NOMPRO")
              .FieldValues("UNIDAD") = "" & Rs.Fields("F7CODMED")
              .FieldValues("CANTIDAD") = Rs.Fields("F3CANPRO")
              .FieldValues("VVENTAUNIT") = Rs.Fields("IMPORTE")
              .FieldValues("PRECIOUNIT") = Rs.Fields("IMPORTE") * (1 + wIgv)
              Rs.MoveNext
              i = i + 1
            Loop
            If Ind = 1 Then
                .Post
                .ADODataset.Requery
            Else
                
            End If
'            If Ind = 0 Then dxImpDatosDet.Dataset.Delete
            Rs.Close
'          End If
'        Next
        sw_nuevo_item = False
'        dxImpDatos.Dataset.ADODataset.Requery
    End With
    dxImpDatosDet.Dataset.Close
    dxImpDatosDet.Dataset.Open

End Sub



