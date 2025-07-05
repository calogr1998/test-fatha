VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frm_cons 
   Appearance      =   0  'Flat
   Caption         =   "Asientos Contables"
   ClientHeight    =   8910
   ClientLeft      =   -6960
   ClientTop       =   3870
   ClientWidth     =   15420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8910
   ScaleWidth      =   15420
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   8355
      Left            =   60
      OleObjectBlob   =   "FRM_CONS.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   15990
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "FRM_CONS.frx":524B
      ToolBars        =   "FRM_CONS.frx":84F3
   End
End
Attribute VB_Name = "frm_cons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim cnn_mov As New ADODB.Connection
'Dim dbconta_cons    As DAO.Database
'Dim CONSULTA        As DAO.Recordset
'Dim sql             As String
'
'Private Sub BtnExitO_Click()
'
'   Unload Me
'
'End Sub
'
'
'
'Private Sub cmdREP_Click()
'Dim dbtabla_conta   As DAO.Database
'Dim tbplanc         As DAO.Recordset
'
'    dbconta_cons.Execute ("delete from REPORTE")
'    datareporte.DatabaseName = wrutatemp & "\db_conta.Mdb"
'    datareporte.RecordSource = "REPORTE"
'    datareporte.Refresh
'
'    Set dbtabla_conta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
'    Set tbplanc = dbtabla_conta.OpenRecordset("cf5pla")
'    tbplanc.Index = "cf5pla"
'
'    sql = "SELECT contable.F5CODCTA,contable.F3DEBHAB, Sum(contable.F3IMPORTE) AS SumaSOLES, "
'    sql = sql + "Sum(contable.F3IMPORTED) AS SumaDOLARES FROM contable GROUP BY contable.F5CODCTA, contable.F3DEBHAB;"
'    Set CONSULTA = dbconta_cons.OpenRecordset(sql)
'
'    CONSULTA.MoveFirst
'    Do While Not CONSULTA.EOF
'        datareporte.Recordset.AddNew
'        'DATAREPORTE.Recordset("COMPROBANTE") = CONSULTA("F3COMPRO")
'        'DATAREPORTE.Recordset("DETALLE") = CONSULTA("F3DETALL")
'
'        tbplanc.Seek "=", CONSULTA("F5CODCTA") & ""
'        If Not tbplanc.NoMatch Then
'            datareporte.Recordset("DETALLE") = tbplanc.Fields("f5nomcta") & ""
'        End If
'
'        datareporte.Recordset("CUENTA") = CONSULTA("F5CODCTA") & ""
'        If CONSULTA("F3DEBHAB") = "D" Then
'            datareporte.Recordset("DEBES") = CONSULTA("SumaSOLES")
'            datareporte.Recordset("DEBED") = CONSULTA("SumaDOLARES")
'        Else
'            datareporte.Recordset("HABERS") = CONSULTA("SumaSOLES")
'            datareporte.Recordset("HABERD") = CONSULTA("SumaDOLARES")
'        End If
'        datareporte.Recordset("empresa") = wnomcia
'        datareporte.Recordset.Update
'        CONSULTA.MoveNext
'    Loop
'    datareporte.Refresh
'    datareporte.Recordset.Close
'
'    tbplanc.Close
'    dbtabla_conta.Close
'
'    Report1.DataFiles(0) = wrutatemp & "\db_conta.mdb"
'    Report1.ReportFileName = wrutatemp & "\REPGEN.rpt"
'    Report1.Action = 1
'
'End Sub
'
'Private Sub cmdtrans_Click()
'Dim dbusuarios  As DAO.Database
'Dim TbUser      As DAO.Recordset
'Dim cnivel      As String
'
'    Set dbusuarios = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
'    Set TbUser = dbusuarios.OpenRecordset("EF2USERS")
'    TbUser.Index = "IDCODUSER"
'
'    TbUser.Seek "=", wusuario
'    If Not TbUser.NoMatch Then
'        cnivel = "" & TbUser.Fields("F2NIVEL")
'    End If
'    TbUser.Close
'    dbusuarios.Close
'
'    If cnivel = "1" Then
'        frmcnttr.Show 1
'    Else
'        MsgBox "Ud. no tiene acceso a esta opción."
'    End If
'
'End Sub
'
'Private Sub Form_Load()
'
'
'    If cnn_mov.State = 1 Then cnn_mov.Close
'    cnn_mov.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\DB_CONTA.MDB" & ";Persist Security Info=False"
'
'    FILL
'End Sub
'
'Private Sub FILL()
'Dim csql As String
'    csql = "SELECT *,"
'    csql = csql & "iif(f3moneda='S',iif(f3debhab='D',f3importe,0),0) as f3importedeb_MN,iif(f3moneda='S',iif(f3debhab='H',f3importe,0),0) as f3importeHAB_MN, "
'    csql = csql & "iif(f3moneda='D',iif(f3debhab='D',f3importe,0),0) as f3importedeb_ME,iif(f3moneda='D',iif(f3debhab='H',f3importe,0),0) as f3importeHAB_ME "
'    csql = csql & "FROM CONTABLE ORDER BY F3COMPRO,F3ELEMEN"
'
'
'
'    With dxDBGrid1
''        .DefaultFields = True
'        .Dataset.ADODataset.ConnectionString = cnn_mov
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.Active = True
'        .KeyField = "item"
'    End With
'End Sub
'
'
'Private Sub imprimir()
'Dim dbcontax As DAO.Database
'Dim tbcontax As DAO.Recordset
'Dim dbtempo As DAO.Database
'Dim tbtempo As DAO.Recordset
'
'On Error GoTo error_bd
'
'   Set dbcontax = OpenDatabase(wrutatemp & "\db_conta.mdb")
'   Set tbcontax = dbcontax.OpenRecordset("contable")
'
'   Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
'   dbtempo.Execute ("delete from conta_rep")
'   Set tbtempo = dbtempo.OpenRecordset("conta_rep")
'
'    If tbcontax.RecordCount <> 0 Then
'       tbcontax.MoveFirst
'       Do While Not tbcontax.EOF
'          tbtempo.AddNew
'          tbtempo.Fields("f3compro") = tbcontax.Fields("f3compro")
'          tbtempo.Fields("f3elemen") = tbcontax.Fields("f3elemen")
'          tbtempo.Fields("f3origen") = tbcontax.Fields("f3origen")
'          tbtempo.Fields("f3fchopr") = tbcontax.Fields("f3fchopr")
'          tbtempo.Fields("f3detall") = tbcontax.Fields("f3detall")
'          tbtempo.Fields("f5codcta") = tbcontax.Fields("f5codcta")
'          tbtempo.Fields("f3cheque") = tbcontax.Fields("f3cheque")
'          tbtempo.Fields("f3nroref") = tbcontax.Fields("f3nroref")
'          tbtempo.Fields("f3debhab") = tbcontax.Fields("f3debhab")
'          If tbcontax.Fields("f3debhab") = "D" Then
'             tbtempo.Fields("f3debes") = tbcontax.Fields("f3importe")
'             tbtempo.Fields("f3debed") = tbcontax.Fields("f3imported")
'          Else
'             tbtempo.Fields("f3habers") = tbcontax.Fields("f3importe")
'             tbtempo.Fields("f3haberd") = tbcontax.Fields("f3imported")
'          End If
'          tbtempo.Fields("f3moneda") = tbcontax.Fields("f3moneda")
'          tbtempo.Fields("f3tipcambd") = tbcontax.Fields("f3tipcambd")
'          tbtempo.Fields("f3tipdoc") = tbcontax.Fields("f3tipdoc")
'          tbtempo.Fields("f3costo") = tbcontax.Fields("f3costo")
'          tbtempo.Fields("empresa") = wnomcia
'          tbtempo.Fields("mes") = Mid(tbcontax.Fields("f3proame"), 5, 2)
'          tbtempo.Update
'          tbcontax.MoveNext
'          If tbcontax.EOF Then Exit Do
'       Loop
'    Else
'       MsgBox "No se han generado los asientos contables. Verifique.", 48, "Compras"
'    End If
'
'    tbcontax.Close
'    dbcontax.Close
'
'    tbtempo.Close
'    dbtempo.Close
'
'    Exit Sub
'
'error_bd:
'    Resume Next
'
'End Sub
'
'Private Sub SSPanel1_Click()
'
'End Sub
