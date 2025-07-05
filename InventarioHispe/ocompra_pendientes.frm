VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ocompra_pendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra Pendientes"
   ClientHeight    =   7005
   ClientLeft      =   1185
   ClientTop       =   1845
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10365
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6900
      Left            =   0
      OleObjectBlob   =   "ocompra_pendientes.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   10335
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "ocompra_pendientes.frx":3B6F
      ToolBars        =   "ocompra_pendientes.frx":874B
   End
End
Attribute VB_Name = "ocompra_pendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection

Private Sub dxDBGrid1_OnDblClick()

    ocompradet_pendientes.Show 1

End Sub

Private Sub Form_Load()
Dim CadSql      As String
    Me.left = 1600
    Me.top = 1050
    '--------- CREA LA TABLA DE LA CABECERA
    'cnombase = wusuario & "OC" & Format(Time, "hh_mm_ss") & ".MDB"
    cnombase = "TEMPLUS.MDB"
    'CREATEDATABASE_N wrutatemp & "\", cnombase
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    cnomtabla = "OC_PEND"
    
    
    'CadSql = "(ITEM TEXT(4),FECHA DATE,NUMORDEN TEXT(8),PROVEEDOR TEXT(100)," _
            & " SOLICITANTE TEXT(8),SOLICITUD TEXT(6),OBRA TEXT(8)," _
            & " OBSERVA TEXT(100))"
    'CREATETABLE_N cnomtabla, CadSql, cnn_form
    
    '--------- CREA LA TABLA DEL DETALLE
    cnomtabla2 = "OC_DETPEND"
    'CadSql = "(ITEM TEXT(4),FENTREGA DATE,CODIGO TEXT(15),DESCRIPCION TEXT(100)," _
            & " UMEDIDA TEXT(5),CANTIDAD DOUBLE,CANT_PEDIDA DOUBLE,AJUSTE DOUBLE)"
    'CREATETABLE_N cnomtabla2, CadSql, cnn_form
    
    Me.MousePointer = vbHourglass
    LOAD_OC_PEND
    Me.MousePointer = vbDefault

End Sub

Private Sub LOAD_OC_PEND()
Dim csql            As String
Dim RSCONSULTA      As New ADODB.Recordset
Dim ncont           As Integer
    
    DELETEREC_LOG cnomtabla, cnn_form

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls

    If ctipoadm_bd = "M" Then
        csql = "SELECT IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4OBSERVA, IF4ORDEN.F4CODPRV, IF4ORDEN.F4CODSOL, IF4ORDEN.F4CODSOLICITUD " _
             & "FROM IF3ORDEN LEFT JOIN IF4ORDEN ON IF3ORDEN.F4NUMORD = IF4ORDEN.F4NUMORD " _
             & "GROUP BY IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4CENTRO, IF4ORDEN.F4OBSERVA, IF4ORDEN.F4CODPRV, IF4ORDEN.F4CODSOL, IF4ORDEN.F4CODSOLICITUD " _
             & "HAVING (((Sum(IF3ORDEN.F3CANFAL))>0)) ORDER BY IF4ORDEN.F4NUMORD;"
    Else
        csql = "SELECT F4NUMORD,F4FECEMI,F4CENTRO,F4OBSERVA,F4CODPRV,F4CODSOL,F4CODSOLICITUD FROM IF4ORDEN " & _
               "WHERE F4NUMORD IN (SELECT DISTINCTROW IF4ORDEN.F4NUMORD " & _
               "FROM IF4ORDEN INNER JOIN IF3ORDEN ON IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD " & _
               "GROUP BY IF4ORDEN.F4NUMORD HAVING (((Sum(IF3ORDEN.F3CANFAL))>0))) ORDER BY F4NUMORD;"
    End If
    
    ncont = 1
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        With dxDBGrid1.Dataset
            RSCONSULTA.MoveFirst
            Do While Not RSCONSULTA.EOF
                .Append
                .FieldValues("ITEM") = ncont
                .FieldValues("NUMORDEN") = RSCONSULTA.Fields("F4NUMORD")
                .FieldValues("FECHA") = RSCONSULTA.Fields("F4FECEMI")
                
                If RsProveedor.State = adStateOpen Then RsProveedor.Close
                RsProveedor.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & RSCONSULTA.Fields("F4CODPRV") & "'", cnn_dbbancos
                If Not RsProveedor.EOF Then
                    .FieldValues("PROVEEDOR") = RsProveedor.Fields("F2NOMPROV") & ""
                End If
                RsProveedor.Close
                
                .FieldValues("SOLICITANTE") = RSCONSULTA.Fields("F4CODSOL") & ""
                .FieldValues("SOLICITUD") = RSCONSULTA.Fields("F4CODSOLICITUD") & ""
                .FieldValues("OBRA") = RSCONSULTA.Fields("F4CENTRO") & ""
                '.FieldValues("OBSERVA") = rsconsulta.Fields("F4OBSERVA") & ""
                RSCONSULTA.MoveNext
                ncont = ncont + 1
            Loop
            .Post
        End With
    End If
    RSCONSULTA.Close

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
           
    dxDBGrid1.OptionEnabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    cnn_form.Close
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub
