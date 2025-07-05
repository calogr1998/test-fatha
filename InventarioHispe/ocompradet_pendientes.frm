VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ocompradet_pendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Orden de Compra"
   ClientHeight    =   4635
   ClientLeft      =   1275
   ClientTop       =   1875
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10590
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4155
      Left            =   0
      OleObjectBlob   =   "ocompradet_pendientes.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   10500
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "ocompradet_pendientes.frx":3AFB
      ToolBars        =   "ocompradet_pendientes.frx":86D7
   End
End
Attribute VB_Name = "ocompradet_pendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cconex_form2        As String
Dim cnn_form2           As New ADODB.Connection
Dim nnumorden           As Double

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    If dxDBGrid1.Columns.FocusedIndex = 6 Then
        If Val(Format(dxDBGrid1.Columns(7).value, "0.00")) > Val(Format(dxDBGrid1.Columns(5).value, "0.00")) - Val(Format(dxDBGrid1.Columns(6).value, "0.00")) Then
            MsgBox "La cantidad de ajuste no puede ser mayor", vbCritical, "Atención"
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns(7).value = Format(0, "#0.00")
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim CadSql      As String
    
    Me.MousePointer = vbHourglass
    cconex_form2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form2.Open cconex_form2
        
    LOAD_OC_DET
    
    Me.MousePointer = vbDefault
End Sub

Private Sub LOAD_OC_DET()
    Dim csql            As String
    Dim RSCONSULTA      As New ADODB.Recordset
    Dim ncont           As Integer
    
    nnumorden = Val(ocompra_pendientes.dxDBGrid1.Columns(1).value)

    DELETEREC_LOG cnomtabla2, cnn_form2

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form2
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls

    csql = "SELECT * FROM IF3ORDEN WHERE F4NUMORD=" & nnumorden & " ORDER BY F3CODPRO"
    
    ncont = 1
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        With dxDBGrid1.Dataset
            RSCONSULTA.MoveFirst
            Do While Not RSCONSULTA.EOF
                .Append
                .FieldValues("ITEM") = ncont
                .FieldValues("CODIGO") = RSCONSULTA.Fields("F3CODPRO") & ""
                .FieldValues("FENTREGA") = RSCONSULTA.Fields("F3FENTREGA")
                
                If rsif5pla.State = adStateOpen Then rsif5pla.Close
                rsif5pla.Open "SELECT F5NOMPRO,F7CODMED FROM IF5PLA WHERE F5CODPRO='" & RSCONSULTA.Fields("F3CODPRO") & "'", cnn_dbbancos
                If Not rsif5pla.EOF Then
                    .FieldValues("DESCRIPCION") = rsif5pla.Fields("F5NOMPRO") & ""
                    .FieldValues("UMEDIDA") = rsif5pla.Fields("F7CODMED") & ""
                End If
                rsif5pla.Close
                
                .FieldValues("CANTIDAD") = Format(Val(RSCONSULTA.Fields("F3CANPRO") & ""), "###,###,##0.00")
                .FieldValues("CANT_PEDIDA") = Format(Val(RSCONSULTA.Fields("F3CANPRO") & "") - Val(RSCONSULTA.Fields("F3CANFAL") & ""), "###,###,##0.00")
                .FieldValues("AJUSTE") = Format(Val(RSCONSULTA.Fields("F3CANPRO") & "") - (Val(RSCONSULTA.Fields("F3CANPRO") & "") - Val(RSCONSULTA.Fields("F3CANFAL") & "")), "###,###,##0.00")
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
    cnn_form2.Close

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Grabar":
            GRABA_AJUSTE
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub

Private Sub GRABA_AJUSTE()
Dim csql            As String
Dim RSCONSULTA      As New ADODB.Recordset
Dim sw              As Boolean

    sw = False
    csql = "SELECT * FROM " & cnomtabla2 & ""
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open csql, cconex_form2, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        RSCONSULTA.MoveFirst
        Do While Not RSCONSULTA.EOF
            csql = "UPDATE IF3ORDEN SET F3AJUSTE=" & RSCONSULTA.Fields("AJUSTE") & _
                   ",F3CANFAL=F3CANFAL-" & RSCONSULTA.Fields("AJUSTE") & _
                   " WHERE F4NUMORD=" & nnumorden & " AND F3CODPRO = '" & RSCONSULTA.Fields("CODIGO") & "'"
            cnn_dbbancos.Execute (csql)
            AlmacenaQuery_sql csql, cnn_dbbancos
            sw = True
            RSCONSULTA.MoveNext
        Loop
    End If
    RSCONSULTA.Close
    
''    If sw = True Then
''        csql = "UPDATE " & cnomtabla & " SET ESTADO = ' ' WHERE NUMORDEN = '" & Format(nnumorden, "00000000") & "'"
''    End If
''    cnn_form2.Execute (csql)

End Sub
