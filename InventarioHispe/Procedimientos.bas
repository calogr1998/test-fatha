Attribute VB_Name = "Procedimientos"
Public codProdtem        As String
Public codtrans        As String
Public addCliFac         As Boolean
Public wserie_fac       As String
Public wserie_bol       As String
Public wRucEmpTrans     As String
Public wDescEmpTrans    As String
Public wCodNivel       As String
Public wDescNivel       As String
Public wCodigoDep     As String
Public wDesDep    As String

Public wUserCaja        As String

Public wLocal       As String


Public almaTrans    As String

Public comisionVen      As Double

Public Function traerCampo(tabla As String, campo As String, campoCom As String, valor As String, Optional condicion As String) As String
    Dim cad As String
    Dim rst As New Recordset
    If IsDate(valor) Then
        cad = "select " & campo & " from " & tabla & " where CVDATE(" & campoCom & ") = '" & valor & "' " & condicion
    Else
        cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
    End If
    rst.Open cad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    traerCampo = ""
    If Not rst.EOF And Not IsNull(rst.Fields(0)) Then traerCampo = rst.Fields(0)
End Function


Public Sub llenaCombo(c As ComboBox, tabla As String, campo As String, Optional cod_campo As String, Optional condicion As String)
    Dim RsFill  As New ADODB.Recordset
    Dim csql As String
    Dim cadd As String
    csql = "select " & campo
    If cod_campo <> "" Then csql = csql & "," & cod_campo
    csql = csql & " from " & tabla
    If condicion <> "" Then csql = csql & " " & condicion
    RsFill.Open csql, cnn_dbbancos
        c.Clear
        Do While Not RsFill.EOF
            cadd = RsFill.Fields(0)
            If cod_campo <> "" Then cadd = cadd & Space(100) & RsFill.Fields(1)
            c.AddItem cadd
            RsFill.MoveNext
        Loop
        c.ListIndex = -1
    RsFill.Close: Set RsFill = Nothing
End Sub

Public Sub AsignaDerechos()
'dim rs as
    For j = 1 To 4
        derecho(j) = False
    Next j
    Set rst = New Recordset
    SQL = "select * from ef2users_der where f2coduser='" & wusuario & "' order by codigo"
    If rst.State = adStateOpen Then rst.Close
    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        Do While Not rst.EOF
            Select Case Val("" & rst("codigo"))
                Case 1
                    derecho(1) = True
                Case 2
                    derecho(2) = True
                Case 3
                    derecho(3) = True
                Case 4
                    derecho(4) = True
                Case 10
                    derecho(10) = True
            End Select
            rst.MoveNext
        Loop
    End If
    rst.Close
    
''    For j = 1 To 4
''        If derecho(j) Then
''            vaTabPro1.Tab = j - 1
''            vaTabPro1.TabState = 0
''        Else
''            vaTabPro1.Tab = j - 1
''            vaTabPro1.TabState = 2
''        End If
''    Next j

End Sub


