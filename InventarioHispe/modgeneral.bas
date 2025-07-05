Attribute VB_Name = "modgeneral"

Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


#If Win32 Then
    'Declaraciones para 32 bits
    Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    'Declaraciones para 16 bits
    Private Declare Function GetPrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lplFileName As String) As Integer
#End If

Type a_MontosFechas
    MONTO As String
    Fecha As String
End Type

Type a_numero
    valor As Integer
End Type

Type a_solicitud
    TOTAL As Byte
    numero As String
End Type

Type a_grabacion
    campo   As String
    valor   As String
    Tipo    As String
End Type
Public sw_GRABA_REGISTRO_logistica    As Boolean
Public Sw_Ejecuta_Sentencia As Boolean
Global lista() As a_solicitud
Global wllamada As Byte
Global oTipoRequerimiento As String

Public Function traerCampo(tabla As String, campo As String, campoCom As String, valor As String, Optional condicion As String) As String
    Dim cad As String
    Dim rst As New Recordset
    If IsDate(valor) Then
        cad = "select " & campo & " from " & tabla & " where CVDATE(" & campoCom & ") = '" & valor & "' " & condicion
    'ElseIf IsNumeric(valor) Then
        'cad = "select " & campo & " from " & tabla & " where " & campoCom & " = " & valor & " " & condicion
    Else
        cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
    End If
    If tabla = "srutas" Then
        cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RUTAS.MDB" & ";Persist Security Info=False"
        If cnn_control.State = 1 Then cnn_control.Close
        cnn_control.Open cconex_control
        rst.Open cad, cnn_control, adOpenForwardOnly, adLockReadOnly
    Else
        rst.Open cad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    traerCampo = ""
    If Not rst.EOF And Not IsNull(rst.Fields(0)) Then traerCampo = rst.Fields(0)
End Function

Public Function OBTIENE_CORRELA_XPAGAR(pconexion As ADODB.Connection)
Dim rspagcab            As New ADODB.Recordset
Dim ncorre              As Double

    If rspagcab.State = adStateOpen Then rspagcab.Close
    rspagcab.Open "SELECT CORRELA FROM PAG_DCTO ORDER BY CORRELA DESC", pconexion, adOpenDynamic, adLockOptimistic
    If Not rspagcab.EOF Then
        rspagcab.MoveFirst
        ncorre = rspagcab.Fields("correla") + 1
    Else
        ncorre = 1
    End If
    rspagcab.Close
    OBTIENE_CORRELA_XPAGAR = ncorre

End Function

Public Sub TRANS_CTASXPAGAR_NEW(ptipo As String, pvia As String, pcorrela As Double, ptipdocu As String, pserdoc As String, pdocum As String, pfecha As Date, pruc As String, pcodigo As String, pmoneda As String, ptipcam As Double, ptotal As Double, pdebhab As String, prefer As String, pfechavenc As Date, pcentro As String, pnomcodigo As String, pconexion As String, preg_com As String, pannorc As String, pordcompra As String)
'On Error GoTo error_trans
'Dim CnTmp As New ADODB.Connection
'Dim amovs_cab(0 To 21)  As a_grabacion
'Dim sw_llena            As Boolean
'Dim ctipo_graba         As String * 1
'Dim cwhere              As String
'Dim ncorrela            As Double
'Dim rspagcab            As New ADODB.Recordset
'    '***abre conexion
'    If CnTmp.State = 1 Then CnTmp.Close
'    CnTmp.Open pconexion
'    '****************
'    sw_llena = False
'    If pcorrela > 0 Then
'        '--------- MODIFICACION
'        If rspagcab.State = adStateOpen Then rspagcab.Close
'        rspagcab.Open "SELECT * FROM PAG_DCTO WHERE CORRELA=" & pcorrela & "", CnTmp, adOpenDynamic, adLockOptimistic
'        If Not rspagcab.EOF Then
'            If Val(Format(rspagcab.Fields("TOTAL"), "0.00")) = Val(Format(rspagcab.Fields("SALDO"), "0.00")) Then
'                sw_llena = True
'                ctipo_graba = "M"
'                cwhere = "CORRELA = " & pcorrela & ""
'                ncorrela = pcorrela
'                CorrelaPagDcto = ncorrela
'            Else
'                sw_llena = False
'                'MsgBox "El documento ya ha sido aplicado.", vbInformation, "CONTROL Plus!"
'
'                MsgBox "El documento ya ha sido aplicado.", vbInformation, wnomcia
'                ncorrela = pcorrela
'                CorrelaPagDcto = ncorrela
'            End If
'        End If
'        rspagcab.Close
'    Else
'        '--------- NUEVO
'        sw_llena = True
'        ctipo_graba = "A"
'        cwhere = ""
'        ncorrela = OBTIENE_CORRELA_XPAGAR(CnTmp)
'        CorrelaPagDcto = ncorrela
'    End If
'    'wcorrela = ncorrela
'    If sw_llena = True Then
'        amovs_cab(0).campo = "VIA_INGR": amovs_cab(0).valor = pvia: amovs_cab(0).Tipo = "T"
'        amovs_cab(1).campo = "CORRELA": amovs_cab(1).valor = ncorrela: amovs_cab(1).Tipo = "N"
'        If Len(Trim(pserdoc)) > 0 Then
'            amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = ptipdocu & pserdoc & "/" & pdocum: amovs_cab(2).Tipo = "T"
'        Else
'            amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = ptipdocu & pdocum: amovs_cab(2).Tipo = "T"
'        End If
'
'        amovs_cab(3).campo = "FCH_COMP": amovs_cab(3).valor = pfecha: amovs_cab(3).Tipo = "F"
'        amovs_cab(4).campo = "PROVEEDOR": amovs_cab(4).valor = pcodigo: amovs_cab(4).Tipo = "T"
'        amovs_cab(5).campo = "PROVEEDORO": amovs_cab(5).valor = pcodigo: amovs_cab(5).Tipo = "T"
'        amovs_cab(6).campo = "MONEDA": amovs_cab(6).valor = pmoneda: amovs_cab(6).Tipo = "T"
'        amovs_cab(7).campo = "MONEDAO": amovs_cab(7).valor = pmoneda: amovs_cab(7).Tipo = "T"
'        amovs_cab(8).campo = "TCAMBIO": amovs_cab(8).valor = ptipcam: amovs_cab(8).Tipo = "N"
'        amovs_cab(9).campo = "TCAMBIOO": amovs_cab(9).valor = ptipcam: amovs_cab(9).Tipo = "N"
'        amovs_cab(10).campo = "TOTAL": amovs_cab(10).valor = ptotal: amovs_cab(10).Tipo = "N"
'        amovs_cab(11).campo = "TOTALO": amovs_cab(11).valor = ptotal: amovs_cab(11).Tipo = "N"
'        amovs_cab(12).campo = "SALDO": amovs_cab(12).valor = ptotal: amovs_cab(12).Tipo = "N"
'        amovs_cab(13).campo = "DEB_HAB": amovs_cab(13).valor = pdebhab: amovs_cab(13).Tipo = "T"
'        amovs_cab(14).campo = "REFERENCIA": amovs_cab(14).valor = prefer: amovs_cab(14).Tipo = "T"
'        amovs_cab(15).campo = "FCH_VCTO": amovs_cab(15).valor = pfechavenc: amovs_cab(15).Tipo = "F"
'        amovs_cab(16).campo = "F4CENTRO": amovs_cab(16).valor = pcentro: amovs_cab(16).Tipo = "T"
'        amovs_cab(17).campo = "NOMPROV": amovs_cab(17).valor = pnomcodigo: amovs_cab(17).Tipo = "T"
'        amovs_cab(18).campo = "REG_COM": amovs_cab(18).valor = Mid(preg_com, 5, 9): amovs_cab(18).Tipo = "T"
'        amovs_cab(19).campo = "RUC": amovs_cab(19).valor = pruc: amovs_cab(19).Tipo = "T"
'        amovs_cab(20).campo = "F4ANNORC": amovs_cab(20).valor = pannorc: amovs_cab(20).Tipo = "T"
'        amovs_cab(21).campo = "F4OCOMPRA": amovs_cab(21).valor = pordcompra: amovs_cab(21).Tipo = "T"
'
'        GRABA_REGISTRO amovs_cab(), "PAG_DCTO", ctipo_graba, 21, pconexion, cwhere
'
'        csql = "UPDATE REGISDOC SET F4CORRELA=" & ncorrela & " WHERE F4MESMOV='" & (Mid(preg_com, 1, 6)) & "' AND F4NUMMOV='" & Format(right(preg_com, 7), "0000000") & "'"
'        CnTmp.Execute (csql)
'
'
'    End If
'    '****cierra conex
'    If CnTmp.State = 1 Then CnTmp.Close
'    Set CnTmp = Nothing
'    'validando registro de compras
'    Exit Sub
'error_trans:
'    MsgBox "Se ha producido el sgte. error : " & Err.Description, 16, wnomcia
'    Resume
'    Exit Sub
End Sub

 


Sub TRANSCTACTE(ptipo As String, pvia As String, pcorrela As Double, ptipdocu As String, pserdoc As String, pdocum As String, pfecha As Date, pruc As String, pcodigo As String, pmoneda As String, ptipcam As Double, ptotal As Double, pdebhab As String, prefer As String, pfechavenc As Date, pcentro As String, pnomcodigo As String, pconexion As ADODB.Connection)
On Error GoTo error_trans
Dim amovs_cab(0 To 17)  As a_grabacion
Dim sw_llena            As Boolean
Dim ctipo_graba         As String * 1
Dim cwhere              As String
Dim ncorrela            As Double

    sw_llena = False
    If pcorrela > 0 Then
        '--------- MODIFICACION
        RSCTA_DCTO.Open "SELECT * FROM CTA_DCTO WHERE TIPO='" & ptipo & "' AND CORRELA=" & pcorrela & "", pconexion, adOpenDynamic, adLockOptimistic
        If Not RSCTA_DCTO.EOF Then
            If Val(Format(RSCTA_DCTO.Fields("TOTAL"), "0.00")) = Val(Format(RSCTA_DCTO.Fields("SALDO"), "0.00")) Then
                sw_llena = True
                ctipo_graba = "M"
                cwhere = "TIPO = '" & ptipo & "' AND CORRELA = " & pcorrela & ""
                ncorrela = pcorrela
            Else
                sw_llena = False
                MsgBox "El documento ya ha sido aplicado.", vbInformation, "Atención"
            End If
        End If
        RSCTA_DCTO.Close
    Else
        '--------- NUEVO
        sw_llena = True
        ctipo_graba = "A"
        cwhere = ""
        ncorrela = OBTIENE_CORRELA(pconexion)
    End If
    
    If sw_llena = True Then
        amovs_cab(0).campo = "TIPO": amovs_cab(0).valor = ptipo: amovs_cab(0).Tipo = "T"
        amovs_cab(1).campo = "VIA_INGR": amovs_cab(1).valor = pvia: amovs_cab(1).Tipo = "T"
        amovs_cab(2).campo = "CORRELA": amovs_cab(2).valor = ncorrela: amovs_cab(2).Tipo = "N"
        amovs_cab(3).campo = "TIPDOCU": amovs_cab(3).valor = ptipdocu: amovs_cab(3).Tipo = "T"
        amovs_cab(4).campo = "SERDOC": amovs_cab(4).valor = pserdoc: amovs_cab(4).Tipo = "T"
        amovs_cab(5).campo = "DOCUM": amovs_cab(5).valor = pdocum: amovs_cab(5).Tipo = "T"
        amovs_cab(6).campo = "FECHA": amovs_cab(6).valor = pfecha: amovs_cab(6).Tipo = "F"
        amovs_cab(7).campo = "RUC": amovs_cab(7).valor = pruc: amovs_cab(7).Tipo = "T"
        amovs_cab(8).campo = "CODIGO": amovs_cab(8).valor = pcodigo: amovs_cab(8).Tipo = "T"
        amovs_cab(9).campo = "MONEDA": amovs_cab(9).valor = pmoneda: amovs_cab(9).Tipo = "T"
        amovs_cab(10).campo = "TIPCAM": amovs_cab(10).valor = ptipcam: amovs_cab(10).Tipo = "N"
        amovs_cab(11).campo = "TOTAL": amovs_cab(11).valor = ptotal: amovs_cab(11).Tipo = "N"
        amovs_cab(12).campo = "SALDO": amovs_cab(12).valor = ptotal: amovs_cab(12).Tipo = "N"
        amovs_cab(13).campo = "DEB_HAB": amovs_cab(13).valor = pdebhab: amovs_cab(13).Tipo = "T"
        amovs_cab(14).campo = "REFERENCIA": amovs_cab(14).valor = prefer: amovs_cab(14).Tipo = "T"
        amovs_cab(15).campo = "FECHA_VCTO": amovs_cab(15).valor = pfechavenc: amovs_cab(15).Tipo = "F"
        amovs_cab(16).campo = "F4CENTRO": amovs_cab(16).valor = pcentro: amovs_cab(16).Tipo = "T"
        amovs_cab(17).campo = "NOMCODIGO": amovs_cab(17).valor = pnomcodigo: amovs_cab(17).Tipo = "T"
        
        GRABA_REGISTRO_logistica amovs_cab(), "CTA_DCTO", ctipo_graba, 17, pconexion, cwhere
        
    End If
        
    Exit Sub
    
error_trans:
    MsgBox "Se ha producido el sgte. error : " & Err.Description, 16, "Error: " & Err.Number
    Exit Sub
    
End Sub

Public Sub CREATETABLE_N(ptabla As String, pcadena As String, pconexion As ADODB.Connection)
On Error Resume Next
Dim query As String
    
    query = "Create table " & ptabla & " " & pcadena
    pconexion.Execute (query)
    'AlmacenaQuery_sql query, pconexion
End Sub

Public Sub DELETEREC_LOG(ptabla As String, pconexion As ADODB.Connection)
On Error Resume Next
    pconexion.Execute ("DELETE * FROM " & ptabla)
    'AlmacenaQuery_sql "DELETE * FROM " & ptabla, pconexion
End Sub
Public Sub DELETEREC_BANCOS(ptabla As String, pconexion As String, pwhere As String)
On Error GoTo ErrorDelete

Dim Cn_Tmp As New ADODB.Connection
Dim CadSql As String
    Cn_Tmp.Open pconexion
    
    
    CadSql = "DELETE * FROM " & ptabla
    
    If Len(Trim(pwhere)) > 0 Then
        CadSql = CadSql & " where " & pwhere
    End If
    
    Cn_Tmp.Execute CadSql
    
    Actualiza_Log CadSql, pconexion
    
    
    If Cn_Tmp.State = 1 Then Cn_Tmp.Close
    Set Cn_Tmp = Nothing
    
    
    Exit Sub

ErrorDelete:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
End Sub



Public Sub CREATEDATABASE_N(pruta As String, pnombase As String)
Dim DBName      As String
Dim cat         As ADOX.Catalog
Dim Tbl         As ADOX.Table

    DBName = pruta & pnombase
    Set cat = New ADOX.Catalog
    Set Tbl = New ADOX.Table
    
    If Len(Dir$(DBName)) Then
        Kill DBName
    End If
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ";"
    cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ";"
            
End Sub

Public Sub GRABA_REGISTRO_logistica(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String)
    On Error GoTo CapturaError
    Dim I           As Integer
    Dim ccampos     As String
    Dim cvalores    As String
    Dim csql        As String
    Dim prSol As New Recordset
    
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For I = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    If parreglo(I).valor <> vbNullString Then
                        cvalores = "'" & parreglo(I).valor & "'"
                    Else
                        cvalores = "Null"
                    End If
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = parreglo(I).valor
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    If parreglo(I).valor <> vbNullString Then
                        cvalores = cvalores & ",'" & parreglo(I).valor & "'"
                    Else
                        cvalores = cvalores & ",Null"
                    End If
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(I).Tipo = "T" Then
                    If parreglo(I).valor <> vbNullString Then
                        cvalores = cvalores & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                    Else
                        cvalores = cvalores & parreglo(I).campo & "=Null"
                    End If
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(I).campo & "=" & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(I).campo & "=" & Format(parreglo(I).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(I).Tipo = "T" Then
                    If parreglo(I).valor <> vbNullString Then
                        cvalores = cvalores & "," & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                    Else
                        cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                    End If
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(I).campo & "=" & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        With objAyudaVale
            .SQLSelectAlter = .NumeroVale
            
            '.TipoVale = left(cnumvale, 1)
            '.CodigoAlmacen = IIf(Not sw_ingreso, Trim(txtalmacen.Text), Trim(txtalmacendes.Text))
            '.NumeroVale = IIf(Not sw_ingreso, Trim(txtnumero.Text), Trim(txtdestino.Text))
            '.Fecha = abofecha.value
            
            If .verificarExistencia Then
                .NumeroVale = .generarNumeroVale
                '.NumeroVale = "I-101108"
                
                cvalores = Replace(cvalores, .SQLSelectAlter, .NumeroVale, 1)
            End If
        End With
        
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If
  
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    pconexion.Execute csql
    sw_GRABA_REGISTRO_logistica = True
'    AlmacenaQuery_sql csql, pconexion
    Actualiza_Log csql, pconexion.ConnectionString
    AlmacenaQuery_sql csql, pconexion
    
    
    Exit Sub
    Resume
CapturaError:

    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atencion"
    Resume Next
End Sub

Public Sub GRABA_REGISTRO(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As String, pwhere As String)
On Error GoTo Error_Graba_Registro
Dim CnSave As New ADODB.Connection
Dim I           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
    
    Sw_Graba_Registro = False
    
    CnSave.Open pconexion
    
    ccampos = "": cvalores = ""
    For I = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    cvalores = "'" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = 0 & parreglo(I).valor
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & ",'" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(I).Tipo = "@" Then
                    cvalores = cvalores & parreglo(I).campo & "=" & parreglo(I).valor & " "
                End If
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(I).campo & "=" & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(I).campo & "=" & Format(parreglo(I).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(I).Tipo = "@" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "=" & parreglo(I).valor & " "
                End If
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(I).campo & "=" & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If

    
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    CnSave.Execute csql
    
    
    Actualiza_Log csql, pconexion
    Sw_Graba_Registro = True
    
    If CnSave.State = 1 Then CnSave.Close
    Set CnSave = Nothing
    
    Exit Sub
    
Error_Graba_Registro:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Select Case Err.Number
    Case 3704
        If CnSave.State = 0 Then CnSave.Open
        Resume
    Case -2147467259
        If MsgBox("¿Desea volver a intentar?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
            For J = 0 To 10000
                If CnSave.State = 1 Then CnSave.Close
                Set CnSave = Nothing
            Next
            'If CnSave.State = 1 Then CnSave.Close
            CnSave.Open pconexion
            Resume
        Else
            Exit Sub
        End If
    End Select
    Sw_Graba_Registro = False
    
    Exit Sub
    
End Sub


Public Sub GRABA_REGISTRO_noenvia(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String)
On Error GoTo CapturaError
Dim I           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
Dim prSol As New Recordset
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For I = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    cvalores = "'" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = 0 & parreglo(I).valor
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(I).campo
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & ",'" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(I).campo & "=" & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(I).campo & "=" & Format(parreglo(I).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(I).Tipo = "T" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "='" & parreglo(I).valor & "'"
                End If
                If parreglo(I).Tipo = "N" Then
                    If Not IsNumeric(parreglo(I).valor) Then
                        cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(I).campo & "=" & 0 & parreglo(I).valor & ""
                    End If
                End If
                If parreglo(I).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(I).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parreglo(I).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(I).valor) Then
                            cvalores = cvalores & "," & parreglo(I).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(I).campo & "=CVDATE('" & parreglo(I).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If
  
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    pconexion.Execute csql
    sw_GRABA_REGISTRO_logistica = True
    Actualiza_Log csql, pconexion.ConnectionString
'    AlmacenaQuery_sql csql, pconexion
    'AlmacenaQuery_sql csql, pconexion
    
    
    Exit Sub
CapturaError:

    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atencion"
    
End Sub


'Public Sub AlmacenaQuery_sql(ByVal sql As String, ConeccionAdodb As ADODB.Connection)
'On Error GoTo CapturaError
'If InStr(UCase(ConeccionAdodb), "DB_BANCOS.MDB") > 0 Or InStr(UCase(ConeccionAdodb), "DB_TABLA.MDB") > 0 Then
'    Dim cnEnvia As New ADODB.Connection
'    If cnEnvia.State = 1 Then cnEnvia.Close
'    cnEnvia.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_envia.mdb;Persist Security Info=False"
'    sql = Replace(sql, "'", "|")
'
'    csql = "insert into querys (wquery) values('" & sql & "')"
'    cnEnvia.Execute csql
'    AlmacenaQuery_sql csql, cnEnvia
'
'    If cnEnvia.State = 1 Then cnEnvia.Close
'    Set cnEnvia = Nothing
'End If
'Exit Sub
'CapturaError:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atención"
'    Exit Sub
'End Sub



Public Sub GRABA_REGISTRO_logistica_DET(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String, parr_det() As Variant, pnumfilas As Integer, pvalores As String, pmes As String, pgraba_saldo As String)

Dim I           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
Dim nfila       As Integer
    
        
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For nfila = 0 To pnumfilas
        For I = 0 To pcantidad
            If Mid(pvalores, I + 1, 1) = "1" Then '--- EL CAMPO ESTA ACTIVO PARA GRABARLO
                If ptipo = "A" Then
                    If Len(ccampos) = 0 Then
                        ccampos = ccampos & parreglo(I).campo
                        If parreglo(I).Tipo = "T" Then
                            cvalores = "'" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            cvalores = parr_det(I, nfila)
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = "'" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = "'" & Format(parr_det(I, nfila), "yyyy-mm-dd") & "'"
                            Else
                                cvalores = "CVDATE('" & parr_det(I, nfila) & "')"
                            End If
                        End If
                    Else
                        ccampos = ccampos & "," & parreglo(I).campo
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & ",'" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            cvalores = cvalores & "," & parr_det(I, nfila) & ""
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & ",'" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & ",'" & Format(parr_det(I, nfila), "yyyy-mm-dd") & "'"
                            Else
                                cvalores = cvalores & ",CVDATE('" & parr_det(I, nfila) & "')"
                            End If
                        End If
                    End If
                End If
                If ptipo = "M" Then
                    If Len(Trim(cvalores)) = 0 Then
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & parreglo(I).campo & "='" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            cvalores = cvalores & parreglo(I).campo & "=" & parr_det(I, nfila) & ""
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & parreglo(I).campo & "='" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & parreglo(I).campo & "=" & Format(parr_det(I, nfila), "yyyy-mm-dd")
                            Else
                                cvalores = cvalores & parreglo(I).campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                            End If
                        End If
                    Else
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & "," & parreglo(I).campo & "='" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            cvalores = cvalores & "," & parreglo(I).campo & "=" & parr_det(I, nfila) & ""
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & "," & parreglo(I).campo & "=" & Format(parr_det(I, nfila), "yyyy-mm-dd")
                            Else
                                cvalores = cvalores & "," & parreglo(I).campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        If ptipo = "A" Then
            cvalores = Replace(cvalores, cnumvale, objAyudaVale.NumeroVale, 1)
            
            csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
            csql = csql & cvalores & ")"
            pconexion.Execute csql
        '    If pgraba_saldo = "*" Then '---- ACTUALIZA CANTIDAD Y PESO
        '        GRABA_SALDO parr_det(1, nfila), parr_det(4, nfila), parr_det(5, nfila), pmes, "I", pconexion
        '    End If
        '    If pgraba_saldo = "A" Then '---- ACTUALIZA CANTIDAD Y COSTO X ALMACEN
        '        GRABA_SALDO_ALM parr_det(1, nfila), Format(parr_det(2, nfila) & "", "0.000"), parr_det(5, nfila), pmes, wtipoguia, pconexion, parr_det(6, nfila), parr_det(10, nfila), "S"
        '    End If
        End If
        
        If ptipo = "M" Then
            csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
            pconexion.Execute csql
        End If
        AlmacenaQuery_sql csql, pconexion
        Actualiza_Log csql, pconexion.ConnectionString
        'AlmacenaQuery_sql csql, pconexion
        
        ccampos = "": cvalores = ""
    Next
    
    sw_GRABA_REGISTRO_logistica = True
End Sub

Public Sub GRABA_SALDO(pcodprod As Variant, pcantidad As Variant, ppeso As Variant, pmes As String, pdestino As String, pconexion As ADODB.Connection)
Dim csql    As String

    rsif5placli.Open "SELECT * FROM IF5PLACLI WHERE F5CODPRO = '" & pcodprod & "'", pconexion, adOpenDynamic, adLockOptimistic
    If Not rsif5placli.EOF Then
        If pdestino = "I" Then
            csql = "UPDATE IF5PLACLI SET F5DEBM" & pmes & " = F5DEBM" & pmes & " + " & Val("" & pcantidad) & ",F5INGPESO" & pmes & " =  F5INGPESO" & pmes & " + " & Val("" & ppeso) & " WHERE F5CODPRO = '" & pcodprod & "'"
        Else
            csql = "UPDATE IF5PLACLI SET F5DEBM" & pmes & " = F5DEBM" & pmes & " - " & Val("" & pcantidad) & ",F5INGPESO" & pmes & " =  F5INGPESO" & pmes & " - " & Val("" & ppeso) & " WHERE F5CODPRO = '" & pcodprod & "'"
        End If
        pconexion.Execute csql
        'AlmacenaQuery_sql csql, pconexion
    End If
    rsif5placli.Close

End Sub

Public Sub ELIMINA_BD_N(pruta As String, pnombase As String)
Dim DBName      As String

    DBName = pruta & "\" & pnombase
    'Kill DBName
    
End Sub

Public Sub GRABA_SALDO_ALM(pcodprod As Variant, pcantidad As Variant, psoles As Variant, pmes As String, pdestino As String, pconexion As ADODB.Connection, palmacen As Variant, pdolares As Variant, poperacion As String)
Dim csql    As String

    rsif6alma.Open "SELECT F5CODPRO FROM IF6ALMA WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'", pconexion, adOpenDynamic, adLockOptimistic
    If Not rsif6alma.EOF Then
        If pdestino = "I" Then
            If poperacion = "S" Then
                csql = "UPDATE IF6ALMA SET F5DEBM" & pmes & " = F5DEBM" & pmes & " + " & Val("" & pcantidad) & _
                ",F5ING" & pmes & " =  F5ING" & pmes & " + " & Val("" & psoles) & _
                ",F5INGD" & pmes & " =  F5INGD" & pmes & " + " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT + " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT + " & Val("" & pcantidad) & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                'AlmacenaQuery_sql csql, pconexion
                
            Else   '------ poperacion = "R"
                csql = "UPDATE IF6ALMA SET F5DEBM" & pmes & " = F5DEBM" & pmes & " - " & Val("" & pcantidad) & _
                ",F5ING" & pmes & " =  F5ING" & pmes & " - " & Val("" & psoles) & _
                ",F5INGD" & pmes & " =  F5INGD" & pmes & " - " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT - " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT - " & pcantidad & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                'AlmacenaQuery_sql csql, pconexion
            End If
        Else
            If poperacion = "S" Then
                csql = "UPDATE IF6ALMA SET F5HABM" & pmes & " = F5HABM" & pmes & " + " & Val("" & pcantidad) & _
                ",F5SAL" & pmes & " =  F5SAL" & pmes & " + " & Val("" & psoles) & _
                ",F5SALD" & pmes & " =  F5SALD" & pmes & " + " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT - " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT - " & pcantidad & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                'AlmacenaQuery_sql csql, pconexion
            Else   '------ poperacion = "R"
                csql = "UPDATE IF6ALMA SET F5HABM" & pmes & " = F5HABM" & pmes & " - " & Val("" & pcantidad) & _
                ",F5SAL" & pmes & " =  F5SAL" & pmes & " - " & Val("" & psoles) & _
                ",F5SALD" & pmes & " =  F5SALD" & pmes & " - " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT + " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT + " & Val("" & pcantidad) & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                'AlmacenaQuery_sql csql, pconexion
            End If
            
        End If
        pconexion.Execute csql
        'AlmacenaQuery_sql csql, pconexion
    End If
    rsif6alma.Close

End Sub

Public Function OBTIENE_CORRELA(pconexion As ADODB.Connection)
Dim ncorre      As Double

    RSCTA_DCTO.Open "SELECT CORRELA FROM CTA_DCTO ORDER BY CORRELA DESC", pconexion, adOpenDynamic, adLockOptimistic
    If Not RSCTA_DCTO.EOF Then
        RSCTA_DCTO.MoveFirst
        ncorre = RSCTA_DCTO.Fields("correla") + 1
    Else
        ncorre = 1
    End If
    RSCTA_DCTO.Close
    OBTIENE_CORRELA = ncorre

End Function

Public Sub ACUMULA_PRODUCTOS(prutatemp As String, pnombase As String, psql1 As String, psql2 As String, pconexion As ADODB.Connection, pnomtabla As String, pconexion_temp As ADODB.Connection, pcantidad As Double, ptipoalm As String, palmacen As String, PCodOri As String, pmedida As Double, ptime As String)
Dim rsdet       As New ADODB.Recordset
Dim rstemp      As New ADODB.Recordset
Dim rsgrupos    As New ADODB.Recordset
Dim CadSql      As String
Dim calmacen    As String

    If sw_creabd = True Then
        DELETEREC_LOG pnomtabla, pconexion_temp
        sw_creabd = False
    End If
        
    rsdet.Open psql2, pconexion, adOpenDynamic, adLockOptimistic
       
    If Not rsdet.EOF Then
        rsdet.MoveFirst
        Do While Not rsdet.EOF
            calmacen = ""
            If ptipoalm = "*" Then
                calmacen = Format(rsdet.Fields("F3GRUPOINS"), "00")
            Else
                calmacen = palmacen
            End If
            If pmedida > pcantidad Then
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & (Val(Format((pcantidad / pmedida) * Val(rsdet.Fields("F3CANTIDAD")), "0.00000"))) & ",'" & PCodOri & "')"
            ElseIf pmedida < pcantidad Then
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & (Val(Format((pcantidad * rsdet.Fields("F3CANTIDAD")) / pmedida, "0.00000"))) & ",'" & PCodOri & "')"
            Else
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & Val(Format(rsdet.Fields("F3CANTIDAD"), "0.000")) & ",'" & PCodOri & "')"
            End If
            pconexion_temp.Execute sql
            'AlmacenaQuery_sql sql, pconexion_temp
            rsdet.MoveNext
        Loop
    Else
        MsgBox "Producto " & PCodOri & " no tiene Receta", vbInformation, "AVISO"
        wFORMULA = 1
        'SQL = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value & "','" & calmacen & "','" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value & "', '" & Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value) & "','" & PCodOri & "')"
        'pconexion_temp.Execute SQL
        Exit Sub
    End If
      
    rsdet.Close

End Sub

Public Function valida(Tipo, car, Optional Texto, Optional mays)
Select Case Tipo
    Case 1  'Todo
        If Not IsMissing(mays) Then
            valida = Asc(UCase(Chr(car)))
        Else
            valida = car
        End If
    Case 2  'solo letras
        If Not ((car >= 65 And car <= 90) Or (car >= 97 And car <= 122) Or car = 8 Or car = 32) Then
            valida = 0
        Else
            If Not IsMissing(mays) Then
                valida = Asc(UCase(Chr(car)))
            Else
                valida = car
            End If
        End If
    Case 3  'numeros sin punto decimal
        If Not (car >= 48 And car <= 57 Or car = 8) Then
            valida = 0
        Else
            valida = car
        End If
    Case 4  'numeros con punto decimal
        If Not (car >= 48 And car <= 57 Or car = 8) Then
            If car = 46 Then
                existedec = InStr(1, Texto, ".")
                If existedec > 0 Then
                    valida = 0
                Else
                    valida = car
                End If
            Else
                valida = 0
            End If
        Else
            valida = car
        End If
End Select
End Function

'Public Function CalculaExistencia(almacen, prod, FECHA)
'Dim rst As New ADODB.Recordset
'Dim rstS As New ADODB.Recordset
'cad = ""
'If Not (almacen = "") Then
'    cad = " and f3vales.f2codalm='" & almacen & "'"
'End If
'
'If ctipoadm_bd = "A" Then
''    'SQL = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr " _
''    '& "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) " _
''    '& "WHERE CVDATE(IF4VALES.F4FECVAL)<=CVDATE('" & fecha & "') And f5codpro='" & prod & "'and f3vales.f2codalm='" & almacen & "'" ' & cad
'    SQL = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) Where CVDATE(if3vales.F4FECVAL)<=CVDATE('" & FECHA & "') AND if3vales.F5CODPRO ='" & Trim(prod) & "' and if3vales.F2CODALM='" & almacen & "'"
'Else
''   ' SQL = "SELECT Sum(If(Left(IF4VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0)) AS ING, Sum(Iff(Left(IF4VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0)) AS EGR "
''   ' SQL = SQL & "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) "
''   ' SQL = SQL & "WHERE IF4VALES.F4FECVAL<='" & fecha & "' And f5codpro='" & prod & "'" & cad
'
'    SQL = "SELECT Left([IF3VALES].[F4NUMVAL],1)='I' AS ING1, IF3VALES.F2CODALM, IF3VALES.F5CODPRO, Sum(IF3VALES.F3CANPRO) AS SumaDeF3CANPRO"
'    SQL = SQL + " FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
'   SQL = SQL + " GROUP BY Left([IF3VALES].[F4NUMVAL],1)='I', IF3VALES.F2CODALM, IF3VALES.F5CODPRO"
'    SQL = SQL + " HAVING (((IF3VALES.F2CODALM)='" & almacen & "') AND ((IF3VALES.F5CODPRO)='" & prod & "'))"
'
'End If
'
'If rst.State = adStateOpen Then rst.Close
'rst.Open UCase(SQL), cnn_dbbancos, adOpenStatic, adLockOptimistic
'
'If Not rst.EOF Then
'    nstock = Val("" & rst!ing1) - Val("" & rst!Egr)
'    CalculaExistencia = nstock
'Else
'    CalculaExistencia = -1
'End If
'rst.Close
'End Function
Public Function CalculaExistencia(almacen, prod, Fecha)
Dim rst As New ADODB.Recordset

cad = ""
If Not (almacen = "") Then
    cad = " and if3vales.f2codalm='" & almacen & "'"
End If
If ctipoadm_bd = "M" Then
    sql = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) Where CVDATE(if3vales.F4FECVAL)<=CVDATE('" & Fecha & "') AND if3vales.F5CODPRO ='" & Trim(prod) & "' and if3vales.F2CODALM='" & almacen & "'"
Else
    sql = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr " _
    & "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) " _
    & "WHERE CVDATE(IF4VALES.F4FECVAL)<=CVDATE('" & Fecha & "') And f5codpro='" & prod & "'" & cad
End If
If rst.State = adStateOpen Then rst.Close
rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    nstock = Val("" & rst("ing")) - Val("" & rst("egr"))
    CalculaExistencia = nstock
Else
    CalculaExistencia = -1
End If
rst.Close
End Function


Public Function VALIDA_FPAGO(pfpago As String)
Dim sw      As Boolean

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select F2DESPAG from ef2forpag where f2forpag='" & Trim(pfpago) & "'", cnn_dbbancos  'cnn_dbbancos
    If rst.EOF = False Then
        wnompag = rst!F2DESPAG & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_FPAGO = sw
        
End Function

Public Function VALIDA_CLIENTE(pcodcli As String)
Dim sw      As Boolean
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where f2codcli='" & pcodcli & "' OR f2NEWRUC='" & pcodcli & "'", cnn_dbbancos
    If rst.EOF = False Then
        wcodcli = "" & rst!F2CODCLI
        wnomcli = "" & rst!F2nomcli
        wruccli = "" & rst!F2NEWRUC
        WDIRCLI = "" & rst!F2DIRCLI
        wforpag = "" & rst!F2FORPAG
        nnumlista = Val("" & rst.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CLIENTE = sw
        
End Function

Public Function VALIDA_CLIENTE_V(pcodcli As String, pcodven As String)
Dim sw      As Boolean
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where (f2codcli='" & pcodcli & "' OR f2newruc='" & pcodcli & "') AND (F2CODVEN='" & pcodven & "')", cnn_dbbancos
    If rst.EOF = False Then
        wcodcli = "" & rst!F2CODCLI
        wnomcli = "" & rst!F2nomcli
        wruccli = "" & rst!F2NEWRUC
        WDIRCLI = "" & rst!F2DIRCLI
        wforpag = "" & rst!F2FORPAG
        nnumlista = Val(rst.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CLIENTE_V = sw
        
End Function

Public Function VALIDA_RESPONSABLE(pcodres As String)
Dim sw      As Boolean

    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from EF2VENDEDORES where F2CODVEN='" & pcodres & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        wnomven = rs.Fields("F2NOMVEN") & ""
        wfacmin_nac = Val(Format("" & rs.Fields("F2FACNAC_MIN"), "0.00"))
        wfacmin_imp = Val(Format("" & rs.Fields("F2FACIMP_MIN"), "0.00"))
        sw = True
    Else
        sw = False
    End If
    rs.Close
    VALIDA_RESPONSABLE = sw
        
End Function

Public Function VALIDA_ORIGEN(PCodOri As String)
Dim sw      As Boolean
Set rst = New ADODB.Recordset
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from Sf1Origenes where F1CodOri='" & Trim(PCodOri) & "'", cnn_dbbancos
    If Not rst.EOF Then
        wcodori = Trim(rst!F1CODORI & "")
        wnomori = Trim(rst!F1NOMORI & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_ORIGEN = sw

End Function
Public Function VALIDA_ALMACENO(pcodalm As String)
Dim sw  As Boolean
Set rs = New ADODB.Recordset

    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select F2NOMALM from ef2almacenes where f2codalm='" & Trim(pcodalm) & "'", cnn_dbbancos
    If rs.EOF = False Then
        WNomPar = rs!F2NOMALM & ""
        sw = True
    Else
        sw = False
    End If
    rs.Close
    VALIDA_ALMACENO = sw
        
End Function

Public Sub GRABAR_ACTUALIZACIONES(Fmes As Integer, Alma As String, cod As String)
Dim amovs_arr(0 To 7) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).Tipo = "N"
amovs_arr(5).campo = "F6FECULT": amovs_arr(5).valor = FecUlt: amovs_arr(5).Tipo = "F"
amovs_arr(6).campo = "F5COSPRO": amovs_arr(6).valor = Cospro: amovs_arr(6).Tipo = "N"
amovs_arr(7).campo = "F5COSPROD": amovs_arr(7).valor = Cosprod: amovs_arr(7).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 7, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Public Sub GRABAR_ACTUALIZACIONES1(Fmes As Integer, Alma As String, cod As String)

Dim amovs_arr(0 To 6) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6FECULT": amovs_arr(4).valor = CVDate(FecUlt): amovs_arr(4).Tipo = "F"
amovs_arr(5).campo = "F5COSPRO": amovs_arr(5).valor = Cospro: amovs_arr(5).Tipo = "N"
amovs_arr(6).campo = "F5COSPROD": amovs_arr(6).valor = Cosprod: amovs_arr(6).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 6, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub Actualizar_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim csql, SSQL As String
    
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset
'''''''''''CAMBIOS EN SQL
SSQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    Debm = 0#: stockact = 0#: ing = 0#: ingd = 0#
    Habm = 0#: StockLog = 0#: sal = 0#: sald = 0#
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")) - psoles
        ingd = RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
    '''''''''CORREGIR MES
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")) - psoles
        sald = RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro
    End If

End If
End Sub

Sub GRABAR_ACTUALIZAALMA(Fmes As String, Alma As String, cod As String)
Dim amovs_arr(0 To 3) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 3, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub GRABAR_ACTUALIZAALMA1(Fmes As String, Alma As String, cod As String)
Dim amovs_arr(0 To 4) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).Tipo = "N"


'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 4, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub ImprimendoV(pcodalm As String, PNUMVAL As String, pcosto As Integer)
    
Dim ITEM, Fila As Integer
    
ITEM = 1
Fila = 1

Set RsStockCab = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset
Set RsProducto = New ADODB.Recordset
Set RSCONSULTA = New ADODB.Recordset

If RsStockCab.State = adStateOpen Then RsStockCab.Close
RsStockCab.Open "SELECT * FROM IF4VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockCab.EOF Then
    
    Printer.ScaleMode = 4
    TituloV PNUMVAL

    If RsStockDet.State = adStateOpen Then RsStockDet.Close
    RsStockDet.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockDet.EOF Then
        Fila = Fila + 6
        Do While Not RsStockDet.EOF
            If RsProducto.State = adStateOpen Then RsProducto.Close
            RsProducto.Open "SELECT * FROM IF5PLA WHERE F5CODPRO ='" & RsStockDet.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsProducto.EOF Then
                writexy Trim(RsStockDet.Fields("F5codpro")), Fila, 1, 0
                writexy left(RsProducto.Fields("F5nompro"), 65), Fila, 11, 0
                writexy Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), Fila, 63, 2
                writexy Trim("" & RsProducto.Fields("f7codmed")), Fila, 80, 0
                Fila = Fila + 1
                If RsProducto.Fields("f5series") = "1" Then
                    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
                    sql = "select * from if3series where f2codalm='" & pcodalm & "' and f4numval='" & PNUMVAL & "' and f5codpro='" & Trim(RsProducto.Fields("f5codpro")) & "' order by f3numser"
                    RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not RSCONSULTA.EOF Then
                        RSCONSULTA.MoveFirst
                        Do While Not RSCONSULTA.EOF
                            writexy ITEM & ".-  E.S.N. ==>" & RSCONSULTA.Fields("F3numser"), Fila, 16, 0
                            Fila = Fila + 1
                            RSCONSULTA.MoveNext
                            If Fila >= 60 Then
                                Printer.NewPage
                                TituloV PNUMVAL
                                Fila = 18
                            End If
                            ITEM = ITEM + 1
                        Loop
                    End If
                    Fila = Fila + 1
                End If
                ITEM = 1
                If Fila >= 60 Then
                   Printer.NewPage
                   TituloV PNUMVAL
                   Fila = 18
                End If
                RsProducto.MoveNext
            End If
            RsStockDet.MoveNext
        Loop
    End If

    Printer.Line (60, Fila + 6)-(88, Fila + 6)
    writexy "Almacén", Fila + 6, 68, 0
    
    Printer.Line (5, Fila + 6)-(33, Fila + 6)
    writexy "Solicitado Por", Fila + 6, 12, 0
    Printer.EndDoc
    
End If

End Sub

Sub TituloV(PNUMVAL As String)
Dim Fila As Integer
Set RsAlmacenes = New ADODB.Recordset
Set RsProveedor = New ADODB.Recordset
Set RsOrigen = New ADODB.Recordset
Set RsPartida = New ADODB.Recordset
    

    If RsOrigen.State = adStateOpen Then RsOrigen.Close
    RsOrigen.Open "SELECT * FROM SF1ORIGENES WHERE F1CODORI = '" & RsStockCab.Fields("F1CODORI") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    CabeceraV
    Printer.FontBold = True
    Printer.FontSize = 15
    writexy IIf(RsOrigen.Fields("F1TIPMOV") = "I", "VALE DE INGRESO", "VALE DE SALIDA"), 6, 34, 0
    Printer.FontBold = False
    Printer.FontSize = 9
    
    Fila = 11
    Printer.FontBold = True

    If Trim(nomcentro) <> "" Then
        writexy "C.COSTO:", 10, 5, 0
        writexy Trim(nomcentro), 10, 15, 0
    End If

    writexy "Nº VALE:", Fila, 5, 0
    writexy PNUMVAL, Fila, 15, 0
    
    writexy "Almacén:", Fila, 60, 0
    
    If Not IsNull(RsStockCab.Fields("F2codalm")) Then
        If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
        RsAlmacenes.Open "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & RsStockCab.Fields("F2CODALM") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsAlmacenes.EOF Then
            writexy Trim(RsAlmacenes.Fields("f2codalm") & " - " & Mid(RsAlmacenes.Fields("F2nomalm"), 1, 21)), Fila, 70, 0
        End If
        
    End If

    Fila = Fila + 1
    
    If Not IsNull(RsStockCab.Fields("F2codprov")) Then
        If RsProveedor.State = adStateOpen Then RsProveedor.Close
        RsProveedor.Open "SELECT * FROM EF2PROVEEDORES WHERE F2NEWRUC = '" & RsStockCab.Fields("F2CODPROV") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProveedor.EOF Then
            writexy "Proveedor:", Fila, 5, 0
            writexy Trim(RsProveedor.Fields("f2codprov") & " - " & RsProveedor.Fields("F2nomprov")), Fila, 15, 0
        End If
        
    End If
    
    If Not IsNull(RsStockCab.Fields("F2codpar")) Then
        If RsStockCab.Fields("f1codori") = "XT0" Or RsStockCab.Fields("f1codori") = "XT1" Then
            If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
            RsAlmacenes.Open "SELECT * FROM EF2ALMACENES", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsAlmacenes.Find "F2codalm = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsAlmacenes.EOF Then
                If RsStockCab.Fields("f1codori") = "XT0" Then
                   writexy "ORIGEN  : ", Fila, 60, 0
                Else
                   writexy "DESTINO : ", Fila, 60, 0
                End If
                writexy UCase(RsAlmacenes.Fields("f2codalm") & "-" & RsAlmacenes.Fields("f2nomalm")), Fila, 70, 0
            End If
        Else
            
            If RsPartida.State = adStateOpen Then RsPartida.Close
            RsPartida.Open "SELECT * FROM EF2PARTIDAS", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsPartida.Find "F2codpar = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsPartida.EOF Then
                writexy "DESTINO : ", Fila, 60, 0
                writexy UCase(RsPartida.Fields("f2nompar")), Fila, 70, 0
            End If
        End If
    Else
        If Val("" & RsStockCab.Fields("F4numord")) > 0 Then
            writexy "O/Compra : ", Fila, 60, 0
            writexy Format(RsStockCab.Fields("F4numord"), "000000"), Fila, 70, 0
        End If
    End If


    Printer.FontBold = False
    writexy "Concepto:", Fila + 1, 5, 0
    writexy Trim(RsOrigen.Fields("F1nomori")), Fila + 1, 15, 0

    writexy "Documento:", Fila + 1, 60, 0
    writexy Trim(RsStockCab.Fields("F1coddoc") & " (" & RsStockCab.Fields("F4numdoc") & ") "), Fila + 1, 70, 0
    
    writexy "Fecha:", Fila + 2, 5, 0
    writexy Trim(Trim(cmes(Month(CVDate(RsStockCab.Fields("F4fecval"))))) & " " & Format(Day(CVDate(RsStockCab.Fields("F4fecval"))), "##") & ", de " & Year(CVDate(RsStockCab.Fields("F4fecval")))), Fila + 2, 15, 0
    
    Printer.FontBold = True
    Printer.Line (1, Fila + 3)-(90, Fila + 3)
    writexy "Código", Fila + 4, 1, 0
    writexy "Artículo", Fila + 4, 11, 0
    writexy "Cantidad", Fila + 4, 63, 0
    writexy "Uni.", Fila + 4, 80, 0
    Printer.Line (1, Fila + 5)-(90, Fila + 5)
    Printer.FontBold = False
    
End Sub

Sub CabeceraV()
Dim csql As String
Set RsParain = New ADODB.Recordset

csql = "SELECT * FROM SF1PARAIN WHERE F1CODEMP = '" & wempresa & "'"
If RsParain.State = adStateOpen Then RsParain.Close
RsParain.Open csql, cnn_control, adOpenDynamic, adLockOptimistic
  
Printer.ScaleMode = 4
Printer.FontName = "Courier New" 'Printer.Fonts (RsParain.Fields("F1FONNAM"))
Printer.FontSize = 12
Printer.FontBold = True
writexy Trim("" & RsParain.Fields("f1nomemp")), 2, 1, 0
Printer.FontSize = 8
Printer.FontBold = False

writexy "Fecha: ", 2, 80, 0
writexy Format(Now, "dd/mm/yyyy"), 2, 88, 0
writexy "Intersys - Inventario", 3, 1, 0
writexy "Página: ", 3, 80, 0
writexy Format(Printer.Page, "###00"), 3, 88, 0

End Sub

Sub Reactualiza_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim SSQL, csql As String
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset

stockact = 0#: Debm = 0#: ing = 0#: ingd = 0#
StockLog = 0#: Habm = 0#: sal = 0#: sald = 0#
sql = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    SSQL = "SELECT F5STOCKACT FROM IF5PLA WHERE F5CODPRO = '" & pcodpro & "'"
    If RsProducto.State = adStateOpen Then RsProducto.Close
    RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProducto.EOF Then
       If ptipmov = "S" Then
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) + pcanpro
       Else
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) - pcanpro
       End If
       csql = "UPDATE IF5PLA SET F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
       cnn_dbbancos.Execute (csql)
       'AlmacenaQuery_sql csql, cnn_dbbancos
       
    End If
    RsProducto.Close
    
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = Val(Format(RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")), "0.00")) - psoles
        ingd = Val(Format(RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = Val(Format(RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")), "0.00")) - psoles
        sald = Val(Format(RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro

    End If
    
End If
End Sub

Public Sub BASE_TEMPORAL(Base As String)
Dim con As String
Set Temp = New ADODB.Connection

con = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
Temp.Open con

End Sub
Public Function VALIDA_PROVEEDOR(pcodprov As String)
Dim RsProveedor As New ADODB.Recordset
On Error Resume Next
Dim sw      As Boolean
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT F2CODPROV, F2NOMPROV, F2NEWRUC " & _
                "FROM EF2PROVEEDORES " & _
                "where f2codprov='" & pcodprov & "' OR F2NEWRUC='" & pcodprov & "'", cnn_dbbancos, 3, 1
    If rs.EOF = False Then
        wcodprov = "" & rs!F2CODPROV
        wnomprov = "" & rs!F2NOMPROV
        wrucprov = "" & rs!F2NEWRUC
        sw = True
    Else
        sw = False
        wcodprov = ""
        wnomprov = ""
        wrucprov = ""
    End If
    rs.Close
    VALIDA_PROVEEDOR = sw
        
End Function

Public Function ObtenerCampo(tabla As String, campo As String, campoCom As String, valor As String, TipoDeComparacion As String, ConexionDeBaseDeDatos As ADODB.Connection) As String
Dim cad As String
Dim rst As New Recordset
If TipoDeComparacion = "F" Then
    cad = "select " & campo & " from " & tabla & " where CVDATE(" & campoCom & ") = '" & valor & "' " & condicion
ElseIf TipoDeComparacion = "T" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
ElseIf TipoDeComparacion = "N" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = " & valor & " " & condicion
End If
If rst.State = 1 Then rst.Close
rst.Open cad, ConexionDeBaseDeDatos, adOpenForwardOnly, adLockReadOnly
ObtenerCampo = ""
If Not rst.EOF And Not IsNull(rst.Fields(0)) Then ObtenerCampo = rst.Fields(0)
End Function


Public Function SeleccionaEnComboLeft(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For I = 0 To NombreDelCombo.ListCount - 1
    If right(NombreDelCombo.List(I), Len(Trim(DatoBuscado))) = DatoBuscado Then
        NombreDelCombo.ListIndex = I
        Exit For
    Else
    End If
Next
End Function

Public Function SeleccionaEnComboRight(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For I = 0 To NombreDelCombo.ListCount - 1
    If UCase(right(NombreDelCombo.List(I), Len(Trim(DatoBuscado)))) = UCase(DatoBuscado) Then
        NombreDelCombo.ListIndex = I
        Exit For
    Else
    End If
Next
End Function

Public Function SeleccionaEnCombo(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For I = 0 To NombreDelCombo.ListCount - 1
    If left(NombreDelCombo.List(I), Len(Trim(DatoBuscado))) = DatoBuscado Then
        NombreDelCombo.ListIndex = I
        Exit For
    Else
    End If
Next
End Function

Public Function CargaRsT(ByVal SentenciaSql As String, ByVal ConexionUsada As ADODB.Connection) As ADODB.Recordset
Dim rst As New ADODB.Recordset
If rst.State = 1 Then rst.Close
rst.Open SentenciaSql, ConexionUsada, 3, 1
Set CargaRsT_DbTabla = rst
End Function
Public Function VALIDA_VENDEDOR(pvendedor As String)
Dim sw      As Boolean
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2USERS where F2CODUSER='" & Trim(pvendedor) & "' AND F2VENDEDOR='*'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomven = rst!F2NOMUSER & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_VENDEDOR = sw
End Function

Public Function VALIDA_COBRADOR(pvendedor As String)
Dim sw      As Boolean

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2USERS where F2CODUSER='" & Trim(pvendedor) & "' AND F2COBRADOR='*'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomven = rst!F2NOMUSER & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_COBRADOR = sw
End Function

Public Function VALIDA_CANAL(pcanal As String)
Dim sw      As Boolean
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2CANALES where F2CANCODIGO='" & Trim(pcanal) & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomcanal = rst!F2CANNOMBRE & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CANAL = sw
End Function

Public Function FileExist(ByVal sFileName As String) As Boolean
Dim hFile As Integer

sFileName = Trim$(sFileName)

FileExist = False

If sFileName = "" Then Exit Function

On Error Resume Next
hFile = FreeFile
Open sFileName For Input Access Read Shared As #hFile
If Err.Number = 0 Then FileExist = True
Close #hFile

End Function

Public Function Repetir(ByVal NumeroDeVeces As Integer, ByVal CaracterRepetido As String) As String
Repetir = ""
For I = 1 To NumeroDeVeces
    Repetir = Repetir & CaracterRepetido
Next
End Function



Public Sub Crea_Campo(pCadenaConexion As String, pNombreDeTabla As String, pNombreDeCampo As String, pTipoCampo As String, pEsNull As Boolean, pValorPorDefecto As String)
On Error GoTo CapturaError
Dim StrContenido As String
Dim SwExiste As Boolean
Dim pAf As New ADOFunctions
Dim pRs As New ADODB.Recordset
Dim SqlCad As String, I As Integer
SwExiste = False
If (FileExist(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl") = False) Then
    SqlCad = "Select * from " & pNombreDeTabla
    Set pRs = pAf.OpenSQLForwardOnly(SqlCad, pCadenaConexion)
    If pRs.State = 1 Then
        For I = 0 To (pRs.Fields.Count - 1)
            If UCase(pRs.Fields(I).Name) = UCase(pNombreDeCampo) Then
                SwExiste = True
                Exit For
            End If
        Next
    End If
Else
    'StrContenido = sGetINI(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", "l")
    If IsDate(StrContenido) Then
        SwExiste = True
    Else
        SwExiste = False
    End If
End If
If pRs.State = 1 Then pRs.Close
Set pRs = Nothing
If SwExiste = False Then
    csql = "Alter table " & pNombreDeTabla & " ADD COLUMN " & pNombreDeCampo & " " & pTipoCampo & IIf(pEsNull = True, " NULL ", " NOT NULL ") & IIf(Len(Trim(pValorPorDefecto)) > 0, " DEFAULT " & pValorPorDefecto, "")
    Call EJECUTA_SENTENCIA(csql, pCadenaConexion)
    'If Sw_Ejecuta_Sentencia = True Then
    sWrtIni wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", str(Date)
    'End If
Else
    If (FileExist(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl") = False) Then
        sWrtIni wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", str(Date)
    End If
End If

Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
End Sub

Public Sub EJECUTA_SENTENCIA(pSentencia As String, pconexion As String)
On Error GoTo Error_GRABA_REGISTRO_logistica

Dim CnExec As New ADODB.Connection
Sw_Ejecuta_Sentencia = False
CnExec.Open pconexion
CnExec.Execute pSentencia
If CnExec.State = 1 Then CnExec.Close
Set CnExec = Nothing
Sw_Ejecuta_Sentencia = True
Exit Sub

Error_GRABA_REGISTRO_logistica:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Select Case Err.Number
    Case 3704
        If CnExec.State = 0 Then CnExec.Open
        Resume
    Case Else
        Sw_Ejecuta_Sentencia = False
        Exit Sub
    End Select
End Sub

Public Function VerificaPermiso(Codigo_de_Permiso As String, Nombre_de_Usuario As String) As Boolean
    Dim Af As New ADOFunctions
    Dim rs As New ADODB.Recordset
    
    csql = "select * from EF2TAREAUSERS where f2coduser='" & Nombre_de_Usuario & "' and f2codtarea='" & Codigo_de_Permiso & "'"
    
    Set rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    If rs.RecordCount > 0 Then
        VerificaPermiso = True
    Else
        VerificaPermiso = False
    End If
    
    If rs.State = 1 Then rs.Close
    
    Set rs = Nothing
End Function

Public Property Get ComputerName() As String
    Dim sName As String
    Dim lRetval As Long
    Dim iPos As Integer
    
    sName = Space$(255)
    lRetval = GetComputerName(sName, 255)
    iPos = InStr(sName, Chr$(0))
    ComputerName = left$(sName, iPos - 1)
End Property

Public Function VerificaAutorizaciones(Codigo_de_Consulta As String, Nombre_de_Usuario As String) As String
    Dim Af As New ADOFunctions
    Dim I As Integer
    Dim rs As New ADODB.Recordset
    
    'Crea_Campo cconex_dbbancos, "EF2AUTORIZADOS", "F2CODUSER", "String", True, ""
    I = 0
    
    csql = "select F3COSTO from EF2AUTORIZADOS where F2CODUSER='" & Nombre_de_Usuario & "' and F2REPORTE='" & Codigo_de_Consulta & "'"
    
    Set rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    If rs.State = 1 Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            
            Do While Not rs.EOF
                I = I + 1
                
                If I = 1 Then
                    VerificaAutorizaciones = "'" & rs!F3COSTO & "'"
                Else
                    VerificaAutorizaciones = VerificaAutorizaciones & ",'" & rs!F3COSTO & "'"
                End If
                
                rs.MoveNext
            Loop
        Else
            VerificaAutorizaciones = "''"
        End If
    End If
    
    If rs.State = 1 Then rs.Close
    
    Set rs = Nothing
End Function

'Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
'Dim sTemp As String * 256
'Dim nLength As Integer
'sTemp = Space$(256)
'nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
'sGetINI = left$(sTemp, nLength)
'End Function

Public Sub sWrtIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Guarda los datos de configuración
    'Los parámetros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
Public Function enviacorreoGmail(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    
    .ServerName = Trim(OriginDominio)
    '.ServerPort = CLng(Val(995))
    .UserName = Trim(OriginMail)
    .Password = Trim(OriginPassword)
    .RelayServer = "smtp.gmail.com"
    .RelayPort = 465
    .Secure = True
    .Options = 5

    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "betania_bk@outlook.com.pe", _
                                          "jparedes@betania.com.pe", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          strMessageHTML, _
                                          2, _
                                          1)
    
    If nError Then
    Exit Function
    End If
    .Priority = 3
'    If nError Then
'        MsgBox ("Unable to create a new message " & .LastErrorString)
'        Exit Function
'    End If
'
'    If .Recipients = 0 Then
'        MsgBox ("There are no recipients for this message")
'        Exit Function
'    End If
    nError = .SendMessage()

    If nError Then
        enviacorreoGmail = False
        MsgBox ("Correo no pudo ser enviado")
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreoGmail = True
    
    End If
End With
End Function

Public Function enviacorreoPOP(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long

 With MailObj
    
''    .ServerName = Trim(OriginDominio)
''    .ServerPort = CLng(Val(110))
''    .UserName = Trim(OriginMail)
''    .Password = Trim(OriginPassword)
''    .Timeout = CLng(Val(60))
''
''    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
''                                          DestinoMail, _
''                                          "betania_bk@outlook.com.pe, psalas@betania.com.pe", _
''                                          "responder.britania@gmail.com", _
''                                          DestinoAsunto, _
''                                          DestinoCuerpo, _
''                                          strMessageHTML, _
''                                          2, _
''                                          1)
''
''    If nError Then
''    Exit Function
''    End If
''    .Priority = 3
''    nError = .SendMessage()
''
''    If nError Then
''        enviacorreoPOP = False
''        MsgBox ("Correo no pudo ser enviado")
''    Else
''        MsgBox ("Correo enviado con éxito")
''        enviacorreoPOP = True
''    End If

    
    'nError = CreateMessage()
    nError = .ComposeMessage(OriginMail, _
                             DestinoMail, _
                             "betania_bk@outlook.com.pe", _
                             "responder.britania@gmail.com", _
                             DestinoAsunto, _
                             DestinoCuerpo, _
                             strMessageHTML, _
                             2, _
                             1)
    
    If nError Then
        MsgBox ("Correo no pudo ser creado")
    End If
    
        .RelayServer = ""
        .RelayPort = 0
    nError = .SendMessage()
    If nError Then
        enviacorreoPOP = False
        MsgBox ("Correo no pudo ser enviado")
        Sw_Act = False
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreoPOP = True
        Sw_Act = True
    End If
End With
End Function

Private Function CreateMessage() As Long
    Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long
        
    CreateMessage = 0
    
    strMessageHTML = ""
    
    '
    nError = InternetMail1.ComposeMessage(editFrom.Text, _
                                          editTo.Text, _
                                          editCc.Text, _
                                          editBcc.Text, _
                                          editSubject.Text, _
                                          editMessageText.Text, _
                                          strMessageHTML, _
                                          nCharacterSet, _
                                          1)
    
    If nError Then
        CreateMessage = nError
        Exit Function
    End If
    
    
End Function

'Public Function enviacorreoPOPa(ByVal MailObj As Object, _
'ByVal OriginDominio As String, _
'ByVal OriginNombre As String, _
'ByVal OriginMail As String, _
'ByVal OriginPassword As String, _
'ByVal DestinoMail As String, _
'ByVal DestinoAsunto As String, _
'ByVal DestinoCuerpo As String) As Boolean
'
'Dim nError As Long
' With MailObj
'
''    .ServerName = Trim(OriginDominio)
''    '.ServerPort = CLng(Val(995))
''    .UserName = Trim(OriginMail)
''    .Password = Trim(OriginPassword)
''    .RelayServer = "smtp.gmail.com"
''    .RelayPort = 465
''    .Secure = True
''    .Options = 5
'
'    .ServerName = Trim(OriginDominio)
'    .ServerPort = CLng(Val(110))
'    .UserName = Trim(OriginMail)
'    .Password = Trim(OriginPassword)
'    .Timeout = CLng(Val(60))
'
'    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
'                                          DestinoMail, _
'                                          "betania_bk@outlook.com.pe, psalas@betania.com.pe", _
'                                          "responder.britania@gmail.com", _
'                                          DestinoAsunto, _
'                                          DestinoCuerpo, _
'                                          strMessageHTML, _
'                                          2, _
'                                          1)
'
'    If nError Then
'    Exit Function
'    End If
'    .Priority = 3
''    If nError Then
''        MsgBox ("Unable to create a new message " & .LastErrorString)
''        Exit Function
''    End If
''
''    If .Recipients = 0 Then
''        MsgBox ("There are no recipients for this message")
''        Exit Function
''    End If
'    nError = .SendMessage()
'
'    If nError Then
'        enviacorreoPOP = False
'        MsgBox ("Correo no pudo ser enviado")
'    Else
'        MsgBox ("Correo enviado con éxito")
'        enviacorreoPOP = True
'    End If
'End With
'End Function
'
'
'
Public Function enviacorreohotmail(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    
    .ServerName = Trim(OriginDominio)
'    .ServerPort = CLng(Val(995))
    .UserName = Trim("OriginMail")
    .Password = Trim("OriginPassword")
    .RelayServer = "smtp.live.com"
    .RelayPort = 587
    '.Secure = True
    .Options = 8


    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "responder.britania@gmail.com", _
                                          "jparedes@betania.com.pe", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          strMessageHTML, _
                                          2, _
                                          1)
    
    If nError Then
    Exit Function
    End If
    .Priority = 3
'    If nError Then
'        MsgBox ("Unable to create a new message " & .LastErrorString)
'        Exit Function
'    End If
'
'    If .Recipients = 0 Then
'        MsgBox ("There are no recipients for this message")
'        Exit Function
'    End If
    nError = .SendMessage()

    If nError Then
        enviacorreohotmail = False
        MsgBox ("Correo no pudo ser enviado")
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreohotmail = True
    End If
End With
End Function


Public Function enviacorreo(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    .ServerName = Trim(OriginDominio)
    .ServerPort = CLng(Val(110))
    .UserName = Trim(OriginMail)
    .Password = Trim(OriginPassword)
    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "", _
                                          "", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          "", _
                                          2, _
                                          1)
    If nError Then
    Exit Function
    End If
    .Priority = 3
    If nError Then
        MsgBox ("Unable to create a new message " & .LastErrorString)
        Exit Function
    End If
    
    If .Recipients = 0 Then
        MsgBox ("There are no recipients for this message")
        Exit Function
    End If
    nError = .SendMessage()
    
    If nError Then
        enviacorreo = False
    Else
        enviacorreo = True
    End If
End With
End Function













