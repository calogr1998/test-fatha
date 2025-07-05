Attribute VB_Name = "Transferencia"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public res_pdf As Long
Private Const ODBC_ADD_DSN = 1 ' Nuevo DSN
Private Const ODBC_CONFIG_DSN = 2 ' Modificar DSN
Private Const ODBC_REMOVE_DSN = 3 ' Eliminar DSN
Private Const ODBC_ADD_SYS_DSN = 4 ' Nuevo DSN de sistema
Private Const ODBC_CONFIG_SYS_DSN = 5 ' Modificar DSN de sistema
Private Const ODBC_REMOVE_SYS_DSN = 6 ' Eliminar DSN de sistema
Private Const vbAPINull As Long = 0 ' Null Pointer
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_FETCH_NEXT As Long = 1

'Declaración de funciones de API
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Long, ByVal fDirection As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, pcbDescription As Integer) As Integer
Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (Env As Long) As Integer

'Constantes
'ruta hasta el servidor (ip/nombre/ruta)
Private Const C_Server = "190.232.93.212"
'usuario
Private Const C_User = "root"
'contraseña
Private Const C_Pass = "infoplus123456"
'base de datos
Private Const C_BD = "test"
'puerto
Private Const C_Port = 3306
'Nombre ODBC de MySql
'(si no tienes ninguno instalado http://dev.mysql.com/downloads/connector/odbc/5.1.html)
Public Const C_MYSQL_ODBC = "MySQL ODBC 5.1 Driver"
Public TCCompra      As String
Public TCVenta     As String
Public ruc_rsocial      As String
Public ruc_direccion    As String
Public ruc_estado       As String
Public ruc_situacion    As String
Public ruc_telefono     As String
Public strODBCtablas As String
Public strTABLAempant As String
Public strODBCsiete As String
Public strODBCsgte As String

Public envia As String


Public Sub AlmacenaQuery_sql(ByVal sql As String, ConeccionAdodb As ADODB.Connection, Optional CentroCC As String)
'Dim cnEnvia As New ADODB.Connection
'If InStr(UCase(ConeccionAdodb), "DB_BANCOS.MDB") > 0 Or InStr(UCase(ConeccionAdodb), "DB_TABLA.MDB") > 0 Then
'    If cnEnvia.State = 1 Then cnEnvia.Close
'    cnEnvia.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_envia.mdb;Persist Security Info=False"
'    sql = Replace(sql, "'", "|")
'    cnEnvia.Execute "insert into querys_pedidos (wquery) values('" & sql & "')"
'End If
'
'If cnEnvia.State = 1 Then cnEnvia.Close
'
'Set cnEnvia = Nothing
''Enviar_Querys
End Sub

Public Sub Enviar_Querys()

'Dim cnn_Envia_Mysql As New ADODB.Connection
'Dim cnn_Envia As New ADODB.Connection
'Dim RsM As New ADODB.Recordset
'Dim cn As String
'Dim Rs As New ADODB.Recordset
'Dim od As String
'
'On Error GoTo error_conexion
'    'strODBC = traerCampo("srutas", "TABLAS", "EMPRESA", wempresa)
'    'envia = traerCampo("srutas", "SIETE", "EMPRESA", wempresa)
'    od = Mid(strODBCtablas, 28, 30)
'    If Len(Trim(strODBCsiete)) = 0 Then Exit Sub
'    If cnn_Envia_Mysql.State = 1 Then cnn_Envia_Mysql.Close
'    If strODBCtablas = "ACCESS" Then
'        StrConexDbEnvia = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_Envia.MDB;Persist Security Info=False"
'        If cnn_Envia_Mysql.State = 1 Then cnn_Envia_Mysql.Close
'        cnn_Envia_Mysql.Open StrConexDbEnvia
'    Else
'        cnn_Envia_Mysql.Open strODBCtablas
'    End If
'
'    StrConexDbEnvia = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_Envia.MDB;Persist Security Info=False"
'    If cnn_Envia.State = 1 Then cnn_Envia.Close
'    cnn_Envia.Open StrConexDbEnvia
'
'    sql = "SELECT * FROM QUERYS_PEDIDOS WHERE enviado= 0 ORDER BY ITEM"
'    If RsM.State = 1 Then RsM.Close
'    RsM.Open sql, cnn_Envia, 3, 1
'    'strTABLA = traerCampo("srutas", "EMPSGTE", "EMPRESA", wempresa)
'    Do While Not RsM.EOF
'        cnn_Envia_Mysql.Execute "insert into " & strODBCsgte & " (item,wquery) values  (" & RsM.Fields("item") & ", '" & RsM.Fields("wquery") & "')"
'        cnn_Envia.Execute "update QUERYS_PEDIDOS set enviado = -1 where item = " & RsM.Fields("item") & ""
'    RsM.MoveNext
'    Loop
'
'Exit Sub
'
'error_conexion:
'    If Err.Number = -2147217900 Then
'        Resume Next
'    ElseIf Err.Number = -2147467259 Then
'        Exit Sub
'    Else
'        MsgBox "El Documento no fue enviado " & Err.Description
'    End If
    
End Sub
Public Sub Recibir_Querys()

'Dim cnn_Recibe_Mysql As New ADODB.Connection
'Dim cnn_Recibe As New ADODB.Connection
'Dim cn As String
'Dim Rs As New ADODB.Recordset
'
'    If cnn_Recibe_Mysql.State = 1 Then cnn_Recibe_Mysql.Close
'    'strODBC = traerCampo("srutas", "TABLAS", "EMPRESA", wempresa)
'    'cnn_Envia_Mysql.Open "Driver={Mysql};DATA SOURCE=" & ControlPlus & ";"
'    cnn_Recibe_Mysql.Open strODBCtablas
'
'    StrConexDbEnvia = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_Recibe.MDB;Persist Security Info=False"
'    If cnn_Recibe.State = 1 Then cnn_Recibe.Close
'    cnn_Recibe.Open StrConexDbEnvia
'
'    'strODBC = traerCampo("srutas", "EMPSGTE", "EMPRESA", wempresa)
'
'    sql = "SELECT * FROM " & strODBCsgte & " WHERE ISNULL(SUC_1) ORDER BY ITEM"
'    If RsM.State = 1 Then RsM.Close
'    RsM.Open sql, cnn_Recibe_Mysql, 3, 1
'    Do While Not RsM.EOF
'        cnn_Recibe.Execute "insert into querys (item,wquery) values ('" & RsM.Fields("item") & "','" & RsM.Fields("wquery") & "')"
'        cnn_Recibe_Mysql.Execute "update " & strODBCsgte & " SET SUC_1=-1 where item = " & RsM.Fields("item")
'    RsM.MoveNext
'    Loop
'
'    sql = "SELECT * FROM QUERYS WHERE  APLICADO= 0 ORDER BY ITEM"
'    If Rs.State = 1 Then Rs.Close
'    Rs.Open sql, cnn_Recibe, 3, 1
'    Do While Not Rs.EOF
'        xcad = Replace(Rs!WQUERY, "|", "'")
'        cnn_dbbancos.Execute xcad
'        cnn_Recibe.Execute "UPDATE QUERYS SET APLICADO=-1 WHERE ITEM =" & Rs!ITEM
'        Rs.MoveNext
'    Loop
End Sub
Public Function IniciaDSN(sDSNname As String) As Boolean
'    'Comprobamos si existe
'    If ExisteDSN(sDSNname) = True Then
'        'Si existe lo eliminamos previamente.
'        If BorrarDSN(sDSNname, C_MYSQL_ODBC) = False Then
'            IniciaDSN = False
'            Exit Function
'        End If
'    End If
'
'    'Creamos el nuevo DSN.
'    IniciaDSN = MySQLCrearDSN(sDSNname)
End Function
 

'Crea un DSN del sistema.
Public Function CrearDSN(sDSN As String, sDriver As String, sAtributos As String, Optional sHwnd As Long = vbAPINull) As Boolean
    'Creamos el DSN (En vez de vbAPINull, empleamos el hwnd del formulario)
    CrearDSN = CBool(SQLConfigDataSource(sHwnd, ODBC_ADD_SYS_DSN, sDriver, sAtributos))
End Function


'Crea un DSN MySQL con los atributos bien seteados.
Public Function MySQLCrearDSN(sDSN As String, _
 Optional sServer As String = C_Server, Optional sBD As String = C_BD, _
 Optional sUser As String = C_User, Optional sPass As String = C_Pass, _
 Optional sPort As Integer = C_Port) As Boolean
'
'    Dim sDriver As String
'    Dim sAtributos As String
'
'    sDriver = C_MYSQL_ODBC
'    sAtributos = "DSN=" & sDSN & Chr(0)
'    sAtributos = sAtributos & "SERVER=" & sServer & Chr(0)
'
'    sAtributos = sAtributos & "PORT=" & sPort & Chr(0)
'
'    sAtributos = sAtributos & "DATABASE=" & sBD & Chr(0)
'
'    sAtributos = sAtributos & "USER=" & sUser & Chr(0)
'
'    sAtributos = sAtributos & "PASSWORD=" & sPass & Chr(0)
'
'    sAtributos = sAtributos & "OPTION=3" & Chr(0)
'
'    'Si queremos resetear la conexión de datos, debemos borrarlo antes
'    If ExisteDSN(sDSN) Then
'        Call BorrarDSN(sDSN, sDriver)
'    End If
'
'    MySQLCrearDSN = CrearDSN(sDSN, sDriver, sAtributos)

End Function


'Elimina un DSN del sistema.
Public Function BorrarDSN(sDSN As String, sDriver As String, Optional sHwnd As Long = vbAPINull) As Boolean
    Dim sAtributos As String
    ' Borramos el DSN (En vez de vbAPINull, empleamos el hwnd del formulario)
    If ExisteDSN(sDSN) Then
        sAtributos = "DSN=" & sDSN
        BorrarDSN = CBool(SQLConfigDataSource(sHwnd, ODBC_REMOVE_SYS_DSN, sDriver, sAtributos))
    Else
        'MsgBox ExIdioma("ModDSN_Contr1")
        BorrarDSN = False
    End If
End Function


'Comprueba si existe un DSN en el sistema.
Public Function ExisteDSN(sDSN As String) As Boolean
    Dim I As Integer, J As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSNActual As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long 'controlador del entorno
    Dim DSNLISTA(100)
    ExisteDSN = False
    For J = 1 To 52
        DSNLISTA(J) = ""
    Next J
    
    J = 1
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until I <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            I = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSNActual = VBA.left(sDSNItem, iDSNLen)
            sDRV = VBA.left(sDRVItem, iDRVLen)
            If sDSN <> Space(iDSNLen) Then
                DSNLISTA(J) = sDSN
                If UCase(sDSN) = UCase(sDSNActual) Then
                    ExisteDSN = True
                    Exit Do
                End If
            End If
        Loop
    End If
End Function

Public Sub valida_sunat(rucc As String)
Dim celda
On Error Resume Next
Web = "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & rucc
principio = rucc
Final = "<br/></small>"
Set XML = CreateObject("Microsoft.XMLHTTP")
XML.Open "POST", Web, False
XML.send
Texto = XML.responseText
posicion1 = InStr(XML.responseText, principio)
posicion2 = InStr(XML.responseText, Final)
ruc_rsocial = Mid(XML.responseText, posicion1 + 14, (posicion2 - posicion1) - 14)

'direccion'
principio = "Direcci"
Final = "Situaci"
posicion1 = InStr(XML.responseText, principio)
posicion2 = InStr(XML.responseText, Final)
ruc_direccion = Mid(XML.responseText, posicion1 + 24, (posicion2 - posicion1) - 52)

' estado
principio = "Estado."
Final = "Agente"
posicion1 = InStr(XML.responseText, principio)
posicion2 = InStr(XML.responseText, Final)
ruc_estado = Mid(XML.responseText, posicion1 + 11, (posicion2 - posicion1) - 53)
ruc_estado = IIf(left(ruc_estado, 1) = "A", "ACTIVO", left(ruc_estado, 8))


' situacion
principio = "Situaci"
Final = "Tel"
posicion1 = InStr(XML.responseText, principio)
posicion2 = InStr(XML.responseText, Final)
ruc_situacion = Trim(Mid(XML.responseText, posicion1 + 18, (posicion2 - posicion1) - 53))

' telefono
principio = "Tel"
Final = "Dependencia"
posicion1 = InStr(XML.responseText, principio)
posicion2 = InStr(XML.responseText, Final)
ruc_telefono = left(Trim(Mid(XML.responseText, posicion1 + 26, (posicion2 - posicion1) - 53)), 7)



If Err <> 0 Then
    MsgBox "Verificar RUC o Verificar conexión a Internet"
End If
Set XML = Nothing

End Sub



