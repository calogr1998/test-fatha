Attribute VB_Name = "ModSk"
Option Explicit

Public Const strNombreFicheroConfigCPgeneral = "\configCP.ini"
Public Const strNombreFicheroConfigCPusuario = "\configCPusuario.ini"
Public Const strNombreFicheroConfigSQLCliente = "\configServidorSQLCliente.ini"

Public strCadenaConexioBdCPlus      As String

Public cnDBTemp                     As ADODB.Connection
Public cnDBContaTabla               As ADODB.Connection

Public cnBdCPlus                    As ADODB.Connection

'-------------------------------------------------------------
Rem Para recuperar informacion desde Backup Db_Bancos.mdb
Public cnn_dbbancos_backup          As New ADODB.Connection
Public StrConexDbBancos_backup      As String
Public cconex_dbbancos_backup       As String
'-------------------------------------------------------------

Public objAyudaBien                 As New ClsBien
Public objAyudaBienAlterno          As New ClsBienAlterno
Public objAyudaBienColor            As New ClsBienColor
Public objAyudaCategoria            As New ClsCategoria
Public objAyudaCliente              As New ClsCliente
Public objAyudaCompra               As New ClsCompra
Public objAyudaComprobante          As New ClsComprobante
Public objAyudaDistrito             As New ClsDistrito
Public objAyudaFamilia              As New ClsFamilia
Public objAyudaEmpresa              As New ClsEmpresa
Public objAyudaFormaPago            As New ClsFormaPago
Public objAyudaGasto                As New ClsGasto
Public objAyudaOrden                As New ClsOrden
Public objAyudaOrdenTrabajo         As New ClsOrdenTrabajo
Public objAyudaOrigen               As New ClsOrigen
Public objAyudaPagDcto              As New ClsPagDcto
Public objAyudaProvDscto            As New ClsProvDscto
Public objAyudaProveedor            As New ClsProveedor
Public objAyudaSector               As New ClsSectorEmpresarial
Public objAyudaSolicitud            As New ClsSolicitud
Public objAyudaSubFamilia           As New ClsSubFamilia
Public objAyudaTCambio              As New ClsTCambio
Public objAyudaTipoDocID            As New ClsTipoDocumentoID
Public objAyudaTipoExistencia       As New ClsTipoExistencia
Public objAyudaTomaInventario       As New ClsTomaInventario
Public objAyudaUM                   As New ClsUM
Public objAyudaVale                 As New ClsVale
Public objAyudaAlmacen              As New ClsAlmacen
Public objAyudaCentroCosto          As New ClsCentroCosto
Public objAyudaMarca                As New ClsMarca
Public objAyudaUsuario              As New ClsUsuario
Public objAyudaTarea                As New ClsTarea
Public objAyudaTareaUsuario         As New ClsTareaUsuario
Public objAyudaCuentaContable       As New ClsCuentaContable




Public bolCorreoEnviado             As Boolean


Type ParametrosContables
    dblDetraccionPorcentaje As Double
    bolDetraccionImpMsj As Boolean
    strDetraccionMensaje As String
    strDetraccionCtaSoles As String
    strDetraccionCtaDolar As String
End Type

Public infoPlusParConta As ParametrosContables

Global Const HA = &H80000005 'CONSTANTE PARA EL COLOR ACTIVADO
                                 'DEL CONTROL TEXTBOX
Global Const DH = &HC0C0C0 'CONSTANTE PARA EL COLOR DESACTIVADO
                           'DEL CONTROL TEXTBOX
Global Const DF = &HC0FFC0    'CONSTANTE PARA EL COLOR DESACTIVADO
                           'DEL CONTROL TEXTBOX

Public Sub abrirCnnDbBancos()
    If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
    
    cnn_dbbancos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos.mdb;Persist Security Info=False"
End Sub

Public Sub abrirCnnnDbBancosBackup()
    If Dir(wrutabancos & "\db_bancos_backup.mdb", vbArchive) <> "" Then
        StrConexDbBancos_backup = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos_backup.mdb;Persist Security Info=False"
        cconex_dbbancos_backup = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos_backup.mdb" & ";Persist Security Info=False"
            
        If cnn_dbbancos_backup.State = 1 Then cnn_dbbancos_backup.Close
        
        cnn_dbbancos_backup.Open cconex_dbbancos_backup
    End If
End Sub

Public Sub abrirCnContaTabla()
    Set cnDBContaTabla = New ADODB.Connection
    
    If cnDBContaTabla.State = 1 Then cnDBContaTabla.Close
    
    cnDBContaTabla.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\db_tabla.mdb;Persist Security Info=False"
End Sub

Public Sub abrirCnTemporal()
    Set cnDBTemp = New ADODB.Connection
    
    If cnDBTemp.State = 1 Then cnDBTemp.Close
    
    cnDBTemp.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnDBTemp.CursorLocation = adUseClient
    cnDBTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0" & _
                                ";Data Source= " & wrutatemp & "Templus.mdb" & _
                                ";Persist Security Info=False"
    
    cnDBTemp.Open
End Sub

Public Sub abrirCnBdCPlus()
    On Error GoTo errAbrirCnBdCPlus
    
    Set cnBdCPlus = New ADODB.Connection
    
    strCadenaConexioBdCPlus = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "CadenaConexionBdCPlus", "l")
    
    If cnBdCPlus.State = 1 Then cnBdCPlus.Close
    
    cnBdCPlus.Open strCadenaConexioBdCPlus
    
    Exit Sub
errAbrirCnBdCPlus:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - AbrirCnBdCPlus"
    
    Err.Clear
End Sub

