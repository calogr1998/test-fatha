VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strDescripcion              As String
Private strDescripcionCorta         As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

'Atributos Adicionales
Private strCodigoExterno2           As String
Private strAnexo                    As String
Private strAnexoImportacion         As String




Private strSQLSelectAlter   As String

Private rstFamilia             As ADODB.Recordset
Private strSQLFamilia          As String

Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let CodigoExterno(ByVal value As String)
    strCodigoExterno = value
End Property

Public Property Get CodigoExterno() As String
    CodigoExterno = strCodigoExterno
End Property

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property
'DescripcionCorta
Public Property Let DescripcionCorta(ByVal value As String)
    strDescripcionCorta = value
End Property

Public Property Get DescripcionCorta() As String
    DescripcionCorta = strDescripcionCorta
End Property



Public Property Let Estado(ByVal value As Boolean)
    bolEstado = value
End Property

Public Property Get Estado() As Boolean
    Estado = bolEstado
End Property



'Propiedad Fecha de Registro de Color de Bien
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Color de Bien
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property



'Atributos Adicionales
'Codigo Externo 2
Public Property Let CodigoExterno2(ByVal value As String)
    strCodigoExterno2 = value
End Property

Public Property Get CodigoExterno2() As String
    CodigoExterno2 = strCodigoExterno2
End Property
'Anexo
Public Property Let Anexo(ByVal value As String)
    strAnexo = value
End Property

Public Property Get Anexo() As String
    Anexo = strAnexo
End Property
'Anexo de Importacion
Public Property Let AnexoImportacion(ByVal value As String)
    strAnexoImportacion = value
End Property

Public Property Get AnexoImportacion() As String
    AnexoImportacion = strAnexoImportacion
End Property



Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    strCodigoExterno = vbNullString
    strDescripcion = vbNullString
    strDescripcionCorta = vbNullString
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strCodigoExterno2 = vbNullString
    strAnexo = vbNullString
    strAnexoImportacion = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerFamilia() As Boolean
    On Error GoTo errObtenerFamilia
    
    Set rstFamilia = New ADODB.Recordset
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "SELECT "
    strSQLFamilia = strSQLFamilia & "* "
    strSQLFamilia = strSQLFamilia & "FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "WHERE "
    strSQLFamilia = strSQLFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        strCodigo = Trim(rstFamilia!F7CODCON & "")
        strCodigoExterno = Trim(rstFamilia!CodExterno & "")
        strDescripcion = Trim(rstFamilia!F7DESCON & "")
        strDescripcionCorta = Trim(rstFamilia!F7DESCONCORTA & "")
        
        bolEstado = CBool(rstFamilia!Estado)
        
        strFecReg = Trim(rstFamilia!FecReg & "")
        strUsuReg = Trim(rstFamilia!UsuReg & "")
        strFecMod = Trim(rstFamilia!FecMod & "")
        strUsuMod = Trim(rstFamilia!UsuMod & "")
        
        strCodigoExterno2 = Trim(rstFamilia!CodExterno2 & "")
        strAnexo = Trim(rstFamilia!Anexo & "")
        strAnexoImportacion = Trim(rstFamilia!AnexoImportacion & "")
        
        obtenerFamilia = True
    Else
        obtenerFamilia = False
    End If
    
    rstFamilia.Close
    
    Set rstFamilia = Nothing
    
    Exit Function
errObtenerFamilia:
    obtenerFamilia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: ObtenerFamilia"
    
    Err.Clear
End Function

Public Sub obtenerConfigFamilia()
    On Error GoTo errObtenerConfigFamilia
    
    Set rstFamilia = New ADODB.Recordset
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "SELECT "
    strSQLFamilia = strSQLFamilia & "* "
    strSQLFamilia = strSQLFamilia & "FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "WHERE "
    strSQLFamilia = strSQLFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        strCodigo = Trim(rstFamilia!F7CODCON & "")
        strCodigoExterno = Trim(rstFamilia!CodExterno & "")
        strDescripcion = Trim(rstFamilia!F7DESCON & "")
        strDescripcionCorta = Trim(rstFamilia!F7DESCONCORTA & "")
        
        bolEstado = CBool(rstFamilia!Estado)
        
        strFecReg = Trim(rstFamilia!FecReg & "")
        strUsuReg = Trim(rstFamilia!UsuReg & "")
        strFecMod = Trim(rstFamilia!FecMod & "")
        strUsuMod = Trim(rstFamilia!UsuMod & "")
        
        strCodigoExterno2 = Trim(rstFamilia!CodExterno2 & "")
        strAnexo = Trim(rstFamilia!Anexo & "")
        strAnexoImportacion = Trim(rstFamilia!AnexoImportacion & "")
    End If
    
    rstFamilia.Close
    
    Set rstFamilia = Nothing
    
    Exit Sub
errObtenerConfigFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: ObtenerFamilia"
    
    Err.Clear
End Sub

Public Function generarCodigoFamilia() As String
    On Error GoTo errGenerarCodigoFamilia
    
    Set rstFamilia = New ADODB.Recordset
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "SELECT TOP 1 "
    strSQLFamilia = strSQLFamilia & "F7CODCON AS CODIGO "
    strSQLFamilia = strSQLFamilia & "FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "ORDER BY "
    strSQLFamilia = strSQLFamilia & "F7CODCON DESC"
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        generarCodigoFamilia = Format(Val(Trim(rstFamilia!Codigo & "")) + 1, "0000")
    Else
        generarCodigoFamilia = "0001"
    End If
    
    rstFamilia.Close
    
    Set rstFamilia = Nothing
    
    strSQLFamilia = vbNullString
    
    Exit Function
errGenerarCodigoFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: GenerarCodigoFamilia"
    
    generarCodigoFamilia = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstFamilia = New ADODB.Recordset
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "SELECT "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01.* "
    strSQLFamilia = strSQLFamilia & "FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "WHERE "
    strSQLFamilia = strSQLFamilia & "F7CODCON = '" & strCodigo & "'"
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstFamilia.Close
    
    strSQLFamilia = vbNullString
    
    Set rstFamilia = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarFamilia() As Boolean
    On Error GoTo errGuardarFamilia
    
    strSQLFamilia = vbNullString
    
    If Not verificarExistencia Then
        strCodigo = generarCodigoFamilia
        
        strSQLFamilia = strSQLFamilia & "INSERT INTO SF7NIVEL01("
        strSQLFamilia = strSQLFamilia & "F7CODCON, CODEXTERNO, F7DESCON, "
        strSQLFamilia = strSQLFamilia & "F7DESCONCORTA, ESTADO, FECREG, USUREG"
        strSQLFamilia = strSQLFamilia & ") "
        strSQLFamilia = strSQLFamilia & "VALUES("
        strSQLFamilia = strSQLFamilia & "'" & strCodigo & "', "
        strSQLFamilia = strSQLFamilia & "'" & strCodigoExterno & "', "
        strSQLFamilia = strSQLFamilia & "'" & strDescripcion & "', "
        strSQLFamilia = strSQLFamilia & "'" & strDescripcionCorta & "', "
        strSQLFamilia = strSQLFamilia & IIf(bolEstado, "TRUE", "FALSE") & ", "
        strSQLFamilia = strSQLFamilia & "CVDATE('" & strFecReg & "'), "
        strSQLFamilia = strSQLFamilia & "'" & strUsuReg & "'"
        strSQLFamilia = strSQLFamilia & ")"
    Else
        strSQLFamilia = strSQLFamilia & "UPDATE "
        strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
        strSQLFamilia = strSQLFamilia & "SET "
        strSQLFamilia = strSQLFamilia & "CODEXTERNO = '" & strCodigoExterno & "', "
        strSQLFamilia = strSQLFamilia & "F7DESCON = '" & strDescripcion & "', "
        strSQLFamilia = strSQLFamilia & "F7DESCONCORTA = '" & strDescripcionCorta & "', "
        strSQLFamilia = strSQLFamilia & "ESTADO = " & IIf(bolEstado, "TRUE, ", "FALSE, ")
        strSQLFamilia = strSQLFamilia & "FECMOD = CVDATE('" & strFecMod & "'), "
        strSQLFamilia = strSQLFamilia & "USUMOD = '" & strUsuMod & "' "
        strSQLFamilia = strSQLFamilia & "WHERE "
        strSQLFamilia = strSQLFamilia & "F7CODCON = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLFamilia
    
    guardarFamilia = True
    
    strSQLSelectAlter = strSQLFamilia
    strSQLFamilia = vbNullString
    
    Exit Function
errGuardarFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: GuardarFamilia"
    
    guardarFamilia = False
    
    Err.Clear
End Function

Public Function eliminarFamilia() As Boolean
    On Error GoTo errEliminarFamilia
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F7NIVEL01", "SF7NIVEL", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarFamilia = False
        
        Exit Function
    End If
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "DELETE FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "WHERE "
    strSQLFamilia = strSQLFamilia & "F7CODCON = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLFamilia
    
    eliminarFamilia = True
    
    strSQLSelectAlter = strSQLFamilia
    strSQLFamilia = vbNullString
    
    Exit Function
errEliminarFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: EliminarFamilia"
    
    eliminarFamilia = False
    
    Err.Clear
End Function

Public Sub listarFamilia(ByVal comboList As Object)
    On Error GoTo errListarFamilia
    
    Set rstFamilia = New ADODB.Recordset
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        strSQLFamilia = vbNullString
        strSQLFamilia = strSQLFamilia & "SELECT "
        strSQLFamilia = strSQLFamilia & "F7CODCON, "
        strSQLFamilia = strSQLFamilia & "F7DESCON "
        strSQLFamilia = strSQLFamilia & "FROM "
        strSQLFamilia = strSQLFamilia & "MAESTROS.SF7NIVEL01 "
        strSQLFamilia = strSQLFamilia & "WHERE "
        strSQLFamilia = strSQLFamilia & "(REPLACE('CODEXTERNO', ' ', '') + '') <> '' AND "
        strSQLFamilia = strSQLFamilia & "ESTADO = 1 "
        strSQLFamilia = strSQLFamilia & "ORDER BY "
        strSQLFamilia = strSQLFamilia & "F7CODCON"
        
        rstFamilia.Open strSQLFamilia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
    
        strSQLFamilia = vbNullString
        strSQLFamilia = strSQLFamilia & "SELECT "
        strSQLFamilia = strSQLFamilia & "F7CODCON, "
        strSQLFamilia = strSQLFamilia & "F7DESCON "
        strSQLFamilia = strSQLFamilia & "FROM "
        strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
        strSQLFamilia = strSQLFamilia & "WHERE "
        strSQLFamilia = strSQLFamilia & "TRIM(CODEXTERNO & '') <> '' AND "
        strSQLFamilia = strSQLFamilia & "ESTADO = TRUE "
        strSQLFamilia = strSQLFamilia & "ORDER BY "
        strSQLFamilia = strSQLFamilia & "F7CODCON"
    
        rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    comboList.Clear
    
    If Not rstFamilia.EOF Then
        comboList.AddItem "(*) - Todos" & Space(150)
        
        Do While Not rstFamilia.EOF
            comboList.AddItem Trim(rstFamilia!F7DESCON & "") & Space(150) & Trim(rstFamilia!F7CODCON & "")
            
            rstFamilia.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstFamilia.Close
    strSQLFamilia = vbNullString
    
    Set rstFamilia = Nothing
    
    Exit Sub
errListarFamilia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: ListarFamilia"
    
    Err.Clear
End Sub

Public Sub listarFamiliaSoloSeleccion(ByVal comboList As ComboBox)
    On Error GoTo errListarFamiliaSoloSeleccion
    
    Set rstFamilia = New ADODB.Recordset
    
    strSQLFamilia = vbNullString
    strSQLFamilia = strSQLFamilia & "SELECT "
    strSQLFamilia = strSQLFamilia & "F7CODCON, "
    strSQLFamilia = strSQLFamilia & "F7DESCON "
    strSQLFamilia = strSQLFamilia & "FROM "
    strSQLFamilia = strSQLFamilia & "SF7NIVEL01 "
    strSQLFamilia = strSQLFamilia & "WHERE "
    strSQLFamilia = strSQLFamilia & "TRIM(CODEXTERNO & '') <> '' AND "
    strSQLFamilia = strSQLFamilia & "ESTADO = TRUE "
    strSQLFamilia = strSQLFamilia & "ORDER BY "
    strSQLFamilia = strSQLFamilia & "F7CODCON"
    
    If rstFamilia.State = 1 Then rstFamilia.Close
    
    rstFamilia.Open strSQLFamilia, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFamilia.EOF Then
        comboList.Clear
        
        Do While Not rstFamilia.EOF
            comboList.AddItem Trim(rstFamilia!Descripcion & "") & Space(50) & Trim(rstFamilia!Codigo & "")
            
            rstFamilia.MoveNext
        Loop
    End If
    
    comboList.ListIndex = -1
    
    rstFamilia.Close
    strSQLFamilia = vbNullString
    
    Set rstFamilia = Nothing
    
    Exit Sub
errListarFamiliaSoloSeleccion:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFamilia: ListarFamiliaSoloSeleccion"
    
    Err.Clear
End Sub

