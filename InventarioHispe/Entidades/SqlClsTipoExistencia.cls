VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsTipoExistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoExterno            As String
Private strDescripcion              As String
Private strAbreviatura              As String

Private strCtaContable              As String
Private strAnexo                    As String
Private strCtaContableImp           As String
Private strAnexoImp                 As String

Private strCtaContableVta           As String
Private strAnexoVta                 As String
Private strCtaContableImpVta        As String
Private strAnexoImpVta              As String

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter   As String

Private rstTipoExistencia             As ADODB.Recordset
Private strSQLTipoExistencia          As String



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

Public Property Let Abreviatura(ByVal value As String)
    strAbreviatura = value
End Property

Public Property Get Abreviatura() As String
    Abreviatura = strAbreviatura
End Property

Public Property Let CtaContable(ByVal value As String)
    strCtaContable = value
End Property

Public Property Get CtaContable() As String
    CtaContable = strCtaContable
End Property

Public Property Let Anexo(ByVal value As String)
    strAnexo = value
End Property

Public Property Get Anexo() As String
    Anexo = strAnexo
End Property

Public Property Let CtaContableImportacion(ByVal value As String)
    strCtaContableImp = value
End Property

Public Property Get CtaContableImportacion() As String
    CtaContableImportacion = strCtaContableImp
End Property

Public Property Let AnexoImportacion(ByVal value As String)
    strAnexoImp = value
End Property

Public Property Get AnexoImportacion() As String
    AnexoImportacion = strAnexoImp
End Property




Public Property Let CtaContableVta(ByVal value As String)
    strCtaContableVta = value
End Property

Public Property Get CtaContableVta() As String
    CtaContableVta = strCtaContableVta
End Property

Public Property Let AnexoVta(ByVal value As String)
    strAnexoVta = value
End Property

Public Property Get AnexoVta() As String
    AnexoVta = strAnexoVta
End Property

Public Property Let CtaContableImportacionVta(ByVal value As String)
    strCtaContableImpVta = value
End Property

Public Property Get CtaContableImportacionVta() As String
    CtaContableImportacionVta = strCtaContableImpVta
End Property

Public Property Let AnexoImportacionVta(ByVal value As String)
    strAnexoImpVta = value
End Property

Public Property Get AnexoImportacionVta() As String
    AnexoImportacionVta = strAnexoImpVta
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
    strAbreviatura = vbNullString
    
    strCtaContable = vbNullString
    strAnexo = vbNullString
    strCtaContableImp = vbNullString
    strAnexoImp = vbNullString
    
    strCtaContableVta = vbNullString
    strAnexoVta = vbNullString
    strCtaContableImpVta = vbNullString
    strAnexoImpVta = vbNullString
    
    bolEstado = False
        
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerTipoExistencia() As Boolean
    On Error GoTo errObtenerTipoExistencia
    
    Set rstTipoExistencia = New ADODB.Recordset
    
    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "SELECT * "
    strSQLTipoExistencia = strSQLTipoExistencia & "FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoExistencia.State = 1 Then rstTipoExistencia.Close
    
    rstTipoExistencia.Open strSQLTipoExistencia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoExistencia.EOF Then
        strCodigo = Trim(rstTipoExistencia!Codigo & "")
        strCodigoExterno = Trim(rstTipoExistencia!CodigoExterno & "")
        strDescripcion = Trim(rstTipoExistencia!Descripcion & "")
        strAbreviatura = Trim(rstTipoExistencia!Abreviatura & "")
        
        strCtaContable = Trim(rstTipoExistencia!CtaContable & "")
        strAnexo = Trim(rstTipoExistencia!Anexo & "")
        strCtaContableImp = Trim(rstTipoExistencia!CtaContableImportacion & "")
        strAnexoImp = Trim(rstTipoExistencia!AnexoImportacion & "")
        
        strCtaContableVta = Trim(rstTipoExistencia!CtaContableVta & "")
        strAnexoVta = Trim(rstTipoExistencia!AnexoVta & "")
        strCtaContableImpVta = Trim(rstTipoExistencia!CtaContableImportacionVta & "")
        strAnexoImpVta = Trim(rstTipoExistencia!AnexoImportacionVta & "")
        
        bolEstado = CBool(rstTipoExistencia!Estado)
        
        obtenerTipoExistencia = True
    Else
        obtenerTipoExistencia = False
    End If
    
    rstTipoExistencia.Close
    
    Set rstTipoExistencia = Nothing
    
    Exit Function
errObtenerTipoExistencia:
    obtenerTipoExistencia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: ObtenerTipoExistencia"
    
    Err.Clear
End Function

Public Sub obtenerConfigTipoExistencia()
    On Error GoTo errObtenerConfigTipoExistencia
    
    Set rstTipoExistencia = New ADODB.Recordset
    
    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "SELECT * "
    strSQLTipoExistencia = strSQLTipoExistencia & "FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoExistencia.State = 1 Then rstTipoExistencia.Close
    
    rstTipoExistencia.Open strSQLTipoExistencia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoExistencia.EOF Then
        strCodigo = Trim(rstTipoExistencia!Codigo & "")
        strCodigoExterno = Trim(rstTipoExistencia!CodigoExterno & "")
        strDescripcion = Trim(rstTipoExistencia!Descripcion & "")
        strAbreviatura = Trim(rstTipoExistencia!Abreviatura & "")
        
        strCtaContable = Trim(rstTipoExistencia!CtaContable & "")
        strAnexo = Trim(rstTipoExistencia!Anexo & "")
        strCtaContableImp = Trim(rstTipoExistencia!CtaContableImportacion & "")
        strAnexoImp = Trim(rstTipoExistencia!AnexoImportacion & "")
        
        strCtaContableVta = Trim(rstTipoExistencia!CtaContableVta & "")
        strAnexoVta = Trim(rstTipoExistencia!AnexoVta & "")
        strCtaContableImpVta = Trim(rstTipoExistencia!CtaContableImportacionVta & "")
        strAnexoImpVta = Trim(rstTipoExistencia!AnexoImportacionVta & "")
        
        bolEstado = CBool(rstTipoExistencia!Estado & "")
    End If
    
    rstTipoExistencia.Close
    
    Set rstTipoExistencia = Nothing
    
    Exit Sub
errObtenerConfigTipoExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: ObtenerTipoExistencia"
    
    Err.Clear
End Sub

Public Function generarCodigoTipoExistencia() As String
    On Error GoTo errGenerarCodigoTipoExistencia

    Dim rsTipoExistencia As New ADODB.Recordset

    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "SELECT "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO "
    strSQLTipoExistencia = strSQLTipoExistencia & "FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "ORDER BY "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO"
    
    If rsTipoExistencia.State = 1 Then rsTipoExistencia.Close

    rsTipoExistencia.Open strSQLTipoExistencia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly

    If Not rsTipoExistencia.EOF Then
        rsTipoExistencia.MoveFirst

        Do While Not rsTipoExistencia.EOF
            strCodigo = Format(Val(Trim(rsTipoExistencia!Codigo & "")) + 1, "00")

            If Not verificarExistencia Then
                Exit Do
            End If

            rsTipoExistencia.MoveNext
        Loop

        generarCodigoTipoExistencia = strCodigo
    Else
        generarCodigoTipoExistencia = "01"
    End If

    rsTipoExistencia.Close

    Set rsTipoExistencia = Nothing

    strSQLTipoExistencia = vbNullString

    Exit Function
errGenerarCodigoTipoExistencia:
    generarCodigoTipoExistencia = vbNullString

    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: GenerarCodigoTipoExistencia"

    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTipoExistencia = New ADODB.Recordset
    
    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "SELECT "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA.* "
    strSQLTipoExistencia = strSQLTipoExistencia & "FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoExistencia.State = 1 Then rstTipoExistencia.Close
    
    rstTipoExistencia.Open strSQLTipoExistencia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoExistencia.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstTipoExistencia.Close
    
    strSQLTipoExistencia = vbNullString
    
    Set rstTipoExistencia = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarTipoExistencia() As Boolean
    On Error GoTo errGuardarTipoExistencia
    
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoTipoExistencia
        End If
        
        strSQLTipoExistencia = vbNullString
        strSQLTipoExistencia = strSQLTipoExistencia & "INSERT INTO MAESTROS.EF2TIPOEXISTENCIA("
        strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO, CODIGOEXTERNO, DESCRIPCION, ABREVIATURA, "
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLE, ANEXO, CTACONTABLEIMPORTACION, ANEXOIMPORTACION, "
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLEVTA, ANEXOVTA, CTACONTABLEIMPORTACIONVTA, ANEXOIMPORTACIONVTA, "
        strSQLTipoExistencia = strSQLTipoExistencia & "ESTADO, FECREG, USUREG) "
        strSQLTipoExistencia = strSQLTipoExistencia & "VALUES("
        strSQLTipoExistencia = strSQLTipoExistencia & "'" & strCodigo & "', "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "'" & strDescripcion & "', '" & strAbreviatura & "', "
        
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strCtaContable <> vbNullString, " '" & strCtaContable & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strAnexo <> vbNullString, " '" & strAnexo & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strCtaContableImp <> vbNullString, " '" & strCtaContableImp & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strAnexoImp <> vbNullString, " '" & strAnexoImp & "'", "''") & ", "
        
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strCtaContableVta <> vbNullString, " '" & strCtaContableVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strAnexoVta <> vbNullString, " '" & strAnexoVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strCtaContableImpVta <> vbNullString, " '" & strCtaContableImpVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(strAnexoImpVta <> vbNullString, " '" & strAnexoImpVta & "'", "''") & ", "
        
        strSQLTipoExistencia = strSQLTipoExistencia & IIf(bolEstado, "1", "0") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "'" & strFecReg & "', '" & strUsuReg & "')"
    Else
        strSQLTipoExistencia = vbNullString
        strSQLTipoExistencia = strSQLTipoExistencia & "UPDATE "
        strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
        strSQLTipoExistencia = strSQLTipoExistencia & "SET "
        strSQLTipoExistencia = strSQLTipoExistencia & "CODIGOEXTERNO = " & IIf(strCodigoExterno <> vbNullString, "'" & strCodigoExterno & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "DESCRIPCION = '" & strDescripcion & "', "
        strSQLTipoExistencia = strSQLTipoExistencia & "ABREVIATURA = '" & strAbreviatura & "', "
        
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLE = " & IIf(strCtaContable <> vbNullString, " '" & strCtaContable & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "ANEXO = " & IIf(strAnexo <> vbNullString, " '" & strAnexo & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLEIMPORTACION = " & IIf(strCtaContableImp <> vbNullString, " '" & strCtaContableImp & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "ANEXOIMPORTACION = " & IIf(strAnexoImp <> vbNullString, " '" & strAnexoImp & "'", "''") & ", "
        
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLEVTA = " & IIf(strCtaContableVta <> vbNullString, " '" & strCtaContableVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "ANEXOVTA = " & IIf(strAnexoVta <> vbNullString, " '" & strAnexoVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "CTACONTABLEIMPORTACIONVTA = " & IIf(strCtaContableImpVta <> vbNullString, " '" & strCtaContableImpVta & "'", "''") & ", "
        strSQLTipoExistencia = strSQLTipoExistencia & "ANEXOIMPORTACIONVTA = " & IIf(strAnexoImpVta <> vbNullString, " '" & strAnexoImpVta & "'", "''") & ", "
        
        strSQLTipoExistencia = strSQLTipoExistencia & "ESTADO = " & IIf(bolEstado, "1, ", "0, ")
        strSQLTipoExistencia = strSQLTipoExistencia & "FECMOD = '" & strFecMod & "', "
        strSQLTipoExistencia = strSQLTipoExistencia & "USUMOD = '" & strUsuMod & "' "
        strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
        strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLTipoExistencia
    
    guardarTipoExistencia = True
    
    strSQLSelectAlter = strSQLTipoExistencia
    strSQLTipoExistencia = vbNullString
    
    Exit Function
errGuardarTipoExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: GuardarTipoExistencia"
    
    guardarTipoExistencia = False
    
    Err.Clear
End Function

Public Function eliminarTipoExistencia() As Boolean
    On Error GoTo errEliminarTipoExistencia
    
    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F5TIPO", "MAESTROS.IF5PLA", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarTipoExistencia = False
        
        Exit Function
    End If
    
    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "DELETE FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLTipoExistencia
    
    eliminarTipoExistencia = True
    
    strSQLSelectAlter = strSQLTipoExistencia
    strSQLTipoExistencia = vbNullString
    
    Exit Function
errEliminarTipoExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: EliminarTipoExistencia"
    
    eliminarTipoExistencia = False
    
    Err.Clear
End Function

Public Sub listarTipoExistencia(ByVal comboList As Object, _
                                ByVal bolIncluirSoloTiposConCodigoExterno As Boolean, _
                                ByVal bolAdicionarOpcionTodosAlListado As Boolean)
    
    On Error GoTo errListarTipoExistencia
    
    Set rstTipoExistencia = New ADODB.Recordset
    
    strSQLTipoExistencia = vbNullString
    strSQLTipoExistencia = strSQLTipoExistencia & "SELECT "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO, "
    strSQLTipoExistencia = strSQLTipoExistencia & "DESCRIPCION "
    strSQLTipoExistencia = strSQLTipoExistencia & "FROM "
    strSQLTipoExistencia = strSQLTipoExistencia & "MAESTROS.EF2TIPOEXISTENCIA "
    strSQLTipoExistencia = strSQLTipoExistencia & "WHERE "
    strSQLTipoExistencia = strSQLTipoExistencia & "ESTADO = 1 "
            
            If bolIncluirSoloTiposConCodigoExterno Then
                strSQLTipoExistencia = strSQLTipoExistencia & "AND TRIM(CODIGOEXTERNO & '') <> '' "
            End If
            
    strSQLTipoExistencia = strSQLTipoExistencia & "ORDER BY "
    strSQLTipoExistencia = strSQLTipoExistencia & "CODIGO"
    
    If rstTipoExistencia.State = 1 Then rstTipoExistencia.Close
    
    rstTipoExistencia.Open strSQLTipoExistencia, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstTipoExistencia.EOF Then
        If bolAdicionarOpcionTodosAlListado Then
            comboList.AddItem "(*) - Todos"
        End If
        
        Do While Not rstTipoExistencia.EOF
            comboList.AddItem Trim(rstTipoExistencia!Descripcion & "") & Space(150) & Trim(rstTipoExistencia!Codigo & "")
            
            rstTipoExistencia.MoveNext
        Loop
            comboList.ListIndex = 0
    End If
    
    rstTipoExistencia.Close
    
    strSQLTipoExistencia = vbNullString
    
    Set rstTipoExistencia = Nothing
    
    Exit Sub
errListarTipoExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoExistencia: ListarTipoExistencia"
    
    Err.Clear
End Sub
