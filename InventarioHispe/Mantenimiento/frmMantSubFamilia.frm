VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMantSubFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Sub-Familia"
   ClientHeight    =   2805
   ClientLeft      =   4710
   ClientTop       =   2445
   ClientWidth     =   4815
   Icon            =   "frmMantSubFamilia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4815
   Begin ActiveToolBars.SSActiveToolBars tlbSubFamilia 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantSubFamilia.frx":058A
      ToolBars        =   "frmMantSubFamilia.frx":382C
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtCodExterno 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkHabilitado 
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Externo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMantSubFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda                As Boolean
Private bolNuevoSubFamilia      As Boolean
Private strCodigo               As String

Private objSubFamilia           As ClsSubFamilia

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoSubFamilia(ByVal Value As Boolean)
    bolNuevoSubFamilia = Value
End Property

Public Property Get NuevoSubFamilia() As Boolean
    NuevoSubFamilia = bolNuevoSubFamilia
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Private Sub limpiarCajas()
    txtCodigo.Text = vbNullString
    txtCodExterno.Text = vbNullString
    chkHabilitado.Value = vbChecked
    txtDescripcion.Text = vbNullString
    
    txtCodigo.Locked = True
    txtCodigo.BackColor = DF
End Sub

Private Sub consultarSubFamilia()
    Set objSubFamilia = New ClsSubFamilia
    
    limpiarCajas
    
    With objSubFamilia
        .Codigo = strCodigo
        
        If .obtenerSubFamilia Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaSubFamilia
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        .listarSubFamiliaSQL
                    Else
                        .listarSubFamilia
                    End If
                End With
            End If
        Else
            txtCodigo.Text = .generarCodigoSubFamilia
        End If
    End With
    
    Set objSubFamilia = Nothing
End Sub

Private Sub validarCajas()
    If Trim(txtCodigo.Text) = vbNullString Then
        MsgBox "El Campo Código es obligatorio.", vbCritical, App.ProductName
        
        txtCodigo.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDescripcion.Text) = vbNullString Then
        MsgBox "El Campo Descripción es obligatorio.", vbCritical, App.ProductName
        
        txtCodigo.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarSubFamilia
    End If
End Sub

Private Sub guardarSubFamilia()
    Set objSubFamilia = New ClsSubFamilia
    
    With objSubFamilia
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarSubFamilia Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodigo = .Codigo
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                guardarSubFamiliaSql
                consultarSubFamiliaSql
            Else
                consultarSubFamilia
            End If
            
            
            
            If Not bolAyuda Then
                MsgBox "Sub-Familia Actualizado.", _
                        vbInformation, App.ProductName
            Else
                objAyudaSubFamilia.Codigo = Trim(txtCodigo.Text)
                objAyudaSubFamilia.Descripcion = Trim(txtDescripcion.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objSubFamilia = Nothing
End Sub

Private Sub eliminarSubFamilia()
    Set objSubFamilia = New ClsSubFamilia
    
    With objSubFamilia
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Color no existente.", vbInformation, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar la Sub-Familia con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarSubFamilia Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodigo = .Codigo
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    eliminarSubFamiliaSql
                    consultarSubFamiliaSql
                Else
                    consultarSubFamilia
                End If
                
                MsgBox "Sub-Familia Eliminado.", _
                    vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objSubFamilia = Nothing
End Sub

Private Sub Form_Load()
    consultarSubFamilia
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolAyuda Then
        With frmListaSubFamilia
        
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                .listarSubFamiliaSQL
            Else
                .listarSubFamilia
            End If
            
            '.Show
        End With
    End If
End Sub

Private Sub tlbSubFamilia_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodigo = vbNullString
            
            consultarSubFamilia
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarSubFamilia
        Case "Salir"
            objAyudaSubFamilia.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub



''--------------------------------------------------------------------------------------------------------------
''------------ SQL ---------------------------------------------------------------------------------------------
''--------------------------------------------------------------------------------------------------------------
Private Sub consultarSubFamiliaSql()
    Set objSqlSubFamilia = New SqlClsSubFamilia
    
    limpiarCajas
    
    With objSqlSubFamilia
        .Codigo = strCodigo
        
        If .obtenerSubFamilia Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaSubFamilia
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        .listarSubFamiliaSQL
                    Else
                        .listarSubFamilia
                    End If
                End With
            End If
        Else
            txtCodigo.Text = .generarCodigoSubFamilia
        End If
    End With
    
    Set objSqlSubFamilia = Nothing
End Sub

Private Sub guardarSubFamiliaSql()
    Set objSqlSubFamilia = New SqlClsSubFamilia
    
    With objSqlSubFamilia
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarSubFamilia Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
'            strCodigo = .Codigo
'
'            consultarSubFamilia
'
'            If Not bolAyuda Then
'                MsgBox "Color Actualizado.", _
'                        vbInformation, App.ProductName
'            Else
'                objAyudaSubFamilia.Codigo = Trim(txtCodigo.Text)
'                objAyudaSubFamilia.Descripcion = Trim(txtDescripcion.Text)
'
'                Unload Me
'            End If
        End If
    End With
    
    Set objSqlSubFamilia = Nothing
End Sub

Private Sub eliminarSubFamiliaSql()
    Set objSqlSubFamilia = New SqlClsSubFamilia
    
    With objSqlSubFamilia
        .Codigo = strCodigo  'Trim(txtCodigo.Text)
'
'        If Not .verificarExistencia Then
'            MsgBox "Color no existente.", vbInformation, App.ProductName
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea Eliminar el Color con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarSubFamilia Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'
'                strCodigo = .Codigo
'
'                consultarSubFamilia
'
'                MsgBox "Color Eliminado.", _
'                    vbInformation, App.ProductName
            End If
'        End If
    End With
    
    Set objSqlSubFamilia = Nothing
End Sub
