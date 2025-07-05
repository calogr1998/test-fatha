VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMantBienColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Color"
   ClientHeight    =   2805
   ClientLeft      =   4710
   ClientTop       =   2445
   ClientWidth     =   4815
   Icon            =   "frmMantBienColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4815
   Begin ActiveToolBars.SSActiveToolBars tlbBienColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantBienColor.frx":058A
      ToolBars        =   "frmMantBienColor.frx":382C
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
Attribute VB_Name = "frmMantBienColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda                As Boolean
Private bolNuevoBienColor       As Boolean
Private strCodigo               As String

Private objBienColor            As ClsBienColor

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoBienColor(ByVal Value As Boolean)
    bolNuevoBienColor = Value
End Property

Public Property Get NuevoBienColor() As Boolean
    NuevoBienColor = bolNuevoBienColor
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

Private Sub consultarBienColor()
    Set objBienColor = New ClsBienColor
    
    limpiarCajas
    
    With objBienColor
        .Codigo = strCodigo
        
        If .obtenerBienColor Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaBienColor
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        .listarBienColorSQL
                    Else
                        .listarBienColor
                    End If
                End With
            End If
        Else
            txtCodigo.Text = .generarCodigoBienColor
        End If
    End With
    
    Set objBienColor = Nothing
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
        guardarBienColor
    End If
End Sub

Private Sub guardarBienColor()
    Set objBienColor = New ClsBienColor
    
    With objBienColor
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarBienColor Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodigo = .Codigo
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                guardarBienColorSql
                consultarBienColorSql
            Else
                consultarBienColor
            End If
            
            
            
            If Not bolAyuda Then
                MsgBox "Color Actualizado.", _
                        vbInformation, App.ProductName
            Else
                objAyudaBienColor.Codigo = Trim(txtCodigo.Text)
                objAyudaBienColor.Descripcion = Trim(txtDescripcion.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objBienColor = Nothing
End Sub

Private Sub eliminarBienColor()
    Set objBienColor = New ClsBienColor
    
    With objBienColor
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Color no existente.", vbInformation, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar el Color con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarBienColor Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodigo = .Codigo
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    eliminarBienColorSql
                    consultarBienColorSql
                Else
                    consultarBienColor
                End If
                
                MsgBox "Color Eliminado.", _
                    vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objBienColor = Nothing
End Sub

Private Sub Form_Load()
    consultarBienColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolAyuda Then
        With frmListaBienColor
        
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                .listarBienColorSQL
            Else
                .listarBienColor
            End If
            
            '.Show
        End With
    End If
End Sub

Private Sub tlbBienColor_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodigo = vbNullString
            
            consultarBienColor
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarBienColor
        Case "Salir"
            objAyudaBienColor.inicializarEntidades
            
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
Private Sub consultarBienColorSql()
    Set objSqlBienColor = New SqlClsBienColor
    
    limpiarCajas
    
    With objSqlBienColor
        .Codigo = strCodigo
        
        If .obtenerBienColor Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaBienColor
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        .listarBienColorSQL
                    Else
                        .listarBienColor
                    End If
                End With
            End If
        Else
            txtCodigo.Text = .generarCodigoBienColor
        End If
    End With
    
    Set objSqlBienColor = Nothing
End Sub

Private Sub guardarBienColorSql()
    Set objSqlBienColor = New SqlClsBienColor
    
    With objSqlBienColor
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarBienColor Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
'            strCodigo = .Codigo
'
'            consultarBienColor
'
'            If Not bolAyuda Then
'                MsgBox "Color Actualizado.", _
'                        vbInformation, App.ProductName
'            Else
'                objAyudaBienColor.Codigo = Trim(txtCodigo.Text)
'                objAyudaBienColor.Descripcion = Trim(txtDescripcion.Text)
'
'                Unload Me
'            End If
        End If
    End With
    
    Set objSqlBienColor = Nothing
End Sub

Private Sub eliminarBienColorSql()
    Set objSqlBienColor = New SqlClsBienColor
    
    With objSqlBienColor
        .Codigo = strCodigo  'Trim(txtCodigo.Text)
'
'        If Not .verificarExistencia Then
'            MsgBox "Color no existente.", vbInformation, App.ProductName
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea Eliminar el Color con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarBienColor Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'
'                strCodigo = .Codigo
'
'                consultarBienColor
'
'                MsgBox "Color Eliminado.", _
'                    vbInformation, App.ProductName
            End If
'        End If
    End With
    
    Set objSqlBienColor = Nothing
End Sub
