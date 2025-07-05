VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5130
   ClientLeft      =   7950
   ClientTop       =   4185
   ClientWidth     =   4320
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030.974
   ScaleMode       =   0  'User
   ScaleWidth      =   4056.247
   Begin VB.PictureBox imgEmpresa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   4095
      TabIndex        =   16
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Frame frmTipoCambio 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Tipo Cambio "
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
      Begin VB.PictureBox picAyudaCambios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         Picture         =   "frmLogin.frx":058A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Tag             =   "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
         ToolTipText     =   "Abrir ayuda Tipo de Cambios - según SUNAT"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtVenta 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2400
         TabIndex        =   6
         Text            =   "0.000"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCompra 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   720
         TabIndex        =   5
         Text            =   "0.000"
         Top             =   840
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483644
         CalendarTrailingForeColor=   -2147483639
         Format          =   138280961
         CurrentDate     =   41134
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Venta"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Compra"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Fecha"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbEmpresa 
      Height          =   315
      ItemData        =   "frmLogin.frx":0B14
      Left            =   1560
      List            =   "frmLogin.frx":0B16
      TabIndex        =   0
      Text            =   "Tower_2020"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1530
      TabIndex        =   1
      Top             =   600
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2760
      TabIndex        =   8
      Top             =   2760
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   990
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Empresa"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   615
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   10
      Top             =   1005
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLogin As New ADODB.Recordset
Dim intFallas As Integer
'--- API para obtener el nombre del equipo
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cmbEmpresa_LostFocus()
    Dim bolValidacion As Boolean
    Dim oEmpresa    As String
    If cmbEmpresa.ListIndex <> -1 Or cmbEmpresa.Text <> Empty Then
        bolValidacion = cargarParametrosEmpresa(Trim(cmbEmpresa.Text & ""))
        
        If bolValidacion Then
            txtUserName.Enabled = True: txtPassword.Enabled = True: cmdOK.Enabled = True
            frmTipoCambio.Enabled = Not verificarTipoCambio(Trim(dtpFecha.Value), txtCompra, txtVenta)
            
            oEmpresa = Mid(cmbEmpresa.Text, 1, InStr(1, cmbEmpresa.Text, "_") - 1)
            With imgEmpresa
                .Picture = LoadPicture(App.Path & "\Logo" & oEmpresa & ".bmp")
                
            End With
            txtUserName.SetFocus
        Else
            cmbEmpresa.SetFocus
        End If
    Else
        MsgBox "Seleccione Empresa.", vbInformation, App.ProductName
        
        cmbEmpresa.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    If cnRutas.State = 1 Then cnRutas.Close
    If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
    
    End
End Sub

Private Sub cmdOk_Click()
    'Validar Cajas
    If cmbEmpresa.Text = Empty Then
        MsgBox "Seleccione la Empresa.", vbInformation, App.ProductName
        cmbEmpresa.SetFocus
        
        Exit Sub
    End If
    
    If txtUserName.Text = Empty Then
        MsgBox "Ingrese su Nombre de usuario.", vbInformation, App.ProductName
        txtUserName.SetFocus
        
        Exit Sub
    End If
    
    If txtPassword.Text = Empty Then
        MsgBox "Ingrese su Contraseña.", vbInformation, App.ProductName
        txtPassword.SetFocus
        
        Exit Sub
    End If
    
    If frmTipoCambio.Enabled Then
        If Val(txtCompra.Text) = 0 Or Val(txtVenta.Text) = 0 Then
            MsgBox "Los campos 'Compra' y 'Venta' son obligatorios," & vbNewLine & _
                    "verifique el ingreso correcto de estos datos.", vbInformation, App.ProductName
            
            txtCompra.SetFocus
        
            Exit Sub
        End If
    End If
    
    wusuario = txtUserName.Text
    validarSesion Trim(txtUserName.Text), Trim(txtPassword.Text)
    
End Sub

Private Sub Form_Load()
    Me.BackColor = vbWhite
    Me.lblLabels(0).BackColor = vbWhite
    Me.lblLabels(1).BackColor = vbWhite
    
    Me.lblLabels(2).BackColor = vbWhite
    Me.Label1.BackColor = vbWhite
    Me.Label2.BackColor = vbWhite
    Me.Label3.BackColor = vbWhite
    frmTipoCambio.BackColor = vbWhite
    
    limpiarCajas
    listarEmpresas cmbEmpresa
End Sub

Public Sub listarEmpresas(ByVal combo As ComboBox)
On Error GoTo errListarEmpresas
Dim XW As String
Dim x As Integer
    If rsRutas.State = 1 Then rsRutas.Close
    
    rsRutas.Open "Select Empresa From Srutas order by Empresa", cnRutas, adOpenForwardOnly, adLockReadOnly
    x = 0
    If Not rsRutas.EOF Then
        combo.Clear
        
        rsRutas.MoveFirst
        
        Do While Not rsRutas.EOF
            combo.AddItem rsRutas!Empresa
            XW = rsRutas!Empresa
            rsRutas.MoveNext
        Loop
    End If
    
    rsRutas.Close
    
    Set rsRutas = Nothing
    combo.Text = XW
    Exit Sub
errListarEmpresas:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbCritical + vbOKOnly, App.ProductName & " - ModInfoPlus: Proc. ListarEmpresas"
    
    Err.Clear
End Sub

Private Sub picAyudaCambios_Click()
    abrirPaginaWeb picAyudaCambios.Tag
End Sub

Private Sub txtcompra_GotFocus()
    seleccionarTextoCaja txtCompra
End Sub

Private Sub txtCompra_LostFocus()
    On Error Resume Next
    
    If IsNumeric(txtCompra.Text) Then
        txtCompra.Text = Format(txtCompra.Text, "#0.000")
    Else
        MsgBox "El campo 'Compra' debe ser númerico.", vbCritical + vbOKOnly, App.ProductName
        
        txtCompra.SetFocus
    End If
End Sub

Private Sub txtpassword_GotFocus()
    seleccionarTextoCaja txtPassword
End Sub

Private Sub txtUserName_GotFocus()
    seleccionarTextoCaja txtUserName
End Sub

Private Sub txtventa_GotFocus()
    seleccionarTextoCaja txtVenta
End Sub

Private Sub txtventa_LostFocus()
    On Error Resume Next
    
    If IsNumeric(txtVenta.Text) Then
        txtVenta.Text = Format(txtVenta.Text, "#0.000")
    Else
        MsgBox "El campo 'Venta' debe ser númerico.", vbCritical + vbOKOnly, App.ProductName
        
        txtVenta.SetFocus
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    KeyAscii = validarKeyPress(KeyAscii)
End Sub

Private Sub limpiarCajas()
    'txtUserName.Text = Empty
    'txtPassword.Text = Empty
    dtpFecha.Value = Format(Now, "Short Date")
    txtCompra.Text = TCCompra
    txtVenta.Text = TCVenta
End Sub

Private Sub grabarTipoCambio()
    Dim Amov_TC(0 To 5) As a_grabacion
    
    Amov_TC(0).campo = "CAMBIO":        Amov_TC(0).valor = Val(txtVenta.Text):  Amov_TC(0).Tipo = "N"
    Amov_TC(1).campo = "CAMBIO_VENTA":  Amov_TC(1).valor = Val(txtVenta.Text):  Amov_TC(1).Tipo = "N"
    Amov_TC(2).campo = "CAMBIOCOMP":    Amov_TC(2).valor = Val(txtCompra.Text): Amov_TC(2).Tipo = "N"
    Amov_TC(3).campo = "FECHA":         Amov_TC(3).valor = dtpFecha.Value:      Amov_TC(3).Tipo = "F"
    Amov_TC(4).campo = "F2CODUSER":     Amov_TC(4).valor = datosUser.codUser:   Amov_TC(4).Tipo = "T"
    Amov_TC(5).campo = "HORAREG":     Amov_TC(5).valor = Format(Time, "Long Time"):  Amov_TC(5).Tipo = "H"
    StrConexDbBancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_bancos.MDB;Persist Security Info=False"
    GRABA_REGISTRO Amov_TC, "CAMBIOS", "A", 5, StrConexDbBancos, ""
End Sub

Private Sub validarSesion(ByVal Usuario As String, pass As String)
    Dim mensaje As String
    
    If rsLogin.State = 1 Then rsLogin.Close
    
    rsLogin.Open "SELECT F2CODUSER, F2NOMUSER, F2DIRUSER, CENTROCOSTO FROM EF2USERS WHERE F2CODUSER = '" & Usuario & "' AND F2PASUSER = '" & pass & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsLogin.EOF Then
        datosUser.codUser = Trim(rsLogin!F2CODUSER & "")
        datosUser.nomUser = Trim(rsLogin!F2NOMUSER & "")
        datosUser.CentroCosto = Trim(rsLogin!CentroCosto & "")
'        If Trim(rsLogin!CentroCosto & "") = "002" Then
'            wSucursal = "PLANTA"
'        Else
'            wSucursal = "OFICINA"
'        End If

        If ComputerName = "SVR-DATA" Then
            wrutatemp = wrutabancos & "\TEMPORALES\BANCOWIN\" & Usuario & "\"
            
            wbasetemp = wrutatemp & "\BASETEMPVENTAS.MDB"
            wfestemp = wrutatemp & "\FES21.MDB"
        Else
            wbasetemp = wrutatemp & "\BASETEMPVENTAS.MDB"
            wfestemp = wrutatemp & "\FES21.MDB"
        End If
        
'        If Len(Trim(rsLogin!F2DIRUSER)) > 0 Then
'            wrutatemp = wrutabancos & "\TEMPORALES\BANCOWIN\" & Usuario & "\"
'
'            wbasetemp = wrutatemp & "\BASETEMPVENTAS.MDB"
'            wfestemp = wrutatemp & "\FES21.MDB"
'        Else
'            wbasetemp = wrutatemp & "\BASETEMPVENTAS.MDB"
'            wfestemp = wrutatemp & "\FES21.MDB"
'        End If

        
        If frmTipoCambio.Enabled Then
            grabarTipoCambio
        End If
        
        MsgBox "Bienvenido(a), " & Trim(rsLogin!F2NOMUSER) & Chr(13) & _
                    "puede iniciar sus labores.", vbInformation + vbOKOnly, App.ProductName & " - Inicio de Sesión"
                 
        Me.Hide
        
'        Lista_Seguimiento.Show
        'mdiVenta.Show
        menu.Show
        
        Unload Me
    Else
        intFallas = intFallas + 1
        
        If intFallas > 3 Then
            MsgBox "Nombre de Usuario ó Contraseña invalidos," & Chr(13) & _
                    "Ud. a superado el margen de intentos fallidos.", vbCritical + vbOKOnly, App.ProductName & " - Error en Datos"
                    
            If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
            
            End
        End If
        
        MsgBox "Nombre de Usuario ó Contraseña invalidos," & Chr(13) & _
                    "vuelva a intentarlo.", vbCritical + vbOKOnly, App.ProductName & " - Error en Datos"
    End If
    
    If rsLogin.State = 1 Then rsLogin.Close
    
    Set rsLogin = Nothing
End Sub

Public Function ComputerName() As String
  '-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
  ComputerName = Space$(260)
  GetComputerName ComputerName, Len(ComputerName)
  ComputerName = left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)
End Function

