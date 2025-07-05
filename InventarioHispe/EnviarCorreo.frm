VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EnviarCorreoOcxAfisca 
   Caption         =   "Enviar Correo"
   ClientHeight    =   7335
   ClientLeft      =   840
   ClientTop       =   2775
   ClientWidth     =   10200
   Icon            =   "EnviarCorreo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10200
   Begin VB.Frame Frame5 
      Caption         =   " Datos de Email "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "aprobaciones@betania.com.pe"
         Top             =   360
         Width           =   8280
      End
      Begin VB.TextBox txtAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   7800
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   1
         ToolTipText     =   "Clic aqui, para agregar archivos adjuntos..."
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adjunto"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Remotos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9975
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   8520
         PasswordChar    =   "*"
         TabIndex        =   19
         Text            =   "4pr0.b3t4"
         Top             =   360
         Width           =   1170
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   18
         Text            =   "aprobaciones@betania.com.pe"
         Top             =   360
         Width           =   2970
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "mail.betania.com.pe"
         Top             =   360
         Width           =   2400
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   195
         Left            =   7680
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor SMTP"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   9975
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   8760
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Enviar Correo"
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   5280
         Width           =   1635
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   8760
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   4140
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   960
         Width           =   9000
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Salir"
         Height          =   315
         Left            =   8280
         TabIndex        =   10
         Top             =   5280
         Width           =   1275
      End
      Begin VB.PictureBox sendmail1 
         Height          =   480
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   23
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblProgresoPorcentaje 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label lblProgresoEstado 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   8160
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   15
      Begin VB.ListBox lstStatus 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   7800
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progreso"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "EnviarCorreoOcxAfisca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Importante: El ocx produce un pequeño Bug y es que no se debe _
establecer la propiedad Visible en False. Por lo tanto para que no se _
vea el ocx en tiempo de ejecución, se estableció el Left y el Top fuera del _
area del form en el Form_Load

'Si tenes idea porque puede ocurrir esto me podes enviar un mail a _
info@recursosvisualbasic.com.ar para ver si lo puedo arreglar y volver a _
compilar.

Option Explicit
Option Compare Text

Private strServidorSMTP             As String
Private strEmailRemitente           As String
Private strEmailContrasena          As String
Private strNombreRemitente          As String
Private strRutaArchivoAdjunto       As String
Private strEmailDestinatario        As String
Private strEmailDestinatarioCC      As String
Private strEmailDestinatarioCCO     As String
Private strEmailAsunto              As String
Private strEmailCuerpo              As String

Private bolEmailEnviado             As Boolean

Private strFichero                  As String

Public Property Let ServidorSMTP(ByVal Value As String)
    strServidorSMTP = Value
End Property

Public Property Get ServidorSMTP() As String
    ServidorSMTP = strServidorSMTP
End Property

Public Property Let EmailRemitente(ByVal Value As String)
    strEmailRemitente = Value
End Property

Public Property Get EmailRemitente() As String
    EmailRemitente = strEmailRemitente
End Property

Public Property Let EmailContrasena(ByVal Value As String)
    strEmailContrasena = Value
End Property

Public Property Get EmailContrasena() As String
    EmailContrasena = strEmailContrasena
End Property

Public Property Let NombreRemitente(ByVal Value As String)
    strNombreRemitente = Value
End Property

Public Property Get NombreRemitente() As String
    NombreRemitente = strNombreRemitente
End Property

Public Property Let RutaArchivoAdjunto(ByVal Value As String)
    strRutaArchivoAdjunto = Value
End Property

Public Property Get RutaArchivoAdjunto() As String
    RutaArchivoAdjunto = strRutaArchivoAdjunto
End Property

Public Property Let EmailDestinatario(ByVal Value As String)
    strEmailDestinatario = Value
End Property

Public Property Get EmailDestinatario() As String
    EmailDestinatario = strEmailDestinatario
End Property

Public Property Let EmailDestinatarioCC(ByVal Value As String)
    strEmailDestinatarioCC = Value
End Property

Public Property Get EmailDestinatarioCC() As String
    EmailDestinatarioCC = strEmailDestinatarioCC
End Property

Public Property Let EmailDestinatarioCCO(ByVal Value As String)
    strEmailDestinatarioCCO = Value
End Property

Public Property Get EmailDestinatarioCCO() As String
    EmailDestinatarioCCO = strEmailDestinatarioCCO
End Property

Public Property Let EmailAsunto(ByVal Value As String)
    strEmailAsunto = Value
End Property

Public Property Get EmailAsunto() As String
    EmailAsunto = strEmailAsunto
End Property

Public Property Let EmailCuerpo(ByVal Value As String)
    strEmailCuerpo = Value
End Property

Public Property Get EmailCuerpo() As String
    EmailCuerpo = strEmailCuerpo
End Property



'EmailEnviado
Public Property Let EmailEnviado(ByVal Value As Boolean)
    bolEmailEnviado = Value
End Property

Public Property Get EmailEnviado() As Boolean
    EmailEnviado = bolEmailEnviado
End Property


Private Sub cmdSend_Click()
'    If Trim(txtFrom.Text) = vbNullString Then
'        MsgBox "Indique el correo electronico del Remitente.", vbInformation + vbOKOnly, App.ProductName
'
'        txtFrom.SetFocus
'
'        Exit Sub
'    End If
'
'    If Dir(Trim(txtAttach.Text), vbArchive) = vbNullString Then
'        MsgBox "Ruta de archivo adjunto invalida, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'        txtFrom.SetFocus
'
'        Exit Sub
'    End If
'
'    If Trim(txtTo.Text) = vbNullString Then
'        MsgBox "Indique el correo electronico del Destinatario.", vbInformation + vbOKOnly, App.ProductName
'
'        txtTo.SetFocus
'
'        Exit Sub
'    End If
'
'    If Trim(txtSubject.Text) = vbNullString Then
'        If MsgBox("No se ha consignado ningún asunto para el Correo, ¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'            txtSubject.SetFocus
'
'            Exit Sub
'        End If
'    End If
'
'    cmdSend.Enabled = False
'
'    Screen.MousePointer = vbHourglass
'
'    With sendmail1
'        'Valida (opcional)
'        .SMTPHostValidacion = VALIDATE_HOST_NONE
'        'Valida la sintaxis de l cuenta (opcional)
'        .ValidarEmail = VALIDATE_SYNTAX
'        'Opcional
'        .Delimitador = ";"
'        'Texto  para visualizar en el campo De (opcional)
'        .FromDisplayName = strNombreRemitente '" Ejemplo "
'        'Requerido (Nombre del servidor SMTP)
'        .SMTPHost = txtServer.Text
'        'Requerido
'        .Remitente = txtFrom.Text
'        'Requerido
'        .Destinatario = txtTo.Text
'        .CcRecipient = strEmailDestinatarioCC
'        .BccRecipient = strEmailDestinatarioCCO
'        'Asunto del mensaje
'        .asunto = txtSubject.Text
'        'Cuerpodel mensaje
'        .mensaje = txtMsg.Text
'        'Adjunto (opcional)
'        .Adjunto = Trim(txtAttach.Text)
'
'        'Opcional (Prioridad del mensaje)
'        .Prioridad = Alta 'Baja
'        'Opcional (si requiere autentificación)
'        .UsarLoginSMTP = True
'        'Requerido si requiere autentificación
'        .Usuario = txtUserName.Text
'        .Password = txtPassword.Text
'
'        txtServer.Text = .SMTPHost
'        'Opcional (por defectoutiliza el Tipo MIME)
'        .Codificacion = MIME_ENCODE
'
'        'Envia el Mail
'        .EnviarEmail
'
'        strEmailDestinatario = Trim(txtTo.Text)
'    End With
'
'    lblProgresoEstado.Caption = vbNullString
'    lblProgresoPorcentaje.Caption = vbNullString
'
'    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    'Codigo por Default
    sendmail1.Move -1000, -1000
    
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    strFichero = App.Path & strNombreFicheroConfigCPgeneral
    
    txtServer.Text = ModUtilitario.sGetINI(strFichero, "ConfigCP", "ServidorCorreoSalienteSMTP", "l")
    
    txtUserName.Text = strEmailRemitente
    txtPassword.Text = strEmailContrasena
    
    txtFrom.Text = strEmailRemitente
    txtAttach.Text = strRutaArchivoAdjunto
    
    txtTo.Text = strEmailDestinatario
    
    strEmailDestinatarioCCO = ModUtilitario.sGetINI(strFichero, "ConfigCP", "EmailCCOpredeterminada", "l")
    
    txtSubject.Text = strEmailAsunto
    
    txtMsg.Text = strEmailCuerpo
    
    lblProgresoEstado.Caption = vbNullString
    lblProgresoPorcentaje.Caption = vbNullString
    
    bolEmailEnviado = False
End Sub

Private Sub sendmail1_SendSuccesful()
    MsgBox "Mensaje enviado correctamente.", vbInformation + vbOKOnly, App.ProductName
    
    lblProgresoPorcentaje.Caption = vbNullString
    
    bolEmailEnviado = True
End Sub

Private Sub sendmail1_Progress(lPercentCompete As Long)
    'Visualiza el porcentaje del progreso del envío en el Label
    lblProgresoPorcentaje.Caption = lPercentCompete & "% completado."
End Sub

Private Sub sendmail1_SendFailed(Explanation As String)
    'En caso de fallar el envío se dispara este evento con la descripción del error
    MsgBox "El envío del Email falló por las posibles razones: " & vbNewLine & _
            Explanation, vbInformation + vbOKOnly, App.ProductName
    
    lblProgresoEstado.Caption = vbNullString
    lblProgresoPorcentaje.Caption = vbNullString
    
    Screen.MousePointer = vbDefault
    
    cmdSend.Enabled = True
    
    bolEmailEnviado = False
End Sub

Private Sub sendmail1_Status(Status As String)
    'Visualiza el estado del envío
    lblProgresoEstado.Caption = Status
End Sub
'Para los adjuntos
Private Sub cmdBrowse_Click()
    On Local Error GoTo errSub
    
    Dim ArchivosAdj()    As String
    Dim i               As Integer
    
    With CommonDialog1
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|Archivos PDF (*.pdf)|*.pdf"
        .FilterIndex = 1
        .DialogTitle = "Seleccionar archivos..."
        .MaxFileSize = &H7FFF
        .Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ArchivosAdj = Split(.FileName, vbNullChar)
    End With

    If UBound(ArchivosAdj) = 0 Then
        If txtAttach.Text = vbNullString Then
            txtAttach.Text = ArchivosAdj(0)
        Else
            txtAttach.Text = txtAttach.Text & ";" & ArchivosAdj(0)
        End If
    ElseIf UBound(ArchivosAdj) > 0 Then
        If right$(ArchivosAdj(0), 1) <> "\" Then ArchivosAdj(0) = ArchivosAdj(0) & "\"
        
        For i = 1 To UBound(ArchivosAdj)
            If txtAttach.Text = "" Then
                txtAttach.Text = ArchivosAdj(0) & ArchivosAdj(i)
            Else
                txtAttach.Text = txtAttach.Text & ";" & ArchivosAdj(0) & ArchivosAdj(i)
            End If
        Next
    Else
        Exit Sub
    End If
    
    Exit Sub
errSub:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    txtAttach.Text = vbNullString
    
    Err.Clear
End Sub
