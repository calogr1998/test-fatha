VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{791923BA-56CB-4A36-9EA3-1B4ED74622AA}#1.0#0"; "csimxctl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form correo 
   Caption         =   "Servicio de Correo"
   ClientHeight    =   5820
   ClientLeft      =   3240
   ClientTop       =   4005
   ClientWidth     =   8115
   LinkTopic       =   "MainForm"
   ScaleHeight     =   5820
   ScaleWidth      =   8115
   Begin InternetMailCtl.InternetMail InternetMail1 
      Left            =   720
      Top             =   4920
      _cx             =   741
      _cy             =   741
      Enabled         =   -1  'True
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "correo.frx":0000
      ToolBars        =   "correo.frx":0CEA
   End
   Begin VB.CommandButton cmdcopia 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1620
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   4500
      TabIndex        =   16
      Top             =   5640
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5565
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox editMessageHTML 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.TextBox editMessageText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   2460
      Width           =   4995
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4680
      Picture         =   "correo.frx":0D64
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1980
      Width           =   375
   End
   Begin VB.ComboBox comboAttachment 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3495
   End
   Begin VB.TextBox editBcc 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   1620
      Width           =   3495
   End
   Begin VB.TextBox editCc 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1260
      Width           =   3975
   End
   Begin VB.TextBox editSubject 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   900
      Width           =   3975
   End
   Begin VB.TextBox editTo 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   540
      Width           =   3975
   End
   Begin VB.TextBox editFrom 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   3975
   End
   Begin VB.Label lblAttachment 
      Caption         =   "&Adjuntar:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label lblCco 
      Caption         =   "&Cco:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label lblCc 
      Caption         =   "&Cc:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label lblSubject 
      Caption         =   "&Asunto:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblTo 
      Caption         =   "&Para:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblFrom 
      Caption         =   "&De:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.Menu menuFile 
      Caption         =   "&Archivo"
      Visible         =   0   'False
      Begin VB.Menu menuNewMessage 
         Caption         =   "&Nuevo Mensaje"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileAttach 
         Caption         =   "&Adjuntar Archivo"
      End
      Begin VB.Menu menuSaveMessage 
         Caption         =   "&Guardar Mensaje"
      End
      Begin VB.Menu menuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuCloseFile 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu menuFormat 
      Caption         =   "F&ormato"
      Visible         =   0   'False
      Begin VB.Menu menuFormatCharSets 
         Caption         =   "Character Set"
         Begin VB.Menu menuFormatCharSet 
            Caption         =   "Western European"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu menuFormatCharSet 
            Caption         =   "Eastern European"
            Index           =   1
         End
         Begin VB.Menu menuFormatCharSet 
            Caption         =   "Cyrillic"
            Index           =   2
         End
         Begin VB.Menu menuFormatCharSet 
            Caption         =   "Greek"
            Index           =   3
         End
         Begin VB.Menu menuFormatCharSet 
            Caption         =   "Turkish"
            Index           =   4
         End
      End
      Begin VB.Menu menuFormatEncodings 
         Caption         =   "Encoding"
         Begin VB.Menu menuFormatEncoding 
            Caption         =   "7 Bit Text"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu menuFormatEncoding 
            Caption         =   "Quoted Printable"
            Index           =   1
         End
      End
      Begin VB.Menu menuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu menuFormatPlainText 
         Caption         =   "Plain Text"
         Checked         =   -1  'True
      End
      Begin VB.Menu menuFormatStyledText 
         Caption         =   "Styled Text (HTML)"
      End
   End
   Begin VB.Menu menuMessage 
      Caption         =   "&Mensaje"
      Visible         =   0   'False
      Begin VB.Menu menuMessageSend 
         Caption         =   "Enviar Mensaje"
      End
      Begin VB.Menu menuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu menuMessagePriorities 
         Caption         =   "Prioridad"
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Highest"
            Index           =   0
         End
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Alta"
            Index           =   1
         End
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Low"
            Index           =   3
         End
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Lowest"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private bShowOptions As Boolean
Dim strMensaje As String, emailto As String, strFirma As String
Dim sw As Boolean

Private Sub Command1_Click()
    lblCco.Visible = True
    editBcc.Visible = True
End Sub

Private Sub cmdcopia_Click()
    If Not sw Then
        lblCco.Visible = True
        editBcc.Visible = True
        sw = True
    Else
        lblCco.Visible = False
        editBcc.Visible = False
        sw = False
    End If
End Sub

'
Private Sub Form_Load()
'On Error Resume Next
'    Load Options
''    bShowOptions = True
'    Updateparam
'    'SaveOptions
'    menuFormatCharSet(0).Tag = mailCharsetISO8859_1
'    menuFormatCharSet(1).Tag = mailCharsetISO8859_2
'    menuFormatCharSet(2).Tag = mailCharsetISO8859_5
'    menuFormatCharSet(3).Tag = mailCharsetISO8859_7
'    menuFormatCharSet(4).Tag = mailCharsetISO8859_9
'
'    menuFormatEncoding(0).Tag = mailEncoding7Bit
'    menuFormatEncoding(1).Tag = mailEncodingQuoted
'
'    menuMessagePriority(0).Tag = "1 (Highest)"
'    menuMessagePriority(1).Tag = "2 (High)"
'    menuMessagePriority(2).Tag = "3 (Normal)"
'    menuMessagePriority(3).Tag = "4 (Low)"
'    menuMessagePriority(4).Tag = "5 (Lowest)"
'
'    If Len(g_strSenderAddress) > 0 Then
'        editFrom.Text = Chr(34) & g_strSenderName & Chr(34) & " <" & g_strSenderAddress & ">"
'    End If
'    emailto = traerCampo("EF2PROVEEDORES", "F2EMAIL", "F2NEWRUC", wrucprov)
'    editSubject.Text = "Orden de Compra N°:" & Left(Right(strFilePathPDF, 11), 7)
'    strFirma = vbCrLf & vbCrLf & vbCrLf & wnomuser & vbCrLf & wcargo & vbCrLf & wnomcia
'    strMensaje = "Le adjuntamos la " & editSubject.Text & vbCrLf & "Agradecere se sirvan atenderlo." & strFirma
'    editMessageText.Text = strMensaje
'    editTo.Text = emailto
'    comboAttachment.AddItem strFilePathPDF
''    comboAttachment.AddItem "c:\bancowin\temporales\0010000001.pdf"
'    comboAttachment.ListIndex = comboAttachment.ListCount - 1
'    lblCco.Visible = False
'    editBcc.Visible = False
End Sub

Private Sub Form_Activate()
'    If bShowOptions Then
'        frmOptions.Show vbModal, correo
'        bShowOptions = False
'    End If
'    editMessageText.SetFocus
End Sub


Private Sub Form_Resize()
    If correo.Width < 5000 Then correo.Width = 5000
    If correo.Height < 5000 Then correo.Height = 5000
    editFrom.Width = (correo.ScaleWidth - editFrom.left) - 160
    editTo.Width = (correo.ScaleWidth - editTo.left) - 160
    editSubject.Width = (correo.ScaleWidth - editSubject.left) - 160
    editCc.Width = (correo.ScaleWidth - editCc.left) - 160
    editBcc.Width = (correo.ScaleWidth - editBcc.left) - 160
    cmdBrowse.left = (correo.ScaleWidth - cmdBrowse.Width) - 160
    cmdcopia.left = (correo.ScaleWidth - cmdcopia.Width) - 160
    comboAttachment.Width = (cmdBrowse.left - comboAttachment.left) - 160
    editMessageText.Height = (correo.ScaleHeight - editMessageText.top) - StatusBar1.Height - 160
    editMessageText.Width = (correo.ScaleWidth - editMessageText.left) - 160
    editMessageHTML.left = editMessageText.left
    editMessageHTML.top = editMessageText.top
    editMessageHTML.Height = (correo.ScaleHeight - editMessageHTML.top) - StatusBar1.Height - 160
    editMessageHTML.Width = (correo.ScaleWidth - editMessageHTML.left) - 160
    ProgressBar1.left = (correo.ScaleWidth - ProgressBar1.Width) - 160
    ProgressBar1.top = (correo.ScaleHeight - StatusBar1.Height) + 40
End Sub

Private Sub menuNewMessage_Click()
'    InternetMail1.ClearMessage
'
'    editFrom.Text = ""
'    editTo.Text = ""
'    editSubject.Text = ""
'    editCc.Text = ""
'    editBcc.Text = ""
'    editMessageText.Text = ""
'    editMessageHTML.Text = ""
'    comboAttachment.Clear
'
'    If Len(g_strSenderAddress) > 0 Then
'        editFrom.Text = Chr(34) & g_strSenderName & Chr(34) & " <" & g_strSenderAddress & ">"
'    End If
End Sub

'
Private Sub menuSaveMessage_Click()
    Dim nError As Long

    nError = CreateMessage()
    If nError Then
        MsgBox "Unable to create message" & _
               InternetMail1.LastErrorString, vbExclamation
        Exit Sub
    End If

    On Error GoTo ExportCanceled
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = ".txt"
    CommonDialog1.DialogTitle = "Save Message"
    CommonDialog1.FileName = ""
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNLongNames + cdlOFNOverwritePrompt
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|" & _
                       "E-Mail Message Files (*.eml)|*.eml|" & _
                       "All Files (*.*)|*.*"

    CommonDialog1.ShowSave
    On Error GoTo 0

    nError = InternetMail1.ExportMessage(CommonDialog1.FileName)
    If nError Then
        MsgBox "Unable to save message to " & _
               CommonDialog1.FileTitle & vbCrLf & _
               InternetMail1.LastErrorString, vbExclamation
        Exit Sub
    End If

ExportCanceled:
End Sub


Private Sub menuFileAttach_Click()
    Dim nError As Long
    
    On Error GoTo AttachCanceled
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = ".txt"
    CommonDialog1.DialogTitle = "Attach File"
    CommonDialog1.FileName = ""
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNLongNames
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    
    CommonDialog1.ShowOpen
    On Error GoTo 0
    
    comboAttachment.AddItem CommonDialog1.FileName
    comboAttachment.ListIndex = comboAttachment.ListCount - 1
    Exit Sub

AttachCanceled:
End Sub

Private Sub menuCloseFile_Click()
    Unload Me
End Sub

Private Sub menuFormatCharSet_Click(Index As Integer)
    Dim nindex As Integer
    
    For nindex = 0 To 4
        If nindex = Index Then
            menuFormatCharSet(nindex).Checked = True
        Else
            menuFormatCharSet(nindex).Checked = False
        End If
    Next
End Sub

'
Private Sub menuFormatEncoding_Click(Index As Integer)
    Dim nindex As Integer
    
    For nindex = 0 To 1
        If nindex = Index Then
            menuFormatEncoding(nindex).Checked = True
        Else
            menuFormatEncoding(nindex).Checked = False
        End If
    Next
End Sub

'
Private Sub menuFormatPlainText_Click()
    editMessageText.Visible = True
    editMessageHTML.Visible = False
    menuFormatPlainText.Checked = True
    menuFormatStyledText.Checked = False
    StatusBar1.SimpleText = "Enter the plain text version of your message"
End Sub

'
Private Sub menuFormatStyledText_Click()
    editMessageText.Visible = False
    editMessageHTML.Visible = True
    editMessageHTML.Text = strMensaje
    menuFormatPlainText.Checked = False
    menuFormatStyledText.Checked = True
    StatusBar1.SimpleText = "Enter the styled text version of your message"
End Sub


Private Sub menuMessageSend_Click()
    Dim nError As Long
    
    nError = CreateMessage()
    
    If nError Then
        StatusBar1.SimpleText = InternetMail1.LastErrorString
        MsgBox "Unable to create a new message" & vbCrLf & _
               InternetMail1.LastErrorString, vbExclamation
        Exit Sub
    End If
    
    If InternetMail1.Recipients = 0 Then
        MsgBox "There are no recipients for this message", _
               vbInformation
        Exit Sub
    End If

    '
    ProgressBar1.Min = 0
    ProgressBar1.Max = InternetMail1.Recipients
    ProgressBar1.value = 0

'    InternetMail1.RelayServer = "recrisa.com.pe"
'    InternetMail1.Options = mailOptionAuthLogin
'    InternetMail1.UserName = wusermail
'    InternetMail1.Password = wpaswmail
'    InternetMail1.UserName = "recrisa"
'    InternetMail1.Password = "RECRE014"
    
    InternetMail1.RelayServer = "127.0.0.1"
    InternetMail1.Options = mailOptionAuthLogin
    InternetMail1.UserName = "hcmarisol"
    InternetMail1.Password = "4209758"
'    InternetMail1.UserName = wusermail
'    InternetMail1.Password = wpaswmail

    With InternetMail1
        .To = ""
        .CC = ""
        
        .Subject = ""
        
        .AttachFile ""
        
        .Message = ""
        
        '.SendMessage()
    End With

    '
    nError = InternetMail1.SendMessage()
    
    If nError Then
        StatusBar1.SimpleText = InternetMail1.LastErrorString
        MsgBox "Unable to send message" & vbCrLf & _
               InternetMail1.LastErrorString, vbExclamation
        Exit Sub
    End If
    
End Sub

'
Private Sub menuMessagePriority_Click(Index As Integer)
    Dim nindex As Integer
    
    For nindex = 0 To 4
        If nindex = Index Then
            menuMessagePriority(nindex).Checked = True
        Else
            menuMessagePriority(nindex).Checked = False
        End If
    Next

End Sub

'
Private Sub menuMessageOptions_Click()
'    frmOptions.Show vbModal, correo
End Sub

Private Sub cmdBrowse_Click()
    menuFileAttach_Click
End Sub

'
Private Sub InternetMail1_OnDelivered(ByVal Address As Variant, ByVal MessageSize As Variant)
    StatusBar1.SimpleText = "Mensaje enviado para " & Address
    MsgBox "Mensaje Enviado Satisfactoriamente", vbInformation, "Sistema de Ventas"
    Unload Me
End Sub

'
Private Sub InternetMail1_OnRecipient(ByVal Address As Variant, Cancel As Variant)
    StatusBar1.SimpleText = "Enviando mensaje para " & Address
    ProgressBar1.value = ProgressBar1.value + 1
End Sub

'
Private Function CreateMessage() As Long
'    Dim strFontName As String
'    Dim strFontSize As String
'    Dim strMessageHTML As String
'    Dim nCharacterSet As Long
'    Dim nEncodingType As Long
'    Dim nIndex As Long
'    Dim nError As Long
'
'    CreateMessage = 0
'
'    '
'    If Len(Trim(editMessageHTML.Text)) = 0 Then
'        strMessageHTML = ""
'    Else
'        strFontName = "Arial"
'        strFontSize = "3"
'        strMessageHTML = "<html><head><title>" & editSubject.Text & "</title></head>" & _
'                         "<body><font " & _
'                         "face=" & Chr(34) & strFontName & Chr(34) & " " & _
'                         "size=" & Chr(34) & strFontSize & Chr(34) & ">" & vbCrLf & _
'                         editMessageHTML.Text & vbCrLf & _
'                         "</font></body></html>"
'    End If
'
'    '
'    nCharacterSet = mailCharsetISO8859_1
'    For nIndex = 0 To 4
'        If menuFormatCharSet(nIndex).Checked Then
'            nCharacterSet = menuFormatCharSet(nIndex).Tag
'            Exit For
'        End If
'    Next
'
'    '
'    nEncodingType = mailEncoding7Bit
'    For nIndex = 0 To 1
'        If menuFormatEncoding(nIndex).Checked Then
'            nEncodingType = menuFormatEncoding(nIndex).Tag
'            Exit For
'        End If
'    Next
'
'    '
'    nError = InternetMail1.ComposeMessage(editFrom.Text, _
'                                          editTo.Text, _
'                                          editCc.Text, _
'                                          editBcc.Text, _
'                                          editSubject.Text, _
'                                          editMessageText.Text, _
'                                          strMessageHTML, _
'                                          nCharacterSet, _
'                                          nEncodingType)
'
'    If nError Then
'        CreateMessage = nError
'        Exit Function
'    End If
'
'
'    '
'    For nIndex = 0 To comboAttachment.ListCount - 1
'        nError = InternetMail1.AttachFile(comboAttachment.List(nIndex))
'        If nError Then
'            CreateMessage = nError
'            Exit Function
'        End If
'    Next
'
'    '
'    For nIndex = 0 To 4
'        If menuMessagePriority(nIndex).Checked Then
'            InternetMail1.Priority = menuMessagePriority(nIndex).Tag
'            Exit For
'        End If
'    Next
'
'    '
'    InternetMail1.Organization = g_strOrganization
    
End Function



Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Enviar"
            menuMessageSend_Click
    End Select
End Sub
