VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{791923BA-56CB-4A36-9EA3-1B4ED74622AA}#1.0#0"; "csimxctl.ocx"
Begin VB.Form MailSend 
   BorderStyle     =   0  'None
   Caption         =   "Internet Mail Control - SendMail Example"
   ClientHeight    =   5190
   ClientLeft      =   2955
   ClientTop       =   3375
   ClientWidth     =   7620
   LinkTopic       =   "MainForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InternetMailCtl.InternetMail InternetMail1 
      Left            =   600
      Top             =   4920
      _cx             =   741
      _cy             =   741
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   5400
      TabIndex        =   16
      Top             =   4860
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   4755
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   767
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
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   2460
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
      Width           =   3675
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4680
      Picture         =   "MailSend.frx":0000
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
      Width           =   3975
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
      Caption         =   "&Attachment:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label lblBcc 
      Caption         =   "&Bcc:"
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
      Caption         =   "&Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblTo 
      Caption         =   "&To:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblFrom 
      Caption         =   "&From:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu menuNewMessage 
         Caption         =   "&New Message"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileAttach 
         Caption         =   "&Attach File"
      End
      Begin VB.Menu menuSaveMessage 
         Caption         =   "&Save Message"
      End
      Begin VB.Menu menuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuCloseFile 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu menuFormat 
      Caption         =   "F&ormat"
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
      Caption         =   "&Message"
      Visible         =   0   'False
      Begin VB.Menu menuMessageSend 
         Caption         =   "Send Message"
      End
      Begin VB.Menu menuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu menuMessagePriorities 
         Caption         =   "Priority"
         Begin VB.Menu menuMessagePriority 
            Caption         =   "Highest"
            Index           =   0
         End
         Begin VB.Menu menuMessagePriority 
            Caption         =   "High"
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
      Begin VB.Menu menuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu menuMessageOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "MailSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Catalyst Internet Mail Control 4.5
' Copyright 2002-2006, Catalyst Development Corporation
' All rights reserved
'
' This product is licensed to you pursuant to the terms of the
' Catalyst license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'
Option Explicit

Private bShowOptions As Boolean

'
' Initialize the controls on the form and load any options
' that were stored (see options.bas module)
'
Private Sub Form_Load()
    CargaOptions
    'bShowOptions = True ' Show the Options dialog on startup
    
    menuFormatCharSet(0).Tag = mailCharsetISO8859_1
    menuFormatCharSet(1).Tag = mailCharsetISO8859_2
    menuFormatCharSet(2).Tag = mailCharsetISO8859_5
    menuFormatCharSet(3).Tag = mailCharsetISO8859_7
    menuFormatCharSet(4).Tag = mailCharsetISO8859_9
    
    menuFormatEncoding(0).Tag = mailEncoding7Bit
    menuFormatEncoding(1).Tag = mailEncodingQuoted

    menuMessagePriority(0).Tag = "1 (Highest)"
    menuMessagePriority(1).Tag = "2 (High)"
    menuMessagePriority(2).Tag = "3 (Normal)"
    menuMessagePriority(3).Tag = "4 (Low)"
    menuMessagePriority(4).Tag = "5 (Lowest)"

    If Len(g_strSenderAddress) > 0 Then
        editFrom.Text = g_strSenderAddress
    End If
    
    Call menuMessageSend_Click
    
    
End Sub

Private Sub CargaOptions()

g_strSenderName = "Infoplus"
g_strSenderAddress = "sistema.solicitudes@gmail.com"
g_strOrganization = wnomcia
g_bRelayMessages = False
'g_strRelayServer = ObtenerCampo("EMAIL", "server_e", "TIPO", "E", "T", cnn_Envia)
g_strRelayServer = "smtp.gmail.com"

'g_nRelayPort = ObtenerCampo("EMAIL", "port_e", "TIPO", "E", "T", cnn_Envia)
g_nRelayPort = 465
comboAttachment.Clear
comboAttachment.AddItem wFileName
comboAttachment.ListIndex = 0
editTo.Text = wDestinatarios
editBcc.Text = ""
editSubject.Text = wSubject
editMessageText.Text = wSubject

InternetMail1.UserName = g_strSenderAddress
InternetMail1.Password = "infoplus1234"
bShowOptions = False
End Sub
Private Sub Form_Activate()
    If bShowOptions Then
       ' MailOptions.Show vbModal, MailSend
        bShowOptions = False
    End If
    
End Sub

'
' Resize controls when the user resizes the form
'
Private Sub Form_Resize()
'    If frmMain.Width < 5000 Then frmMain.Width = 5000
'    If frmMain.Height < 5000 Then frmMain.Height = 5000
'    editFrom.Width = (frmMain.ScaleWidth - editFrom.left) - 160
'    editTo.Width = (frmMain.ScaleWidth - editTo.left) - 160
'    editSubject.Width = (frmMain.ScaleWidth - editSubject.left) - 160
'    editCc.Width = (frmMain.ScaleWidth - editCc.left) - 160
'    editBcc.Width = (frmMain.ScaleWidth - editBcc.left) - 160
'    cmdBrowse.left = (frmMain.ScaleWidth - cmdBrowse.Width) - 160
'    comboAttachment.Width = (cmdBrowse.left - comboAttachment.left) - 160
'    editMessageText.Height = (frmMain.ScaleHeight - editMessageText.top) - StatusBar1.Height - 160
'    editMessageText.Width = (frmMain.ScaleWidth - editMessageText.left) - 160
'    editMessageHTML.left = editMessageText.left
'    editMessageHTML.top = editMessageText.top
'    editMessageHTML.Height = (frmMain.ScaleHeight - editMessageHTML.top) - StatusBar1.Height - 160
'    editMessageHTML.Width = (frmMain.ScaleWidth - editMessageHTML.left) - 160
'    ProgressBar1.left = (frmMain.ScaleWidth - ProgressBar1.Width) - 160
'    ProgressBar1.top = (frmMain.ScaleHeight - StatusBar1.Height) + 40
End Sub
'
' Clear the current message
'
Private Sub menuNewMessage_Click()
    InternetMail1.ClearMessage
    
    editFrom.Text = ""
    editTo.Text = ""
    editSubject.Text = ""
    editCc.Text = ""
    editBcc.Text = ""
    editMessageText.Text = ""
    editMessageHTML.Text = ""
    comboAttachment.Clear
    
    If Len(g_strSenderAddress) > 0 Then
        editFrom.Text = Chr(34) & g_strSenderName & Chr(34) & " <" & g_strSenderAddress & ">"
    End If
End Sub
'
' Save the current message to a file. The CreateMessage helper function
' is called, and then the user is prompted for a file to save the
' message to
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
'
' Select a file to be attached to the message and add it to the
' combo box control which lists the current file attachments.
' Note that this does not actually attach the file to the current
' message; that is done in the CreateMessage helper function
'
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
'
' Select a character set that will be used when composing the
' message; the default is ISO-8859-1 which is used by Western
' European languages
'
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
' Select an encoding type for the message; the default is
' 7-bit plain text.
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
' Display the multiline edit control which contains the plain text
' version of the message
'
Private Sub menuFormatPlainText_Click()
    editMessageText.Visible = True
    editMessageHTML.Visible = False
    menuFormatPlainText.Checked = True
    menuFormatStyledText.Checked = False
    StatusBar1.SimpleText = "Enter the plain text version of your message"
End Sub
'
' Display the multiline edit control which contains the styled (HTML)
' text version of the message. See the CreateMessage helper function
' for how the sample handles HTML text. A more sophisticated approach
' would be to actually create an HTML editor, but that is beyond the
' scope of this example.
'
Private Sub menuFormatStyledText_Click()
    editMessageText.Visible = False
    editMessageHTML.Visible = True
    menuFormatPlainText.Checked = False
    menuFormatStyledText.Checked = True
    StatusBar1.SimpleText = "Enter the styled text version of your message"
End Sub

' Call the CreateMessage helper function to create the message and
' then deliver it to the specified recipients
'
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
    ' Update the progress bar control with the total number of
    ' recipients so that we can display the progress of the
    ' delivery process; see the OnRecipient and OnDeliver events
    ' for the InternetMail control
    '
    ProgressBar1.Min = 0
    ProgressBar1.Max = InternetMail1.Recipients
    ProgressBar1.Value = 0
    
    '
    ' The user has specified that a relay server is to be used
    ' (under the Message | Options menu) then set those properties
    ' so that our message will always be routed through that server
    '
    'If g_bRelayMessages Then
        InternetMail1.RelayServer = g_strRelayServer
        InternetMail1.RelayPort = g_nRelayPort
    'Else
     '   InternetMail1.RelayServer = ""
     '   InternetMail1.RelayPort = 0
    'End If
    InternetMail1.Secure = True
    InternetMail1.Options = 5
    '
    ' Begin the process of delivering the message
    '
    nError = InternetMail1.SendMessage()
    If nError Then
        StatusBar1.SimpleText = InternetMail1.LastErrorString
        MsgBox "Unable to send message" & vbCrLf & _
               InternetMail1.LastErrorString, vbExclamation
        
        Exit Sub
    End If
    InternetMail1.Disconnect
    
    Unload Me
End Sub
'
' Select a priority for the message; the default is normal
' priority
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
' Display the options dialog which allows the user to enter
' some default values and select whether or not a relay server
' should be used to deliver messages
'
Private Sub menuMessageOptions_Click()
    'frmOptions.Show vbModal, frmMain
End Sub

Private Sub cmdBrowse_Click()
    menuFileAttach_Click
End Sub
'
' This event is called immediately after a message has been successfully
' delivered to a recipient
'
Private Sub InternetMail1_OnDelivered(ByVal Address As Variant, ByVal MessageSize As Variant)
    StatusBar1.SimpleText = "Message delivered to " & Address
End Sub
'
' This event is called immediately before a message is delivered to
' the specified recipient; if the Cancel argument is set to True, then
' the message will not be delivered
'
Private Sub InternetMail1_OnRecipient(ByVal Address As Variant, Cancel As Variant)
    StatusBar1.SimpleText = "Sending message to " & Address
    ProgressBar1.Value = ProgressBar1.Value + 1
End Sub
'
' A helper function which uses the ComposeMessage method to create a
' message, followed by attaching each of the listed files to the
' message. This function is called when the user selects the menu option
' to send a message or save the message to a text file
'
Private Function CreateMessage() As Long
    Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long
        
    CreateMessage = 0
    
    '
    ' If the user has entered any HTML text, then make sure that it is
    ' properly formed HTML and specify a font and font size to use;
    ' this is hard-coded, but an application would obviously want to
    ' make something like font selection customizable
    '
    If Len(Trim(editMessageHTML.Text)) = 0 Then
        strMessageHTML = ""
    Else
        strFontName = "Arial"
        strFontSize = "3"
        strMessageHTML = "<html><head><title>" & editSubject.Text & "</title></head>" & _
                         "<body><font " & _
                         "face=" & Chr(34) & strFontName & Chr(34) & " " & _
                         "size=" & Chr(34) & strFontSize & Chr(34) & ">" & vbCrLf & _
                         editMessageHTML.Text & vbCrLf & _
                         "</font></body></html>"
    End If
    
    '
    ' Determine what character set was selected by the user
    '
    nCharacterSet = mailCharsetISO8859_1
    For nindex = 0 To 4
        If menuFormatCharSet(nindex).Checked Then
            nCharacterSet = menuFormatCharSet(nindex).Tag
            Exit For
        End If
    Next
    
    '
    ' Determine what encoding type was selected by the user
    '
    nEncodingType = mailEncoding7Bit
    For nindex = 0 To 1
        If menuFormatEncoding(nindex).Checked Then
            nEncodingType = menuFormatEncoding(nindex).Tag
            Exit For
        End If
    Next
    
    '
    ' Use the ComposeMessage method to do all of the hard work of
    ' creating the actual message
    '
    nError = InternetMail1.ComposeMessage(editFrom.Text, _
                                          editTo.Text, _
                                          editCc.Text, _
                                          editBcc.Text, _
                                          editSubject.Text, _
                                          editMessageText.Text, _
                                          strMessageHTML, _
                                          nCharacterSet, _
                                          nEncodingType)
    
    If nError Then
        CreateMessage = nError
        Exit Function
    End If
    
    '
    ' Attach each file that was selected by the user
    '
    For nindex = 0 To comboAttachment.ListCount - 1
        nError = InternetMail1.AttachFile(comboAttachment.List(nindex))
        If nError Then
            CreateMessage = nError
            Exit Function
        End If
    Next
    
    '
    ' Set the Priority property to the message priority that
    ' was selected by the user
    '
    For nindex = 0 To 4
        If menuMessagePriority(nindex).Checked Then
            InternetMail1.Priority = menuMessagePriority(nindex).Tag
            Exit For
        End If
    Next

    '
    ' Set the Organization property to the name of the current user's
    ' organization or company; this is specified in the options dialog
    '
    InternetMail1.Organization = g_strOrganization
    
End Function



