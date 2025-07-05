VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4215
   ClientLeft      =   2670
   ClientTop       =   2055
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "OptionsForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameSender 
      Caption         =   "Sender"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6075
      Begin VB.TextBox editOrganization 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1020
         Width           =   4395
      End
      Begin VB.TextBox editSenderAddress 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   660
         Width           =   4395
      End
      Begin VB.TextBox editSenderName 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   300
         Width           =   4395
      End
      Begin VB.Label lblOrganization 
         Caption         =   "Organizacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblSenderAddress 
         Caption         =   "Correo:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblName 
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame frameServer 
      Caption         =   "Servidor"
      Height          =   1755
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   6075
      Begin VB.TextBox editRelayPort 
         Height          =   315
         Left            =   1020
         TabIndex        =   13
         Top             =   1260
         Width           =   795
      End
      Begin VB.TextBox editRelayServer 
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         Top             =   900
         Width           =   2415
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Send all messages through the specified relay server"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Send all messages directly to the recipient"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   3375
      End
      Begin VB.Label lblRelayPort 
         Caption         =   "Port:"
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblRelayServer 
         Caption         =   "Name:"
         Height          =   255
         Left            =   450
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Const g_strRelayServer = "127.0.0.1"
'Const g_nRelayPort = 25

Private Sub Form_Load()
    LoadOptions
    UpdateForm False
End Sub

Private Sub editSenderName_Change()
    UpdateForm True
End Sub

Private Sub editSenderName_GotFocus()
    editSenderName.SelStart = 0
    editSenderName.SelLength = Len(editSenderName.Text)
End Sub

Private Sub editSenderAddress_Change()
    UpdateForm True
End Sub

Private Sub editSenderAddress_GotFocus()
    editSenderAddress.SelStart = 0
    editSenderAddress.SelLength = Len(editSenderAddress.Text)
End Sub

Private Sub editOrganization_Change()
    UpdateForm True
End Sub

Private Sub editOrganization_GotFocus()
    editOrganization.SelStart = 0
    editOrganization.SelLength = Len(editOrganization.Text)
End Sub

Private Sub optServer_Click(Index As Integer)
    UpdateForm True
End Sub

Private Sub editRelayServer_Change()
    UpdateForm True
End Sub

Private Sub editRelayServer_GotFocus()
    editRelayServer.SelStart = 0
    editRelayServer.SelLength = Len(editRelayServer.Text)
End Sub

Private Sub editRelayPort_Change()
    UpdateForm True
End Sub

Private Sub editRelayPort_GotFocus()
    editRelayPort.SelStart = 0
    editRelayPort.SelLength = Len(editRelayPort.Text)
End Sub

Private Sub cmdOK_Click()
    UpdateForm True
    
    If g_bRelayMessages And Len(g_strRelayServer) = 0 Then
        MsgBox "A relay server host name or IP address must be specified", vbExclamation
        editRelayServer.SetFocus
        Exit Sub
    End If
    
    SaveOptions
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdApply_Click()
    UpdateForm True
    
    If g_bRelayMessages And Len(g_strRelayServer) = 0 Then
        MsgBox "A relay server host name or IP address must be specified", vbExclamation
        editRelayServer.SetFocus
        Exit Sub
    End If
    
    SaveOptions
    cmdApply.Enabled = False
End Sub

Private Sub UpdateForm(bValidate As Boolean)
        Static bUpdating As Boolean
        
        If bUpdating Then Exit Sub
        bUpdating = True
        
    If bValidate Then
        Dim bModified As Boolean
        
        If wnomuser <> g_strSenderName Then bModified = True
        If wcorreouser <> g_strSenderAddress Then bModified = True
        If editOrganization.Text <> g_strOrganization Then bModified = True
        
        g_strSenderName = Trim(wnomuser)
        g_strSenderAddress = Trim(wcorreouser)
        g_strOrganization = Trim(editOrganization.Text)
        
        If optServer(0).Value = True Then
            If g_bRelayMessages = True Then bModified = True
            g_bRelayMessages = False
        Else
            If g_bRelayMessages = False Then bModified = True
            g_bRelayMessages = True
        End If
        
        If g_bRelayMessages Then
            If editRelayServer.Text <> g_strRelayServer Then bModified = True
            If Val(editRelayPort.Text) <> g_nRelayPort Then bModified = True

'            g_strRelayServer = Trim(editRelayServer.Text)
'            g_nRelayPort = val(editRelayPort)
        End If
        
        cmdApply.Enabled = bModified
    Else
        editSenderName.Text = g_strSenderName
        editSenderAddress.Text = g_strSenderAddress
        editOrganization.Text = g_strOrganization

        If g_bRelayMessages Then
            optServer(0).Value = False
            optServer(1).Value = True
        Else
            optServer(0).Value = True
            optServer(1).Value = False
        End If

        editRelayServer.Text = g_strRelayServer
        editRelayPort.Text = CStr(g_nRelayPort)
    End If
    
    lblRelayServer.Enabled = g_bRelayMessages
    lblRelayPort.Enabled = g_bRelayMessages
    editRelayServer.Enabled = g_bRelayMessages
    editRelayPort.Enabled = g_bRelayMessages
    
    bUpdating = False
End Sub
