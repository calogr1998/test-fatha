VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdetalle 
   ClientHeight    =   2580
   ClientLeft      =   3330
   ClientTop       =   4635
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7935
   Begin MSFlexGridLib.MSFlexGrid grddetalle 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin Threed.SSPanel pnldetalle 
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      _Version        =   65536
      _ExtentX        =   13917
      _ExtentY        =   3916
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars2 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   2
         Tools           =   "frmdetalle.frx":0000
         ToolBars        =   "frmdetalle.frx":197C
      End
   End
End
Attribute VB_Name = "frmdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    wf5codpro = ""
    Unload Me
End If
End Sub

Private Sub Form_Load()
wf5codpro = ""
End Sub

Private Sub grddetalle_DblClick()
wf5codpro = grddetalle.TextMatrix(grddetalle.Row, 2)
Unload Me
End Sub

Private Sub grddetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    grddetalle_DblClick
End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub SSActiveToolBars2_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "idaceptar"
        grddetalle_DblClick
    Case "idcancelar"
        wf5codpro = ""
        Unload Me
End Select
End Sub
