VERSION 5.00
Begin VB.Form lee_codigosbarra 
   Caption         =   "Códigos de Barra"
   ClientHeight    =   1815
   ClientLeft      =   2895
   ClientTop       =   2550
   ClientWidth     =   5370
   LinkTopic       =   "Form2"
   ScaleHeight     =   1815
   ScaleWidth      =   5370
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5055
      Begin VB.TextBox txtcodigo 
         Height          =   330
         Left            =   1710
         TabIndex        =   2
         Top             =   495
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ESC - Salir"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4095
         TabIndex        =   3
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   990
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
   End
End
Attribute VB_Name = "lee_codigosbarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtcodigo_Change()
    
    If Len(txtcodigo) = 6 Then
        wcodigo_barra = txtcodigo.Text
        wsw_codbarra = 2
        Me.Hide
    End If

End Sub

Private Sub txtcodigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        wcodigo_barra = txtcodigo.Text
        wsw_codbarra = 2
        Unload Me
    End If

End Sub
