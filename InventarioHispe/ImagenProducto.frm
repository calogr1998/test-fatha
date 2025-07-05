VERSION 5.00
Begin VB.Form ImagenProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   1800
   ClientTop       =   1920
   ClientWidth     =   6795
   Icon            =   "ImagenProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Image ImgProducto 
      Height          =   5595
      Left            =   0
      Top             =   0
      Width           =   5115
   End
End
Attribute VB_Name = "ImagenProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
If FileExist(App.Path & "\Articulos\" & wNameImage & ".jpg") = True Then
    ImgProducto.Picture = LoadPicture(App.Path & "\Articulos\" & wNameImage & ".jpg")
Else
    ImgProducto.Picture = LoadPicture(App.Path & "\Articulos\Default.jpg")
End If

Me.Width = ImgProducto.Width + 200
Me.Height = ImgProducto.Height + 500

Me.left = (Screen.Width - ImgProducto.Width) / 2
Me.top = (Screen.Height - ImgProducto.Height) / 2

End Sub
