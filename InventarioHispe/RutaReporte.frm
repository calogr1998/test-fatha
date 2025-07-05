VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RutaReporte 
   Caption         =   "Form3"
   ClientHeight    =   930
   ClientLeft      =   9450
   ClientTop       =   3870
   ClientWidth     =   1950
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   930
   ScaleWidth      =   1950
   Begin MSComDlg.CommonDialog caja 
      Left            =   720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "RutaReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRuta As String
Dim nTipoFile As Integer
Private Sub Form_Load()
    Me.left = 0
    Me.top = 0
    Ruta = ""
    Select Case TipoFile
    Case 0
        caja.DialogTitle = "Exportar a *.pdf"
        caja.Filter = "Archivo Acrobat (*.pdf)|*.pdf"
    Case 1
        caja.DialogTitle = "Exportar a *.xls"
        caja.Filter = "Archivo Excel (*.xls)|*.xls"
    Case 2
        caja.DialogTitle = "Exportar a *.gif"
        caja.Filter = "Archivo GIF (*.gif)|*.gif"
    Case 3
        caja.DialogTitle = "Exportar a *.rtf"
        caja.Filter = "Archivo RTF (*.rtf)|*.rtf"
    End Select
    caja.CancelError = False
    caja.ShowSave
    Ruta = caja.FileName
End Sub
Public Property Get Ruta() As String
    Ruta = strRuta
End Property

Public Property Let Ruta(ByVal vNewValue As String)
    strRuta = vNewValue
End Property

Public Property Get TipoFile() As Integer
    TipoFile = nTipoFile
End Property

Public Property Let TipoFile(ByVal vNewValue As Integer)
    nTipoFile = vNewValue
End Property
