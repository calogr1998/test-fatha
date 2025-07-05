VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCommon 
   Caption         =   "Common Dialog"
   ClientHeight    =   975
   ClientLeft      =   1740
   ClientTop       =   2280
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   2415
   Begin MSComDlg.CommonDialog caja 
      Left            =   180
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRuta As String

Private Sub Form_Load()
Select Case TipoExporta
Case 1
    Ruta = ""
    caja.DialogTitle = "Guardar como"
    caja.Filter = "Excel (*.xls)|*.xls"
    caja.CancelError = False
    caja.ShowSave
    Ruta = caja.FileName
Case 2
    Ruta = ""
    caja.DialogTitle = "Guardar como"
    caja.Filter = "Acrobat (*.pdf)|*.pdf"
    caja.CancelError = False
    caja.ShowSave
    Ruta = caja.FileName
End Select
End Sub

Public Property Get Ruta() As String
    Ruta = strRuta
End Property

Public Property Let Ruta(ByVal vNewValue As String)
    strRuta = vNewValue
End Property
