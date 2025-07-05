VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form ReportedeValorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reporte de Productos no valorizados :::"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   Icon            =   "ReportedeValorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   2565
      TabIndex        =   5
      Top             =   1275
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Procesar..."
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   3690
      Begin VB.OptionButton Option4 
         Caption         =   "Producción"
         Height          =   480
         Left            =   1890
         TabIndex        =   4
         Top             =   585
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Importaciones"
         Height          =   480
         Left            =   1890
         TabIndex        =   3
         Top             =   135
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Compras Locales"
         Height          =   480
         Left            =   180
         TabIndex        =   2
         Top             =   585
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos..."
         Enabled         =   0   'False
         Height          =   480
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   1695
      End
   End
End
Attribute VB_Name = "ReportedeValorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SSCommand1_Click()
Dim sql As String, codigo As String * 3

If Option2.Value = True Then codigo = "XCO": nombre = Option2.Caption
If Option3.Value = True Then codigo = "XIM": nombre = Option3.Caption
If Option4.Value = True Then codigo = "XCP": nombre = Option4.Caption


sql = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL " & _
" FROM SF1ORIGENES INNER JOIN IF4VALES ON SF1ORIGENES.F1CODORI = IF4VALES.F1CODORI " & _
" WHERE (((SF1ORIGENES.F1CODORI) = '" & codigo & "')) AND (Select Count(F3VALVTA) From " & _
" IF3VALES Where IF3VALES.F2CODALM =  IF4VALES.F2CODALM AND IF3VALES.F4NUMVAL = IF4VALES.F4NUMVAL) " & _
" > 0 ORDER BY F4FECVAL;"
Acr_Prod_no_val.Label10.Caption = "Los Vales listados son de " & nombre
Acr_Prod_no_val.lblempresa.Caption = wempresa
Acr_Prod_no_val.fldfecha.Text = Date
Acr_Prod_no_val.DataControl1.ConnectionString = cnn_dbbancos
Acr_Prod_no_val.DataControl1.Source = sql
Acr_Prod_no_val.Show 1

End Sub
