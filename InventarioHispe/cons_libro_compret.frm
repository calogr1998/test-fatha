VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form cons_libro_compret 
   Caption         =   "Reporte Mensual de los Comprobantes de Retención"
   ClientHeight    =   2130
   ClientLeft      =   4080
   ClientTop       =   3060
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5700
   Begin Threed.SSCommand cmdaceptar 
      Height          =   465
      Left            =   1395
      TabIndex        =   2
      Top             =   1575
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1365
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   5550
      _Version        =   65536
      _ExtentX        =   9790
      _ExtentY        =   2408
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.ComboBox cmbmes 
         Height          =   315
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1860
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   1
         Top             =   765
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   1620
         TabIndex        =   6
         Top             =   315
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   1620
         TabIndex        =   5
         Top             =   855
         Width           =   285
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   465
      Left            =   2880
      TabIndex        =   3
      Top             =   1575
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "cons_libro_compret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bancos  As ADODB.Connection
Dim sql     As String
Dim anno    As String

Private Sub cmbmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Text1.SetFocus
    End If

End Sub

Private Sub cmdaceptar_Click()
Dim rsnuevo1    As ADODB.Recordset
Dim tbparam     As ADODB.Recordset
Dim nuevo       As ADODB.Recordset

    Set tbparam = New ADODB.Recordset
    Set nuevo = New ADODB.Recordset
    Set rsnuevo1 = New ADODB.Recordset
    
    acr_compret.Label13.Caption = wnomcia
    acr_compret.Label14.Caption = Format(Now, "dd/mm/yyyy")
    
    
    sql = "Select * from RETENDOC where Right(FECHA,4) & Mid(FECHA,4,2)='" & Text1.Text & "' & '" & Right(cmbmes.Text, 2) & "' order by SERIE,NUM_DOCUMENTO"
    If rsnuevo1.State = adStateOpen Then rsnuevo1.Close
    rsnuevo1.Open sql, bancos, adOpenDynamic, adLockOptimistic
    
    
    If Not rsnuevo1.EOF Then
        acr_compret.DataControl1.ConnectionString = bancos
        acr_compret.DataControl1.Source = sql
        acr_compret.Label17.Caption = Left(cmbmes.Text, 20)
        acr_compret.Label18.Caption = "Intersys - Sistema de Compras"
        acr_compret.Show 1
    Else
        MsgBox "No existen registros para ser procesados", vbInformation, "Atencion"
        cmbmes.SetFocus
    End If
    rsnuevo1.Close
    
End Sub

Private Sub cmdaceptar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdsalir.SetFocus
    End If

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Left = 4500
    Me.Top = 2500
    
    Set bancos = New ADODB.Connection
    
    cmbmes.AddItem "Enero    " & Space(80) & "01"
    cmbmes.AddItem "Febrero  " & Space(80) & "02"
    cmbmes.AddItem "Marzo    " & Space(80) & "03"
    cmbmes.AddItem "Abril    " & Space(80) & "04"
    cmbmes.AddItem "Mayo     " & Space(80) & "05"
    cmbmes.AddItem "Junio    " & Space(80) & "06"
    cmbmes.AddItem "Julio    " & Space(80) & "07"
    cmbmes.AddItem "Agosto   " & Space(80) & "08"
    cmbmes.AddItem "Setiembre" & Space(80) & "09"
    cmbmes.AddItem "Octubre  " & Space(80) & "10"
    cmbmes.AddItem "Noviembre" & Space(80) & "11"
    cmbmes.AddItem "Diciembre" & Space(80) & "12"
    cmbmes.ListIndex = 0
    
    Text1.Text = Year(Date)
    
    With bancos
     .Provider = "Microsoft.JET.OLEDB.4.0; Data Source=" & wrutabancos & "\db_bancos.mdb; Persist Security Info=False"
     .Open
    End With
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdaceptar.SetFocus
End If
KeyAscii = TxtNum(KeyAscii)

End Sub

Function TxtNum(KeyAscii As Integer) As Integer
  If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 32 Then
      TxtNum = KeyAscii
  Else
      TxtNum = 0
      Beep
  End If
End Function


