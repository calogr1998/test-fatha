VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form frmorigen 
   Caption         =   "Orígenes"
   ClientHeight    =   5055
   ClientLeft      =   3960
   ClientTop       =   3720
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6300
   Begin Threed.SSFrame SSFrame1 
      Height          =   4245
      Left            =   135
      TabIndex        =   9
      Top             =   135
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   7488
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame SSFrame3 
         Height          =   915
         Left            =   225
         TabIndex        =   13
         Top             =   2475
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   1614
         _StockProps     =   14
         Caption         =   "Visualizar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optcosto 
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   405
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "&Ninguno"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optcosto 
            Height          =   240
            Index           =   1
            Left            =   1980
            TabIndex        =   5
            Top             =   405
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "&Solo Costo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optcosto 
            Height          =   240
            Index           =   2
            Left            =   3780
            TabIndex        =   6
            Top             =   405
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "&Costo,Igv y Total"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   915
         Left            =   225
         TabIndex        =   12
         Top             =   1485
         Width           =   5640
         _Version        =   65536
         _ExtentX        =   9948
         _ExtentY        =   1614
         _StockProps     =   14
         Caption         =   "Tipo Movimiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption opttipo 
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   405
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "&Ingreso"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption opttipo 
            Height          =   285
            Index           =   1
            Left            =   1980
            TabIndex        =   3
            Top             =   405
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "&Salida"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSCheck chkpartida 
         Height          =   195
         Left            =   900
         TabIndex        =   7
         Top             =   3735
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Mostrar Proveedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtdescripcion 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   1
         Top             =   855
         Width           =   4560
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   0
         Top             =   405
         Width           =   960
      End
      Begin Threed.SSCheck chkprecio 
         Height          =   195
         Left            =   3240
         TabIndex        =   8
         Top             =   3735
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Mostrar Precio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   405
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "frmorigen.frx":0000
      ToolBars        =   "frmorigen.frx":4BDC
   End
End
Attribute VB_Name = "frmorigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Private Sub chkpartida_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    chkprecio.SetFocus
End If
End Sub

Private Sub Form_Activate()
If Not sw_nuevo_doc Then
    SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
    CargaOrigen
Else
    SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
End If
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
        SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
        sw_nuevo_mant = True
        Call nuevo
    Case "ID_Eliminar"
        Call eliminar
    Case "ID_Grabar"
        If Len(Trim(txtcodigo.Text)) = 0 Then
            MsgBox "Debe Ingresar Código del Orígen", vbInformation, "Sistema de Logística"
            txtcodigo.Text = ""
            txtcodigo.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(txtdescripcion.Text)) = 0 Then
            MsgBox "Debe Ingresar Descripción del Orígen", vbInformation, "Sistema de Logística"
            txtdescripcion.Text = ""
            txtdescripcion.SetFocus
            Exit Sub
        End If
            
        Screen.MousePointer = vbHourglass
        Call grabar
        Screen.MousePointer = vbDefault
    Case "ID_Imprimir"
        Call imprimir
    Case "ID_Lista"
'        ListaOrigen.adoctasctes.Refresh
        Unload Me
End Select
End Sub

Private Sub txtcodigo_GotFocus()
txtcodigo.SelStart = 0
txtcodigo.SelLength = Len(txtcodigo.Text)
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtdescripcion.SetFocus
Else
    KeyAscii = VALIDA(1, KeyAscii, , True)
End If
End Sub

Private Sub txtdescripcion_GotFocus()
txtdescripcion.SelStart = 0
txtdescripcion.SelLength = Len(txtdescripcion.Text)
End Sub

Public Sub grabar()

wf1codori = Trim(txtcodigo.Text)
wf1nomori = Trim(txtdescripcion.Text)
wf1tipo = IIf(opttipo(0).Value, "I", "S")
wf1partida = IIf(chkpartida.Value, "1", "0")
wf1precio = IIf(chkprecio.Value, "1", "0")
If optcosto(0).Value Then
    wf1costo = ""
ElseIf optcosto(1).Value Then
    wf1costo = "1"
Else
    wf1costo = "*"
End If

If Not sw_nuevo_mant Then
    SQL = "delete from sf1origenes where f1codori='" & wf1codori & "'"
    cnn_dbbancos.Execute SQL
Else
    If rst.State = adStateOpen Then rst.Close
    SQL = "select f1codori from sf1origenes where f1codori='" & wf1codori & "'"
    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        MsgBox "El Código de Orígen " & wf1codori & " ya Existe", vbExclamation, "Sistema de Logística"
        txtcodigo.SetFocus
        Exit Sub
    End If
    rst.Close
End If

cab = "insert into sf1origenes (f1codori,f1nomori,f1tipmov,f1partida,f1costo,f1precio)"
det = "values ('" & wf1codori & "','" & wf1nomori & "','" & wf1tipo & "','" & wf1partida & "','" & wf1costo & _
"','" & wf1precio & "')"
cnn_dbbancos.Execute cab & " " & det

sw_nuevo_mant = False
SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True

MsgBox "El Orígen " & wf1codori & "-" & wf1nomori & Chr(13) & "  ha sido Actualizado", vbInformation, "Sistema de Logística"
End Sub

Public Sub CargaOrigen()
Set rst = New ADODB.Recordset

txtcodigo.Text = wconcepto

If rst.State = adStateOpen Then rst.Close
SQL = "select * from sf1origenes where f1codori='" & wconcepto & "'"
rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    txtcodigo.Enabled = False
    txtdescripcion.Text = "" & rst("f1nomori")
    If rst("f1tipmov") = "I" Then
        opttipo(0).Value = True
    Else
        opttipo(1).Value = True
    End If
    
    If rst("f1costo") = "*" Then
        optcosto(2).Value = True
    ElseIf rst("f1costo") = "1" Then
        optcosto(1).Value = True
    Else
        optcosto(0).Value = True
    End If
    
    If rst("f1partida") = "1" Then
        chkpartida.Value = True
    Else
        chkpartida.Value = False
    End If
    
    If rst("f1precio") = "1" Then
        chkprecio.Value = True
    Else
        chkprecio.Value = False
    End If
        
    txtdescripcion.SetFocus
End If
rst.Close
End Sub

Public Sub nuevo()
txtcodigo.Text = ""
txtdescripcion.Text = ""
opttipo(0).Value = True
optcosto(0).Value = True
chkpartida = False
chkprecio.Value = False
txtcodigo.Enabled = True
txtcodigo.SetFocus
End Sub

Public Sub eliminar()
resp = MsgBox("¿Esta Seguro de Eliminar el Orígen " & txtcodigo.Text & "-" & txtdescripcion.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Sistema de Logística")
If resp = vbYes Then
    'Verifica si Orígen a Eliminar ha sido utilizado
    If rst.State = adStateOpen Then rst.Close
    SQL = "select f1codori from if4vales where f1codori='" & txtcodigo.Text & "'"
    rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If rst.EOF Then
        weliminar = True
    Else
        weliminar = False
    End If
    rst.Close
    
    If weliminar Then
        SQL = "delete from sf1origenes where f1codori='" & txtcodigo.Text & "'"
        cnn_dbbancos.Execute SQL
        MsgBox "El Orígen " & txtcodigo.Text & "-" & txtdescripcion.Text & " ha sido Eliminado", vbInformation, "Sistema de Logística"
        Call nuevo
    Else
        MsgBox "El Orígen " & txtcodigo.Text & "-" & txtdescripcion.Text & " ha sido utilizado" & Chr(13) & "No Podrá Eliminarlo", vbInformation, "Sistema de Logística"
    End If
End If
End Sub

Public Sub imprimir()
With rptorigen
    .datos.ConnectionString = cnn_dbbancos
    SQL = "select f1codori,f1nomori,iif(f1tipmov='I','INGRESOS','SALIDAS') as tipo from sf1origenes order by 3,f1nomori"
    .datos.Source = SQL
    .fldfecha.Text = Format(Date, "DD/MM/YYYY")
    .lblcia.Caption = wnomcia
    .Caption = "Orígenes"
    .Show vbModal
End With
End Sub
