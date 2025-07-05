VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form mant_vendedores 
   Caption         =   "Actualizacion de Vendedores"
   ClientHeight    =   2970
   ClientLeft      =   2025
   ClientTop       =   1485
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   7785
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7620
      Begin VB.TextBox txtdscto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4890
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "0"
         Top             =   1560
         Width           =   645
      End
      Begin Threed.SSCheck chkvendedor 
         Height          =   330
         Left            =   360
         TabIndex        =   3
         Top             =   1575
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Vendedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtnombres 
         Height          =   315
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   2
         Top             =   990
         Width           =   5820
      End
      Begin VB.TextBox txtcodigo 
         Height          =   315
         Left            =   1530
         MaxLength       =   8
         TabIndex        =   1
         Top             =   495
         Width           =   1140
      End
      Begin Threed.SSCheck chkcobrador 
         Height          =   330
         Left            =   2385
         TabIndex        =   4
         Top             =   1575
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Cobrador"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "% Dscto."
         Height          =   210
         Left            =   4080
         TabIndex        =   8
         Top             =   1605
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   210
         Left            =   360
         TabIndex        =   7
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   540
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   270
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_vendedores.frx":0000
      ToolBars        =   "mant_vendedores.frx":7E74
   End
End
Attribute VB_Name = "mant_vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SQL         As String

Private Sub chkvendedor_Click(Value As Integer)

    If chkvendedor.Value = True Then
        Label3.Visible = True
        txtdscto.Visible = True
    Else
        Label3.Visible = False
        txtdscto.Visible = False
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    If sw_nuevo_doc = True Then
        nuevo
    Else
        actualizacion lista_vendedores.dxDBGrid1.Columns(0).Value
    End If

End Sub

Private Sub nuevo()
    
    txtcodigo.Text = ""
    txtnombres.Text = ""
    txtdscto.Text = 0#
    chkvendedor.Value = False
    chkcobrador.Value = False
    
End Sub

Private Sub actualizacion(cod)

    SQL = "select * from EF2USERS where F2CODUSER='" & cod & "'"
    If rsusuarios.State = adStateOpen Then rsusuarios.Close
    rsusuarios.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsusuarios.EOF Then
        txtcodigo.Text = Trim("" & rsusuarios.Fields("F2CODUSER"))
        txtnombres.Text = Trim("" & rsusuarios.Fields("F2NOMUSER"))
        If Trim("" & rsusuarios.Fields("F2VENDEDOR")) = "*" Then
            chkvendedor.Value = True
        End If
        If Trim("" & rsusuarios.Fields("F2COBRADOR")) = "*" Then
            chkcobrador.Value = True
        End If
    End If
    rsusuarios.Close

End Sub

Private Sub eliminar()

    Beep
    If MsgBox("Está seguro de eliminar el vendedor.", 36, "Atención") = 6 Then
        SQL = "select F2CODUSER from EF2USERS where F2CODUSER='" & txtcodigo.Text & "'"
        If rsusuarios.State = adStateOpen Then rsusuarios.Close
        rsusuarios.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsusuarios.EOF Then
            SQL = "DELETE * from EF2USERS where F2CODUSER='" & txtcodigo.Text & "'"
            cnn_dbbancos.Execute SQL
            nuevo
        Else
            Beep
        End If
        rsusuarios.Close
        txtcodigo.SetFocus
    End If

End Sub

Private Sub grabar()
On Error GoTo graba
Dim amovs(0 To 4) As a_grabacion

    SQL = "select F2CODUSER from EF2USERS where F2CODUSER='" & txtcodigo.Text & "'"
    If rsusuarios.State = adStateOpen Then rsusuarios.Close
    rsusuarios.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsusuarios.EOF Then
        sw = 0
    Else
        sw = 1
    End If
    rsusuarios.Close
    amovs(0).campo = "F2CODUSER": amovs(0).valor = txtcodigo.Text: amovs(0).TIPO = "T"
    amovs(1).campo = "F2NOMUSER": amovs(1).valor = txtnombres.Text: amovs(1).TIPO = "T"
    amovs(2).campo = "F2VENDEDOR": amovs(2).valor = IIf(chkvendedor.Value = True, "*", " "): amovs(2).TIPO = "T"
    amovs(3).campo = "F2COBRADOR": amovs(3).valor = IIf(chkcobrador.Value = True, "*", " "): amovs(3).TIPO = "T"
    amovs(4).campo = "F2PORCDSCTO": amovs(4).valor = txtdscto.Text: amovs(4).TIPO = "N"
    
    If sw = 1 Then
        GRABA_REGISTRO amovs(), "EF2USERS", "A", 4, cnn_dbbancos, ""
    Else
        GRABA_REGISTRO amovs(), "EF2USERS", "M", 4, cnn_dbbancos, "F2CODUSER='" & txtcodigo.Text & "'"
    End If
    Exit Sub
    
graba:
    If Err = 3186 Then
        For i% = 1 To 10000
        Next i%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'    lista_vendedores.adoctasctes.Refresh
    
End Sub

Private Sub txtcodigo_GotFocus()

    txtcodigo.SelStart = 0: txtcodigo.SelLength = Len(txtcodigo.Text)
    
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtnombres.SetFocus
    End If
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            nuevo
            txtcodigo.SetFocus
        Case "ID_Grabar"
            grabar
        Case "ID_Eliminar"
            eliminar
        Case "ID_Lista"
            Unload Me
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub txtcodigo_LostFocus()

    If Len(Trim(txtcodigo.Text)) > 0 Then
        SQL = "select F2NOMUSER,F2VENDEDOR,F2COBRADOR from EF2USERS where F2CODUSER='" & txtcodigo.Text & "'"
        If rsusuarios.State = adStateOpen Then rsusuarios.Close
        rsusuarios.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsusuarios.EOF Then
            txtnombres.Text = "" & rsusuarios.Fields("F2NOMUSER")
            txtdscto.Text = Val("" & rsusuarios.Fields("F2PORCDSCTO"))
            If Trim("" & rsusuarios.Fields("F2VENDEDOR")) = "*" Then
                chkvendedor.Value = True
            End If
            If Trim("" & rsusuarios.Fields("F2COBRADOR")) = "*" Then
                chkcobrador.Value = True
            End If
        End If
        rsusuarios.Close
    End If

End Sub
