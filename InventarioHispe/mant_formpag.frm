VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form mant_formpag 
   Caption         =   "Mantenimiento de Formas de Pago"
   ClientHeight    =   5430
   ClientLeft      =   3660
   ClientTop       =   2880
   ClientWidth     =   6450
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
   ScaleHeight     =   5430
   ScaleWidth      =   6450
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   6315
      Begin VB.Frame Frame4 
         Caption         =   " Letras "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   315
         TabIndex        =   16
         Top             =   3645
         Width           =   5685
         Begin VB.TextBox txtinteres 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4455
            TabIndex        =   10
            Top             =   405
            Width           =   915
         End
         Begin VB.TextBox txtnumletras 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1485
            TabIndex        =   9
            Top             =   405
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            Height          =   210
            Left            =   3735
            TabIndex        =   18
            Top             =   495
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N° de Letras"
            Height          =   210
            Left            =   270
            TabIndex        =   17
            Top             =   495
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Sistema "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   315
         TabIndex        =   15
         Top             =   2610
         Width           =   5685
         Begin Threed.SSOption optsistema 
            Height          =   240
            Index           =   0
            Left            =   405
            TabIndex        =   6
            Top             =   405
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Compras"
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
         Begin Threed.SSOption optsistema 
            Height          =   240
            Index           =   1
            Left            =   2250
            TabIndex        =   7
            Top             =   405
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Ventas"
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
         Begin Threed.SSOption optsistema 
            Height          =   240
            Index           =   2
            Left            =   4185
            TabIndex        =   8
            Top             =   405
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Ambos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   315
         TabIndex        =   14
         Top             =   1620
         Width           =   5685
         Begin Threed.SSOption opttipo 
            Height          =   330
            Index           =   0
            Left            =   1170
            TabIndex        =   4
            Top             =   360
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   582
            _StockProps     =   78
            Caption         =   "Financiado"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Height          =   330
            Index           =   1
            Left            =   3555
            TabIndex        =   5
            Top             =   360
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   582
            _StockProps     =   78
            Caption         =   "Contado"
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
      End
      Begin VB.TextBox txtcodigo 
         Height          =   330
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   0
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtdescrip 
         Height          =   330
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   1
         Top             =   810
         Width           =   4380
      End
      Begin VB.TextBox txtdias 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1575
         TabIndex        =   2
         Top             =   1215
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   210
         Left            =   360
         TabIndex        =   13
         Top             =   495
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   210
         Left            =   360
         TabIndex        =   12
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° de Dias"
         Height          =   210
         Left            =   360
         TabIndex        =   11
         Top             =   1305
         Width           =   750
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   135
      Top             =   4950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "mant_formpag.frx":0000
      ToolBars        =   "mant_formpag.frx":4BDC
   End
End
Attribute VB_Name = "mant_formpag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL             As String

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_mant = True
            nuevo_formpag
        Case "ID_Grabar"
            grabar_formpag
        Case "ID_Eliminar"
            eliminar_formpag
        Case "ID_Imprimir":
            With Acr_formpag
                .DataControl1.ConnectionString = cnn_dbbancos
                .DataControl1.Source = "Select * FROM EF2FORPAG ORDER BY F2FORPAG"
                .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                .lblempresa.Caption = wnomcia
                .Show 1
            End With
        Case "ID_Lista"
'            lista_formpag.adoctasctes.Refresh
            Unload Me
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()

    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    If sw_nuevo_mant = True Then
        nuevo_formpag
    Else
        actualizacion_formpag lista_formpag.dxDBGrid1.Columns(0).Value
        txtcodigo.Enabled = False
    End If

End Sub

Private Sub actualizacion_formpag(cod)

    SQL = "select * from ef2forpag where f2forpag='" & cod & "'"
    If rsformpag.State = adStateOpen Then rsformpag.Close
    rsformpag.Open SQL, cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not rsformpag.EOF Then
        txtcodigo.Text = "" & rsformpag.Fields("f2forpag")
        txtdescrip.Text = "" & rsformpag.Fields("f2despag")
        txtdias.Text = "" & rsformpag.Fields("f2dias")
        If "" & rsformpag.Fields("F2TIPO") = "F" Then
            opttipo(0).Value = True
        Else
            opttipo(1).Value = True
        End If
        If "" & rsformpag.Fields("F2TIPO_FP") = "C" Then
            optsistema(0).Value = True
        End If
        If "" & rsformpag.Fields("F2TIPO_FP") = "V" Then
            optsistema(1).Value = True
        End If
        If "" & rsformpag.Fields("F2TIPO_FP") = "A" Then
            optsistema(2).Value = True
        End If
        txtnumletras.Text = Val("" & rsformpag.Fields("F2CANTLETRAS"))
        txtinteres.Text = Val("" & rsformpag.Fields("F2INTERES"))
    End If

End Sub

Private Sub nuevo_formpag()
    
    txtcodigo.Text = ""
    txtdescrip.Text = ""
    txtdias.Text = Format(0, "00")
    opttipo(0).Value = True
    optsistema(2).Value = True
    txtnumletras.Text = ""
    txtinteres.Text = ""
    
    txtcodigo.Enabled = True

End Sub

Private Sub grabar_formpag()
On Error GoTo ERROR_GRABA
Dim amovs(0 To 6)   As a_grabacion
Dim ctiposis        As String

    SQL = "select * from ef2forpag where f2forpag='" & txtcodigo.Text & "'"
    If rsformpag.State = adStateOpen Then rsformpag.Close
    rsformpag.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsformpag.EOF Then
        sw = 0
    Else
        sw = 1
    End If
    rsformpag.Close
    amovs(0).campo = "f2forpag": amovs(0).valor = txtcodigo.Text: amovs(0).TIPO = "T"
    amovs(1).campo = "f2despag": amovs(1).valor = txtdescrip.Text: amovs(1).TIPO = "T"
    amovs(2).campo = "f2dias": amovs(2).valor = Val(Format(txtdias.Text, "00")): amovs(2).TIPO = "N"
    amovs(3).campo = "F2TIPO": amovs(3).valor = IIf(opttipo(0).Value = True, "F", "C"): amovs(3).TIPO = "T"
    ctiposis = ""
    If optsistema(0).Value = True Then
        ctiposis = "C"
    End If
    If optsistema(1).Value = True Then
        ctiposis = "V"
    End If
    If optsistema(2).Value = True Then
        ctiposis = "A"
    End If
    amovs(4).campo = "F2TIPO_FP": amovs(4).valor = ctiposis: amovs(4).TIPO = "T"
    amovs(5).campo = "F2CANTLETRAS": amovs(5).valor = Val(txtnumletras.Text & ""): amovs(5).TIPO = "N"
    amovs(6).campo = "F2INTERES": amovs(6).valor = Val(txtinteres.Text & ""): amovs(6).TIPO = "N"
    If sw = 1 Then
        GRABA_REGISTRO amovs(), "ef2forpag", "A", 6, cnn_dbbancos, ""
    Else
        GRABA_REGISTRO amovs(), "ef2forpag", "M", 6, cnn_dbbancos, "f2forpag='" & txtcodigo.Text & "'"
    End If
    txtcodigo.Enabled = False
    Exit Sub

ERROR_GRABA:
    If Err.Number = 3186 Then
        If MsgBox("La base de datos se encuentra bloqueada por otro usuario. Desea reintentar ?", vbYesNo, "Atención") = vbYes Then
            Resume
        Else
            MsgBox "No se realizaron los cambios.", vbInformation, "Atención"
            Exit Sub
        End If
    Else
        MsgBox "Se ha producido el sgte. error : " & Err.Description, vbInformation, "Atención"
        Exit Sub
    End If

End Sub

Private Sub eliminar_formpag()
On Error GoTo ERROR_ELIMINA
 
    Beep
    If MsgBox("¿Está seguro de eliminar la Forma de Pago?", 36, "Atención") = 6 Then
        SQL = "select f2forpag from ef2forpag where f2forpag='" & txtcodigo.Text & "' "
        If rsformpag.State = adStateOpen Then rsformpag.Close
        rsformpag.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsformpag.EOF Then
            SQL = "DELETE * from ef2forpag where f2forpag='" & txtcodigo.Text & "' "
            cnn_dbbancos.Execute SQL
            txtcodigo.Enabled = True
            nuevo_formpag
        Else
            Beep
        End If
        rsformpag.Close
        txtdescrip.SetFocus
    End If
    Exit Sub

ERROR_ELIMINA:
    If Err.Number = 3186 Then
        If MsgBox("La base de datos se encuentra bloqueada por otro usuario. Desea reintentar ?", vbYesNo, "Atención") = vbYes Then
            Resume
        Else
            MsgBox "No se realizaron los cambios.", vbInformation, "Atención"
            Exit Sub
        End If
    Else
        MsgBox "Se ha producido el sgte. error : " & Err.Description, vbInformation, "Atención"
        Exit Sub
    End If

End Sub

Private Sub txtcodigo_GotFocus()

    txtcodigo.SelStart = 0: txtcodigo.SelLength = Len(txtcodigo.Text)
    
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdescrip.SetFocus
    End If

End Sub

Private Sub txtcodigo_LostFocus()

    If Len(Trim(txtcodigo.Text)) > 0 Then
        SQL = "select f2forpag from ef2forpag where f2forpag='" & txtcodigo.Text & "' "
        If rsformpag.State = adStateOpen Then rsformpag.Close
        rsformpag.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsformpag.EOF Then
            MsgBox "Codigo ya existe.", vbInformation, "Mensaje"
            txtcodigo.Text = ""
            txtcodigo.SetFocus
       End If
    End If
    
End Sub

Private Sub txtdescrip_GotFocus()

    txtdescrip.SelStart = 0: txtdescrip.SelLength = Len(txtdescrip.Text)
    
End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdias.SetFocus
    End If
    
End Sub

Private Sub txtdias_GotFocus()

    txtdias.SelStart = 0: txtdias.SelLength = Len(txtdias.Text)
    
End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdescrip.SetFocus
    End If
    
End Sub

Private Sub txtdias_LostFocus()
    
    txtdias.Text = Format(txtdias.Text, "00")
    
End Sub
