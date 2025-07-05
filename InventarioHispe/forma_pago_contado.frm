VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form forma_pago_contado 
   Caption         =   "Forma de Pago : Contado"
   ClientHeight    =   5565
   ClientLeft      =   2145
   ClientTop       =   1395
   ClientWidth     =   8805
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
   LockControls    =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   8805
   Begin Threed.SSPanel SSPanel1 
      Height          =   4920
      Left            =   90
      TabIndex        =   13
      Top             =   135
      Width           =   8610
      _Version        =   65536
      _ExtentX        =   15187
      _ExtentY        =   8678
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtmontopagado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   180
         Width           =   1410
      End
      Begin Threed.SSFrame fracheque 
         Height          =   1365
         Left            =   180
         TabIndex        =   23
         Top             =   3330
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   2408
         _StockProps     =   14
         Caption         =   " Datos del Cheque "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCheck chkdiferido 
            Height          =   240
            Left            =   225
            TabIndex        =   7
            Top             =   900
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "  Diferido"
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
         Begin Threed.SSPanel pnlbanco 
            Height          =   330
            Left            =   4455
            TabIndex        =   26
            Top             =   405
            Width           =   3660
            _Version        =   65536
            _ExtentX        =   6456
            _ExtentY        =   582
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin VB.TextBox txtbanco 
            Height          =   315
            Left            =   3870
            MaxLength       =   2
            TabIndex        =   6
            Top             =   405
            Width           =   465
         End
         Begin VB.TextBox txtnumcheque 
            Height          =   315
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   5
            Top             =   405
            Width           =   1410
         End
         Begin aBoxCtl.aBox abofechavenc 
            Height          =   315
            Left            =   4455
            TabIndex        =   8
            Top             =   855
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            ABoxType        =   ""
            MinValue        =   "D01000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   1
            AlignmentVertical=   2
            HideSelection   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ApplyTextFormat =   -1  'True
            TextFormat      =   "dd/mm/yyyy"
            Text            =   "07/02/2003"
            DateFormat      =   "dd/mm/yyyy"
            FocusDateFormat =   1
            NegativeForeColor=   255
            NumberFormat    =   17
            DecimalPlaces   =   0
            HotAppearance   =   2
            CalendarTrailingForeColor=   -2147483629
            BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowButton      =   1
            ButtonPicture   =   "forma_pago_contado.frx":0000
            ButtonWidth     =   21
            UpDownWidth     =   14
            NullText        =   ""
            BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalcBtnHotStyle =   4
            CalcBackColor   =   -2147483643
            CalcBtnBackColor=   -2147483643
            CalcBtnDigitColor=   -2147483646
            CalcBtnFuntionColor=   8388736
            CalcDisplayFrameColor=   65535
            CalcHeaderBackColor=   -2147483646
         End
         Begin VB.Label lblfechavenc 
            Caption         =   "Fecha Venc."
            Height          =   240
            Left            =   3285
            TabIndex        =   27
            Top             =   900
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   210
            Left            =   3285
            TabIndex        =   25
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cheque"
            Height          =   210
            Left            =   270
            TabIndex        =   24
            Top             =   450
            Width           =   780
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2400
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   4233
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txttotdol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   1800
            Width           =   1320
         End
         Begin VB.TextBox txttotsol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2295
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   1800
            Width           =   1320
         End
         Begin VB.TextBox txtchedol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5310
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1125
            Width           =   1320
         End
         Begin VB.TextBox txtchesol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2295
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   1125
            Width           =   1320
         End
         Begin VB.TextBox txtefedol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5310
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   675
            Width           =   1320
         End
         Begin VB.TextBox txtefesol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2295
            TabIndex        =   1
            Text            =   "0.00"
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label Label12 
            Caption         =   "S/."
            Height          =   240
            Left            =   1890
            TabIndex        =   29
            Top             =   1845
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "US$"
            Height          =   240
            Left            =   4815
            TabIndex        =   28
            Top             =   1845
            Width           =   330
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   210
            Left            =   810
            TabIndex        =   22
            Top             =   1845
            Width           =   345
         End
         Begin VB.Line Line2 
            X1              =   5130
            X2              =   6795
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line1 
            X1              =   2160
            X2              =   3825
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            Height          =   210
            Left            =   810
            TabIndex        =   21
            Top             =   1170
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dólares"
            Height          =   210
            Left            =   5580
            TabIndex        =   20
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Soles"
            Height          =   210
            Left            =   2835
            TabIndex        =   19
            Top             =   315
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Efectivo"
            Height          =   210
            Left            =   810
            TabIndex        =   18
            Top             =   720
            Width           =   585
         End
      End
      Begin VB.TextBox txttc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7695
         TabIndex        =   0
         Text            =   "0.000"
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtmontodoc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label lblmontopag 
         Caption         =   "US$"
         Height          =   240
         Left            =   5310
         TabIndex        =   31
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monto pagado"
         Height          =   210
         Left            =   4185
         TabIndex        =   30
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T/C"
         Height          =   210
         Left            =   7290
         TabIndex        =   16
         Top             =   225
         Width           =   240
      End
      Begin VB.Label lblmoneda 
         Caption         =   "US$"
         Height          =   240
         Left            =   1890
         TabIndex        =   15
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto del documento"
         Height          =   210
         Left            =   225
         TabIndex        =   14
         Top             =   225
         Width           =   1560
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   315
      Top             =   5175
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "forma_pago_contado.frx":0352
      ToolBars        =   "forma_pago_contado.frx":4F2E
   End
End
Attribute VB_Name = "forma_pago_contado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmoneda             As String
Dim nmonto              As Double
Dim dfecha              As Date
Dim ccodigo             As String
Dim cnombre             As String
Dim cdocumento          As String
Dim ncorrela            As Double
Dim amovs_cab(0 To 18)  As a_grabacion
Dim amovs(0 To 10)      As a_grabacion
Dim amovs_doc(0 To 0)   As a_grabacion

Private Sub abofechavenc_GotFocus()

    abofechavenc.FocusSelect = True

End Sub

Private Sub chkdiferido_Click(Value As Integer)

    If Value = True Then
        lblfechavenc.Visible = True
        abofechavenc.Visible = True
        abofechavenc.Value = Format(Date, "DD/MM/YYYY")
    Else
        lblfechavenc.Visible = False
        abofechavenc.Visible = False
    End If

End Sub

Private Sub chkdiferido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If abofechavenc.Visible = True Then
            abofechavenc.SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()

    sw_ayuda = False
    
    '----------------------- VARIABLES : PARAMETROS DE FACTURACION
    cmoneda = "S"
    nmonto = 1000
    dfecha = "25/01/2002"
    ccodigo = "0001"
    cnombre = "CLIENTE 1"
    cdocumento = "Fac001/0000001"
    ncorrela = 1
    '-----------------------------------------
    
    If cmoneda = "S" Then
        lblmoneda.Caption = "S/."
        lblmontopag.Caption = "S/."
    Else
        lblmoneda.Caption = "US$"
        lblmontopag.Caption = "US$"
    End If
    txtmontodoc.Text = Format(nmonto, "###,###,##0.00")
    
    rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & dfecha & "')", cnn_dbbancos
    If Not rscambios.EOF Then
        txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
    Else
        txttc.Text = Format(0, "0.000")
    End If
    rscambios.Close
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Grabar":
            grabar
        Case "ID_Salir":
            Unload Me
    End Select

End Sub

Private Sub txtbanco_DblClick()

    txtbanco_KeyDown 113, 0

End Sub

Private Sub txtbanco_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodban = ""
        hlp_bancos.Show 1
        sw_ayuda = False
        If Len(Trim(wcodban)) > 0 Then
            txtbanco.Text = wcodban
        End If
        txtbanco_KeyPress 13
    End If

End Sub

Private Sub txtbanco_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkdiferido.SetFocus
    End If

End Sub

Private Sub txtbanco_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtbanco.Text)) > 0 Then
            rsbancos.Open "SELECT * FROM BANCOS WHERE CODIGO='" & txtbanco.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsbancos.EOF Then
                pnlbanco.Caption = Trim(rsbancos.Fields("BANCO") & "")
            Else
                MsgBox "Código del banco no existe. Verifique.", vbInformation, "Atención"
                txtbanco.SetFocus
            End If
            rsbancos.Close
        End If
    End If

End Sub

Private Sub txtchedol_GotFocus()

    txtchedol.SelStart = 0
    txtchedol.SelLength = Len(txtchedol.Text)
    
End Sub

Private Sub txtchedol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If fracheque.Visible = True Then
            txtnumcheque.SetFocus
        End If
    End If

End Sub

Private Sub txtchedol_LostFocus()

    txtchedol.Text = Format(txtchedol.Text, "###,###,##0.00")
    If Val(Format(txtchedol.Text, "0.00")) > 0# Then
        If calcula_total() = True Then
        Else
            MsgBox "El monto de pago no puede ser mayor al monto del documento.", vbInformation, "Atención"
            txtchedol.SetFocus
        End If
    End If
    
    If Val(Format(txtchesol.Text, "0.00")) > 0# Or Val(Format(txtchedol.Text, "0.00")) > 0# Then
        fracheque.Visible = True
    Else
        fracheque.Visible = False
    End If
    
End Sub

Private Sub txtchesol_GotFocus()

    txtchesol.SelStart = 0
    txtchesol.SelLength = Len(txtchesol.Text)
    
End Sub

Private Sub txtchesol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtchedol.SetFocus
    End If
    
End Sub

Private Sub txtchesol_LostFocus()

    txtchesol.Text = Format(txtchesol.Text, "###,###,##0.00")
    If Val(Format(txtchesol.Text, "0.00")) > 0# Then
        If calcula_total() = True Then
        Else
            MsgBox "El monto de pago no puede ser mayor al monto del documento.", vbInformation, "Atención"
            txtchesol.SetFocus
        End If
    End If
    
    If Val(Format(txtchesol.Text, "0.00")) > 0# Or Val(Format(txtchedol.Text, "0.00")) > 0# Then
        fracheque.Visible = True
    Else
        fracheque.Visible = False
    End If

End Sub

Private Sub txtefedol_GotFocus()

    txtefedol.SelStart = 0
    txtefedol.SelLength = Len(txtefedol.Text)
    
End Sub

Private Sub txtefedol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtchesol.SetFocus
    End If

End Sub

Private Sub txtefedol_LostFocus()

    txtefedol.Text = Format(txtefedol.Text, "###,###,##0.00")
    If Val(Format(txtefedol.Text, "0.00")) > 0# Then
        If calcula_total() = True Then
        Else
            MsgBox "El monto de pago no puede ser mayor al monto del documento.", vbInformation, "Atención"
            txtefedol.SetFocus
        End If
    End If
    
End Sub

Private Sub txtefesol_GotFocus()

    txtefesol.SelStart = 0
    txtefesol.SelLength = Len(txtefesol.Text)

End Sub

Private Sub txtefesol_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtefedol.SetFocus
    End If

End Sub

Private Sub txtefesol_LostFocus()

    txtefesol.Text = Format(txtefesol.Text, "###,###,##0.00")
    If Val(Format(txtefesol.Text, "0.00")) > 0# Then
        If calcula_total() = True Then
        Else
            MsgBox "El monto de pago no puede ser mayor al monto del documento.", vbInformation, "Atención"
            txtefesol.SetFocus
        End If
    End If

End Sub

Private Sub txtnumcheque_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtbanco.SetFocus
    End If

End Sub

Private Function calcula_total()
Dim nsoles      As Double
Dim ndolar      As Double
Dim ntotal      As Double
Dim sw_funcion  As Boolean

    sw_funcion = True

    nsoles = Val(Format(txtefesol.Text, "0.00")) + Val(Format(txtchesol.Text, "0.00"))
    ndolar = Val(Format(txtefedol.Text, "0.00")) + Val(Format(txtchedol.Text, "0.00"))
    
    txttotsol.Text = Format(nsoles, "###,###,##0.00")
    txttotdol.Text = Format(ndolar, "###,###,##0.00")
    
    If cmoneda = "S" Then
        ntotal = Val(Format(txttotsol.Text, "0.00"))
        If Val(Format(txttotdol.Text, "0.00")) > 0# Then
            ntotal = ntotal + Val(Format(Val(Format(txttotdol.Text, "0.00")) * txttc.Text, "0.00"))
        End If
    Else
        ntotal = Val(Format(txttotdol.Text, "0.00"))
        If Val(Format(txttotsol.Text, "0.00")) > 0# Then
            ntotal = ntotal + Val(Format(Val(Format(txttotsol.Text, "0.00")) / txttc.Text, "0.00"))
        End If
    End If
    
    If ntotal > nmonto Then
        sw_funcion = False
    Else
        txtmontopagado.Text = Format(ntotal, "###,###,##0.00")
    End If
    
    calcula_total = sw_funcion
    
End Function

Private Sub txttc_GotFocus()

    txttc.SelStart = 0
    txttc.SelLength = Len(txttc.Text)
    
End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtefesol.SetFocus
    End If

End Sub

Private Sub txttc_LostFocus()

    txttc.Text = Format(txttc.Text, "0.000")

End Sub

Private Sub grabar()
Dim ntotal      As Double
Dim nmto        As Double

    SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
    If calcula_total() = True Then
        If Val(Format(txtchesol.Text, "0.00")) > 0# Or Val(Format(txtchedol.Text, "0.00")) > 0# Then
            If Len(Trim(txtnumcheque.Text)) = 0 Then
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                MsgBox "Falta ingresar el número del cheque. Verifique.", vbInformation, "Atención"
                txtnumcheque.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtbanco.Text)) = 0 Then
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                MsgBox "Falta ingresar el banco. Verifique.", vbInformation, "Atención"
                txtnumcheque.SetFocus
                Exit Sub
            End If
        End If
        
        If Val(Format(txtefesol.Text, "0.00")) > 0# Or Val(Format(txtefedol.Text, "0.00")) > 0# Then
            nmto = 0#
            If cmoneda = "S" Then
                nmto = Val(Format(txtefesol.Text, "0.00")) + Val(Format(Val(Format(txtefedol.Text, "0.00")) * txttc.Text, "0.00"))
            Else
                nmto = Val(Format(txtefedol.Text, "0.00")) + Val(Format(Val(Format(txtefesol.Text, "0.00")) / txttc.Text, "0.00"))
            End If
            GRABA_MOVS "Efe", Format(dfecha, "dd/mm/yyyy"), Format(dfecha, "dd/mm/yyyy"), cmoneda, txttc.Text, nmto, 0, "H", "COBRANZA DIA " & Format(dfecha, "DD/MM/YYYY"), cnombre
        End If
        
        If Val(Format(txtchesol.Text, "0.00")) > 0# Or Val(Format(txtchedol.Text, "0.00")) > 0# Then
            nmto = 0#
            If cmoneda = "S" Then
                nmto = Val(Format(txtchesol.Text, "0.00")) + Val(Format(Val(Format(txtchedol.Text, "0.00")) * txttc.Text, "0.00"))
            Else
                nmto = Val(Format(txtchedol.Text, "0.00")) + Val(Format(Val(Format(txtchesol.Text, "0.00")) / txttc.Text, "0.00"))
            End If
            GRABA_MOVS "Efe", Format(dfecha, "dd/mm/yyyy"), Format(dfecha, "dd/mm/yyyy"), cmoneda, txttc.Text, nmto, 0, "H", "COBRANZA DIA " & Format(dfecha, "DD/MM/YYYY"), cnombre
        End If
        
    Else
        MsgBox "El monto de pago no puede ser mayor al monto del documento.", vbInformation, "Atención"
        txtefesol.SetFocus
    End If

End Sub

Private Sub GRABA_MOVS(ptipdocu As String, pdocum As String, pfecha As Date, pmoneda As String, ptipcamb As Double, pmonto As Double, psaldo As Double, pdebhab As String, prefer As String, pnomcodigo As String)
Dim ncorrcomp       As Double
Dim nsaldodoc       As Double
Dim nsaldo          As Double

    ncorrcomp = OBTIENE_CORRELA(cnn_dbbancos)
    amovs_cab(0).campo = "TIPO": amovs_cab(0).valor = "C": amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "VIA_INGR": amovs_cab(1).valor = "2": amovs_cab(1).TIPO = "T"
    amovs_cab(2).campo = "CORRELA": amovs_cab(2).valor = ncorrcomp: amovs_cab(2).TIPO = "N"
    amovs_cab(3).campo = "TIPDOCU": amovs_cab(3).valor = ptipdocu: amovs_cab(3).TIPO = "T"
    amovs_cab(4).campo = "SERDOC": amovs_cab(4).valor = "": amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "DOCUM": amovs_cab(5).valor = pdocum: amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "FECHA": amovs_cab(6).valor = pfecha: amovs_cab(6).TIPO = "F"
    amovs_cab(7).campo = "RUC": amovs_cab(7).valor = cruc: amovs_cab(7).TIPO = "T"
    amovs_cab(8).campo = "CODIGO": amovs_cab(8).valor = ccodigo: amovs_cab(8).TIPO = "T"
    amovs_cab(9).campo = "MONEDA": amovs_cab(9).valor = pmoneda: amovs_cab(9).TIPO = "T"
    amovs_cab(10).campo = "TIPCAM": amovs_cab(10).valor = ptipcamb: amovs_cab(10).TIPO = "N"
    amovs_cab(11).campo = "TOTAL": amovs_cab(11).valor = pmonto: amovs_cab(11).TIPO = "N"
    amovs_cab(12).campo = "SALDO": amovs_cab(12).valor = psaldo: amovs_cab(12).TIPO = "N"
    amovs_cab(13).campo = "DEB_HAB": amovs_cab(13).valor = pdebhab: amovs_cab(13).TIPO = "T"
    amovs_cab(14).campo = "REFERENCIA": amovs_cab(14).valor = prefer: amovs_cab(14).TIPO = "T"
    amovs_cab(15).campo = "FCH_REPO": amovs_cab(15).valor = pfecha: amovs_cab(15).TIPO = "F"
    amovs_cab(16).campo = "ANO_REPO": amovs_cab(16).valor = Year(pfecha): amovs_cab(16).TIPO = "N"
    amovs_cab(17).campo = "NRO_REPO": amovs_cab(17).valor = Day(pfecha) & Format(Month(pfecha), "00"): amovs_cab(17).TIPO = "T"
    amovs_cab(18).campo = "NOMCODIGO": amovs_cab(18).valor = pnomcodigo: amovs_cab(18).TIPO = "T"
    

    amovs(0).campo = "TIPO": amovs(0).valor = "C": amovs(0).TIPO = "T"
    amovs(1).campo = "CODIGO": amovs(1).valor = ccodigo: amovs(1).TIPO = "T"
    amovs(2).campo = "CORR_COMP": amovs(2).valor = ncorrcomp: amovs(2).TIPO = "N"
    amovs(3).campo = "CORR_DCTO": amovs(3).valor = ncorrela: amovs(3).TIPO = "N"
    If pmoneda = "S" Then
        amovs(4).campo = "IMPUTASO": amovs(4).valor = pmonto: amovs(4).TIPO = "N"
        amovs(5).campo = "IMPUTADO": amovs(5).valor = 0: amovs(5).TIPO = "N"
    Else
        amovs(4).campo = "IMPUTASO": amovs(4).valor = 0: amovs(4).TIPO = "N"
        amovs(5).campo = "IMPUTADO": amovs(5).valor = pmonto: amovs(5).TIPO = "N"
    End If
    amovs(6).campo = "TCAMBIO": amovs(6).valor = ptipcamb: amovs(6).TIPO = "N"
    amovs(7).campo = "FCH_REPO": amovs(7).valor = pfecha: amovs(7).TIPO = "F"
    amovs(8).campo = "ANO_REPO": amovs(8).valor = Year(pfecha): amovs(8).TIPO = "T"
    amovs(9).campo = "NRO_REPO": amovs(9).valor = Day(pfecha) & Format(Month(pfecha), "00"): amovs(9).TIPO = "N"
    amovs(10).campo = "FCH_MVTO": amovs(10).valor = pfecha: amovs(10).TIPO = "F"
    
    nsaldo = 0#
    If pmoneda = "S" Then '---- moneda del anticipo
        If lblmoneda.Caption = "S/." Then         '---- moneda del documento
            nsaldo = pmonto
        Else
            nsaldo = Format(pmonto / ptipcamb, "0.00")
        End If
    Else
        If lblmoneda.Caption = "US$" Then
            nsaldo = pmonto
        Else
            nsaldo = Format(pmonto * ptipcamb, "0.00")
        End If
    End If
    
    nsaldodoc = 0#
    RSCTA_DCTO.Open "SELECT SALDO FROM CTA_DCTO WHERE CORRELA=" & ncorrela & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RSCTA_DCTO.EOF Then
        nsaldodoc = Val(RSCTA_DCTO.Fields("SALDO") & "")
    End If
    RSCTA_DCTO.Close
    
    amovs_doc(0).campo = "SALDO": amovs_doc(0).valor = nsaldodoc - nsaldo: amovs_doc(0).TIPO = "N"
    
    GRABA_REGISTRO amovs_cab(), "CTA_DCTO", "A", 18, cnn_dbbancos, ""
    GRABA_REGISTRO amovs(), "CTA_MVTO", "A", 10, cnn_dbbancos, ""
    GRABA_REGISTRO amovs_doc(), "CTA_DCTO", "M", 0, cnn_dbbancos, "TIPO='C' AND CORRELA =" & ncorrela & ""
            
End Sub
