VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form FrmPSistema 
   Caption         =   "Parámetros del Sistema"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "FrmPSistema.frx":0000
      ToolBars        =   "FrmPSistema.frx":1980
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   2550
      Left            =   4155
      TabIndex        =   0
      Top             =   1590
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   4498
      _StockProps     =   14
      Caption         =   "Datos Generales"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.ComboBox CmbCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1272
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2064
         Width           =   1590
      End
      Begin VB.ComboBox CmbTiplet 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1272
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   384
         Width           =   2625
      End
      Begin VB.TextBox Txtscale 
         Height          =   288
         Left            =   1272
         TabIndex        =   7
         Top             =   720
         Width           =   444
      End
      Begin VB.TextBox TxtTasas 
         Height          =   288
         Left            =   1272
         TabIndex        =   6
         Top             =   1056
         Width           =   588
      End
      Begin VB.TextBox Txtporte 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1272
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   1392
         Width           =   588
      End
      Begin VB.TextBox Txtmoneda 
         Height          =   300
         Left            =   1272
         TabIndex        =   4
         Top             =   1710
         Width           =   348
      End
      Begin VB.TextBox Txtcodcli 
         Height          =   288
         Left            =   3552
         TabIndex        =   3
         Top             =   1728
         Width           =   348
      End
      Begin VB.TextBox Txttasad 
         Height          =   288
         Left            =   3312
         TabIndex        =   2
         Top             =   1080
         Width           =   588
      End
      Begin VB.TextBox Txtigv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3312
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1392
         Width           =   588
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cod.Cliente:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   22
         Left            =   2496
         TabIndex        =   18
         Top             =   1776
         Width           =   852
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Tipo Costo:"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   144
         TabIndex        =   17
         Top             =   2112
         Width           =   816
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Moneda:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   13
         Left            =   144
         TabIndex        =   16
         Top             =   1776
         Width           =   636
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Porte:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   144
         TabIndex        =   15
         Top             =   1416
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Scale Mode:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   144
         TabIndex        =   14
         Top             =   744
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "IGV.:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   2844
         TabIndex        =   13
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Tasa  i  $  %"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   2352
         TabIndex        =   12
         Top             =   1104
         Width           =   876
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Tasa  i  S/.   % "
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   144
         TabIndex        =   11
         Top             =   1104
         Width           =   1092
      End
      Begin VB.Label LabelTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Tipo de Letra:"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   144
         TabIndex        =   10
         Top             =   432
         Width           =   996
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1500
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   3990
      _Version        =   65536
      _ExtentX        =   7048
      _ExtentY        =   2646
      _StockProps     =   14
      Caption         =   "Rutas"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox txtconta 
         Height          =   285
         Left            =   948
         MaxLength       =   100
         TabIndex        =   21
         Top             =   720
         Width           =   2835
      End
      Begin VB.TextBox Txtcompras 
         Height          =   285
         Left            =   948
         MaxLength       =   100
         TabIndex        =   20
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Contabilid."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   48
         TabIndex        =   23
         Top             =   756
         Width           =   732
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Compras:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   432
         Width           =   660
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1065
      Left            =   120
      TabIndex        =   24
      Top             =   1590
      Width           =   3990
      _Version        =   65536
      _ExtentX        =   7048
      _ExtentY        =   1884
      _StockProps     =   14
      Caption         =   "Otros Documentos"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox Txtnumlet 
         Height          =   285
         Left            =   2928
         MaxLength       =   7
         TabIndex        =   28
         Top             =   288
         Width           =   765
      End
      Begin VB.TextBox Txtnumped 
         Height          =   285
         Left            =   948
         MaxLength       =   7
         TabIndex        =   27
         Top             =   648
         Width           =   780
      End
      Begin VB.TextBox Txtnumord 
         Height          =   285
         Left            =   948
         MaxLength       =   7
         TabIndex        =   26
         Top             =   288
         Width           =   765
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   780
         Left            =   3984
         TabIndex        =   25
         Top             =   288
         Width           =   24
         _Version        =   65536
         _ExtentX        =   42
         _ExtentY        =   1376
         _StockProps     =   14
         Caption         =   "SSFrame4"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "N° Letra:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   2205
         TabIndex        =   31
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "N° Pedido:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   15
         Left            =   96
         TabIndex        =   30
         Top             =   696
         Width           =   768
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "N° Orden:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   23
         Left            =   96
         TabIndex        =   29
         Top             =   348
         Width           =   708
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1305
      Left            =   120
      TabIndex        =   32
      Top             =   240
      Width           =   8070
      _Version        =   65536
      _ExtentX        =   14245
      _ExtentY        =   2307
      _StockProps     =   14
      Caption         =   "Datos de la Empresa"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox Txtnomemp 
         Height          =   285
         Left            =   975
         MaxLength       =   70
         TabIndex        =   36
         Top             =   288
         Width           =   4245
      End
      Begin VB.TextBox TxtLogEmp 
         Height          =   285
         Left            =   975
         MaxLength       =   70
         TabIndex        =   35
         Top             =   600
         Width           =   6915
      End
      Begin VB.TextBox TxtDirEmp 
         Height          =   285
         Left            =   975
         MaxLength       =   70
         TabIndex        =   34
         Top             =   900
         Width           =   6930
      End
      Begin VB.TextBox TxtRucEmp 
         Height          =   285
         Left            =   6150
         MaxLength       =   11
         TabIndex        =   33
         Top             =   288
         Width           =   1740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Logo :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   75
         TabIndex        =   40
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Dirección:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   75
         TabIndex        =   39
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Nombre :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   75
         TabIndex        =   38
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "R.U.C.:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   17
         Left            =   5472
         TabIndex        =   37
         Top             =   336
         Width           =   528
      End
   End
End
Attribute VB_Name = "FrmPSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WCONT           As Integer
Dim wname(200)      As String * 30
Dim sw_cabecera     As Boolean

Private Sub CmbCosto_Change()
If Trim(CmbCosto.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub CmbTiplet_Change()
If Trim(CmbTiplet.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub CmbTiplet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
Parametros

End Sub

Private Sub Form_Load()
Me.Height = 5025
Me.Width = 8745
sw_cabecera = False
    For WCONT = 0 To Printer.FontCount - 1
        wname(WCONT) = Printer.Fonts(WCONT)
    Next WCONT
    
On Error GoTo ERROR01
    CmbTiplet.Clear
    For WCONT = 0 To Printer.FontCount
        CmbTiplet.AddItem wname(WCONT)
    Next WCONT
    CmbTiplet.ListIndex = 0

    CmbCosto.Clear
    CmbCosto.AddItem "Promedio"
    CmbCosto.AddItem "P.E.P.S."
    CmbCosto.ListIndex = 0
ERROR01: Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo seguir

    RsParametro.Close
    
Exit Sub
seguir: Resume Next
End Sub

Private Sub GRABAR()
Dim csql As String

Set RsParametro = New ADODB.Recordset
SQL = "SELECT * FROM SF1PARAIN WHERE F1CODEMP = '" & wempresa & "'"
If RsParametro.State = adStateOpen Then RsParametro.Close
RsParametro.Open SQL, cnn_control, adOpenDynamic, adLockOptimistic
If Not RsParametro.EOF Then
    csql = "UPDATE SF1PARAIN SET F1NOMEMP = '" & Txtnomemp.Text & "',F1LOGEMP = '" & TxtLogEmp.Text & "' ,F1DIREMP = '" & TxtDirEmp.Text & "' ,F1RUCEMP = '" & TxtRucEmp.Text & "',F1COMPRAS = '" & Txtcompras.Text & "', " & _
    " F1NUMORD = '" & Txtnumord.Text & "' ,F1NUMPED = '" & Txtnumped.Text & "' ,F1NUMLET = '" & Txtnumlet.Text & "' , F1FONNAM = '" & CmbTiplet.ListIndex & "',F1SCALE = " & Val(Format(Txtscale.Text, "#0")) & ",F1TASASOL = " & Val(Format(TxtTasas.Text, "#0.00")) & " , " & _
    " F1TASADOL = " & Val(Format(Txttasad.Text, "#0.00")) & ",F1MONEDA = '" & Txtmoneda.Text & "',F1CODCLI = '" & Txtcodcli.Text & "' ,F1TIPCOSTO = '" & Format(CmbCosto.ListIndex, "0") & "',f1contabilidad = '" & Trim(txtconta.Text) & "' WHERE F1CODEMP = '" & wempresa & "'"
    cnn_control.Execute (csql)
    
    gmoneda = "" & Txtmoneda.Text
    rutaconta = txtconta.Text & "\"
    RutaCom = Txtcompras.Text & "\"
End If

End Sub

Private Sub Parametros()
    
On Error GoTo ERROR02
Set RsParametro = New ADODB.Recordset
SQL = "SELECT * FROM SF1PARAIN WHERE F1CODEMP ='" & wempresa & "'"
If RsParametro.State = adStateOpen Then RsParametro.Close
RsParametro.Open SQL, cnn_control, adOpenDynamic, adLockOptimistic
If Not RsParametro.EOF Then

    Txtnomemp.Text = IIf(IsNull(RsParametro.Fields("F1NOMEMP")), "", RsParametro.Fields("F1NOMEMP"))
    TxtLogEmp.Text = IIf(IsNull(RsParametro.Fields("F1LOGEMP")), "", RsParametro.Fields("F1LOGEMP"))
    TxtDirEmp.Text = IIf(IsNull(RsParametro.Fields("F1DIREMP")), "", RsParametro.Fields("F1DIREMP"))
    TxtRucEmp.Text = IIf(IsNull(RsParametro.Fields("F1RUCEMP")), "", RsParametro.Fields("F1RUCEMP"))

    Txtcompras.Text = "" & RsParametro.Fields("F1COMPRAS")
    Txtnumlet.Text = IIf(IsNull(RsParametro.Fields("F1NUMLET")), "", RsParametro.Fields("F1NUMLET"))
    Txtnumord.Text = IIf(IsNull(RsParametro.Fields("F1NUMORD")), "", RsParametro.Fields("F1NUMORD"))
    Txtnumped.Text = IIf(IsNull(RsParametro.Fields("F1NUMPED")), "", RsParametro.Fields("F1NUMPED"))
    
    CmbTiplet.ListIndex = IIf(IsNull(RsParametro.Fields("F1FONNAM")), 9, RsParametro.Fields("F1FONNAM"))
    Txtscale.Text = IIf(IsNull(RsParametro.Fields("F1SCALE")), "4", RsParametro.Fields("F1SCALE"))
    TxtTasas.Text = IIf(IsNull(RsParametro.Fields("F1TASASOL")), " ", RsParametro.Fields("F1TASASOL"))
    Txttasad.Text = IIf(IsNull(RsParametro.Fields("F1TASADOL")), " ", RsParametro.Fields("F1TASADOL"))
    Txtmoneda.Text = IIf(IsNull(RsParametro.Fields("F1MONEDA")), "S", RsParametro.Fields("F1MONEDA"))
    Txtcodcli.Text = IIf(IsNull(RsParametro.Fields("F1CODCLI")), "C", RsParametro.Fields("F1CODCLI"))
    CmbCosto.ListIndex = Val(Format(RsParametro.Fields("F1TIPCOSTO"), "0"))
    txtconta.Text = "" & RsParametro.Fields("F1contabilidad")

ERROR02: Resume Next
End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
Case "ID_Grabar"
    
        Me.MousePointer = 11
        GRABAR
        MsgBox "Datos Grabados", vbInformation, "Sistema de Inventarios"
        sw_cabecera = False
        Me.MousePointer = 1
Case "ID_Salir"
    If sw_cabecera = True Then
        If MsgBox("Desea Guardar los Cambios", vbYesNo + vbInformation, "Sistema de Inventarios") = vbYes Then
            GRABAR
            sw_cabecera = False
            Unload Me
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Select

End Sub


Private Sub Txtcodcli_Change()
If Trim(Txtcodcli.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtcompras_Change()
If Trim(Txtcompras.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtcompras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtconta_Change()
If Trim(txtconta.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub TxtDirEmp_Change()
If Trim(TxtDirEmp.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub TxtDirEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtigv_Change()
If Trim(Txtigv.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtigv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtLogEmp_Change()
If Trim(TxtLogEmp.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub TxtLogEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtmoneda_Change()
If Trim(Txtmoneda.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtmoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtnomemp_Change()
If Trim(Txtnomemp.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtnomemp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtnumlet_Change()
If Trim(Txtnumlet.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtnumlet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtnumord_Change()
If Trim(Txtnumord.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtnumord_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtnumped_Change()
If Trim(Txtnumped.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtnumped_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtporte_Change()
If Trim(Txtporte.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtRucEmp_Change()
If Trim(TxtRucEmp.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub TxtRucEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txtscale_Change()
If Trim(Txtscale.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtscale_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txttasad_Change()
If Trim(Txttasad.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txttasad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtTasas_Change()
If Trim(TxtTasas.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub TxtTasas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


