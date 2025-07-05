VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilRecalcularCostoPromedio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recalcular Costo Promedio"
   ClientHeight    =   4215
   ClientLeft      =   2355
   ClientTop       =   2535
   ClientWidth     =   11370
   Icon            =   "frmUtilRecalcularCostoPromedio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   11370
   Begin VB.Frame Frame1 
      Caption         =   " Proceso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   11175
      Begin MSComctlLib.ProgressBar pgbProceso2 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbProceso1 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblProceso2 
         Caption         =   "Proceso 2"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   10575
      End
      Begin VB.Label lblProceso1 
         Caption         =   "Proceso 1"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   10575
      End
   End
   Begin VB.CommandButton cmdOperacion 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOperacion 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame fraDatos 
      Caption         =   " Datos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11175
      Begin VB.TextBox txtCodAlmacen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtCodProducto 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   3960
         TabIndex        =   0
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   128647169
         CurrentDate     =   41939
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   5880
         TabIndex        =   16
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   128647169
         CurrentDate     =   41939
      End
      Begin VB.Label Label5 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblAlmacen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   8775
      End
      Begin VB.Label Label2 
         Caption         =   "hasta"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   8775
      End
      Begin VB.Label Label3 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Iniciar desde"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmUtilRecalcularCostoPromedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFichero As String

Private Sub listarAnnosVale()
    objAyudaVale.listarAnnoValeSoloSeleccion cmbAnno
    'objSqlAyudaVale.listarAnnoValeSoloSeleccion cmbAnno
    
    If cmbAnno.ListCount > 0 Then
        cmbAnno.ListIndex = cmbAnno.ListCount - 1
    End If
End Sub

Private Sub listarMesesVale()
    objAyudaVale.listarMesValeSoloSeleccion cmbMes, Trim(cmbAnno.Text)
    
    'objSqlAyudaVale.listarMesValeSoloSeleccion cmbMes, Trim(cmbAnno.Text)
    
    If cmbMes.ListCount > 0 Then
        cmbMes.ListIndex = cmbMes.ListCount - 1
    End If
End Sub

Private Sub inicializarControles()
    strFichero = App.Path & strNombreFicheroConfigCPgeneral
    
    txtCodAlmacen.Text = vbNullString
        lblalmacen.Caption = vbNullString
    
    dtpDesde.MinDate = ModUtilitario.sGetINI(strFichero, "ConfigCP", "FechaInicioOperacionesCP", "l")
    dtpDesde.MaxDate = Date
    
    txtCodProducto.Text = vbNullString
        lblProducto.Caption = "Todos los Productos (*)"
    
    lblProceso1.Caption = vbNullString
    pgbProceso1.Value = 0
    lblProceso2.Caption = vbNullString
    pgbProceso2.Value = 0
    
    cmdOperacion(0).Enabled = True
    cmdOperacion(1).Enabled = True
    fraDatos.Enabled = True
End Sub

Private Sub cmbAnno_Click()
    listarMesesVale
End Sub

Private Sub cmbMes_Click()
    DoEvents
    
    If CDate(DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)) < CDate(ModUtilitario.sGetINI(strFichero, "ConfigCP", "FechaInicioOperacionesCP", "l")) Then
        dtpDesde.Value = ModUtilitario.sGetINI(strFichero, "ConfigCP", "FechaInicioOperacionesCP", "l")
        dtpHasta.Value = DateSerial(Year(dtpDesde.Value), Month(dtpDesde.Value) + 1, 0)
    Else
        dtpDesde.Value = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
        dtpHasta.Value = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 1, 0)
    End If
End Sub

Private Sub cmdOperacion_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim strMonedaPredeterminada As String
            
            If Trim(txtCodAlmacen.Text) = vbNullString Or Trim(lblalmacen.Caption) = vbNullString Then
                MsgBox "Almacen no especificado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            strMonedaPredeterminada = ModUtilitario.sGetINI(strFichero, "ConfigCP", "MonedaPredeterminada", "l")
            
            cmdOperacion(0).Enabled = False
            cmdOperacion(1).Enabled = False
            fraDatos.Enabled = False
''
''            Me.MousePointer = vbHourglass
''
''            abrirCnTemporal
''
''            cnDBTemp.Execute "DELETE FROM TMPUTILORIGENES"
''
''            abrirCnTemporal
''
''            cnDBTemp.Execute "DELETE FROM TMPUTILVALECAB"
''
''            abrirCnTemporal
''
''            cnDBTemp.Execute "DELETE FROM TMPUTILVALEDET"
''
''
''
''
''            SqlCad = vbNullString
''            SqlCad = SqlCad & "INSERT INTO TMPUTILORIGENES IN '" & wrutatemp & "Templus.mdb' "
''            SqlCad = SqlCad & "SELECT * FROM SF1ORIGENES"
''
''            abrirCnnDbBancos
''
''            abrirCnTemporal
''
''            cnn_dbbancos.Execute SqlCad
''
''
''            SqlCad = vbNullString
''            SqlCad = SqlCad & "INSERT INTO TMPUTILVALECAB IN '" & wrutatemp & "Templus.mdb' "
''            SqlCad = SqlCad & "SELECT * FROM IF4VALES WHERE F1CODORI NOT IN ('XCS')"
''
''            abrirCnnDbBancos
''
''            abrirCnTemporal
''
''            cnn_dbbancos.Execute SqlCad
''
''            SqlCad = vbNullString
''            SqlCad = SqlCad & "INSERT INTO TMPUTILVALEDET IN '" & wrutatemp & "Templus.mdb' "
''            SqlCad = SqlCad & "SELECT "
''            SqlCad = SqlCad & "DET.* "
''            SqlCad = SqlCad & "FROM "
''            SqlCad = SqlCad & "IF3VALES AS DET "
''            SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON (CAB.F4NUMVAL = DET.F4NUMVAL) AND (CAB.F2CODALM = DET.F2CODALM) "
''            SqlCad = SqlCad & "WHERE "
''            SqlCad = SqlCad & "CAB.F1CODORI NOT IN ('XCS')"
''
''            abrirCnnDbBancos
''
''            abrirCnTemporal
''
''            cnn_dbbancos.Execute SqlCad
            
''            Me.MousePointer = vbDefault
            
            With objAyudaVale
'                .recalcularCostoPromedioProducto Trim(txtCodAlmacen.Text), Trim(dtpDesde.Value & ""), Trim(dtpHasta.Value & ""), , _
                                                    IIf(Trim(lblProducto.Caption) <> "Todos los Productos (*)", Trim(txtCodProducto.Text), vbNullString), lblProceso1, pgbProceso1


                .recalcularCostoPromedioProducto Trim(dtpDesde.Value & ""), txtCodAlmacen.Text, _
                                                     "S", _
                                                    IIf(Trim(lblProducto.Caption) <> "Todos los Productos (*)", Trim(txtCodProducto.Text), vbNullString), lblProceso1, pgbProceso1
            End With
            
'            With objSqlAyudaVale
'                .recalcularCostoPromedioProducto Trim(dtpDesde.value & ""), _
'                                                    Trim(dtpHasta.value & ""), _
'                                                    strMonedaPredeterminada, _
'                                                    IIf(Trim(lblProducto.Caption) <> "Todos los Productos (*)", Trim(txtCodProducto.Text), vbNullString), lblProceso1, pgbProceso1
'            End With
            
            inicializarControles
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdOperacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    ModUtilitario.deshabilitarBotonCerrarForm frmUtilRecalcularCostoPromedio
    
    inicializarControles
    
    listarAnnosVale
End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodAlmacen_DblClick()
    txtCodAlmacen_KeyDown 113, 0
End Sub

Private Sub txtCodAlmacen_GotFocus()
    txtCodAlmacen.SelStart = 0: txtCodAlmacen.SelLength = Len(txtCodAlmacen.Text)
End Sub

Private Sub txtCodAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        wcod_alm = ""
        ayuda_almacen.Show 1
        Unload ayuda_almacen
        If Len(Trim(wcod_alm)) > 0 Then
            txtCodAlmacen.Text = wcod_alm
            lblalmacen.Caption = wnomalmacen
            txtCodAlmacen_KeyPress 13
        End If
    End If
End Sub

Private Sub txtCodAlmacen_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtCodAlmacen.Text = Format(txtCodAlmacen.Text, "00")
        sql = "select f2codalm,f2nomalm from ef2almacenes where f2codalm = '" & txtCodAlmacen.Text & "'"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            gcodalm = Rs.Fields("F2codalm")
            lblalmacen.Caption = Rs.Fields("F2nomalm") & ""
            'txtproducto.SetFocus
        Else
            
            Beep
            txtCodAlmacen.Text = ""
        End If
        Rs.Close
    End If
End Sub

Private Sub txtCodAlmacen_LostFocus()
     wcod_alm = txtCodAlmacen.Text
     txtCodAlmacen.Text = Format(txtCodAlmacen.Text, "00")
        sql = "select f2codalm,f2nomalm from ef2almacenes where f2codalm = '" & txtCodAlmacen.Text & "'"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            gcodalm = Rs.Fields("F2codalm")
            lblalmacen.Caption = Rs.Fields("F2nomalm") & ""
            'txtproducto.SetFocus
        End If
End Sub

Private Sub txtCodProducto_DblClick()
    txtCodProducto_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProducto_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtCodProducto
End Sub

Private Sub txtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                .Ayuda = True
                .TieneMovimientoAlmacen = True
                .InsumoOP = False
                '.CadenaCorte = InputBox(
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    txtCodProducto.Text = objAyudaBien.Codigo
                    lblProducto.Caption = objAyudaBien.Descripcion
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodProducto_LostFocus()
    If Trim(txtCodProducto.Text) <> vbNullString Then
        txtCodProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        lblProducto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtCodProducto.Text), "T")
        
        If Trim(txtCodProducto.Text) = vbNullString Then
            lblProducto.Caption = "Todos los Productos (*)"
        End If
    Else
        lblProducto.Caption = "Todos los Productos (*)"
    End If
End Sub
