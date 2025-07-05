VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MovDiario_xAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento Diario por Almacén"
   ClientHeight    =   2010
   ClientLeft      =   1590
   ClientTop       =   3090
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6210
   Begin VB.TextBox Txtcodalm 
      Height          =   315
      Left            =   1620
      MaxLength       =   3
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   135
      Width           =   450
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   780
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   1376
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
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97124353
         CurrentDate     =   40611
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1845
         TabIndex        =   4
         Top             =   360
         Width           =   450
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   2835
      TabIndex        =   2
      Top             =   1440
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSPanel Txtnomalm 
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      Top             =   135
      Width           =   3330
      _Version        =   65536
      _ExtentX        =   5874
      _ExtentY        =   556
      _StockProps     =   15
      ForeColor       =   -2147483640
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Almacén"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   810
      TabIndex        =   6
      Top             =   180
      Width           =   630
   End
End
Attribute VB_Name = "MovDiario_xAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nmes    As String
Dim amovs_cab(0 To 0)  As a_grabacion


Private Sub cmdresp_Click(Index As Integer)
    If Index = 0 Then
        If Trim(Txtcodalm.Text) <> "" Then
            If TxtFecha.value <= Format(Date, "DD/MM/YYYY") Then
                Procesa_Movimiento
            Else
                MsgBox "Ingrese correctamente la Fecha", vbCritical, "Sistema de Inventario"
                TxtFecha.SetFocus
            End If
        Else
            MsgBox "Ingrese el Almacen", vbCritical, "Sistema de Inventario"
            Txtcodalm.SetFocus
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    TxtFecha.value = Format(Date, "DD/MM/YYYY")
    Me.MousePointer = vbDefault
    cnombase = "TEMPLUS.MDB"
    If cnn_form.State = 1 Then cnn_form.Close
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    Me.top = 1200
    

End Sub

Private Sub Txtcodalm_DblClick()
    
    Txtcodalm_KeyDown 113, 0
    
End Sub

Private Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Me.MousePointer = vbHourglass
        wcod_alm = "" & Txtcodalm.Text
        ayuda_almacen.Show 1
        Txtcodalm.Text = wcod_alm
        Me.MousePointer = vbDefault
        Txtcodalm_KeyPress 13
    End If

End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Txtcodalm.Text = Format(Txtcodalm.Text, "00")
        sql = "select f2codalm,f2nomalm from ef2almacenes where f2codalm = '" & Txtcodalm.Text & "'"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            gcodalm = Rs.Fields("F2codalm")
            Txtnomalm.Caption = Rs.Fields("F2nomalm") & ""
            TxtFecha.SetFocus
        Else
            
            Beep
            Txtcodalm.Text = ""
        End If
        Rs.Close
    End If
    
End Sub


Private Sub Procesa_Movimiento()
On Error GoTo ERRORBD:
    Me.MousePointer = vbHourglass
    nmes = Format(Month(TxtFecha.value), "00")
    
    'CREANDO O ELIMINANDO LA BD
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "DROP TABLE MOV_DIARIO_ALMACEN", cnn_form, adOpenDynamic, adLockOptimistic
    If Rs.State = adStateOpen Then Rs.Close
    '--------------------------
    
    If Trim(Txtcodalm.Text) = "" Then
        'no permitir esto
    Else
        sql = "SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS SALDO INTO MOV_DIARIO_ALMACEN IN '" & wrutatemp & "\TEMPLUS.MDB' "
        sql = sql & "FROM IF4VALES INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) "
        sql = sql & "Where IF4VALES.F4FECVAL<=CVDATE('" & Format(TxtFecha.value, "DD/MM/YYYY") & "') AND IF3VALES.f2codalm = '" & Txtcodalm.Text & "' "
        sql = sql & "GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED;"
    End If
    
    If ctipoadm_bd = "M" Then
        cnn_form.Execute sql
        'AlmacenaQuery_sql sql, cnn_form
    Else
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
    End If
    
    
    'para calcular los productos que tienen movimientos ese dia
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "DROP TABLE MOV_DIARIO_DIA", cnn_form, adOpenDynamic, adLockOptimistic
    If Rs.State = adStateOpen Then Rs.Close

        sql = "SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB,  Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0)) AS INGRESO,Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0)) AS SALIDA INTO MOV_DIARIO_DIA IN '" & wrutatemp & "\TEMPLUS.MDB' "
        sql = sql & "FROM IF4VALES INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) "
        sql = sql & "Where IF4VALES.F4FECVAL=CVDATE('" & Format(TxtFecha.value, "DD/MM/YYYY") & "') AND IF3VALES.f2codalm = '" & Txtcodalm.Text & "' "
        sql = sql & "GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB;"
    
    If ctipoadm_bd = "M" Then
        cnn_form.Execute sql
        'AlmacenaQuery_sql sql, cnn_form
    Else
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
    End If
    
    If Trim(Txtcodalm.Text) = "" Then
        SQLX = "SELECT MOV_DIARIO_ALMACEN.INGRESOS, MOV_DIARIO_ALMACEN.SALIDAS, MOV_DIARIO_ALMACEN.SALDOANT, MOV_DIARIO_ALMACEN.SALDOACT, MOV_DIARIO_ALMACEN.F5CODPRO, MOV_DIARIO_ALMACEN.F5CODFAB, MOV_DIARIO_ALMACEN.F5NOMPRO, MOV_DIARIO_ALMACEN.F7CODMED From MOV_DIARIO_ALMACEN GROUP BY MOV_DIARIO_ALMACEN.INGRESOS, MOV_DIARIO_ALMACEN.SALIDAS, MOV_DIARIO_ALMACEN.SALDOANT, MOV_DIARIO_ALMACEN.SALDOACT, MOV_DIARIO_ALMACEN.F5CODPRO, MOV_DIARIO_ALMACEN.F5CODFAB, MOV_DIARIO_ALMACEN.F5NOMPRO, MOV_DIARIO_ALMACEN.F7CODMED, MOV_DIARIO_ALMACEN.F2CODALM HAVING (((Sum(MOV_DIARIO_ALMACEN.INGRESOS))>0) AND ((Sum(MOV_DIARIO_ALMACEN.SALIDAS))>0)) ORDER BY MOV_DIARIO_ALMACEN.F5CODFAB;"
    Else
        'OJO AQUI REEMPLAZAR SELECT
        SQLX = "SELECT DISTINCTROW MOV_DIARIO_ALMACEN.F5CODPRO, MOV_DIARIO_ALMACEN.F5CODFAB, MOV_DIARIO_ALMACEN.F5NOMPRO, MOV_DIARIO_ALMACEN.F7CODMED AS UM, [MOV_DIARIO_ALMACEN]![SALDO]+[MOV_DIARIO_DIA]![SALIDA]-[MOV_DIARIO_DIA]![INGRESO] AS SALDOANT, MOV_DIARIO_DIA.INGRESO AS INGRESOS, MOV_DIARIO_DIA.SALIDA AS SALIDAS, MOV_DIARIO_ALMACEN.SALDO AS SALDOACT"
        SQLX = SQLX + " FROM MOV_DIARIO_ALMACEN INNER JOIN MOV_DIARIO_DIA ON (MOV_DIARIO_ALMACEN.F5CODPRO = MOV_DIARIO_DIA.F5CODPRO) AND (MOV_DIARIO_ALMACEN.F2CODALM = MOV_DIARIO_DIA.F2CODALM) WHERE (MOV_DIARIO_ALMACEN.F2CODALM = '" & Txtcodalm.Text & "' )"
       ' SQLX = "SELECT MOV_DIARIO_ALMACEN.INGRESOS, MOV_DIARIO_ALMACEN.SALIDAS, MOV_DIARIO_ALMACEN.SALDOANT, MOV_DIARIO_ALMACEN.SALDOACT, MOV_DIARIO_ALMACEN.F5CODPRO, MOV_DIARIO_ALMACEN.F5CODFAB, MOV_DIARIO_ALMACEN.F5NOMPRO, MOV_DIARIO_ALMACEN.F7CODMED From MOV_DIARIO_ALMACEN GROUP BY MOV_DIARIO_ALMACEN.INGRESOS, MOV_DIARIO_ALMACEN.SALIDAS, MOV_DIARIO_ALMACEN.SALDOANT, MOV_DIARIO_ALMACEN.SALDOACT, MOV_DIARIO_ALMACEN.F5CODPRO, MOV_DIARIO_ALMACEN.F5CODFAB, MOV_DIARIO_ALMACEN.F5NOMPRO, MOV_DIARIO_ALMACEN.F7CODMED, MOV_DIARIO_ALMACEN.F2CODALM HAVING (((Sum(MOV_DIARIO_ALMACEN.INGRESOS))>0) AND ((Sum(MOV_DIARIO_ALMACEN.SALIDAS))>0) AND ((MOV_DIARIO_ALMACEN.F2CODALM)='" & Trim(Txtcodalm.Text) & "')) ORDER BY MOV_DIARIO_ALMACEN.F5CODFAB;"
    End If
    
    Me.MousePointer = vbDefault
    
    With acr_MovDiario_xAlmacen
        
        .fldempresa.Text = wnomcia
        .fldtitulo.Text = "Del  " & Format(TxtFecha.value, "dd/mm/yyyy")
        .fldFecha.Text = Format(Date, "dd/mm/yyyy")
        .fldcodalmacen.Visible = True
        .fldnomalmacen.Visible = True
        .lblAlmacen.Visible = True
        .fldcodalmacen.Text = Txtcodalm.Text
        .fldnomalmacen.Text = Txtnomalm.Caption
        .datconexion.ConnectionString = cnn_form
        .datconexion.Source = SQLX
        .Show vbModal
        
    End With
ERRORBD:

If Err.Number = -2147217865 Then
    Resume Next
End If
    
End Sub

Private Sub grabar()

    amovs_cab(0).campo = "SALDOANT": amovs_cab(0).valor = (Rs.Fields("SALDOANT") + Debm) - Habm: amovs_cab(0).TIPO = "N"
    
    '------- ACTUALIZAR STOCKS
    'GRABA_REGISTRO_logistica amovs_cab(), "MOV_DIARIO_ALMACEN", "M", 0, cnn_form, "F2CODALM = '" & rs.Fields("f2codalm") & "' AND F5CODPRO = '" & rs.Fields("f5codpro") & "'"

End Sub

Private Sub grabar_dia()
Dim amovs_cabD(0 To 1)  As a_grabacion

    amovs_cabD(0).campo = "SALIDAS": amovs_cabD(0).valor = Habm: amovs_cabD(0).TIPO = "N"
    amovs_cabD(1).campo = "INGRESOS": amovs_cabD(1).valor = Debm: amovs_cabD(1).TIPO = "N"
    
    '------- ACTUALIZAR STOCKS
    GRABA_REGISTRO_logistica amovs_cabD(), "MOV_DIARIO_ALMACEN", "M", 1, cnn_form, "F2CODALM = '" & Rs.Fields("f2codalm") & "' AND F5CODPRO = '" & Rs.Fields("f5codpro") & "'"

End Sub

Private Sub txtfecha_GotFocus()
TxtFecha.value = Format(Date, "DD/MM/YYYY")
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdresp(0).SetFocus
End If
End Sub

