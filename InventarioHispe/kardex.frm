VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form kardex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex"
   ClientHeight    =   4320
   ClientLeft      =   825
   ClientTop       =   3405
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8670
   Begin VB.TextBox txtOrigen 
      Height          =   285
      Left            =   1005
      MaxLength       =   50
      TabIndex        =   21
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Frame SSFrame1 
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
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin VB.TextBox txtCcosto 
         Height          =   285
         Left            =   885
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Frame SSFrame3 
         Caption         =   " Moneda "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   2760
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
         Begin Threed.SSOption optmoneda 
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optmoneda 
            Height          =   240
            Index           =   1
            Left            =   1200
            TabIndex        =   20
            Top             =   360
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Dólares"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   " Tipo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   2760
         TabIndex        =   15
         Top             =   1920
         Width           =   5535
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   0
            Left            =   1800
            TabIndex        =   16
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Valorizado"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   300
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "No Valorizado"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   2
            Left            =   3120
            TabIndex        =   27
            Top             =   300
            Width           =   2010
            _Version        =   65536
            _ExtentX        =   3545
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Resumen Valorizado"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   " Rango de Fechas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
         Begin VB.ComboBox cmbMes 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker aboHasta 
            Height          =   315
            Left            =   840
            TabIndex        =   11
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   135921665
            CurrentDate     =   40611
         End
         Begin MSComCtl2.DTPicker aboDesde 
            Height          =   315
            Left            =   840
            TabIndex        =   12
            Top             =   765
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   135921665
            CurrentDate     =   40611
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   28
            Top             =   375
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   14
            Top             =   760
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   13
            Top             =   1215
            Width           =   420
         End
      End
      Begin VB.TextBox txtAlmacen 
         Height          =   285
         Left            =   885
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtProducto 
         Height          =   285
         Left            =   885
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label pnlCcosto 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   26
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label pnlOrigen 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   25
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "C.Costo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   645
      End
      Begin VB.Label pnlAlmacen 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2655
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label pnlProducto 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2655
         TabIndex        =   6
         Top             =   600
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtmedida 
      Height          =   285
      Left            =   9480
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "kardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wanioo As String


Private Sub aboDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub aboHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbMes_Change()
    Dim mesSeleccionado As String
    Dim fechaInicio As Date
    Dim fechaFin As Date
    
    mesSeleccionado = cmbMes.Text ' Obtener el mes seleccionado en texto
    
    Select Case mesSeleccionado
        Case "ENERO"
            fechaInicio = CDate("01/01/" & wanioo)
            fechaFin = CDate("31/01/" & wanioo)
        Case "FEBRERO"
            fechaInicio = CDate("01/02/" & wanioo)
            ' Verificar si es año bisiesto
            If (CInt(wanioo) Mod 4 = 0 And CInt(wanioo) Mod 100 <> 0) Or (CInt(wanioo) Mod 400 = 0) Then
                fechaFin = CDate("29/02/" & wanioo)
            Else
                fechaFin = CDate("28/02/" & wanioo)
            End If
        Case "MARZO"
            fechaInicio = CDate("01/03/" & wanioo)
            fechaFin = CDate("31/03/" & wanioo)
        Case "ABRIL"
            fechaInicio = CDate("01/04/" & wanioo)
            fechaFin = CDate("30/04/" & wanioo)
        Case "MAYO"
            fechaInicio = CDate("01/05/" & wanioo)
            fechaFin = CDate("31/05/" & wanioo)
        Case "JUNIO"
            fechaInicio = CDate("01/06/" & wanioo)
            fechaFin = CDate("30/06/" & wanioo)
        Case "JULIO"
            fechaInicio = CDate("01/07/" & wanioo)
            fechaFin = CDate("31/07/" & wanioo)
        Case "AGOSTO"
            fechaInicio = CDate("01/08/" & wanioo)
            fechaFin = CDate("31/08/" & wanioo)
        Case "SEPTIEMBRE"
            fechaInicio = CDate("01/09/" & wanioo)
            fechaFin = CDate("30/09/" & wanioo)
        Case "OCTUBRE"
            fechaInicio = CDate("01/10/" & wanioo)
            fechaFin = CDate("31/10/" & wanioo)
        Case "NOVIEMBRE"
            fechaInicio = CDate("01/11/" & wanioo)
            fechaFin = CDate("30/11/" & wanioo)
        Case "DICIEMBRE"
            fechaInicio = CDate("01/12/" & wanioo)
            fechaFin = CDate("31/12/" & wanioo)
    End Select
    
    ' Asignar las fechas seleccionadas a los controles
    aboDesde.Value = Format(fechaInicio, "dd/mm/yyyy")
    aboHasta.Value = Format(fechaFin, "dd/mm/yyyy")
End Sub


Private Sub cmbMes_Click()
    Dim mesSeleccionado As String
    Dim fechaInicio As Date
    Dim fechaFin As Date
    
    mesSeleccionado = Trim(left(cmbMes.Text, 15)) ' Obtener el mes seleccionado en texto
    
    Select Case mesSeleccionado
        Case "ENERO"
            fechaInicio = CDate("01/01/" & wanioo)
            fechaFin = CDate("31/01/" & wanioo)
        Case "FEBRERO"
            fechaInicio = CDate("01/02/" & wanioo)
            ' Verificar si es año bisiesto
            If (CInt(wanioo) Mod 4 = 0 And CInt(wanioo) Mod 100 <> 0) Or (CInt(wanioo) Mod 400 = 0) Then
                fechaFin = CDate("29/02/" & wanioo)
            Else
                fechaFin = CDate("28/02/" & wanioo)
            End If
        Case "MARZO"
            fechaInicio = CDate("01/03/" & wanioo)
            fechaFin = CDate("31/03/" & wanioo)
        Case "ABRIL"
            fechaInicio = CDate("01/04/" & wanioo)
            fechaFin = CDate("30/04/" & wanioo)
        Case "MAYO"
            fechaInicio = CDate("01/05/" & wanioo)
            fechaFin = CDate("31/05/" & wanioo)
        Case "JUNIO"
            fechaInicio = CDate("01/06/" & wanioo)
            fechaFin = CDate("30/06/" & wanioo)
        Case "JULIO"
            fechaInicio = CDate("01/07/" & wanioo)
            fechaFin = CDate("31/07/" & wanioo)
        Case "AGOSTO"
            fechaInicio = CDate("01/08/" & wanioo)
            fechaFin = CDate("31/08/" & wanioo)
        Case "SETIEMBRE"
            fechaInicio = CDate("01/09/" & wanioo)
            fechaFin = CDate("30/09/" & wanioo)
        Case "OCTUBRE"
            fechaInicio = CDate("01/10/" & wanioo)
            fechaFin = CDate("31/10/" & wanioo)
        Case "NOVIEMBRE"
            fechaInicio = CDate("01/11/" & wanioo)
            fechaFin = CDate("30/11/" & wanioo)
        Case "DICIEMBRE"
            fechaInicio = CDate("01/12/" & wanioo)
            fechaFin = CDate("31/12/" & wanioo)
    End Select
    
    ' Asignar las fechas seleccionadas a los controles
    aboDesde.Value = Format(fechaInicio, "dd/mm/yyyy")
    aboHasta.Value = Format(fechaFin, "dd/mm/yyyy")
End Sub



Private Sub cmdaceptar_Click()
    If Trim(txtAlmacen.Text) <> "" Then
            If DateDiff("d", CDate(aboDesde.Value), CDate(aboHasta.Value)) >= 0 Then
                imprimir
            Else
                MsgBox "La fecha de Inicio debe ser menor que fecha del final", vbCritical, "Sistema de Logística"
                aboDesde.SetFocus
            End If
    Else
        MsgBox "Ingresar el codigo de Almacen", vbCritical, "Sistema de Logística"
        txtAlmacen.SetFocus
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1200
    wanioo = "20" & right(wempresa, 2)

    listarMesesVale
    aboDesde.Value = Format("01/" & Month(Date) & "/" & wanioo, "dd/mm/yyyy")
    aboHasta.Value = Format(Date, "dd/mm/yyyy")
    Me.MousePointer = vbDefault
End Sub
Private Sub listarMesesVale()
    objAyudaVale.listarMesValeSoloSeleccion cmbMes, wanioo
    
    If cmbMes.ListCount > 0 Then
        cmbMes.ListIndex = cmbMes.ListCount - 1
    End If
End Sub

Private Sub imprimir()
    Dim csql        As String
    Dim cad         As String, cad2 As String
    Dim i           As Integer
    
    Dim rpt1 As New acr_kardex
    
    ccodprod = txtProducto.Text
    wcod_alm = txtAlmacen.Text
    
    rpt1.CodigoAlmacen = Trim(txtAlmacen.Text)
    rpt1.CodigoProducto = Trim(txtProducto.Text)
    rpt1.Fecha = Format(aboDesde.Value, "Short Date")
    rpt1.CodigoMoneda = IIf(CBool(optmoneda(0).Value), "S", "D")
    
    rpt1.fldempresa.Text = wnomcia
    
    If Len(pnlCcosto.Caption) > 0 And Len(txtCcosto.Text) > 0 Then
        cad2 = "C.C: " & pnlCcosto.Caption
    End If
    
    If Len(pnlOrigen.Caption) > 0 And Len(txtOrigen.Text) > 0 Then
        cad2 = cad2 & " - Origen: " & pnlOrigen.Caption
    End If
    
    'rpt1.fldtitulo.Text = cad2 & " - Del  " & Format(aboDesde.Value, "dd/mm/yyyy") & "  al  " & Format(aboHasta.Value, "dd/mm/yyyy")
'    rpt1.fldFecha.Text = Format(Date, "dd/mm/yyyy")
    rpt1.lblalmacen.Visible = True
'    rpt1.fldcodalmacen.Visible = True
    rpt1.fldnomalmacen.Visible = True
    rpt1.fldcodalmacen.Text = txtAlmacen.Text
    rpt1.fldnomalmacen.Text = pnlAlmacen.Caption
    rpt1.datconexion.ConnectionString = cnn_dbbancos
    rpt1.lblPeriodo.Caption = Trim(left(cmbMes, 15)) & " " & wanioo
    rpt1.lblSaldoInicial.Caption = aboDesde.Value
    If opttipo(0).Value = True Then
        If optmoneda(0).Value = True Then
            
                csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, IF4VALES.F4CENTRO,"
                csql = csql & "IF4VALES.F4TIPDOC,IF4VALES.NUMORDEN,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC) AS F4SERDOC,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC)AS F4NUMDOC, IF3VALES.F3VALVTA, IF4VALES.F4FECVAL, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,NULL) AS ENTRADAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,NULL) AS SALIDAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA,NULL) AS ENTRADACU, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA,NULL) AS SALIDACU, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA * IF3VALES.F3CANPRO,NULL) AS ENTRADACOS, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA * IF3VALES.F3CANPRO,NULL) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI, EF7MEDIDAS.UsuMod, SF1ORIGENES.F1CODORIEXTERNO "
                csql = csql & "FROM ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED "
                csql = csql & "Where "
                csql = csql & "IF4VALES.F1CODORI NOT IN ('XCS','SIN') AND "
                    
                    If Len(Trim(ccodprod)) > 0 Then
                        csql = csql & "(IF5PLA.F5CODPRO='" & ccodprod & "') And "
                    End If
                    
                    ' Modificación para múltiples orígenes
                    If Len(Trim(txtOrigen.Text)) > 0 Then
                        Dim origenes As String
                        Dim origenArray() As String
                        
                        origenArray = Split(Trim(txtOrigen.Text), "|")
                        
                        ' Si hay más de un origen, usar IN, de lo contrario mantener el formato original
                        If UBound(origenArray) > 0 Then
                            origenes = "('" & Join(origenArray, "','") & "')"
                            csql = csql & "(IF4VALES.F1CODORI IN " & origenes & ") And "
                        Else
                            csql = csql & "(IF4VALES.F1CODORI='" & Trim(txtOrigen.Text) & "') And "
                        End If
                    End If

                    If Len(Trim(txtCcosto.Text)) > 0 Then
                        csql = csql & "(IF4VALES.F4CENTRO='" & txtCcosto.Text & "') And "
                    End If

                csql = csql & "((IF4VALES.F4FECVAL) >=cvdate('" & aboDesde.Value & "') "
                csql = csql & "And (IF4VALES.F4FECVAL) <= cvdate('" & aboHasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') "
                csql = csql & "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL"
            
'            rpt1.LblIngreso.Caption = "S/"
'            rpt1.LblSalida.Caption = "S/"
'            rpt1.LblSaldos.Caption = "S/"
        Else
                csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, "
                csql = csql & "IF4VALES.F4TIPDOC,IF4VALES.NUMORDEN,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC) AS F4SERDOC,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC)AS F4NUMDOC, IF3VALES.F3VALDOL, IF4VALES.F4FECVAL, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALDOL * IF3VALES.F3CANPRO ,0) AS ENTRADACOS, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALDOL * IF3VALES.F3CANPRO,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI "
                csql = csql & "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) "
                csql = csql & "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) "
                csql = csql & "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI "
                csql = csql & "Where "
                csql = csql & "IF4VALES.F1CODORI NOT IN ('XCS') AND "
                
                ' Modificación para múltiples orígenes
                If Len(Trim(txtOrigen.Text)) > 0 Then

                    
                    origenArray = Split(Trim(txtOrigen.Text), "|")
                    
                    ' Si hay más de un origen, usar IN, de lo contrario mantener el formato original
                    If UBound(origenArray) > 0 Then
                        origenes = "('" & Join(origenArray, "','") & "')"
                        csql = csql & "(IF4VALES.F1CODORI IN " & origenes & ") And "
                    Else
                        csql = csql & "(IF4VALES.F1CODORI='" & Trim(txtOrigen.Text) & "') And "
                    End If
                End If
                
                csql = csql & "((IF4VALES.F4FECVAL) >=cvdate('" & aboDesde.Value & "') And (IF4VALES.F4FECVAL) <=cvdate('" & aboHasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') "
                csql = csql & "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL"

'            rpt1.LblIngreso.Caption = "US$"
'            rpt1.LblSalida.Caption = "US$"
'            rpt1.LblSaldos.Caption = "US$"
        End If
        rpt1.datconexion.ConnectionString = cnn_dbbancos
        rpt1.datconexion.Source = csql
        rpt1.Caption = App.ProductName & " - Kardex Valorizado"
        rpt1.Show 'vbModal
    ElseIf opttipo(1).Value = True Then
        Dim rpt2 As New acr_kardex_nv
        
        If Len(pnlCcosto.Caption) > 0 And Len(txtCcosto.Text) > 0 Then
            cad2 = "C.C: " & pnlCcosto.Caption
        End If
        
        If Len(pnlOrigen.Caption) > 0 And Len(txtOrigen.Text) > 0 Then
            cad2 = cad2 & " - Origen: " & pnlOrigen.Caption
        End If
        
        If optmoneda(0).Value = True Then
            cad = cad & " S/"
        Else
            cad = cad & " US$"
        End If
        
        rpt2.CodigoAlmacen = Trim(txtAlmacen.Text)
        rpt2.CodigoProducto = Trim(txtProducto.Text)
        rpt2.Fecha = Format(aboDesde.Value, "Short Date")
        
        rpt2.fldempresa.Text = wnomcia
        rpt2.fldtitulo.Text = cad2 & " - Del  " & Format(aboDesde.Value, "dd/mm/yyyy") & "  al  " & Format(aboHasta.Value, "dd/mm/yyyy") & "   " & cad
        rpt2.fldFecha.Text = Format(Date, "dd/mm/yyyy")
        rpt2.lblalmacen.Visible = True
        rpt2.fldcodalmacen.Visible = True
        rpt2.fldnomalmacen.Visible = True
        rpt2.fldcodalmacen.Text = txtAlmacen.Text
        rpt2.fldnomalmacen.Text = pnlAlmacen.Caption
        rpt2.fldcodprod.Text = txtProducto.Text
        rpt2.fldnomprod.Text = pnlProducto.Caption
        rpt2.fldmedida.Text = txtmedida.Text
        rpt2.datconexion.ConnectionString = cnn_dbbancos

                csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, "
                csql = csql & "iif(IF4VALES.F4TIPDOC = '86', 'Guía',IF4VALES.F4TIPDOC) as F4TIPDOC,IF4VALES.NUMORDEN,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC) AS F4SERDOC,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC)AS F4NUMDOC, IF3VALES.F3VALVTA, IF4VALES.F4FECVAL, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,NULL) AS ENTRADAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,NULL) AS SALIDAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA * IF3VALES.F3CANPRO,NULL) AS ENTRADACOS, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA * IF3VALES.F3CANPRO,NULL) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI "
                csql = csql & "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) "
                csql = csql & "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) "
                csql = csql & "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI "
                csql = csql & "Where "
                
                ' Modificación para múltiples orígenes
                If Len(Trim(txtOrigen.Text)) > 0 Then

                    
                    origenArray = Split(Trim(txtOrigen.Text), "|")
                    
                    ' Si hay más de un origen, usar IN, de lo contrario mantener el formato original
                    If UBound(origenArray) > 0 Then
                        origenes = "('" & Join(origenArray, "','") & "')"
                        csql = csql & "(IF4VALES.F1CODORI IN " & origenes & ") And "
                    Else
                        csql = csql & "(IF4VALES.F1CODORI='" & Trim(txtOrigen.Text) & "') And "
                    End If
                End If
                
                csql = csql & "((IF4VALES.F4FECVAL) >=cvdate('" & aboDesde.Value & "') And (IF4VALES.F4FECVAL) <=cvdate('" & aboHasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') "
                csql = csql & "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL"
        rpt2.datconexion.ConnectionString = cnn_dbbancos
        rpt2.datconexion.Source = csql
        rpt2.Caption = App.ProductName & " - Kardex No Valorizado"
        rpt2.Show 'vbModal
        
        ElseIf opttipo(2).Value = True Then
        
            Dim rpt3 As New acr_kardex_resumen
    
            ccodprod = txtProducto.Text
            wcod_alm = txtAlmacen.Text
            
            rpt3.CodigoAlmacen = Trim(txtAlmacen.Text)
            rpt3.CodigoProducto = Trim(txtProducto.Text)
            rpt3.Fecha = Format(aboDesde.Value, "Short Date")
            rpt3.CodigoMoneda = IIf(CBool(optmoneda(0).Value), "S", "D")
            
            rpt3.fldempresa.Text = wnomcia
            
            If Len(pnlCcosto.Caption) > 0 And Len(txtCcosto.Text) > 0 Then
                cad2 = "C.C: " & pnlCcosto.Caption
            End If
            
            If Len(pnlOrigen.Caption) > 0 And Len(txtOrigen.Text) > 0 Then
                cad2 = cad2 & " - Origen: " & pnlOrigen.Caption
            End If
            
            rpt3.fldtitulor.Text = "Del  " & Format(aboDesde.Value, "dd/mm/yyyy") & "  al  " & Format(aboHasta.Value, "dd/mm/yyyy")
            rpt3.fldFecha.Text = Format(Date, "dd/mm/yyyy")
            rpt3.lblalmacen.Visible = True
            rpt3.fldcodalmacen.Visible = True
            rpt3.fldnomalmacen.Visible = True
            rpt3.fldcodalmacen.Text = txtAlmacen.Text
            rpt3.fldnomalmacen.Text = pnlAlmacen.Caption
            rpt3.datconexion.ConnectionString = cnn_dbbancos
        
        If optmoneda(0).Value = True Then
                csql = "SELECT IF4VALES.F4NUMVAL, IF4VALES.F4CENTRO, IF4VALES.F4TIPDOC, IF4VALES.NUMORDEN, IIf(IF4VALES.F4TIPDOC='86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC) AS F4SERDOC,IIf(IF4VALES.F4TIPDOC='86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC) AS F4NUMDOC, Sum(IF3VALES.F3VALVTA) AS SumaDeF3VALVTA, IF4VALES.F4FECVAL, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,Null)) AS ENTRADAK, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,Null)) AS SALIDAK, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,Null)) AS ENTRADACOS, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,Null)) AS SALIDACOS, IF3VALES.F2CODALM, SF1ORIGENES.F1NOMORI "
                csql = csql & "FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL =IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
                If Len(txtCcosto.Text) > 0 Then
                    csql = csql & "Where IF4VALES.F4CENTRO = '" & txtCcosto.Text & "' "
                End If
                
                csql = csql & "GROUP BY IF4VALES.F4NUMVAL, IF4VALES.F4CENTRO, IF4VALES.F4TIPDOC, IF4VALES.NUMORDEN, IIf(IF4VALES.F4TIPDOC='86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC), IIf(IF4VALES.F4TIPDOC='86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC), IF4VALES.F4FECVAL, IF3VALES.F2CODALM, SF1ORIGENES.F1NOMORI, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI "
                csql = csql & "HAVING (((IF4VALES.F4FECVAL)>=cvdate('" & aboDesde.Value & "') And (IF4VALES.F4FECVAL)<= cvdate('" & aboHasta.Value & "')) AND ((IF3VALES.F2CODALM)='" & wcod_alm & "') AND ((IF4VALES.F1CODORI) Not In ('XCS') And (IF4VALES.F1CODORI)='" & txtOrigen.Text & "')) "
                csql = csql & "ORDER BY IF4VALES.F4FECVAL, IF4VALES.F4NUMVAL;"

'            rpt3.LblIngreso.Caption = "S/"
'            rpt3.LblSalida.Caption = "S/"
        Else

                csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, "
                csql = csql & "IF4VALES.F4TIPDOC,IF4VALES.NUMORDEN,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4SERGUIA,IF4VALES.F4SERDOC) AS F4SERDOC,iif(IF4VALES.F4TIPDOC = '86',IF4VALES.F4NUMGUIA,IF4VALES.F4NUMDOC)AS F4NUMDOC, IF3VALES.F3VALDOL, IF4VALES.F4FECVAL, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALDOL * IF3VALES.F3CANPRO ,0) AS ENTRADACOS, "
                csql = csql & "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALDOL * IF3VALES.F3CANPRO,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI "
                csql = csql & "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) "
                csql = csql & "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) "
                csql = csql & "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI "
                csql = csql & "Where "
                
                ' Modificación para múltiples orígenes
                If Len(Trim(txtOrigen.Text)) > 0 Then

                    
                    origenArray = Split(Trim(txtOrigen.Text), "|")
                    
                    ' Si hay más de un origen, usar IN, de lo contrario mantener el formato original
                    If UBound(origenArray) > 0 Then
                        origenes = "('" & Join(origenArray, "','") & "')"
                        csql = csql & "(IF4VALES.F1CODORI IN " & origenes & ") And "
                    Else
                        csql = csql & "(IF4VALES.F1CODORI='" & Trim(txtOrigen.Text) & "') And "
                    End If
                End If
                
                csql = csql & "((IF4VALES.F4FECVAL) >=cvdate('" & aboDesde.Value & "') And (IF4VALES.F4FECVAL) <=cvdate('" & aboHasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') "
                csql = csql & "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL"

'            rpt3.LblIngreso.Caption = "US$"
'            rpt3.LblSalida.Caption = "US$"
        End If
        rpt3.datconexion.ConnectionString = cnn_dbbancos
        rpt3.datconexion.Source = csql
        rpt3.Caption = App.ProductName & " - Kardex Valorizado"
        rpt3.Show 'vbModal
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
End Sub

Private Sub optmoneda_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub opttipo_Click(index As Integer, Value As Integer)
    If index = 1 Then
        SSFrame3.Visible = False
    Else
        SSFrame3.Visible = True
    End If
End Sub

Private Sub opttipo_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If index = 1 Then
            cmdAceptar.SetFocus
        Else
            optmoneda(0).SetFocus
        End If
    End If
End Sub

Private Sub txtAlmacen_DblClick()
    txtAlmacen_KeyDown vbKeyF2, 0
End Sub

Private Sub txtalmacen_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtAlmacen
End Sub

Private Sub txtAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_almacen
                wcod_alm = vbNullString
                
                .Show 1
                
                If Trim(wcod_alm) <> vbNullString Then
                    txtAlmacen.Text = wcod_alm
                    pnlAlmacen.Caption = wnomalmacen
                End If
            End With
        Case vbKeyReturn
            txtAlmacen.Text = Format(txtAlmacen.Text, "00")
            pnlAlmacen.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T")
            
            If Trim(pnlAlmacen.Caption) = vbNullString Then
                txtAlmacen.SetFocus
            Else
                ModUtilitario.pulsarTecla vbKeyTab
            End If
    End Select
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtAlmacen_LostFocus()
    If Trim(txtAlmacen.Text) <> vbNullString And Trim(pnlAlmacen.Caption) = vbNullString Then
        txtAlmacen.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T")
        pnlAlmacen.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T")
    End If
End Sub

Private Sub txtccosto_DblClick()
txtccosto_KeyDown 113, 0
End Sub

Private Sub txtccosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
        sw_ayuda = True
        wcodcosto = "": wdescosto = ""
        Ayuda_Centros.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            txtCcosto.Text = wcodcosto
            pnlCcosto.Caption = wdescosto
'            txtccosto_KeyPress 13
        End If
    End If
End Sub

Private Sub txtOrigen_Change()
'    txtOrigen_KeyDown vbKeyF2, 0
End Sub

Private Sub txtOrigen_DblClick()
    txtOrigen_KeyDown vbKeyF2, 0
End Sub

Private Sub txtOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            
            With frmListaOrigen
'                objAyudaOrigen.inicializarEntidades
                
                .Ayuda = True
                
                objAyudaOrigen.inicializarEntidades
                
                .Show 1
                
'                If objAyudaOrigen.Codigo <> vbNullString Then
                    
                    txtOrigen.Text = objAyudaOrigen.Codigo
                    pnlOrigen.Caption = objAyudaOrigen.Descripcion
                        pnlOrigen.ToolTipText = objAyudaOrigen.Descripcion
                    
                    ModUtilitario.pulsarTecla vbKeyTab
'                End If
            End With
    End Select
End Sub

Private Sub txtProducto_DblClick()
    txtProducto_KeyDown vbKeyF2, 0
End Sub

Private Sub txtProducto_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtProducto
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                .Ayuda = True
                .InsumoOP = False
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                objAyudaBien.inicializarEntidades
                
                .Show 1
                
                If objAyudaBien.Codigo <> vbNullString Then
                    objAyudaBien.obtenerConfigBien
                    
                    txtProducto.Text = objAyudaBien.Codigo
                    pnlProducto.Caption = objAyudaBien.Descripcion
                        pnlProducto.ToolTipText = objAyudaBien.Descripcion
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
    End Select
End Sub

Private Sub txtProducto_LostFocus()
    If Trim(txtProducto.Text) <> vbNullString And Trim(pnlProducto.Caption) = vbNullString Then
        txtProducto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "IF5PLA", "F5CODPRO", Trim(txtProducto.Text), "T")
        pnlProducto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(txtProducto.Text), "T")
            pnlProducto.ToolTipText = pnlProducto.Caption
    End If
End Sub

