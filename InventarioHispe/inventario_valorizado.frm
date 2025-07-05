VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form inventario_valorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   6270
   ClientLeft      =   1455
   ClientTop       =   1935
   ClientWidth     =   6645
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6645
   Begin Threed.SSPanel SSPanel1 
      Height          =   6120
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   6540
      _Version        =   65536
      _ExtentX        =   11536
      _ExtentY        =   10795
      _StockProps     =   15
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
      BevelOuter      =   1
      Begin Threed.SSFrame Framoneda 
         Height          =   750
         Left            =   315
         TabIndex        =   12
         Top             =   1530
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
         Caption         =   " Moneda "
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
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption optsoles 
            Height          =   240
            Left            =   720
            TabIndex        =   1
            Top             =   300
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Soles"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optdolares 
            Height          =   240
            Left            =   3915
            TabIndex        =   2
            Top             =   300
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Dólares "
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
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   750
         Left            =   315
         TabIndex        =   13
         Top             =   780
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
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
         Begin VB.TextBox txtalmacen 
            Height          =   330
            Left            =   1035
            MaxLength       =   2
            TabIndex        =   0
            Top             =   270
            Width           =   465
         End
         Begin Threed.SSPanel pnlalmacen 
            Height          =   330
            Left            =   1575
            TabIndex        =   14
            Top             =   270
            Width           =   4200
            _Version        =   65536
            _ExtentX        =   7408
            _ExtentY        =   582
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   15
            Top             =   315
            Width           =   630
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   660
         Left            =   315
         TabIndex        =   16
         Top             =   120
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1164
         _StockProps     =   14
         Caption         =   " Fecha "
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
         Alignment       =   2
         Font3D          =   3
         Begin MSComCtl2.DTPicker abohasta 
            Height          =   315
            Left            =   2400
            TabIndex        =   29
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   111214593
            CurrentDate     =   40611
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   210
            Left            =   1785
            TabIndex        =   17
            Top             =   360
            Width           =   420
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   750
         Left            =   315
         TabIndex        =   18
         Top             =   2280
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
         Caption         =   " Tipo "
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
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   0
            Left            =   720
            TabIndex        =   3
            Top             =   300
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Valorizado"
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
         End
         Begin Threed.SSOption opttipo 
            Height          =   240
            Index           =   1
            Left            =   3915
            TabIndex        =   4
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
      Begin Threed.SSFrame SSFrame5 
         Height          =   750
         Left            =   315
         TabIndex        =   20
         Top             =   3600
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
         Caption         =   " Impresión "
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
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption optimpresion 
            Height          =   240
            Index           =   0
            Left            =   765
            TabIndex        =   7
            Top             =   300
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Detallado"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optimpresion 
            Height          =   240
            Index           =   1
            Left            =   3915
            TabIndex        =   8
            Top             =   300
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Resumido"
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
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   750
         Left            =   315
         TabIndex        =   21
         Top             =   4320
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
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
         Begin VB.TextBox txtmarca 
            Height          =   330
            Left            =   1035
            MaxLength       =   3
            TabIndex        =   9
            Top             =   270
            Width           =   465
         End
         Begin Threed.SSPanel pnlmarca 
            Height          =   330
            Left            =   1575
            TabIndex        =   22
            Top             =   270
            Width           =   4200
            _Version        =   65536
            _ExtentX        =   7408
            _ExtentY        =   582
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
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   23
            Top             =   315
            Width           =   450
         End
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   750
         Left            =   315
         TabIndex        =   24
         Top             =   5040
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
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
         Begin VB.TextBox txtlinea 
            Height          =   330
            Left            =   1035
            MaxLength       =   3
            TabIndex        =   10
            Top             =   270
            Width           =   465
         End
         Begin Threed.SSPanel pnllinea 
            Height          =   330
            Left            =   1575
            TabIndex        =   25
            Top             =   270
            Width           =   4200
            _Version        =   65536
            _ExtentX        =   7408
            _ExtentY        =   582
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
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Línea"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   26
            Top             =   315
            Width           =   390
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   750
         Left            =   315
         TabIndex        =   19
         Top             =   3000
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   1323
         _StockProps     =   14
         Caption         =   " Orden "
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
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption optorden 
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Por Línea"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption optorden 
            Height          =   240
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Por Marca"
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
         End
         Begin Threed.SSOption optorden 
            Height          =   240
            Index           =   2
            Left            =   2760
            TabIndex        =   27
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Por código"
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
         End
         Begin Threed.SSOption optorden 
            Height          =   240
            Index           =   3
            Left            =   4080
            TabIndex        =   28
            Top             =   300
            Width           =   1650
            _Version        =   65536
            _ExtentX        =   2910
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Por cod. fabricante"
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
         End
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "inventario_valorizado.frx":0000
      ToolBars        =   "inventario_valorizado.frx":1978
   End
End
Attribute VB_Name = "inventario_valorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nniveles            As Integer
Dim nlonnivel01         As Integer
Dim nlonnivel02         As Integer
Dim nlonnivel03         As Integer
Dim nlonnivel04         As Integer
Dim nlonnivel05         As Integer
Dim cnombase            As String
Dim cnomtabla           As String
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection
Dim nmes                As Integer
Dim sw_ayuda            As Boolean
Dim sw_ayuda_linea      As Boolean

Private Sub abohasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAlmacen.SetFocus
    End If
    
End Sub

Private Sub cmdaceptar_Click()

    
End Sub

Private Sub cmdsalir_Click()

End Sub

Private Sub Form_Load()
    Dim CadSql      As String

    Me.MousePointer = vbHourglass
    Me.top = 1065
    Me.left = 1605
    cnombase = "TEMPLUS.MDB"
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    

    aboHasta.Value = Format(Date, "DD/MM/YYYY")
    
    'si hay un solo almacen lo pone de una vez
    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    rsalmacen.Open "SELECT * FROM EF2ALMACENES", cnn_dbbancos
    rsalmacen.MoveFirst
        txtAlmacen.Text = Trim(rsalmacen.Fields("F2CODALM") & "")
        pnlalmacen.Caption = Trim(rsalmacen.Fields("F2NOMALM") & "")
    rsalmacen.MoveNext
    If Not rsalmacen.EOF Then
        txtAlmacen.Text = ""
        pnlalmacen.Caption = ""
    End If
    rsalmacen.Close
    Me.MousePointer = vbDefault
    
End Sub

Private Sub BUSCA_NIVELES()
Dim I   As Integer

    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM SF1PARAIN WHERE F1CODEMP ='" & wempresa & "'", cnn_control
    If Not rscontrol.EOF Then
        nniveles = IIf(Val("" & rscontrol.Fields("F1niveles")) > 5, 5, Val("" & rscontrol.Fields("F1niveles")))
        
        nlonnivel01 = Val("" & rscontrol.Fields("F1LONNIV1"))
        nlonnivel02 = Val("" & rscontrol.Fields("F1LONNIV2"))
        nlonnivel03 = Val("" & rscontrol.Fields("F1LONNIV3"))
        nlonnivel04 = Val("" & rscontrol.Fields("F1LONNIV4"))
        nlonnivel05 = Val("" & rscontrol.Fields("F1LONNIV5"))
        
        If nniveles > 0 Then
            For I = 1 To nniveles
                'chknivel(I - 1).Visible = True
                'chknivel(I - 1).Caption = " " & Left(rscontrol.Fields("F1NIVEL0" & Format(I, "0")), 10)
            Next
        End If
    End If
    rscontrol.Close
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_form.Close
    
End Sub

Private Sub optsoles_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        opttipo(1).SetFocus
    End If

End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        optorden(1).SetFocus
    End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo Error_Reporte
    
    Dim sql         As String
    Dim csqlmon     As String
    Dim calmacen    As String
    Dim dfhasta     As Date
    Dim cdescmoneda As String
    Dim corden      As String
    Dim cwherealma  As String
    Dim cbasetemp   As String
    Dim ccrea       As String
    Dim chaving     As String
    Dim cwheremarca As String
    Dim cwherelinea As String
    Dim cformat     As String
    Dim cwhere_formato  As String
    Dim nivel01     As String
    Dim cwherenulos As String
    Dim cwherenulos2 As String
    
    Select Case Tool.ID
        Case "ID_Procesar"
            Me.MousePointer = vbHourglass
            
            BUSCA_NIVELES
            
            cformat = ""
            If wf1decimal_cantidad > 0 Then
                cformat = String(wf1decimal_cantidad, "0")
            End If
                        
            calmacen = Trim(txtAlmacen.Text)
            cwherealma = ""
            cwheremarca = ""
            cwherenulos = " Val(IIf(IsNull([TBEXISTENCIAS].[CANTIDAD]),0,[TBEXISTENCIAS].[CANTIDAD]))>0 "
            If Len(Trim(calmacen)) > 0 Then
                cwherealma = "And ((EF2ALMACENES.f2codalm) = '" & calmacen & "') "
            Else
                cwherealma = ""
            End If
            dfhasta = Format(aboHasta.Value, "DD/MM/YYYY")
            csqlmon = ""
            If optdolares.Value = True Then
                cdescmoneda = "(EXPRESADO EN DOLARES AMERICANOS)"
                csqlmon = "Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD,Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO,(IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, IIF(CANTIDAD>0,[VALOR_VENTA]/[CANTIDAD],0) AS COSTO_UNITARIO, "
            Else
                cdescmoneda = "(EXPRESADO EN NUEVOS SOLES)"
                csqlmon = "Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD,Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, IIF(CANTIDAD>0,[VALOR_VENTA]/[CANTIDAD],0) AS COSTO_UNITARIO, "
            End If
            If optorden(0).Value = True Then
            Else
                corden = "ORDER BY First(EF2MARCAS.F2DESMAR),First(IF5PLA.F5CODPRO)"
            End If
            
            If optimpresion(1).Value = True Then
                cbasetemp = wrutatemp & "\TEMPLUS.MDB"
                ccrea = "INTO TEMP_INV IN '" & cbasetemp & "' "
                
                sql = "DROP TABLE TEMP_INV"
                cnn_form.Execute sql
                'AlmacenaQuery_sql sql, cnn_form
            End If
            
            If Len(Trim(txtMarca.Text)) > 0 Then
                cwheremarca = " AND EF2MARCAS.F2DESMAR='" & Trim(pnlmarca.Caption) & "' "
            Else
                cwheremarca = ""
            End If
            
            sql = "SELECT IF3VALES.F5CODPRO AS CODPRO, "
            sql = sql & csqlmon
            sql = sql & "First(IF5PLA.F5NOMPRO) AS NOMPRO, First(IF5PLA.F5FACTOR) AS F5FACTOR, "
            sql = sql & "First(IF5PLA.F5CODFAB) AS CODFAB, "
            sql = sql & "First(IF5PLA.F5PREVTA) AS F5PREVTA, "
            sql = sql & "First(IF5PLA.F7CODMED) AS CODMED, "
            sql = sql & "First(EF2MARCAS.F2DESMAR) AS DESMAR, "
            sql = sql & "First(SF7NIVEL01.F7DESCON) AS NIVEL01 "
            sql = sql & ccrea
            sql = sql & "FROM (((IF5PLA INNER JOIN (IF4VALES INNER JOIN IF3VALES ON "
            sql = sql & "(IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND "
            sql = sql & "(IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON "
            sql = sql & "IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) "
            sql = sql & "INNER JOIN SF7NIVEL01 ON LEFT(IF5PLA.F5CODPRO,2) = SF7NIVEL01.F7CODCON) "
            sql = sql & "LEFT JOIN EF2MARCAS ON "
            sql = sql & "IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR) INNER JOIN EF2ALMACENES ON "
            sql = sql & "IF4VALES.F2CODALM = EF2ALMACENES.F2CODALM "
            sql = sql & "WHERE ((IF4VALES.F4FECVAL) <= cvdate('" & dfhasta & "')) "
            sql = sql & cwherealma
            sql = sql & cwheremarca
            sql = sql & " AND IF5PLA.F5DESCONTINUADO='N' "
            sql = sql & "GROUP BY IF3VALES.F5CODPRO "
            sql = sql & corden
            
            If optimpresion(0).Value = True Then
                If opttipo(0).Value = True Then
                    With acr_inventario_valorizado
                        Screen.MousePointer = vbHourglass
                        .Caption = wnomcia
                        If ctipoadm_bd = "M" Then
                            .datos.ConnectionString = cnn_form
                        Else
                            .datos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cnn_dbbancos & ""
                        End If
                        .lblmoneda.Caption = cdescmoneda
                        .lblfecha.Caption = "Al " & Format(aboHasta.Value, "DD/MM/YYYY")
                        .datos.Source = sql
    
                        .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                        .lblEmpresa.Caption = wnomcia
                        If pnlalmacen.Caption = "" Then
                            .LabelAlmacen.Caption = "(TODOS LOS ALMACENES)"
                        Else
                            .LabelAlmacen.Caption = pnlalmacen.Caption
                        End If
                        
                        If optorden(0).Value = True Then
                            .fldquiebre.DataField = "NIVEL01"
                            .GroupHeader1.DataField = "NIVEL01"
                        Else
                            .fldquiebre.DataField = "DESMAR"
                            .GroupHeader1.DataField = "DESMAR"
                        End If
                        Screen.MousePointer = vbDefault
                        .Show vbModal
                    End With
                Else
                    With acr_inventario_novalorizado
                        cwhere_formato = ""
                        If wf1formato_inventario = "1" Then
                            '.lblfactor.Visible = False
                            '.fldfactor.Visible = False
                            cwhere_formato = " TBEXISTENCIAS.cantidad > 0 "
                        Else
                            '.lblfactor.Visible = False
                            '.fldfactor.Visible = False
                        End If
                        
    '                    If wf1decimal_cantidad > 0 Then
    '                        .fldsubcantidad.OutputFormat = "#,##0" & "." & cformat
    '                        .fldsaldo.OutputFormat = "#,##0" & "." & cformat
    '                    Else
    '                        .fldsubcantidad.OutputFormat = "#,##0"
    '                        .fldsaldo.OutputFormat = "#,##0"
    '                    End If
                        '---------------------------------------------------------------------
                        cbasetemp = wrutatemp & "\TEMPLUS.MDB"
                        ccrea = "INTO TBEXISTENCIAS IN '" & cbasetemp & "' "
                        
                        sql = "DROP TABLE TBEXISTENCIAS"
                        cnn_form.Execute sql
                        'AlmacenaQuery_sql sql, cnn_form
                                                        
                        chaving = "Having ((if3vales.f2codalm) = '" & calmacen & "') "
                        
                        If Len(Trim(txtMarca.Text)) > 0 Then
                            cwheremarca = "WHERE TBPRODUCTOS.DESMAR='" & Trim(pnlmarca.Caption) & "' "
                        Else
                            cwheremarca = ""
                        End If
                                                        
                        cwherealma = ""
                        
                        If Len(Trim(calmacen)) > 0 Then
                            cwherealma = "AND ((IF3VALES.F2CODALM) = '" & calmacen & "') "
                        Else
                            cwherealma = ""
                        End If
                        
                        sql = "SELECT if3vales.F2CODALM, if3vales.F5CODPRO, Sum(IIf(Left([IF3VALES].[f4numval],1)='I'," & _
                              "[IF3VALES].[f3canpro])) AS ing, Sum(IIf(Left([IF3VALES].[f4numval],1)='S'," & _
                              "[IF3VALES].[f3canpro])) AS egr, Format(IIf(IsNull(ing),0,ing)-IIf(IsNull(egr),0,egr),'0.00') AS CANTIDAD " & _
                              ccrea & _
                              "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND " & _
                              "(IF4VALES.F2CODALM = if3vales.F2CODALM) " & _
                              "WHERE (((IF4VALES.F4FECVAL) <= CVDATE('" & dfhasta & "'))) " & cwherealma & _
                              "GROUP BY if3vales.F2CODALM, if3vales.F5CODPRO " & _
                              "ORDER BY if3vales.F2CODALM, if3vales.F5CODPRO;"
                        
                        If ctipoadm_bd = "M" Then
                            cnn_form.Execute (sql)
                             'AlmacenaQuery_sql sql, cnn_form
                        Else
                            cnn_dbbancos.Execute (sql)
                             'AlmacenaQuery_sql sql, cnn_form
                        End If
                        
                        '---------------------------------------------------------------------
                        ccrea = "INTO TBPRODUCTOS IN '" & cbasetemp & "' "
                        
                        sql = "DROP TABLE TBPRODUCTOS"
                        cnn_form.Execute sql
                         'AlmacenaQuery_sql sql, cnn_form
                         
                        If nlonnivel01 = 2 Then
                            nivel01 = "LEFT(IF5PLA.F5CODPRO,2)"
                        Else
                            nivel01 = "LEFT(IF5PLA.F5CODPRO,3)"
                        End If
                        
                        sql = "SELECT IF5PLA.F5CODPRO AS CODPRO, IF5PLA.F5NOMPRO AS NOMPRO, IF5PLA.F5FACTOR AS F5FACTOR, " & _
                              "IF5PLA.F5CODFAB AS CODFAB, EF7MEDIDAS.F7SIGMED as F7CODMED, IF5PLA.F5PREVTA, EF2MARCAS.F2DESMAR AS DESMAR, " & _
                              "SF7NIVEL01.F7DESCON AS NIVEL01,IF5PLA.F5FOB AS COSTO_UNITARIO " & _
                              ccrea & _
                              "FROM ((IF5PLA LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR) LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) " & _
                              "LEFT JOIN SF7NIVEL01 ON " & _
                              nivel01 & _
                              " = SF7NIVEL01.F7CODCON " & _
                              "WHERE IF5PLA.F5TIPO='P' AND F5DESCONTINUADO='N'"
                        If ctipoadm_bd = "M" Then
                            cnn_form.Execute (sql)
                             'AlmacenaQuery_sql sql, cnn_form
                        Else
                            cnn_dbbancos.Execute sql
                            'AlmacenaQuery_sql sql, cnn_dbbancos
                        End If
                        '---------------------------------------------------------------------
                       
                            If optorden(0).Value = True Then corden = "ORDER BY TBPRODUCTOS.NIVEL01,TBPRODUCTOS.CODPRO"
                            If optorden(1).Value = True Then corden = "ORDER BY TBPRODUCTOS.DESMAR, RIGHT(TBPRODUCTOS.NOMPRO,4)"
                            If optorden(2).Value = True Then corden = "ORDER BY TBPRODUCTOS.CODPRO"
                            If optorden(3).Value = True Then corden = "ORDER BY TBPRODUCTOS.CODFAB"
                        If Len(Trim(txtlinea.Text)) > 0 Then
                            If Len(Trim(cwheremarca)) > 0 Then
                                cwherelinea = " AND TBPRODUCTOS.NIVEL01='" & pnllinea.Caption & "' "
                            Else
                                cwherelinea = "WHERE TBPRODUCTOS.NIVEL01='" & pnllinea.Caption & "' "
                            End If
                        Else
                            cwherelinea = ""
                        End If
        
                        If (Len(cwheremarca) > 0 Or Len(cwherelinea) > 0) And Len(cwhere_formato) > 0 Then
                            cwhere_formato = " AND " & cwhere_formato
                        Else
                            If (Len(cwheremarca) = 0 Or Len(cwherelinea) = 0) And Len(cwhere_formato) > 0 Then
                                cwhere_formato = "WHERE " & cwhere_formato
                            End If
                        End If
                        
                        If Len(cwheremarca) > 0 Or Len(cwherelinea) > 0 Or Len(cwhere_formato) > 0 Then
                            cwherenulos = " AND " & cwherenulos
                        Else
                            cwherenulos = "WHERE " & cwherenulos
                        End If
                            
                        sql = vbNullString
                        sql = sql & "SELECT DISTINCTROW "
                        sql = sql & "TBPRODUCTOS.CODPRO, "
                        sql = sql & "LEFT(TBPRODUCTOS.NOMPRO,70) AS NOMPRO, "
                        sql = sql & "TBPRODUCTOS.F5FACTOR, "
                        sql = sql & "TBPRODUCTOS.DESMAR, "
                        sql = sql & "TBPRODUCTOS.F7CODMED, "
                        sql = sql & "TBPRODUCTOS.CODFAB, "
                        sql = sql & "TBPRODUCTOS.F5PREVTA, "
                        sql = sql & "TBPRODUCTOS.NIVEL01, "
                        sql = sql & "TBEXISTENCIAS.ING, "
                        sql = sql & "TBEXISTENCIAS.EGR, "
                        sql = sql & "IIF(ISNULL(TBEXISTENCIAS.CANTIDAD),0,TBEXISTENCIAS.CANTIDAD) AS CANTIDAD, "
                        sql = sql & "TBPRODUCTOS.COSTO_UNITARIO "
                        sql = sql & "FROM "
                        sql = sql & "TBPRODUCTOS "
                        sql = sql & "LEFT JOIN TBEXISTENCIAS ON TBPRODUCTOS.CODPRO = TBEXISTENCIAS.F5CODPRO "
                        sql = sql & cwheremarca & cwherelinea & cwhere_formato & cwherenulos & corden
                              
                        Screen.MousePointer = vbHourglass
                        .Caption = wnomcia
                        .datos.ConnectionString = cnn_form
                        .lblfecha.Caption = "Al " & Format(aboHasta.Value, "DD/MM/YYYY")
                        .datos.Source = sql
                        .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                        .Fldhora.Text = Time
                        .lblEmpresa.Caption = wnomcia
                        If pnlalmacen.Caption = "" Then
                            .LabelAlmacen.Caption = "(TODOS LOS ALMACENES)"
                        Else
                            .LabelAlmacen.Caption = pnlalmacen.Caption
                        End If
                        If optorden(0).Value = True Then
                            .fldquiebre.DataField = "NIVEL01"
                            .GroupHeader1.DataField = "NIVEL01"
                        ElseIf optorden(1).Value = True Then
                            .fldquiebre.DataField = "DESMAR"
                            .GroupHeader1.DataField = "DESMAR"
                        End If
                        Screen.MousePointer = vbDefault
                        .Show vbModal
                    End With
                End If
            Else
                If ctipoadm_bd = "M" Then
                    cnn_form.Execute (sql)
                     'AlmacenaQuery_sql sql, cnn_form
                Else
                    cnn_dbbancos.Execute (sql)
                    AlmacenaQuery_sql sql, cnn_dbbancos
                    
                End If
                If opttipo(0).Value = True Then
                    If optorden(0).Value = True Then
                        sql = "SELECT DISTINCTROW TEMP_INV.NIVEL01 AS DESCRIPCION, Sum(TEMP_INV.VALOR_VENTA) AS VALOR " & _
                              "FROM TEMP_INV " & _
                              "GROUP BY TEMP_INV.NIVEL01 " & _
                              "ORDER BY TEMP_INV.NIVEL01;"
                    Else
                        sql = "SELECT DISTINCTROW TEMP_INV.DESMAR AS DESCRIPCION, Sum(TEMP_INV.VALOR_VENTA) AS VALOR " & _
                              "FROM TEMP_INV " & _
                              "GROUP BY TEMP_INV.DESMAR " & _
                              "ORDER BY TEMP_INV.DESMAR;"
                    End If
                Else
                    If optorden(0).Value = True Then
                        sql = "SELECT DISTINCTROW TEMP_INV.NIVEL01 AS DESCRIPCION, Sum(TEMP_INV.CANTIDAD) AS VALOR " & _
                              "FROM TEMP_INV " & _
                              "GROUP BY TEMP_INV.NIVEL01 " & _
                              "ORDER BY TEMP_INV.NIVEL01;"
                    Else
                        sql = "SELECT DISTINCTROW TEMP_INV.DESMAR AS DESCRIPCION, Sum(TEMP_INV.CANTIDAD) AS VALOR " & _
                              "FROM TEMP_INV " & _
                              "GROUP BY TEMP_INV.DESMAR " & _
                              "ORDER BY TEMP_INV.DESMAR;"
                    End If
                End If
                
                With acr_resumen_inventario_valorizado
                    Screen.MousePointer = vbHourglass
                    .Caption = wnomcia
                    .datos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cnn_form & ""
                    .lblmoneda.Caption = cdescmoneda
                    .lblfecha.Caption = "Al " & Format(aboHasta.Value, "DD/MM/YYYY")
                    .datos.Source = sql
                    .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                    .lblEmpresa.Caption = wnomcia
                    If pnlalmacen.Caption = "" Then
                        .LabelAlmacen.Caption = "(TODOS LOS ALMACENES)"
                    Else
                        .LabelAlmacen.Caption = pnlalmacen.Caption
                    End If
                    Screen.MousePointer = vbDefault
                    .Show vbModal
                End With
                
            End If
            
            Me.MousePointer = vbDefault
        Case "ID_Salir"
            Unload Me
    End Select
Exit Sub
    
Error_Reporte:
    If Err.Number = -2147217865 Then
        Resume Next
    Else
        MsgBox "ERROR N° " & Err.Number & Space(2) & Err.Description, 16, "AVISO"
        Resume Next
    End If
    Exit Sub
    
End Sub

Private Sub txtalmacen_Change()
    
    pnlalmacen.Caption = ""
    
End Sub

Private Sub txtAlmacen_DblClick()

    txtAlmacen_KeyDown 113, 0
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtAlmacen.SelStart = 0
    txtAlmacen.SelLength = Len(txtAlmacen.Text)

End Sub

Private Sub txtAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            txtAlmacen.Text = wcod_alm
            pnlalmacen.Caption = wnomalmacen
            txtAlmacen_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        optdolares.SetFocus
    End If

End Sub

Private Sub txtAlmacen_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtAlmacen.Text)) > 0 Then
            wnomalmacen = ""
            If VALIDA_ALMACEN(txtAlmacen.Text) = True Then
                pnlalmacen.Caption = wnomalmacen
            Else
                MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                txtAlmacen.SetFocus
            End If
        Else
            pnlalmacen.Caption = "TODOS LOS ALMACENES"
        End If
    End If

End Sub

Private Sub txtlinea_DblClick()
    
    txtlinea_KeyDown 113, 0
    
End Sub

Private Sub txtlinea_GotFocus()
    
    txtlinea.SelStart = 0: txtlinea.SelLength = Len(txtlinea.Text)
    
End Sub

Private Sub txtlinea_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        wcodlinea = ""
        sw_ayuda_linea = True
''''        ayuda_nivel01.Show 1
        sw_ayuda_linea = False
        If Len(Trim(wcodlinea)) > 0 Then
            txtlinea.Text = wcodlinea
            pnllinea.Caption = wnomlinea
        End If
    End If

End Sub

Private Sub txtlinea_KeyPress(KeyAscii As Integer)
Dim rsfield     As New ADODB.Recordset

If KeyAscii = 13 Then

    If sw_ayuda_linea = False Then
        If Len(Trim(txtlinea.Text)) > 0 Then
            sql = "SELECT F7DESCON FROM SF7NIVEL01 WHERE F7CODCON = '" & txtlinea.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rsfield.EOF Then
                pnllinea.Caption = Trim("" & rsfield.Fields("F7DESCON"))
            Else
                txtlinea.Text = ""
                txtlinea.SetFocus
                pnllinea.Caption = ""
                MsgBox "Código de línea no existe. Verifique.", vbInformation, "Atención"
                
            End If
            rsfield.Close
            Set rsfield = Nothing
        Else
            pnllinea.Caption = ""
        End If
    End If
End If

End Sub

Private Sub txtlinea_LostFocus()
'txtlin
End Sub

Private Sub txtmarca_DblClick()

    txtmarca_KeyDown 113, 0

End Sub

Private Sub txtmarca_GotFocus()

    txtMarca.SelStart = 0: txtMarca.SelLength = Len(txtMarca.Text)

End Sub

Private Sub txtmarca_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodmar = ""
        sw_ayuda_marca = ""
        ayuda_marcas.Show 1
        sw_ayuda = False
        If Len(Trim(wcodmar)) > 0 Then
            txtMarca.Text = wcodmar
            txtmarca_KeyPress 13
        End If
    End If

End Sub

Private Sub txtmarca_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtlinea.SetFocus
    End If

End Sub

Private Sub txtmarca_LostFocus()
Dim rsfield     As New ADODB.Recordset

    If sw_ayuda = False Then
        If Len(Trim(txtMarca.Text)) > 0 Then
            sql = "select F2DESMAR from EF2MARCAS where F2CODMAR = '" & txtMarca.Text & "'"
            If rsfield.State = 1 Then rsfield.Close
            rsfield.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rsfield.EOF Then
                pnlmarca.Caption = Trim("" & rsfield.Fields("F2DESMAR"))
            Else
                txtMarca.Text = ""
                pnlmarca.Caption = ""
                MsgBox "Código de marca no existe. Verifique.", vbInformation, "Atención"
                txtMarca.SetFocus
            End If
            rsfield.Close
            Set rsfield = Nothing
        Else
            pnlmarca.Caption = ""
        End If
    End If
        
End Sub
