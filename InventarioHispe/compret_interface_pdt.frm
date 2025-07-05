VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form compret_interface_pdt 
   Caption         =   "Generar interface para el PDT"
   ClientHeight    =   2880
   ClientLeft      =   4155
   ClientTop       =   2100
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4230
   Begin Threed.SSPanel SSPanel1 
      Height          =   1995
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4020
      _Version        =   65536
      _ExtentX        =   7091
      _ExtentY        =   3519
      _StockProps     =   15
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
      Begin VB.TextBox txtanno 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   2
         Top             =   630
         Width           =   690
      End
      Begin VB.ComboBox cmbmes 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1035
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   495
         TabIndex        =   6
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label lblcompro 
         Caption         =   "Nº Comprobante "
         Height          =   195
         Left            =   450
         TabIndex        =   5
         Top             =   1620
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   210
         Left            =   495
         TabIndex        =   4
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Comprobantes de Retención"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   90
         Width           =   3315
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   2295
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   2115
      TabIndex        =   8
      Top             =   2295
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
Attribute VB_Name = "compret_interface_pdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsretcab        As New ADODB.Recordset
Dim rsretdet        As New ADODB.Recordset
Dim rsregisofi      As New ADODB.Recordset
Dim cproame         As String

Private Sub GENERA_RET_INTERFACE()
Dim csql            As String
Dim Csql2           As String
Dim nmes            As Integer
Dim ccadena         As String
Dim crazon          As String
Dim capepat         As String
Dim capemat         As String
Dim cnombres        As String
Dim ntotal          As Double
Dim cmes            As String
Dim cnummov         As String
Dim ccomp           As String
Dim cfile           As String
    
    nmes = Right(cmbmes.Text, 2)
    cproame = Format(txtanno.Text, "0000") & Format(nmes, "00")
    
    Open Trim(wrutabancos) & Trim(wusuario) & ".TXT" For Output As #1
    
    wfile = "0626" & wrucempresa & cproame & ".TXT"
    'Open Trim(wrutabancos) & cfile For Output As #1
    
    csql = "SELECT * FROM RETENDOC WHERE MONTH(FECHA)=" & nmes & " ORDER BY FECHA"
    If rsretcab.State = adStateOpen Then rsretcab.Close
    rsretcab.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsretcab.EOF Then
    
        rsretcab.MoveFirst
        Do While Not rsretcab.EOF
            If rsretcab.Fields("ANULADO") & "" <> "S" Then
                lblcompro.Caption = "Nº Comprobante " & rsretcab.Fields("SERIE") & "/" & rsretcab.Fields("NUM_DOCUMENTO")
                lblcompro.Refresh
                
                'If rsretcab.Fields("SERIE") = "001" And rsretcab.Fields("NUM_DOCUMENTO") = "0000099" Then
                '    SS$ = 2
                'End If
                
                Csql2 = "SELECT * FROM RETENMOV WHERE SERIE_D='" & rsretcab.Fields("SERIE") & "' AND NUM_DOCUMENTOS='" & rsretcab.Fields("NUM_DOCUMENTO") & "' ORDER BY FECHA_EMISION"
                If rsretdet.State = adStateOpen Then rsretdet.Close
                rsretdet.Open Csql2, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsretdet.EOF Then
                    rsretdet.MoveFirst
                    Do While Not rsretdet.EOF
                        ccadena = ""
                        ccadena = Format(rsretcab.Fields("RUC") & "", "00000000000")
                        
                        crazon = "": capepat = "": capemat = "": cnombres = ""
                        If rsproveedor.State = adStateOpen Then rsproveedor.Close
                        rsproveedor.Open "SELECT F2APEPAT,F2APEMAT,F2NOMBRES FROM EF2PROVEEDORES WHERE F2NEWRUC='" & rsretcab.Fields("RUC") & "' AND F2TIPO_PERSONA='N'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsproveedor.EOF Then
                            capepat = Trim(Left(rsproveedor.Fields("F2APEPAT") & "", 20))
                            capemat = Trim(Left(rsproveedor.Fields("F2APEMAT") & "", 20))
                            cnombres = Trim(Left(rsproveedor.Fields("F2NOMBRES") & "", 20))
                        Else
                            crazon = Trim(Left(rsretcab.Fields("NOMBRE"), 40))
                        End If
                        rsproveedor.Close
                        ccadena = ccadena & "|" & crazon & "|" & capepat & "|" & capemat & "|" & cnombres & "|"
                        ccadena = ccadena & Format(rsretcab.Fields("SERIE") & "", "0000") & "|"
                        ccadena = ccadena & Format(rsretcab.Fields("NUM_DOCUMENTO") & "", "00000000") & "|"
                        ccadena = ccadena & Format(rsretcab.Fields("FECHA") & "", "DD/MM/YYYY") & "|"
                        Rem NSE ccadena = ccadena & Format(rsretdet.Fields("MONTO_PAGO") & "", "0.00") & "|"
                        ccadena = ccadena & Format(rsretcab.Fields("BASE") & "", "0.00") & "|"
                        ccadena = ccadena & Format(rsretdet.Fields("TIPO") & "", "00") & "|"
                        ccadena = ccadena & Format(rsretdet.Fields("SERIE") & "", "0000") & "|"
                        ccadena = ccadena & Format(rsretdet.Fields("NUMERO_CORRELA") & "", "00000000") & "|"
                        ccadena = ccadena & Format(rsretdet.Fields("FECHA_EMISION") & "", "DD/MM/YYYY") & "|"
                        
                        ntotal = 0#
                        cmes = Mid(rsretdet.Fields("REGCOMP") & "", 1, 2)
                        cnummov = Mid(rsretdet.Fields("REGCOMP") & "", 3, 7)
                        ccomp = "SELECT F4TOTAL FROM REGISOFI WHERE F4MESMOV='" & cmes & "' AND F4NUMMOV='" & cnummov & "'"
                        If rsregisofi.State = adStateOpen Then rsregisofi.Close
                        rsregisofi.Open ccomp, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsregisofi.EOF Then
                            ntotal = Val(rsregisofi.Fields("F4TOTAL") & "")
                        End If
                        rsregisofi.Close
                        
                        ccadena = ccadena & Format(ntotal, "0.00") & "|"
                        
                        Print #1, ccadena
                        
                        rsretdet.MoveNext
                    Loop
                End If
                rsretdet.Close
            End If
            
            rsretcab.MoveNext
        Loop
        
        'MsgBox "Fin del proceso.", vbInformation, "Atención"
        
    Else
        'MsgBox "No existen movimientos en el mes para ser procesados.", vbInformation, "Atención"
    End If
    
    Close #1
    
    
    
    rsretcab.Close

End Sub

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            GENERA_RET_INTERFACE
            frmView.Show 1
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

    txtanno.Text = wanno
    
    cmbmes.AddItem "Enero     " & Space(80) & "01"
    cmbmes.AddItem "Febrero   " & Space(80) & "02"
    cmbmes.AddItem "Marzo     " & Space(80) & "03"
    cmbmes.AddItem "Abril     " & Space(80) & "04"
    cmbmes.AddItem "Mayo      " & Space(80) & "05"
    cmbmes.AddItem "Junio     " & Space(80) & "06"
    cmbmes.AddItem "Julio     " & Space(80) & "07"
    cmbmes.AddItem "Agosto    " & Space(80) & "08"
    cmbmes.AddItem "Setiembre " & Space(80) & "09"
    cmbmes.AddItem "Octubre   " & Space(80) & "10"
    cmbmes.AddItem "Noviembre " & Space(80) & "11"
    cmbmes.AddItem "Diciembre " & Space(80) & "12"
    
    cmbmes.ListIndex = Month(Date) - 1
    
End Sub
