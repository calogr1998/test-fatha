VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form generar_ret 
   Caption         =   "Generación de Asientos Contables"
   ClientHeight    =   2790
   ClientLeft      =   3795
   ClientTop       =   2205
   ClientWidth     =   4200
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
   ScaleHeight     =   2790
   ScaleWidth      =   4200
   Begin Threed.SSPanel SSPanel1 
      Height          =   1995
      Left            =   90
      TabIndex        =   4
      Top             =   90
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
      Begin VB.ComboBox cmbmes 
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1035
         Width           =   2445
      End
      Begin VB.TextBox txtanno 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   0
         Top             =   630
         Width           =   690
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
         TabIndex        =   8
         Top             =   90
         Width           =   3315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   210
         Left            =   495
         TabIndex        =   7
         Top             =   675
         Width           =   300
      End
      Begin VB.Label lblcompro 
         Caption         =   "Nº Comprobante "
         Height          =   195
         Left            =   450
         TabIndex        =   6
         Top             =   1620
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   495
         TabIndex        =   5
         Top             =   1080
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   765
      TabIndex        =   2
      Top             =   2250
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2250
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
Attribute VB_Name = "generar_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsretcab        As New ADODB.Recordset
Dim rsretdet        As New ADODB.Recordset
Dim cconex_form     As String
Dim cnn_form        As New ADODB.Connection
Dim rscontable      As New ADODB.Recordset
Dim cf1ctaretencion As String
Dim cf1origen_ret   As String
Dim nelemen         As Integer
Dim ncompro         As Integer
Dim cproame         As String
Dim ntc             As Double
Dim rsregisdoc      As New ADODB.Recordset

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            GENERA_RET
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub GENERA_RET()
Dim csql            As String
Dim Csql2           As String
Dim nmes            As Integer

    nmes = right(cmbmes.Text, 2)
    cproame = Format(txtanno.Text, "0000") & Format(nmes, "00")
    csql = "SELECT IF4VALES.F4TIPCAM, IF4VALES.F4MONEDA, IF3VALES.* " & _
    "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) WHERE MONTH(F4FECVAL)=" & nmes & " ORDER BY F4NUMVAL"
    If rsretcab.State = adStateOpen Then rsretcab.Close
    rsretcab.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsretcab.EOF Then
    
        If rsparam_com.State = adStateOpen Then rsparam_com.Close
        rsparam_com.Open "SELECT * FROM PARAM_COM WHERE F1CODEMP ='" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
        If Not rsparam_com.EOF Then
            cf1ctaretencion = Trim("" & rsparam_com.Fields("F1CTARETENCION"))
            cf1origen_ret = Trim("" & rsparam_com.Fields("F1ORIGEN_RET"))
        End If
        rsparam_com.Close
        
        cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\templus.mdb;Persist Security Info=False" '"\DB_CONTA.MDB;Persist Security Info=False"
        cnn_form.Open cconex_form
        cnn_form.Execute ("DELETE * FROM CONTABLE")
        
        rscontable.Open "SELECT * FROM CONTABLE", cnn_form, adOpenDynamic, adLockOptimistic
        
        ncompro = 0
        rsretcab.MoveFirst
        Do While Not rsretcab.EOF
            If Len(Trim(rsretcab.Fields("TRANSFERIDO") & "")) = 0 Then
                
                lblcompro.Caption = "Nº Comprobante " & rsretcab.Fields("SERIE") & "/" & rsretcab.Fields("NUM_DOCUMENTO")
                lblcompro.Refresh
                
                ntc = 0#
                nelemen = 0
                ncompro = ncompro + 1
                If rscambios.State = adStateOpen Then rscambios.Close
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & rsretcab.Fields("FECHA") & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rscambios.EOF Then
                    ntc = Val("" & rscambios.Fields("CAMBIO"))
                End If
                rscambios.Close
    
                GENERA_CAB
                If rsretcab.Fields("ANULADO") & "" <> "S" Then
                    Csql2 = "SELECT * FROM RETENMOV WHERE SERIE_D='" & rsretcab.Fields("SERIE") & "' AND NUM_DOCUMENTOS='" & rsretcab.Fields("NUM_DOCUMENTO") & "' ORDER BY FECHA_EMISION"
                    If rsretdet.State = adStateOpen Then rsretdet.Close
                    rsretdet.Open Csql2, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsretdet.EOF Then
                        rsretdet.MoveFirst
                        Do While Not rsretdet.EOF
                            GENERA_DET
                            rsretdet.MoveNext
                        Loop
                    End If
                    rsretdet.Close
                End If
            End If
            
            rsretcab.MoveNext
        Loop
        
        rscontable.Close
        cnn_form.Close
        
        MsgBox "Fin del proceso.", vbInformation, "Atención"
        
    Else
        MsgBox "No existen movimientos en el mes para ser procesados.", vbInformation, "Atención"
    End If
    rsretcab.Close

End Sub

Private Sub Form_Load()
    
    txtanno.Text = wanno
    
    cmbmes.AddItem " Enero     " & Space(80) & "01"
    cmbmes.AddItem " Febrero   " & Space(80) & "02"
    cmbmes.AddItem " Marzo     " & Space(80) & "03"
    cmbmes.AddItem " Abril     " & Space(80) & "04"
    cmbmes.AddItem " Mayo      " & Space(80) & "05"
    cmbmes.AddItem " Junio     " & Space(80) & "06"
    cmbmes.AddItem " Julio     " & Space(80) & "07"
    cmbmes.AddItem " Agosto    " & Space(80) & "08"
    cmbmes.AddItem " Setiembre " & Space(80) & "09"
    cmbmes.AddItem " Octubre   " & Space(80) & "10"
    cmbmes.AddItem " Noviembre " & Space(80) & "11"
    cmbmes.AddItem " Diciembre " & Space(80) & "12"
    
    cmbmes.ListIndex = Month(Date) - 1
    
End Sub

Private Sub GENERA_CAB()

    nelemen = nelemen + 1
    rscontable.AddNew
    rscontable.Fields("F3PROAME") = cproame
    rscontable.Fields("F3COMPRO") = cf1origen_ret & Format(ncompro, "00000")
    rscontable.Fields("F3ELEMEN") = Format(nelemen, "0000")
    rscontable.Fields("F3ORIGEN") = cf1origen_ret
    rscontable.Fields("F3FCHOPR") = rsretcab.Fields("FECHA")
    rscontable.Fields("F5CODCTA") = cf1ctaretencion
    rscontable.Fields("F3CHEQUE") = rsretcab.Fields("NUM_DOCUMENTO") & ""
    rscontable.Fields("F3NROREF") = rsretcab.Fields("NUM_DOCUMENTO") & ""
    
    If rsretcab.Fields("ANULADO") & "" <> "S" Then
        rscontable.Fields("F3DETALL") = rsretcab.Fields("NOMBRE") & ""
        rscontable.Fields("F3IMPORTE") = rsretcab.Fields("RETENIDO")
        If ntc > 0# Then
            rscontable.Fields("F3TIPCAMBD") = ntc
            rscontable.Fields("F3IMPORTED") = Format(rsretcab.Fields("RETENIDO") / ntc, "0.00")
        End If
    Else
        rscontable.Fields("F3DETALL") = "A N U L A D O"
        rscontable.Fields("F3IMPORTE") = 0#
        rscontable.Fields("F3TIPCAMBD") = 0#
        rscontable.Fields("F3IMPORTED") = 0#
    End If
        
    rscontable.Fields("F3DEBHAB") = "H"
    rscontable.Fields("F3MONEDA") = "S"
    rscontable.Fields("F3TIPDOC") = "Val"
    rscontable.Fields("F3RUC") = rsretcab.Fields("RUC") & ""
    rscontable.Fields("F3SERDOC") = rsretcab.Fields("SERIE") & ""
    rscontable.Fields("F3COMP_RETENCION") = rsretcab.Fields("SERIE") & "/" & rsretcab.Fields("NUM_DOCUMENTO") & ""
    rscontable.Update

End Sub

Private Sub GENERA_DET()
Dim ccomp       As String
Dim cmes        As String
Dim cnummov     As String

    nelemen = nelemen + 1
    rscontable.AddNew
    rscontable.Fields("F3PROAME") = cproame
    rscontable.Fields("F3COMPRO") = cf1origen_ret & Format(ncompro, "00000")
    rscontable.Fields("F3ELEMEN") = Format(nelemen, "0000")
    rscontable.Fields("F3ORIGEN") = cf1origen_ret
    rscontable.Fields("F3FCHOPR") = rsretcab.Fields("FECHA") 'rsretdet.Fields("FECHA_EMISION")
    rscontable.Fields("F3DETALL") = rsretcab.Fields("NOMBRE") & ""
    
    If Len(Trim(rsretdet.Fields("REGCOMP") & "")) > 0 Then
        cmes = Mid(rsretdet.Fields("REGCOMP") & "", 1, 2)
        cnummov = Mid(rsretdet.Fields("REGCOMP") & "", 3, 7)
        ccomp = "SELECT F4CTACONT FROM REGISDOC WHERE F4MESMOV='" & cmes & "' AND F4NUMMOV='" & cnummov & "'"
        If rsregisdoc.State = adStateOpen Then rsregisdoc.Close
        rsregisdoc.Open ccomp, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsregisdoc.EOF Then
            rscontable.Fields("F5CODCTA") = rsregisdoc.Fields("F4CTACONT") & ""
        End If
        rsregisdoc.Close
    End If
    
    rscontable.Fields("F3CHEQUE") = rsretdet.Fields("NUMERO_CORRELA") & ""
    rscontable.Fields("F3NROREF") = rsretdet.Fields("NUMERO_CORRELA") & ""
    rscontable.Fields("F3IMPORTE") = rsretdet.Fields("IMPORTE_RETENIDO")
    If rsretdet.Fields("TIPO") & "" = "07" Then
        rscontable.Fields("F3DEBHAB") = "H"
    Else
        rscontable.Fields("F3DEBHAB") = "D"
    End If
    rscontable.Fields("F3MONEDA") = "S"
    If ntc > 0# Then
        rscontable.Fields("F3TIPCAMBD") = ntc
        rscontable.Fields("F3IMPORTED") = Format(rsretdet.Fields("IMPORTE_RETENIDO") / ntc, "0.00")
    End If
    
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open "SELECT F2ABREV FROM DOCUMENTOS WHERE F2CODDOC='" & rsretdet.Fields("TIPO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        rscontable.Fields("F3TIPDOC") = rsdocumentos.Fields("F2ABREV") & ""
    End If
    rsdocumentos.Close
    
    rscontable.Fields("F3RUC") = rsretcab.Fields("RUC") & ""
    rscontable.Fields("F3SERDOC") = rsretdet.Fields("SERIE") & ""
    rscontable.Fields("F3COMP_RETENCION") = rsretcab.Fields("SERIE") & "/" & rsretcab.Fields("NUM_DOCUMENTO") & ""
    rscontable.Update
    
End Sub
