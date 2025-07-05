VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form transferir_ret 
   Caption         =   "Transferencia de Asientos Contables"
   ClientHeight    =   2880
   ClientLeft      =   4395
   ClientTop       =   2130
   ClientWidth     =   4380
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
   ScaleHeight     =   2880
   ScaleWidth      =   4380
   Begin Threed.SSPanel SSPanel1 
      Height          =   1995
      Left            =   180
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
         Height          =   330
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
      Left            =   855
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
      Left            =   2250
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
Attribute VB_Name = "transferir_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TRANSFERIR_COMP_RET()
Dim conex_cnt       As String
Dim cnn_cnt         As New ADODB.Connection
Dim rscnt           As New ADODB.Recordset
Dim cconex_form     As String
Dim cnn_form        As New ADODB.Connection
Dim cmes            As String
Dim conex_conta     As String
Dim cnn_conta       As New ADODB.Connection
        
    conex_cnt = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wcontacnt & "\CNT_CONT.MDB" & ";Persist Security Info=False"
    cnn_cnt.Open conex_cnt
    
    If rscnt.State = adStateOpen Then rscnt.Close
    rscnt.Open "SELECT * FROM CF1CNT WHERE F1DIR ='" & wempresa & "'", cnn_cnt, adOpenDynamic, adLockOptimistic
    If Not rscnt.EOF Then
        wdg1 = Val(rscnt.Fields("F1DGRAD1") & "")
        wdg2 = Val(rscnt.Fields("F1DGRAD2") & "")
        wdg3 = Val(rscnt.Fields("F1DGRAD3") & "")
        wdg4 = Val(rscnt.Fields("F1DGRAD4") & "")
        wdg5 = Val(rscnt.Fields("F1DGRAD5") & "")
    End If
    rscnt.Close
    cnn_cnt.Close
    
    cmes = Right(cmbmes.Text, 2)
    
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\DB_CONTA.MDB;Persist Security Info=False"
    cnn_form.Open cconex_form
    
    conex_conta = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\DB_MOV" & cmes & ".MDB" & ";Persist Security Info=False"
    cnn_conta.Open conex_conta
    
    cconex_dbtabla = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\DB_TABLA.MDB;Persist Security Info=False"
    If cnn_dbtabla.State = adStateOpen Then cnn_dbtabla.Close
    cnn_dbtabla.Open cconex_dbtabla
    
    cconex_analisis = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\DB_ANALI.MDB;Persist Security Info=False"
    If cnn_analisis.State = adStateOpen Then cnn_analisis.Close
    cnn_analisis.Open cconex_analisis
    
    TRANSFIERE_ASIENTOS cnn_conta, cnn_form, cmes, cnn_dbbancos, "R", "RETENDOC"

    cnn_conta.Close

    cnn_form.Close
    
    cnn_dbtabla.Close
    
    cnn_analisis.Close

End Sub

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            Me.MousePointer = 11
            TRANSFERIR_COMP_RET
            Me.MousePointer = 1
            MsgBox "Fin del proceso.", vbInformation, "Atención"
        Case 1:
            Unload Me
    End Select

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
