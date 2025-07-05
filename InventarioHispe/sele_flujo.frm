VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frm_SeleFlujo 
   Caption         =   "Flujo"
   ClientHeight    =   3870
   ClientLeft      =   2550
   ClientTop       =   2670
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5655
   Begin ComctlLib.Toolbar tblbar 
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   -15
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Imprimir Flujo y Gastos"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CRFlujos 
      Left            =   3720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3630
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   5565
      _Version        =   65536
      _ExtentX        =   9816
      _ExtentY        =   6403
      _StockProps     =   15
      BackColor       =   12632256
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
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1725
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Data dc_bancos 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1395
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2235
         Visible         =   0   'False
         Width           =   2355
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "sele_flujo.frx":0000
         Height          =   2610
         Left            =   135
         OleObjectBlob   =   "sele_flujo.frx":0018
         TabIndex        =   1
         Top             =   120
         Width           =   5325
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   720
         Left            =   105
         TabIndex        =   2
         Top             =   2760
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
         _ExtentY        =   1270
         _StockProps     =   14
         Caption         =   " Búsqueda "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Buscado 
            Height          =   300
            Left            =   945
            TabIndex        =   3
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   345
            Width           =   495
         End
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   2070
      Top             =   -135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":0D37
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":1051
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":136B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":1475
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":178F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":1AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":1BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":1ECD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":21E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":2501
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sele_flujo.frx":281B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_SeleFlujo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub IMPRIMIR()
Dim codigo As String
Dim dbtempo As DAO.Database
Dim tmpflujotab As DAO.Recordset

    Set dbtempo = Workspaces(0).OpenDatabase(wrutatemp & "db_tempo.MDB")
    dbtempo.Execute ("Delete * From TmpFlujoGastos")
    Set tmpflujotab = dbtempo.OpenRecordset("TmpFlujoGastos")

    Data2.Recordset.MoveFirst
    Do While Not Data2.Recordset.EOF
        codigo = Trim(Data2.Recordset.Fields("Cod_fjo"))
        Data1.RecordSource = "select * from bf9gin where GRUPOFLUJO ='" & codigo & "'"
        Data1.Refresh
        If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                tmpflujotab.AddNew
                tmpflujotab.Fields("CODFLUJO") = Data2.Recordset.Fields("Cod_fjo")
                tmpflujotab.Fields("DESFLUJO") = Data2.Recordset.Fields("descripcion")
                tmpflujotab.Fields("CODIGO") = Data1.Recordset.Fields("CODIGO")
                tmpflujotab.Fields("NOMBRE") = Data1.Recordset.Fields("NOMBRE")
                tmpflujotab.Fields("GRUPOFLUJO") = Data1.Recordset.Fields("GRUPOFLUJO")
                tmpflujotab.Update
                Data1.Recordset.MoveNext
            Loop
        End If
        Data2.Recordset.MoveNext
    Loop
    CRFlujos.DataFiles(0) = wrutabancos & "\db_tabla.mdb"
    CRFlujos.ReportFileName = wrutatemp & "flujos.rpt"
    CRFlujos.Action = 1

End Sub


Private Sub Buscado_Change()
  
   dc_bancos.Recordset.FindFirst "codigo like " & "'" & Trim(Buscado.Text) & "*" & "'"
  
End Sub

Private Sub Buscado_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then DBGrid1.SetFocus
  
End Sub

Private Sub DBGrid1_DblClick()
   
   DBGrid1_KeyPress 13
      
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

   DBGrid1_KeyPress KeyCode
  
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

   If Not (LLAMADA = "ayuda") Then
      Select Case KeyAscii
         Case 13:
            If dc_bancos.Recordset.RecordCount > 0 Then
               sw = False
               cod_bank = DBGrid1.Columns(0)
'               frm_showFlujo.Show 1
            End If
         Case 27:
            Unload Me
      End Select
   Else
      If KeyAscii = 13 Then
         LLAMADA = ""
         cod_grupo = DBGrid1.Columns(0)
         des_grupo = DBGrid1.Columns(1)
         Unload Me
      End If
      If KeyAscii = 27 Then
        Unload Me
      End If
   End If
   
End Sub

Private Sub Form_Activate()
'  dc_prov.DatabaseName = wrutabanco & "\DB_BANCOS.mdb"
'  dc_prov.RecordSource = "select * from grupos where tipo = 'G' order by cod_grup"
'  dc_prov.Refresh

   
End Sub

Private Sub Form_Load()

    Data1.DatabaseName = wrutabancos & "\DB_BANCOS.MDB"
    Data1.RecordSource = "bf9gin"
    Data2.DatabaseName = wrutabancos & "\db_tabla.mdb"
    Data2.RecordSource = "flujocaja"
    
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cconexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    cnn_dbbancos.Open cconexion
    
    dc_bancos.DatabaseName = wrutabancos & "\db_bancos.mdb"
    If wdestino = "I" Then
        dc_bancos.RecordSource = "select * from grupos_flujo where left(codigo,1)='C'"
    Else
        If wdestino = "E" Then
            dc_bancos.RecordSource = "select * from grupos_flujo where left(codigo,1)='P'"
        Else
            dc_bancos.RecordSource = "select * from grupos_flujo"
        End If
    End If
    dc_bancos.Refresh
        
    DBGrid1.Refresh
    
End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Index
        Case 1:
            cod_bank = ""
            sw = True
            frm_showFlujo.Show 1
        Case 2:
            If DBGrid1.Row < 0 Then
                MsgBox "Seleccione el código a modificar.", 48, "Bancos"
                DBGrid1.SetFocus
                Exit Sub
            Else
                sw = False
                cod_bank = DBGrid1.Columns(0)
'                frm_showFlujo.Show 1
            End If
        Case 3: ' IMPRIMIR GRUPO DE FLUJOS
'            IMPRIMIR
        Case 4: ' IMPRIMIR FLUJO Y GASTOS
'            llena_temp
'            CRFlujos.DataFiles(0) = wrutatemp & "db_tempo.mdb"
'            CRFlujos.ReportFileName = wrutatemp & "flujoygastos2.rpt"
'            CRFlujos.Action = 1
        Case 5:
            Unload Me
    End Select

End Sub

Private Sub llena_temp()
Dim codigo As String
Dim dbtempo As DAO.Database
Dim tmpflujotab As DAO.Recordset

    Set dbtempo = Workspaces(0).OpenDatabase(wrutatemp & "db_tempo.MDB")
    dbtempo.Execute ("Delete * From TmpFlujoGastos")
    Set tmpflujotab = dbtempo.OpenRecordset("TmpFlujoGastos")

    Data2.Recordset.MoveFirst
    Do While Not Data2.Recordset.EOF
        codigo = Trim(Data2.Recordset.Fields("Cod_fjo"))
        Data1.RecordSource = "select * from bf9gin where GRUPOFLUJO ='" & codigo & "'"
        Data1.Refresh
        If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                tmpflujotab.AddNew
                tmpflujotab.Fields("CODFLUJO") = Data2.Recordset.Fields("Cod_fjo")
                tmpflujotab.Fields("DESFLUJO") = Data2.Recordset.Fields("descripcion")
                tmpflujotab.Fields("CODIGO") = Data1.Recordset.Fields("CODIGO")
                tmpflujotab.Fields("NOMBRE") = Data1.Recordset.Fields("NOMBRE")
                tmpflujotab.Fields("GRUPOFLUJO") = Data1.Recordset.Fields("GRUPOFLUJO")
                tmpflujotab.Update
                Data1.Recordset.MoveNext
            Loop
        End If
        Data2.Recordset.MoveNext
    Loop
    tmpflujotab.Close
    dbtempo.Close
            
End Sub
