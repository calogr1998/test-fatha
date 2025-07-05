VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmselegrupos 
   Caption         =   "Selección de Grupos de Gastos"
   ClientHeight    =   3540
   ClientLeft      =   2355
   ClientTop       =   1815
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7560
   Begin ComctlLib.Toolbar tblbar 
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
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
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3270
      Left            =   45
      TabIndex        =   1
      Top             =   75
      Width           =   7440
      _Version        =   65536
      _ExtentX        =   13123
      _ExtentY        =   5768
      _StockProps     =   15
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.Data dc_prov 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2700
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   2145
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmselegrupos.frx":0000
         Height          =   3015
         Left            =   180
         OleObjectBlob   =   "frmselegrupos.frx":0016
         TabIndex        =   2
         Top             =   120
         Width           =   7170
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   2340
      Top             =   -90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":0A05
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":0D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1039
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1143
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":145D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1777
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1881
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":1EB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmselegrupos.frx":21CF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmselegrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
                If dc_prov.Recordset.RecordCount > 0 Then
                    sw = False
                    cod_prove = DBGrid1.Columns(0)
                    frmreggrupos.Show 1
                    dc_prov.Refresh
                    DBGrid1.Refresh
                End If
            Case 27:
                Unload Me
        End Select
    Else
        If KeyAscii = 13 Then
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
  dc_prov.DatabaseName = wrutabancos & "\DB_BANCOS.mdb"
  dc_prov.RecordSource = "select * from grupos where tipo = 'G' order by cod_grup"
  dc_prov.Refresh
  DBGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    T_PROV.Close
'    dbfactur.Close

End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Index
        Case 1:
            sw = True
            cod_prove = ""
            frmreggrupos.Show 1
            dc_prov.Refresh
            DBGrid1.Refresh
        Case 2:
            If DBGrid1.Row < 0 Then
                MsgBox "Seleccione el grupo a modificar.", 48, "Bancos"
                DBGrid1.SetFocus
                Exit Sub
            Else
                sw = False
                cod_prove = DBGrid1.Columns(0)
                frmreggrupos.Show 1
                dc_prov.Refresh
                DBGrid1.Refresh

            End If
        Case 3:
            frmimpgrupos.Show 1
        Case 4:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
cconexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
cnn_dbbancos.Open cconexion
'  Set dbfactur = OpenDatabase(wrutabanco & "\DB_BANCOS.MDB")
  
End Sub

