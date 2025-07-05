VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmconbd 
   Appearance      =   0  'Flat
   Caption         =   "Consistencia de Bases de Datos"
   ClientHeight    =   1560
   ClientLeft      =   3945
   ClientTop       =   2085
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1560
   ScaleWidth      =   3975
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   630
      TabIndex        =   1
      Top             =   1080
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   3735
      TabIndex        =   3
      Top             =   90
      Width           =   3795
      Begin Crystal.CrystalReport cryreporte 
         Left            =   2700
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.TextBox txtmes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   0
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   315
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   2025
      TabIndex        =   2
      Top             =   1080
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "frmconbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            PROCESA
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

    txtmes.Text = Format(mes, "00")

End Sub

Private Sub PROCESA()
Dim CONSULTA As DAO.Recordset
Dim dbtempo  As DAO.Database
Dim tbtempo  As DAO.Recordset

    Me.MousePointer = 11
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")

    Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
    dbtempo.Execute ("delete * from temp_consis")
    Set tbtempo = dbtempo.OpenRecordset("temp_consis")

    If TbCabRegis.RecordCount > 0 Then
        '------------- verifica si la cabecera tiene detalles
        Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where f4mesmov = '" & txtmes.Text & "' order by f4nummov")
        If CONSULTA.RecordCount > 0 Then
            CONSULTA.MoveFirst
            Do While Not CONSULTA.EOF
                TbDetRegis.Index = "IDMESNUM"
                TbDetRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                If TbDetRegis.NoMatch Then
                    tbtempo.AddNew
                    tbtempo.Fields("mesmov") = CONSULTA.Fields("f4mesmov")
                    tbtempo.Fields("nummov") = CONSULTA.Fields("f4nummov")
                    tbtempo.Fields("empresa") = wnomcia
                    tbtempo.Fields("descrip") = "No tiene detalle."
                    tbtempo.Update
                End If
                CONSULTA.MoveNext
            Loop
            '------------- verifica si el detalle tiene cabecera
            Set CONSULTA = dbcompras.OpenRecordset("select * from regismov where f4mesmov = '" & txtmes.Text & "' order by f4nummov")
            If CONSULTA.RecordCount > 0 Then
                CONSULTA.MoveFirst
                Do While Not CONSULTA.EOF
                    TbCabRegis.Index = "IDMESNUM"
                    TbCabRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                    If TbCabRegis.NoMatch Then
                        tbtempo.AddNew
                        tbtempo.Fields("mesmov") = CONSULTA.Fields("f4mesmov")
                        tbtempo.Fields("nummov") = CONSULTA.Fields("f4nummov")
                        tbtempo.Fields("empresa") = wnomcia
                        tbtempo.Fields("descrip") = "No tiene cabecera."
                        tbtempo.Update
                    End If
                    CONSULTA.MoveNext
                Loop
            End If
            '------------- verifica si la cabecera esta en el archivo oficial
            Set CONSULTA = dbcompras.OpenRecordset("select * from regisdoc where f4mesmov = '" & txtmes.Text & "' order by f4nummov")
            If CONSULTA.RecordCount > 0 Then
                CONSULTA.MoveFirst
                Do While Not CONSULTA.EOF
                    TbOfiRegis.Index = "IDMESNUM"
                    TbOfiRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                    If TbOfiRegis.NoMatch Then
                        tbtempo.AddNew
                        tbtempo.Fields("mesmov") = CONSULTA.Fields("f4mesmov")
                        tbtempo.Fields("nummov") = CONSULTA.Fields("f4nummov")
                        tbtempo.Fields("empresa") = wnomcia
                        tbtempo.Fields("descrip") = "No existe en el archivo oficial."
                        tbtempo.Update
                    End If
                    CONSULTA.MoveNext
                Loop
            End If
            '------------- verifica si el registro del archivo oficial tiene cabecera
            Set CONSULTA = dbcompras.OpenRecordset("select * from regisofi where f4mesmov = '" & txtmes.Text & "' order by f4nummov")
            If CONSULTA.RecordCount > 0 Then
                CONSULTA.MoveFirst
                Do While Not CONSULTA.EOF
                    TbCabRegis.Index = "IDMESNUM"
                    TbCabRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                    If TbCabRegis.NoMatch Then
                        tbtempo.AddNew
                        tbtempo.Fields("mesmov") = CONSULTA.Fields("f4mesmov")
                        tbtempo.Fields("nummov") = CONSULTA.Fields("f4nummov")
                        tbtempo.Fields("empresa") = wnomcia
                        tbtempo.Fields("descrip") = "El registro del archivo oficial no tiene cabecera."
                        tbtempo.Update
                    End If
                    CONSULTA.MoveNext
                Loop
            End If
            '------------------------------
        Else
            MsgBox "No existen registros para ser procesados.", 48, "Compras"
        End If
        If tbtempo.RecordCount > 0 Then
            cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
            cryreporte.ReportFileName = wrutatemp & "\rptconbd.rpt"
            cryreporte.Action = 1
        Else
            MsgBox "No existen registros inconsistentes.", 48, "Compras"
        End If
    Else
        MsgBox "No existen registros para ser procesados.", 48, "Compras"
    End If
         
    TbCabRegis.Close
    TbOfiRegis.Close
    TbDetRegis.Close
    CONSULTA.Close
    dbcompras.Close

    tbtempo.Close
    dbtempo.Close
    Me.MousePointer = 1

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtmes.Text = Format(txtmes.Text, "00")
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            cmdresp(0).SetFocus
        Else
            MsgBox "Mes incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub

Private Sub txtmes_LostFocus()

    If Len(Trim(txtmes.Text)) > 0 Then
        txtmes.Text = Format(txtmes.Text, "00")
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            cmdresp(0).SetFocus
        Else
            MsgBox "Mes incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub

