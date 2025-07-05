VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmcontc 
   Appearance      =   0  'Flat
   Caption         =   "Consistencia de Tipos de Cambio"
   ClientHeight    =   1635
   ClientLeft      =   3795
   ClientTop       =   2310
   ClientWidth     =   3705
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
   ScaleHeight     =   1635
   ScaleWidth      =   3705
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   1890
      TabIndex        =   4
      Top             =   1125
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   1125
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   915
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   3420
      TabIndex        =   1
      Top             =   90
      Width           =   3480
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
         Top             =   315
         Width           =   420
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
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmcontc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            PROCESA
            MsgBox "Fin del proceso.", 48, "Compras"
            cmdresp(1).SetFocus
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

    txtmes.Text = mes

End Sub

Private Sub PROCESA()
Dim CONSULTA As DAO.Recordset

    txtmes.Text = Format(txtmes.Text, "00")
    If MsgBox("Está seguro de realizar el proceso ?", 36, "Atención") = 6 Then
        Me.MousePointer = 11
        Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
        TbCabRegis.Index = "idmesnum"
        Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
        TbOfiRegis.Index = "idmesnum"

        Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
        Set tbcambios = dbcomtabla.OpenRecordset("cambios")
        tbcambios.Index = "cambios"
        
        Set CONSULTA = dbcompras.OpenRecordset("Select * From REGISDOC Where F4MESMOV = '" + txtmes.Text + "' order by F4NUMMOV")
        If CONSULTA.RecordCount > 0 Then
            CONSULTA.MoveFirst
            Do While Not CONSULTA.EOF
                tbcambios.Seek "=", CONSULTA.Fields("f4fecha")
                If Not tbcambios.NoMatch Then
                    If Format(tbcambios.Fields("cambio"), "#0.000") <> Format(CONSULTA.Fields("f4tipcam"), "#0.000") Then
                        TbCabRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                        If Not TbCabRegis.NoMatch Then
                            TbCabRegis.Edit
                            TbCabRegis.Fields("f4tipcam") = tbcambios.Fields("cambio")
                            TbCabRegis.Update
                            TbOfiRegis.Seek "=", CONSULTA.Fields("f4mesmov"), CONSULTA.Fields("f4nummov")
                            If Not TbOfiRegis.NoMatch Then
                                TbOfiRegis.Edit
                                TbOfiRegis.Fields("f4tipcam") = tbcambios.Fields("cambio")
                                If CONSULTA.Fields("f4moneda") = "D" Then
                                    TbOfiRegis.Fields("f4basimp") = IIf(CONSULTA.Fields("f4basimp") <> 0, Format(CONSULTA.Fields("f4basimp") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4monina") = IIf(CONSULTA.Fields("f4monina") <> 0, Format(CONSULTA.Fields("f4monina") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4igv") = IIf(CONSULTA.Fields("f4igv") <> 0, Format(CONSULTA.Fields("f4igv") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4otrimp") = IIf(CONSULTA.Fields("f4otrimp") <> 0, Format(CONSULTA.Fields("f4otrimp") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4total") = IIf(CONSULTA.Fields("f4total") <> 0, Format(CONSULTA.Fields("f4total") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4montoret") = IIf(CONSULTA.Fields("f4montoret") <> 0, Format(CONSULTA.Fields("f4montoret") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                    TbOfiRegis.Fields("f4fonavi") = IIf(CONSULTA.Fields("f4fonavi") <> 0, Format(CONSULTA.Fields("f4fonavi") * TbCabRegis.Fields("f4tipcam"), "#0.00"), 0#)
                                End If
                                TbOfiRegis.Update
                            End If
                        End If
                    End If
                End If
                CONSULTA.MoveNext
            Loop
        End If
        TbCabRegis.Close
        TbOfiRegis.Close
        CONSULTA.Close
        dbcompras.Close
        tbcambios.Close
        dbcomtabla.Close
        Me.MousePointer = 1
    End If

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            cmdresp(0).SetFocus
        Else
            MsgBox "El mes es incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub

