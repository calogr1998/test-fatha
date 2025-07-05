VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Reg_Documentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Documentos"
   ClientHeight    =   3270
   ClientLeft      =   3360
   ClientTop       =   5670
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "Reg_Documentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   635
      ButtonWidth     =   1667
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir     "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   5340
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Documentos.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Documentos.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Documentos.frx":0B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Documentos.frx":10DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4395
      Begin VB.TextBox TxtCodDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1620
         Width           =   540
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo"
         Height          =   615
         Left            =   180
         TabIndex        =   5
         Top             =   2040
         Width           =   3855
         Begin VB.OptionButton OptTipo 
            Caption         =   "Ambos"
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Por Pagar"
            Height          =   195
            Index           =   1
            Left            =   1500
            TabIndex        =   7
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Por Cobrar"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.TextBox txtcodigo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtdescrip 
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   3
         Top             =   585
         Width           =   2445
      End
      Begin VB.TextBox txtabrev 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   2
         Top             =   900
         Width           =   540
      End
      Begin VB.TextBox txtdebhab 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label Label 
         Caption         =   "Código CONCAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1665
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Debe/Haber"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Abreviatura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   855
      End
   End
End
Attribute VB_Name = "Reg_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Rs As New ADODB.Recordset
'Dim Amov_Doc(0 To 20) As a_grabacion
'Dim T_PROV As New ADODB.Recordset
'
'Private Sub chkobliga_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      chkserie.SetFocus
'   End If
'End Sub
'
'Private Sub chkpagos_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      chkobliga.SetFocus
'   End If
'
'End Sub
'
'Private Sub chksaldos_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      chkpagos.SetFocus
'   End If
'
'End Sub
'
'Private Sub chktransfer_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      chksaldos.SetFocus
'   End If
'
'End Sub
'
'
'
'Private Sub Form_Load()
'
'If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
'cnn_dbbancos.Open
'
'    If sw = False Then  'modifica
'        txtcodigo.Text = Cod_Prove
'        csql = "SELECT * FROM DOCUMENTOS WHERE F2CODDOC='" & txtcodigo.Text & "'"
'        If T_PROV.State = 1 Then T_PROV.Close
'        T_PROV.Open csql, cnn_dbbancos, 3, 1
'        txtcodigo.Enabled = False
'        If T_PROV.RecordCount > 0 Then
'            txtdescrip.Text = T_PROV.Fields("f2desdoc") & ""
'            txtabrev.Text = T_PROV.Fields("f2abrev") & ""
'            txtdebhab.Text = T_PROV.Fields("f2debhab") & ""
'            TxtCodDoc.Text = T_PROV.Fields("F2ABREV_CONCAR") & ""
'
'            Select Case T_PROV!F2TIPO
'            Case "C"
'                OptTipo(0).Value = True
'            Case "P"
'                OptTipo(1).Value = True
'            Case "A"
'                OptTipo(2).Value = True
'            End Select
'
'        End If
'    End If
'
'End Sub
'
'Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)
'
'   Select Case Button.Index
'      Case 1:
'         grabar
'         Unload Me
'      Case 2:
'         eliminar
'         Unload Me
'      Case 3:
'         Unload Me
'   End Select
'
'End Sub
'
'Private Sub grabar()
'On Error GoTo error_graba
'
'
'
'
'    Amov_Doc(0).campo = "f2desdoc": Amov_Doc(0).valor = Trim(txtdescrip.Text) & "": Amov_Doc(0).TIPO = "T"
'    Amov_Doc(1).campo = "f2abrev": Amov_Doc(1).valor = Trim(txtabrev.Text) & "": Amov_Doc(1).TIPO = "T"
'    Amov_Doc(2).campo = "f2debhab": Amov_Doc(2).valor = Trim(txtdebhab.Text) & "": Amov_Doc(2).TIPO = "T"
'
'    If OptTipo(0).Value = True Then
'        Amov_Doc(3).campo = "F2TIPO": Amov_Doc(3).valor = "C": Amov_Doc(3).TIPO = "T"
'    ElseIf OptTipo(1).Value = True Then
'        Amov_Doc(3).campo = "F2TIPO": Amov_Doc(3).valor = "P": Amov_Doc(3).TIPO = "T"
'    ElseIf OptTipo(2).Value = True Then
'        Amov_Doc(3).campo = "F2TIPO": Amov_Doc(3).valor = "A": Amov_Doc(3).TIPO = "T"
'    End If
'    Amov_Doc(4).campo = "F2ABREV_CONCAR": Amov_Doc(4).valor = TxtCodDoc.Text: Amov_Doc(4).TIPO = "T"
'    'Comprueba si existe
'    csql = "SELECT * FROM DOCUMENTOS WHERE F2CODDOC='" & txtcodigo.Text & "'"
'    If T_PROV.State = 1 Then T_PROV.Close
'    T_PROV.Open csql, cnn_dbbancos, 3, 1
'    '*******************
'    If T_PROV.RecordCount > 0 Then
'        GRABA_REGISTRO Amov_Doc, "documentos", "M", 4, cnn_dbbancos, "F2CODDOC='" & txtcodigo.Text & "'"
'    Else
'
'        Amov_Doc(5).campo = "f2coddoc": Amov_Doc(5).valor = CorrelaDocumento: Amov_Doc(5).TIPO = "T"
'        GRABA_REGISTRO Amov_Doc, "documentos", "A", 5, cnn_dbbancos, ""
'    End If
'
'
'    Exit Sub
'
'error_graba:
'    Select Case Err
'        Case 3186:
'            MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos", vbCritical, "CONTROL Plus!"
'            For I = 1 To 10000
'            Next
'            Resume
'        Case 3163:
'            MsgBox Err.Description & "  Verifique la base de datos.", 48, "CONTROL Plus!"
'            Resume Next
'        Case Else
'            MsgBox Err.Description, 48, "CONTROL Plus!"
'            Resume Next
'    End Select
'
'End Sub
'
'Private Function CorrelaDocumento() As String
'    csql = "SELECT top 1 f2coddoc FROM DOCUMENTOS order by f2coddoc desc"
'    If Rs.State = 1 Then Rs.Close
'    Rs.Open csql, cnn_dbbancos, 3, 1
'    If Rs.RecordCount > 0 Then
'        CorrelaDocumento = Format(Val(Rs!f2coddoc & "") + 1, "00")
'    Else
'        CorrelaDocumento = "01"
'    End If
'End Function
'
'Private Sub eliminar()
'
'   If MsgBox("Está seguro de eliminar el registro ? ", vbYesNo, "CONTROL Plus!") = vbYes Then
'      T_PROV.Seek "=", txtcodigo.Text
'      If T_PROV.RecordCount > 0 Then
'         T_PROV.Delete
'      Else
'         MsgBox "El registro no puede eliminarse porque aún no ha sido grabado", 48, "CONTROL Plus!"
'      End If
'   End If
'
'End Sub
'
'Private Sub tblbara_ButtonClick(ByVal Button As ComctlLib.Button)
'
'End Sub
'
'Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'   Select Case Trim(Button.Caption)
'      Case "Grabar":
'         grabar
'         Unload Me
'      Case "Eliminar":
'         eliminar
'         Unload Me
'      Case "Salir":
'         Unload Me
'   End Select
'End Sub
'
'Private Sub txtabrev_GotFocus()
'
'   txtabrev.SelStart = 0
'   txtabrev.SelLength = Len(txtabrev.Text)
'
'End Sub
'
'Private Sub txtabrev_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      txtdebhab.SetFocus
'   End If
'
'End Sub
'
'Private Sub TxtCodDoc_DblClick()
'Call TxtCodDoc_KeyDown(113, 0)
'End Sub
'
'Private Sub TxtCodDoc_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 113 Then
'    cCodDocConcar = ""
'    Ayuda_Documentos_CONCAR.Show 1
'    Unload Ayuda_Documentos_CONCAR
'    If Len(Trim(cCodDocConcar)) > 0 Then
'        TxtCodDoc.Text = cCodDocConcar
'    End If
'End If
'End Sub
'
'Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      txtdescrip.SetFocus
'   End If
'
'End Sub
'
'Private Sub txtdebhab_GotFocus()
'
'   txtdebhab.SelStart = 0
'   txtdebhab.SelLength = Len(txtdebhab.Text)
'
'End Sub
'
'Private Sub txtdebhab_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      chkTransfer.SetFocus
'   End If
'
'End Sub
'
'Private Sub txtdebhab_LostFocus()
'
'    If Len(Trim(txtdebhab.Text)) = 0 Then
'        MsgBox "Falta indicar el debe o haber. Verifique.", 48, "CONTROL Plus!"
'        txtdebhab.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtdescrip_GotFocus()
'
'   txtdescrip.SelStart = 0
'   txtdescrip.SelLength = Len(txtdescrip.Text)
'
'End Sub
'
'Private Sub txtdescrip_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'      txtabrev.SetFocus
'   End If
'
'End Sub
