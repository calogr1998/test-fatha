VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form mant_centros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Centros de Costos"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   6495
      Begin VB.TextBox txtdescrip 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   900
         Width           =   4785
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   495
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   135
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_centros.frx":0000
      ToolBars        =   "mant_centros.frx":7E74
   End
End
Attribute VB_Name = "mant_centros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Codigo          As String
Dim rsconceptos        As New ADODB.Recordset
Dim sql             As String
Dim wopcion         As Byte

Private Sub Form_Activate()
If sw_mant_ayuda = True Then
    SSActiveToolBars1.Tools(3).Visible = False
    SSActiveToolBars1.Tools(4).Visible = False
    'SSActiveToolBars1.Tools(10).Visible = False
Else
    SSActiveToolBars1.Tools(3).Visible = True
    SSActiveToolBars1.Tools(4).Visible = True
End If
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            nuevo_centro
        Case "ID_Grabar":
            grabar_centro
            If sw_mant_ayuda = True Then
                wcodcosto = txtcodigo.Text
                wdescosto = txtdescrip.Text
                sw_mant_ayuda = False
                Me.Hide
            End If

        Case "ID_Eliminar"
            eliminar_centro
'        Case "ID_Imprimir":
'            With Acr_Marcas
'                .DataControl1.ConnectionString = cnn_dbbancos
'                .DataControl1.Source = "Select * From ef2marcas order by f2codmar"
'                .fldfecha.Text = Format(Date, "DD/MM/YYYY")
'                .lblempresa.Caption = wnomcia
'                .Show 1
'            End With
        Case "ID_Lista"
            wopcion = 1
            Unload Me
        Case "ID_Salir"
            wopcion = 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
Me.MousePointer = 11
Me.left = 1600
Me.top = 1050

wopcion = 0
If sw_nuevo_doc = True Then
    nuevo_centro
Else
    actualizacion_centro Lista_Centros.dxDBGrid(nGridActive).Columns(0).Value
End If
Me.MousePointer = 1
End Sub

Private Sub actualizacion_centro(cod)
    
    sql = "select * from centros where f3costo='" & cod & "'"
    If rsconceptos.State = adStateOpen Then rsconceptos.Close
    rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsconceptos.EOF Then
        txtcodigo.Text = "" & rsconceptos.Fields("f3costo")
        txtdescrip.Text = rsconceptos.Fields("f3descrip") & ""
    End If
    
End Sub

Private Sub nuevo_centro()
    genera_cod
    txtcodigo.Enabled = True
    txtcodigo.Text = Codigo
    txtcodigo.Enabled = False
    txtdescrip.Text = ""
End Sub

Private Sub grabar_centro()
On Error GoTo graba
Dim amovs(0 To 1) As a_grabacion

sql = "select * from centros where f3costo='" & txtcodigo.Text & "' "
If rsconceptos.State = adStateOpen Then rsconceptos.Close
rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rsconceptos.EOF Then
    sw = 0
Else
    sw = 1
End If
rsconceptos.Close
amovs(0).campo = "f3costo": amovs(0).valor = txtcodigo.Text: amovs(0).TIPO = "T"
amovs(1).campo = "f3descrip": amovs(1).valor = txtdescrip.Text: amovs(1).TIPO = "T"

If sw = 1 Then
    GRABA_REGISTRO amovs(), "centros", "A", 1, cnn_dbbancos, ""
Else
    GRABA_REGISTRO amovs(), "centros", "M", 1, cnn_dbbancos, "f3costo='" & txtcodigo.Text & "'"
End If
txtcodigo.Enabled = False
Exit Sub


graba:
    If Err = 3186 Then
        For I% = 1 To 10000
        Next I%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If

End Sub

Private Sub eliminar_centro()
 Beep
    If MsgBox("¿Está seguro de eliminar el Centro...?", 36, "Atención") = 6 Then
        sql = "select f3costo from centros where f3costo='" & txtcodigo.Text & "' "
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
        rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconceptos.EOF Then
            If ctipoadm_bd = "M" Then
                sql = "DELETE from centros where f3costo='" & txtcodigo.Text & "' "
            Else
                sql = "DELETE * from centros where f3costo='" & txtcodigo.Text & "' "
            End If
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            txtcodigo.Enabled = True
            nuevo_centro
        Else
            Beep
        End If
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
'        txtcodigo.SetFocus
    End If

End Sub

Private Sub txtcodigo_GotFocus()
txtcodigo.SelStart = 0
txtcodigo.SelLength = Len(txtcodigo.Text)
End Sub


Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtdescrip.SetFocus
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    If Len(Trim(txtcodigo.Text)) > 0 Then
        sql = "select f3costo from centros where f3costo='" & txtcodigo.Text & "'"
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
        rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconceptos.EOF Then
            MsgBox "Código de Centro existe. Verifique.", vbInformation, "Atención"
            txtcodigo.SetFocus
        End If
        rsconceptos.Close
    End If
End Sub

Private Sub txtdescrip_GotFocus()
txtdescrip.SelStart = 0
txtdescrip.SelLength = Len(txtdescrip.Text)
End Sub

Private Sub genera_cod()
sql = "select f3costo from centros order by f3costo desc"
If rsconceptos.State = adStateOpen Then rsconceptos.Close
rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rsconceptos.EOF Then
    Codigo = rsconceptos.Fields("f3costo") + 1
    Codigo = Format(Codigo, "000")
Else
    Codigo = 1
    Codigo = Format(Codigo, "000")
End If
rsconceptos.Close
End Sub


