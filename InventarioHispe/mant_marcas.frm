VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form mant_marcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Marcas"
   ClientHeight    =   2460
   ClientLeft      =   2655
   ClientTop       =   2775
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6750
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   6495
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         Height          =   960
         Left            =   270
         TabIndex        =   7
         Top             =   1620
         Width           =   6000
         Begin Threed.SSOption opttipo 
            Height          =   285
            Index           =   0
            Left            =   855
            TabIndex        =   2
            Top             =   450
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Nacional"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption opttipo 
            Height          =   285
            Index           =   1
            Left            =   3960
            TabIndex        =   3
            Top             =   450
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Importado"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
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
         TabIndex        =   0
         Top             =   450
         Width           =   735
      End
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
         TabIndex        =   1
         Top             =   900
         Width           =   4785
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
         TabIndex        =   6
         Top             =   495
         Width           =   495
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
         TabIndex        =   5
         Top             =   945
         Width           =   855
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   360
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_marcas.frx":0000
      ToolBars        =   "mant_marcas.frx":7E74
   End
End
Attribute VB_Name = "mant_marcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Codigo          As String
Dim rsmarcas        As New ADODB.Recordset
Dim sql             As String
Dim wopcion         As Byte

Private Sub Form_Activate()
If sw_mant_ayuda = True Then
    SSActiveToolBars1.Tools(3).Visible = False
    SSActiveToolBars1.Tools(4).Visible = False
Else
    SSActiveToolBars1.Tools(3).Visible = True
    SSActiveToolBars1.Tools(4).Visible = True
End If

If sw_nuevo_doc = False Then
    SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If wopcion = 0 Then
        Lista_marcas.dxDBGrid1.Dataset.Close
        Lista_marcas.dxDBGrid1.Dataset.Open
        Lista_marcas.dxDBGrid1.Dataset.ADODataset.Requery
    End If
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            nuevo_marca
        Case "ID_Grabar":
            grabar_marca
        
            If sw_mant_ayuda = True Then
                wcodmar = txtCodigo.Text
                wmarca = txtdescrip.Text
                sw_mant_ayuda = False
                Me.Hide
            End If
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
        Case "ID_Eliminar"
            eliminar_marca
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
    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1050
    
    wopcion = 0
    If sw_nuevo_doc = True Then
        nuevo_marca
    Else
        actualizacion_marca Lista_marcas.dxDBGrid1.Columns(0).value
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub actualizacion_marca(cod)
    
    sql = "select * from ef2marcas where f2codmar='" & cod & "'"
    If rsmarcas.State = adStateOpen Then rsmarcas.Close
    rsmarcas.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsmarcas.EOF Then
        txtCodigo.Text = "" & rsmarcas.Fields("f2codmar")
        txtdescrip.Text = rsmarcas.Fields("f2desmar") & ""
        If rsmarcas.Fields("F2ORIGEN") & "" = "N" Then
            opttipo(0).value = True
        Else
            opttipo(1).value = True
        End If
    End If
    
End Sub

Private Sub nuevo_marca()
    genera_cod
    txtCodigo.Enabled = True
    txtCodigo.Text = Codigo
    txtCodigo.Enabled = False
    txtdescrip.Text = ""
    opttipo(0).value = True
End Sub

Private Sub grabar_marca()
On Error GoTo graba
Dim amovs(0 To 2) As a_grabacion

sql = "select * from ef2marcas where f2codmar='" & txtCodigo.Text & "' "
If rsmarcas.State = adStateOpen Then rsmarcas.Close
rsmarcas.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rsmarcas.EOF Then
    sw = 0
Else
    sw = 1
End If
rsmarcas.Close
amovs(0).campo = "f2codmar": amovs(0).valor = txtCodigo.Text: amovs(0).Tipo = "T"
amovs(1).campo = "f2desmar": amovs(1).valor = txtdescrip.Text: amovs(1).Tipo = "T"
amovs(2).campo = "F2ORIGEN": amovs(2).valor = IIf(opttipo(0).value = True, "N", "I"): amovs(2).Tipo = "T"

If sw = 1 Then
    GRABA_REGISTRO_logistica amovs(), "ef2marcas", "A", 2, cnn_dbbancos, ""
    MsgBox "Se actualizó la Marca", vbInformation
Else
    GRABA_REGISTRO_logistica amovs(), "ef2marcas", "M", 2, cnn_dbbancos, "f2codmar='" & txtCodigo.Text & "'"
    MsgBox "Se actualizó la Marca", vbInformation
End If
txtCodigo.Enabled = False
Exit Sub


graba:
    If Err = 3186 Then
        For i% = 1 To 10000
        Next i%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If

End Sub

Private Sub eliminar_marca()
 Beep
    If MsgBox("¿Está seguro de eliminar la Marca...?", 36, "Atención") = 6 Then
        sql = "select f2codmar from ef2marcas where f2codmar='" & txtCodigo.Text & "' "
        If rsmarcas.State = adStateOpen Then rsmarcas.Close
        rsmarcas.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsmarcas.EOF Then
            sql = "DELETE from ef2marcas where f2codmar='" & Trim(txtCodigo.Text) & "'"
            cnn_dbbancos.Execute sql
             'AlmacenaQuery_sql sql, cnn_dbbancos
            txtCodigo.Enabled = True
            nuevo_marca
        Else
            Beep
        End If
        If rsmarcas.State = adStateOpen Then rsmarcas.Close
'        txtcodigo.SetFocus
    End If

End Sub

Private Sub txtcodigo_GotFocus()
txtCodigo.SelStart = 0
txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtdescrip.SetFocus
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    If Len(Trim(txtCodigo.Text)) > 0 Then
        sql = "select f2codmar from ef2marcas where f2codmar='" & txtCodigo.Text & "'"
        If rsmarcas.State = adStateOpen Then rsmarcas.Close
        rsmarcas.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsmarcas.EOF Then
            MsgBox "Código de marca existe. Verifique.", 16, "Atención"
            txtCodigo.SetFocus
        End If
        rsmarcas.Close
    End If
End Sub

Private Sub txtdescrip_Change()
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
End Sub

Private Sub txtdescrip_GotFocus()
txtdescrip.SelStart = 0
txtdescrip.SelLength = Len(txtdescrip.Text)
End Sub

Private Sub genera_cod()
sql = "select f2codmar from ef2marcas order by f2codmar desc"
If rsmarcas.State = adStateOpen Then rsmarcas.Close
rsmarcas.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rsmarcas.EOF Then
    Codigo = rsmarcas.Fields("f2codmar") + 1
    Codigo = Format(Codigo, "000")
Else
    Codigo = 1
    Codigo = Format(Codigo, "000")
End If
rsmarcas.Close
End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        opttipo(0).SetFocus
    End If

End Sub
