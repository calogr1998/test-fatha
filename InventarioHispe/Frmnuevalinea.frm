VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form nuevalinea 
   ClientHeight    =   2175
   ClientLeft      =   2505
   ClientTop       =   1935
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   ScaleHeight     =   2175
   ScaleWidth      =   6075
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   285
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   0
      Top             =   1215
      Width           =   3855
   End
   Begin ActiveToolBars.SSActiveToolBars atbmenu2 
      Left            =   225
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "Frmnuevalinea.frx":0000
      ToolBars        =   "Frmnuevalinea.frx":1974
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   675
      TabIndex        =   2
      Top             =   675
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   675
      TabIndex        =   1
      Top             =   1260
      Width           =   840
   End
End
Attribute VB_Name = "nuevalinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSnivel As New ADODB.Recordset
Private Sub Form_Load()

If sw_nuevo_mant = True Then
txtcodigo.Text = GeneraCodigo
txtdescripcion.Text = ""
Frmnuevalinea.Caption = "Nueva Linea"
Else
Frmnuevalinea.Caption = "Modificar"

        actualizacion_nivel linea1.dxDBGrid1.Columns.ColumnByFieldName("F7CODCON").Value
End If
End Sub

Private Sub actualizacion_nivel(cod)

   SQL = "select * from sf7nivel01 where f7codcon ='" & cod & "'"
    If RSnivel.State = adStateOpen Then RSnivel.Close
    RSnivel.Open SQL, cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not RSnivel.EOF Then
        txtcodigo.Text = "" & RSnivel.Fields("f7codcon")
        txtdescripcion.Text = "" & RSnivel.Fields("f7descon")
    End If

End Sub

Private Sub atbmenu2_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "idgrabar"
        grabar
    Case "idsalir"
        Unload Me
End Select
End Sub

Public Function GeneraCodigo()
Dim rst As New ADODB.Recordset

SQL = "select max(f7codcon) from sf7nivel01"
rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    If Val("" & rst(0).Value) = 0 Then
        GeneraCodigo = "01"
    Else
        If Val(rst(0).Value) > 0 Then
            GeneraCodigo = Val(rst(0).Value) + 1
        Else
            GeneraCodigo = ""
        End If
    End If
End If
End Function

Public Sub grabar()
If Trim(txtdescripcion.Text) = "" Then
    MsgBox "Debe Ingresar la Descripción", vbInformation, "Sistema de Logística"
    txtdescripcion.Text = ""
    txtdescripcion.SetFocus
    Exit Sub
End If

If sw_nuevo_mant = True Then
    SQL = "insert into sf7nivel01 (f7codcon, f7descon) " _
    & "values ('" & txtcodigo.Text & "','" & txtdescripcion.Text & "')"
Else
    SQL = "update sf7nivel01 set f7descon='" & txtdescripcion.Text & "' where f7codcon='" & txtcodigo.Text & "'"
End If
cnn_dbbancos.Execute SQL
Unload Me
linea1.dxDBGrid1.Dataset.Close
linea1.dxDBGrid1.Dataset.Open
End Sub

