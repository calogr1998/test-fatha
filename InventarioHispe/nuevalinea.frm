VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form nuevalinea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Lineas"
   ClientHeight    =   2205
   ClientLeft      =   2490
   ClientTop       =   1920
   ClientWidth     =   5895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5895
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtdescripcion 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   4
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1010
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "nuevalinea.frx":0000
      ToolBars        =   "nuevalinea.frx":7E74
   End
End
Attribute VB_Name = "nuevalinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSnivel As New ADODB.Recordset
Dim Codigo As String
Private Sub Form_Load()
Me.MousePointer = vbHourglass
If sw_nuevo_mant = True Then
    nuevo_nivel
'    txtdescripcion.SetFocus
    nuevalinea.Caption = "Nueva Linea"
Else
    nuevalinea.Caption = "Modificar"
    actualizacion_nivel linea1.dxDBGrid1.Columns.ColumnByFieldName("F7CODCON").value
End If
Me.MousePointer = vbDefault
End Sub

Private Sub actualizacion_nivel(cod)
   sql = "select * from sf7nivel01 where f7codcon ='" & cod & "'"
    If RSnivel.State = adStateOpen Then RSnivel.Close
    RSnivel.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not RSnivel.EOF Then
        txtCodigo.Text = "" & RSnivel.Fields("f7codcon")
        txtDescripcion.Text = "" & RSnivel.Fields("f7descon")
    End If
End Sub

Public Function GeneraCod()
Dim rst As New ADODB.Recordset

sql = "select f7codcon from sf7nivel01 where left(f7codcon,1) between '0' and '9'  order by f7codcon desc"
rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    If Val("" & rst(0).value) = 0 Then
        Codigo = "01"
    Else
        If Val(rst(0).value) > 0 Then
            Codigo = Format(Val(rst(0).value) + 1, "00")
        Else
            Codigo = ""
        End If
    End If
End If
End Function

Public Sub grabar()
If Trim(txtDescripcion.Text) = "" Then
    MsgBox "Debe Ingresar la Descripción", vbInformation, "Sistema de Logística"
    txtDescripcion.Text = ""
    txtDescripcion.SetFocus
    Exit Sub
End If

If sw_nuevo_mant = True Then
    sql = "insert into sf7nivel01 (f7codcon, f7descon) " _
    & "values ('" & txtCodigo.Text & "','" & txtDescripcion.Text & "')"
Else
    sql = "update sf7nivel01 set f7descon='" & txtDescripcion.Text & "' where f7codcon='" & txtCodigo.Text & "'"
End If
cnn_dbbancos.Execute sql
AlmacenaQuery_sql sql, cnn_dbbancos
Actualiza_Log sql, cnn_dbbancos.ConnectionString
MsgBox "Linea de Productos actualizada correctamente"
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
        sw_nuevo_doc = True
        nuevo_nivel
    Case "ID_Grabar"
        grabar
    Case "ID_Eliminar"
        eliminar
    Case "ID_Imprimir":
'        With Acr_Linea
'            .DataControl1.ConnectionString = cnn_dbbancos
'            .DataControl1.Source = "Select * From sf7nivel01 order by f7codcon"
'            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
'            .lblempresa.Caption = wnomcia
'            .Show 1
'        End With
    Case "ID_Lista"
        wopcion = 1
        linea1.dxDBGrid1.Dataset.Close
        linea1.dxDBGrid1.Dataset.Open
        Unload Me
    Case "ID_Salir"
        wopcion = 1
        linea1.dxDBGrid1.Dataset.Close
        linea1.dxDBGrid1.Dataset.Open
        Unload Me
End Select
End Sub

Private Sub nuevo_nivel()
    GeneraCod
    txtCodigo.Text = Codigo
    txtDescripcion.Text = ""
End Sub

Public Sub eliminar()
    xcod = txtCodigo.Text
    xdes = txtDescripcion.Text
    resp = MsgBox("¿Está Seguro de Eliminar la Línea " & xcod & " " & xdes & "?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística")
    If resp = vbYes Then
        sql = "select * from sf7nivel02 where f7nivel01='" & xcod & "'"
        If RSnivel.State = adStateOpen Then RSnivel.Close
        RSnivel.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not RSnivel.EOF Then
            MsgBox "La Línea " & xcod & " " & xdes & " Tiene Registros Asociados. Elimine Primero Estos.", vbInformation, "Sistema de Logística"
            RSnivel.Close
            Exit Sub
        End If
        RSnivel.Close
        
        sql = "delete from sf7nivel01 where f7codcon='" & xcod & "'"
        cnn_dbbancos.Execute sql
        AlmacenaQuery_sql sql, cnn_dbbancos
        nuevo_nivel
    End If

End Sub
