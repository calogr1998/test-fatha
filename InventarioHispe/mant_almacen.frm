VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form mant_almacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Almacenes"
   ClientHeight    =   4020
   ClientLeft      =   4515
   ClientTop       =   3735
   ClientWidth     =   8235
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8235
   Begin VB.Frame Frame1 
      Height          =   3705
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.CommandButton Command1 
         Caption         =   "Asociar Conceptos"
         Height          =   330
         Left            =   225
         TabIndex        =   18
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         Height          =   825
         Left            =   225
         TabIndex        =   12
         Top             =   2295
         Width           =   7800
         Begin VB.OptionButton opttipo 
            Caption         =   "Obra Civil"
            Height          =   210
            Index           =   4
            Left            =   6165
            TabIndex        =   17
            Top             =   405
            Width           =   1455
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Repuestos"
            Height          =   210
            Index           =   3
            Left            =   4815
            TabIndex        =   16
            Top             =   405
            Width           =   1095
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Reservados"
            Height          =   210
            Index           =   2
            Left            =   3420
            TabIndex        =   15
            Top             =   405
            Width           =   1185
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Muestras/Demostraciones"
            Height          =   210
            Index           =   1
            Left            =   1035
            TabIndex        =   14
            Top             =   405
            Width           =   2220
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Venta"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   13
            Top             =   405
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtcentro 
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox txtdireccion 
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   1575
         Width           =   5370
      End
      Begin VB.TextBox txtruc 
         Height          =   285
         Left            =   1530
         MaxLength       =   11
         TabIndex        =   3
         Top             =   1215
         Width           =   1590
      End
      Begin VB.TextBox txtalmacen 
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   855
         Width           =   5370
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1530
         TabIndex        =   1
         Top             =   495
         Width           =   690
      End
      Begin Threed.SSPanel pnlcentro 
         Height          =   285
         Left            =   2340
         TabIndex        =   6
         Top             =   1935
         Width           =   4560
         _Version        =   65536
         _ExtentX        =   8043
         _ExtentY        =   503
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo"
         Height          =   210
         Left            =   360
         TabIndex        =   11
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   210
         Left            =   360
         TabIndex        =   7
         Top             =   540
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_almacen.frx":0000
      ToolBars        =   "mant_almacen.frx":7E74
   End
End
Attribute VB_Name = "mant_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Codigo      As String
Dim rstemphlpc  As New ADODB.Recordset
Dim sw_hlp      As Boolean
Dim sql         As String


Private Sub Command1_Click()
    conceptos_almacenes.Show 1
End Sub

Private Sub Form_Activate()
  
  If sw_mant_ayuda = True Then
    SSActiveToolBars1.Tools(3).Visible = False
    SSActiveToolBars1.Tools(4).Visible = False
  Else
    SSActiveToolBars1.Tools(3).Visible = True
    SSActiveToolBars1.Tools(4).Visible = True
  End If

End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass

    Me.left = 1500
    Me.top = 980
    
    If sw_nuevo_doc = True Then
        nuevo_almacen
    Else
        actualizacion_almacen lista_almacen.dxDBGrid1.Columns(0).value
    End If
    
    sw_hlp = False
    Me.MousePointer = vbDefault
End Sub

Private Sub nuevo_almacen()
    
    txtCodigo.Enabled = True
    genera_cod
    txtCodigo.Text = Codigo
    wcod_alm = Codigo
    txtAlmacen.Text = ""
    txtruc.Text = ""
    TxtDireccion.Text = ""
    txtcentro.Text = ""
    pnlcentro.Caption = ""
    opttipo(0).value = True
    
End Sub

Private Sub actualizacion_almacen(cod)
    
    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        sql = "SELECT * FROM MAESTROS.EF2ALMACENES WHERE F2CODALM = '" & cod & "'"
        rsalmacen.Open sql, cnBdCPlus, adOpenDynamic, adLockOptimistic
    Else
        sql = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & cod & "'"
        rsalmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    End If
    
    If Not rsalmacen.EOF Then
        txtCodigo.Text = "" & rsalmacen.Fields("f2codalm")
        txtAlmacen.Text = "" & rsalmacen.Fields("F2NOMALM")
        txtruc.Text = "" & rsalmacen.Fields("F2RUCALM")
        TxtDireccion.Text = "" & rsalmacen.Fields("F2DIRALM")
        txtcentro.Text = "" & rsalmacen.Fields("F4CENTRO")
        
        If rstemphlpc.State = adStateOpen Then rstemphlpc.Close
        
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            sql = "SELECT F3COSTO, F3DESCRIP FROM MAESTROS.CENTROS WHERE F3COSTO = '" & txtcentro.Text & "'"
            rstemphlpc.Open sql, cnBdCPlus, adOpenDynamic, adLockOptimistic
        Else
            sql = "SELECT F3COSTO, F3DESCRIP FROM CENTROS WHERE F3COSTO = '" & txtcentro.Text & "'"
            rstemphlpc.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        End If
        
        If Not rstemphlpc.EOF Then
            pnlcentro.Caption = "" & rstemphlpc.Fields("F3DESCRIP")
        End If
        
        rstemphlpc.Close
        
        If Trim("" & rsalmacen.Fields("F2TIPO")) = "0" Or Len(Trim("" & rsalmacen.Fields("F2TIPO"))) = 0 Then opttipo(0).value = True
        If Trim("" & rsalmacen.Fields("F2TIPO")) = "1" Then opttipo(1).value = True
        If Trim("" & rsalmacen.Fields("F2TIPO")) = "2" Then opttipo(2).value = True
        If Trim("" & rsalmacen.Fields("F2TIPO")) = "3" Then opttipo(3).value = True
        If Trim("" & rsalmacen.Fields("F2TIPO")) = "4" Then opttipo(4).value = True
    End If
    
    rsalmacen.Close

End Sub

Private Sub genera_cod()

    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        sql = "select F2CODALM from MAESTROS.ef2almacenes order by f2codalm desc"
        If rsalmacen.State = adStateOpen Then rsalmacen.Close
        rsalmacen.Open sql, cnBdCPlus, adOpenDynamic, adLockOptimistic
        If Not rsalmacen.EOF Then
            Codigo = rsalmacen.Fields("f2codalm") + 1
            Codigo = Format(Codigo, "00")
        Else
            Codigo = 1
            Codigo = Format(Codigo, "00")
        End If
    Else
        sql = "select F2CODALM from ef2almacenes order by f2codalm desc"
        If rsalmacen.State = adStateOpen Then rsalmacen.Close
        rsalmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsalmacen.EOF Then
            Codigo = rsalmacen.Fields("f2codalm") + 1
            Codigo = Format(Codigo, "00")
        Else
            Codigo = 1
            Codigo = Format(Codigo, "00")
        End If
    End If
    
End Sub

Private Sub eliminar_almacen()

    Beep
    If MsgBox("Está seguro de eliminar el almacén.", 36, "Atención") = 6 Then
    
    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        sql = "select f2codalm from MAESTROS.ef2almacenes where f2codalm = '" & txtCodigo.Text & "'"
        rsalmacen.Open sql, cnBdCPlus, adOpenDynamic, adLockOptimistic
    Else
        sql = "select f2codalm from ef2almacenes where f2codalm='" & txtCodigo.Text & "'"
        rsalmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    End If
    
    If Not rsalmacen.EOF Then
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            sql = "DELETE from MAESTROS.ef2almacenes where f2codalm='" & txtCodigo.Text & "'"
            cnBdCPlus.Execute sql
            txtCodigo.Enabled = True
            nuevo_almacen
        Else
            sql = "DELETE * from ef2almacenes where f2codalm='" & txtCodigo.Text & "'"
            cnn_dbbancos.Execute sql
            txtCodigo.Enabled = True
            nuevo_almacen
        End If
    Else
        Beep
    End If
    
    rsalmacen.Close
    txtCodigo.SetFocus
    End If

End Sub

Private Sub grabar_almacen()
On Error GoTo graba
Dim ctipoalm        As String
Dim amovs(0 To 31)  As a_grabacion
    
    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        sql = "SELECT * FROM MAESTROS.EF2ALMACENES WHERE F2CODALM = '" & txtCodigo.Text & "'"
        rsalmacen.Open sql, cnBdCPlus, adOpenDynamic, adLockOptimistic
    Else
        sql = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & txtCodigo.Text & "'"
        rsalmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    End If
    
    If Not rsalmacen.EOF Then
        sw = 0
    Else
        sw = 1
    End If
    
    amovs(0).campo = "f2codalm": amovs(0).valor = txtCodigo.Text: amovs(0).Tipo = "T"
    amovs(1).campo = "F2NOMALM": amovs(1).valor = txtAlmacen.Text: amovs(1).Tipo = "T"
    amovs(2).campo = "F2RUCALM": amovs(2).valor = txtruc.Text: amovs(2).Tipo = "T"
    amovs(3).campo = "F2DIRALM": amovs(3).valor = TxtDireccion.Text: amovs(3).Tipo = "T"
    amovs(4).campo = "F4CENTRO": amovs(4).valor = txtcentro.Text: amovs(4).Tipo = "T"
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
    Else
        If sw = 1 Then
            amovs(5).campo = "F1VALING01": amovs(5).valor = "I-010000": amovs(5).Tipo = "T"
            amovs(6).campo = "F1VALING02": amovs(6).valor = "I-020000": amovs(6).Tipo = "T"
            amovs(7).campo = "F1VALING03": amovs(7).valor = "I-030000": amovs(7).Tipo = "T"
            amovs(8).campo = "F1VALING04": amovs(8).valor = "I-040000": amovs(8).Tipo = "T"
            amovs(9).campo = "F1VALING05": amovs(9).valor = "I-050000": amovs(9).Tipo = "T"
            amovs(10).campo = "F1VALING06": amovs(10).valor = "I-060000": amovs(10).Tipo = "T"
            amovs(11).campo = "F1VALING07": amovs(11).valor = "I-070000": amovs(11).Tipo = "T"
            amovs(12).campo = "F1VALING08": amovs(12).valor = "I-080000": amovs(12).Tipo = "T"
            amovs(13).campo = "F1VALING09": amovs(13).valor = "I-090000": amovs(13).Tipo = "T"
            amovs(14).campo = "F1VALING10": amovs(14).valor = "I-100000": amovs(14).Tipo = "T"
            amovs(15).campo = "F1VALING11": amovs(15).valor = "I-110000": amovs(15).Tipo = "T"
            amovs(16).campo = "F1VALING12": amovs(16).valor = "I-120000": amovs(16).Tipo = "T"
            amovs(17).campo = "F1VALSAL01": amovs(17).valor = "S-010000": amovs(17).Tipo = "T"
            amovs(18).campo = "F1VALSAL02": amovs(18).valor = "S-020000": amovs(18).Tipo = "T"
            amovs(19).campo = "F1VALSAL03": amovs(19).valor = "S-030000": amovs(19).Tipo = "T"
            amovs(20).campo = "F1VALSAL04": amovs(20).valor = "S-040000": amovs(20).Tipo = "T"
            amovs(21).campo = "F1VALSAL05": amovs(21).valor = "S-050000": amovs(21).Tipo = "T"
            amovs(22).campo = "F1VALSAL06": amovs(22).valor = "S-060000": amovs(22).Tipo = "T"
            amovs(23).campo = "F1VALSAL07": amovs(23).valor = "S-070000": amovs(23).Tipo = "T"
            amovs(24).campo = "F1VALSAL08": amovs(24).valor = "S-080000": amovs(24).Tipo = "T"
            amovs(25).campo = "F1VALSAL09": amovs(25).valor = "S-090000": amovs(25).Tipo = "T"
            amovs(26).campo = "F1VALSAL10": amovs(26).valor = "S-100000": amovs(26).Tipo = "T"
            amovs(27).campo = "F1VALSAL11": amovs(27).valor = "S-110000": amovs(27).Tipo = "T"
            amovs(28).campo = "F1VALSAL12": amovs(28).valor = "S-120000": amovs(28).Tipo = "T"
            amovs(29).campo = "F1ULTINV": amovs(29).valor = Format(Date, "DD/MM/YYYY"): amovs(29).Tipo = "T"
            amovs(30).campo = "F1ULTSAL": amovs(30).valor = Format(Date, "DD/MM/YYYY"): amovs(30).Tipo = "T"
        Else
            amovs(5).campo = "F1VALING01": amovs(5).valor = rsalmacen.Fields("F1VALING01"): amovs(5).Tipo = "T"
            amovs(6).campo = "F1VALING02": amovs(6).valor = rsalmacen.Fields("F1VALING02"): amovs(6).Tipo = "T"
            amovs(7).campo = "F1VALING03": amovs(7).valor = rsalmacen.Fields("F1VALING03"): amovs(7).Tipo = "T"
            amovs(8).campo = "F1VALING04": amovs(8).valor = rsalmacen.Fields("F1VALING04"): amovs(8).Tipo = "T"
            amovs(9).campo = "F1VALING05": amovs(9).valor = rsalmacen.Fields("F1VALING05"): amovs(9).Tipo = "T"
            amovs(10).campo = "F1VALING06": amovs(10).valor = rsalmacen.Fields("F1VALING06"): amovs(10).Tipo = "T"
            amovs(11).campo = "F1VALING07": amovs(11).valor = rsalmacen.Fields("F1VALING07"): amovs(11).Tipo = "T"
            amovs(12).campo = "F1VALING08": amovs(12).valor = rsalmacen.Fields("F1VALING08"): amovs(12).Tipo = "T"
            amovs(13).campo = "F1VALING09": amovs(13).valor = rsalmacen.Fields("F1VALING09"): amovs(13).Tipo = "T"
            amovs(14).campo = "F1VALING10": amovs(14).valor = rsalmacen.Fields("F1VALING10"): amovs(14).Tipo = "T"
            amovs(15).campo = "F1VALING11": amovs(15).valor = rsalmacen.Fields("F1VALING11"): amovs(15).Tipo = "T"
            amovs(16).campo = "F1VALING12": amovs(16).valor = rsalmacen.Fields("F1VALING12"): amovs(16).Tipo = "T"
            amovs(17).campo = "F1VALSAL01": amovs(17).valor = rsalmacen.Fields("F1VALSAL01"): amovs(17).Tipo = "T"
            amovs(18).campo = "F1VALSAL02": amovs(18).valor = rsalmacen.Fields("F1VALSAL02"): amovs(18).Tipo = "T"
            amovs(19).campo = "F1VALSAL03": amovs(19).valor = rsalmacen.Fields("F1VALSAL03"): amovs(19).Tipo = "T"
            amovs(20).campo = "F1VALSAL04": amovs(20).valor = rsalmacen.Fields("F1VALSAL04"): amovs(20).Tipo = "T"
            amovs(21).campo = "F1VALSAL05": amovs(21).valor = rsalmacen.Fields("F1VALSAL05"): amovs(21).Tipo = "T"
            amovs(22).campo = "F1VALSAL06": amovs(22).valor = rsalmacen.Fields("F1VALSAL06"): amovs(22).Tipo = "T"
            amovs(23).campo = "F1VALSAL07": amovs(23).valor = rsalmacen.Fields("F1VALSAL07"): amovs(23).Tipo = "T"
            amovs(24).campo = "F1VALSAL08": amovs(24).valor = rsalmacen.Fields("F1VALSAL08"): amovs(24).Tipo = "T"
            amovs(25).campo = "F1VALSAL09": amovs(25).valor = rsalmacen.Fields("F1VALSAL09"): amovs(25).Tipo = "T"
            amovs(26).campo = "F1VALSAL10": amovs(26).valor = rsalmacen.Fields("F1VALSAL10"): amovs(26).Tipo = "T"
            amovs(27).campo = "F1VALSAL11": amovs(27).valor = rsalmacen.Fields("F1VALSAL11"): amovs(27).Tipo = "T"
            amovs(28).campo = "F1VALSAL12": amovs(28).valor = rsalmacen.Fields("F1VALSAL12"): amovs(28).Tipo = "T"
            amovs(29).campo = "F1ULTINV": amovs(29).valor = rsalmacen.Fields("F1ULTINV"): amovs(29).Tipo = "F"
            amovs(30).campo = "F1ULTSAL": amovs(30).valor = rsalmacen.Fields("F1ULTSAL"): amovs(30).Tipo = "F"
        End If
    End If
    
    rsalmacen.Close
    Set rsalmacen = Nothing
    
    ctipoalm = ""
    If opttipo(0).value = True Then ctipoalm = "0"
    If opttipo(1).value = True Then ctipoalm = "1"
    If opttipo(2).value = True Then ctipoalm = "2"
    If opttipo(3).value = True Then ctipoalm = "3"
    If opttipo(4).value = True Then ctipoalm = "4"
    
    amovs(31).campo = "F2TIPO": amovs(31).valor = ctipoalm: amovs(31).Tipo = "T"
    
    If sw = 1 Then
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            GRABA_REGISTRO_logistica amovs(), "MAESTROS.EF2ALMACENES", "A", 4, cnBdCPlus, ""
        Else
            GRABA_REGISTRO_logistica amovs(), "EF2ALMACENES", "A", 31, cnn_dbbancos, ""
        End If
        
        MsgBox "Se a creado un nuevo Almacén", vbInformation
    Else
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            GRABA_REGISTRO_logistica amovs(), "MAESTROS.EF2ALMACENES", "M", 5, cnBdCPlus, "f2codalm = '" & txtCodigo.Text & "'"
        Else
            GRABA_REGISTRO_logistica amovs(), "EF2ALMACENES", "M", 31, cnn_dbbancos, "f2codalm='" & txtCodigo.Text & "'"
        End If
        
        MsgBox "Se actualizó correctamente el Almacén", vbInformation
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

Private Sub Form_Unload(Cancel As Integer)
'lista_almacen.adoctasctes.Refresh
End Sub

Private Sub txtcentro_DblClick()

    txtcentro_KeyDown 113, 0

End Sub

Private Sub txtcodigo_GotFocus()

    txtCodigo.SelStart = 0: txtCodigo.SelLength = Len(txtCodigo.Text)
    
End Sub

Private Sub txtcentro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_hlp = True
        Ayuda_Centros.Show 1
        If Len(Trim(wcodcosto)) > 0 Then
            txtcentro.Text = wcodcosto
            pnlcentro.Caption = wdescosto
        End If
    End If
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAlmacen.SetFocus
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtAlmacen.SelStart = 0: txtAlmacen.SelLength = Len(txtAlmacen.Text)
    
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtruc.SetFocus
    End If
    
End Sub

Private Sub txtcentro_LostFocus()

    If sw_hlp = False Then
        If Len(Trim(txtcentro.Text)) > 0 Then
            sql = "select F3COSTO,F3DESCRIP from CENTROS where F3COSTO='" & txtcentro.Text & "'"
            If rstemphlpc.State = adStateOpen Then rstemphlpc.Close
            rstemphlpc.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rstemphlpc.EOF Then
                pnlcentro.Caption = rstemphlpc.Fields("F3DESCRIP") & ""
            Else
                MsgBox "Ingrese codigo válido.", 16, "Atención"
                txtcentro.Text = ""
                pnlcentro.Caption = ""
                txtcentro.SetFocus
                Exit Sub
            End If
            rstemphlpc.Close
        End If
    End If
    
End Sub

Private Sub txtruc_GotFocus()

    txtruc.SelStart = 0: txtruc.SelLength = Len(txtruc.Text)
    
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtDireccion.SetFocus
    End If
    
End Sub

Private Sub txtdireccion_GotFocus()

    TxtDireccion.SelStart = 0: TxtDireccion.SelLength = Len(TxtDireccion.Text)
    
End Sub

Private Sub txtdireccion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtcentro.SetFocus
    End If
    
End Sub

Private Sub txtcentro_GotFocus()

    txtcentro.SelStart = 0: txtcentro.SelLength = Len(txtcentro.Text)
    
End Sub

Private Sub txtcentro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtAlmacen.SetFocus
    End If
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            nuevo_almacen
            txtCodigo.SetFocus
        Case "ID_Grabar"
            Me.MousePointer = vbHourglass
            grabar_almacen
            Me.MousePointer = vbDefault
            If sw_mant_ayuda = True Then
              wcod_alm = txtCodigo.Text
              wnomalmacen = txtAlmacen.Text
              sw_mant_ayuda = False
              Me.Hide
            End If
        Case "ID_Eliminar"
            eliminar_almacen
'        Case "ID_Imprimir":
'               With Acr_Almacen
'                    .DataControl1.ConnectionString = cnn_dbbancos
'                    .DataControl1.Source = "select * from ef2almacenes order by f2codalm"
'                    .fldfecha.Text = Format(Date, "DD/MM/YYYY")
'                    .lblempresa.Caption = wempresa
'                    .Show 1
'                End With
        Case "ID_Lista"
'            lista_almacen.adoctasctes.Refresh
            Unload Me
        Case "ID_Salir"
           ' lista_almacen.adoctasctes.Refresh
            Unload Me
    End Select
    
End Sub
