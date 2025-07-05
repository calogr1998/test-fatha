VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form mant_conceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Conceptos"
   ClientHeight    =   4890
   ClientLeft      =   4260
   ClientTop       =   2460
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6615
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6495
      Begin Threed.SSCheck SSCheck3 
         Height          =   285
         Left            =   315
         TabIndex        =   13
         Top             =   1350
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Incluir almacen de destino"
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
      Begin VB.Frame Frame4 
         Caption         =   "Almacen Destino"
         Height          =   870
         Left            =   270
         TabIndex        =   10
         Top             =   1755
         Width           =   6000
         Begin VB.ComboBox cmbalmacen 
            Height          =   315
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   315
            Width           =   3120
         End
         Begin Threed.SSCheck SSCheck2 
            Height          =   195
            Left            =   3825
            TabIndex        =   12
            Top             =   405
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Todos"
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
      Begin VB.Frame Frame3 
         Height          =   540
         Left            =   270
         TabIndex        =   8
         Top             =   3375
         Width           =   6000
         Begin Threed.SSCheck SSCheck1 
            Height          =   195
            Left            =   225
            TabIndex        =   9
            Top             =   225
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Costo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
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
         TabIndex        =   5
         Top             =   855
         Width           =   4785
      End
      Begin VB.TextBox txtcodigo 
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
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         Height          =   780
         Left            =   270
         TabIndex        =   1
         Top             =   2610
         Width           =   6000
         Begin Threed.SSOption opttipo 
            Height          =   285
            Index           =   0
            Left            =   855
            TabIndex        =   2
            Top             =   315
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Ingreso"
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
            Top             =   315
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Salida"
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
         Left            =   315
         TabIndex        =   7
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
         Left            =   315
         TabIndex        =   6
         Top             =   540
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   45
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_conceptos.frx":0000
      ToolBars        =   "mant_conceptos.frx":7E74
   End
End
Attribute VB_Name = "mant_conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Codigo          As String
Dim rsconceptos        As New ADODB.Recordset
Dim sql             As String
Dim wopcion         As Byte

Private Sub Command1_Click()
    Almacenes_Conceptos.Show 1
End Sub

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

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.Width = 6870
    Me.left = 1600
    Me.top = 1050
    
    wopcion = 0
    
    habilita_combo
    
    If sw_nuevo_doc = True Then
        txtcodigo.Enabled = True
        nuevo_concepto
        SSCheck3.value = False
        cmbAlmacen.ListIndex = -1
        cmbAlmacen.Enabled = False
        SSCheck2.value = False
        SSCheck2.Enabled = False
        txtcodigo.TabIndex = 0
    Else
        actualizacion_concepto lista_conceptos.dxDBGrid1.Columns(0).value
        txtcodigo.Enabled = False
    End If
    If lista_conceptos.dxDBGrid1.Columns.ColumnByFieldName("f1tipmov").value = "I" Then
        cmbAlmacen.Enabled = False
    End If
    
    If sw_mant_ayuda = True Then
        If wtipmov = "I" Then
            opttipo(0).value = True
            opttipo(0).Enabled = True
            opttipo(1).Enabled = False
        Else
            opttipo(1).value = True
            opttipo(1).Enabled = True
            opttipo(0).Enabled = False
        End If
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub nuevo_concepto()

    txtcodigo.Text = Codigo
    txtdescrip.Text = ""
    opttipo(0).value = True
    SSCheck2.value = False
    SSCheck3.value = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If wopcion = 0 Then
    With Lista_marcas.dxDBGrid1.Dataset
      .Close
      .Open
      .ADODataset.Requery
    End With
  End If
End Sub

Private Sub opttipo_Click(Index As Integer, value As Integer)
    If opttipo(0).value = True Then
        SSCheck1.Enabled = True
    Else
        SSCheck1.Enabled = False
        SSCheck1.value = False
    End If
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
  Select Case Tool.Id
      Case "ID_Nuevo"
          sw_nuevo_doc = True
          txtcodigo.Enabled = True
          nuevo_concepto
      Case "ID_Grabar":
          grabar_concepto
          If sw_mant_ayuda = True Then
            wconcepto = txtcodigo.Text
            wnomconcepto = txtdescrip.Text
            sw_mant_ayuda = False
            Me.Hide
          End If
          
      Case "ID_Eliminar"
          eliminar_concepto
'      Case "ID_Imprimir":
'        With Acr_Conceptos
'            .DataControl1.ConnectionString = cnn_dbbancos
'            .DataControl1.Source = "Select f1codori,f1nomori,IIF(f1tipmov='S','Salida','Ingreso') as tipo From sf1origenes order by f1tipmov"
'            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
'            .lblempresa.Caption = wnomcia
'            .Show 1
'        End With

      Case "ID_Lista"
          wopcion = 1
'          lista_conceptos.dxDBGrid1.Dataset.Close
'          lista_conceptos.dxDBGrid1.Dataset.Open
          Unload Me
      Case "ID_Salir"
          wopcion = 1
'          lista_conceptos.dxDBGrid1.Dataset.Close
'          lista_conceptos.dxDBGrid1.Dataset.Open
          Unload Me
  End Select

End Sub

Private Sub grabar_concepto()
On Error GoTo graba
Dim amovs2(0 To 4) As a_grabacion

sql = "select * from sf1origenes where f1codori='" & txtcodigo.Text & "' "
If rsconceptos.State = adStateOpen Then rsconceptos.Close
rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rsconceptos.EOF Then
    sw = 0
Else
    sw = 1
End If
rsconceptos.Close
amovs2(0).campo = "f1codori": amovs2(0).valor = txtcodigo.Text: amovs2(0).Tipo = "T"
amovs2(1).campo = "f1nomori": amovs2(1).valor = txtdescrip.Text: amovs2(1).Tipo = "T"
amovs2(2).campo = "f1tipmov": amovs2(2).valor = IIf(opttipo(0).value = True, "I", "S"): amovs2(2).Tipo = "T"
amovs2(3).campo = "f1costo": amovs2(3).valor = IIf(SSCheck1.value = True, "*", ""): amovs2(3).Tipo = "T"
amovs2(4).campo = "codalmdes": amovs2(4).valor = IIf(SSCheck2.value = True, "99", right(cmbAlmacen.Text, 2)): amovs2(4).Tipo = "T"
    If sw = 1 Then
        GRABA_REGISTRO_logistica amovs2(), "sf1origenes", "A", 4, cnn_dbbancos, ""
        asocia_almacenes txtcodigo.Text
    Else
        GRABA_REGISTRO_logistica amovs2(), "sf1origenes", "M", 4, cnn_dbbancos, "f1codori='" & txtcodigo.Text & "'"
    End If
    ' SE DESACTIVA POR DUPLICAR LA TABLA "ALMACEN_CONCEPTOS"
'If sw_mant_ayuda = True Then
'    cnn_dbbancos.Execute "insert into almacen_concepto (f2codalm,f1codori) values ('" & wcod_alm & "','" & txtcodigo.Text & "')"
'    txtcodigo.Enabled = False
'End If
    
If sw = 1 Then
    MsgBox "Se a creado el nuevo concepto", vbInformation, "Sistema de Logistica"
Else
    MsgBox "El Concepto se Actualizó", vbInformation, "Sistema de Logistica"
End If
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

Private Sub actualizacion_concepto(cod)
    
    sql = "select * from sf1origenes where f1codori='" & cod & "'"
    If rsconceptos.State = adStateOpen Then rsconceptos.Close
    rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsconceptos.EOF Then
        txtcodigo.Text = "" & rsconceptos.Fields("f1codori")
        txtdescrip.Text = rsconceptos.Fields("f1nomori") & ""
        If rsconceptos.Fields("f1tipmov") & "" = "I" Then
            opttipo(0).value = True
        Else
            opttipo(1).value = True
        End If
        
        If rsconceptos.Fields("codalmdes").value = "99" Then
            SSCheck2.value = True
        Else
            SSCheck2.value = False
            For i = 0 To cmbAlmacen.ListCount - 1
                If rsconceptos.Fields("codalmdes").value = right(cmbAlmacen.List(i), 2) Then
                    cmbAlmacen.ListIndex = i
                    Exit For
                End If
            Next
        End If
        
        If cmbAlmacen.Text = "" Then
            If SSCheck2.value = True Then
                SSCheck3.value = True
                SSCheck3.Enabled = True
                cmbAlmacen.Enabled = False
                cmbAlmacen.ListIndex = -1
            Else
                SSCheck3.value = False
                SSCheck2.Enabled = False
                cmbAlmacen.Enabled = False
                'SSCheck3.Enabled = False
            End If
        Else
            SSCheck3.value = True
            cmbAlmacen.Enabled = True
            SSCheck2.Enabled = True
        End If
        
        If rsconceptos.Fields("f1costo") = "*" Then
            SSCheck1.value = True
        End If
    End If
End Sub

Private Sub eliminar_concepto()
 Beep
    If MsgBox("¿Está seguro de eliminar el Concepto...?", 36, "Atención") = 6 Then
        sql = "select f1codori from sf1origenes where f1codori='" & txtcodigo.Text & "' "
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
        rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconceptos.EOF Then
            If ctipoadm_bd = "M" Then
            csql = "delete from almacen_concepto where f1codori='" & txtcodigo.Text & "'"
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                
                csql = "DELETE from sf1origenes where f1codori='" & txtcodigo.Text & "' "
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                
            Else
                csql = "delete * from almacen_concepto where f1codori='" & txtcodigo.Text & "'"
                cnn_dbbancos.Execute csql
                 'AlmacenaQuery_sql sql, cnn_dbbancos
                 
                csql = "DELETE * from sf1origenes where f1codori='" & txtcodigo.Text & "' "
                cnn_dbbancos.Execute csql
                 'AlmacenaQuery_sql sql, cnn_dbbancos
            End If
            
            txtcodigo.Enabled = True
            nuevo_concepto
        Else
            Beep
        End If
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
        txtcodigo.SetFocus
    End If

End Sub

Private Sub SSCheck2_Click(value As Integer)
    If SSCheck2.value = True Then
      cmbAlmacen.Enabled = False
      cmbAlmacen.ListIndex = -1
    Else
      cmbAlmacen.Enabled = True
    End If

End Sub

Private Sub SSCheck3_Click(value As Integer)
    If SSCheck3.value = False Then
        cmbAlmacen.ListIndex = -1
        cmbAlmacen.Enabled = False
        SSCheck2.value = False
        SSCheck2.Enabled = False
    Else
        cmbAlmacen.Enabled = True
        SSCheck2.Enabled = True
    End If
End Sub

Private Sub txtcodigo_GotFocus()
txtcodigo.SelStart = 0
txtcodigo.SelLength = Len(txtcodigo.Text)
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtdescrip.SetFocus
    End If

End Sub

Private Sub txtcodigo_LostFocus()

    If Len(Trim(txtcodigo.Text)) > 0 Then
        sql = "select f1codori from sf1origenes where f1codori='" & txtcodigo.Text & "'"
        If rsconceptos.State = adStateOpen Then rsconceptos.Close
        rsconceptos.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconceptos.EOF Then
            MsgBox "Código de concepto existe. Verifique.", vbInformation, "Atención"
            txtcodigo.SetFocus
        End If
        rsconceptos.Close
    End If

End Sub

Private Sub txtdescrip_GotFocus()
txtdescrip.SelStart = 0
txtdescrip.SelLength = Len(txtdescrip.Text)
End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       ' opttipo(0).SetFocus
    End If

End Sub

Private Sub habilita_combo()
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select f2codalm,f2nomalm from ef2almacenes order by f2nomalm asc", cnn_dbbancos, adOpenStatic, adLockReadOnly
    X = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            cmbAlmacen.AddItem Rs.Fields("f2nomalm") & "" & Space(75) & Rs.Fields("F2CODALM") & ""
            Rs.MoveNext
            X = X + 1
        Loop
    End If
    Rs.Close
    
    If X = 1 Then
        cmbAlmacen.ListIndex = 0
    End If
End Sub

Private Sub asocia_almacenes(Codigo As String)
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select f2codalm from ef2almacenes", cnn_dbbancos, adOpenStatic, adLockReadOnly
    X = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            sql = "INSERT INTO ALMACEN_CONCEPTO (F2CODALM,F1CODORI) VALUES ('" _
                & Rs.Fields("F2CODALM") & "','" & Codigo & "')"
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Rs.MoveNext
        Loop
    End If
    Rs.Close
End Sub
