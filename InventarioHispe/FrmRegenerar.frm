VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmRegenerar 
   Caption         =   "Regenerar Saldos"
   ClientHeight    =   1710
   ClientLeft      =   2445
   ClientTop       =   2565
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   1710
   ScaleWidth      =   4935
   Begin Threed.SSCommand BtnExit 
      Height          =   345
      Left            =   2475
      TabIndex        =   0
      Top             =   1170
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand BtnSaldos 
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   1170
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   614
      _StockProps     =   78
      Caption         =   "&Iniciar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   450
      Left            =   135
      TabIndex        =   2
      Top             =   510
      Width           =   4590
      _Version        =   65536
      _ExtentX        =   8096
      _ExtentY        =   794
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin ComctlLib.ProgressBar gauge 
         Height          =   255
         Left            =   195
         TabIndex        =   3
         Top             =   90
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   45
      End
   End
   Begin VB.Label LabelRe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Regenerando ...."
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   1470
   End
End
Attribute VB_Name = "FrmRegenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim wcodalm     As String
Dim wcodpro     As String
Dim wval        As Double
Dim PROMEDIO    As Integer
Dim amovs_cab(0 To 8)  As a_grabacion
Dim Arreglo(0 To 5)     As a_grabacion
Dim amovs_det(0 To 3)  As a_grabacion

Private Sub Actualiza_Costos()
On Error GoTo Depura
Dim wsaldoc         As Double
Dim wsaldos, wvalsol, SalPep   As Double
Dim wsaldod, wvaldol   As Double
Dim ValVta, ValDol, Totite, TotDol As Double
Dim wcanpro, wtemp      As Double
Dim wultinv             As Variant
Dim wcodigo, SQL        As String
Dim rsconsulta          As ADODB.Recordset
    
    Set RsStockDet = New ADODB.Recordset
    Set RsAlmacenes = New ADODB.Recordset
    Set RsMovAlmacen = New ADODB.Recordset
    Set rsconsulta = New ADODB.Recordset
    Set RsProducto = New ADODB.Recordset
    
    
    SQL = "SELECT * FROM IF3VALES"
    
    If RsStockDet.State = adStateOpen Then RsStockDet.Close
    RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockDet.EOF Then
        RsStockDet.MoveFirst
        Do While Not RsStockDet.EOF
            wcodalm = Trim(RsStockDet.Fields("F2codalm"))
            SQL = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & wcodalm & "'"
            If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
            RsAlmacenes.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsAlmacenes.EOF Then wultinv = CVDate(RsAlmacenes.Fields("F1ultinv"))
            RsAlmacenes.MoveFirst
            
            Do While Not RsAlmacenes.EOF
                wcodpro = Trim(RsStockDet.Fields("F5CODPRO"))
                
                SQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"
                If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
                RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                
                'SI NO LO ENCUENTRA, LO CREA
                If RsMovAlmacen.EOF Then
                    SQL = "INSERT INTO IF6ALMA(F2CODALM,F5CODPRO) VALUES('" & wcodalm & "','" & wcodpro & "')"
                    cnn_dbbancos.Execute (SQL)
                End If
                
                SQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"
                If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
                   RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsMovAlmacen.EOF Then
                    RsMovAlmacen.MoveFirst
                    Do While Not RsMovAlmacen.EOF
                        If CVDate(RsStockDet.Fields("F4fecval")) > CVDate(wultinv) Then
                            wmes = Month(CVDate(RsStockDet.Fields("F4fecval")))
                            If Left(RsStockDet.Fields("F4numval"), 1) = "S" Then
                                If PROMEDIO = 1 Then ' para el valvta de los movimiento de salida
                                    wcanpro = Val(Format(RsStockDet.Fields("F3canpro"), "#0.00"))
                    
                                    SQL = "Select * from IF3VALES where F2CODALM='" & wcodalm & "' and F5CODPRO='" & wcodpro & "' and left(F4numval,1)='I' and F3SALPEP>0 ORDER BY f4fecval,f4numval"
                                    If rsconsulta.State = adStateOpen Then rsconsulta.Close
                                    rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                                    If Not rsconsulta.EOF Then
                                        rsconsulta.MoveFirst
                                        wtemp = 0
                                        Do While wtemp < wcanpro
                    
                                            If rsconsulta.Fields("f3salpep") <= (wcanpro - wtemp) Then
                                                wtemp = wtemp + rsconsulta.Fields("F3salpep")
                            
                                                SQL = "UPDATE IF3VALES SET F3SALPEP = '0' WHERE F2CODALM='" & wcodalm & "' and F5CODPRO='" & wcodpro & "' and left(F4numval,1)='I'"
                                                cnn_dbbancos.Execute (SQL)
                                                wvalsol = wvalsol + (Val(Format(rsconsulta.Fields("F3salpep"), "#0.00")) * Val(Format(rsconsulta.Fields("f3valvta"), "#0.00")))
                                                
                                            Else
                                                SalPep = "" & Val(rsconsulta.Fields("F3salpep") - (wcanpro - wtemp))
                                                SQL = "UPDATE IF3VALES SET F3SALPEP = " & SalPep & " WHERE F2CODALM='" & wcodalm & "' and F5CODPRO='" & wcodpro & "' and left(F4numval,1)='I'"
                                                cnn_dbbancos.Execute (SQL)
                                                wvalsol = wvalsol + ((wcanpro - wtemp) * Val(Format(rsconsulta.Fields("f3valvta"), "#0.00")))
                                                wtemp = wcanpro
                                            End If
                                        rsconsulta.MoveNext
                                        If rsconsulta.EOF Then Exit Do
                                        Loop
                                    End If
                                    wvalsol = IIf(wcanpro = 0#, 0#, wvalsol / wcanpro)
                                    rsconsulta.Close
                                End If
                
                                ValVta = Format(RsMovAlmacen.Fields("f5cospro"), "#0.000")
                                ValDol = Format(RsMovAlmacen.Fields("f5cosprod"), "#0.000")
                                Totite = Format(RsMovAlmacen.Fields("f5cospro") * Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")), "#0.000")
                                TotDol = Format(RsMovAlmacen.Fields("f5cosprod") * Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")), "#0.000")
                                
                                SQL = "UPDATE IF3VALES SET F3valvta =" & ValVta & ",F3valdol =" & ValDol & ",F3totite =" & Totite & ",F3totdol =" & TotDol & " WHERE F2CODALM='" & wcodalm & "' and F5CODPRO='" & wcodpro & "' and left(F4numval,1)='S'"
                                cnn_dbbancos.Execute (SQL)
                                
                                stockact = RsMovAlmacen.Fields("F6STOCKACT") - Val(Format(RsStockDet.Fields("F3canpro"), "#0.000"))
                                Habm = RsMovAlmacen.Fields("f5habm" & Format(wmes, "00")) + Val(Format(RsStockDet.Fields("F3canpro"), "#0.000"))
                                sal = IIf(IsNull(RsMovAlmacen.Fields("f5sal" & Format(wmes, "00"))), 0, RsMovAlmacen.Fields("f5sal" & Format(wmes, "00"))) + (Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")) * RsStockDet.Fields("F3valvta"))
                                sald = Val("" & RsMovAlmacen.Fields("f5sald" & Format(wmes, "00"))) + (Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")) * RsStockDet.Fields("F3valdol"))
                                
                                GRABAR_COSTOS
                            Else
                                If Trim(RsStockDet.Fields("F3JCG")) = "*" Then
                
                                    ValVta = Format(RsMovAlmacen.Fields("f5cospro"), "#0.000")
                                    ValDol = Format(RsMovAlmacen.Fields("f5cosprod"), "#0.000")
                                    Totite = Format(RsMovAlmacen.Fields("f5cospro") * Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")), "#0.000")
                                    TotDol = Format(RsMovAlmacen.Fields("f5cosprod") * Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")), "#0.000")
                            
                                    SQL = "UPDATE IF3VALES SET F3valvta =" & ValVta & ",F3valdol =" & ValDol & ",F3totite =" & Totite & ",F3totdol =" & TotDol & " WHERE F2CODALM='" & wcodalm & "' and F5CODPRO='" & wcodpro & "' and left(F4numval,1)='I'"
                                    cnn_dbbancos.Execute (SQL)
                                End If
                                If (Val(Format("" & RsMovAlmacen.Fields("f6stockact"), "#0.00")) + Val(Format(RsStockDet.Fields("f3canpro"), "#0.00"))) <> 0 Then
                                    Cospro = Format(((Val(Format("" & RsMovAlmacen.Fields("f6stockact"), "#0.000")) * Val(Format(RsMovAlmacen.Fields("f5cospro"), "#0.000"))) + (Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")) * Val(Format(RsStockDet.Fields("f3valvta"), "#0.000")))) / (Val(Format(RsMovAlmacen.Fields("f6stockact"), "#0.000")) + Val(Format(RsStockDet.Fields("f3canpro"), "#0.000"))), "#0.000")
                                    Cosprod = Format(((Val(Format("" & RsMovAlmacen.Fields("f6stockact"), "#0.000")) * Val(Format(RsMovAlmacen.Fields("f5cosprod"), "#0.000"))) + (Val(Format(RsStockDet.Fields("f3canpro"), "#0.000")) * Val(Format(RsStockDet.Fields("f3valdol"), "#0.000")))) / (Val(Format(RsMovAlmacen.Fields("f6stockact"), "#0.000")) + Val(Format(RsStockDet.Fields("f3canpro"), "#0.000"))), "#0.000")
                                End If
                                stockact = RsMovAlmacen.Fields("F6STOCKACT") + Val(Format(RsStockDet.Fields("F3canpro"), "#0.000"))
                                Debm = RsMovAlmacen.Fields("f5debm" & Format(wmes, "00")) + Val(Format(RsStockDet.Fields("F3canpro"), "#0.000"))
                                
                                ing = Val("" & RsMovAlmacen.Fields("f5ing" & Format(wmes, "00"))) + Val(Format(RsStockDet.Fields("F3totite"), "#0.000"))
                                ingd = IIf(IsNull(RsMovAlmacen.Fields("f5ingd" & Format(wmes, "00"))), 0, RsMovAlmacen.Fields("f5ingd" & Format(wmes, "00"))) + Val(Format(RsStockDet.Fields("F3totdol"), "#0.000"))
                            
                                GRABAR_COSTOS2
                            
                            End If
                        End If
                        RsMovAlmacen.MoveNext
                        If RsMovAlmacen.EOF Then
                             RsMovAlmacen.Close
                            Exit Do
                        End If
                        'gauge.Value = gauge.Value + 1
                    Loop
                End If
                RsAlmacenes.MoveNext
                If RsAlmacenes.EOF Then
                RsAlmacenes.Close
                Exit Do
                End If
            Loop
            RsStockDet.MoveNext
            If RsStockDet.EOF Then Exit Do
        Loop
        
        '--------------------------------------------------------------------------
        '--------------------------------------------------------------------------
        '--------------------------------------------------------------------------
        Dim rsif6       As New ADODB.Recordset
        Dim ncantidad   As Double
        Dim csql        As String
        
        cnn_dbbancos.Execute ("UPDATE IF5PLA SET F5STOCKACT =0.00")
        
        csql = "SELECT F5CODPRO FROM IF5PLA"
        If RsProducto.State = adStateOpen Then RsProducto.Close
        RsProducto.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProducto.EOF Then
            RsProducto.MoveFirst
            Do While Not RsProducto.EOF
                ncantidad = 0
                csql = "SELECT F6STOCKACT FROM IF6ALMA WHERE F5CODPRO='" & RsProducto.Fields("F5CODPRO") & "'"
                If rsif6.State = adStateOpen Then rsif6.Close
                rsif6.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsif6.EOF Then
                    rsif6.MoveFirst
                    Do While Not rsif6.EOF
                        ncantidad = ncantidad + Val(rsif6.Fields("F6STOCKACT") & "")
                        rsif6.MoveNext
                    Loop
                    
                    If ncantidad > 0 Then
                        csql = "UPDATE IF5PLA SET F5STOCKACT=" & ncantidad & " WHERE F5CODPRO='" & RsProducto.Fields("F5CODPRO") & "'"
                        cnn_dbbancos.Execute (csql)
                    End If
                    
                End If
                rsif6.Close
                RsProducto.MoveNext
            Loop
        End If
        RsProducto.Close
        
        '--------------------------------------------------------------------------
        '--------------------------------------------------------------------------
        '--------------------------------------------------------------------------
        
        'gauge.Value = 0#
    
''        SQL = "SELECT * FROM IF5PLA"
''        If RsProducto.State = adStateOpen Then RsProducto.Close
''        RsProducto.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
''        If Not RsProducto.EOF Then
''            'wval = RsProducto.RecordCount
''            'gauge.Max = RsProducto.RecordCount
''            RsProducto.MoveFirst
''
''            Do While Not RsProducto.EOF
''                Rem JCG SQL = "UPDATE IF5PLA SET F5STOCKACT =" & ACTUAL & ""
''                Rem JCG cnn_dbbancos.Execute SQL
''                RsProducto.MoveNext
''
''                For i = 1 To 2
''                    SQL = "SELECT * FROM IF6ALMA WHERE F2CODALM ='" & Format(i, "00") & "' AND F5CODPRO ='" & RsProducto.Fields("F5CODPRO") & "'"
''                    If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
''                    RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
''                    If Not RsMovAlmacen.EOF Then
''                        Do While Not RsMovAlmacen.EOF
''                            Stock = Val(Format(RsProducto.Fields("F5STOCKACT"), "0.00")) + Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.00"))
''                            SQL = "UPDATE IF5PLA SET F5STOCKACT =" & Stock & " WHERE F5CODPRO ='" & wcodpro & "'"
''                            cnn_dbbancos.Execute SQL
''                            RsMovAlmacen.MoveNext
''                        Loop
''                        If RsMovAlmacen.EOF Then Exit Do
''                        RsProducto.MoveNext
''                    End If
''                Next i
''            Loop
''        End If
    
    'For valor% = Val(gauge.Value) To gauge.Max
    '    gauge.Value = gauge.Value + 1
    'Next
    End If

Exit Sub
Depura:
    MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atención"
    Resume

End Sub

Private Sub BtnExit_Click()
    
    Unload Me

End Sub

Private Sub BtnSaldos_Click()
    
    Me.MousePointer = 11
    If MsgBox("Desea iniciar la regeneración?..", 36, "Inventarios") = 6 Then
        LabelRe.Visible = True
        
        Regenerar_Saldos
        
        LabelRe.Visible = False
    End If
    Me.MousePointer = 1

End Sub

Private Sub Form_Load()
Set cnn_control = New ADODB.Connection
'Set CNN_PERSONAL = New ADODB.Connection
'Set DBBANCOWIN = New ADODB.Connection

    cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CONTROL.MDB;Persist Security Info=False"
    If cnn_control.State = adStateOpen Then cnn_control.Close
    cnn_control.Open cconex_control
    
''    cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Wruta & "EMPRESA.MDB;Persist Security Info=False"
''    If CNN_PERSONAL.State = adStateOpen Then CNN_PERSONAL.Close
''    CNN_PERSONAL.Open cconex_empresa
    
''    cconex_inventa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ruta & "INVENTA.MDB;Persist Security Info=False"
''    If DBBANCOWIN.State = adStateOpen Then DBBANCOWIN.Close
''    DBBANCOWIN.Open cconex_inventa

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo seguir

    TBPRODUCTO.Close
    TbStockCab.Close
    TbStockDet.Close
    TbMovAlmacen.Close
    TBALMACEN.Close
    TbParametro.Close

    DbInventa.Close
    dbempresa.Close
    dbcontrol.Close

    Set dbcontrol = Nothing
    Set DbInventa = Nothing
    Set dbempresa = Nothing

    Exit Sub
seguir: Resume Next
    
End Sub

Private Sub Limpiar_Datos()
'On Error GoTo seguido

Dim stockact, CANPRO  As Double
Dim SQL    As String

gauge.Value = 0#
'gauge.Max = 30
Set RsMovAlmacen = New ADODB.Recordset
Set RsAlmacenes = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset
Set RsProducto = New ADODB.Recordset
Set RsParametro = New ADODB.Recordset

    '-------------------------------------------------------------------------
    'STOCK
    SQL = "SELECT * FROM EF2ALMACENES "
    If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
    RsAlmacenes.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    Do While Not RsAlmacenes.EOF
        wcodalm = RsAlmacenes.Fields("F2CODALM")
        wmes = Month(CVDate(RsAlmacenes.Fields("F1ULTINV")))
        For F = wmes To 12
           'LIMPIAR_REGISTROS
           cnn_dbbancos.Execute "UPDATE IF6ALMA SET F5DEBM" & Format(F, "00") & "=0.00 ,F5HABM" & Format(F, "00") & "=0.00,F5ING" & Format(F, "00") & "=0.00,F5SAL" & Format(F, "00") & _
                               "=0.00,F5INGD" & Format(F, "00") & "=0.00,F5SALD" & Format(F, "00") & "=0.00  WHERE F2CODALM = '" & wcodalm & "'"
        Next F
        RsAlmacenes.MoveNext
    Loop
    RsAlmacenes.Close

SQL = "SELECT * FROM IF6ALMA"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsMovAlmacen.EOF Then
    RsMovAlmacen.MoveFirst
    Do While Not RsMovAlmacen.EOF
        wcodalm = "" & RsMovAlmacen.Fields("F2codalm")
        wcodpro = "" & RsMovAlmacen.Fields("F5CODPRO")
        
        SQL = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & wcodalm & "'"
        If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
        RsAlmacenes.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsAlmacenes.EOF Then
            wmes = Month(CVDate(RsAlmacenes.Fields("F1ULTINV")))
            RsAlmacenes.MoveFirst
            Do While Not RsAlmacenes.EOF
                wstock = 0#
                For I% = 0 To wmes - 1
                    wstock = wstock + (Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(I%, "00")), "#0.000")) - Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(I%, "00")), "#0.000")))
                Next I%
                
                'BUSCA VALES DE FECHA DEL MISMO MES DEL CIERRE PARA TENER EL STOCK REAl
                ing = 0#: sal = 0#: Debm = 0#: Habm = 0#: ingd = 0#: sald = 0#
                SQL = "SELECT * FROM IF3VALES WHERE F2CODALM ='" & Trim(wcodalm) & "' AND F5CODPRO ='" & wcodpro & "'"
                
                If RsStockDet.State = adStateOpen Then RsStockDet.Close
                RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsStockDet.EOF Then
                    RsStockDet.MoveFirst
                    Do While Not RsStockDet.EOF
                        If CVDate(RsStockDet.Fields("f4fecval")) <= CVDate(RsAlmacenes.Fields("F1ultinv")) And wmes = Month(CVDate(RsStockDet.Fields("f4fecval"))) Then
                            If Left(RsStockDet.Fields("f4numval"), 1) = "I" Then 'ingreso
                                wstock = wstock + RsStockDet.Fields("f3canpro")
                                Debm = Debm + RsStockDet.Fields("f3canpro")
                                ing = ing + Val("" & RsStockDet.Fields("F3TOTITE"))
                                ingd = ingd + Val("" & RsStockDet.Fields("F3TOTDOL"))
                            Else 'salida
                                wstock = wstock - RsStockDet.Fields("f3canpro")
                                Habm = Habm + RsStockDet.Fields("f3canpro")
                                sal = sal + Val("" & RsStockDet.Fields("F3TOTITE"))
                                sald = sald + Val("" & RsStockDet.Fields("F3TOTDOL"))
                            End If
                        End If
                        RsStockDet.MoveNext
                        If RsStockDet.EOF Then Exit Do
                    Loop
                End If
''                '-------------------------------------------------------------------------
''                 'STOCK
''                 For F = wmes To 12
''
''                    LIMPIAR_REGISTROS
''
''                 Next F
                 grabar
                
                
                RsAlmacenes.MoveNext
                If RsAlmacenes.EOF Then Exit Do
            Loop
        End If
        RsMovAlmacen.MoveNext
        If RsMovAlmacen.EOF Then Exit Do
'        gauge.Value = gauge.Value + 1
    Loop
    
    If RsStockDet.RecordCount = 0 Then
        MsgBox "No se Registraron Movimientos...", vbInformation, "Inventario"
        Exit Sub
    End If
    
    Actualiza_Costos

End If
Exit Sub
seguido:
MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atención"
Resume Next

End Sub
            

Private Sub Regenerar_Saldos()
    
    Limpiar_Datos
    For valor% = gauge.Value To gauge.Max
        If gauge.Value = gauge.Max Then Exit For
        gauge.Value = gauge.Value + 1
    Next
    MsgBox "Fin del Proceso", 64, "Inventario"
    BtnExit.SetFocus

End Sub

Private Sub grabar()

    amovs_cab(0).campo = "F6STOCKACT": amovs_cab(0).valor = wstock: amovs_cab(0).TIPO = "N"
    amovs_cab(1).campo = "F5DEBM" & Format(wmes, "00"): amovs_cab(1).valor = Debm: amovs_cab(1).TIPO = "N"
    amovs_cab(2).campo = "F5HABM" & Format(wmes, "00"): amovs_cab(2).valor = Habm: amovs_cab(2).TIPO = "N"
    amovs_cab(3).campo = "F5ING" & Format(wmes, "00"): amovs_cab(3).valor = ing: amovs_cab(3).TIPO = "N"
    amovs_cab(4).campo = "F5SAL" & Format(wmes, "00"): amovs_cab(4).valor = sal: amovs_cab(4).TIPO = "N"
    amovs_cab(5).campo = "F5INGD" & Format(wmes, "00"): amovs_cab(5).valor = ingd: amovs_cab(5).TIPO = "N"
    amovs_cab(6).campo = "F5SALD" & Format(wmes, "00"): amovs_cab(6).valor = sald: amovs_cab(6).TIPO = "N"
    
    '-------------------------------------------------------------------------
    Cospro = 0#
    Cosprod = 0#
               
    If RsMovAlmacen.Fields("f5costoini") > 0# And RsMovAlmacen.Fields("f5debm00") > 0 Then
        Cospro = Format(Val("" & RsMovAlmacen.Fields("f5ing00")) / RsMovAlmacen.Fields("f5debm00"), "#0.000")
        Cosprod = Format(Val("" & RsMovAlmacen.Fields("f5ingd00")) / RsMovAlmacen.Fields("f5debm00"), "#0.000")
    End If
    
    amovs_cab(7).campo = "F5COSPRO": amovs_cab(7).valor = Cospro: amovs_cab(7).TIPO = "N"
    amovs_cab(8).campo = "F5COSPROD": amovs_cab(8).valor = Cosprod: amovs_cab(8).TIPO = "N"
    
    '------- ACTUALIZAR STOCKS
    GRABA_REGISTRO amovs_cab(), "IF6ALMA", "M", 8, cnn_dbbancos, "F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"

End Sub

Private Sub GRABAR_COSTOS()

    amovs_det(0).campo = "F6STOCKACT": amovs_det(0).valor = stockact: amovs_det(0).TIPO = "N"
    amovs_det(1).campo = "F5HABM" & Format(wmes, "00"): amovs_det(1).valor = Habm: amovs_det(1).TIPO = "N"
    amovs_det(2).campo = "F5SAL" & Format(wmes, "00"): amovs_det(2).valor = sal: amovs_det(2).TIPO = "N"
    amovs_det(3).campo = "F5SALD" & Format(wmes, "00"): amovs_det(3).valor = sald: amovs_det(3).TIPO = "N"
    
    ctipo = "M"
    
    '------- ACTUALIZAR STOCKS
    GRABA_REGISTRO amovs_det(), "IF6ALMA", ctipo, 3, cnn_dbbancos, "F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"
    
End Sub

Private Sub GRABAR_COSTOS2()

    amovs_det(0).campo = "F6STOCKACT": amovs_det(0).valor = stockact: amovs_det(0).TIPO = "N"
    amovs_det(1).campo = "F5DEBM" & Format(wmes, "00"): amovs_det(1).valor = Debm: amovs_det(1).TIPO = "N"
    amovs_det(2).campo = "F5ING" & Format(wmes, "00"): amovs_det(2).valor = ing: amovs_det(2).TIPO = "N"
    amovs_det(3).campo = "F5INGD" & Format(wmes, "00"): amovs_det(3).valor = ingd: amovs_det(3).TIPO = "N"
    ctipo = "M"
    
    '------- ACTUALIZAR STOCKS
    GRABA_REGISTRO amovs_det(), "IF6ALMA", "M", 3, cnn_dbbancos, "F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"

End Sub

Private Sub LIMPIAR_REGISTROS()
    
    Arreglo(0).campo = "F5DEBM" & Format(F, "00"): Arreglo(0).valor = 0#: Arreglo(0).TIPO = "N"
    Arreglo(1).campo = "F5HABM" & Format(F, "00"): Arreglo(1).valor = 0#: Arreglo(1).TIPO = "N"
    Arreglo(2).campo = "F5ING" & Format(F, "00"): Arreglo(2).valor = 0#: Arreglo(2).TIPO = "N"
    Arreglo(3).campo = "F5SAL" & Format(F, "00"): Arreglo(3).valor = 0#: Arreglo(3).TIPO = "N"
    Arreglo(4).campo = "F5INGD" & Format(F, "00"): Arreglo(4).valor = 0#: Arreglo(4).TIPO = "N"
    Arreglo(5).campo = "F5SALD" & Format(F, "00"): Arreglo(5).valor = 0#: Arreglo(5).TIPO = "N"
        
    '------- ACTUALIZAR STOCKS
    GRABA_REGISTRO Arreglo(), "IF6ALMA", "M", 5, cnn_dbbancos, "F2CODALM = '" & wcodalm & "' AND F5CODPRO = '" & wcodpro & "'"

End Sub

