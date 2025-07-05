VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmExcel 
   Caption         =   "Cargar Excel"
   ClientHeight    =   5385
   ClientLeft      =   1845
   ClientTop       =   1920
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7785
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exportar Access"
      Height          =   1095
      Left            =   3840
      TabIndex        =   4
      Top             =   4200
      Width           =   3615
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importar"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   3615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo Excel :"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog Excel 
      Left            =   2280
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread FPSpread1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      _Version        =   196608
      _ExtentX        =   13573
      _ExtentY        =   6588
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmExcel.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCargar 
         Caption         =   "Cargar Archivo Excel"
      End
      Begin VB.Menu MnuDescargar 
         Caption         =   "Descargar Archivo"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim handle As Integer
Dim rstemp As ADODB.Recordset
Dim cn As ADODB.Connection

'Private Sub conectar()
'Set cn = New ADODB.Connection
'cconex_form.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ".mdb;Persist Security Info=False"
'cn.Open
'End Sub
Private Sub cargartemp()
Dim i As Long
Dim j As Long
Dim h As Long

For j = 2 To FPSpread1.MaxRows + 2
    For i = 1 To FPSpread1.MaxCols
        FPSpread1.Row = j: FPSpread1.col = i
        If FPSpread1.Value = "" Then
            Exit Sub
        Else
            rstemp.AddNew
            Select Case xvale
            Case 1
                h = i
                'xx = Str(Format(Val(FPSpread1.Value), "0"))
                'xval = Val("" & FPSpread1.Value)
                If Right(FPSpread1.Value, 5) = 0 Then
                    xval = Val("" & FPSpread1.Value)
                Else
                    xval = "" & FPSpread1.Value
                    xval = -1
                End If
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xval = "" & FPSpread1.Value
                End If
                rstemp(0) = xval: h = h + 1: FPSpread1.col = h
                If Right(FPSpread1.Value, 5) = 0 Then
                    xvalC = Val("" & FPSpread1.Value)
                Else
                    xvalC = "" & FPSpread1.Value
                    xvalC = -1
                End If
                'XVALC = Val("" & FPSpread1.Value)
                If xvalC > 0 Then
                    xvalC = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xvalC = "" & FPSpread1.Value
                End If
                rstemp(1) = xvalC: h = h + 1: FPSpread1.col = h
                rstemp(2) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp(3) = Str(Format(Val(FPSpread1.Value), "0")): h = h + 1: FPSpread1.col = h
                rstemp(4) = Str(Format(Val(FPSpread1.Value), "0")): h = h + 1: FPSpread1.col = h
                rstemp.Update
            Case 2
                h = i
                xval = Val("" & FPSpread1.Value)
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xval = "" & FPSpread1.Value
                End If
                rstemp(0) = xval: h = h + 1: FPSpread1.col = h
                If Right(FPSpread1.Value, 5) = 0 Then
                    xvalC = Val("" & FPSpread1.Value)
                Else
                    xvalC = "" & FPSpread1.Value
                    xvalC = -1
                End If
                'xvalC = Val("" & FPSpread1.Value)
                If xvalC > 0 Then
                    xvalC = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xvalC = "" & FPSpread1.Value
                End If
                rstemp(1) = xvalC: h = h + 1: FPSpread1.col = h
                rstemp(2) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp(3) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp.Update
            Case 3
                h = i
                xval = Val("" & FPSpread1.Value)
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xval = "" & FPSpread1.Value
                End If
                rstemp(0) = xval: h = h + 1: FPSpread1.col = h
                'xvalC = Val("" & FPSpread1.Value)
                If Right(FPSpread1.Value, 5) = 0 Then
                    xvalC = Val("" & FPSpread1.Value)
                Else
                    xvalC = "" & FPSpread1.Value
                    xvalC = -1
                End If
                If xvalC > 0 Then
                    xvalC = "" & Trim("" & Str(Format(Val(FPSpread1.Value), "0")))
                Else
                    xvalC = "" & FPSpread1.Value
                End If
                rstemp(1) = xvalC: h = h + 1: FPSpread1.col = h
                rstemp(2) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp(3) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp(4) = FPSpread1.Value: h = h + 1: FPSpread1.col = h
                rstemp.Update
            End Select
            
          End If
          Exit For
    Next
Next
End Sub
Private Sub cmdExportar_Click()
Dim SQL As String
Dim RS As ADODB.Recordset

cargartemp
'conectar
Set RS = New ADODB.Recordset

If rstemp.RecordCount <> 0 Then
    rstemp.MoveFirst
End If
'---------------------
    i = 0
    Select Case xvale
    Case 1
        With vale_ingreso.dxDBGrid1.Dataset   'vale_salida.dxDBGrid1.Dataset
            sw_nuevo_temp = False
            sw_nuevo_item = True
            Do While Not rstemp.EOF
                i = i + 1
                If sw_nuevo_temp = False Then
                    .Edit
                    sw_nuevo_temp = True
                Else
                    .Append
                End If
                .FieldValues("ITEM") = i
                If Right(rstemp(1), 5) = 0 Then
                    xval = Val("" & rstemp(1))
                Else
                    xval = "" & rstemp(1)
                    xval = -1
                End If
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(rstemp(1)), "0")))
                Else
                    xval = "" & rstemp(1)
                End If
                If rsif5pla.State = adStateOpen Then rsif5pla.Close
                rsif5pla.Open "SELECT F5CODPRO,F7CODMED,F5STOCKACT FROM IF5PLA WHERE F5CODFAB='" & xval & "' AND F5MARCA='" & Trim("" & rstemp(0)) & "'", cnn_dbbancos
                .FieldValues("CODPROD") = "" & rsif5pla.Fields("F5CODPRO")
                .FieldValues("CODFAB") = "" & Trim(xval)
                .FieldValues("MARCA") = "" & Trim(rstemp(0))
                .FieldValues("DESCRIPCION") = "" & Trim(rstemp(2))
                .FieldValues("UMEDIDA") = "" & rsif5pla.Fields("F7CODMED")
                '.FieldValues("STOCKACTUAL") = Format(0, "###,##0.00")
                .FieldValues("CANTIDAD") = Format(Val("" & rstemp(3)), "###,##0.00")
                .FieldValues("COSTOUNI") = Format(Val("" & rstemp(4)), "###,##0.00")
                .FieldValues("IGV") = Format(0, "###,##0.00")
                .FieldValues("TOTAL") = Format(0, "###,##0.00")
                .Post
                rsif5pla.Close
                rstemp.MoveNext
            Loop
        
         End With
    '------------------------------------
    'Vale Salida
    Case 2
            With vale_salida.dxDBGrid1.Dataset
            sw_nuevo_temp = False
            sw_nuevo_item = True
            Do While Not rstemp.EOF
                i = i + 1
                If sw_nuevo_temp = False Then
                    .Edit
                    sw_nuevo_temp = True
                Else
                    .Append
                End If
                .FieldValues("ITEM") = i
                'xval = Val("" & rstemp(1))
                If Right(rstemp(1), 5) = 0 Then
                    xval = Val("" & rstemp(1))
                Else
                    xval = "" & rstemp(1)
                    xval = -1
                End If
                
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(rstemp(1)), "0")))
                Else
                    xval = "" & rstemp(1)
                End If
                If rsif5pla.State = adStateOpen Then rsif5pla.Close
                rsif5pla.Open "SELECT F5CODPRO,F7CODMED,F5STOCKACT FROM IF5PLA WHERE F5CODFAB='" & xval & "' AND F5MARCA='" & Trim("" & rstemp(0)) & "'", cnn_dbbancos
                .FieldValues("CODPROD") = "" & rsif5pla.Fields("F5CODPRO")
                .FieldValues("CODFAB") = "" & Trim(xval)
                .FieldValues("MARCA") = "" & Trim(rstemp(0))
                .FieldValues("DESCRIPCION") = "" & Trim(rstemp(2))
                .FieldValues("UMEDIDA") = "" & rsif5pla.Fields("F7CODMED")
                .FieldValues("STOCKACTUAL") = Format(Val("" & rsif5pla.Fields("F5STOCKACT")), "###,##0.00")
                .FieldValues("CANTIDAD") = Format(Val("" & rstemp(3)), "###,##0.00")
                .Post
                rsif5pla.Close
                rstemp.MoveNext
            Loop
        
        End With
    Case 3
            With FrmTransferencias.dxDBGrid1.Dataset    'vale_salida.dxDBGrid1.Dataset
            sw_nuevo_temp = False
            sw_nuevo_item = True
            Do While Not rstemp.EOF
                i = i + 1
                If sw_nuevo_temp = False Then
                    .Edit
                    sw_nuevo_temp = True
                Else
                    .Append
                End If
                .FieldValues("F3ITEM") = i
                'xval = Val("" & rstemp(1))
                If Right(rstemp(1), 5) = 0 Then
                    xval = Val("" & rstemp(1))
                Else
                    xval = "" & rstemp(1)
                    xval = -1
                End If
                If xval > 0 Then
                    xval = "" & Trim("" & Str(Format(Val(rstemp(1)), "0")))
                Else
                    xval = "" & rstemp(1)
                End If
                If rsif5pla.State = adStateOpen Then rsif5pla.Close
                rsif5pla.Open "SELECT F5CODPRO,F7CODMED,F5STOCKACT FROM IF5PLA WHERE F5CODFAB='" & xval & "' AND F5MARCA='" & Trim("" & rstemp(0)) & "'", cnn_dbbancos
                .FieldValues("F5CODPRO") = "" & rsif5pla.Fields("F5CODPRO")
                .FieldValues("F5CODFAB") = "" & Trim(xval)
                .FieldValues("MARCA") = "" & Trim(rstemp(0))
                .FieldValues("F5NOMPRO") = "" & Trim(rstemp(2))
                .FieldValues("F6VALPRO") = Format(0, "###,##0.00") ' "" & Trim(rstemp(2))
                .FieldValues("F6CANMOV") = Format(Val("" & rstemp(3)), "###,##0.00")
                .FieldValues("F7SIGMED") = "" & rsif5pla.Fields("F7CODMED")
                .FieldValues("F6total") = Format(Val("" & rstemp(4)), "###,##0.00")
                .Post
                rsif5pla.Close
                rstemp.MoveNext
                If rstemp.EOF Then Exit Do
            Loop
        
        End With
End Select
'------------------------

'Do While Not rstemp.EOF '"Total,) " &
'    sql = "insert into DETALLE(Marca,CODFAB," & _
'          "Descripcion,Cantidad)" & _
'        "values('" & rstemp(0) & "','" & rstemp(1) & "'," & _
'        "'" & rstemp(2) & "','" & rstemp(3) & "')"
'    rs.Open sql, cnn_form, adOpenDynamic, adLockOptimistic
'    rstemp.MoveNext
'Loop

End Sub
Private Sub Combo1_Click()
  F = FPSpread1.ImportExcelSheet(handle, Combo1.ListIndex)
End Sub
Private Sub Form_Load()

FPSpread1.MaxCols = 15
FPSpread1.MaxRows = 1000

Set rstemp = New ADODB.Recordset
Select Case xvale
    Case 1
        rstemp.Fields.Append "Marca", adVarChar, 10
        rstemp.Fields.Append "CodInterno", adVarChar, 15
        rstemp.Fields.Append "Descripcion", adVarChar, 100
        rstemp.Fields.Append "Cantidad", adInteger
        rstemp.Fields.Append "Costo", adDouble
    Case 2
        rstemp.Fields.Append "Marca", adVarChar, 10
        rstemp.Fields.Append "CodInterno", adVarChar, 15
        rstemp.Fields.Append "Descripcion", adVarChar, 100
        rstemp.Fields.Append "Cantidad", adInteger
    Case 3
        rstemp.Fields.Append "Marca", adVarChar, 10
        rstemp.Fields.Append "CodInterno", adVarChar, 15
        rstemp.Fields.Append "Descripcion", adVarChar, 100
        rstemp.Fields.Append "Cantidad", adInteger
        rstemp.Fields.Append "Costo", adDouble
End Select
rstemp.Open
End Sub

Private Sub mnuCargar_Click()
 Dim List() As String
    Dim ListCount As Integer
    
    ReDim List(1)
        
    With Excel
        .FileName = "*.xls"
        .DialogTitle = "Selecione el archivo de Excel"
        .Filter = "Excel files|(*.xls)"
        .FilterIndex = 0
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        F = FPSpread1.GetExcelSheetList(.FileName, List, ListCount, (App.Path & "\log.txt"), handle, True)
        If (ListCount - 1 > 1) Then
            ReDim List(ListCount - 1)
            F = FPSpread1.GetExcelSheetList(.FileName, List, ListCount, (App.Path & "\log.txt"), handle, False)
        End If
        xlfile = .FileName
    End With
    Combo1.Clear
    For i = 0 To ListCount - 1
        Combo1.AddItem (List(i))
    Next i
   ' Combo1.ListIndex = 0
End Sub

Private Sub MnuDescargar_Click()
    vale_ingreso.Show 1
End Sub

Private Sub mnuSalir_Click()
    'vale_salida.dxDBGrid1.Dataset.Refresh
    
    Unload Me
End Sub
