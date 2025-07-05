VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form lista_productos 
   Caption         =   "Productos"
   ClientHeight    =   7665
   ClientLeft      =   1530
   ClientTop       =   690
   ClientWidth     =   10380
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
   ScaleHeight     =   7665
   ScaleWidth      =   10380
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   225
      Top             =   7245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "lista_productos.frx":0000
      ToolBars        =   "lista_productos.frx":32A8
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   7050
      Left            =   45
      OleObjectBlob   =   "lista_productos.frx":336C
      TabIndex        =   0
      Top             =   90
      Width           =   10275
   End
End
Attribute VB_Name = "lista_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VSize           As Long
Dim HSize           As Long
Dim Values()        As Variant

Private Sub dxDBGrid1_OnUnboundAddRecord(ByVal RecNo As Long, ByVal Bookmark As Variant, ByVal Fields As Variant)

    If RecNo >= 0 And RecNo <= VSize Then
        VSize = VSize + 1
        ReDim Preserve Values(HSize - 1, VSize)
        For J = 0 To HSize - 1
            For i = VSize - 1 To RecNo Step -1
                 Values(J, i) = Values(J, i - 1)
            Next
        Next
        For i = 0 To HSize - 1
            Values(i, RecNo) = Fields(i)
        Next
    End If

End Sub

Private Sub dxDBGrid1_OnUnboundDeleteRecord(ByVal RecNo As Long, ByVal Bookmark As Variant)

    If RecNo >= 0 And RecNo < VSize Then
        For J = 0 To HSize - 1
            For i = RecNo To VSize - 1
                Values(J, i) = Values(J, i + 1)
            Next
        Next
        VSize = VSize - 1
        ReDim Preserve Values(HSize - 1, VSize)
    End If
End Sub

Private Sub dxDBGrid1_OnUnboundGetRecord(ByVal RecNo As Long, ByVal Bookmark As Variant, Fields As Variant)

    If RecNo >= 0 And RecNo < VSize Then
        Dim Arr() As Variant
        ReDim Arr(HSize - 1)
        For i = 0 To HSize - 1
            Arr(i) = Values(i, RecNo)
        Next
        Fields = Arr
    End If

End Sub

Private Sub dxDBGrid1_OnUnboundSetRecord(ByVal RecNo As Long, ByVal Bookmark As Variant, ByVal Fields As Variant)

    If RecNo >= 0 And RecNo < VSize Then
        For i = 0 To HSize - 1
            If Not IsEmpty(Fields(i)) Then
                Values(i, RecNo) = Fields(i)
            End If
        Next
    End If

End Sub

Private Sub AddRecords()
Dim ncontador       As Long
Dim nfil            As Integer
Dim cnomcli         As String
Dim csql            As String
  
    csql = "SELECT COUNT(F5CODPRO) AS TOT_ITEM FROM IF5PLA"
    rsif5pla.Open csql, cnn_personal, adOpenDynamic, adLockOptimistic
    ncontador = Val("" & rsif5pla.Fields("TOT_ITEM"))
    rsif5pla.Close

    dxDBGrid1.Dataset.DisableControls
    dxDBGrid1.Dataset.UserDataset.UserRecordCount = ncontador

    VSize = ncontador
    ReDim Values(HSize - 1, ncontador)

    nfil = 0
    csql = "SELECT F5CODPRO,F5NOMPRO,F5CODFAB,F7CODMED,F5MARCA FROM IF5PLA ORDER BY F5CODPRO"
    
    rsif5pla.Open csql, dbbancowin, adOpenDynamic, adLockOptimistic
    If Not rsif5pla.EOF Then
        Do While Not rsif5pla.EOF
            Values(0, nfil) = rsif5pla.Fields("F5CODPRO") & ""
            Values(1, nfil) = rsif5pla.Fields("F5CODFAB") & ""
            Values(2, nfil) = rsif5pla.Fields("F5NOMPRO") & ""
            rsmarcas.Open "SELECT * FROM EF2MARCAS WHERE F2CODMAR='" & rsif5pla.Fields("F5MARCA") & "'", cnn_empresa, adOpenDynamic, adLockOptimistic
            If Not rsmarcas.EOF Then
                Values(3, nfil) = rsmarcas.Fields("F2DESMAR") & ""
            Else
                Values(3, nfil) = ""
            End If
            rsmarcas.Close
            Values(4, nfil) = rsif5pla.Fields("F7CODMED") & ""
            nfil = nfil + 1
            rsif5pla.MoveNext
        Loop
    End If
    rsif5pla.Close

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Open
   
End Sub

Private Sub AddColumns(ByVal Num As Long)
Dim NewColumn As dxGridColumn
  
    With dxDBGrid1
        .Dataset.Close
        .Columns.DestroyColumns
        For i = 0 To Num - 1
            Select Case i
                Case 0:
                    Set NewColumn = .Columns.Add(gedTextEdit)
                    NewColumn.FieldName = "Column" & i
                    NewColumn.ObjectName = "UColumn" & i
                    NewColumn.Caption = "Codigo"
                    NewColumn.Alignment = taLeftJustify
                    NewColumn.RowIndex = 0
                Case 1:
                    Set NewColumn = .Columns.Add(gedTextEdit)
                    NewColumn.FieldName = "Column" & i
                    NewColumn.ObjectName = "UColumn" & i
                    NewColumn.Caption = "Cod. Fab."
                    NewColumn.Alignment = taLeftJustify
                    NewColumn.RowIndex = 0
                Case 2:
                    Set NewColumn = .Columns.Add(gedTextEdit)
                    NewColumn.FieldName = "Column" & i
                    NewColumn.ObjectName = "UColumn" & i
                    NewColumn.Caption = "Descripcion"
                    NewColumn.Alignment = taLeftJustify
                    NewColumn.RowIndex = 0
                Case 3:
                    Set NewColumn = .Columns.Add(gedDateEdit)
                    NewColumn.FieldName = "Column" & i
                    NewColumn.ObjectName = "UColumn" & i
                    NewColumn.Caption = "Marca"
                    NewColumn.Alignment = taLeftJustify
                    NewColumn.RowIndex = 0
                Case 4:
                    Set NewColumn = .Columns.Add(gedTextEdit)
                    NewColumn.FieldName = "Column" & i
                    NewColumn.ObjectName = "UColumn" & i
                    NewColumn.Caption = "U.M."
                    NewColumn.Alignment = taLeftJustify
                    NewColumn.RowIndex = 0
            End Select
        Next
        
        For i = 0 To HSize - 1
            J = 12
            If i = 2 Then J = 70
            If i = 3 Then J = 15
            If i = 4 Then J = 8
            .Dataset.UserDataset.SetFieldSize i, J
        Next
        
    End With
  
End Sub

Private Sub CONFIGURA_GRID()
    
    With dxDBGrid1
       .DatasetType = dtUnbound
       .Options.Set (egoAutoWidth)
       .Options.Set (egoShowGroupPanel)
       .Options.Set (egoLoadAllRecords)
       .Options.Set (egoAutoSort)
       .Columns.HeaderPanelRowCount = 1
       .KeyField = "UColumn0field"
    End With

    HSize = 5
        
    AddColumns HSize
    
    AddRecords
    
End Sub

Private Sub dxDBGrid1_OnDblClick()

    sw_nuevo_documento = False
    mant_productos.Show 1

End Sub

Private Sub Form_Activate()

    CONFIGURA_GRID
    
End Sub

Private Sub Form_Load()

    Me.AutoRedraw = False
    
    Me.Height = 8040
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    
    sw_nuevo_documento = True
    
    Me.AutoRedraw = True
                
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_documento = True
            mant_productos.Show 1
        Case "ID_Imprimir":
            'IMPRIMIR
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub
