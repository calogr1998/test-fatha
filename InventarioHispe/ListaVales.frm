VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form ListaVales 
   Caption         =   "Vales"
   ClientHeight    =   7245
   ClientLeft      =   900
   ClientTop       =   600
   ClientWidth     =   10335
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
   ScaleHeight     =   7245
   ScaleWidth      =   10335
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   405
      Top             =   -90
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
      Tools           =   "ListaVales.frx":0000
      ToolBars        =   "ListaVales.frx":32A8
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6645
      Left            =   135
      OleObjectBlob   =   "ListaVales.frx":3364
      TabIndex        =   0
      Top             =   315
      Width           =   10125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5490
      TabIndex        =   1
      Top             =   90
      Width           =   4650
   End
End
Attribute VB_Name = "ListaVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditLookUp  As Boolean
Dim i           As Byte

Private Sub dxDBGrid1_OnDblClick()
    
    sw_nuevo_documento = False
    Me.MousePointer = 11
    If wtipoguia = "I" Then
        vale_ingreso.Show 1
    Else
        vale_salida.Show 1
    End If
    Me.MousePointer = 1
    
End Sub

Private Sub Form_Load()
    
    Me.AutoRedraw = False
    Me.Height = 8040
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    
    
    sw_nuevo_documento = True
    If wtipoguia = "S" Then
        Me.Caption = "Vale de Salida"
    Else
        Me.Caption = "Vale de Ingreso"
    End If
    Me.AutoRedraw = True
    PROCESO

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = 11
            If wtipoguia = "I" Then
                vale_ingreso.Show 1
            Else
                vale_salida.Show 1
            End If
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    End Select

End Sub

Public Sub PROCESO()
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = dbbancowin
        If wtipoguia = "S" Then

            SQL = "SELECT A.F2CODALM,A.F4NUMVAL,A.F4NUMDOC,A.F4FECVAL,B.F1NOMORI " & _
            "FROM IF4VALES AS A, SF1ORIGENES AS B WHERE MID(F4NUMVAL,1,1) = 'S' AND " & _
            "A.F1CODORI=B.F1CODORI ORDER BY A.F2CODALM,A.F4NUMVAL"

        Else

            SQL = "SELECT A.F2CODALM,A.F4NUMVAL,A.F4NUMDOC,A.F4FECVAL,B.F1NOMORI " & _
            "FROM IF4VALES AS A, SF1ORIGENES AS B WHERE MID(F4NUMVAL,1,1) = 'I' AND " & _
            "A.F1CODORI=B.F1CODORI ORDER BY A.F2CODALM,A.F4NUMVAL"

         End If

        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SQL
        .Dataset.Active = True
        .KeyField = "F2CODALM"
    End With
 
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
 
    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.M.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With
    
End Sub
