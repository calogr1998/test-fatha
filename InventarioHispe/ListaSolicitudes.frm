VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lista_solicitudes 
   Caption         =   "Solicitudes de Materiales"
   ClientHeight    =   7200
   ClientLeft      =   1155
   ClientTop       =   1440
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
   ScaleHeight     =   7200
   ScaleWidth      =   10380
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6735
      Left            =   45
      OleObjectBlob   =   "ListaSolicitudes.frx":0000
      TabIndex        =   0
      Top             =   360
      Width           =   10290
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   45
      Top             =   90
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
      Tools           =   "ListaSolicitudes.frx":2F1A
      ToolBars        =   "ListaSolicitudes.frx":61C2
   End
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   630
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+Enter -> Buscar Siguiente  /  Shift+Enter -> Encontrar Nuevo"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5625
      TabIndex        =   1
      Top             =   90
      Width           =   4650
   End
End
Attribute VB_Name = "lista_solicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rslista     As ADODB.Recordset
Dim i           As Byte
Dim DBName      As String
Dim EditLookUp  As Boolean

Private Sub Form_Load()

    Me.AutoRedraw = False
    Me.Height = 8040
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    sw_nuevo_documento = True
    Me.AutoRedraw = True
        
    PROCESO
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = 11
            solicitud.Show 1
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    
    End Select
    
End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    sw_nuevo_documento = False
    Me.MousePointer = 11
    solicitud.Show 1
    Me.MousePointer = 1

End Sub

Public Sub PROCESO()
Dim csql        As String
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dblogi
        csql = "select cod_solicitud,cs_fecha,cs_codsolicitante,cs_moneda,cs_codcosto,cs_total,anulado " & _
              " from tb_cabsolicitud WHERE anulado ='N' order by cod_solicitud DESC"
       
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "cod_solicitud"
    End With

End Sub

Private Sub cmdColor_Click(Index As Integer)
    
    With cdColor
        .CancelError = True
        On Error GoTo ErrHandler
        .Flags = cdlCCRGBInit
        Select Case Index
            Case 0: .Color = dxDBGrid1.AutoSearchColor
            Case 1: .Color = dxDBGrid1.AutoSearchTextColor
        End Select
        .ShowColor
        Select Case Index
            Case 0:
                dxDBGrid1.AutoSearchColor = .Color
                lblresult.BackColor = .Color
            Case 1:
                dxDBGrid1.AutoSearchTextColor = .Color
                lblresult.ForeColor = .Color
        End Select
    End With
    
ErrHandler:
    Exit Sub
    
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

Private Sub Form_Resize()
On Error Resume Next

    'picTop.Move 0, 0, ScaleWidth
    'dxDBGrid1.Move 0, picTop.Height, ScaleWidth, ScaleHeight - picTop.Height

End Sub
