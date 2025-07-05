VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form FrmModPrecios 
   Caption         =   "Modificación de Precios"
   ClientHeight    =   6855
   ClientLeft      =   1755
   ClientTop       =   1560
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9660
   Begin VB.Frame Frame2 
      Height          =   5910
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   10410
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5565
         Left            =   135
         OleObjectBlob   =   "FrmModPrecios.frx":0000
         TabIndex        =   5
         Top             =   225
         Width           =   10215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   10410
      Begin VB.TextBox Txtcodalm 
         Height          =   285
         Left            =   3810
         MaxLength       =   2
         TabIndex        =   1
         Top             =   315
         Width           =   510
      End
      Begin Threed.SSPanel PnlNomAlm 
         Height          =   285
         Left            =   4365
         TabIndex        =   2
         Top             =   315
         Width           =   3930
         _Version        =   65536
         _ExtentX        =   6932
         _ExtentY        =   503
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
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Almacén Origen:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   2205
         TabIndex        =   3
         Top             =   345
         Width           =   1170
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "FrmModPrecios.frx":7C23
      ToolBars        =   "FrmModPrecios.frx":E14F
   End
End
Attribute VB_Name = "FrmModPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wsql            As String
Dim msgdev          As String
Dim swvisible_fob   As Boolean


Private Sub Form_Activate()
    Txtcodalm.SetFocus
End Sub

Private Sub Form_Load()

    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    
    Procesando
    swvisible_fob = False
    msgdev = ""

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        Case "ID_Imprimir"
            With Acr_ListaGeneralPrecios
                .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cnn_dbbancos & ""
                If Len(Txtcodalm.Text) > 0 Then
                    SQL = "SELECT A.F5CODPRO AS CODIGO,A.F5NOMPRO AS NOMBRE,A.F7CODMED AS UMEDIDA,A.F5VALVTA AS VVTA1,A.F5IGVVTA AS IGV1,A.F5PREVTA AS PVTA1,A.F5VALVTA2 AS VVTA2,A.F5IGVVTA2 AS IGV2,A.F5PREVTA2 AS PVTA2,A.F5VALVTA3 AS VVTA3,A.F5IGVVTA3 AS IGV3,A.F5PREVTA3 AS PVTA3,A.F5VALVTA4 AS VVTA4,A.F5IGVVTA4 AS IGV4,A.F5PREVTA4 AS PVTA4,A.F5VALVTA5 AS VVTA5,A.F5IGVVTA5 AS IGV5,A.F5PREVTA5 AS PVTA5,A.F5FOB AS FOB,A.F5FACTOR AS FACTOR FROM IF5PLA A,IF6ALMA B WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & Trim(Txtcodalm.Text) & "' ORDER BY A.F5CODPRO"
                Else
                    SQL = "SELECT F5CODPRO AS CODIGO,F5NOMPRO AS NOMBRE,a.F7CODMED AS UMEDIDA,F5VALVTA AS VVTA1,F5IGVVTA AS IGV1,F5PREVTA AS TOTAL1,F5VALVTA2 AS VVTA2,F5IGVVTA2 AS IGV2,F5PREVTA2 AS TOTAL2,F5VALVTA3 AS VVTA3,F5IGVVTA3 AS IGV3,F5PREVTA3 AS TOTAL3,F5VALVTA4 AS VVTA4,F5IGVVTA4 AS IGV4,F5PREVTA4 AS TOTAL4,F5VALVTA5 AS VVTA5,F5IGVVTA5 AS IGV5,F5PREVTA5 AS TOTAL5,F5FOB AS FOB,F5FACTOR AS FACTOR FROM IF5PLA ORDER BY F5CODPRO"
                End If
                .DataControl1.Source = SQL
                .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                .lblempresa.Caption = wnomcia
                .LabelAlmacen.Caption = PnlNomalm.Caption
                .Show vbModal
            
            End With
            
        Case "ID_Lista_FOB":
            msgdev = InputBox("Ingrese su Contraseña ...", "Modificación de Lista de Precios")
            If rs.State = adStateOpen Then rs.Close
            rs.Open "select * from ef2USERS where f2coduseR= '" & wusuario & "' AND f2pass_autoriza_documentos='" & msgdev & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                'msgdev = "0"
                swvisible_fob = True
                dxDBGrid1.Columns.ColumnByFieldName("FOB").Visible = True
                dxDBGrid1.Columns.ColumnByFieldName("FACTOR").Visible = True
            Else
                MsgBox "Ud. no esta autorizado para visualizar los Precios FOB Y Factor ...", vbInformation, "Atención "
                swvisible_fob = False
                dxDBGrid1.Columns.ColumnByFieldName("FOB").Visible = False
                dxDBGrid1.Columns.ColumnByFieldName("FACTOR").Visible = False
            End If
            rs.Close
        Case "ID_Salir"
            Unload Me
            
    End Select
    
End Sub

Private Sub Txtcodalm_DblClick()
    
    Txtcodalm_KeyDown 113, 0
    
End Sub

Private Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        hlp_almacenes.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            Txtcodalm.Text = wcod_alm
            PnlNomalm.Caption = wnomalmacen
            Txtcodalm_KeyPress 13
        End If
    End If

End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Len(Trim(Txtcodalm.Text)) > 0 Then
            wnomalmacen = ""
            If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
                PnlNomalm.Caption = wnomalmacen
                Procesando
            Else
                MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                Txtcodalm.SetFocus
            End If
        End If

    End If
    
End Sub

Private Sub txtcodalm_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(Txtcodalm.Text)) > 0 Then
            If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
                PnlNomalm.Caption = wnomalmacen
            Else
                MsgBox "El código del almacén no existe. Verifique.", vbCritical, "Atención"
                Txtcodalm.SetFocus
            End If
        End If
    End If

End Sub

Private Sub Procesando()
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        
        If Len(Txtcodalm.Text) > 0 Then
            wsql = "SELECT A.F5CODPRO AS CODIGO,A.F5NOMPRO AS NOMBRE,A.F5VALVTA AS VVTA1,A.F5IGVVTA AS IGV1,A.F5PREVTA AS TOTAL1,A.F5VALVTA2 AS VVTA2,A.F5IGVVTA2 AS IGV2,A.F5PREVTA2 AS TOTAL2,A.F5VALVTA3 AS VVTA3,A.F5IGVVTA3 AS IGV3,A.F5PREVTA3 AS TOTAL3,A.F5VALVTA4 AS VVTA4,A.F5IGVVTA4 AS IGV4,A.F5PREVTA4 AS TOTAL4,A.F5VALVTA5 AS VVTA5,A.F5IGVVTA5 AS IGV5,A.F5PREVTA5 AS TOTAL5,F5FOB AS FOB,F5FACTOR AS FACTOR FROM IF5PLA A,IF6ALMA B WHERE A.F5CODPRO=B.F5CODPRO AND B.F2CODALM='" & Trim(Txtcodalm.Text) & "' ORDER BY A.F5CODPRO"
        Else
            wsql = "SELECT F5CODPRO AS CODIGO,F5NOMPRO AS NOMBRE,F5VALVTA AS VVTA1,F5IGVVTA AS IGV1,F5PREVTA AS TOTAL1,F5VALVTA2 AS VVTA2,F5IGVVTA2 AS IGV2,F5PREVTA2 AS TOTAL2,F5VALVTA3 AS VVTA3,F5IGVVTA3 AS IGV3,F5PREVTA3 AS TOTAL3,F5VALVTA4 AS VVTA4,F5IGVVTA4 AS IGV4,F5PREVTA4 AS TOTAL4,F5VALVTA5 AS VVTA5,F5IGVVTA5 AS IGV5,F5PREVTA5 AS TOTAL5,F5FOB AS FOB,F5FACTOR AS FACTOR FROM IF5PLA ORDER BY F5CODPRO"
        End If
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = wsql
        .Dataset.Active = True
        .KeyField = "ITEM"
    End With

End Sub
