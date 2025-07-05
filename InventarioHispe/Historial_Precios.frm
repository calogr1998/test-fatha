VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{EAEA378F-B941-4FBA-893A-680F0D58F786}#1.0#0"; "sptbdock.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Historial_Precios 
   Caption         =   "Historial de Precios"
   ClientHeight    =   5700
   ClientLeft      =   3195
   ClientTop       =   2415
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin TabDock.TTabDock TTabDock1 
      Left            =   7200
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame FraGrafico 
      Height          =   3255
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   6615
      Begin MSChart20Lib.MSChart MSPrecios 
         Height          =   4095
         Left            =   0
         OleObjectBlob   =   "Historial_Precios.frx":0000
         TabIndex        =   4
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame FraFechas 
      Caption         =   "Fechas"
      Height          =   750
      Left            =   480
      TabIndex        =   0
      Top             =   60
      Width           =   6195
      Begin MSComCtl2.DTPicker aboHasta 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker aboDesde 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   40611
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "Historial_Precios.frx":270C
      ToolBars        =   "Historial_Precios.frx":4D10
   End
End
Attribute VB_Name = "Historial_Precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Af As New ADOFunctions
Dim strNomTemporal  As String
Dim StrCn As String

Private Sub aboDesde_Change()
Carga_Recursos aboDesde.Value, abohasta.Value
End Sub

Private Sub Form_Load()
    StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "templus.mdb;Persist Security Info=False"
    
    strNomTemporal = ""
    aboDesde.Value = Format("01/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
    abohasta.Value = Date
    Carga_Recursos aboDesde.Value, abohasta.Value
    
    TTabDock1.AddForm Historial_Precios_Productos, tdDocked, tdAlignLeft, "Historial_Precios_Productos"
    TTabDock1.DockedForms.ITEM("Historial_Precios_Productos").Panel.Width = 2500
    TTabDock1.FormShow "Historial_Precios_Productos"

End Sub

Private Sub Form_Resize()
On Error Resume Next
FraFechas.Move TTabDock1.DockedForms.ITEM("Historial_Precios_Productos").Panel.Width, 0, Me.ScaleWidth - TTabDock1.DockedForms.ITEM("Historial_Precios_Productos").Panel.Width, 870
FraGrafico.Move TTabDock1.DockedForms.ITEM("Historial_Precios_Productos").Panel.Width, FraFechas.Height, Me.ScaleWidth - TTabDock1.DockedForms.ITEM("Historial_Precios_Productos").Panel.Width, Me.ScaleHeight - FraFechas.Height
MSPrecios.Move 0, 0, FraGrafico.Width, FraGrafico.Height
End Sub

Private Sub Carga_Recursos(pFechaInicio As Variant, pFechaFin As Variant)
Dim Rs As New ADODB.Recordset
If IsDate(pFechaFin) And IsDate(pFechaFin) Then
    If Len(Trim(strNomTemporal)) > 0 Then
        Historial_Precios_Productos.dxDBGrid.Dataset.Close
        csql = "drop table " & strNomTemporal
        EJECUTA_SENTENCIA csql, StrCn
    End If
    strNomTemporal = "RECURSOS_" & Format(Now, "ddmmyyyy") & "_" & Format(Time, "hhmmss")
    csql = "SELECT distinct IF3ORDEN.F3CODPRO AS F5CODPRO, IF5PLA.F5NOMPRO "
    csql = csql & "INTO " & strNomTemporal & " IN '" & wrutatemp & "templus.mdb" & "' "
    csql = csql & "FROM IF4ORDEN INNER JOIN (IF3ORDEN LEFT JOIN IF5PLA ON IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO) "
    csql = csql & "ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
    csql = csql & "Where (((IF4ORDEN.F4FECEMI) >= #" & Format(pFechaInicio, "mm/dd/yyyy") & "# And (IF4ORDEN.F4FECEMI) <= #" & Format(pFechaFin, "mm/dd/yyyy") & "#) "
    csql = csql & "And ((IF4ORDEN.F4COLOCADA) = True)) "
    'csql = csql & "ORDER BY IF4ORDEN.F3CODPRO"
    EJECUTA_SENTENCIA csql, cconex_dbbancos
    If Sw_Ejecuta_Sentencia = True Then
        csql = "Alter table " & strNomTemporal & " ADD COLUMN GRAFICA YESNO"
        EJECUTA_SENTENCIA csql, StrCn
    End If
    With Historial_Precios_Productos.dxDBGrid
        csql = "select * from " & strNomTemporal & " ORDER BY F5CODPRO"
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = StrCn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
         .KeyField = "F5CODPRO"
    End With
'    Set Rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
'    If Rs.State = 1 Then
'        If Rs.RecordCount > 0 Then
'            Rs.MoveFirst
'            ListRecursos.ListItems.Clear
'            Do While Not Rs.EOF
'                ListRecursos.ListItems.Add , "*" & Rs!f3codpro, Rs!F5NOMPRO & ""
'                Rs.MoveNext
'            Loop
'        End If
'    End If
'    If Rs.State = 1 Then Rs.Close
'    Set Rs = Nothing
    MSPrecios.Title = ""
    Carga_Grafica MSPrecios.Title, pFechaInicio, pFechaFin
End If
End Sub


Public Sub Carga_Grafica(pCodigos_de_Recursos As String, pFechaInicio As Variant, pFechaFin As Variant)
Dim Rs As New ADODB.Recordset
Dim DblAnt As Double
If IsDate(pFechaFin) And IsDate(pFechaFin) And Len(Trim(pCodigos_de_Recursos)) > 0 Then
    DblAnt = 0
    csql = "TRANSFORM Avg(LIPRE.F3PRECOS) AS [El Valor] "
    csql = csql & "SELECT LIPRE.F3CODPRO, LIPRE.F5NOMPRO "
    csql = csql & "FROM ["
    csql = csql & "SELECT IF4ORDEN.F4FECEMI, IF3ORDEN.F3CODPRO, IF5PLA.F5NOMPRO, IF3ORDEN.F3PRECOS "
    csql = csql & "FROM IF4ORDEN INNER JOIN (IF3ORDEN LEFT JOIN IF5PLA ON IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO) "
    csql = csql & "ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
    csql = csql & "Where (((IF4ORDEN.F4FECEMI) >= #" & Format(pFechaInicio, "mm/dd/yyyy") & "# And (IF4ORDEN.F4FECEMI) <= #" & Format(pFechaFin, "mm/dd/yyyy") & "#) "
    csql = csql & "And ((IF4ORDEN.F4COLOCADA) = True)) "
    csql = csql & "And IF3ORDEN.F3CODPRO IN (" & pCodigos_de_Recursos & ") "
    csql = csql & "ORDER BY IF4ORDEN.F4FECEMI"
    csql = csql & "]. AS LIPRE "
    csql = csql & "GROUP BY LIPRE.F3CODPRO, LIPRE.F5NOMPRO "
    csql = csql & "PIVOT Format([F4FECEMI],'Short Date')"
    Set Rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    If Rs.State = 1 Then
        If Rs.RecordCount > 0 Then
            With MSPrecios
                .Legend.Location.Visible = True
                .ColumnCount = Rs.RecordCount
                
                .RowCount = Rs.Fields.Count - 2
                
                If .RowCount = 1 Then
                    .ChartType = VtChChartType2dBar
                Else
                    .ChartType = VtChChartType2dLine
                End If
                
                
                For i = 2 To Rs.Fields.Count - 1
                    .Row = i - 1
                    .RowLabel = Rs.Fields(i).Name & ""
                Next
                
                Rs.MoveFirst
                R = 1
                Do While Not Rs.EOF
                    
                    .Column = R '.RowLabel = Rs.Name & ""
                    .ColumnLabel = Rs(1).Value & ""
                    For i = 2 To Rs.Fields.Count - 1
                        .Row = i - 1
                        
                        If Val(Rs(i).Value & "") > 0 Then
                            DblAnt = Val(Rs(i).Value & "")
                            .Data = Val(Rs(i).Value & "")
                        Else
                            If DblAnt > 0 Then
                                .Data = DblAnt
                            End If
                        End If
                        
                    Next
                    R = R + 1
                    Rs.MoveNext
                Loop
                .Refresh
            End With
        Else
            MSPrecios.ColumnCount = 0
            MSPrecios.RowCount = 0
        End If
    Else
        MSPrecios.ColumnCount = 0
        MSPrecios.RowCount = 0
    End If
    If Rs.State = 1 Then Rs.Close
    Set Rs = Nothing
Else
    MSPrecios.ColumnCount = 0
    MSPrecios.RowCount = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Trim(strNomTemporal)) > 0 Then
    
    Historial_Precios_Productos.dxDBGrid.Dataset.Close
    csql = "drop table " & strNomTemporal
    EJECUTA_SENTENCIA csql, StrCn
End If
End Sub

Private Sub ListRecursos_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
Dim StrCodigo1 As String
Dim StrCodigo2 As String
StrCodigo1 = "'" & Mid(ITEM.Key, 2) & "'"
StrCodigo2 = ",'" & Mid(ITEM.Key, 2) & "'"
If ITEM.Checked = True Then
    If Len(Trim(MSPrecios.Title)) = 0 Then
        MSPrecios.Title = MSPrecios.Title & StrCodigo1
    Else
        MSPrecios.Title = MSPrecios.Title & StrCodigo2
    End If
Else
    MSPrecios.Title = Replace(MSPrecios.Title, StrCodigo2, "")
    MSPrecios.Title = Replace(MSPrecios.Title, StrCodigo1, "")
    If InStr(MSPrecios.Title, "'") = 2 Then
        MSPrecios.Title = Replace(MSPrecios.Title, ",", "")
    End If
End If

Carga_Grafica MSPrecios.Title, aboDesde.Value, abohasta.Value
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
 Select Case Tool.Id
        Case "ID_Procesar"
            Carga_Grafica MSPrecios.Title, aboDesde.Value, abohasta.Value
        Case "ID_Salir"
            Unload Me
End Select
End Sub

Private Sub TTabDock1_PanelResize(ByVal Panel As TabDock.TTabDockHost)
Form_Resize
End Sub
