VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmReportAlm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reporte de Almacenes :::"
   ClientHeight    =   4920
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7410
   Icon            =   "FrmReportAlm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   816
      Left            =   4896
      TabIndex        =   11
      Top             =   630
      Width           =   2436
      Begin VB.TextBox TxtStkCom 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   288
         Left            =   1224
         MaxLength       =   2
         TabIndex        =   14
         Top             =   468
         Width           =   888
      End
      Begin VB.TextBox TxtStkDis 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   288
         Left            =   1224
         TabIndex        =   12
         Top             =   168
         Width           =   888
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Stk. Comp. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   204
         TabIndex        =   15
         Top             =   492
         Width           =   924
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Stk. Dispo. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   204
         TabIndex        =   13
         Top             =   192
         Width           =   912
      End
   End
   Begin VB.Frame Frame3 
      Height          =   810
      Left            =   90
      TabIndex        =   7
      Top             =   630
      Width           =   4728
      Begin VB.TextBox TxtCodProd 
         Height          =   288
         Left            =   936
         TabIndex        =   8
         Top             =   312
         Width           =   465
      End
      Begin Threed.SSPanel PnlProd 
         Height          =   492
         Left            =   1428
         TabIndex        =   9
         Top             =   216
         Width           =   3192
         _Version        =   65536
         _ExtentX        =   5630
         _ExtentY        =   868
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   108
         TabIndex        =   10
         Top             =   336
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   612
      Left            =   84
      TabIndex        =   0
      Top             =   0
      Width           =   4740
      Begin VB.TextBox txtalmacen 
         Height          =   288
         Left            =   960
         MaxLength       =   2
         TabIndex        =   1
         Top             =   216
         Width           =   465
      End
      Begin Threed.SSPanel pnlalmacen 
         Height          =   288
         Left            =   1452
         TabIndex        =   2
         Top             =   216
         Width           =   3180
         _Version        =   65536
         _ExtentX        =   5609
         _ExtentY        =   508
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   168
         TabIndex        =   3
         Top             =   228
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   612
      Left            =   4884
      TabIndex        =   4
      Top             =   0
      Width           =   2436
      Begin aBoxCtl.aBox aboFecha 
         Height          =   336
         Left            =   864
         TabIndex        =   5
         Top             =   180
         Width           =   1392
         _ExtentX        =   2461
         _ExtentY        =   582
         ABoxType        =   ""
         MinValue        =   "D10000101"
         MaxValue        =   "D99991231"
         ABoxStyle       =   2
         AlignmentVertical=   2
         HideSelection   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusSelect     =   -1  'True
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "27/04/2007"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         NumberFormat    =   17
         DecimalPlaces   =   0
         HotAppearance   =   2
         CalendarTrailingForeColor=   -2147483629
         BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowButton      =   1
         ButtonPicture   =   "FrmReportAlm.frx":000C
         ButtonWidth     =   22
         UpDownWidth     =   14
         NullText        =   ""
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalcBtnHotStyle =   4
         CalcBackColor   =   -2147483643
         CalcBtnBackColor=   -2147483643
         CalcBtnDigitColor=   -2147483646
         CalcBtnFuntionColor=   8388736
         CalcDisplayFrameColor=   65535
         CalcHeaderBackColor=   -2147483646
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   204
         TabIndex        =   6
         Top             =   228
         Width           =   612
      End
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   336
      Left            =   6360
      TabIndex        =   16
      Top             =   1524
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   593
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxImpDatos 
      Height          =   2715
      Left            =   90
      OleObjectBlob   =   "FrmReportAlm.frx":035E
      TabIndex        =   18
      Top             =   1950
      Width           =   7230
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   336
      Left            =   5364
      TabIndex        =   17
      Top             =   1524
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   593
      _StockProps     =   78
      Caption         =   "&Consultar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "FrmReportAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Formato()
    dxImpDatos.Columns(0).Width = 50
    dxImpDatos.Columns(0).Caption = "Fecha"
    dxImpDatos.Columns(1).Width = 30
    dxImpDatos.Columns(1).Caption = "T/D"
    dxImpDatos.Columns(2).Width = 50
    dxImpDatos.Columns(2).Caption = "Número"
    dxImpDatos.Columns(3).Width = 100
    dxImpDatos.Columns(3).Caption = "Origen"
    dxImpDatos.Columns(4).Width = 50
    dxImpDatos.Columns(4).Caption = "T. Doc."
    dxImpDatos.Columns(5).Width = 70
    dxImpDatos.Columns(5).Caption = "Número Doc."
    dxImpDatos.Columns(6).Width = 100
    dxImpDatos.Columns(6).Caption = "cantidad"

End Sub

Private Sub abofecha_LostFocus()
If aboFecha.Value < Date Then
    MsgBox "La fecha que esta indicando es menor a la actual, el stock a mostrar no es real..Cuidado", vbInformation + vbSystemModal, "Mensaje: Validación"
Else
    MsgBox "No puede hacer un corte a una fecha mayor a la actual...Verifique los datos", vbInformation + vbSystemModal, "Mensaje: Validación"
    aboFecha.Value = Date
End If
End Sub

Private Sub cmdaceptar_Click()
Dim StrSql As String
StrSql = ""
'StrSQL = StrSQL & "TRANSFORM Sum(IF3VALES.F3CANPRO) As Stock"
'StrSQL = StrSQL & " SELECT IF5PLA.F5NOMPRO AS Producto "
'StrSQL = StrSQL & " FROM IF5PLA INNER JOIN (IF4VALES INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = " & _
'"IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = " & _
'"IF3VALES.F2CODALM) "
'
'StrSQL = StrSQL & "WHERE IF4VALES.F4FECVAL <= #" & aboFecha.Value & "# AND IF4VALES.F1CODORI)='XV1'"
'
'StrSQL = StrSQL & "GROUP BY IF3VALES.F2CODALM"
'StrSQL = StrSQL & "ORDER BY IF5PLA.F5NOMPRO"
'StrSQL = StrSQL & "PIVOT IF3VALES.F2CODALM);"

If txtalmacen.Text <> "" Then
    If TxtCodProd.Text <> "" Then
        StrSql = "TRANSFORM Format(sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) ,'0.00') " & _
        " As CANTIDAD SELECT Left(IF5PLA.F5NOMPRO,2) As Grupo, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED as F5NOMPRO FROM IF5PLA INNER JOIN " _
        & " (IF4VALES INNER JOIN IF3VALES ON IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL AND IF4VALES.F2CODALM = " _
        & " IF3VALES.F2CODALM) ON (IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " _
        & " WHERE IF4VALES.F4FECVAL <= #" & Format(aboFecha.Value, "MM/DD/YYYY") & "# " _
        & " AND IF3VALES.F2CODALM = '" & txtalmacen.Text & "' And IF5PLA.F5CODPRO = '" & TxtCodProd.Text & "' GROUP BY " _
        & " Left(IF5PLA.F5NOMPRO,2), IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED ORDER BY " & _
          " Left(IF5PLA.F5NOMPRO,2) PIVOT IF3VALES.F2CODALM;"
    Else
        StrSql = "TRANSFORM Format(sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) ,'0.00') " & _
        " As CANTIDAD SELECT Left(IF5PLA.F5NOMPRO,2) As Grupo, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED as F5NOMPRO FROM IF5PLA INNER JOIN " _
        & " (IF4VALES INNER JOIN IF3VALES ON IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL AND IF4VALES.F2CODALM = " _
        & " IF3VALES.F2CODALM) ON (IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " _
        & " WHERE IF4VALES.F4FECVAL <= #" & Format(aboFecha.Value, "MM/DD/YYYY") & "# " _
        & " AND IF3VALES.F2CODALM = '" & txtalmacen.Text & "' GROUP BY " _
        & " Left(IF5PLA.F5NOMPRO,2), IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED ORDER BY " & _
          " Left(IF5PLA.F5NOMPRO,2) PIVOT IF3VALES.F2CODALM;"
    End If
Else
    If TxtCodProd.Text <> "" Then
        StrSql = "TRANSFORM Format(Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)),'0.00') As Stock SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED as F5NOMPRO FROM IF5PLA INNER JOIN " _
        & " (IF4VALES INNER JOIN IF3VALES ON IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL AND IF4VALES.F2CODALM = " _
        & " IF3VALES.F2CODALM) ON (IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " _
        & " WHERE IF4VALES.F4FECVAL <= #" & Format(aboFecha.Value, "MM/DD/YYYY") & "# " _
        & " AND IF5PLA.F5CODPRO = '" & TxtCodProd.Text & "' GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED " _
        & " ORDER BY  IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED PIVOT IF3VALES.F2CODALM;"
    Else
        StrSql = "TRANSFORM Format(Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)),'0.00') As Stock SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED as F5NOMPRO FROM IF5PLA INNER JOIN " _
        & " (IF4VALES INNER JOIN IF3VALES ON IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL AND IF4VALES.F2CODALM = " _
        & " IF3VALES.F2CODALM) ON (IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " _
        & " WHERE IF4VALES.F4FECVAL <= #" & Format(aboFecha.Value, "MM/DD/YYYY") & "# " _
        & " GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED " _
        & " ORDER BY IF5PLA.F5NOMPRO + '  '  + IF5PLA.F7CODMED PIVOT IF3VALES.F2CODALM;"
    End If
End If
If SSCommand1.Tag = "" Then
    If txtalmacen.Text <> "" Then
        If txtalmacen.Text = "01" Then
            acr_StkAlm_novalorizado.fldsaldo.DataField = ""
        Else
            acr_StkAlm_novalorizado.Field14.DataField = ""
        End If
    End If
    
    Screen.MousePointer = 11
    acr_StkAlm_novalorizado.lblempresa.Caption = wempresa
    acr_StkAlm_novalorizado.Label4.Caption = acr_StkAlm_novalorizado.Label4.Caption & IIf(txtalmacen.Text = "", "En todos los Almacenes", "En el Almacen " & txtalmacen.Text) & " - " & pnlalmacen.Caption
    
    acr_StkAlm_novalorizado.fldfecha.Text = Format(aboFecha.Value, "DD/MM/YYYY")
    
    acr_StkAlm_novalorizado.datos.ConnectionString = cnn_dbbancos
    acr_StkAlm_novalorizado.datos.Source = StrSql
    acr_StkAlm_novalorizado.Show 1
    Screen.MousePointer = 0
    'dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    'dxDBGrid1.Dataset.Active = False
    'dxDBGrid1.Dataset.ADODataset.CommandType = cmdText
    'dxDBGrid1.Dataset.ADODataset.CommandText = StrSQL
    'dxDBGrid1.Dataset.Active = True
    'MsgBox cnn_dbbancos
Else
    If rs.State = 1 Then rs.Close
    
    rs.Open StrSql, cnn_dbbancos, adOpenDynamic, adLockReadOnly
    If txtalmacen.Text <> "" Then TxtStkDis.Text = rs(txtalmacen.Text) Else _
    TxtStkDis.Text = Val(rs("01")) + Val(rs("40"))
End If
End Sub

Private Sub Form_Load()
aboFecha.Value = Format(Date, "DD/MM/YYYY")
End Sub

Private Sub SSCommand1_Click()
Dim sql As String

If wcodproducto <> "" And wcod_alm <> "" Then
    sql = "SELECT IF3VALES.F4FECVAL, Left(IF4VALES.[F4NUMVAL],1) AS TD, " & _
    " Mid(IF4VALES.[F4NUMVAL],3,Len(IF4VALES.[F4Numval])-2) AS Numero, SF1ORIGENES.F1NOMORI, " & _
    " IF4VALES.F1CODDOC, IF4VALES.F4NUMFAC, Sum([F3CANPRO]) AS Cantidad" & _
    " FROM SF1ORIGENES INNER JOIN (IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = " & _
    " IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON SF1ORIGENES.F1CODORI " & _
    " = IF4VALES.F1CODORI Where IF3VALES.F5CODPRO = '" & wcodproducto & "' And If4VAlES.F2CODALM = '" & wcod_alm & "' Group By " & _
    " IF3VALES.F4FECVAL, Left(IF4VALES.[F4NUMVAL],1) , " & _
    " Mid(IF4VALES.[F4NUMVAL],3,Len(IF4VALES.[F4Numval])-2), SF1ORIGENES.F1NOMORI, " & _
    " IF4VALES.F1CODDOC, IF4VALES.F4NUMFAC;"
Else
    sql = "SELECT IF3VALES.F4FECVAL, Left(IF4VALES.[F4NUMVAL],1) AS TD, " & _
    " Mid(IF4VALES.[F4NUMVAL],3,Len(IF4VALES.[F4Numval])-2) AS Numero, SF1ORIGENES.F1NOMORI, " & _
    " IF4VALES.F1CODDOC, IF4VALES.F4NUMFAC, Sum([F3CANPRO]) AS Cantidad" & _
    " FROM SF1ORIGENES INNER JOIN (IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = " & _
    " IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON SF1ORIGENES.F1CODORI " & _
    " = IF4VALES.F1CODORI Group By " & _
    " IF3VALES.F4FECVAL, Left(IF4VALES.[F4NUMVAL],1) , " & _
    " Mid(IF4VALES.[F4NUMVAL],3,Len(IF4VALES.[F4Numval])-2), SF1ORIGENES.F1NOMORI, " & _
    " IF4VALES.F1CODDOC, IF4VALES.F4NUMFAC;"
End If
    dxImpDatos.Dataset.Active = False
    dxImpDatos.DefaultFields = True
    dxImpDatos.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    dxImpDatos.Dataset.ADODataset.CommandText = sql
    dxImpDatos.Dataset.Active = True
    dxImpDatos.KeyField = "F4FECVAL"
    dxImpDatos.Dataset.Refresh
    Call Formato
    SSCommand1.Tag = "PR"
    Call cmdaceptar_Click
    SSCommand1.Tag = ""
End Sub

Private Sub txtalmacen_Change()
    
    pnlalmacen.Caption = ""
    
End Sub

Private Sub txtalmacen_DblClick()

    txtalmacen_KeyDown 113, 0
    
End Sub

Private Sub txtalmacen_GotFocus()

    txtalmacen.SelStart = 0
    txtalmacen.SelLength = Len(txtalmacen.Text)

End Sub

Private Sub txtalmacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            txtalmacen.Text = wcod_alm
            pnlalmacen.Caption = wnomalmacen
            Call SSCommand1_Click
            txtalmacen_KeyPress 13
        Else
            TxtStkDis.Text = "0"
            TxtStkCom.Text = "0"
        End If
    End If
    
End Sub

Private Sub txtalmacen_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        optdolares.SetFocus
'    End If

End Sub

Private Sub txtalmacen_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtalmacen.Text)) > 0 Then
            wnomalmacen = ""
            If VALIDA_ALMACEN(txtalmacen.Text) = True Then
                pnlalmacen.Caption = wnomalmacen
            Else
                MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                txtalmacen.SetFocus
            End If
        Else
            pnlalmacen.Caption = "TODOS LOS ALMACENES"
        End If
    End If

End Sub

Private Sub TxtCodProd_DblClick()
Call TxtCodProd_KeyDown(vbKeyF2, 1)
End Sub

Private Sub TxtCodProd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    wcodproducto = TxtCodProd.Text
    sw_ayuda_prod = True
    wtipoguia = "S"
    ayuda_productos.Show 1
    If wcodproducto = "" Then
        TxtCodProd.Text = ""
        PnlProd.Caption = "TODOS LOS PRODUCTOS"
        TxtStkDis.Text = "0"
        TxtStkCom.Text = "0"
        
    Else
        TxtCodProd.Text = wcodproducto
        PnlProd.Caption = wdesproducto
        Call SSCommand1_Click
    End If
    
End If
End Sub

Private Sub TxtCodProd_LostFocus()
If TxtCodProd.Text = "" Then
    PnlProd.Caption = "TODOS LOS PRODUCTOS"
End If
End Sub
