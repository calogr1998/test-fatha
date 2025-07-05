VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form AnaliMovMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reporte de Ingresos por movimientos :::"
   ClientHeight    =   855
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5775
   Icon            =   "FrmAnaliMovMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   5604
      Begin aBoxCtl.aBox aboFecha1 
         Height          =   336
         Left            =   720
         TabIndex        =   2
         Top             =   252
         Width           =   1320
         _ExtentX        =   2328
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
         Text            =   "17/03/2007"
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
         ButtonPicture   =   "FrmAnaliMovMes.frx":014A
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
      Begin aBoxCtl.aBox aboFecha2 
         Height          =   336
         Left            =   2796
         TabIndex        =   4
         Top             =   252
         Width           =   1296
         _ExtentX        =   2275
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
         Text            =   "17/03/2007"
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
         ButtonPicture   =   "FrmAnaliMovMes.frx":049C
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
      Begin Threed.SSCommand cmdaceptar 
         Height          =   330
         Left            =   4230
         TabIndex        =   5
         Top             =   255
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2074
         _ExtentY        =   593
         _StockProps     =   78
         Caption         =   "&Procesar..."
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   2196
         TabIndex        =   3
         Top             =   300
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   132
         TabIndex        =   1
         Top             =   300
         Width           =   528
      End
   End
End
Attribute VB_Name = "AnaliMovMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
'SQL = "SELECT Left([IF4VALES].[F4NUMVAL],1) AS [IS], IF5PLA.F5CTACON1 as CTA, TBGRUPOS.NOMGRUPO as Grupo, " & _
'" IF4VALES.F1CODORI as Descripcion, Sum(IF3VALES.F3VALDOL) AS DOLARES,Sum(iif(Left([IF4VALES].[F4NUMVAL],1) =) 'I', " & _
'" IF3VALES.F3VALVTA,0),Sum(iif(IF3VALES.F3VALDOL) AS DOLARES, Sum(iif(Left([IF4VALES].[F4NUMVAL],1) =) 'S', " & _
'" IF3VALES.F3VALVTA,0), iif(Left([IF4VALES].[F4NUMVAL],1) = 'S',IF3VALES.F3VALDOL,0) as Dolares ,IF3VALES.F4FECVAL " & _
'"FROM IF4VALES INNER JOIN ((IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) INNER JOIN TBGRUPOS ON IF5PLA.F5CTACON1 = TBGRUPOS.F5CODCTA) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) " & _
'"GROUP BY Left([IF4VALES].[F4NUMVAL],1), IF5PLA.F5CTACON1, TBGRUPOS.NOMGRUPO, IF4VALES.F1CODORI, IF3VALES.F4FECVAL;"

'SQL = "SELECT  Left([IF4VALES].[F4NUMVAL],1) AS [IS], IF5PLA.F5CTACON1 as CTA, " & _
'        " TBGRUPOS.NOMGRUPO as Grupo,  " & _
'        " SF1ORIGENES.F1NOMORI as Descripcion," & _
'        " Sum(iif(Left([IF4VALES].[F4NUMVAL],1) = 'I', IF3VALES.F3VALDOL,0)) AS DOLARESENT," & _
'        " Sum(iif(Left([IF4VALES].[F4NUMVAL],1) = 'S', IF3VALES.F3VALDOL,0)) AS DOLARESSAL," & _
'        " Sum(iif(Left([IF4VALES].[F4NUMVAL],1) = 'I', IF3VALES.F3VALVTA,0)) AS SOLESENT, " & _
'        " Sum(iif(Left([IF4VALES].[F4NUMVAL],1) = 'S', IF3VALES.F3VALVTA,0)) AS SOLESSAL, " & _
'        " IF3VALES.F4FECVAL " & _
'        " FROM IF4VALES INNER JOIN ((IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = " & _
'        " IF3VALES.F5CODPRO) INNER JOIN TBGRUPOS ON IF5PLA.F5CTACON1 = TBGRUPOS.F5CODCTA) " & _
'        " INNER JOIN SF1ORIGENES ON ( IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) ON (IF4VALES" & _
'        ".F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)" & _
'        " GROUP BY " & _
'        " Left([IF4VALES].[F4NUMVAL],1), IF5PLA.F5CTACON1, TBGRUPOS.NOMGRUPO, " & _
'        " IF4VALES.F1CODORI, IF3VALES.F4FECVAL;"

'Dim con1 As String
'con1 = "(SELECT IIf(Month([IF4VALES].[F4FECVAL])=1,'Enero',IIf(Month([IF4VALES].[F4FECVAL]) " & _
'" =2,'Febrero','Marzo')) AS MESES, [IF4VALES].[F1CODORI], [IF3VALES].[F5CODPRO],  " & _
'" [IF3VALES].[F3CANPRO], [F5NOMPRO]+' '+[F7CODMED] AS Producto, [SF1ORIGENES].[F1NOMORI], " & _
'" [IF5PLA].[F7CODMED] FROM (IF4VALES INNER JOIN SF1ORIGENES ON [IF4VALES].[F1CODORI]= " & _
'" [SF1ORIGENES].[F1CODORI]) INNER JOIN (IF3VALES INNER JOIN IF5PLA ON [IF3VALES].[F5CODPRO] " & _
'" =[IF5PLA].[F5CODPRO]) ON ([IF4VALES].[F2CODALM]=[IF3VALES].[F2CODALM]) AND ([IF4VALES].[F4NUMVAL]" & _
'" =[IF3VALES].[F4NUMVAL]))"
'
'SQL = "TRANSFORM Sum(Consulta1.F3CANPRO) AS SumaDeF3CANPRO SELECT Consulta1.F5CODPRO AS " & _
'" Codigo, Consulta1.Producto, Consulta1.F1NOMORI AS Origen, Sum(Consulta1.F3CANPRO) AS " & _
'" Total FROM " & con1 & " as Consulta1 GROUP BY Consulta1.F5CODPRO, Consulta1.Producto, Consulta1.F1NOMORI " & _
'" ORDER BY Consulta1.Producto, Consulta1.F1NOMORI PIVOT Consulta1.[Meses];"


'MsgBox SQL
'
'If rs.State = 1 Then rs.Close
'rs.Open SQL, cnn_dbbancos, adOpenDynamic, adLockReadOnly
'Dim rtemp As New Recordset
'rtemp.Fields.Append "Cuenta", adVarChar, 15
'rtemp.Fields.Append "Descripcion", adVarChar, 60
'rtemp.Fields.Append "EntradasD", adVarChar, 15
'rtemp.Fields.Append "SalidasD", adVarChar, 15
'rtemp.Fields.Append "SALDOFIND", adVarChar, 15
'rtemp.Fields.Append "EntradasMN", adVarChar, 15
'rtemp.Fields.Append "SalidasMN", adVarChar, 15
'rtemp.Fields.Append "SALDOFINMN", adVarChar, 15
'rtemp.Open
'
'Do While rs.EOF = False
'    rtemp.AddNew
'    rtemp("Cuenta") = rs("F5CTACON1")
'    rtemp("Descripcion") = rs("NOMGRUPO")
'    If rs("IS") = "I" Then
'        rtemp("EntradasMN") = rs("SOLES")
'        rtemp("EntradasMN") = rs("DOLARES")
'    Else
'        rtemp("EntradasD") = rs("SOLES")
'        rtemp("SalidasD") = rs("DOLARES")
'    End If
'    rtemp.Update
'    rs.MoveNext
'Loop
'acr_AnaliMovMes.datos.ConnectionString = cnn_dbbancos
'acr_AnaliMovMes.datos.Source = rtemp
'acr_AnaliMovMes.Show 1

SQL = "SELECT BF9GIN.CUENTA, BF9GIN.NOMBRE, (SELECT Sum(IIf(Left([If4Vales].[f4numval],1)='I',[F3CANPRO] * [F3VALVTA],([F3CANPRO] * [F3VALVTA])* -1))" & _
" FROM IF4VALES INNER JOIN (BF9GIN BF9 INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
" ON BF9.CUENTA = IF5PLA.F5CTACON1) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = " & _
" IF3VALES.F2CODALM) Where BF9.CUENTA = BF9GIN.CUENTA AND IF4VALES.F4FECVAL < CVDATE('" & aboFecha1.Value & _
"')) As SaldoAnteriorSOL, (SELECT Sum(IIf(Left([If4Vales].[f4numval],1)='I',[F3CANPRO] * [F3VALDOL],([F3CANPRO] * [F3VALDOL])* -1))" & _
" FROM IF4VALES INNER JOIN (BF9GIN BF9 INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
" ON BF9.CUENTA = IF5PLA.F5CTACON1) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = " & _
" IF3VALES.F2CODALM) Where BF9.CUENTA = BF9GIN.CUENTA AND IF4VALES.F4FECVAL < CVDATE('" & aboFecha1.Value & _
"')) As SaldoAnteriorDOl, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI," & _
" Sum(IIf(Left([If4Vales].[F4numval],1)='I',[F3CANPRO]*[F3VALVTA],0)) AS SolesEnt, " & _
" Sum(IIf(Left([If4vales].[F4numval],1)='I',[F3CANPRO]*[F3VALDOL],0)) AS DolaresEnt, " & _
" Sum(IIf(Left([If4Vales].[F4numval],1)='S',([F3CANPRO]*[F3VALVTA]),0)) AS SolesSal, " & _
" Sum(IIf(Left([If4vales].[F4numval],1)='S',([F3CANPRO]*[F3VALDOL]),0)) AS DolaresSal " & _
" FROM (SF1ORIGENES INNER JOIN IF4VALES ON SF1ORIGENES.F1CODORI = IF4VALES.F1CODORI) " & _
" INNER JOIN ((BF9GIN INNER JOIN IF5PLA ON BF9GIN.CUENTA = IF5PLA.F5CTACON1) INNER JOIN " & _
" IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) " & _
" AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) WHERE IF4VALES.F4FECVAL BETWEEN CVDATE('" & aboFecha1.Value & _
"') AND CVDATE('" & aboFecha2.Value & "') GROUP BY BF9GIN.CUENTA, BF9GIN.NOMBRE, " & _
" IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI;"

acr_AnaliMovMes.datos.ConnectionString = cnn_dbbancos
acr_AnaliMovMes.datos.Source = SQL
acr_AnaliMovMes.Show 1
End Sub

Private Sub Form_Load()
aboFecha1.Value = Date
aboFecha2.Value = Date
End Sub
