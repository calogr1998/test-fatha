VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{F7E69521-3C28-11D2-B3E7-00AA00B42B7C}#3.1#0"; "FPTAB30.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form ProcesaVale 
   Caption         =   "Ventas "
   ClientHeight    =   7230
   ClientLeft      =   1065
   ClientTop       =   1035
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9525
   Begin TabproADOLib.fpTabProADO fpTabProADO1 
      Height          =   6315
      Left            =   180
      TabIndex        =   6
      Top             =   810
      Width           =   9150
      _Version        =   196609
      _ExtentX        =   16140
      _ExtentY        =   11139
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   0
      TabCount        =   2
      Tab             =   1
      AlignTextH      =   1
      AlignTextV      =   1
      ThreeD          =   -1  'True
      ShowFocusRect   =   0   'False
      ApplyTo         =   2
      ActiveTabBold   =   0   'False
      AlignPictureV   =   2
      OffsetFromClientTop=   -1  'True
      ChamferedWidth  =   2
      ChamferedHeight =   2
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      BookRingShowHole=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      ThreeDInnerWidthActive=   1
      ActiveTabInflate=   2
      DrawFocusRect   =   1
      DataField       =   ""
      DataMember      =   ""
      TabCaption      =   "ProcesaVale.frx":0000
      PageEarMarkPictureNext=   "ProcesaVale.frx":067F
      PageEarMarkPicturePrev=   "ProcesaVale.frx":069B
      EarMarkPictureNext=   "ProcesaVale.frx":06B7
      EarMarkPicturePrev=   "ProcesaVale.frx":06D3
      Begin VB.Frame Moneda 
         Caption         =   "Moneda"
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   -24014
         TabIndex        =   11
         Top             =   -16049
         Width           =   8880
         Begin VB.OptionButton optmoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   2205
            TabIndex        =   3
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optmoneda 
            Caption         =   "D�lares"
            Height          =   195
            Index           =   1
            Left            =   5670
            TabIndex        =   4
            Top             =   270
            Width           =   1140
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4875
         Left            =   -23789
         OleObjectBlob   =   "ProcesaVale.frx":06EF
         TabIndex        =   7
         Top             =   -21134
         Width           =   8565
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   5460
         Left            =   225
         OleObjectBlob   =   "ProcesaVale.frx":3F92
         TabIndex        =   9
         Top             =   630
         Width           =   8655
      End
   End
   Begin VB.TextBox txttc 
      Height          =   285
      Left            =   8640
      TabIndex        =   2
      Top             =   360
      Width           =   645
   End
   Begin VB.TextBox txtnomtienda 
      BackColor       =   &H80000004&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   405
      Width           =   3975
   End
   Begin VB.TextBox txtcodtienda 
      Height          =   285
      Left            =   1005
      TabIndex        =   0
      Top             =   405
      Width           =   645
   End
   Begin aBoxCtl.aBox txtfecha 
      Height          =   315
      Left            =   6435
      TabIndex        =   1
      Top             =   360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
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
      Text            =   "14/01/2003"
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
      ButtonPicture   =   "ProcesaVale.frx":8173
      ButtonWidth     =   21
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
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "ProcesaVale.frx":84C5
      ToolBars        =   "ProcesaVale.frx":C415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "T.Cambio"
      Height          =   240
      Left            =   7830
      TabIndex        =   12
      Top             =   405
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5895
      TabIndex        =   10
      Top             =   405
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tienda"
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   450
      Width           =   855
   End
End
Attribute VB_Name = "ProcesaVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Values()            As Variant
Dim amovs_cab(0 To 6) As a_grabacion
Dim amovs_det(0 To 5) As a_grabacion
Dim ctipo               As String * 1
Dim cvalores            As String
Dim cmes                As String * 2
Dim RSDETALLE           As New ADODB.Recordset
Dim nfil                As Integer
Dim rsdetp              As New ADODB.Recordset
Dim CCNUMVALE           As String
Dim avales_cab(0 To 15) As a_grabacion
Dim avales_det(0 To 12) As a_grabacion

Dim WSQL            As String
Dim WSQL1           As String
Dim wnomtablatemp   As String
Dim sw_leetemp      As Boolean

Private Sub GENERA_VALE_SALIDA()
Dim csql            As String
Dim rsconsulta      As New ADODB.Recordset
Dim ntc             As Double
Dim rscontrol       As New ADODB.Recordset
Dim cconcepto_venta  As String
        
    rscambios.Open "SELECT * FROM CAMBIOS WHERE FECHA= CVDate( '" & txtfecha.Value & "' )", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rscambios.EOF Then
        ntc = Val("" & rscambios.Fields("CAMBIO"))
    Else
        ntc = 0#
    End If
    rscambios.Close
    
    cconcepto_venta = ""
    rscontrol.Open "SELECT * FROM SF1PARAIN WHERE F1CODEMP ='" & wempresa & "'", cnn_control, adOpenDynamic, adLockOptimistic
    If Not rscontrol.EOF Then
        cconcepto_venta = rscontrol.Fields("F1CONC_VENTA") & ""
    End If
    rscontrol.Close
    
    'sw_nuevo_documento = True
'    csql = "SELECT VENTASMATERIA.ALMACEN From VENTASMATERIA GROUP BY VENTASMATERIA.ALMACEN"
'    rsconsulta.Open csql, cnn_form, adOpenDynamic, adLockOptimistic
'    If Not rsconsulta.EOF Then
'        rsconsulta.MoveFirsif5pla
'        Do While Not rsconsulta.EOF
            '---- OJO :  CODIGO CONCEPTO DEBE SER PARAMETRO
            GENERA_VALE gcodalm, txtfecha.Value, cconcepto_venta, ntc, "S", "S", ""
'            rsconsulta.MoveNext
'        Loop
'    End If
'    rsconsulta.Close

End Sub

Private Function GENERA_NUMVALE(palmacen As String, pmes As String, ptipo As String)
Dim cnumvale    As String

    rsalmacen.Open "SELECT * FROM EF2ALMACENES WHERE F2CODALM='" & palmacen & "'", cnn_dbbancos
    If Not rsalmacen.EOF Then
        If ptipo = "I" Then
            cnumvale = Mid(rsalmacen.Fields("F1VALING" & pmes) & "", 1, 4) & Format(Val(Mid(rsalmacen.Fields("F1VALING" & pmes) & "", 5, 4)) + 1, "0000")
        Else
            cnumvale = Mid(rsalmacen.Fields("F1VALSAL" & pmes) & "", 1, 4) & Format(Val(Mid(rsalmacen.Fields("F1VALSAL" & pmes) & "", 5, 4)) + 1, "0000")
        End If
    End If
    rsalmacen.Close
    GENERA_NUMVALE = cnumvale
    
End Function

Private Sub GENERA_VALE(palmacen As String, pfecha As Date, pconcepto As String, ptipcamb As Double, pmoneda As String, ptipo As String, pnumdoc As String)
Dim cnumvale        As String
Dim ccampo          As String

    If sw_nuevo_documento = True Then
        cnumvale = GENERA_NUMVALE(palmacen, Format(Month(pfecha), "00"), "S")
        CCNUMVALE = cnumvale
        'txtnumero.Text = cnumvale
        ctipo = "A"
    Else
        cnumvale = txtnumero.Text
        CCNUMVALE = cnumvale
        ctipo = "M"
    End If
    
    '-------------------------------------------------------
    '------------------------- ASIGNA DATOS DE LA CABECERA
    avales_cab(0).campo = "F4NUMVAL": avales_cab(0).valor = cnumvale: avales_cab(0).TIPO = "T"
    avales_cab(1).campo = "F2CODALM": avales_cab(1).valor = palmacen: avales_cab(1).TIPO = "T"
    avales_cab(2).campo = "F4FECVAL": avales_cab(2).valor = pfecha: avales_cab(2).TIPO = "F"
    avales_cab(3).campo = "F1CODORI": avales_cab(3).valor = pconcepto: avales_cab(3).TIPO = "T"
    avales_cab(4).campo = "F4TIPCAM": avales_cab(4).valor = ptipcamb: avales_cab(4).TIPO = "N"
    avales_cab(5).campo = "F2CODPROV": avales_cab(5).valor = "": avales_cab(5).TIPO = "T"
    avales_cab(6).campo = "F4CENTRO": avales_cab(6).valor = "": avales_cab(6).TIPO = "T"
    avales_cab(7).campo = "F4MONEDA": avales_cab(7).valor = pmoneda: avales_cab(7).TIPO = "T"
    avales_cab(8).campo = "F4SERGUIA": avales_cab(8).valor = "": avales_cab(8).TIPO = "T"
    avales_cab(9).campo = "F4NUMGUIA": avales_cab(9).valor = "": avales_cab(9).TIPO = "T"
    avales_cab(10).campo = "F4TIPDOC": avales_cab(10).valor = "": avales_cab(10).TIPO = "T"
    avales_cab(11).campo = "F4SERDOC": avales_cab(11).valor = "": avales_cab(11).TIPO = "T"
    avales_cab(12).campo = "F4NUMDOC": avales_cab(12).valor = pnumdoc: avales_cab(12).TIPO = "T"
    If ctipo = "A" Then
        avales_cab(13).campo = "F4FECGRA": avales_cab(13).valor = Format(Date, "dd/mm/yyyy"): avales_cab(13).TIPO = "F"
        avales_cab(14).campo = "F4USEGRA": avales_cab(14).valor = wusuario: avales_cab(14).TIPO = "T"
    Else
        avales_cab(13).campo = "F4FECMOD": avales_cab(13).valor = Format(Date, "dd/mm/yyyy"): avales_cab(13).TIPO = "F"
        avales_cab(14).campo = "F4USEMOD": avales_cab(14).valor = wusuario: avales_cab(14).TIPO = "T"
    End If
    avales_cab(15).campo = "F4OBSERVA": avales_cab(15).valor = "": avales_cab(15).TIPO = "T"
    '-------------------------------------------------------
    '------------------------- ASIGNA DATOS DEL DETALLE
    
    avales_det(0).campo = "F4NUMVAL": avales_det(0).valor = "": avales_det(0).TIPO = "T"
    avales_det(1).campo = "F5CODPRO": avales_det(1).valor = "": avales_det(1).TIPO = "T"
    avales_det(2).campo = "F3CANPRO": avales_det(2).valor = "": avales_det(2).TIPO = "N"
    avales_det(3).campo = "F3VALVTA": avales_det(3).valor = "": avales_det(3).TIPO = "N"
    avales_det(4).campo = "F3IGV": avales_det(4).valor = "": avales_det(4).TIPO = "N"
    avales_det(5).campo = "F3TOTITE": avales_det(5).valor = "": avales_det(5).TIPO = "N"
    avales_det(6).campo = "F2CODALM": avales_det(6).valor = "": avales_det(6).TIPO = "T"
    avales_det(7).campo = "F4FECVAL": avales_det(7).valor = "": avales_det(7).TIPO = "F"
    avales_det(8).campo = "F3VALDOL": avales_det(8).valor = "": avales_det(8).TIPO = "N"
    avales_det(9).campo = "F3IGVDOL": avales_det(9).valor = "": avales_det(9).TIPO = "N"
    avales_det(10).campo = "F3TOTDOL": avales_det(10).valor = "": avales_det(10).TIPO = "N"
    avales_det(11).campo = "": avales_det(11).valor = "": avales_det(11).TIPO = "T"
    avales_det(12).campo = "F3GRUPO": avales_det(12).valor = "": avales_det(12).TIPO = "T"
    
    '------------------- CALCULA NUMERO DE FILAS
    nitems = 0
    RSDETALLE.Open "SELECT COUNT(ITEM) AS NITEM FROM VENTASMATERIA WHERE ALMACEN = '" & palmacen & "'", cnn_form
    If Not RSDETALLE.EOF Then
        nitems = Val("" & RSDETALLE.Fields("NITEM"))
    End If
    RSDETALLE.Close
    '---------------------------------------------
    ReDim Values(12, nitems)
    
    RSDETALLE.Open "SELECT * FROM VENTASMATERIA WHERE ALMACEN = '" & palmacen & "'", cnn_form
    If Not RSDETALLE.EOF Then
        nfil = 0
        RSDETALLE.MoveFirst
        Do While Not RSDETALLE.EOF
            Values(0, nfil) = cnumvale
            Values(1, nfil) = RSDETALLE.Fields("CODIGO") & ""
            Values(2, nfil) = RSDETALLE.Fields("CANTIDAD") & ""
            If pmoneda = "S" Then
                Values(3, nfil) = Val(RSDETALLE.Fields("COSTOUNI") & "")
                Values(4, nfil) = Val(RSDETALLE.Fields("IGV") & "")
                Values(5, nfil) = Val(RSDETALLE.Fields("TOTAL") & "")
                
                Values(8, nfil) = Val(Format(Val(RSDETALLE.Fields("COSTOUNI") & "") / Val(Format(txttc.Text, "0.00")), "0.00"))
                Values(9, nfil) = Val(Format(Val(RSDETALLE.Fields("IGV") & "") / Val(Format(txttc.Text, "0.00")), "0.00"))
                Values(10, nfil) = Val(Format(Val(RSDETALLE.Fields("TOTAL") & "") / Val(Format(txttc.Text, "0.00")), "0.00"))
            Else
                Values(8, nfil) = Val(RSDETALLE.Fields("COSTOUNI") & "")
                Values(9, nfil) = Val(RSDETALLE.Fields("IGV") & "")
                Values(10, nfil) = Val(RSDETALLE.Fields("TOTAL") & "")
                
                Values(3, nfil) = Val(Format(Val(RSDETALLE.Fields("COSTOUNI") & "") * Val(Format(txttc.Text, "0.00")), "0.00"))
                Values(4, nfil) = Val(Format(Val(RSDETALLE.Fields("IGV") & "") * Val(Format(txttc.Text, "0.00")), "0.00"))
                Values(5, nfil) = Val(Format(Val(RSDETALLE.Fields("TOTAL") & "") * Val(Format(txttc.Text, "0.00")), "0.00"))
            End If
            Values(6, nfil) = palmacen
            Values(7, nfil) = pfecha
            Values(11, nfil) = ptipo
            Values(12, nfil) = RSDETALLE.Fields("GRUPO") & ""
            RSDETALLE.MoveNext
            nfil = nfil + 1
        Loop
   
    End If
    RSDETALLE.Close
    
    cvalores = "1111111111101"
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    cmes = Format(Month(pfecha), "00")
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        GRABA_REGISTRO avales_cab(), "IF4VALES", ctipo, 15, cnn_dbbancos, ""
        
        If sw_graba_registro = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_DET avales_det(), "IF3VALES", ctipo, 12, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, "A"
        End If
        
        If ptipo = "I" Then
            ccampo = "F1VALING" & cmes
        Else
            ccampo = "F1VALSAL" & cmes
        End If
        ACTUALIZA_ALMA_VALE cnumvale, ccampo, palmacen
        
    Else    '--- Modificaci�n
        
        '-------------------------------------------------------
        '------- GRABA CABECERA
        GRABA_REGISTRO avales_cab(), "IF4VALES", ctipo, 15, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & palmacen & "'"
        '-------------------------------------------------------
        '------- RESTA LOS SALDOS
        rsif3vales.Open "SELECT * FROM IF3VALES WHERE F3NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif3vales.EOF Then
            Do While Not rsif3vales.EOF
                GRABA_SALDO rsif3vales.Fields("F5CODPRO") & "", rsif3vales.Fields("F3CANPRO"), rsif3vales.Fields("F3VALVTA"), cmes, "I", cnn_dbbancos
                rsif3vales.MoveNext
            Loop
        End If
        rsif3vales.Close
        '-------------------------------------------------------
        '------- GRABA DETALLE
        cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F3NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'")
        GRABA_REGISTRO_DET avales_det(), "IF3VALES", "A", 12, cnn_dbbancos, "F3NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & palmacen & "'", Values(), nfil - 1, cvalores, cmes, "A"
    End If
    '-------------------------------------------------------
    '-------------------------------------------------------

End Sub

Private Sub ACTUALIZA_ALMA_VALE(pnumvale As String, pcampo As String, palmacen As String)
Dim csql    As String
        
    csql = "UPDATE EF2ALMACENES SET " & pcampo & " =  '" & pnumvale & "' WHERE '" & pnumvale & "' > " & pcampo & " AND F2CODALM='" & palmacen & "'"
    cnn_dbbancos.Execute csql
    
End Sub

Private Sub Form_Load()
Dim CSQL1   As String

    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1010

    FlagControl = False
    
    sw_nuevo_documento = True
    
    cnombase = wusuario & "PRODUC" & Format(Time, "hh_mm_ss") & ".MDB"
    '--- conexion a la base de datos temporal --------'
    
    CREATEDATABASE_N wrutatemp & "\", cnombase
    
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    
    cnomtabla = "VENTASDIA"
    
    WSQL = "(ITEM TEXT(2),CODIGO TEXT(10),DESCRIPCION TEXT(50),UNIDAD TEXT(3),CANTIDAD DOUBLE,VALVTAS DOUBLE,VALVTAD DOUBLE,NUMVALE TEXT(8),GRUPO TEXT(1))"
    CREATETABLE_N cnomtabla, WSQL, cnn_form
    DELETEREC_N cnomtabla, cnn_form
    
    cnomtabla2 = "VENTASMATERIA"
    
    WSQL1 = "(ITEM TEXT(2),CODIGO TEXT(10),DESCRIPCION TEXT(50),UNIDAD TEXT(3),CANTIDAD DOUBLE,COSTOUNI DOUBLE,IGV DOUBLE,TOTAL DOUBLE,NUMVALE TEXT(8),GRUPO TEXT(1),ALMACEN TEXT(2))"
    CREATETABLE_N cnomtabla2, WSQL1, cnn_form
    DELETEREC_N cnomtabla2, cnn_form
      
    'ANTES CREAR LA TABLA
    wnomtablatemp = "PRODUCTOS"
    
    CSQL1 = "(ITEM TEXT(2),GRUPO TEXT(1),CODPROD TEXT(15),DESCRIPCION TEXT(100)," & _
             "UMEDIDA TEXT(3),MEDIDA DOUBLE,CANTIDAD DOUBLE,TOTAL DOUBLE,ALMACEN TEXT(2),COSTOUNI DOUBLE,IGV DOUBLE,CODORI TEXT(10))"
    
    CREATETABLE_N wnomtablatemp, CSQL1, cnn_form
    '-------
    'Call Nuevo
    txtfecha.Value = Format(Now, "dd/mm/yyyy")
    optmoneda(0).Value = True
    
End Sub

Private Sub Conf_Grid()
    
    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With

    With dxDBGrid2.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
    Call AdicionaItem1

End Sub

Private Sub AdicionaItem1()

    DELETEREC_N cnomtabla2, cnn_form
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    'adiciono la primera fila
    
    dxDBGrid1.Option = egoSmartRefresh
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls
        
    WSQL = "SELECT DISTINCTROW TBVENTA_DET.F5CODPRO, TBVENTA_DET.F5NOMPRO,TBVENTA_DET.F5GRUPO, TBVENTA_DET.F7CODMED, TBVENTA_DET.F3CANPRO, TBVENTA_DET.F3VALVTA, TBVENTA_CAB.F4ESTNUL, TBVENTA_CAB.F4TIPMON, TBVENTA_CAB.F4TIPCAM " _
           & " FROM TBVENTA_CAB INNER JOIN TBVENTA_DET ON (TBVENTA_CAB.F4NUMDOC = TBVENTA_DET.F4NUMDOC) AND (TBVENTA_CAB.F4SERDOC = TBVENTA_DET.F4SERDOC) " _
           & " Where ((TBVENTA_CAB.F2CODALM = '" & gcodalm & "') AND (TBVENTA_CAB.F4ESTNUL <> 'S') AND (TBVENTA_CAB.F4FECEMI=CVDate('" & Format(txtfecha.Value, "DD/MM/YYYY") & "'))) ORDER BY TBVENTA_DET.F5CODPRO;"
    
    If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
    rsdocumentos.Open WSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsdocumentos.EOF Then
        With dxDBGrid1.Dataset
            i = 1: rsdocumentos.MoveFirst
            Do While Not rsdocumentos.EOF
                .Append
                .FieldValues("ITEM") = i
                .FieldValues("CODIGO") = rsdocumentos.Fields("F5CODPRO")
                .FieldValues("DESCRIPCION") = rsdocumentos.Fields("F5NOMPRO")
                .FieldValues("UNIDAD") = rsdocumentos.Fields("F7CODMED")
                .FieldValues("CANTIDAD") = Format(rsdocumentos.Fields("F3CANPRO"), "###,##0.00")
                If rsdocumentos.Fields("F4TIPMON") = "S" Then
                    .FieldValues("VALVTAS") = Format(Val(rsdocumentos.Fields("F3VALVTA") & ""))
                    .FieldValues("VALVTAD") = Format(Val(.FieldValues("VALVTAS") & "") / Val(rsdocumentos.Fields("F4TIPCAM") & ""), "###,##0.00")
                Else
                    .FieldValues("VALVTAD") = Format(Val(rsdocumentos.Fields("F3VALVTA") & ""))
                    .FieldValues("VALVTAS") = Format(Val(.FieldValues("VALVTAD") & "") * Val(rsdocumentos.Fields("F4TIPCAM") & ""), "###,##0.00")
                End If
                .FieldValues("NUMVALE") = ""
                .FieldValues("GRUPO") = rsdocumentos.Fields("F5GRUPO")
                rsdocumentos.MoveNext
                If Not rsdocumentos.EOF Then
                    Do While .FieldValues("CODIGO") = rsdocumentos.Fields("F5CODPRO")
                        .FieldValues("CANTIDAD") = .FieldValues("CANTIDAD") + Format(rsdocumentos.Fields("F3CANPRO"), "###,##0.00")
                        If rsdocumentos.Fields("F4TIPMON") = "S" Then
                            .FieldValues("VALVTAS") = .FieldValues("VALVTAS") + Format(Val(rsdocumentos.Fields("F3VALVTA") & ""))
                            .FieldValues("VALVTAD") = Format(Val(.FieldValues("VALVTAS") & "") / Val(rsdocumentos.Fields("F4TIPCAM") & ""), "###,##0.00")
                        Else
                            .FieldValues("VALVTAD") = Val(.FieldValues("VALVTAD") & "") + Format(Val(rsdocumentos.Fields("F3VALVTA") & ""))
                            .FieldValues("VALVTAS") = Format(Val(.FieldValues("VALVTAD") & "") * Val(rsdocumentos.Fields("F4TIPCAM") & ""), "###,##0.00")
                        End If
                        rsdocumentos.MoveNext
                        If rsdocumentos.EOF Then Exit Do
                    Loop
                End If
                i = i + 1
            Loop
            .Post
        End With
    End If
    rsdocumentos.Close
    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.Dataset.Refresh
    dxDBGrid1.OptionEnabled = True


End Sub

Private Sub LEE_TEMPORAL()
Dim nregistro   As Integer
Dim MEDIDA      As Double

    sw_creabd = True
    sw_leetemp = False
    wFORMULA = 0
    nregistro = 1
    Do While nregistro <= dxDBGrid1.Dataset.RecordCount
        dxDBGrid1.Dataset.RecNo = nregistro
        If RS.State = adStateOpen Then RS.Close
        RS.Open "SELECT * FROM IF4FORMULA WHERE F4CODPRO='" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value & "" & "'"
        If Not RS.EOF Then
            MEDIDA = RS.Fields("F4FBASE")
        Else
            MEDIDA = 0
        End If
        RS.Close
        PROCESA_PRODUCTOS dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value & "", dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value & "", Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value & ""), MEDIDA
        If wFORMULA = 1 Then
            SQL = "INSERT INTO " & wnomtablatemp & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value & "','" & gcodalm & "','" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value & "', " & Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value) & ",'" & PCodOri & "')"
            cnn_form.Execute SQL
            wFORMULA = 0
        End If
        nregistro = nregistro + 1
        sw_leetemp = True
    Loop
    
End Sub

Private Sub PROCESA_PRODUCTOS(pgrupo As String, pproducto As String, pcantidad As Double, pmedida As Double)
Dim CSQL1       As String
Dim Csql2       As String

    If Len(Trim(pgrupo)) = 0 Then
        s = 0
    End If

    'wbasetemp = wusuario & "TEMP" & Format(Time, "hh_mm_ss") & ".MDB"
    Rem EMB CSQL1 = "SELECT * FROM IF4FORMULA WHERE F4GRUPO ='" & pgrupo & "' AND F4CODPRO ='" & pproducto & "'"
    Rem EMB Csql2 = "SELECT * FROM IF3FORMULA WHERE F3GRUPO ='" & pgrupo & "' AND F3CODPRO ='" & pproducto & "' ORDER BY F3GRUPOINS,F3CODPROINS"
    CSQL1 = "SELECT * FROM IF4FORMULA WHERE F4CODPRO ='" & pproducto & "'"
    Csql2 = "SELECT * FROM IF3FORMULA WHERE F3CODPRO ='" & pproducto & "' ORDER BY F3CODPROINS"
    
    ACUMULA_PRODUCTOS wrutatemp, cnombase, CSQL1, Csql2, cnn_dbbancos, wnomtablatemp, cnn_form, pcantidad, "", gcodalm, pproducto, pmedida, ""
    
End Sub

Private Sub AdicionaItem2()
'Dim csql    As String
'Dim rstempo As New ADODB.Recordset

    DELETEREC_N cnomtabla2, cnn_form
    dxDBGrid2.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid2.Dataset.Active = True
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open
    
    dxDBGrid2.Option = egoSmartRefresh
    dxDBGrid2.OptionEnabled = False
    dxDBGrid2.Dataset.DisableControls
    
    'csql = "SELECT VENTASMATERIA.CODIGO, Sum(VENTASMATERIA.CANTIDAD) AS SumaDeCANTIDAD" & _
    '       " From VENTASMATERIA GROUP BY VENTASMATERIA.CODIGO;"
    
    'If rstempo.State = adStateOpen Then rstempo.Close
    'rstempo.Open csql, cnn_form, adOpenStatic, adLockReadOnly
    'If Not rstempo.EOF Then
        'rstempo.MoveFirst
        'Do While Not rstempo.EOF
            With dxDBGrid2.Dataset
                i = 1: RS.MoveFirst
                Do While Not RS.EOF
                    .Append
                    .FieldValues("ITEM") = i
                    .FieldValues("CODIGO") = RS.Fields("CODPROD")
                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    rsif5pla.Open "SELECT * FROM IF5PLA WHERE F5CODPRO='" & RS.Fields("CODPROD") & "" & "' ", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsif5pla.EOF Then
                        .FieldValues("DESCRIPCION") = rsif5pla.Fields("F5NOMPRO")
                        .FieldValues("UNIDAD") = rsif5pla.Fields("F7CODMED")
                    End If
                    rsif5pla.Close
                    .FieldValues("CANTIDAD") = Format(RS.Fields("CANTIDAD"), "###,##0.00")
                    .FieldValues("NUMVALE") = ""
                    '.FieldValues("GRUPO") = RS.Fields("GRUPO")
                    .FieldValues("ALMACEN") = RS.Fields("ALMACEN")
                    RS.MoveNext
                    i = i + 1
                Loop
                .Post
            End With
        'Loop
    'End If
    'rstempo.Close
    'Set rstempo = Nothing
    
    dxDBGrid2.Dataset.EnableControls
        
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open
    dxDBGrid2.Dataset.Refresh
    dxDBGrid2.OptionEnabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    dxDBGrid2.Dataset.Close
    cnn_form.Close
    
    
    ELIMINA_BD_N wrutatemp, cnombase

End Sub



Private Sub optmoneda_Click(Index As Integer)
    Select Case Index
        Case 0
                dxDBGrid1.Columns(8).Visible = True
                dxDBGrid1.Columns(6).Visible = False
        Case 1
                dxDBGrid1.Columns(8).Visible = False
                dxDBGrid1.Columns(6).Visible = True
    End Select
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Ventas"
            Me.MousePointer = 11
            Call Conf_Grid
            Call LEE_TEMPORAL
            If sw_leetemp = True Then
                Call Llena_MatPrima
            Else
                MsgBox "No se realizaron Ventas para la " & txtnomtienda.Text & " para esta Fecha", vbInformation + vbDefaultButton1, "Ventas"
            End If
            Me.MousePointer = 1
        Case "ID_Generar"
            If MsgBox("Desea Generar el Vale de Salida de la Materia Prima", vbYesNo + vbQuestion, "Ventas") = vbYes Then
                Me.MousePointer = 11
                Call GENERA_VALE_SALIDA
                Me.MousePointer = 1
                'MsgBox "El Vale de Salida fue Generado con �xito", vbInformation + vbDefaultButton1, "Ventas"
                SSActiveToolBars1.Tools.ITEM("ID_GENERAR").Enabled = True
                If MsgBox("Desea Imprimir Vale", vbYesNo + vbInformation, "") = vbYes Then
                   'ImprimeVale CCNUMVALE, gcodalm
                End If
                
            End If
        Case "ID_Salir"
                Unload Me
    End Select

End Sub


Private Sub txtcodtienda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtcodtienda.Text = Format(txtcodtienda.Text, "00")
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open "select * from tbtablas where tabtipo='12' and tabcodigo='" & UCase(Trim(txtcodtienda.Text)) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        txtnomtienda.Text = rs.Fields("tabdescripcion")
'        gcodalm = Format(rs.Fields("tabcantidad"), "00")
'    End If
'    rs.Close OJO!!!!
    If RS.State = adStateOpen Then RS.Close
    RS.Open "select * from EF2ALMACENES where F2CODALM='" & UCase(Trim(txtcodtienda.Text)) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        txtnomtienda.Text = RS.Fields("F2NOMALM")
        gcodalm = Format(RS.Fields("F2CODALM"), "00")
    End If
    RS.Close
    'SendKeys "TAB"
 End If
End Sub

Private Sub txtcodtienda_LostFocus()
    If RS.State = adStateOpen Then RS.Close
    RS.Open "select * from tbtablas where tabtipo='12' and tabcodigo='" & UCase(Trim(txtcodtienda.Text)) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        txtnomtienda.Text = RS.Fields("tabdescripcion")
        gcodalm = Format(RS.Fields("tabcantidad"), "00")
    End If
    RS.Close
End Sub

Public Sub Llena_MatPrima()
Dim csql    As String

    Rem NSE RS.Open "SELECT * FROM " & wnomtablatemp & " ORDER BY CODPROD,GRUPO, ALMACEN ", cnn_form, adOpenDynamic, adLockOptimistic
    If RS.State = adStateOpen Then RS.Close
    csql = "SELECT Sum(CANTIDAD) AS CANTIDAD, CODPROD, ALMACEN" & _
           " From " & wnomtablatemp & _
           " GROUP BY CODPROD, ALMACEN " & _
           " ORDER BY CODPROD, ALMACEN;"
    RS.Open csql, cnn_form, adOpenStatic, adLockReadOnly
    If Not RS.EOF Then
        Call AdicionaItem2
    End If
    RS.Close
    
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "TAB"
    End If
End Sub

Private Sub TxtFecha_LostFocus()

    If IsDate(txtfecha.Value) Then
        If rscambios.State = adStateOpen Then rscambios.Close
        rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & txtfecha.Value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rscambios.EOF Then
            txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
        Else
            txttc.Text = Format(0, "0.000")
        End If
        rscambios.Close
    Else
        MsgBox "Fecha incorrecta. Verifique.", vbCritical, "Atenci�n"
        txtfecha.SetFocus
    End If

End Sub

Private Sub txttc_GotFocus()

    txttc.SelStart = 0: txttc.SelLength = Len(txttc.Text)

End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txttc.Text = Format(txttc.Text, "0.000")
        SendKeys "{tab}"
    End If
    
End Sub
