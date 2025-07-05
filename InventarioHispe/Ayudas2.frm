VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Ayudas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayudas"
   ClientHeight    =   6090
   ClientLeft      =   1065
   ClientTop       =   1365
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtnumero 
      Height          =   240
      Left            =   990
      TabIndex        =   5
      Top             =   5580
      Width           =   2715
   End
   Begin VB.TextBox txtdocumento 
      Height          =   240
      Left            =   3600
      TabIndex        =   3
      Top             =   5220
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtserie 
      Height          =   240
      Left            =   720
      TabIndex        =   2
      Top             =   5220
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4605
      Left            =   135
      TabIndex        =   6
      Top             =   45
      Width           =   8250
      _Version        =   65536
      _ExtentX        =   14552
      _ExtentY        =   8123
      _StockProps     =   15
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4380
         Left            =   90
         OleObjectBlob   =   "Ayudas2.frx":0000
         TabIndex        =   7
         Top             =   45
         Width           =   8025
      End
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   4770
      Width           =   5820
   End
   Begin aBoxCtl.aBox abodesde 
      Height          =   255
      Left            =   6300
      TabIndex        =   4
      Top             =   5220
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   450
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
      Text            =   "22/04/2002"
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
      ButtonPicture   =   "Ayudas2.frx":1940
      ButtonWidth     =   17
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
   Begin VB.Label lblnumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   5580
      Width           =   600
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Emision:"
      Height          =   195
      Left            =   5130
      TabIndex        =   10
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbldocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Documento:"
      Height          =   195
      Left            =   2295
      TabIndex        =   9
      Top             =   5265
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblserie 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serie:"
      Height          =   195
      Left            =   225
      TabIndex        =   8
      Top             =   5220
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   4815
      Width           =   885
   End
End
Attribute VB_Name = "Ayudas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim cadena As String, cserie As String, cdocumento As String, cfecha As String, cnumero As String
Dim SQL As String

Private Sub abodesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If wtipoayuda = "G" Then
            cfecha = Format("" & abodesde.Value, "dd/mm/yyyy")
        
               If IsDate(cfecha) Then
                
                     If Len(gcoduse) > 0 Then
                         SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                         "where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F2CODCLI= '" & Trim(gcoduse) & "' and F4FECEMI >= CVDATE('" & cfecha & "') and F4FECEMI<=CVDATE('" & cfecha & "')"
        
                     Else
        
                         SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                         "Where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F4FECEMI >=CVDATE('" & cfecha & "') and F4FECEMI<=CVDATE('" & cfecha & "')"
        
                    End If

               Else
                    SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                         "where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F2CODCLI= '" & Trim(gcoduse) & "'"
                    txtserie.SetFocus
               End If
         Else
              If wtipoayuda = "FAC" And KeyAscii = 13 Then
                   cfecha = Format("" & abodesde.Value, "dd/mm/yyyy")
                   If IsDate(cfecha) Then
                        SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                        "Where F4ESTNUL = 'N' and F4TIPODOCU = '01' and F4FECEMI >=CVDATE('" & cfecha & "') and F4FECEMI<=CVDATE('" & cfecha & "')"
                   Else
                        txtserie.SetFocus
                   End If
              Else
                    If wtipoayuda = "B" Then
                        cfecha = Format("" & abodesde.Value, "dd/mm/yyyy")
                           If IsDate(cfecha) Then
                                SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                                "Where F4ESTNUL = 'N' and F4TIPODOCU = '03' and F4FECEMI >=CVDATE('" & cfecha & "') and F4FECEMI<=CVDATE('" & cfecha & "')"
                           Else
                               txtserie.SetFocus
                           End If
                    Else
                        If wtipoayuda = "N" Then
                            cfecha = Format("" & abodesde.Value, "dd/mm/yyyy")
                            If IsDate(cfecha) Then
                                SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                                "Where F4ESTNUL = 'N' and F4TIPODOCU = '07' and F4FECEMI >=CVDATE('" & cfecha & "') and F4FECEMI<=CVDATE('" & cfecha & "')"
                            Else
                                txtserie.SetFocus
                            End If
                        End If
                   End If
             End If
          End If
        End If
        
        PROCEDIMIENTO
        CABECERA
End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    wcodigos = "" & dxDBGrid1.Columns(0).Value
    wdescripcion = "" & dxDBGrid1.Columns(1).Value
    Me.Hide

End Sub

Private Sub Form_Load()

    abodesde.Value = Format(Now, "dd/mm/yyyy")
    Me.Left = 3000
    Me.Top = 980
        
    'INGRESO DEL TIPO DE AYUDA EN VARIABLE wtipoayuda y vale
    
    wtipoayuda = UCase$(wtipoayuda)
    wvale = "i"
    wvale = UCase$(wvale)
       
    Select Case wtipoayuda
            Case "A"
                    'Ayuda de Almacenes
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "F"
                    'Ayuda de Forma de Pago
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "M"
                    'Ayuda de Marcas
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "P"
                    'Ayuda de Proveedores
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "T"
                    'Ayuda de Transporte
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "V"
                    'Ayuda de Vendedores
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "Z"
                    'Ayuda de Zonas
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "D"
                    'Ayuda de Medidas
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "G"
                    'Ayuda de Guias
                    CONECTAR
                    LLENAR
                    CABECERA
            Case "C"
                    'Ayuda de Cotizaciones
                    CONECTAR
                    LLENAR
                    CABECERA
                    
            Case "PE"
                    'Ayuda de Pedidos
                    CONECTAR
                    LLENAR
                    CABECERA
           
            Case "CE"
                    'Ayuda de Centros de Costo
                    CONECTAR
                    LLENAR
                    CABECERA
                       
           Case "FAC"
                    'Ayuda de Facturas
                    CONECTAR
                    LLENAR
                    CABECERA
           Case "B"
                    'Ayuda de Boletas
                    CONECTAR
                    LLENAR
                    CABECERA
           Case "N"
                    'Ayuda de Nota de Credito
                    CONECTAR
                    LLENAR
                    CABECERA
           Case "VAL"
                   'Ayuda de Vales
                    'CONECTAR
                    BASE_TEMPORAL_V
                    TABLA_TEMPORAL_V
                    CONECTAR
                    LLENAR
                    
                    With dxDBGrid1
                        .DefaultFields = True
                        '.Dataset.ADODataset.ConnectionString = temp
                        SQL = "Select * from " & DBTable1 & ""
                        .Dataset.Active = False
                        .Dataset.ADODataset.CommandText = SQL
                         .Dataset.Active = True
                        .KeyField = "ITEM"
                    End With
    
                    CABECERA
                    
    End Select
    
End Sub

Public Sub CONECTAR()
    
    If wtipoayuda = "A" Or wtipoayuda = "F" Or wtipoayuda = "P" Or wtipoayuda = "T" Or wtipoayuda = "V" Or wtipoayuda = "Z" Or wtipoayuda = "D" Then
    
        With dxDBGrid1
            .DefaultFields = True
            .Dataset.ADODataset.ConnectionString = cn
        End With
    Else
        If wtipoayuda = "VAL" Then
            With dxDBGrid1
                .DefaultFields = True
                .Dataset.ADODataset.ConnectionString = temp
            End With
        Else
            If wtipoayuda = "G" Or wtipoayuda = "C" Or wtipoayuda = "PE" Or wtipoayuda = "FAC" Or wtipoayuda = "B" Or wtipoayuda = "N" Then
                With dxDBGrid1
                    .DefaultFields = True
                    .Dataset.ADODataset.ConnectionString = cn2
                End With
            Else
                If wtipoayuda = "CE" Then
                    With dxDBGrid1
                        .DefaultFields = True
                        .Dataset.ADODataset.ConnectionString = cn3
                    End With
                Else
                    If wtipoayuda = "M" Then
                         With dxDBGrid1
                            .DefaultFields = True
                            .Dataset.ADODataset.ConnectionString = cn1
                         End With
                    
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Public Sub CABECERA()

    Select Case wtipoayuda
    Case "A"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Almacen": .Columns(1).Width = 120: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Direccion": .Columns(2).Width = 120: .Columns(2).DisableEditor = True
           End With
           
    Case "F"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 250: .Columns(1).DisableEditor = True
           End With
    
    Case "M"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
           End With
    
    Case "P"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Proveedor": .Columns(1).Width = 250: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Ruc": .Columns(2).Width = 80: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Direccion": .Columns(3).Width = 250: .Columns(3).DisableEditor = True
           End With
    Case "T"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 205: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Direccion": .Columns(2).Width = 230: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Ruc": .Columns(3).Width = 65: .Columns(3).DisableEditor = True
                .Columns(4).Caption = "Numero": .Columns(4).Width = 65: .Columns(4).DisableEditor = True
           End With
    Case "V"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Vendedor": .Columns(1).Width = 110: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Direccion": .Columns(2).Width = 130: .Columns(2).DisableEditor = True
           End With
    Case "Z"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 150: .Columns(1).DisableEditor = True
           End With
    Case "D"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "Codigo": .Columns(0).Width = 70: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 150: .Columns(1).DisableEditor = True
           End With
    Case "G"
           MEDIDAS
           With dxDBGrid1
                .Columns(0).Caption = "No Serie": .Columns(0).Width = 100: .Columns(0).DisableEditor = True
                .Columns(1).Caption = "No Documento": .Columns(1).Width = 100: .Columns(1).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                
                .Columns(2).Caption = "Fecha Emision": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
            End With
    
    Case "C"
            MEDIDAS
            With dxDBGrid1
                .Columns(0).Caption = "No Cotizacion": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Fecha Emision": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Codigo": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
             End With
             
    Case "PE"
             MEDIDAS
             With dxDBGrid1
                .Columns(0).Caption = "No Pedido": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Fecha Emision": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Codigo": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
                .Columns(4).Caption = "Total Factura": .Columns(4).Width = 100: .Columns(4).DisableEditor = True: .Columns(4).DecimalPlaces = 2
                
            End With
    Case "CE"
             MEDIDAS
             With dxDBGrid1
                .Columns(0).Caption = "No Costo": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "Descripcion": .Columns(1).Width = 200: .Columns(1).DisableEditor = True
             End With
             
    Case "FAC"
             MEDIDAS
             With dxDBGrid1
                .Columns(0).Caption = "No Serie": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "No Documento": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Fecha Emision": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
             End With
     Case "B"
            MEDIDAS
            With dxDBGrid1
                .Columns(0).Caption = "No Serie": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "No Documento": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Fecha Emision": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
            End With
            
    Case "N"
            MEDIDAS
            With dxDBGrid1
                .Columns(0).Caption = "No Serie": .Columns(0).Width = 100: .Columns(0).DisableEditor = True: .Columns(0).Color = &HC0FFFF
                .Columns(1).Caption = "No Documento": .Columns(1).Width = 100: .Columns(1).DisableEditor = True
                .Columns(2).Caption = "Fecha Emision": .Columns(2).Width = 100: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Cliente": .Columns(3).Width = 200: .Columns(3).DisableEditor = True
            End With
    
    Case "VAL"
            MEDIDAS
            With dxDBGrid1
                .Columns(0).Caption = "ITEM": .Columns(0).Width = 10
                .Columns(1).Caption = "Vale": .Columns(1).Width = 80: .Columns(1).DisableEditor = True: .Columns(1).Color = &HC0FFFF
                .Columns(2).Caption = "Fecha": .Columns(2).Width = 80: .Columns(2).DisableEditor = True
                .Columns(3).Caption = "Nombre": .Columns(3).Width = 100: .Columns(3).DisableEditor = True
                .Columns(0).Visible = False
    
            End With
    
    End Select
    
End Sub

Public Sub LLENAR()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

    Select Case wtipoayuda
    Case "A"
        SQL = "SELECT F2CODALM,F2NOMALM,F2DIRALM FROM EF2ALMACENES"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODALM"
    Case "F"
        SQL = "SELECT F2FORPAG,F2DESPAG FROM EF2FORPAG"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2FORPAG"
    Case "M"
        SQL = "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODMAR"
    Case "P"
        SQL = "SELECT F2CODPROV,F2NOMPROV,F2NEWRUC,F2DIRPROV FROM EF2PROVEEDORES"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODPROV"
    Case "T"
        SQL = "SELECT F2CODTRA,F2NOMTRA,F2DIRTRA,F2RUCTRA,F2PLATRA FROM EF2TRANSPORT"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODTRA"
    Case "V"
        SQL = "SELECT F2CODVEN,F2NOMVEN,F2DIRVEN FROM EF2VENDEDORES"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODVEN"
    Case "Z"
        SQL = "SELECT F2CODZON,F2DESZON FROM EF2ZONAS"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODZON"
    Case "D"
        SQL = "SELECT F7CODMED,F7NOMMED FROM EF7MEDIDAS"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F7CODMED"
    
    Case "G"
        If Len(gcoduse) > 0 Then
            SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
            "where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F2CODCLI= '" & Trim(gcoduse) & "'"
            
        Else
            SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
            "Where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09'"
       End If
       dxDBGrid1.Dataset.ADODataset.CommandText = SQL
       dxDBGrid1.Dataset.Active = True
       dxDBGrid1.KeyField = "F4NUMDOC"
       
       
    Case "C"
         
        SQL = "SELECT F4NUMCOT,F4FECEMI,F2CODCLI,F2NOMCLI FROM TBCOTIZA_CAB WHERE F4ESTNUL='N'"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F4NUMCOT"
        
    Case "PE"
        SQL = "SELECT F4NUMPED,F4FECEMI,F2CODCLI,F2NOMCLI,F4TOTFAC FROM TBPEDIDO_CAB WHERE F4ESTNUL='N'"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F4NUMPED"
        
    Case "CE"
        SQL = "SELECT F3COSTO,F3DESCRIP FROM CENTROS"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F3COSTO"
    
    Case "FAC"
        SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
            "WHERE F4ESTNUL='N' AND F4TIPODOCU='01'"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F4NUMDOC"
    
    Case "B"
        SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
        "WHERE F4ESTNUL='N' AND F4TIPODOCU='03'"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F4NUMDOC"
    
    Case "N"
       SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
        "WHERE F4ESTNUL='N' AND F4TIPODOCU='07'"
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F4NUMDOC"
        
    Case "VAL"
    
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N DBTable1, temp
    X = 1
       SQL = "SELECT F4NUMVAL,F4FECVAL,F1CODORI FROM IF4VALES WHERE LEFT(F4NUMVAL,1)='" & wvale & "'"
       If rs.State = adStateOpen Then rs.Close
       rs.Open SQL, dbbancowin, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
       Do While Not rs.EOF
           
           numerovale = Trim(rs.Fields("F4NUMVAL"))
           fecha = Trim(rs.Fields("F4FECVAL"))
           codigo = Trim(rs.Fields("F1CODORI"))
            
        SQL = "SELECT * FROM SF1ORIGENES WHERE F1CODORI='" & codigo & "'"
            If rs1.State = adStateOpen Then rs1.Close
            rs1.Open SQL, dbbancowin, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
               
                codigo1 = Trim(rs1.Fields("F1CODORI"))
                nombre = Trim(rs1.Fields("F1NOMORI"))
               
            End If
            
                csql = "INSERT INTO " & DBTable1 & " (ITEM,VALE,FECHA,NOMBRE)" & _
                       "VALUES (" & X & " ,'" & numerovale & "','" & fecha & "','" & nombre & "')"
                                        
                                    
                temp.Execute (csql)
                X = X + 1
    
       rs.MoveNext
       Loop
       End If
    
    End Select
    
End Sub

Private Sub TxtDescripcion_Change()

    Select Case wtipoayuda
    
    Case "A"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
    
        If Len(cadena) > 0 Then
            SQL = "SELECT F2CODALM,F2NOMALM,F2DIRALM FROM EF2ALMACENES WHERE F2NOMALM >='" & cadena & "' and F2NOMALM <='" & cadena & "z" & "' order by F2NOMALM"
            PROCEDIMIENTO
            CABECERA
    
        Else
            SQL = "SELECT F2CODALM,F2NOMALM,F2DIRALM FROM EF2ALMACENES order by F2NOMALM"
            PROCEDIMIENTO
            CABECERA
        End If
    
    Case "F"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F2FORPAG,F2DESPAG FROM EF2FORPAG WHERE F2DESPAG >='" & cadena & "' and F2DESPAG <='" & cadena & "z" & "' ORDER BY F2DESPAG"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2FORPAG,F2DESPAG FROM EF2FORPAG ORDER BY F2DESPAG "
            PROCEDIMIENTO
            CABECERA
        End If
       
    Case "M"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS WHERE F2DESMAR >='" & cadena & "' and F2DESMAR <='" & cadena & "z" & "' ORDER BY F2DESMAR"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2CODMAR,F2DESMAR FROM EF2MARCAS ORDER BY F2DESMAR  "
            PROCEDIMIENTO
            CABECERA
        End If
    
    Case "P"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F2CODPROV,F2NOMPROV,F2NEWRUC,F2DIRPROV FROM EF2PROVEEDORES WHERE F2NOMPROV >='" & cadena & "' and F2NOMPROV <='" & cadena & "z" & "' ORDER BY F2NOMPROV"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2CODPROV,F2NOMPROV,F2NEWRUC,F2DIRPROV FROM EF2PROVEEDORES ORDER BY F2NOMPROV"
            PROCEDIMIENTO
            CABECERA
        End If
    
    Case "T"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F2CODTRA,F2NOMTRA,F2DIRTRA,F2RUCTRA,F2PLATRA FROM EF2TRANSPORT WHERE F2NOMTRA >='" & cadena & "' and F2NOMTRA <='" & cadena & "z" & "' ORDER BY F2NOMTRA"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2CODTRA,F2NOMTRA,F2DIRTRA,F2RUCTRA,F2PLATRA FROM EF2TRANSPORT ORDER BY F2NOMTRA"
            PROCEDIMIENTO
            CABECERA
        End If
        
    Case "V"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            
            SQL = "SELECT F2CODVEN,F2NOMVEN,F2DIRVEN FROM EF2VENDEDORES WHERE F2NOMVEN >='" & cadena & "' and F2NOMVEN <='" & cadena & "z" & "' ORDER BY F2NOMVEN"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2CODVEN,F2NOMVEN,F2DIRVEN FROM EF2VENDEDORES ORDER BY F2NOMVEN"
            PROCEDIMIENTO
            CABECERA
        End If
    
    Case "Z"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            
            SQL = "SELECT F2CODZON,F2DESZON FROM EF2ZONAS WHERE F2DESZON >='" & cadena & "' and F2DESZON <='" & cadena & "z" & "' ORDER BY F2DESZON"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F2CODZON,F2DESZON FROM EF2ZONAS ORDER BY F2DESZON"
            PROCEDIMIENTO
            CABECERA
        End If
     
    Case "D"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F7CODMED,F7NOMMED FROM EF7MEDIDAS WHERE F7NOMMED >='" & cadena & "' and F7NOMMED <='" & cadena & "z" & "' ORDER BY F7NOMMED"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F7CODMED,F7NOMMED FROM EF7MEDIDAS ORDER BY F7NOMMED"
            PROCEDIMIENTO
            CABECERA
        End If
       
    Case "CE"
        cadena = Trim(TxtDescripcion.Text)
        cadena = UCase$(cadena)
        
        If Len(cadena) > 0 Then
            SQL = "SELECT F3COSTO,F3DESCRIP FROM CENTROS WHERE F3DESCRIP >='" & cadena & "' and F3DESCRIP <='" & cadena & "z" & "' order by F3COSTO"
            PROCEDIMIENTO
            CABECERA
        Else
            SQL = "SELECT F3COSTO,F3DESCRIP FROM CENTROS ORDER BY F3COSTO"
            PROCEDIMIENTO
            CABECERA
        End If
       
    End Select

End Sub

Public Sub PROCEDIMIENTO()
        
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = SQL
    dxDBGrid1.Dataset.Open
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Refresh
        
End Sub

Public Sub MEDIDAS()

    Select Case wtipoayuda
    Case "A"
           Ayudas.Width = 5600
           Ayudas.Height = 4500
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3600
           lblDescripcion.Left = 300
           TxtDescripcion.Top = 3600
           TxtDescripcion.Left = 1300
           Ayudas.Caption = "Ayuda de Almacenes"
           TxtDescripcion.Width = 3300
    
    Case "F"
    
           Ayudas.Width = 5600
           Ayudas.Height = 4500
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3600
           lblDescripcion.Left = 300
           TxtDescripcion.Top = 3600
           TxtDescripcion.Left = 1300
           Ayudas.Caption = "Ayuda de Forma de Pago"
           TxtDescripcion.Width = 4000
    
    Case "M"
    
           Ayudas.Width = 5600
           Ayudas.Height = 4500
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3600
           lblDescripcion.Left = 300
           TxtDescripcion.Top = 3600
           TxtDescripcion.Left = 1300
           Ayudas.Caption = "Ayuda de Marcas"
           TxtDescripcion.Width = 3000
           
    Case "P"
    
           Ayudas.Width = 8600
           Me.Left = 1500
           Ayudas.Caption = "Ayuda de Proveedores"
           Me.Height = 5600
           TxtDescripcion.Left = 1300
           
    Case "T"
           
           Ayudas.Width = 8600
           Me.Left = 1500
           Ayudas.Caption = "Ayuda de Transporte"
           TxtDescripcion.Width = 5100
           Me.Height = 5600
           TxtDescripcion.Left = 1300
            
    Case "V"
           Me.Left = 1900
           Ayudas.Width = 5600
           Ayudas.Height = 4400
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3500
           lblDescripcion.Left = 400
           TxtDescripcion.Top = 3500
           TxtDescripcion.Left = 1380
           Ayudas.Caption = "Ayuda de Vendedores"
           TxtDescripcion.Width = 3000
           
    Case "Z"
    
           Ayudas.Width = 4400
           Ayudas.Height = 4200
           dxDBGrid1.Width = 3800
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 4100
           SSPanel1.Height = 3300
           lblDescripcion.Top = 3400
           lblDescripcion.Left = 360
           TxtDescripcion.Top = 3400
           TxtDescripcion.Left = 1400
           Ayudas.Caption = "Ayuda de Zonas"
           TxtDescripcion.Width = 2000
            
    Case "D"
    
           Ayudas.Width = 4400
           Ayudas.Height = 4200
           dxDBGrid1.Width = 3800
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 4100
           SSPanel1.Height = 3300
           lblDescripcion.Top = 3400
           lblDescripcion.Left = 300
           TxtDescripcion.Top = 3400
           TxtDescripcion.Left = 1400
           Ayudas.Caption = "Ayuda de Medidas"
           TxtDescripcion.Width = 2000
    
    Case "G"
            
           Ayudas.Width = 7800
           Ayudas.Height = 4100
           Ayudas.Top = 1200
           Ayudas.Left = 2500
           dxDBGrid1.Width = 7200
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 7400
           SSPanel1.Height = 3100
           TxtDescripcion.Top = 3600
           TxtDescripcion.Left = 1300
           TxtDescripcion.Width = 2000
           Ayudas.Caption = "Ayuda de Guias"
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           txtserie.Visible = True
           txtdocumento.Visible = True
           abodesde.Visible = True
           lblserie.Visible = True
           lbldocumento.Visible = True
           lblfecha.Visible = True
           lblserie.Top = 3300
           lbldocumento.Top = 3300
           lblfecha.Top = 3300
           txtserie.Top = 3300
           txtdocumento.Top = 3300
           abodesde.Top = 3300
    
    
    Case "C"
          
           Me.Left = 1900
           Ayudas.Width = 8350
           Ayudas.Height = 4100
           dxDBGrid1.Width = 7800
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 8000
           SSPanel1.Height = 3150
           Ayudas.Caption = "Ayuda de Cotizaciones"
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           lblnumero.Top = 3300
           lblnumero.Left = 450
           txtnumero.Top = 3300
           txtnumero.Left = 1150
           txtnumero.Width = 1500
           
    
    Case "PE"
    
           Me.Left = 1900
           Ayudas.Width = 8350
           Ayudas.Height = 4300
           dxDBGrid1.Width = 7800
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 8000
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3500
           lblDescripcion.Left = 400
           TxtDescripcion.Top = 3500
           TxtDescripcion.Left = 1500
           Ayudas.Caption = "Ayuda de Pedidos"
           TxtDescripcion.Width = 3000
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           lblnumero.Top = 3500
           lblnumero.Left = 450
           txtnumero.Top = 3500
           txtnumero.Left = 1150
           txtnumero.Width = 1500
    
    Case "CE"
           Ayudas.Width = 5600
           Ayudas.Height = 4300
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           lblDescripcion.Top = 3500
           lblDescripcion.Left = 300
           TxtDescripcion.Top = 3500
           TxtDescripcion.Left = 1300
           Ayudas.Caption = "Ayuda de Centros de Costo"
           TxtDescripcion.Width = 3300
    
    Case "FAC"
            
           Ayudas.Width = 7800
           Ayudas.Height = 4100
           Ayudas.Top = 1200
           Ayudas.Left = 2500
           dxDBGrid1.Width = 7200
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 7400
           SSPanel1.Height = 3100
           Ayudas.Caption = "Ayuda de Facturas"
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           txtserie.Visible = True
           txtdocumento.Visible = True
           abodesde.Visible = True
           lblserie.Visible = True
           lbldocumento.Visible = True
           lblfecha.Visible = True
           lblserie.Top = 3300
           lbldocumento.Top = 3300
           lblfecha.Top = 3300
           txtserie.Top = 3300
           txtdocumento.Top = 3300
           abodesde.Top = 3300
           lblnumero.Visible = False
           txtnumero.Visible = False
          
    Case "B"
           Ayudas.Width = 7800
           Ayudas.Height = 4100
           Ayudas.Top = 1200
           Ayudas.Left = 2500
           dxDBGrid1.Width = 7200
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 7400
           SSPanel1.Height = 3100
           Ayudas.Caption = "Ayuda de Boletas"
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           txtserie.Visible = True
           txtdocumento.Visible = True
           abodesde.Visible = True
           lblserie.Visible = True
           lbldocumento.Visible = True
           lblfecha.Visible = True
           lblserie.Top = 3300
           lbldocumento.Top = 3300
           lblfecha.Top = 3300
           txtserie.Top = 3300
           txtdocumento.Top = 3300
           abodesde.Top = 3300
           lblnumero.Visible = False
           txtnumero.Visible = False
          
    Case "N"
           Ayudas.Width = 7800
           Ayudas.Height = 4100
           Ayudas.Top = 1200
           Ayudas.Left = 2500
           dxDBGrid1.Width = 7200
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 7400
           SSPanel1.Height = 3100
           Ayudas.Caption = "Ayuda de Nota de Credito"
           lblDescripcion.Visible = False
           TxtDescripcion.Visible = False
           txtserie.Visible = True
           txtdocumento.Visible = True
           abodesde.Visible = True
           lblserie.Visible = True
           lbldocumento.Visible = True
           lblfecha.Visible = True
           lblserie.Top = 3300
           lbldocumento.Top = 3300
           lblfecha.Top = 3300
           txtserie.Top = 3300
           txtdocumento.Top = 3300
           abodesde.Top = 3300
           lblnumero.Visible = False
           txtnumero.Visible = False
    
    
    Case "VAL"
           Ayudas.Width = 5600
           Ayudas.Height = 4500
           dxDBGrid1.Width = 5000
           dxDBGrid1.Height = 3000
           SSPanel1.Width = 5300
           SSPanel1.Height = 3350
           'lblDescripcion.top = 3600
           'lblDescripcion.left = 300
           'TxtDescripcion.top = 3600
           'TxtDescripcion.left = 1300
           Ayudas.Caption = "Ayuda de Vales"
           'TxtDescripcion.Width = 3300
    
    End Select

End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1_OnDblClick
    End If

End Sub

Private Sub txtnumero_Change()

    Select Case wtipoayuda
    
    Case "C"
        cnumero = Trim(txtnumero.Text)
        If Len(cnumero) > 0 Then
            SQL = "SELECT F4NUMCOT,F4FECEMI,F2CODCLI,F2NOMCLI FROM TBCOTIZA_CAB WHERE F4ESTNUL='N' AND F4NUMCOT >='" & cnumero & "' AND F4NUMCOT <= '" & cnumero & "z" & "' order by F4NUMCOT"
        Else
            SQL = "SELECT F4NUMCOT,F4FECEMI,F2CODCLI,F2NOMCLI FROM TBCOTIZA_CAB WHERE F4ESTNUL='N' order by F4NUMCOT"
        End If
        PROCEDIMIENTO
        CABECERA
    
    Case "PE"
        cnumero = Trim(txtnumero.Text)
        If Len(cnumero) > 0 Then
            SQL = "SELECT F4NUMPED,F4FECEMI,F2CODCLI,F2NOMCLI,F4TOTFAC FROM TBPEDIDO_CAB WHERE F4ESTNUL='N' AND F4NUMPED >='" & cnumero & "' AND F4NUMPED <= '" & cnumero & "z" & "' order by F4NUMPED"
        Else
            SQL = "SELECT F4NUMPED,F4FECEMI,F2CODCLI,F2NOMCLI,F4TOTFAC FROM TBPEDIDO_CAB WHERE F4ESTNUL='N' ORDER BY F4NUMPED"
        End If
        PROCEDIMIENTO
        CABECERA
    
    End Select

End Sub

Private Sub txtserie_Change()

    If wtipoayuda = "G" Then
        cserie = Trim(txtserie)

        If Len(cserie) > 0 Then
            If Len(gcoduse) > 0 Then
                 SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                 "where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F2CODCLI= '" & Trim(gcoduse) & "' and f4serdoc >='" & cserie & "' and f4serdoc<='" & cserie & "z" & "'"
             Else
                 SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                 "Where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and f4serdoc >='" & cserie & "' and f4serdoc<='" & cserie & "z" & "'"
            End If
        Else
            txtdocumento.SetFocus
        End If
    Else
        If wtipoayuda = "FAC" Then
            cserie = Trim(txtserie)
                If Len(cserie) > 0 Then
                    SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                        "WHERE F4ESTNUL='N' AND F4TIPODOCU='01' and f4serdoc >='" & cserie & "' and f4serdoc<='" & cserie & "z" & "'"
                Else
                    txtdocumento.SetFocus
                End If
        'End If
        Else
            If wtipoayuda = "B" Then
                cserie = Trim(txtserie)
                    If Len(cserie) > 0 Then
                        SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                        "WHERE F4ESTNUL='N' AND F4TIPODOCU='03' and f4serdoc >='" & cserie & "' and f4serdoc<='" & cserie & "z" & "'"
                    Else
                        txtdocumento.SetFocus
                    End If
    '       End If
    '    End If
        
            Else
                   If wtipoayuda = "N" Then
                       cserie = Trim(txtserie)
                       If Len(cserie) > 0 Then
                           SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                           "WHERE F4ESTNUL='N' AND F4TIPODOCU='07' and f4serdoc >='" & cserie & "' and f4serdoc<='" & cserie & "z" & "'"
                       Else
                           txtdocumento.SetFocus
                       End If
                   End If
            End If
       End If
    End If
    PROCEDIMIENTO
    CABECERA
    
End Sub

Private Sub txtdocumento_Change()

    If wtipoayuda = "G" Then
        cdocumento = Trim(txtdocumento)
    
           If Len(txtdocumento) > 0 Then
                 If Len(gcoduse) > 0 Then
                     SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                     "where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F2CODCLI= '" & Trim(gcoduse) & "' and F4NUMDOC >='" & cdocumento & "' and F4NUMDOC<='" & cdocumento & "z" & "'"
                 Else
                     SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                     "Where F4ESTNUL = 'N' and F4ESTFAC = 'N' and F4TIPODOCU = '09' and F4NUMDOC >='" & cdocumento & "' and F4NUMDOC<='" & cdocumento & "z" & "'"
                End If
           
            Else
                abodesde.SetFocus
            End If
     Else
         If wtipoayuda = "FAC" Then
            cdocumento = Trim(txtdocumento)
                If Len(txtdocumento) > 0 Then
                    SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                    "Where F4ESTNUL = 'N' and F4TIPODOCU = '01' and F4NUMDOC >='" & cdocumento & "' and F4NUMDOC<='" & cdocumento & "z" & "'"
                Else
                    abodesde.SetFocus
                End If
        Else
            If wtipoayuda = "B" Then
                cdocumento = Trim(txtdocumento)
                   If Len(txtdocumento) > 0 Then
                       SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                       "Where F4ESTNUL = 'N' and F4TIPODOCU = '03' and F4NUMDOC >='" & cdocumento & "' and F4NUMDOC<='" & cdocumento & "z" & "'"
                   Else
                       abodesde.SetFocus
                   End If
            Else
                If wtipoayuda = "N" Then
                     cdocumento = Trim(txtdocumento)
                     If Len(txtdocumento) > 0 Then
                        SQL = "SELECT F4SERDOC,F4NUMDOC,F4FECEMI,F2NOMCLI FROM TBVENTA_CAB " & _
                        "Where F4ESTNUL = 'N' and F4TIPODOCU = '07' and F4NUMDOC >='" & cdocumento & "' and F4NUMDOC<='" & cdocumento & "z" & "'"
                     Else
                        abodesde.SetFocus
                     End If
    
                End If
            End If
        End If
    End If
    PROCEDIMIENTO
    CABECERA
    
End Sub

Public Sub BASE_TEMPORAL_V()

    Set temp = New ADODB.Connection
    
    usuario = "Johana"
    basetemporal = usuario & "Centros" & Format(Time, "hh_mm_ss") & ".MDB"
    CREATEDATABASE_N "C:\Bancowin\", CStr(basetemporal)
    CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=C:\Bancowin\" & basetemporal & "; Persist Security Info=False"
    temp.Open CON
    
End Sub

Public Sub TABLA_TEMPORAL_V()
    
    DBTable1 = "Vales"
    SQL = "(ITEM Text(5), VALE Text(20),FECHA Date,NOMBRE Text(100))"
    CREATETABLE_N DBTable1, CStr(SQL), temp

End Sub
