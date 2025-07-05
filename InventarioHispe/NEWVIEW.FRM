VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmView 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Preliminar"
   ClientHeight    =   6420
   ClientLeft      =   1725
   ClientTop       =   1020
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "NEWVIEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   9435
   Begin MSComDlg.CommonDialog CAJA 
      Left            =   3825
      Top             =   4680
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.TextBox txtPag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   396
      Left            =   1296
      TabIndex        =   9
      Text            =   "1"
      Top             =   5952
      Width           =   540
   End
   Begin Threed.SSCheck OPCPRE 
      Height          =   300
      Left            =   6576
      TabIndex        =   8
      Top             =   5952
      Width           =   2796
      _Version        =   65536
      _ExtentX        =   4932
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Imprimir con Valor predeterminado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   396
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   5952
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   698
      _StockProps     =   78
      Caption         =   "<<"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   396
      Index           =   1
      Left            =   768
      TabIndex        =   3
      Top             =   5952
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   698
      _StockProps     =   78
      Caption         =   "<"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   396
      Index           =   2
      Left            =   1824
      TabIndex        =   4
      Top             =   5952
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   698
      _StockProps     =   78
      Caption         =   ">"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   396
      Index           =   3
      Left            =   2355
      TabIndex        =   5
      Top             =   5952
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   698
      _StockProps     =   78
      Caption         =   ">>"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   390
      Index           =   4
      Left            =   4905
      TabIndex        =   6
      Top             =   5955
      Width           =   525
      _Version        =   65536
      _ExtentX        =   917
      _ExtentY        =   698
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "NEWVIEW.frx":27A2
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   390
      Index           =   6
      Left            =   4380
      TabIndex        =   7
      ToolTipText     =   "Genera un TXT"
      Top             =   5955
      Width           =   525
      _Version        =   65536
      _ExtentX        =   917
      _ExtentY        =   698
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "NEWVIEW.frx":2CE4
   End
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5820
      Index           =   0
      Left            =   48
      MaxLength       =   20000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   48
      Width           =   9312
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   390
      Index           =   5
      Left            =   5895
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   5955
      Width           =   525
      _Version        =   65536
      _ExtentX        =   917
      _ExtentY        =   698
      _StockProps     =   78
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "NEWVIEW.frx":3226
   End
   Begin Threed.SSCommand ComOpc 
      Height          =   390
      Index           =   7
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   5955
      Width           =   525
      _Version        =   65536
      _ExtentX        =   917
      _ExtentY        =   698
      _StockProps     =   78
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "NEWVIEW.frx":3330
   End
   Begin VB.Label LBL 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Página   2 de  14"
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
      Height          =   210
      Left            =   2970
      TabIndex        =   1
      Top             =   6030
      Width           =   1200
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Posicion    As Long
Dim PosIni      As Long
Dim Fila        As Integer

Private Sub ComOpc_Click(Index As Integer)
Dim COMANDO     As String
Dim Id          As Variant
Dim cfile       As String

    Me.MousePointer = 11
    Select Case Index
      Case 0
        Texto(PosIni).Visible = False
        Texto(0).Visible = True
        PosIni = 0
        LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
        txtPag.Text = Str(PosIni + 1)
        txtPag.SetFocus
      Case 1
        If PosIni <> 0 Then
            Texto(PosIni).Visible = False
            PosIni = PosIni - 1
            Texto(PosIni).Visible = True
        End If
        LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
        txtPag.Text = Str(PosIni + 1)
        txtPag.SetFocus
      Case 2
        If PosIni <> Posicion Then
            Texto(PosIni).Visible = False
            PosIni = PosIni + 1
            Texto(PosIni).Visible = True
        End If
        LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
        txtPag.Text = Str(PosIni + 1)
        txtPag.SetFocus
      Case 3
        Texto(PosIni).Visible = False
        Texto(Posicion).Visible = True
        PosIni = Posicion
        LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
        txtPag.Text = Str(PosIni + 1)
        txtPag.SetFocus
      Case 4
        IMPRIME
      Case 5
        Unload Me
        Exit Sub
      Case 6
        COMANDO = Trim(wusuario) & ".TXT"
        Id = Shell("NOTEPAD " + COMANDO$, 1)
      Case 7
        If Len(Trim(wfile)) > 0 Then
            cfile = Trim(wrutabancos) & Trim(wusuario) & ".TXT"
            Open Trim(wrutabancos) & "\COPIA.BAT" For Output As #2
            Print #2, "COPY " & cfile & Space(1) & "A:\" & wfile
            Close #2
            Id = Shell(Trim(wrutabancos) & "\COPIA.BAT")
        End If
    End Select
    Me.MousePointer = 1
    
End Sub

Private Sub Form_Load()
Dim X       As Integer
Dim xCad    As String
Dim xLinea  As String

    Me.MousePointer = 11
'''''    Me.Width = menu.Width - 100
'''''    Me.Height = menu.Height - 360
'''''    Texto(0).Width = Me.Width - 80
'''''    Texto(0).Height = Me.Height - 950
'''''    For X = 0 To 7
'''''       ComOpc(X).Top = Texto(0).Height + 100
'''''    Next X
'''''    OPCPRE.Top = Texto(0).Height + 100
'''''    txtPag.Top = Texto(0).Height + 100
'''''    LBL.Top = Texto(0).Height + 100
'''''    Me.Refresh
    '****
    PosIni = 0
    Posicion = 0
    xCad = ""
    Open Trim(wrutabancos) & (wusuario) & ".TXT" For Input As #1
    Do While Not EOF(1)
        xLinea = ""
        Line Input #1, xLinea
        If xLinea <> "" Then
            If Asc(xLinea) = 12 Then
                Posicion = Posicion + 1
                Load Texto(Posicion)
                Texto(Posicion).Text = ""
                xLinea = ""
            End If
        End If
        Texto(Posicion).Text = Texto(Posicion).Text & Chr(13) & Chr(10) & xLinea
    Loop
    '***
    CAJA.Min = 1: CAJA.Max = Posicion + 1
    LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
    Me.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    App.Title = ""
    Close #1
End Sub

Private Sub IMPRIME()
On Error GoTo Fin_Inprime
Dim Destino     As Integer
Dim xFontSize   As String
Dim xFontName   As String
Dim xLinea      As String

    Fila = 1
    Destino = 1
    CAJA.Flags = &H0&
    CAJA.FromPage = 0: CAJA.ToPage = 0
    CAJA.CancelError = True
    CAJA.Action = 5
    
    SeteaImpresora Destino

    If App.Title = "Ministerio" Then
       If Destino = 1 Then
          xFontName = Printer.FontName
          xFontSize = Printer.FontSize
          Printer.FontName = "Lucida Console"
          Printer.FontSize = 6
        End If
    End If
    Select Case CAJA.Flags
           Case 16 'Imprimir Todos
                Open Trim(wusuario) & ".Txt" For Input As #2
                Do While Not EOF(2)
                    Line Input #2, xLinea
                    writexy 1, Fila, 1, RTrim(xLinea)
                    Fila = Fila + 1
                    If xLinea <> "" Then
                        If Asc(xLinea) = 12 Then
                            Printer.NewPage
                            Fila = 1
                        End If
                    End If
                Loop
                Printer.EndDoc
                Close #2
           Case 17 'Imprimir por Seleccion
                'Aqui va el Codigo de Seleccion
           Case 18 'Imprimir por Rango
                ImprimirRango
    End Select
    If App.Title = "Ministerio" Then
       If Destino = 1 Then
          Printer.FontName = xFontName
          Printer.FontSize = xFontSize
        End If
    End If
Exit Sub
Fin_Inprime:
    Select Case Err
           Case 55
                  Close #2
                  Resume
           Case 32755
                Exit Sub
           Case Else:
                MsgBox "Ocurrio el Siguiente Error:" + Chr(13) + Error, 32, "Advertencia"
                Exit Sub
    End Select
End Sub

Private Sub ImprimirRango()
Dim n As Integer
Dim xCadena As String, xAux As String
Dim xIni As Integer, xFin As Integer
Dim xPos As Integer
On Error GoTo ErrRango
   Fila = 1
   xIni = CAJA.FromPage
   xFin = CAJA.ToPage
   If (xIni > 0) And (xFin >= xIni) Then
       For n = xIni - 1 To xFin - 1
           xCadena = Texto(n) + Chr(13)
           Do
            xPos = InStr(xCadena, Chr(13))
            xAux = Left(xCadena, (xPos - 1))
            xCadena = Right(xCadena, Len(xCadena) - IIf(xPos < Len(xCadena), xPos + 1, xPos))
            'WriteXY 1, Fila, 1, xAux
            writexy xAux, Fila, 1, 1
            Fila = Fila + 1
           Loop Until InStr(xCadena, Chr(13)) = 0
           If xFin > 1 Then
              Fila = 1
              Printer.NewPage
           End If
       Next n
       Printer.EndDoc
    End If
Exit Sub
ErrRango:
    MsgBox "Ocurrio el Siguiente Error..!" + Chr(13) + Err, 32, "Advertencia"
    Exit Sub
End Sub

Private Function Pagina(Direc) As String
    
'    Dim Cadena   As String
'    Dim Car      As String * 1
'
'    If POSICION + 2 = LOF(1) Then POSICION = PosIni
'
'    If Direc Then
'      If PosIni > POSICION Then
'        POSICION = PosIni
'      Else
'        PosIni = POSICION
'      End If
'    Else
 '     If PosIni > POSICION Then
''        PosIni = POSICION
'      Else
'        POSICION = PosIni
'      End If
'    End If
 '
 '   If Direc Then
'      If POSICION > 0 Then
''        POSICION = POSICION + 3
 '     Else
 '       POSICION = POSICION + 1
 '     End If
 '   Else
 '     POSICION = POSICION - 1
 '   End If
 '   If POSICION > 0 Then
 '   Seek #1, POSICION
 '   Do While POSICION < LOF(1) And POSICION > 0
 '     If Direc = 0 Then
 '       Seek #1, POSICION
 '     End If
 '     Car = Input(1, #1)
 '     If Asc(Car) = 12 Then
 '       If Asc(Mid(Cadena, 1, 1)) = 13 And Asc(Mid(Cadena, 2, 1)) = 10 Then
 '         Cadena = Mid(Cadena, 3, Len(Cadena) - 2)
 '       End If
 '       Pagina = Cadena
 '       Exit Function
 '     Else
 '       If Direc Then
 '         POSICION = POSICION + 1
 '         Cadena = Cadena & Car
 '       Else
'          POSICION = POSICION - 1
'          Cadena = Car & Cadena
'        End If
'      End If
'    Loop
'    End If
'    If Trim(Cadena) <> "" Then
'      Pagina = Cadena
'    Else
'      Pagina = texto.Text
'    End If
End Function

Private Sub SeteaImpresora(PImpresion As Integer)
On Error GoTo ErrPrn
    
    If PImpresion = 0 Then
        Open Trim(wusuario) & ".TXT" For Output As #1
    Else
        Printer.ScaleMode = 4
        Printer.FontName = "Courier"
        Printer.FontSize = 12
    End If
    Exit Sub

ErrPrn:
    
    If Err = 380 Then  'Invalid Property Value
        Printer.FontName = "Courier"
        Printer.FontSize = 7.5
    Else
        MsgBox Error$(Err)
    End If
    Exit Sub

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Texto_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtPag_Change()
    If Val(txtPag.Text) > 1 And Val(txtPag.Text) < (Posicion + 1) Then
        Texto(PosIni).Visible = False
        PosIni = Val(txtPag.Text) - 1
        Texto(PosIni).Visible = True
        LBL.Caption = "Página " & Format(PosIni + 1, "###") & " De " & Format(Posicion + 1, "###")
    End If
End Sub

