VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form New_Importaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costos Estimados de Importación"
   ClientHeight    =   9375
   ClientLeft      =   330
   ClientTop       =   2055
   ClientWidth     =   14295
   ControlBox      =   0   'False
   Icon            =   "New_Importaciones.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabImporta 
      Height          =   7875
      Left            =   60
      TabIndex        =   8
      Top             =   420
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   13891
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "New_Importaciones.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dxDBMerca"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Detalle de Importación"
      TabPicture(1)   =   "New_Importaciones.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdInc"
      Tab(1).Control(1)=   "ChkShow"
      Tab(1).Control(2)=   "dxDBImport"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Otros Gastos"
      TabPicture(2)   =   "New_Importaciones.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dxDBGastos"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   10395
         Begin VB.ComboBox CboAgente 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox TxtZonaDet 
            Height          =   315
            Left            =   7980
            TabIndex        =   58
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox TxtCant 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5520
            TabIndex        =   57
            Text            =   "0.00"
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox CboUm 
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   480
            Width           =   915
         End
         Begin VB.ComboBox CboZona 
            Height          =   315
            Left            =   4500
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1080
            Width           =   2775
         End
         Begin VB.ComboBox CboImporta 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   480
            Width           =   3495
         End
         Begin VB.ComboBox CboVia 
            Height          =   315
            Left            =   7980
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label7 
            Caption         =   "Agente"
            Height          =   195
            Left            =   240
            TabIndex        =   64
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            Caption         =   "U.M."
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4080
            TabIndex        =   62
            Top             =   540
            Width           =   450
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            Caption         =   "Zona"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4020
            TabIndex        =   61
            Top             =   1140
            Width           =   450
         End
         Begin VB.Label Label14 
            Caption         =   "Embarcador"
            Height          =   195
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Via"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   6780
            TabIndex        =   59
            Top             =   540
            Width           =   990
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2895
         Left            =   10620
         TabIndex        =   44
         Top             =   360
         Width           =   3375
         Begin VB.TextBox TxtNumOri 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1020
            MaxLength       =   15
            TabIndex        =   48
            Top             =   600
            Width           =   2115
         End
         Begin VB.TextBox TxtNumFac 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1620
            MaxLength       =   15
            TabIndex        =   47
            Top             =   180
            Width           =   1545
         End
         Begin VB.TextBox TxtSerie 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1020
            MaxLength       =   7
            TabIndex        =   46
            Top             =   180
            Width           =   540
         End
         Begin VB.TextBox TxtRefere 
            Height          =   1740
            Left            =   1020
            MaxLength       =   80
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            Caption         =   "Imp.Origen"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   660
            Width           =   780
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            Caption         =   "Fac./Flete"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            Caption         =   "Referencia"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   49
            Top             =   1020
            Width           =   780
         End
      End
      Begin VB.CommandButton CmdInc 
         Caption         =   """%"""
         Height          =   255
         Left            =   -61560
         TabIndex        =   6
         ToolTipText     =   "Calcula Incidencias"
         Top             =   420
         Width           =   495
      End
      Begin VB.CheckBox ChkShow 
         Appearance      =   0  'Flat
         Caption         =   "Muestra sólo rubros usados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3315
         Begin VB.TextBox TxtImportacion 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.TextBox txtnumero 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1500
            MaxLength       =   7
            TabIndex        =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   1680
         End
         Begin aBoxCtl.aBox aboFecha 
            Height          =   285
            Left            =   1500
            TabIndex        =   1
            Top             =   780
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            ABoxType        =   ""
            MinValue        =   "D10000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   2
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
            Text            =   "13/06/2012"
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
            ButtonPicture   =   "New_Importaciones.frx":0060
            ButtonWidth     =   19
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
         Begin VB.Label LblImportacion 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Importación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   420
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label LblNumero 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "N° Presupuesto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   420
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Fecha de Cálculo"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1290
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   3540
         TabIndex        =   9
         Top             =   360
         Width           =   6975
         Begin VB.TextBox TxtRucPrv 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1200
            MaxLength       =   11
            TabIndex        =   2
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox TxtCodCli 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1200
            MaxLength       =   11
            TabIndex        =   3
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label LblTelPrv 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5400
            TabIndex        =   68
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label LblDirPrv 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1200
            TabIndex        =   67
            Top             =   540
            Width           =   3735
         End
         Begin VB.Label LblNomPrv 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2760
            TabIndex        =   66
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label LblNomCli 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2760
            TabIndex        =   65
            Top             =   900
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "Dirección"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Telf."
            Height          =   195
            Left            =   5040
            TabIndex        =   11
            Top             =   600
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   900
            Width           =   1080
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBMerca 
         Height          =   4485
         Left            =   120
         OleObjectBlob   =   "New_Importaciones.frx":05FA
         TabIndex        =   40
         Top             =   3300
         Width           =   13890
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBImport 
         Height          =   5085
         Left            =   -74880
         OleObjectBlob   =   "New_Importaciones.frx":53BF
         TabIndex        =   41
         Top             =   720
         Width           =   13830
      End
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   18
         Top             =   5760
         Width           =   13815
         Begin VB.TextBox TxtUtil 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10560
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label LblGastoFin 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   42
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label LblTotalS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9120
            TabIndex        =   34
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label LblIgvS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7680
            TabIndex        =   33
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label LblCostoS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   32
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label LblCostoI 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   28
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label LblIgvI 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7680
            TabIndex        =   27
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label LblTotalI 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9120
            TabIndex        =   26
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OTROS GASTOS"
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   180
            Width           =   6075
         End
         Begin VB.Label LblTot 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9120
            TabIndex        =   24
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label LblIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7680
            TabIndex        =   23
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label LblCosto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   22
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFC0&
            BorderWidth     =   3
            X1              =   120
            X2              =   10560
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "UTILIDAD BRUTA DITEC ESTIMADA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   6075
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VALOR DE VENTA SUGERIDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   1140
            Width           =   6075
         End
         Begin VB.Label Label15 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COSTO TOTAL DE IMPORTACION CON FINANCIAMIENTOS Y COMISIONES US$"
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   540
            Width           =   6075
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGastos 
         Height          =   6885
         Left            =   -74880
         OleObjectBlob   =   "New_Importaciones.frx":AA7D
         TabIndex        =   43
         Top             =   720
         Width           =   13890
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   635
      ButtonWidth     =   1879
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar   "
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar  "
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resumen"
            Object.ToolTipText     =   "Resumen de Importación"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "O.Comp."
            Object.ToolTipText     =   "Importar Orden de Compra"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "T.Comp."
            Object.ToolTipText     =   "Generar Costo conTrabajo Complementario"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir        "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   9240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":FCA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":10240
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":107DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":10D74
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":1130E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":118A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":11E42
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":123DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "New_Importaciones.frx":12976
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   60
      TabIndex        =   29
      Top             =   8280
      Width           =   14115
      Begin VB.CheckBox ChkAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Aprobar Presupuesto"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   11640
         Picture         =   "New_Importaciones.frx":12F10
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1155
      End
      Begin VB.CheckBox ChkCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Cerrar Presupuesto"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   12840
         Picture         =   "New_Importaciones.frx":1349A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label PnlFactorConFin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00000000000000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   36
         Top             =   480
         Width           =   2715
      End
      Begin VB.Label PnlFactorSinFin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00000000000000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   35
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Factor Con Financiammiento"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   540
         Width           =   2010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factor Sin Financiammiento"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1950
      End
   End
End
Attribute VB_Name = "New_Importaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zFlete As Double
Dim nGasFin As Double
Dim ngFob As Double
Dim ngDesaduana As Double
Dim ngAdela As Double
Dim nfFob As Double
Dim nfDesaduana As Double
Dim nfAdela As Double
Dim nComi As Double
Dim TbCabImport1 As New ADODB.Recordset
Dim sw_nuevo_doc1 As Boolean
Dim sw_detalle1 As Boolean
Dim sw_nuevo_item1 As Boolean
Dim tempo As New ADODB.Connection
Dim sw_cabecera1 As Boolean
Dim TbDetOrden1 As New ADODB.Recordset
Dim tbProducto1 As New ADODB.Recordset
Dim Precios As Boolean
Dim cvalores As String
Dim cmes As String
Dim falta As String
Dim wcantord    As Integer
Dim amovs_cab(0 To 40)  As a_grabacion
Dim amovs_det(0 To 23)  As a_grabacion
Dim amovs_cab1(0 To 7)  As a_grabacion
Dim amovs_det1(0 To 10)  As a_grabacion
Dim factor As Double
Dim ctipo As String
Dim RSDETALLE As New ADODB.Recordset
Dim flag As Boolean
Dim nFlete As Double
Dim Temp As New ADODB.Connection
Dim tbcostosimp1 As New ADODB.Recordset
Dim TbAgente1 As New ADODB.Recordset
Dim cambio As Double
Dim Nfob As Double
Dim sw_cabecera As Boolean
Dim sw_detalle As Boolean
'Dim sw_nuevo_doc As Boolean
Dim sw_nuevo_item As Boolean
Dim sw_ayuda As Boolean
Dim sw_ayuda_oc     As Boolean
Dim wgrabar As Byte
Public wnumimp As String
Dim inicio As Boolean
Dim rsttemp As ADODB.Recordset
Dim graba As Boolean
Public CODPROV As String

Private Sub abofecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub abofechaconfirma_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtOrdenDespacho.SetFocus
End If
End Sub


Private Function CalculaFlete() As Double
 
Select Case Trim(CboUm.Text)
Case "KG"
    Select Case right(Me.CboImporta.Text, 4)
        Case "0010"
            If IsNumeric(TxtCant.Text) And Val(TxtCant.Text) > 0 Then
                sql = "SELECT TIPO, PESO_DE, PESO_A, FACTOR, FAC_UNID,TAR_PAD, ID_ZONA, COSTO,tip_cam,mon "
                sql = sql & "From IMPORT_TARIFA_0010 "
                sql = sql & "WHERE (((TIPO)='a') "
                sql = sql & "AND ((PESO_DE)<=" & Val(TxtCant.Text) & ") "
                sql = sql & "AND ((PESO_A)>=" & Val(TxtCant.Text) & ") "
                sql = sql & "AND ((ID_ZONA)=" & Val(left(CboZona.Text, 2)) & "))"
                If Rs.State = 1 Then Rs.Close
                Rs.Open sql, cnn_dbbancos, 3, 1
                If Rs.RecordCount > 0 Then
                    If Rs!factor = True Then
                        PESO_PADRE = traerCampo("IMPORT_TARIFA", "PESO_DE", "ID_TARIFA", Rs!TAR_PAD, "")
                        COSTO_PADRE = traerCampo("IMPORT_TARIFA", "COSTO", "ID_TARIFA", Rs!TAR_PAD, "")
                        UNID_EXCEDE = (Val(TxtCant.Text) - Val(PESO_PADRE)) / Val(Rs!FAC_UNID & "")
                        VALOR_FLETE = (UNID_EXCEDE * Val(Rs!costo & "")) + COSTO_PADRE
                        
                        If Rs!mon = "S" Then
                            nFlete = VALOR_FLETE / Val(Rs!tip_cam & "")
                        ElseIf Rs!mon = "D" Then
                            nFlete = VALOR_FLETE
                        End If
                        nFlete = Fix(nFlete) / 10
                        nFlete = Round(nFlete, 0)
                        nFlete = nFlete * 10
                        'MsgBox nFlete
                        CalculaFlete = nFlete
                    Else
                        If Rs!mon = "S" Then
                            nFlete = Rs!costo / Val(Rs!tip_cam & "")
                        ElseIf Rs!mon = "D" Then
                            nFlete = Rs!costo
                        End If
                        nFlete = Fix(nFlete) / 10
                        nFlete = Round(nFlete, 0) + 1
                        nFlete = nFlete * 10
        '                MsgBox nFlete
                        CalculaFlete = nFlete
                    End If
                Else
                    MsgBox "No existe escala registrada para este peso", vbInformation, "Flete"
                End If
            End If
        Case "0295"
            NFOBTOT = Me.dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue
            npeso = Val(TxtCant.Text)
            Dim SiAfecto As Double, NoAfecto As Double
            If Val(NFOBTOT) > 0 And Val(npeso) > 0 Then
                sql = "SELECT * FROM IMPORT_TARIFA_0380 order by id_tarifa"
                If Rs.State = 1 Then Rs.Close
                Rs.Open sql, cnn_dbbancos, 3, 1
                Do While Not Rs.EOF
                    If Rs!afecto & "" = -1 Then
                        Select Case Rs!Base & ""
                        
                        Case "P"
                        
                        Case "F"
                        
                        Case "-"
                        
                        End Select
                    Else
                    End If
                    Rs.MoveNext
                Loop
            End If
    End Select
End Select
End Function

Private Sub CboImporta_Click()
If left(CboImporta.Text, 4) <> "0000" Then
    CargaImport
    GridImport
End If
End Sub

Private Sub CboImporta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub CboUm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub CboVia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub CboZona_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub ChkAprobar_Click()
nval = ObtenerCampo("import_cab", "f4aprobado", "f4numimp", txtnumero.Text, "T", cnn_dbbancos)
If Len(Trim(nval)) = 0 Then
    
    ChkAprobar.Value = 0
    MsgBox "El presupuesto de importación no ha sido guardado", vbInformation, "Sistema de Logística"
    Exit Sub
End If
If nval = True Then Exit Sub
If MsgBox("¿Está seguro de aprobar el presupuesto?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
        amovs_cab(0).campo = "f4aprobado": amovs_cab(0).valor = IIf(ChkAprobar.Value = 1, -1, 0): amovs_cab(0).TIPO = "N"
        amovs_cab(1).campo = "f4importacion": amovs_cab(1).valor = NuevaImportacion: amovs_cab(1).TIPO = "T"
        GRABA_REGISTRO amovs_cab(), "import_cab", "M", 1, cnn_dbbancos, "f4numimp='" & txtnumero.Text & "'"
        
        TxtImportacion.Text = amovs_cab(1).valor
        TxtImportacion.Visible = True
        LblImportacion.Visible = True
        txtnumero.Visible = False
        LblNumero.Visible = False
        ChkAprobar.Enabled = False
        For I = 1 To 9
            amovs_cab(0).campo = "f4NUMIMP": amovs_cab(0).valor = txtnumero.Text: amovs_cab(0).TIPO = "T"
            amovs_cab(1).campo = "ID_SEG": amovs_cab(1).valor = I: amovs_cab(1).TIPO = "N"
            If I = 1 Then
                amovs_cab(2).campo = "ESTADO": amovs_cab(2).valor = I: amovs_cab(2).TIPO = "N"
            Else
                amovs_cab(2).campo = "ESTADO": amovs_cab(2).valor = 0: amovs_cab(2).TIPO = "N"
            End If
            GRABA_REGISTRO amovs_cab(), "import_MOV", "A", 1, cnn_dbbancos, ""
        Next

End If

End Sub
Private Function NuevaImportacion() As String
    csql = "select top 1 f4importacion from import_cab where f4aprobado=-1 order by f4importacion desc"
    If Rs.State = 1 Then Rs.Close
    Rs.Open csql, cnn_dbbancos, 3, 1
    If Rs.RecordCount > 0 Then
        NuevaImportacion = "I" & Format(Val(Mid(Rs!f4importacion & "", 2, 9)) + 1, "000000000")
    Else
        NuevaImportacion = "I000000001"
    End If
End Function

Private Sub ChkShow_Click()
GridImport
GridGastos
End Sub

Private Sub CmdFlete_Click()
SwCalc = True
CalculaFlete
SwCalc = False
End Sub

Private Sub ChkShow_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub CmdGastos_Click()
Nfob = Val(dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue & "")
NCOSTO = Val(dxDBImport.Columns.ColumnByFieldName("t2costo").SummaryFooterValue & "")

zFlete = 0
sql = "SELECT Sum(Val(TMP_IMP_DET2.T2Total+'')) AS Total "
sql = sql & "FROM TMP_IMP_DET2 INNER JOIN IMP_EMBxPLANT "
sql = sql & "ON (TMP_IMP_DET2.T2Grupo = IMP_EMBxPLANT.F1GRUPO) "
sql = sql & "AND (TMP_IMP_DET2.T2SubGrupo = IMP_EMBxPLANT.F1SUBGRUPO) "
sql = sql & "GROUP BY TMP_IMP_DET2.T2Calcular, IMP_EMBxPLANT.F1EMBARCA, "
sql = sql & "IMP_EMBxPLANT.F1ADUANA HAVING (((TMP_IMP_DET2.T2Calcular)=-1) "
sql = sql & "AND ((IMP_EMBxPLANT.F1EMBARCA)='" & left(CboImporta.Text, 4) & "') "
sql = sql & "AND ((IMP_EMBxPLANT.F1ADUANA)=-1))"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, tempo, 3, 1
If Rs.RecordCount > 0 Then zFlete = Val(Rs!TOTAL & "")
nGasFin = 0
nComi = 0

nfFob = (LblCostoS.Caption * FMFob.Text * FIFob.Text * 0.014)
nfDesaduana = (LblCostoS.Caption * FmDesaduana.Text * FIDesaduana.Text * 0.014)
nfAdela = (LblCostoS.Caption * FMAdela.Text * FIAdela.Text * 0.014)

'nComi = (Nfob + (Nfob * GMFob.Text * GIFob.Text / 100) + (zflete * GMDesaduana.Text * GIDesaduana.Text / 100) + (ncosto + nComi + ngFob + ngDesaduana) * ((1 + (TxtUtil.Text / 100)) * (1 + (wigv / 100)) * (GIAdela.Text / 100)) / (1 - ((1 + (TxtUtil.Text / 100)) * (1 + (wigv / 100)) * (GIAdela.Text / 100)))) * 2.5 / 100

ngFob = (Nfob * GMFob.Text * GIFob.Text / 100)

ngDesaduana = (zFlete * GMDesaduana.Text * GIDesaduana.Text / 100)

X = (1 + (wIgv / 100)) * (GIAdela.Text / 100) * GMAdela.Text
nutil = (1 + (TxtUtil.Text / 100))
ncom = (TxtCom.Text / 100)
nval = (NCOSTO + ngFob + ngDesaduana)
zup = (nval + (nval * nutil * X)) * nutil
zdown = 1 - (nutil * ncom) - (((X ^ 2) + (X * ncom)) * (nutil ^ 2))
xtotal = zup / zdown
nComi = xtotal * ncom
LblComision.Caption = Format(nComi, "#,##0.00")
ngAdela = (NCOSTO + nComi + ngFob + ngDesaduana) * ((1 + (TxtUtil.Text / 100)) * (1 + (wIgv / 100)) * (GIAdela.Text / 100)) / (1 - ((1 + (TxtUtil.Text / 100)) * (1 + (wIgv / 100)) * (GIAdela.Text / 100)))

nGasFin = nfFob + nfDesaduana + nfAdela + ngFob + ngDesaduana + ngAdela + nComi
LblGastoFin.Caption = nGasFin
Call TxtGastoFin_LostFocus
Call TxtUtil_LostFocus
End Sub

Private Sub CmdInc_Click()
On Error Resume Next
If tempo.State = 0 Then tempo.Open
sql = "update tmp_imp_det2 set t2inciden=''"
tempo.Execute sql
  'AlmacenaQuery_sql sql, tempo
sql = "update tmp_imp_det2 set t2inciden=val(t2total)*100/val(" & Format(LblCostoI.Caption, "#0.00") & ") where t2cal_inc=-1 and val(t2total&'')>0"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
'GridImport
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
If dxDBGrid1.Dataset.State = 2 Then dxDBGrid1.Dataset.Post
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Cant." Or Column.Caption = "Fob. Unit" Then
    Text = Format(Text, "#,#0.00")
End If

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 46 Or KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Or dxDBGrid1.Dataset.RecNo = 0 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
            sw_nuevo_item = False
        End If
    End If
End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)

Set RSDETALLE = New ADODB.Recordset
Select Case KeyCode
Case 113: '--- F2
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName
    Case "F5CODPRO":
        wcodgasto = "": wnomgasto = ""
        '------------------------------------------------------
        sw_ayuda = True
        ayuda_gastos.Show 1
        sw_ayuda = False
        If Len(Trim(wcodgasto)) > 0 Then
            sql = "SELECT IIF(ISNULL(BF9GIN.GRUPOFLUJO),'9999',BF9GIN.GRUPOFLUJO) AS GRUPO, IIF(ISNULL(GRUPOS_FLUJO.NOMBRE),'OTROS GASTOS',GRUPOS_FLUJO.NOMBRE) AS NOMBRE " & _
                "FROM BF9GIN LEFT JOIN GRUPOS_FLUJO ON BF9GIN.GRUPOFLUJO = GRUPOS_FLUJO.CODIGO WHERE BF9GIN.CODIGO='" & wcodgasto & "'"
            If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
            RSDETALLE.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSDETALLE.EOF Then
                dxDBGrid1.Dataset.Edit
                '--------------------------------------------------------------------
                dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodgasto
                dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wnomgasto
                dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value = RSDETALLE.Fields("GRUPO")
                dxDBGrid1.Columns.ColumnByFieldName("NOMBRE").Value = RSDETALLE.Fields("NOMBRE")
                dxDBGrid1.Columns.ColumnByFieldName("F3CHECK").Value = True
                dxDBGrid1.Dataset.Post
            End If
        End If
        dxDBGrid1.Columns.FocusedIndex = 3
    End Select
'Case 115: '--- F4
'        If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logistica") = vbYes Then
'        sw_nuevo_item = True
''        sw_detalle = True
'            If dxDBGrid1.Count = 1 Then
'                dxDBGrid1.Dataset.Delete
'                AdicionaItem
'            Else
'                dxDBGrid1.Dataset.Delete
''                RENUMERARITEMS
'            End If
'        End If
End Select

End Sub

Private Sub CmdInc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub dxDBGastos_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim nmes  As Double, nInteres As Double
With Me.dxDBGastos
    If .Columns.ColumnByFieldName("t3calcular").Value = True Then
        
        .Columns.ColumnByFieldName("t3calcular").Value = False
        .Dataset.Edit
        .Columns.ColumnByFieldName("t3costo").Value = ""
        .Columns.ColumnByFieldName("t3igv").Value = ""
        .Columns.ColumnByFieldName("t3total").Value = ""
        .Columns.ColumnByFieldName("t3inciden").Value = ""
        .Dataset.Post
    Else
        .Dataset.Edit
        ntotal = 0
        .Columns.ColumnByFieldName("t3calcular").Value = True
        nmes = Val(.Columns.ColumnByFieldName("t3meses").Value & "")
        nInteres = Val(.Columns.ColumnByFieldName("t3interes").Value & "")
        
        Nfob = Val(dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue & "")
        NCOSTO = Val(dxDBImport.Columns.ColumnByFieldName("t2costo").SummaryFooterValue & "")

        Select Case Me.dxDBGastos.Columns.ColumnByFieldName("T3KEY").Value & ""
        
        Case "00010001", "00010002", "00010003"
            ngFob = GastoFob(Nfob, "00020001")
            ngDesaduana = GastoDesaduana("00020002")
            tInteres = 2
            tMes = 1
            
            X = (1 + (wIgv / 100)) * (tInteres / 100) * tMes
            nutil = (1 + (TxtUtil.Text / 100))
            ncom = (ObtieneComision / 100)
            nval = (NCOSTO + ngFob + ngDesaduana)
            zup = (nval + (nval * nutil * X)) * nutil
            zdown = 1 - (nutil * ncom) - (((X ^ 2) + (X * ncom)) * (nutil ^ 2))
            xtotal = zup / zdown
            ntotal = (Format(xtotal, "#0.00") * nmes * nInteres / 100 * 0.014)
        Case "00020001"
            ntotal = GastoFob(Nfob, "00020001")
        Case "00020002"
            ntotal = GastoDesaduana("00020002")
        Case "00020003"
            ngFob = GastoFob(Nfob, "00020001")
            ngDesaduana = GastoDesaduana("00020002")
            X = (1 + (wIgv / 100)) * (nInteres / 100) * nmes
            nutil = (1 + (TxtUtil.Text / 100))
            ncom = (ObtieneComision / 100)
            nval = (NCOSTO + ngFob + ngDesaduana)
            zup = (nval + (nval * nutil * X)) * nutil
            zdown = 1 - (nutil * ncom) - (((X ^ 2) + (X * ncom)) * (nutil ^ 2))
            xtotal = zup / zdown
            nComi = xtotal * ncom
            ntotal = (NCOSTO + nComi + ngFob + ngDesaduana) * (nutil * (1 + (wIgv / 100)) * (nInteres / 100)) / (1 - (nutil * (1 + (wIgv / 100)) * (nInteres / 100)))
            
        Case "00030001"
            ngFob = GastoFob(Nfob, "00020001")
            ngDesaduana = GastoDesaduana("00020002")
            'obtener meses e interes de adelanto
            sql = "SELECT * FROM TMP_IMP_DET3 where TMP_IMP_DET3.t3key='00020003'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, tempo, 3, 1
            If Rs.RecordCount > 0 Then
                nmes = Rs!t3meses
                nInteres = Rs!t3interes
            End If
            'obtener porcentaje de comision
            sql = "SELECT * FROM TMP_IMP_DET3 where TMP_IMP_DET3.t3key='00030001'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, tempo, 3, 1
            If Rs.RecordCount > 0 Then
                ncom = Val(Rs!t3interes & "") / 100
            End If
            '********************************
            X = (1 + (wIgv / 100)) * (nInteres / 100) * nmes
            nutil = (1 + (TxtUtil.Text / 100))
            nval = (NCOSTO + ngFob + ngDesaduana)
            zup = (nval + (nval * nutil * X)) * nutil
            zdown = 1 - (nutil * ncom) - (((X ^ 2) + (X * ncom)) * (nutil ^ 2))
            xtotal = zup / zdown
            ntotal = xtotal * ncom
        End Select
        If dxDBGastos.Columns.ColumnByFieldName("T3KEY").Value = "00030001" Then
            .Columns.ColumnByFieldName("t3costo").Value = ntotal
            .Columns.ColumnByFieldName("t3igv").Value = ""
            .Columns.ColumnByFieldName("t3total").Value = ""
            .Columns.ColumnByFieldName("t3inciden").Value = ""
        Else
            .Columns.ColumnByFieldName("t3costo").Value = ntotal
            .Columns.ColumnByFieldName("t3igv").Value = ntotal * wIgv / 100
            .Columns.ColumnByFieldName("t3total").Value = Val(.Columns.ColumnByFieldName("t3costo").Value) + Val(.Columns.ColumnByFieldName("t3igv").Value)
            .Columns.ColumnByFieldName("t3inciden").Value = ""
        End If
        .Dataset.Post
        If .Columns.ColumnByFieldName("t3total").Value = 0 Then
            .Columns.FocusedIndex = 5
        End If
    End If
End With
LblGastoFin.Caption = Format(Val(Me.dxDBGastos.Columns.ColumnByFieldName("T3COSTO").SummaryFooterValue), "###,###,##0.00")
CalculaMontos
End Sub
Private Function ObtieneComision() As Double
Dim Com As New ADODB.Recordset
ObtieneComision = 0
sql = "SELECT * FROM TMP_IMP_DET3 where TMP_IMP_DET3.t3key='00030001'"
If Com.State = 1 Then Com.Close
Com.Open sql, tempo, 3, 1
If Com.RecordCount > 0 Then
    ObtieneComision = Com!t3interes
End If
If Com.State = 1 Then Com.Close
Set Com = Nothing
End Function

Private Sub TrasladaCalculos()

End Sub

Private Function GastoFob(ValorFOB As Double, CodigoItem As String) As Double
Dim tMes As Double, tInteres As Double
'obtener meses e interes de adelanto
If tempo.State = 0 Then tempo.Open
sql = "SELECT * FROM TMP_IMP_DET3 where TMP_IMP_DET3.t3key='" & CodigoItem & "'"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, tempo, 3, 1
If Rs.RecordCount > 0 Then
    tMes = Val(Rs!t3meses & "")
    tInteres = Val(Rs!t3interes & "")
End If
GastoFob = (ValorFOB * tMes * tInteres / 100)
End Function
Private Function GastoDesaduana(CodigoItem As String) As Double
Dim tMes As Double, tInteres As Double

'obteniendo flete
zFlete = 0
sql = "SELECT Sum(Val(TMP_IMP_DET2.T2Total+'')) AS Total "
sql = sql & "FROM TMP_IMP_DET2 INNER JOIN IMP_EMBxPLANT "
sql = sql & "ON (TMP_IMP_DET2.T2Grupo = IMP_EMBxPLANT.F1GRUPO) "
sql = sql & "AND (TMP_IMP_DET2.T2SubGrupo = IMP_EMBxPLANT.F1SUBGRUPO) "
sql = sql & "GROUP BY TMP_IMP_DET2.T2Calcular, IMP_EMBxPLANT.F1EMBARCA, "
sql = sql & "IMP_EMBxPLANT.F1ADUANA HAVING (((TMP_IMP_DET2.T2Calcular)=-1) "
sql = sql & "AND ((IMP_EMBxPLANT.F1EMBARCA)='" & right(CboImporta.Text, 4) & "') "
sql = sql & "AND ((IMP_EMBxPLANT.F1ADUANA)=-1))"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, tempo, 3, 1
If Rs.RecordCount > 0 Then zFlete = Val(Rs!TOTAL & "")
'obtener meses e interes de adelanto
sql = "SELECT * FROM TMP_IMP_DET3 where TMP_IMP_DET3.t3key='" & CodigoItem & "'"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, tempo, 3, 1
If Rs.RecordCount > 0 Then
    tMes = Val(Rs!t3meses & "")
    tInteres = Val(Rs!t3interes & "")
End If

GastoDesaduana = (zFlete * tMes * tInteres / 100)
End Function

Private Sub dxDBGastos_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Incidencia" And Len(Trim(Text)) > 0 Then
    Text = Format(Text, "#0.00") & " %"
End If
If Column.Caption = "Costo" Or Column.Caption = "%" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,##0.00")
End If
End Sub

Private Sub dxDBGastos_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Incidencia" And Len(Trim(Text)) > 0 Then
    Text = Format(Text, "#0.00") & " %"
End If
If Column.Caption = "Costo" Or Column.Caption = "%" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,##0.00")
End If
End Sub

Private Sub dxDBGastos_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Incidencia" And Len(Trim(Text)) > 0 Then
    Text = Format(Text, "#0.00") & " %"
End If
If Column.Caption = "Costo" Or Column.Caption = "%" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,##0.00")
End If
End Sub

Private Sub dxDBImport_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim nEsp As Double
With Me.dxDBImport
'MsgBox .Columns.ColumnByFieldName("t2tipo").Value & ""
    If .Columns.ColumnByFieldName("t2calcular").Value = True Then
    'If State = .Columns.ColumnByFieldName("t2calcular").Value Then
        .Dataset.Edit
        .Columns.ColumnByFieldName("t2calcular").Value = False
        
        .Columns.ColumnByFieldName("t2costo").Value = ""
        .Columns.ColumnByFieldName("t2igv").Value = ""
        .Columns.ColumnByFieldName("t2total").Value = ""
        .Columns.ColumnByFieldName("t2inciden").Value = ""
        .Dataset.Post
    Else
        If tempo.State = 0 Then tempo.Open
        .Dataset.Edit
        ntotal = 0
        .Columns.ColumnByFieldName("t2calcular").Value = True
        nEsp = Val(dxDBImport.Columns.ColumnByFieldName("T2esp").Value & "")
        Select Case Me.dxDBImport.Columns.ColumnByFieldName("T2KEY").Value & ""
        
        Case "00010001"
            ntotal = Me.dxDBMerca.Columns.ColumnByFieldName("T1FOBTOT").SummaryFooterValue
        Case "00010002"
            ntotal = CalculaFlete
        Case "00010003"
            
            csql = "SELECT [TMP_IMP_DET2].[T2Total] From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2KEY)='00010002')) "
            csql = csql & "OR (((TMP_IMP_DET2.T2KEY)='00010001'))"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            If ntotal * nEsp / 100 < 75 Then
                ntotal = 75
            Else
                ntotal = ntotal * nEsp / 100
            End If
        
        Case "00020001"
            csql = "SELECT * FROM TMP_IMP_DET1 ORDER BY T1MODELO DESC"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                cod = Rs!T1CodPro
                csql = "SELECT IF5PLA.F5CODPRO, PARTIDA.F5PARTARA, PARTIDA.AD_VALOREM, "
                csql = csql & "PARTIDA.PERCEPCION, PARTIDA.PERMISO "
                csql = csql & "FROM IF5PLA INNER JOIN PARTIDA ON "
                csql = csql & "IF5PLA.F5PARTARA = PARTIDA.F5PARTARA "
                csql = csql & "WHERE (((IF5PLA.F5CODPRO)='" & cod & "'))"
                If Rs.State = 1 Then Rs.Close
                Rs.Open csql, cnn_dbbancos, 3, 1
                If Rs.RecordCount > 0 Then
                    Rs.MoveFirst
                    ADVAL = Val(Rs!AD_VALOREM & "")
                    dxDBImport.Columns.ColumnByFieldName("T2esp").Value = ADVAL
                    csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
                    csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0001') "
                    csql = csql & "OR ((TMP_IMP_DET2.T2Grupo)='0002'))"
                    If Rs.State = 1 Then Rs.Close
                    Rs.Open csql, tempo, 3, 1
                    If Rs.RecordCount > 0 Then Rs.MoveFirst
                    Do While Not Rs.EOF
                        ntotal = ntotal + Val(Rs!t2total & "")
                        Rs.MoveNext
                    Loop
                    ntotal = ntotal * ADVAL / 100
                End If
            End If
        Case "00020002"
            csql = "SELECT [TMP_IMP_DET2].[T2Total] From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0001')) "
            csql = csql & "OR (((TMP_IMP_DET2.T2KEY)='00020001'))"
            If tempo.State = 1 Then tempo.Close
            tempo.Open
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * nEsp / 100
        Case "00020003"
            csql = "SELECT [TMP_IMP_DET2].[T2Total] From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0001')) "
            csql = csql & "OR (((TMP_IMP_DET2.T2KEY)='00020001'))"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * nEsp / 100
        Case "00020004"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE (TMP_IMP_DET2.T2Grupo='0001') "
            csql = csql & "OR (TMP_IMP_DET2.T2Grupo='0002' AND TMP_IMP_DET2.T2KEY<>'00020004')"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
                Do While Not Rs.EOF
                    ntotal = ntotal + Val(Rs!t2total & "")
                    Rs.MoveNext
                Loop
                ntotal = ntotal * nEsp / 100
            
            
        Case "00030003"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0003') "
            csql = csql & "AND ((TMP_IMP_DET2.T2KEY)<>'00030003'))"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * wIgv / 100
        Case "00040002", "00040001"
            If Val(Me.dxDBImport.Columns.ColumnByFieldName("T2VALDEF").Value & "") > 0 Then
                ntotal = Val(Me.dxDBImport.Columns.ColumnByFieldName("T2VALDEF").Value & "")
            End If
        Case "00040009"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0004') "
            csql = csql & "AND ((TMP_IMP_DET2.T2KEY)<>'00040009'))"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * wIgv / 100
        Case "00050002"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE TMP_IMP_DET2.T2KEY='00010002'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * wIgv / 100
        Case "00050001"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE TMP_IMP_DET2.T2KEY='00010003'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * wIgv / 100
        Case "00050009"
            sql = "select * from IMPORT_CAB WHERE F4NUMIMP='" & TxtNumOri.Text & "'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open sql, cnn_dbbancos, 3, 1
            If Rs.RecordCount > 0 Then
                ntotal = Val(Rs!F4COSTOS & "") - Val(Rs!F4COSTOI & "")
                ntotal = 0.9 * ntotal / 3 * 0.68
            Else
                MsgBox "Debe ingresar un Codigo de Costo de Importación Origen Válido", vbCritical, "Sistema de Logística"
                .Columns.ColumnByFieldName("t2calcular").Value = False
                .Columns.ColumnByFieldName("t2costo").Value = ""
                .Columns.ColumnByFieldName("t2igv").Value = ""
                .Columns.ColumnByFieldName("t2total").Value = ""
                .Columns.ColumnByFieldName("t2inciden").Value = ""
            End If
        Case "00070001"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE TMP_IMP_DET2.T2KEY='00010002'"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * nEsp / 100
        Case "00070008"
            csql = "SELECT TMP_IMP_DET2.T2Total From TMP_IMP_DET2 "
            csql = csql & "WHERE (((TMP_IMP_DET2.T2Grupo)='0007') "
            csql = csql & "AND ((TMP_IMP_DET2.T2KEY)<>'00070008'))"
            If Rs.State = 1 Then Rs.Close
            Rs.Open csql, tempo, 3, 1
            If Rs.RecordCount > 0 Then Rs.MoveFirst
            Do While Not Rs.EOF
                ntotal = ntotal + Val(Rs!t2total & "")
                Rs.MoveNext
            Loop
            ntotal = ntotal * wIgv / 100
        End Select
        
        If .Columns.ColumnByFieldName("t2tipo").Value & "" = "M" Then
            .Columns.ColumnByFieldName("t2costo").Value = ntotal
            .Columns.ColumnByFieldName("t2igv").Value = ""
            .Columns.ColumnByFieldName("t2total").Value = ntotal
        ElseIf .Columns.ColumnByFieldName("t2tipo").Value & "" = "I" Then
            .Columns.ColumnByFieldName("t2costo").Value = ""
            .Columns.ColumnByFieldName("t2igv").Value = ntotal
            .Columns.ColumnByFieldName("t2total").Value = ntotal
        Else
            .Columns.ColumnByFieldName("t2total").Value = ntotal
            .Columns.ColumnByFieldName("t2inciden").Value = ""
        End If
        .Dataset.Post
        If .Columns.ColumnByFieldName("t2total").Value = 0 Then
            .Columns.FocusedIndex = 5
        End If
    End If
End With
CalculaMontos
End Sub

Private Sub dxDBImport_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Incidencia" And Len(Trim(Text)) > 0 Then
    Text = Format(Text, "#0.00") & " %"
End If
If Column.Caption = "Costo" Or Column.Caption = "%" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,#0.00")
End If
End Sub

Private Sub dxDBImport_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Costo" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,#0.00")
End If
Color = RGB(215, 253, 255)
End Sub

Private Sub dxDBImport_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Incidencia" Then
    Text = Format(Text, "#0.00") & " %"
End If
If Column.Caption = "Costo" Or Column.Caption = "I.G.V." Or Column.Caption = "Total" Then
    Text = Format(Text, "#,#0.00")
End If
Color = RGB(255, 255, 255)
End Sub

Private Sub dxDBImport_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
With Me.dxDBImport
If .Columns.FocusedColumn.Caption = "Total" Then
    ntotal = .Columns.ColumnByFieldName("t2total").Value
    .Dataset.Edit
    If .Columns.ColumnByFieldName("t2tipo").Value & "" = "M" Then
        .Columns.ColumnByFieldName("t2costo").Value = ntotal
        .Columns.ColumnByFieldName("t2igv").Value = ""
    ElseIf .Columns.ColumnByFieldName("t2tipo").Value & "" = "I" Then
        .Columns.ColumnByFieldName("t2costo").Value = ""
        .Columns.ColumnByFieldName("t2igv").Value = ntotal
    Else
        .Columns.ColumnByFieldName("t2costo").Value = ""
        .Columns.ColumnByFieldName("t2igv").Value = ""
        .Columns.ColumnByFieldName("t2inciden").Value = ""
    End If
    If Val(.Columns.ColumnByFieldName("t2total").Value & "") > 0 Then
        .Columns.ColumnByFieldName("t2calcular").Value = True
    Else
        .Columns.ColumnByFieldName("t2calcular").Value = False
        .Columns.ColumnByFieldName("t2costo").Value = ""
        .Columns.ColumnByFieldName("t2igv").Value = ""
        .Columns.ColumnByFieldName("t2total").Value = ""
        .Columns.ColumnByFieldName("t2inciden").Value = ""
    End If
    .Dataset.Post
End If
End With
End Sub

Private Sub dxDBMerca_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Cant." Then
    Text = Format(Text, "#,#0.000")
End If
If Column.Caption = "FOB Unit." Or Column.Caption = "FOB Total" Or Column.Caption = "Costo Unit." Or Column.Caption = "Venta Unit." Then
    Text = Format(Text, "#,#0.00")
End If
If Column.Caption = "% Utilidad" Then
    Text = Format(Text, "#,#0.00")
End If
End Sub

Private Sub dxDBMerca_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Cant." Then
    Text = Format(Text, "#,#0.000")
End If
If Column.Caption = "FOB Unit." Or Column.Caption = "FOB Total" Or Column.Caption = "Costo Unit." Or Column.Caption = "Venta Unit." Then
    Text = Format(Text, "#,#0.00")
End If
If Column.Caption = "%Utilidad" Then
    Text = Format(Text, "#,#0.00")
End If
End Sub

Private Sub dxDBMerca_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
wcodproducto = ""
ayuda_productos.Show 1
If wcodproducto <> "" Then
    dxDBMerca.Dataset.Edit
    dxDBMerca.Columns.ColumnByFieldName("t1codpro").Value = wcodproducto
    dxDBMerca.Columns.ColumnByFieldName("t1modelo").Value = ObtenerCampo("if5pla", "f5modelo", "f5codpro", wcodproducto, "T", cnn_dbbancos)
    dxDBMerca.Columns.ColumnByFieldName("t1nompro").Value = ObtenerCampo("if5pla", "f5nompro", "f5codpro", wcodproducto, "T", cnn_dbbancos)
    dxDBMerca.Columns.ColumnByFieldName("t1desmar").Value = ObtenerCampo("ef2marcas", "f2desmar", "f2codmar", ObtenerCampo("if5pla", "f5nompro", "f5codpro", wcodproducto, "T", cnn_dbbancos), "T", cnn_dbbancos)
    dxDBMerca.Dataset.Post
End If
End Sub

Private Sub dxDBMerca_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
If dxDBMerca.Columns.FocusedIndex = 1 Then
    If Len(Trim(ObtenerCampo("IF5PLA", "F5NOMPRO", "F5CODPRO", Me.dxDBMerca.Columns.ColumnByFieldName("T1CODPRO").Value & "", "T", cnn_dbbancos))) = 0 Then
        MsgBox "El producto no existe", vbCritical, "Sistema de Logística"
        dxDBMerca.Dataset.Edit
        dxDBMerca.Columns.ColumnByFieldName("T1CODPRO").Value = ""
        dxDBMerca.Columns.ColumnByFieldName("T1NOMPRO").Value = ""
        dxDBMerca.Dataset.Post
    Else
        dxDBMerca.Dataset.Edit
        dxDBMerca.Columns.ColumnByFieldName("T1NOMPRO").Value = ObtenerCampo("IF5PLA", "F5NOMPRO", "F5CODPRO", Me.dxDBMerca.Columns.ColumnByFieldName("T1CODPRO").Value & "", "T", cnn_dbbancos)
        dxDBMerca.Dataset.Post
    End If
End If
If Me.dxDBMerca.Columns.FocusedIndex = 4 Or Me.dxDBMerca.Columns.FocusedIndex = 3 Then
    With Me.dxDBMerca
        nfobunit = .Columns.ColumnByFieldName("t1fobuni").Value
        ncant = .Columns.ColumnByFieldName("t1cantidad").Value
        mcostounit = nFlete + (nfobunit * ncant)
        nmargen = .Columns.ColumnByFieldName("t1margen").Value
        nventaunit = mcostounit * 100 / 80
        
        .Dataset.Edit
        .Columns.ColumnByFieldName("t1fobtot").Value = .Columns.ColumnByFieldName("t1cantidad").Value * .Columns.ColumnByFieldName("t1fobuni").Value
        
        .Dataset.Post
        
        'dxDBImport.Dataset.Edit
        For I = 1 To dxDBImport.Dataset.RecordCount
            dxDBImport.Dataset.RecNo = I
            If dxDBImport.Columns.ColumnByName("DBCALCULAR").Value = True Then
                NSAVECOSTO = Val(dxDBImport.Columns.ColumnByName("DBCOSTO").Value & "")
                NSAVEIGV = Val(dxDBImport.Columns.ColumnByName("DBIGV").Value & "")
                NSAVETOTAL = Val(dxDBImport.Columns.ColumnByName("DBTOTAL").Value & "")
                dxDBImport.Dataset.Edit
                dxDBImport.Columns.ColumnByName("DBCALCULAR").Value = False
                dxDBImport.Dataset.Post
                Call dxDBImport_OnCheckEditToggleClick(dxDBImport.Columns.ColumnByName("DBCALCULAR"), Node, "", cbsChecked)
                If dxDBImport.Columns.ColumnByName("DBTOTAL").Value = 0 Then
                    dxDBImport.Dataset.Edit
                    dxDBImport.Columns.ColumnByName("DBCOSTO").Value = IIf(NSAVECOSTO > 0, NSAVECOSTO, "")
                    dxDBImport.Columns.ColumnByName("DBIGV").Value = IIf(NSAVEIGV > 0, NSAVEIGV, "")
                    dxDBImport.Columns.ColumnByName("DBTOTAL").Value = IIf(NSAVETOTAL > 0, NSAVETOTAL, "")
                    dxDBImport.Dataset.Post
                End If
            End If
        Next
        Call CmdInc_Click
        dxDBImport.Dataset.RecNo = 1
        dxDBImport.Dataset.Refresh
        'dxDBImport.Dataset.Post
        For I = 1 To dxDBGastos.Dataset.RecordCount
            dxDBGastos.Dataset.RecNo = I
            If dxDBGastos.Columns.ColumnByName("DBCALCULAR").Value = True Then
                dxDBGastos.Dataset.Edit
                dxDBGastos.Columns.ColumnByName("DBCALCULAR").Value = False
                dxDBGastos.Dataset.Post
                Call dxDBGastos_OnCheckEditToggleClick(dxDBGastos.Columns.ColumnByName("DBCALCULAR"), Node, "", cbsChecked)
            End If
        Next
        dxDBGastos.Dataset.RecNo = 1
    End With

End If
End Sub

Private Sub dxDBMerca_OnHeaderButtonClick()
MsgBox "x"
End Sub

Private Sub dxDBMerca_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item1 = False Then
        If Action = daInsert Then
            With dxDBMerca.Columns
                .ColumnByFieldName("T1CANTIDAD").Value = Format(0)
                .ColumnByFieldName("T1FOBUNI").Value = Format(0, "0.0000")
                .ColumnByFieldName("T1FOBTOT").Value = Format(0, "0.0000")
                .ColumnByFieldName("T1COSUNI").Value = Format(0, "0.0000")
                .ColumnByFieldName("T1VTAUNI").Value = Format(0, "0.0000")
                .ColumnByFieldName("T1MARGEN").Value = Format(0, "0.0000")
                .FocusedIndex = 0
            End With
        End If
    End If
End Sub


Private Sub dxDBMerca_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item1 = False Then
        If Action = daInsert Then
            If dxDBMerca.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBMerca.Columns.ColumnByFieldName("t1codpro").Value & "")) = 0 And Len(Trim(dxDBProductos.Columns.ColumnByFieldName("F3CODPRO").Value & "")) = 0 Or chkcerrar.Value Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBMerca.Dataset.Refresh
        End If
    End If
End Sub

Private Sub dxDBMerca_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
Case 46
    If MsgBox("¿Desea eliminar este producto?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
        sql = "delete  * from TMP_IMP_DET1 where t1codpro='" & Me.dxDBMerca.Columns.ColumnByFieldName("t1item").Value & "'"
        tempo.Execute sql
        'AlmacenaQuery_sql sql, tempo
        
        dxDBMerca.Dataset.Delete
        dxDBMerca.Dataset.Refresh
'        SQL = "select * from TMP_IMP_DET1 order by t1item"
'        If rs.State = 1 Then rs.Close
'        rs.Open SQL, tempo, 3, 1
'        If rs.RecordCount > 0 Then
'            I = 1
'            Do While Not rs.EOF
'                DoEvents
'                SQL = "update TMP_IMP_DET1 set t1item=" & I & " where t1item=" & Val(rs!T1ITEM & "")
'                tempo.Execute SQL
'                rs.MoveNext: I = I + 1
'            Loop
'        End If
        
    End If
Case 13 And Shift = 1
    '*****************
    If tempo.State = 0 Then tempo.Open
    sql = "SELECT TOP 1 T1ITEM FROM TMP_IMP_DET1 ORDER BY T1ITEM DESC"
    If Rs.State = 1 Then Rs.Close
    Rs.Open sql, tempo, 3, 1
    If Rs.RecordCount > 0 Then
        nitem = Val(Rs!t1item & "") + 1
    Else
        nitem = 1
    End If
    sql = "insert into TMP_IMP_DET1(t1numimp,T1ITEM) values ('" & Me.txtnumero.Text & "'," & nitem & ")"
    tempo.Execute sql
    'AlmacenaQuery_sql sql, tempo
    GridMercancias
    Me.dxDBMerca.Columns.FocusedIndex = 1
    Me.dxDBMerca.Dataset.Last
End Select
End Sub

Private Sub dxDBProductos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Select Case dxDBProductos.Columns.FocusedColumn.FieldName
    Case "f2codprov"
        wtipprov = "E"
        ayuda_proveedores.Show vbModal
        wtipprov = ""
        If Trim(wcodprov) <> "" Then
            dxDBProductos.Dataset.Edit
            dxDBProductos.Columns.ColumnByFieldName("f2codprov").Value = wrucprov
            dxDBProductos.Columns.ColumnByFieldName("f2nomprov").Value = wnomprov
        End If
    Case "F5CODPRO"
'        hlp_productos.Show vbModal
        ayuda_productos.Show vbModal
        If Trim(wcodproducto) <> "" Then
            dxDBProductos.Dataset.Edit
            dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
            dxDBProductos.Columns.ColumnByFieldName("F3CODFAB").Value = wcodfab
            dxDBProductos.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
            dxDBProductos.Columns.ColumnByFieldName("F5PARTARA").Value = wpartar
            dxDBProductos.Columns.ColumnByFieldName("F5UNIMED").Value = wmedida
'            dxDBProductos.Columns.FocusedIndex = 6
        End If
End Select
End Sub

Private Sub dxDBProductos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Calcula_Costos
If dxDBProductos.Columns.FocusedColumn.FieldName = "F5CODPRO" Then
If Len(Trim(dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").Value)) > 0 Then
   wf5codpro = dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").Value
   If rst.State = adStateOpen Then rst.Close
   sql = "select f5codfab,f5nompro,f7codmed,f5partara,f5valvta,f5factor,f5codpro,f5marca from if5pla where f5codpro='" & wf5codpro & "'"
   rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
   If Not rst.EOF Then
       dxDBProductos.Dataset.Edit
       dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").Value = "" & rst("f5codpro")
       dxDBProductos.Columns.ColumnByFieldName("F3CODFAB").Value = "" & rst("f5codFAB")
       dxDBProductos.Columns.ColumnByFieldName("F5NOMPRO").Value = "" & rst("f5nompro")
       dxDBProductos.Columns.ColumnByFieldName("F5PARTARA").Value = "" & rst("f5partara")
       dxDBProductos.Columns.ColumnByFieldName("F5UNIMED").Value = "" & rst("f7codmed")
       wmarca = "" & rst("F5MARCA")
       If rsttemp.State = adStateOpen Then rsttemp.Close
       sql = "select f2desmar from ef2marcas where f2codmar='" & wmarca & "'"
       rsttemp.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
       If Not rsttemp.EOF Then
           dxDBProductos.Columns.ColumnByFieldName("F5CODMARCA").Value = wmarca
           dxDBProductos.Columns.ColumnByFieldName("F5MARCA").Value = "" & rsttemp("f2desmar")
       End If
       rsttemp.Close
       dxDBProductos.Dataset.Post
   Else
       MsgBox "El Producto no Existe", vbInformation, "Sistema de Logística"
       dxDBProductos.Dataset.Edit
       dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").Value = ""
       dxDBProductos.Columns.ColumnByFieldName("F3CODFAB").Value = ""
       dxDBProductos.Columns.ColumnByFieldName("F5NOMPRO").Value = ""
       dxDBProductos.Columns.ColumnByFieldName("F5PARTARA").Value = ""
       dxDBProductos.Columns.ColumnByFieldName("F5UNIMED").Value = ""
       dxDBProductos.Columns.ColumnByFieldName("F3PREFOB").Value = "0.0000"
       dxDBProductos.Columns.ColumnByFieldName("F3TOTAL").Value = "0.0000"
       dxDBProductos.Columns.ColumnByFieldName("F3CANTIDAD").Value = "0.00"
       dxDBProductos.Dataset.Post
   End If
   rst.Close
End If
ElseIf dxDBProductos.Columns.FocusedColumn.FieldName = "f2codprov" Then
    wf2codprov = dxDBProductos.Columns.ColumnByFieldName("F2CODPROV").Value
    If Trim(wf2codprov) <> "" Then
        If rst.State = adStateOpen Then rst.Close
        sql = "select f2codprov,f2nomprov,f2newruc from ef2proveedores where f2newruc='" & wf2codprov & "'"
        rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not rst.EOF Then
            dxDBProductos.Dataset.Edit
            dxDBProductos.Columns.ColumnByFieldName("f2codprov").Value = "" & rst("f2newruc")
            dxDBProductos.Columns.ColumnByFieldName("f2nomprov").Value = "" & rst("f2nomprov")
            dxDBProductos.Dataset.Post
        Else
            MsgBox "El Proovedor no Existe", vbInformation, "Sistema de Logística"
            dxDBProductos.Dataset.Edit
            dxDBProductos.Columns.ColumnByFieldName("f2codprov").Value = ""
            dxDBProductos.Columns.ColumnByFieldName("f2nomprov").Value = ""
            dxDBProductos.Dataset.Post
        End If
        rst.Close
    End If
End If

dxDBProductos.Dataset.Edit
Calcula_New_Importaciones dxDBProductos.Columns.FocusedIndex
dxDBProductos.Dataset.Post
dxDBProductos.Dataset.Refresh
'Calcula_Costos
End Sub


Private Sub dxDBProductos_OnEditing(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, Allow As Boolean)
If dxDBProductos.Columns.FocusedColumn.FieldName = "F5CODPRO" And dxDBProductos.Columns.ColumnByFieldName("F5MANUAL").Value = "N" Then
    dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").DisableEditor = True
Else
    dxDBProductos.Columns.ColumnByFieldName("F5CODPRO").DisableEditor = False
End If
End Sub


Private Sub dxDBProductos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 115 Then
    If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
        sw_nuevo_item = True
        If dxDBProductos.Dataset.RecNo = 1 Then
            dxDBProductos.Dataset.Delete
            AdicionaItem
        Else
            dxDBProductos.Dataset.Delete
        End If
        sw_nuevo_item = False
    End If
End If
End Sub

Private Sub CargaUM()
If Rs.State = 1 Then Rs.Close
sql = "SELECT * FROM EF7MEDIDAS WHERE F7IMPORT=TRUE ORDER BY F7NOMMED"
Rs.Open sql, cnn_dbbancos, 3, 1
CboUm.Clear
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        CboUm.AddItem Rs!f7codmed
        Rs.MoveNext
    Loop
    CboUm.ListIndex = 0
End If
End Sub

Private Sub CargaZonas()
If Rs.State = 1 Then Rs.Close
sql = "SELECT * FROM import_zonas ORDER BY id_zona"
Rs.Open sql, cnn_dbbancos, 3, 1
CboZona.Clear
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        CboZona.AddItem Format(Rs!ID_ZONA, "00") & Space(3) & Rs!DETALLE
        Rs.MoveNext
    Loop
    CboZona.ListIndex = 0
End If
End Sub
Private Sub CargaImportadores()
If Rs.State = 1 Then Rs.Close
sql = "SELECT F2CODPROV,F2NOMPROV, F2NEWRUC, F2NOMABREV "
sql = sql & "From EF2PROVEEDORES "
sql = sql & "Where TIPPROV LIKE '%EMB%' ORDER BY F2NOMABREV"
Rs.Open sql, cnn_dbbancos, 3, 1
CboImporta.Clear
CboImporta.AddItem left("Seleccione el Agente" & Space(100), 100) & Format("0000", "0000")
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        CboImporta.AddItem left(Rs!F2NOMPROV & Space(100), 100) & Format(Rs!F2CODPROV, "0000")
        Rs.MoveNext
    Loop
    CboImporta.ListIndex = 0
End If

End Sub

Private Sub CargaAgentes()
If Rs.State = 1 Then Rs.Close
sql = "SELECT F2CODPROV,F2NOMPROV, F2NEWRUC, F2NOMABREV "
sql = sql & "From EF2PROVEEDORES "
sql = sql & "Where TIPPROV LIKE '%AGT%' ORDER BY F2NOMABREV"
Rs.Open sql, cnn_dbbancos, 3, 1
CboAgente.Clear
CboAgente.AddItem left("Seleccione el Agente" & Space(100), 100) & Format("0000", "0000")
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        CboAgente.AddItem left(Rs!F2NOMPROV & Space(100), 100) & Format(Rs!F2CODPROV, "0000")
        Rs.MoveNext
    Loop
    CboAgente.ListIndex = 0
End If
End Sub
Private Sub dxDBProductos_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Cant." Or Column.Caption = "Fob. Unit" Or Column.Caption = "Fob. Total" Then
    Text = Format(Text, "#,#0.00")
End If
End Sub

Private Sub dxDBProductos_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If Column.Caption = "Fob. Total" Then
    Text = Format(Text, "#,#0.00")
End If
End Sub








Private Sub FIAdela_GotFocus()
FIAdela.SelStart = 0: FIAdela.SelLength = Len(FIAdela.Text)
End Sub

Private Sub FIAdela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub FIDesaduana_GotFocus()
FIDesaduana.SelStart = 0: FIDesaduana.SelLength = Len(FIDesaduana.Text)
End Sub

Private Sub FIDesaduana_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub FIFob_GotFocus()
FIFob.SelStart = 0: FIFob.SelLength = Len(FIFob.Text)
End Sub

Private Sub FIFob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub FMAdela_GotFocus()
FMAdela.SelStart = 0: FMAdela.SelLength = Len(FMAdela.Text)
End Sub

Private Sub FMAdela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub FmDesaduana_GotFocus()
FmDesaduana.SelStart = 0: FmDesaduana.SelLength = Len(FmDesaduana.Text)
End Sub

Private Sub FmDesaduana_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub FMFob_GotFocus()
FMFob.SelStart = 0: FMFob.SelLength = Len(FMFob.Text)
End Sub

Private Sub FMFob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
    BASE_TEMPORAL

 '   Me.left = 1550
  '  Me.top = 700
    CargaVias
    CargaUM
    CargaZonas
    CargaImportadores
    CargaAgentes
    sw_ayuda_oc = False
    sw_nuevo_doc1 = True
    sw_detalle1 = False
    sw_cabecera1 = False
    sw_nuevo_item1 = False
    If sw_nuevo_documento = True Then
    
        sql = "select top 1 f4numimp from IMPORT_CAB order by f4numimp desc"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, 3, 1
        If Rs.RecordCount > 0 Then
            txtnumero.Text = Format(Val(Rs!f4numimp & "") + 1, "0000000")
            GOC = Format(Val(Rs!f4numimp & "") + 1, "0000000")
        Else
            txtnumero.Text = Format(1, "0000000")
            GOC = Format(1, "0000000")
        End If
        Call PROCEDIMIENTO_NUEVO
        
        Me.Toolbar.Buttons.ITEM(2).Visible = False
        Me.Toolbar.Buttons.ITEM(3).Visible = True
    Else
        If Val(GOC & "") > 0 Then
            txtnumero.Text = GOC
            Call txtnumero_KeyPress(13)
        Else
            sql = "Select f4numimp from Import_cab"
            If TbCabImport1.State = adStateOpen Then TbCabImport1.Close
            TbCabImport1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            
            If TbCabImport1.EOF Then
                WNUMERO1 = 1
            Else
                TbCabImport1.MoveLast
                WNUMERO1 = Val("" & TbCabImport1.Fields("F4NUMIMP")) + 1
            End If
        End If
        'txtnumero.Text = Format(WNUMERO1, "0000000")
        sw_nuevo_doc = True
        sw_detalle = False
        sw_nuevo_item = False
        sw_ayuda = False
        '**************
        aboFecha.Value = Format(Date, "dd/mm/yyyy")
    
        '**************
        'CargaImport
        'GridImport
        'CargaMercancias
        'GridMercancias
    
        llena_items
        
        
         
        Set rst = New ADODB.Recordset
        
        If Val(wnumimp) > 0 Then
            inicio = True
            txtnumero.Text = wnumimp
            txtnumero_KeyPress vbKeyReturn
            inicio = False
        End If
        End If

    
End Sub
Private Sub GridImport()
If ChkShow.Value = 1 Then
    wfiltro = " where val(t2total & '')>0"
Else
    wfiltro = ""
End If
    If tempo.State = 1 Then tempo.Close
    With Me.dxDBImport
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = tempo
        .Dataset.ADODataset.CommandText = "select *,t2NRO1,t2descripcion from TMP_IMP_DET2 " & wfiltro & " order by t2NRO1,t2NRO2"
        .Dataset.Active = True
        .KeyField = "T2KEY"
        .m.FullExpand
    End With
End Sub

Private Sub GridGastos()
If ChkShow.Value = 1 Then
    wfiltro = " where val(t3total & '')>0"
Else
    wfiltro = ""
End If
    If tempo.State = 1 Then tempo.Close
    With Me.dxDBGastos
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = tempo
        .Dataset.ADODataset.CommandText = "select *,t3NRO1,t3descripcion from TMP_IMP_DET3 " & wfiltro & " order by t3NRO1,t3NRO2"
        .Dataset.Active = True
        .KeyField = "T3KEY"
        .m.FullExpand
    End With
End Sub

Private Sub GridMercancias()
    SwCalc = True
    If tempo.State = 1 Then tempo.Close
    tempo.Open
    With Me.dxDBMerca
        .Dataset.ADODataset.ConnectionString = tempo
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = "select * from TMP_IMP_DET1 order by t1item"
        .Dataset.Active = True
        .KeyField = "t1item"
        
    End With
    SwCalc = False
End Sub
Private Sub CargaMercancias()
sql = "SELECT IMPORT_DET1.F4NumImp, IMPORT_DET1.F3NumOrd, IMPORT_DET1.F5CodPro, "
sql = sql & "IF5PLA.F5MODELO, IMPORT_DET1.F3MERCA , EF2MARCAS.F2DESMAR, IMPORT_DET1.F3Cantidad, "
sql = sql & "IMPORT_DET1.F3FobUni,IMPORT_DET1.F3FobTot, IMPORT_DET1.F3CosUni, IMPORT_DET1.F3VtaUni, "
sql = sql & "IMPORT_DET1.F3Margen,IMPORT_DET1.f3item "
sql = sql & "FROM (IF5PLA INNER JOIN IMPORT_DET1 ON IF5PLA.F5CODPRO = IMPORT_DET1.F5CodPro) "
sql = sql & "INNER JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR "
sql = sql & "WHERE (((IMPORT_DET1.F4NumImp)='" & Trim(txtnumero.Text) & "'))"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
If tempo.State = 0 Then tempo.Open
sql = "delete * from TMP_IMP_DET1"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo

If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
        VALORES = ""
        For I = 0 To (Rs.Fields.Count - 1)
            Select Case I
            Case Is >= 6: VALORES = VALORES & "," & Rs.Fields(I) & ""
            Case Else:
                If I = 0 Then
                    VALORES = VALORES & "'" & Rs.Fields(I) & "'"
                Else
                    VALORES = VALORES & ",'" & Rs.Fields(I) & "'"
                End If
            End Select
        Next
        sql = "insert into TMP_IMP_DET1 values (" & VALORES & ")"
        tempo.Execute sql
        'AlmacenaQuery_sql sql, tempo
        Rs.MoveNext
    Loop
Else

End If
End Sub
Private Sub CargaVias()
CboVia.Clear
CboVia.AddItem "Seleccione Via"
CboVia.AddItem "Aereo"
CboVia.AddItem "Maritimo"
CboVia.ListIndex = 0
End Sub

Private Sub CargaImport()

sql = "SELECT ([IMP_PLANT_DET].[GRUPO]+[IMP_PLANT_DET].[SUBGRUPO]) AS codigo, "
sql = sql & "IMP_PLANT_CAB.NRO1, IMP_PLANT_DET.GRUPO, IMP_PLANT_DET.NRO2, "
sql = sql & "IMP_PLANT_DET.SUBGRUPO, IMP_PLANT_DET.CAL_INC, IMP_PLANT_DET.ESP, "
sql = sql & "IMP_PLANT_DET.VALDEF, 0 AS F4CALCULAR, IMP_PLANT_CAB.DESCRIPCION, "
sql = sql & "IMP_PLANT_DET.DETALLE, IMP_PLANT_DET.TIPO, '' AS F4COSTO, '' AS F4IGV, '' AS F4TOTAL, "
sql = sql & "'' AS F4INCIDEN FROM IMP_PLANT_CAB INNER JOIN (IMP_EMBxPLANT INNER "
sql = sql & "JOIN IMP_PLANT_DET ON (IMP_EMBxPLANT.F1SUBGRUPO = IMP_PLANT_DET.SUBGRUPO) "
sql = sql & "AND (IMP_EMBxPLANT.F1GRUPO = IMP_PLANT_DET.GRUPO)) "
sql = sql & "ON IMP_PLANT_CAB.GRUPO = IMP_PLANT_DET.GRUPO "
sql = sql & "WHERE (((IMP_EMBxPLANT.F1EMBARCA)='" & right(CboImporta.Text, 4) & "') "
sql = sql & "AND (([IMP_PLANT_DET].[GRUPO]+[IMP_PLANT_DET].[SUBGRUPO]) Not In "
sql = sql & "(SELECT F4GRUPO+F4SUBGRUPO FROM  IMPORT_DET2 WHERE "
sql = sql & "IMPORT_DET2.F4NumImp='" & txtnumero.Text & "'))) "
sql = sql & "ORDER BY IMP_PLANT_CAB.NRO1, IMP_PLANT_DET.NRO2 "
sql = sql & "Union All "
sql = sql & "SELECT IMPORT_DET2.F4Grupo+IMPORT_DET2.F4SubGrupo AS codigo, "
sql = sql & "IMP_PLANT_CAB.NRO1, IMPORT_DET2.F4Grupo, IMP_PLANT_DET.NRO2, "
sql = sql & "IMPORT_DET2.F4SubGrupo, IMP_PLANT_DET.CAL_INC, IMPORT_DET2.F4Esp, "
sql = sql & "IMP_PLANT_DET.VALDEF, IMPORT_DET2.F4Calcular, IMP_PLANT_CAB.DESCRIPCION, "
sql = sql & "IMP_PLANT_DET.DETALLE, IMP_PLANT_DET.TIPO, IMPORT_DET2.F4Costo, "
sql = sql & "IMPORT_DET2.F4Igv, IMPORT_DET2.F4Total, IMPORT_DET2.F4Inciden "
sql = sql & "FROM IMP_PLANT_CAB INNER JOIN (IMP_PLANT_DET INNER JOIN IMPORT_DET2 "
sql = sql & "ON (IMP_PLANT_DET.GRUPO = IMPORT_DET2.F4Grupo) "
sql = sql & "AND (IMP_PLANT_DET.SUBGRUPO = IMPORT_DET2.F4SubGrupo)) "
sql = sql & "ON (IMP_PLANT_CAB.GRUPO = IMP_PLANT_DET.GRUPO) "
sql = sql & "AND (IMP_PLANT_CAB.GRUPO = IMP_PLANT_DET.GRUPO) "
sql = sql & "WHERE (((IMPORT_DET2.F4NumImp)='" & txtnumero.Text & "'));"

If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
If tempo.State = 0 Then tempo.Open

sql = "delete * from TMP_IMP_DET2"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
        amovs_cab(0).campo = "T2KEY": amovs_cab(0).valor = Rs!Codigo & "": amovs_cab(0).TIPO = "T"
        amovs_cab(1).campo = "T2Grupo": amovs_cab(1).valor = Rs!GRUPO & "": amovs_cab(1).TIPO = "T"
        amovs_cab(2).campo = "T2SubGrupo": amovs_cab(2).valor = Rs!subGRUPO & "": amovs_cab(2).TIPO = "T"
        amovs_cab(3).campo = "T2Descripcion": amovs_cab(3).valor = Rs!Descripcion & "": amovs_cab(3).TIPO = "T"
        amovs_cab(4).campo = "T2Detalle": amovs_cab(4).valor = Rs!DETALLE & "": amovs_cab(4).TIPO = "T"
        amovs_cab(5).campo = "T2Tipo": amovs_cab(5).valor = Rs!TIPO & "": amovs_cab(5).TIPO = "T"
        amovs_cab(6).campo = "T2Calcular": amovs_cab(6).valor = IIf((Rs!F4CALCULAR & "") = False, 0, -1): amovs_cab(6).TIPO = "N"
        amovs_cab(7).campo = "T2Costo": amovs_cab(7).valor = IIf(Len(Trim(Rs!F4Costo & "")) = 0, "NULL", Rs!F4Costo & ""): amovs_cab(7).TIPO = "N"
        amovs_cab(8).campo = "T2Igv": amovs_cab(8).valor = IIf(Len(Trim(Rs!F4IGV & "")) = 0, "NULL", Rs!F4IGV & ""): amovs_cab(8).TIPO = "N"
        amovs_cab(9).campo = "T2Total": amovs_cab(9).valor = IIf(Len(Trim(Rs!F4Total & "")) = 0, "NULL", Rs!F4Total & ""): amovs_cab(9).TIPO = "N"
        amovs_cab(10).campo = "T2Inciden": amovs_cab(10).valor = IIf(Len(Trim(Rs!F4Inciden & "")) = 0, "NULL", Rs!F4Inciden & ""): amovs_cab(10).TIPO = "N"
        amovs_cab(11).campo = "t2valdef": amovs_cab(11).valor = IIf(Len(Trim(Rs!valdef & "")) = 0, "NULL", Rs!valdef & ""): amovs_cab(11).TIPO = "N"
        amovs_cab(12).campo = "t2nro1": amovs_cab(12).valor = IIf(Len(Trim(Rs!NRO1 & "")) = 0, "NULL", Rs!NRO1 & ""): amovs_cab(12).TIPO = "N"
        amovs_cab(13).campo = "t2nro2": amovs_cab(13).valor = IIf(Len(Trim(Rs!nro2 & "")) = 0, "NULL", Rs!nro2 & ""): amovs_cab(13).TIPO = "N"
        amovs_cab(14).campo = "t2cal_inc": amovs_cab(14).valor = IIf((Rs!cal_inc & "") = False, 0, -1): amovs_cab(14).TIPO = "N"
        amovs_cab(15).campo = "t2eSP": amovs_cab(15).valor = IIf(Len(Trim(Rs!ESP & "")) = 0, "NULL", Rs!ESP & ""): amovs_cab(15).TIPO = "N"
        GRABA_REGISTRO amovs_cab(), "TMP_IMP_DET2", "A", 15, tempo, ""

        Rs.MoveNext
    Loop
Else

End If
End Sub

Private Sub CargaGastos()

sql = "SELECT ([IMP_PLANT_DET].[GRUPO]+[IMP_PLANT_DET].[SUBGRUPO]) AS codigo, "
sql = sql & "WHERE (((IMPORT_DET2.F4NumImp)='" & txtnumero.Text & "'));"
sql = ""
sql = "SELECT IMPORT_GASTOS_CAB.GRUPO+IMPORT_GASTOS_det.subGRUPO as indDex, "
sql = sql & "IMPORT_GASTOS_CAB.GRUPO, IMPORT_GASTOS_CAB.NRO1, IMPORT_GASTOS_CAB.DESCRIPCION, "
sql = sql & "IMPORT_GASTOS_DET.SUBGRUPO, IMPORT_GASTOS_DET.NRO2, IMPORT_GASTOS_DET.DETALLE, "
sql = sql & "IMPORT_DET3.F3Meses, IMPORT_DET3.F3Interes,IMPORT_DET3.F3CALCULAR, IMPORT_DET3.F3Costo, "
sql = sql & "IMPORT_DET3.F3Igv, IMPORT_DET3.F3Total, IMPORT_DET3.F3Inciden "
sql = sql & "FROM IMPORT_DET3 INNER JOIN (IMPORT_GASTOS_CAB INNER JOIN IMPORT_GASTOS_DET "
sql = sql & "ON IMPORT_GASTOS_CAB.GRUPO = IMPORT_GASTOS_DET.GRUPO) "
sql = sql & "ON (IMPORT_GASTOS_DET.SUBGRUPO = IMPORT_DET3.F3SubGrupo) "
sql = sql & "AND (IMPORT_DET3.F3Grupo = IMPORT_GASTOS_DET.GRUPO) "
sql = sql & "Where (((IMPORT_DET3.F3NumImp) = '" & txtnumero.Text & "')) "
sql = sql & "Union All "
sql = sql & "SELECT IMPORT_GASTOS_CAB.GRUPO+IMPORT_GASTOS_det.subGRUPO AS indDex, "
sql = sql & "IMPORT_GASTOS_CAB.GRUPO, IMPORT_GASTOS_CAB.NRO1, IMPORT_GASTOS_CAB.DESCRIPCION, "
sql = sql & "IMPORT_GASTOS_DET.SUBGRUPO, IMPORT_GASTOS_DET.NRO2, IMPORT_GASTOS_DET.DETALLE, "
sql = sql & "IMPORT_GASTOS_DET.MESES, IMPORT_GASTOS_DET.INTERES,IMPORT_GASTOS_DET.CALCULAR, "
sql = sql & "IMPORT_GASTOS_DET.COSTO, "
sql = sql & "IMPORT_GASTOS_DET.IGV, IMPORT_GASTOS_DET.TOTAL, 0 AS F3Inciden "
sql = sql & "FROM IMPORT_GASTOS_CAB INNER JOIN IMPORT_GASTOS_DET "
sql = sql & "ON IMPORT_GASTOS_CAB.GRUPO = IMPORT_GASTOS_DET.GRUPO "
sql = sql & "WHERE (((IMPORT_GASTOS_CAB.GRUPO+IMPORT_GASTOS_det.subGRUPO) "
sql = sql & "Not In (SELECT IMPORT_GASTOS_CAB.GRUPO+IMPORT_GASTOS_det.subGRUPO as indDex "
sql = sql & "FROM IMPORT_DET3 INNER JOIN (IMPORT_GASTOS_CAB INNER JOIN IMPORT_GASTOS_DET "
sql = sql & "ON IMPORT_GASTOS_CAB.GRUPO = IMPORT_GASTOS_DET.GRUPO) "
sql = sql & "ON (IMPORT_GASTOS_DET.SUBGRUPO = IMPORT_DET3.F3SubGrupo) "
sql = sql & "AND (IMPORT_DET3.F3Grupo = IMPORT_GASTOS_DET.GRUPO) "
sql = sql & "WHERE (((IMPORT_DET3.F3NumImp)='" & txtnumero.Text & "')) )))"

If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
If tempo.State = 0 Then tempo.Open

sql = "delete * from TMP_IMP_DET3"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
        amovs_cab(0).campo = "T3KEY": amovs_cab(0).valor = Rs!InDdex & "": amovs_cab(0).TIPO = "T"
        amovs_cab(1).campo = "T3Grupo": amovs_cab(1).valor = Rs!GRUPO & "": amovs_cab(1).TIPO = "T"
        amovs_cab(2).campo = "T3SubGrupo": amovs_cab(2).valor = Rs!subGRUPO & "": amovs_cab(2).TIPO = "T"
        amovs_cab(3).campo = "T3Descripcion": amovs_cab(3).valor = Rs!Descripcion & "": amovs_cab(3).TIPO = "T"
        amovs_cab(4).campo = "T3Detalle": amovs_cab(4).valor = Rs!DETALLE & "": amovs_cab(4).TIPO = "T"
        amovs_cab(5).campo = "T3Meses": amovs_cab(5).valor = IIf(Len(Trim(Rs!F3MESES & "")) = 0, "NULL", Rs!F3MESES & "") & "": amovs_cab(5).TIPO = "N"
        amovs_cab(6).campo = "T3Calcular": amovs_cab(6).valor = IIf((Rs!F3CALCULAR & "") = False, 0, -1): amovs_cab(6).TIPO = "N"
        amovs_cab(7).campo = "T3Costo": amovs_cab(7).valor = IIf(Len(Trim(Rs!F3Costo & "")) = 0, "NULL", Rs!F3Costo & ""): amovs_cab(7).TIPO = "N"
        amovs_cab(8).campo = "T3Igv": amovs_cab(8).valor = IIf(Len(Trim(Rs!F3IGV & "")) = 0, "NULL", Rs!F3IGV & ""): amovs_cab(8).TIPO = "N"
        amovs_cab(9).campo = "T3Total": amovs_cab(9).valor = IIf(Len(Trim(Rs!F3TOTAL & "")) = 0, "NULL", Rs!F3TOTAL & ""): amovs_cab(9).TIPO = "N"
        amovs_cab(10).campo = "T3Inciden": amovs_cab(10).valor = IIf(Len(Trim(Rs!F3Inciden & "")) = 0, "NULL", Rs!F3Inciden & ""): amovs_cab(10).TIPO = "N"
        amovs_cab(11).campo = "t3INTERES": amovs_cab(11).valor = IIf(Len(Trim(Rs!F3INTERES & "")) = 0, "NULL", Rs!F3INTERES & ""): amovs_cab(11).TIPO = "N"
        amovs_cab(12).campo = "t3nro1": amovs_cab(12).valor = IIf(Len(Trim(Rs!NRO1 & "")) = 0, "NULL", Rs!NRO1 & ""): amovs_cab(12).TIPO = "N"
        amovs_cab(13).campo = "t3nro2": amovs_cab(13).valor = IIf(Len(Trim(Rs!nro2 & "")) = 0, "NULL", Rs!nro2 & ""): amovs_cab(13).TIPO = "N"
        GRABA_REGISTRO amovs_cab(), "TMP_IMP_DET3", "A", 13, tempo, ""

        Rs.MoveNext
    Loop
Else

End If
End Sub

Public Sub BASE_TEMPORAL()
Set tempo = New ADODB.Connection

base_temp = "TMP_IMP.MDB"

CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "TMP_IMP.MDB; Persist Security Info=False"
tempo.Open CON

End Sub

Public Sub TABLA_TEMPORAL()
DBTable = "tmp_costos"
End Sub
Public Sub BASE_TEMPORAL1()
Set tempo = New ADODB.Connection

base_temp = "TMP_IMP.MDB"

CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & base_temp & "; Persist Security Info=False"
tempo.Open CON

End Sub

Public Sub TABLA_TEMPORAL1()
DBTable1 = "TmpDet_Import"
End Sub

Private Sub Conf_Grid1()
    
    With dxDBProductos.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
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
        '.Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        
    End With
  
    Call AdicionaItem1
    
End Sub

Private Sub Conf_Grid()
    
    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
'        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
'        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
'        .Set (egoTabThrough)
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
        '.Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        
    End With
  
    Call AdicionaItem
    
End Sub

Private Sub AdicionaItem()

Dim sw_nuevo_temp   As Boolean

dxDBGrid1.Dataset.Active = False

If sw_nuevo_doc = False Then

    DELETEREC_N DBTable, Temp
    dxDBGrid1.Dataset.Refresh
End If

dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
dxDBGrid1.Dataset.Active = True
dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open

With dxDBGrid1.Dataset

sw_nuevo_temp = False
sw_nuevo_item = True
For I = 1 To 1

    If sw_nuevo_temp = True Then
        If sw_nuevo_doc = True Then
            .Edit
        Else
            .Append
        End If
            sw_nuevo_temp = True
        Else
            .Append
    End If

    .FieldValues("F3ITEM") = I
    .FieldValues("F5CODPRO") = ""
    .FieldValues("F5NOMPRO") = ""
    .FieldValues("GRUPO") = ""
    .FieldValues("NOMBRE") = ""
'    .FieldValues("F3CHECK") = True
    .FieldValues("F3PRESUPUESTO") = Format(0, "###,##0.00")
    .FieldValues("F3SOLES") = Format(0, "###,##0.00")
    .FieldValues("F3DOLARES") = Format(0, "###,##0.00")
    
Next
    .Post
    sw_nuevo_item = False

End With

dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open

End Sub
Private Sub AdicionaItem1()

Dim sw_nuevo_temp1   As Boolean

dxDBProductos.Dataset.Active = False

If sw_nuevo_doc1 = False Then
    DELETEREC_N "TmpDet_Import", tempo
    dxDBProductos.Dataset.Refresh
End If

dxDBProductos.Dataset.ADODataset.ConnectionString = tempo
dxDBProductos.Dataset.Active = True
dxDBProductos.Dataset.Close
dxDBProductos.Dataset.Open

With dxDBProductos.Dataset

sw_nuevo_temp1 = False
sw_nuevo_item1 = True
For I = 1 To 1

    If sw_nuevo_temp1 = True Then
        If sw_nuevo_doc = True Then
            .Edit
        Else
            .Append
        End If
            sw_nuevo_temp = True
        Else
            .Append
    End If

    .FieldValues("F3ITEM1") = I
    .FieldValues("F3NUMORD") = ""
    '.FieldValues("F3DOCUM") = ""
    .FieldValues("F5CODPRO") = ""
    .FieldValues("F3CODFAB") = ""
    .FieldValues("F5NOMPRO") = ""
    .FieldValues("F5UNIMED") = ""
    .FieldValues("F3CANTIDAD") = Format(0, "0.000")
    .FieldValues("F3PREFOB") = Format(0, "0.0000")
    .FieldValues("f5advalorem") = Format(0, "0.0000")
    .FieldValues("advalorem") = Format(0, "0.0000")
    .FieldValues("base") = Format(0, "0.0000")
    .FieldValues("F3TOTAL") = Format(0, "0.0000")
    .FieldValues("F3PRECOS") = Format(0, "0.0000")
    .FieldValues("F3MARGEN") = Format(0, "0.0000")
    .FieldValues("F3VALVTA") = Format(0, "0.0000")
    .FieldValues("F3DSCTO") = Format(0, "0.0000")
    .FieldValues("F3VTANET") = Format(0, "0.0000")
    .FieldValues("F3PREUNI") = Format(0, "0.0000")
    .FieldValues("F3FLETE") = Format(0, "0.00")
    .FieldValues("CANTIDAD") = Format(0)
    .FieldValues("F5MANUAL") = "S"

Next

    dxDBProductos.Dataset.Post
    sw_nuevo_item1 = False

End With
dxDBProductos.Dataset.Close
dxDBProductos.Dataset.Open
End Sub


Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.ColumnByFieldName("F3PRESUPUESTO").Value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value = Format(0, "###,##0.00")
'            dxDBGrid1.Columns.ColumnByFieldName("F3CHECK").Value = True
            dxDBGrid1.Columns.FocusedIndex = 0
        End If
    End If
    
End Sub


Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                'If Len(Trim(dxDBGrid1.Columns(1).Value & "")) = 0 Then
                If Len("" & dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid1.Dataset.Refresh
        End If
    End If
    
End Sub




Private Sub Form_Unload(Cancel As Integer)

    If sw_ayuda_oc = True Then
        Unload New_Importaciones
    End If

End Sub

Private Sub fpTabProADO1_TabActivate(TabToActivate As Integer)
Dim I As Integer

    If TabToActivate = 0 Then
        If flag = False Then
            GRABACIONES
            flag = True
            
            If Val(New_New_Importaciones.dxDBProductos.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue) > 0 Then
                factor = dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").SummaryFooterValue / _
                (Val(Format(New_New_Importaciones.dxDBProductos.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.000")) _
                + Val(Format(New_New_Importaciones.dxDBProductos.Columns.ColumnByFieldName("F3FLETE").SummaryFooterValue, "0.000"))) + 1
                PnlFactor.Caption = Format(factor, "0.00000000000000000000")
            Else
                PnlFactor.Caption = "0.00000000000000000000"
            End If
            
            Calcula_Costos
            If dxDBProductos.Dataset.RecordCount > 0 Then
                For I = 1 To dxDBProductos.Dataset.RecordCount
                    dxDBProductos.Dataset.RecNo = I
                    dxDBProductos.Dataset.Edit
                    Calcula_New_Importaciones 7
                    dxDBProductos.Dataset.Post
                Next I
            End If
        End If
    Else
        If TabToActivate = 1 Then
            YY = 0
        Else
            If TabToActivate = 2 Then
                flag = False
            End If
        End If
    End If
    
End Sub

Private Sub fpTabProADO1_TabPageShown(ActiveTab As Integer, ActivePage As Integer)
If chkcerrar Then
    SSPanel1.Enabled = False
    dxDBGrid1.Enabled = False
Else
    dxDBGrid1.Enabled = True
End If
End Sub

Private Sub GIAdela_GotFocus()
GIAdela.SelStart = 0: GIAdela.SelLength = Len(GIAdela.Text)
End Sub

Private Sub GIAdela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub GIDesaduana_GotFocus()
GIDesaduana.SelStart = 0: GIDesaduana.SelLength = Len(GIDesaduana.Text)
End Sub

Private Sub GIDesaduana_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub GIFob_GotFocus()
GIFob.SelStart = 0: GIFob.SelLength = Len(GIFob.Text)
End Sub

Private Sub GIFob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub GMAdela_GotFocus()
GMAdela.SelStart = 0: GMAdela.SelLength = Len(GMAdela.Text)
End Sub

Private Sub GMAdela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub GMDesaduana_GotFocus()
GMDesaduana.SelStart = 0: GMDesaduana.SelLength = Len(GMDesaduana.Text)
End Sub

Private Sub GMDesaduana_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub GMFob_GotFocus()
GMFob.SelStart = 0: GMFob.SelLength = Len(GMFob.Text)
End Sub


Private Sub GMFob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub







Private Sub LblCostoI_Change()
If sw_nuevo_doc1 = False Then

End If
End Sub

Private Sub ActualizaIncidencias()
sql = "update TMP_IMP_DET2 SET C=0 WHERE T2Cal_Inc=-1"
End Sub

Private Sub LblCostoS_Change()
If sw_nuevo_doc1 = False Then
    If Val(dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue) > 0 Then
        Me.PnlFactorSinFin.Caption = Format(Val(Format(Me.dxDBImport.Columns.ColumnByFieldName("t2costo").SummaryFooterValue, "#0.00")) / Me.dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue, "#0.00000000000000000000")
        Me.PnlFactorConFin.Caption = Format(Val(Format(LblCostoI.Caption, "#0.00")) / Me.dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue, "#0.00000000000000000000")
    Else
        Me.PnlFactorSinFin.Caption = "0.00000000000000000000"
        Me.PnlFactorConFin.Caption = "0.00000000000000000000"
    End If
End If
End Sub






Private Sub PnlFactorConFin_Change()
If tempo.State = 0 Then tempo.Open
sql = "update TMP_IMP_DET1 set t1cosuni=t1fobuni*" & Val(PnlFactorConFin)
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
GridMercancias

End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Trim(Button.Caption)
    Case "idproforma"
        With rptproforma
            .datos.ConnectionString = cnn_dbbancos
            sql = "SELECT IMPORT_DET.*, IF5PLA.F5MARCA, EF2MARCAS.F2DESMAR, IF5PLA.F7CODMED, IF5PLA.F5TEXTO_ING, IMPORT_CAB.F4PROFORMA " _
            & "FROM IMPORT_CAB INNER JOIN (IMPORT_DET INNER JOIN (IF5PLA LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = " _
            & "EF2MARCAS.F2CODMAR) ON (IMPORT_DET.F5CODMARCA = IF5PLA.F5MARCA) AND (IMPORT_DET.F3CodFab = " _
            & "IF5PLA.F5CODFAB)) ON IMPORT_CAB.F4NumImp = IMPORT_DET.F4NumImp " _
            & " WHERE IMPORT_CAB.F4NUMIMP='" & txtnumero.Text & "' order by f2desmar"
            .Caption = "Proforma"
            .txtnum.Text = txtnumero.Text
            .txtembarcador.Text = TxtNomPrv.Text
            .txtdir.Text = TxtDirPrv.Text
            .txtfecha.Text = aboFecha.Value
            .datos.Source = sql
            .Show vbModal
        End With
    
    Case "Resumen"
        'Call costeo
        sql = "SELECT T3.T3Nro1, T3.T3Nro2, T3.T3Descripcion, T3.T3Detalle, "
        sql = sql & "T3.T3Meses, T3.T3Interes, T3.T3Costo, T3.T3Igv, T3.T3Total, "
        sql = sql & "T3.T3Inciden FROM TMP_IMP_DET3 AS T3 "
        sql = sql & "ORDER BY T3.T3Nro1, T3.T3Nro2"
        With Imp_Gastos
            m_sDBName = wrutatemp & "\TMP_IMP.MDB"
            .dcRptData.DatabaseName = m_sDBName
            .dcRptData.RecordSource = sql
            .Show 1
        End With
    
    Case "Nuevo"
        sw_nuevo_documento = True
        sql = "select top 1 f4numimp from IMPORT_CAB order by f4numimp desc"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, 3, 1
        If Rs.RecordCount > 0 Then
            txtnumero.Text = Format(Val(Rs!f4numimp & "") + 1, "0000000")
        Else
            txtnumero.Text = Format(1, "0000000")
        End If
        Call PROCEDIMIENTO_NUEVO
        
        Me.Toolbar.Buttons.ITEM(2).Visible = False
        Me.Toolbar.Buttons.ITEM(3).Visible = True
    Case "Grabar"
        If Trim(LblNomPrv.Caption) = "" Then
            MsgBox "Debe Ingresar Proveedor", vbInformation, "Sistema de Logística"
            TxtRucPrv.SetFocus
            Exit Sub
        End If
        If left(CboVia.Text, 1) = "S" Then
            MsgBox "Debe Seleccionar la Via de Importación", vbInformation, "Sistema de Logística"
            CboVia.SetFocus
            Exit Sub
        End If
        If left(CboImporta.Text, 4) = "0000" Then
            MsgBox "Debe Seleccionar un Embarcador", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
        If left(CboAgente.Text, 4) = "0000" Then
            MsgBox "Debe Seleccionar un Agente", vbInformation, "Sistema de Logística"
            CboAgente.SetFocus
            Exit Sub
        End If
        If left(Me.CboZona.Text, 2) = "0000" Then
            MsgBox "Debe Seleccionar la Zona de Origen", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
        If Len(Trim(Me.Txtcodcli.Text)) = 0 Then
            MsgBox "Debe Ingresar un Código de Cliente", vbInformation, "Sistema de Logística"
            Txtcodcli.SetFocus
            Exit Sub
        End If
        If Val(Me.dxDBMerca.Columns.ColumnByFieldName("t1fobuni").Value) = 0 Then
            MsgBox "El valor FOB Unitario tiene que ser mayor a cero.", vbInformation, "Sistema de Logística"
            dxDBMerca.SetFocus
            dxDBMerca.Columns.FocusedIndex = 1
            Exit Sub
        End If
        Me.MousePointer = 11
        dxDBMerca.Dataset.Edit
        If dxDBMerca.Dataset.State = dsEdit Or dxDBMerca.Dataset.State = dsInsert Then
             dxDBMerca.Dataset.Post
             sw_detalle1 = True
        End If
        If sw_cabecera1 = True Or sw_detalle1 = True Then
            
            If Me.dxDBMerca.Count = 0 Then
                Me.MousePointer = 1
                MsgBox "Debe Seleccionar Los Productos de la Importación", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                GRABAR2
                Me.MousePointer = 1
                MsgBox "La Importación Nº " & txtnumero.Text & " ha sido Actualizada.", vbInformation, "Sistema de Logística"
            End If
            
            sw_nuevo_doc1 = False
            gfalta = falta
            gcanti = wcantord
            sw_nuevo_doc1 = False
            sw_detalle1 = False
            sw_cabecera1 = False
             
        End If
        Me.Toolbar.Buttons.ITEM(2).Visible = True
        Me.MousePointer = 1
    Case "T.Comp."
        TxtNumOri.Text = txtnumero.Text
        txtnumero.Text = Format(Val(txtnumero.Text) + 1, "0000000")
        If Trim(TxtRucPrv.Text) = "" Then
            MsgBox "Debe Ingresar Embarcador", vbInformation, "Sistema de Logística"
            TxtRucPrv.SetFocus
            Exit Sub
        End If
        If left(CboVia.Text, 1) = "S" Then
            MsgBox "Debe Seleccionar la Via de Importación", vbInformation, "Sistema de Logística"
            CboVia.SetFocus
            Exit Sub
        End If
        If left(CboImporta.Text, 4) = "0000" Then
            MsgBox "Debe Seleccionar el Importador", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
        If left(Me.CboZona.Text, 2) = "0000" Then
            MsgBox "Debe Seleccionar el Importador", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
        If Len(Trim(Me.Txtcodcli.Text)) = 0 Then
            MsgBox "Debe Ingresar un Código de Cliente", vbInformation, "Sistema de Logística"
            Txtcodcli.SetFocus
            Exit Sub
        End If
        Me.MousePointer = 11
        dxDBMerca.Dataset.Edit
        If dxDBMerca.Dataset.State = dsEdit Or dxDBMerca.Dataset.State = dsInsert Then
             dxDBMerca.Dataset.Post
             sw_detalle1 = True
        End If
        If sw_cabecera1 = True Or sw_detalle1 = True Then
            GRABAR2
            If Me.dxDBMerca.Count = 0 Then
                Me.MousePointer = 1
                MsgBox "Debe Seleccionar Los Productos de la Importación", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                Me.MousePointer = 1
                MsgBox "La Importación Nº " & txtnumero.Text & " ha sido Actualizada.", vbInformation, "Sistema de Logística"
            End If
            
            sw_nuevo_doc1 = False
            gfalta = falta
            gcanti = wcantord
            sw_nuevo_doc1 = False
            sw_detalle1 = False
            sw_cabecera1 = False
             
        End If
        Me.Toolbar.Buttons.ITEM(2).Visible = True
        Me.MousePointer = 1
    Case "O.Comp."
            Me.MousePointer = 11
            sw_ayuda_oc = True
            wtipoc = "I"
            importar_ocompra.Show 1
            AGREGA_OCIMP

            GridMercancias
            TxtSerie.SetFocus
            Me.MousePointer = 1
    
    Case "Imprimir"
        Me.MousePointer = vbHourglass
        sql = "Select * from TmpDet_Import order by f5marca"
        
        

    sql = "SELECT T2KEY, T2Nro1, T2Grupo AS GRUPO, T2Nro2, T2SubGrupo, T2Descripcion AS Descripcion, "
    sql = sql & "T2Detalle, IIf(Len(Trim(T2TOTAL+''))>0 And Len(Trim(T2ESP+''))>0,"
    sql = sql & "T2Detalle+' ('+Format(T2ESP,'0.00')+'%)',T2Detalle) AS DETALLE, "
    sql = sql & "T2ESP, T2Cal_Inc, T2ValDef, T2Tipo, T2Calcular, T2Costo AS F4COSTO, "
    sql = sql & "T2Igv AS F4IGV, T2Total AS F4TOTAL, T2Inciden AS F4INCIDEN "
    sql = sql & "From TMP_IMP_DET2 "
    sql = sql & "ORDER BY TMP_IMP_DET2.T2Nro1, TMP_IMP_DET2.T2Nro2"
        With RegImporta2
            'superior
            .Caption = "COSTOS ESTIMADOS DE IMPORTACION"
            .DataControl.ConnectionString = tempo
            .DataControl.Source = sql
            .Label1.Caption = wnomcia
            .Label27.Caption = aboFecha.Value
            .LblProv.Caption = New_Importaciones.LblNomPrv.Caption
            .LblCli.Caption = New_Importaciones.LblNomCli.Caption
            .LblViaTip.Caption = UCase(New_Importaciones.CboZona.Text)
            .LblViaDes.Caption = UCase(New_Importaciones.TxtZonaDet.Text)
            .LblRef.Caption = New_Importaciones.TxtRefere.Text
            .LblEmb.Caption = UCase(New_Importaciones.CboVia.Text)
            .LblEmb.Alignment = ddTXLeft
            .LblFobTot.Caption = Format(New_Importaciones.dxDBMerca.Columns.ColumnByFieldName("t1fobTOT").SummaryFooterValue, "#,#0.00")
            If New_Importaciones.CboUm.Text = "KG" Then
                .LblUmPeso.Caption = New_Importaciones.CboUm.Text
                .LblPeso.Caption = Format(New_Importaciones.TxtCant.Text, "#,#0.00")
                .LblPeso.Alignment = ddTXLeft
            Else
                .LblUmVol.Caption = New_Importaciones.CboUm.Text
                .LblVol.Caption = Format(New_Importaciones.TxtCant.Text, "#,#0.00")
                .LblVol.Alignment = ddTXLeft
            End If
            'inferior
            .FldGasFin.Text = New_Importaciones.LblGastoFin.Caption
            .FldCostoI.Text = New_Importaciones.LblCostoI.Caption
            .FldIgvI.Text = New_Importaciones.LblIgvI.Caption
            .FldTotalI.Text = New_Importaciones.LblTotalI.Caption
            
            .FldCostoS.Text = New_Importaciones.LblCostoS.Caption
            .FldIgvS.Text = New_Importaciones.LblIgvS.Caption
            .FldTotalS.Text = New_Importaciones.LblTotalS.Caption
            .FldSinFin.Text = Format(New_Importaciones.PnlFactorSinFin.Caption, "#0.000000")
            .FldConFin.Text = Format(New_Importaciones.PnlFactorConFin.Caption, "#0.000000")
            .FldCostoU.Text = New_Importaciones.lblcosto.Caption
            .FldIgvU.Text = New_Importaciones.LblIgv.Caption
            .FldTotalU.Text = New_Importaciones.LblTot.Caption
            .FldUtil.Text = New_Importaciones.TxtUtil.Text
            .Refresh
            .Show vbModal
        End With
        Me.MousePointer = vbDefault
    
    Case "idtraduccion"
        Me.MousePointer = vbHourglass
        sql = "Select * from TmpDet_Import order by f5marca"
        RegImporta.DataControl1.ConnectionString = tempo
        RegImporta.DataControl1.Source = sql
        RegImporta.Caption = "Traducción"
        RegImporta.Label1.Caption = wempresa
        'RegImporta.txtproforma.Text = New_Importaciones.txtnumproforma.Text
        RegImporta.Label27.Caption = aboFecha.Value
        RegImporta.Label28.Caption = New_Importaciones.LblNomPrv.Caption
        RegImporta.Label29.Caption = New_Importaciones.LblDirPrv.Caption
        RegImporta.Label30.Caption = New_Importaciones.LblTelPrv.Caption
        RegImporta.Label31.Caption = New_Importaciones.TxtRefere.Text
        RegImporta.Show vbModal
        Me.MousePointer = vbDefault
        
    Case "Eliminar"
        Me.MousePointer = 11
        elimina txtnumero.Text
        'sw_nuevo_doc1 = True
        'factor = Format(0, "0.00000000000000000000")
        '
        'nuevo
        'dxDBProductos.Dataset.Close
        'DELETEREC_N "TmpDet_Import", tempo
        'AdicionaItem1
        Me.MousePointer = 1
        Unload Me
        
    
    Case "ID_Calculadora"
        Me.MousePointer = 11
        Dim X As Variant
        X = Shell("Calc.exe", 1)
        Me.MousePointer = 1
        
    Case "Salir"
        Me.MousePointer = 11
        If chkcerrar.Value = 0 Then
            If sw_cabecera1 = True Or sw_detalle1 = True Then
                If MsgBox("Desea Grabar el Movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                    GRABAR2
                    sw_nuevo_doc1 = False
                    sw_cabecera1 = False
                    sw_detalle1 = False
                End If
            End If
            '------------------ VERIFICA SI SE GRABO LA IMPORTACION
            Set rsimport_cab = New ADODB.Recordset
            If rsimport_cab.State = adStateOpen Then rsimport_cab.Close
            rsimport_cab.Open "SELECT F4NUMIMP FROM IMPORT_CAB WHERE F4NUMIMP = '" & txtnumero.Text & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
            If rsimport_cab.EOF Then
                
                csql = ("DELETE * FROM TB_COSTEOCAB WHERE F4NUMIMP='" & txtnumero.Text & "'")
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
                
                csql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP='" & txtnumero.Text & "'")
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
             
                
            End If
            rsimport_cab.Close
            Set rsimport_cab = Nothing
        End If
        Me.MousePointer = 1
        '-----------------------------------------------------------
        Unload Me
End Select
End Sub



Private Sub TxtAdelanto_GotFocus()
TxtAdelanto.SelStart = 0: TxtAdelanto.SelLength = Len(TxtAdelanto.Text)
End Sub

Private Sub TxtCant_GotFocus()
TxtCant.Text = Format(TxtCant.Text, "#0.00")
TxtCant.SelStart = 0: TxtCant.SelLength = Len(TxtCant.Text)
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13:    SendKeys "{Tab}"
End Select
End Sub

Private Sub TxtCant_LostFocus()
TxtCant.Text = Format(Val(TxtCant.Text), "#,#0.00")
End Sub

Private Sub Txtcodcli_DblClick()
Call TxtCodCli_KeyUp(113, 0)
End Sub

Private Sub Txtcodcli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCodCli_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    ayuda_clientes.Show 1
    Txtcodcli.Text = wruccli
    LblNomCli.Caption = wnomcliprov
End If
End Sub

Private Sub TxtCostoI_Change()
CalculaMontos
End Sub

Private Sub txtCostoS_Change()
CalculaMontos
End Sub

Private Sub TxtDesaduana_Change()
TxtDesaduana.SelStart = 0: TxtDesaduana.SelLength = Len(TxtDesaduana.Text)
End Sub

Private Sub TxtCom_GotFocus()
TxtCom.SelStart = 0: TxtCom.SelLength = Len(TxtCom.Text)
End Sub

Private Sub TxtCom_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    dxDBMerca.SetFocus
    dxDBMerca.Columns.FocusedIndex = 1
End If

End Sub

Private Sub TxtFob_GotFocus()
txtFOB.SelStart = 0: txtFOB.SelLength = Len(txtFOB.Text)
End Sub

Private Sub TxtDirPrv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub TxtGastoFin_GotFocus()
TxtGastoFin.SelStart = 0: TxtGastoFin.SelLength = Len(TxtGastoFin.Text)
End Sub

Private Sub TxtGastoFin_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    SendKeys "{tab}"
End Select
End Sub

Private Sub TxtGastoFin_LostFocus()
CalculaMontos
LblGastoFin.Caption = Format(LblGastoFin.Caption, "#,###,#0.00")
End Sub

Private Sub TxtIgvI_Change()
CalculaMontos
End Sub

Private Sub TxtIgvS_Change()
CalculaMontos
End Sub

Private Sub Txtnomcli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtNomPrv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub



Private Sub txtnumero_GotFocus()

    txtnumero.SelStart = 0: txtnumero.SelLength = Len(txtnumero.Text)

End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
    
    Set TbCabImport1 = New ADODB.Recordset
    
    If KeyAscii = 13 Then
        txtnumero.Text = Format(txtnumero.Text, "0000000")
        sql = "Select * from Import_Cab where F4NUMIMP='" & txtnumero.Text & "'"
        If TbCabImport1.State = adStateOpen Then TbCabImport1.Close
        TbCabImport1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not TbCabImport1.EOF Then
            actualiza
            'CargaImport
            'GridImport
            CargaMercancias
            CargaGastos
            GridGastos
            GridMercancias
            sw_nuevo_doc1 = False
        Else
            PROCEDIMIENTO_NUEVO
            sw_nuevo_doc1 = True
        End If
                
    SendKeys "{tab}"
    End If

End Sub

Private Sub txtnumero_LostFocus()
  If sw_ayuda = False Then
        If Len(Trim(txtnumero.Text)) > 0 Then
            llena_items
        End If
  End If
End Sub

Private Sub llena_items()
csql = "select * from import_cab where f4numimp='" & Me.txtnumero.Text & "'"
If rst.State = 1 Then rst.Close
rst.Open csql, cnn_dbbancos, 3, 1
If rst.RecordCount > 0 Then
    Me.aboFecha.Value = rst!F4FECHA & ""
    'datos del proveedor
    Me.TxtRucPrv.Text = ObtenerCampo("EF2PROVEEDORES", "f2newruc", "f2codprov", rst!F4CODPRV & "", "T", cnn_dbbancos)
    Me.LblNomPrv.Caption = ObtenerCampo("EF2PROVEEDORES", "f2nomprov", "f2codprov", rst!F4CODPRV & "", "T", cnn_dbbancos)
    Me.LblDirPrv.Caption = ObtenerCampo("EF2PROVEEDORES", "f2dirprov", "f2codprov", rst!F4CODPRV & "", "T", cnn_dbbancos)
    Me.LblTelPrv.Caption = ObtenerCampo("EF2PROVEEDORES", "f2telprov", "f2codprov", rst!F4CODPRV & "", "T", cnn_dbbancos)
    'datos del cliente
    Me.Txtcodcli.Text = ObtenerCampo("EF2CLIENTES", "f2newruc", "f2codcli", Format(rst!f4cliente, "0000") & "", "T", cnn_dbbancos)
    Me.LblNomCli.Caption = ObtenerCampo("EF2CLIENTES", "f2nomcli", "f2codcli", Format(rst!f4cliente, "0000") & "", "T", cnn_dbbancos)
    'importador
    Call SeleccionaEnComboRight(rst!f4importador & "", CboImporta)
    Call SeleccionaEnComboRight(rst!f4AGENTE & "", CboAgente)
    Call SeleccionaEnCombo(rst!f4undmed & "", CboUm)
    Call SeleccionaEnCombo(rst!f4via, CboVia)
    Call SeleccionaEnCombo(rst!f4ZONA, CboZona)
    Me.TxtCant.Text = Format(rst!f4medida & "", "#,#0.00")
    Me.TxtZonaDet.Text = rst!F4ZONADES & ""
    Me.TxtRefere.Text = rst!F4REFERE & ""
    'factura del flete
    Me.TxtSerie.Text = rst!f4serie & ""
    Me.TxtNumFac.Text = rst!f4numfac & ""
    'importacion totales
    'Me.TxtCostoI.Caption = Format(Val(rs!f4costoi & ""), "#,#0.00")
    'Me.TxtIgvI.Caption = Format(Val(rs!f4igvi & ""), "#,#0.00")
    'Me.TxtTotalI.Caption = Format(Val(rs!f4toti & ""), "#,#0.00")
    'totales sugeridos
    'Me.txtCostoS.Text = Format(Val(rs!f4costos & ""), "#,#0.00")
    'Me.TxtIgvS.Text = Format(Val(rs!f4igvs & ""), "#,#0.00")
    'Me.TxtTotalS.Text = Format(Val(rs!f4tots & ""), "#,#0.00")
    '*****
    Me.LblGastoFin.Caption = Format(Val(rst!f4gasfin & ""), "#,#0.00")
    Me.PnlFactorSinFin.Caption = Format(rst!F4FACTORSINFIN & "", "0.00000000000000000000")
    Me.PnlFactorConFin.Caption = Format(rst!F4FACTORCONFIN & "", "0.00000000000000000000")
    Me.TxtUtil.Text = Format(Val(rst!F4utilidad & ""), "#,#0.00")
    Me.TxtNumOri.Text = rst!F4numori & ""
    Me.TxtImportacion.Text = rst!f4importacion & ""
    
    

    
    
    If rst!f4cerrado & "" = "N" Then
        chkcerrar.Value = 0
    Else
        chkcerrar.Value = 1
        Toolbar.Buttons.ITEM(3).Visible = False
        Toolbar.Buttons.ITEM(4).Visible = False
    End If
    If rst!f4aprobado & "" = True Then
        ChkAprobar.Value = 1
        TxtImportacion.Visible = True
        LblImportacion.Visible = True
        txtnumero.Visible = False
        LblNumero.Visible = False
    Else
        ChkAprobar.Value = 0
        TxtImportacion.Visible = False
        LblImportacion.Visible = False
        txtnumero.Visible = True
        LblNumero.Visible = True
    End If
    CalculaMontos
Else
    nuevo
End If
End Sub
Private Sub txtOrdenCompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtOrdenDespacho.SetFocus
End If
End Sub


Private Sub txtnumproforma_GotFocus()
txtnumproforma.SelStart = 0: txtnumproforma.SelLength = Len(txtnumproforma.Text)
End Sub

Private Sub txtnumproforma_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtnumproforma.Text = Format(txtnumproforma.Text, "0000000")
    TxtRucPrv.SetFocus
Else
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End If
End Sub


Private Sub txtOrdenDespacho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    aboFechaProDespacho.SetFocus
End If
End Sub

Private Sub aboEmision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    abofechaconfirma.SetFocus
End If
End Sub

Private Sub aboFechaProDespacho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbtipoembarque.SetFocus
End If
End Sub

Private Sub cmbtipoembarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCompañiaTransporte.SetFocus
End If
End Sub

Private Sub txtCompañiaTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtProforma.SetFocus
End If
End Sub

Private Sub txtProforma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   aboFechaSalida.SetFocus
End If
End Sub

Private Sub aboFechaSalida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   aboFecProArrPuerto.SetFocus
End If
End Sub

Private Sub aboFecProArrPuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   aboFecLLegEmbarque.SetFocus
End If
End Sub

Private Sub aboFecLLegEmbarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   aboFecProgInspeccion.SetFocus
End If
End Sub

Private Sub aboFecProgInspeccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   aboFecInspeccion.SetFocus
End If
End Sub

Private Sub aboFecInspeccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtCertificado.SetFocus
End If
End Sub

Private Sub TxtNumOri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtRucPrv_Change()
    'TxtNomPrv.Text = ""
    If Trim(TxtRucPrv.Text) <> "" And sw_cabecera1 = False Then
        sw_cabecera1 = True
        cargar_datos
    End If
End Sub
Private Sub TxtRucPrv_DblClick()
    TxtRucPrv_KeyUp 113, 0
End Sub


Private Sub TxtRucPrv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        cargar_datos
    End If
    If KeyCode = 113 Then
        Me.MousePointer = 11
        wrucprov = "" & TxtRucPrv.Text
        sw_ocompra = True
        wtipprov = "E"
        ayuda_proveedores.Show 1
        wtipprov = ""
'        hlp_proveedores.Show 1
        TxtRucPrv.Text = wrucprov
        LblNomPrv.Caption = wnomprov
        LblDirPrv.Caption = wdirprov
        Me.MousePointer = 1
        TxtRucPrv_KeyPress 13
    End If

End Sub

Private Sub TxtRucPrv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Public Sub cargar_datos()
'If Len(Trim(TxtRucPrv.Text)) = 0 Then Exit Sub
Set Tbproveedor1 = New ADODB.Recordset

sql = "Select * from EF2PROVEEDORES where f2newruc='" & Trim(TxtRucPrv.Text) & "'"
If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
Tbproveedor1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not Tbproveedor1.EOF And Len(TxtRucPrv.Text) > 0 Then
'    TXTCODPRO.Text = "" & Tbproveedor1.Fields("F2CODPROV")
    LblNomPrv.Caption = "" & Tbproveedor1.Fields("F2NOMPROV")
    LblDirPrv.Caption = "" & Tbproveedor1.Fields("F2DIRPROV")
    TxtRucPrv.Text = "" & Tbproveedor1.Fields("f2newruc")
    LblTelPrv.Caption = "" & Tbproveedor1.Fields("F2TELPROV")
Else
    MsgBox "Codigo de ruc no existe.Ingrese un codigo de ruc", vbInformation, "Atencion"
    TxtNomPrv.Text = ""
    TxtDirPrv.Text = ""
    TxtRucPrv.Text = ""
    TxtTelPrv = ""
    TxtRucPrv.SetFocus
End If

End Sub

Private Sub actualiza()
    Dim Tbproveedor1    As New ADODB.Recordset

    sql = "Select F2NOMPROV,F2DIRPROV,F2CODPROV,F2NEWRUC,F2TELPROV from EF2PROVEEDORES where F2CODPROV='" & TbCabImport1.Fields("F4CodPrv") & "'"
    If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
    Tbproveedor1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not Tbproveedor1.EOF Then
        LblNomPrv.Caption = "" & Tbproveedor1.Fields("F2NOMPROV")
        LblDirPrv.Caption = "" & Tbproveedor1.Fields("F2DIRPROV")
'        TXTCODPRO.Text = "" & Tbproveedor1.Fields("F2CODPROV")
'        TxtRucPrv.Text = "" & Tbproveedor1.Fields("F2NEWRUC")
'        TxtTelPrv.Text = "" & Tbproveedor1.Fields("F2TELPROV")
    End If
    Tbproveedor1.Close
    Set Tbproveedor1 = Nothing
    
'    txtnumproforma.Text = "" & TbCabImport1.Fields("f4proforma")
    TxtSerie.Text = "" & TbCabImport1.Fields("F4SERIE")
    TxtNumFac.Text = "" & TbCabImport1.Fields("F4NUMFAC")
    aboFecha.Value = "" & Format(TbCabImport1.Fields("F4FECHA"), "DD/MM/YYYY")
    TxtRefere.Text = "" & TbCabImport1.Fields("F4REFERE")
    PnlFactorSinFin.Caption = "" & Format(TbCabImport1.Fields("F4FACTORSINFIN"), "###,##0.00000000000000000000")
    If "" & TbCabImport1.Fields("F4CERRADO") = "S" Then
        chkcerrar.Value = 1
        chkcerrar.Caption = "Importación Cerrada"
        activar True
    Else
        chkcerrar.Value = 0
        chkcerrar.Caption = "&Cerrar Importación"
    End If
    If "" & TbCabImport1.Fields("F4aprobado") = True Then
        ChkAprobar.Value = 1
        ChkAprobar.Enabled = False
    Else
        chkcerrar.Value = 0
        ChkAprobar.Enabled = True
    End If
    'Seguimiento
'    aboFecInspeccion.Value = "" & TbCabImport1("F4FECINSPE")
'    aboFecLLegEmbarque.Value = "" & TbCabImport1("F4FECLLEGADA")
'    aboFecProArrPuerto.Value = "" & TbCabImport1("F4FECPUERTO")
'    aboEmision.Value = "" & TbCabImport1("F4FECEMISION")
'    abofechaconfirma.Value = "" & TbCabImport1("F4FECEMBARCADOR")
'    aboFechaProDespacho.Value = "" & TbCabImport1("F4FECDESPACHO")
'    aboFechaSalida.Value = "" & TbCabImport1("F4FECSALIDA")
'    aboFecProgInspeccion.Value = "" & TbCabImport1("F4FECPROGINSPE")
    
'    txtOrdenDespacho.Text = "" & TbCabImport1("F4DESPACHO")
'    txtCompañiaTransporte.Text = "" & TbCabImport1("F4TRANSPORTE")
'    txtproforma.Text = "" & TbCabImport1("F4PROEMBARCA")
'    txtCertificado.Text = "" & TbCabImport1("F4CERTIFICADO")
    

    
    'LLena_DataGrid
    sw_cabecera1 = False: sw_detalle1 = False
End Sub
Private Sub CalculaMontos()
 
Me.MousePointer = 11
LblCostoI.Caption = Format(Me.dxDBImport.Columns.ColumnByFieldName("t2costo").SummaryFooterValue + Format(LblGastoFin.Caption, "#0.00"), "#,#0.00")
LblIgvI.Caption = Format(Val(Format(LblCostoI.Caption, "#0.00")) * wIgv / 100, "#,#0.00")
LblTotalI.Caption = Format(Val(Format(LblCostoI.Caption, "#0.00")) + Val(Format(LblIgvI.Caption, "#0.00")), "#,#0.00")
nutil = Format(TxtUtil.Text, "#0.00")
LblCostoS.Caption = Format((Val(Format(LblCostoI.Caption, "#0.00")) * nutil / 100) + Val(Format(LblCostoI.Caption, "#0.00")), "#,#0.00")
LblIgvS.Caption = Format(Val(Format(LblCostoS.Caption, "#0.00")) * wIgv / 100, "#,#0.00")
LblTotalS.Caption = Format(Val(Format(LblCostoS.Caption, "#0.00")) + Val(Format(LblIgvS.Caption, "#0.00")), "#,#0.00")
lblcosto.Caption = Format(Val(Format(LblCostoS.Caption, "#0.00")) - Val(Format(LblCostoI.Caption, "#0.00")), "#,#0.00")
LblIgv.Caption = Format(Val(Format(LblIgvS.Caption, "#0.00")) - Val(Format(LblIgvI.Caption, "#0.00")), "#,#0.00")
LblTot.Caption = Format(Val(Format(LblTotalS.Caption, "#0.00")) - Val(Format(LblTotalI.Caption, "#0.00")), "#,#0.00")
Me.MousePointer = 1

    If Val(dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue) > 0 Then
        Me.PnlFactorSinFin.Caption = Format(Val(Format(Me.dxDBImport.Columns.ColumnByFieldName("t2costo").SummaryFooterValue & "", "#0.00")) / Me.dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue, "#0.00000000000000000000")
        Me.PnlFactorConFin.Caption = Format(Val(Format(LblCostoI.Caption, "#0.00")) / Me.dxDBMerca.Columns.ColumnByFieldName("t1fobtot").SummaryFooterValue, "#0.00000000000000000000")
    Else
        Me.PnlFactorSinFin.Caption = "0.00000000000000000000"
        Me.PnlFactorConFin.Caption = "0.00000000000000000000"
    End If
End Sub

Private Sub LLena_DataGrid()
Dim X As Integer
Dim SQL1 As String
Dim unidad As String

    Set TbDetImport1 = New ADODB.Recordset
    Set tbProducto1 = New ADODB.Recordset

    CONT = 1
    dxDBProductos.Dataset.Close
    DELETEREC_N "TmpDet_Import", tempo
    SQL1 = "Select F5CODPRO,F5NomPro,f5modelo,f5advalorem,F5FACTOR,f5marca,f5codfab from IF5PLA"
    If tbProducto1.State = adStateOpen Then tbProducto1.Close
    tbProducto1.Open SQL1, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    sql = "Select * from Import_Det where F4NUMIMP='" & txtnumero.Text & "'"
    If TbDetImport1.State = adStateOpen Then TbDetImport1.Close
    TbDetImport1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    If Not TbDetImport1.EOF Then
        X = 1
        Do While TbDetImport1.Fields("F4NUMIMP") = txtnumero.Text
            numimporta = TbDetImport1.Fields("F4NumImp")
            numorden = TbDetImport1.Fields("F3NumOrd")
            codfabrica = TbDetImport1.Fields("F3CodFab")
            codproducto = TbDetImport1.Fields("F5CodPro")
            F5MARCA = "" & TbDetImport1.Fields("f5marca")
            F5CODMARCA = "" & TbDetImport1.Fields("f5codmarca")
            
            tbProducto1.Filter = "f5codpro='" & codproducto & "' and f5marca='" & F5CODMARCA & "'"
            If Not tbProducto1.EOF Then
                NOMPRODUCTO = tbProducto1.Fields("F5NomPro")
                ADVALOREM = Val("" & tbProducto1.Fields("f5advalorem"))
                'xf5factor = "" & tbProducto1.Fields("f5factor")
            End If
            tbProducto1.Filter = adFilterNone
            
            cantidad1 = Format(TbDetImport1.Fields("f3Cantidad"), "0.00")
            preciounit = Format(TbDetImport1.Fields("F3Preuni"), "0.0000")
            total1 = Format(TbDetImport1.Fields("F3Total"), "0.0000")
            f3flete = Val(Format("" & TbDetImport1.Fields("F3flete"), "0.00"))
            preccosto = Format(TbDetImport1.Fields("F3PreCos"), "0.0000")
            margen = Format(TbDetImport1.Fields("F3Margen"), "0.0000")
            ValVta = Format(TbDetImport1.Fields("F3ValVta"), "0.0000")
            descuento = Format(TbDetImport1.Fields("F3Dscto"), "0.0000")
            preuni = Format(Val(Format(TbDetImport1.Fields("F3ValVta"), "0.0000")) + Val(Format(TbDetImport1.Fields("F3ValVta"), "0.0000")) * 0.18, "0.0000")
            cantidad2 = Format(TbDetImport1.Fields("f3Cantidad"), "0.0000")
            vtaneta = Format(TbDetImport1.Fields("F3VTANET"), "0.0000")
            ADVALOREM1 = ADVALOREM
            unidad = TbDetImport1.Fields("F3UniMed") & ""
            F2CODPROV = "" & TbDetImport1.Fields("F2codprov")
            F2NOMPROV = "" & TbDetImport1.Fields("F2nomprov")
            f5partara = "" & TbDetImport1.Fields("F5partara")
            f5manual = "" & TbDetImport1.Fields("F5manual")
            f3costototal = Format(Val(Format(TbDetImport1.Fields("f3costototal"), "0.0000")))
            
            csql = "INSERT INTO " & "TmpDet_Import" & " (F4NUMIMP,F3NUMORD,F3CODFAB,F5CODPRO," & _
            "F5NOMPRO,F3CANTIDAD,F3PREFOB,F3TOTAL,F3PRECOS,F3MARGEN,F3VALVTA,F3DSCTO,F3PREUNI,CANTIDAD,F3VTANET,F5UNIMED,F3ITEM1,f2codprov,f2nomprov,f5partara,f5manual,f3costototal,f3flete,f5marca,f5codmarca) VALUES('" & numimporta & "','" & _
            numorden & "','" & codfabrica & "','" & codproducto & "','" & _
            NOMPRODUCTO & "'," & cantidad1 & "," & preciounit & "," & _
            total1 & "," & preccosto & "," & margen & "," & ValVta & "," & descuento & ", " & preuni & "," & cantidad2 & "," & vtaneta & ",'" & unidad & "'," & CONT & ",'" & F2CODPROV & "','" & F2NOMPROV & "','" & f5partara & "','" & f5manual & "'," & f3costototal & "," & f3flete & ",'" & F5MARCA & "','" & F5CODMARCA & "')"
            tempo.Execute (csql)
            'AlmacenaQuery_sql csql, tempo
            
            CONT = CONT + 1
            TbDetImport1.MoveNext
            If TbDetImport1.EOF Then
                Exit Do
            End If
        Loop
        dxDBProductos.Dataset.ADODataset.ConnectionString = tempo
        dxDBProductos.Dataset.Active = True
        dxDBProductos.Dataset.Open
    End If
    tbProducto1.Close
    dxDBProductos.Dataset.First
    dxDBProductos.Columns.FocusedIndex = 0
End Sub
Private Sub TxtNomPrv_Change()
If Trim(TxtNomPrv.Text) <> "" And sw_cabecera1 = False Then
    sw_cabecera1 = True
End If
End Sub

Private Sub TxtDirPrv_Change()
If Trim(TxtDirPrv.Text) <> "" And sw_cabecera1 = False Then
    sw_cabecera1 = True
End If
End Sub

Private Sub txtserie_LostFocus()
'    If txtserie.Text = "" Then
'        MsgBox ("Ingrese Nº de Serie y Nº de Factura."), vbInformation, "Atencion"
'    End If
End Sub

Private Sub TxtTelPrv_Change()
If Trim(TxtTelPrv.Text) <> "" And sw_cabecera1 = False Then
    sw_cabecera1 = True
End If
End Sub

Private Sub AGREGA_OCIMP()
    If importar_ocompra.dxDBGrid1.Count > 0 Then
        CODPROV = importar_ocompra.dxDBGrid1.Columns.ColumnByFieldName("F4CODPRV").Value
        BUSCA_OCOMPRA Trim(importar_ocompra.dxDBGrid1.Columns(0).Value)
        sw_Ord = True
    Else
        sw_Ord = False
    End If
End Sub

Private Sub BUSCA_OCOMPRA(pocompra As String)
Dim I      As Integer
    
If rsif4orden.State = adStateOpen Then rsif4orden.Close
rsif4orden.Open "SELECT * FROM IF4ORDEN WHERE F4NUMORD = '" & (pocompra) & "' and f4local='0'", cnn_dbbancos, 3, 1
If rsif4orden.RecordCount > 0 Then
    'SELECCIONA LA VIA DE IMPORTACION
    CboVia.ListIndex = Val(rsif4orden!F4VIATRANS & "")
    Call SeleccionaEnComboRight((rsif4orden!F4CODCLI & ""), CboImporta)
    
    Set Tbproveedor1 = New ADODB.Recordset
    sql = "Select * from EF2PROVEEDORES where f2newruc='" & Trim(CODPROV) & "'"
    If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
    Tbproveedor1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not Tbproveedor1.EOF And Len(CODPROV) > 0 Then
        LblNomPrv.Caption = "" & Tbproveedor1.Fields("F2NOMPROV")
        LblDirPrv.Caption = "" & Tbproveedor1.Fields("F2DIRPROV")
        TxtRucPrv.Text = "" & Tbproveedor1.Fields("f2newruc")
        LblTelPrv.Caption = "" & Tbproveedor1.Fields("F2TELPROV")
    End If
    If rsif3orden.State = adStateOpen Then rsif3orden.Close
    rsif3orden.Open "SELECT * FROM IF3ORDEN WHERE F4NUMORD = '" & (pocompra) & "'", cnn_dbbancos, 3, 1
    If rsif3orden.RecordCount > 0 Then
        sw_Orden = True
        If tempo.State = 1 Then tempo.Close
        tempo.Open
        csql = "DELETE * FROM TMP_IMP_DET1"
        tempo.Execute csql
        'AlmacenaQuery_sql csql, tempo
        
        n = 1
        Do While Not rsif3orden.EOF
            'sql = "insert into TMP_IMP_DET1"
            'sql = sql & "(T1NumImp,T1NumOrd,T1CodPro,T1Modelo,T1NomPro,T1DesMar,"
            'sql = sql & "T1Cantidad,T1FobUni,T1FobTot) "
            'sql = sql & "values ('" & txtnumero.Text & "','" & rsif3orden!F4NUMORD
            'sql = sql & "','" & rsif3orden!f3codpro
            'sql = sql & "','" & ObtenerCampo("if5pla", "f5modelo", "f5codpro", rsif3orden!f3codpro, "T", cnn_dbbancos)
            'sql = sql & "','" & ObtenerCampo("if5pla", "f5NOMPRO", "f5codpro", rsif3orden!f3codpro, "T", cnn_dbbancos)
            'sql = sql & "','" & rsif3orden!F5MARCA
            'sql = sql & "'," & rsif3orden!f3canfal & "," & rsif3orden!F3PRECOS & ","
            'sql = sql & rsif3orden!F3TOTAL & ")"
            
            amovs_det(0).campo = "t1item": amovs_det(0).valor = n: amovs_det(0).TIPO = "N"
            amovs_det(1).campo = "T1NumImp": amovs_det(1).valor = txtnumero.Text: amovs_det(1).TIPO = "T"
            amovs_det(2).campo = "T1NumOrd": amovs_det(2).valor = rsif3orden!F4NUMORD & "": amovs_det(2).TIPO = "T"
            amovs_det(3).campo = "T1CodPro": amovs_det(3).valor = rsif3orden!f3codpro & "": amovs_det(3).TIPO = "T"
            amovs_det(4).campo = "T1Modelo": amovs_det(4).valor = ObtenerCampo("if5pla", "f5modelo", "f5codpro", rsif3orden!f3codpro, "T", cnn_dbbancos): amovs_det(4).TIPO = "N"
            amovs_det(5).campo = "T1NomPro": amovs_det(5).valor = rsif3orden!F5NOMPRO & "": amovs_det(5).TIPO = "T"
            amovs_det(6).campo = "T1DesMar": amovs_det(6).valor = rsif3orden!F5MARCA & "": amovs_det(6).TIPO = "T"
            amovs_det(7).campo = "T1Cantidad": amovs_det(7).valor = Val(rsif3orden!F3CANPRO & ""): amovs_det(7).TIPO = "N"
            amovs_det(8).campo = "T1FobUni": amovs_det(8).valor = Val(rsif3orden!F3PRECOS & ""): amovs_det(8).TIPO = "N"
            amovs_det(9).campo = "T1FobTot": amovs_det(9).valor = Val(rsif3orden!F3TOTAL & ""): amovs_det(9).TIPO = "N"
            
            GRABA_REGISTRO amovs_det, "TMP_IMP_DET1", "A", 9, tempo, ""
            
            'tempo.Execute sql
            n = n + 1
            rsif3orden.MoveNext
        Loop
        sw_nuevo_item = False
    End If
    rsif3orden.Close
End If
rsif4orden.Close

End Sub

Private Sub Agrega_Items()
Dim WTOTAL   As Double
Dim CONT     As Integer
Dim rsproduc    As New ADODB.Recordset

    Set TbDetOrden1 = New ADODB.Recordset
        
    WTOTAL = 0#: I% = 0: CONT = 1
    
    'SQL = "Select F5CODPRO,F5NomPro,F7CodMed,F5VALVTA,F5PARTARA,F5CODFAB,F5MARCA from IF5PLA"
    'If rsproduc.State = adStateOpen Then rsproduc.Close
    'rsproduc.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    xcont = dxDBProductos.Dataset.RecordCount
    xfab = dxDBProductos.Dataset.FieldValues("f3codfab")
    dxDBProductos.Dataset.Close
    
    If xcont = 1 And xfab = "" Then
        DELETEREC_N "TmpDet_Import", tempo
    End If
    Set rst = New ADODB.Recordset
    With dxDBProductos.Dataset
        For I = 1 To hlp_ocompra.Grid1.SelBookmarks.Count
        'hlp_ocompra.Grid1.Bookmark = hlp_ocompra.Grid1.SelBookmarks.item(X)
        hlp_ocompra.Grid1.Bookmark = hlp_ocompra.Grid1.SelBookmarks.ITEM(I - 1)
        xf4codprv = hlp_ocompra.Grid1.Columns(5)
        If rst.State = adStateOpen Then rst.Close
        sql = "select f2nomprov from ef2proveedores where f2newruc='" & xf4codprv & "'"
        rst.Open sql, cnn_dbbancos, adOpenStatic
        If Not rst.EOF Then
            xproveedor = rst("f2nomprov")
        Else
            xproveedor = ""
        End If
        
        If Len(Trim(hlp_ocompra.Grid1.Columns(0))) > 0 Then
            sql = "Select * from IF3ORDEN where F4NUMORD=" & Val(Format(hlp_ocompra.Grid1.Columns(0), "0000000")) & " "
            If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
            TbDetOrden1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not TbDetOrden1.EOF Then
                Do While TbDetOrden1.Fields("F4NumOrd") = Val(Format(hlp_ocompra.Grid1.Columns(0), "0000000"))
                    If Val(Format(TbDetOrden1.Fields("F3CanFal"), "0.000")) > 0# Then
                        Columna1 = "" & Format(hlp_ocompra.Grid1.Columns(0), "0000000")
                        Columna3 = "" & TbDetOrden1.Fields("F3CodPro")
                        Columna4 = "" & TbDetOrden1.Fields("F3CodFab")
                        'rsproduc.Find "F5CODFAB='" & TbDetOrden1.Fields("F3CODFAB") & "'"
                        
                        sql = "Select F5CODPRO,F5NomPro,F7CodMed,F5VALVTA,F5PARTARA,F5CODFAB,F5MARCA from IF5PLA WHERE F5CODFAB='" & TbDetOrden1.Fields("F3CODFAB") & "' and f5marca='" & TbDetOrden1.Fields("f5codmarca") & "'"
                        If rsproduc.State = adStateOpen Then rsproduc.Close
                        rsproduc.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
                        If Not rsproduc.EOF Then
                            Columna5 = "" & Mid(rsproduc.Fields("F5NomPro"), 1, 30)
                            Columna6 = "" & rsproduc.Fields("F7CodMed")
                            Columna12 = 0 & rsproduc.Fields("F5VALVTA")
                            Columna22 = "" & rsproduc.Fields("f5partara")
                            Columna24 = "" & rsproduc.Fields("f5codfab")
                            'Columna25 = "" & rsproduc.Fields("").Value
                            If IsNull(ADVALOREM) Then   'Giannina
                                ADVALOREM = 0#
                            End If
                            
                            wmarca = "" & rsproduc.Fields("f5marca")
                            If rsttemp.State = adStateOpen Then rsttemp.Close
                            sql = "select f2desmar from ef2marcas where f2codmar='" & wmarca & "'"
                            rsttemp.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                            If Not rsttemp.EOF Then
                                Columna25 = "" & rsttemp("f2desmar")
                                Columna26 = wmarca
                            End If
                            rsttemp.Close
                            
                        End If
                        Columna7 = Format(TbDetOrden1.Fields("F3CanFal"), "0.00")
                        Columna8 = Val(Format(TbDetOrden1.Fields("F3Precos"), "0.0000"))
                        Columna9 = Val(Format(TbDetOrden1.Fields("F3Total"), "0.0000"))
                        Columna10 = Format(0, "0.00")
                        Columna11 = Format(0, "0.00")
                        Columna12 = Format(0, "0.00")
                        Columna13 = Format(0, "0.00")
                        Columna14 = Format(TbDetOrden1.Fields("F3CanFal"), "0.00")
                        Columna16 = Format(ADVALOREM * Columna8, "0.0000")
                        
                        Columna23 = "N"
                        If Columna16 = "" Then  'Giannina
                            Columna16 = 0#
                        End If
                        Columna17 = Format(Val(Format(TbDetOrden1.Fields("F3Precos"), "0.0000")) + Format(ADVALOREM * Val(Format(TbDetOrden1.Fields("F3Precos"), "0.0000")), "0.0000"), "0.0000")
                        If Columna17 = "" Then  'Giannina
                            Columna17 = 0#
                        End If
                        Columna20 = xf4codprv
                        Columna21 = xproveedor
                        
                        rsproduc.Close
                        Set rsproduc = Nothing
                        
                        WTOTAL = WTOTAL + Val(Format(TbDetOrden1.Fields("F3Total"), "0.0000"))
                        
                    End If
                    csql = "INSERT INTO " & "TmpDet_Import" & " (F3NUMORD,F3CODFAB,F5CODPRO,F5NOMPRO," & _
                    "F5UNIMED,F3CANTIDAD,F3PREFOB,F3TOTAL,F3PRECOS,F3MARGEN,F3VALVTA,F3DSCTO,CANTIDAD,advalorem,base,F3ITEM1,f2codprov,f2nomprov,f5partara,f5manual,f5marca,f5codmarca) VALUES('" & Columna1 & "','" & _
                    Columna24 & "','" & Columna3 & "','" & Columna5 & "','" & _
                    Columna6 & "'," & Columna7 & "," & Columna8 & "," & _
                    Columna9 & "," & Columna10 & "," & Columna11 & "," & Columna12 & "," & Columna13 & "," & Columna14 & "," & Columna16 & "," & Columna17 & "," & CONT & ",'" & Columna20 & "','" & Columna21 & "','" & Columna22 & "','" & Columna23 & "','" & Columna25 & "','" & Columna26 & "')"
                    tempo.Execute (csql)
                    'AlmacenaQuery_sql csql, tempo
                    
'                    csql = "INSERT INTO " & "TmpDet_Import" & " (F3NUMORD,F3CODFAB,F5CODPRO,F5NOMPRO,F3CANTIDAD," & _
'                    "F3PREFOB,F3TOTAL,F3PRECOS,F3MARGEN,F3VALVTA,F3DSCTO,CANTIDAD,F5UNIMED,F3ITEM1,f2codprov,f2nomprov," & _
'                    "f5partara,f5manual,f5marca,f5codmarca,advalorem,BASE) VALUES('" & _
'                    numorden & "','" & codfabrica & "','" & codproducto & "','" & NOMPRODUCTO & "'," & cantidad1 & _
'                    "," & preciounit & "," & total1 & "," & preccosto & "," & margen & "," & ValVta & "," & descuento & "," & cantidad2 & ",'" & unidad & "'," & CONT & ",'" & F2CODPROV & "','" & F2NOMPROV & _
'                    "','" & f5partara & "','" & f5manual & "','" & F5MARCA & "','" & f5codmarca & "',0,0)"
'                    tempo.Execute (csql)
                    
                    CONT = CONT + 1
                    TbDetOrden1.MoveNext
                    If TbDetOrden1.EOF Then
                        Exit Do
                    End If
                Loop
            End If
        End If
        Next I
        dxDBProductos.Dataset.ADODataset.ConnectionString = tempo
        dxDBProductos.Dataset.Active = True
        dxDBProductos.Dataset.Open
    End With
    'rsproduc.Close
    dxDBProductos.Dataset.Edit
    Calcula_New_Importaciones 7
    dxDBProductos.Dataset.Post
    dxDBProductos.Dataset.First
    dxDBProductos.Columns.FocusedIndex = 0
End Sub
Private Sub txtserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtSerie.Text = Format$(TxtSerie.Text, "000")
    TxtNumFac.SetFocus
End If
End Sub

Private Sub Txtnumfac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtNumFac.Text = Format$(TxtNumFac.Text, "0000000")
    SendKeys "{tab}"
End If
End Sub

Private Sub TxtRefere_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim rsimport_cab    As New ADODB.Recordset

Select Case Tool.Id
    Case "idproforma"
        With rptproforma
            .datos.ConnectionString = cnn_dbbancos
            sql = "SELECT IMPORT_DET.*, IF5PLA.F5MARCA, EF2MARCAS.F2DESMAR, IF5PLA.F7CODMED, IF5PLA.F5TEXTO_ING, IMPORT_CAB.F4PROFORMA " _
            & "FROM IMPORT_CAB INNER JOIN (IMPORT_DET INNER JOIN (IF5PLA LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = " _
            & "EF2MARCAS.F2CODMAR) ON (IMPORT_DET.F5CODMARCA = IF5PLA.F5MARCA) AND (IMPORT_DET.F3CodFab = " _
            & "IF5PLA.F5CODFAB)) ON IMPORT_CAB.F4NumImp = IMPORT_DET.F4NumImp " _
            & " WHERE IMPORT_CAB.F4NUMIMP='" & txtnumero.Text & "' order by f2desmar"
            .Caption = "Proforma"
            .txtnum.Text = txtnumero.Text
            .txtembarcador.Text = TxtNomPrv.Text
            .txtdir.Text = TxtDirPrv.Text
            .txtfecha.Text = aboFecha.Value
            .datos.Source = sql
            .Show vbModal
        End With
    
    Case "idcosteo"
        Call costeo
    
    Case "ID_Nuevo"
         txtnumero.Text = Format(WNUMERO1, "0000000")
         Call PROCEDIMIENTO_NUEVO
    
    Case "ID_Grabar"
        If Trim(TxtNomPrv.Text) = "" Then
            MsgBox "Debe Ingresar Embarcador", vbInformation, "Sistema de Logística"
            TxtRucPrv.SetFocus
            Exit Sub
        End If
        If left(CboVia.Text, 1) = "S" Then
            MsgBox "Debe Seleccionar la Via de Importación", vbInformation, "Sistema de Logística"
            CboVia.SetFocus
            Exit Sub
        End If
        If left(CboImporta.Text, 4) = "0000" Then
            MsgBox "Debe Seleccionar el Importador", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
        If left(Me.CboZona.Text, 2) = "0000" Then
            MsgBox "Debe Seleccionar el Importador", vbInformation, "Sistema de Logística"
            CboImporta.SetFocus
            Exit Sub
        End If
                
        Me.MousePointer = 11
        dxDBMerca.Dataset.Edit
        If dxDBMerca.Dataset.State = dsEdit Or dxDBMerca.Dataset.State = dsInsert Then
             dxDBMerca.Dataset.Post
             sw_detalle1 = True
        End If
        If sw_cabecera1 = True Or sw_detalle1 = True Then
            GRABAR2
            If Not graba Then
                Me.MousePointer = 1
                MsgBox "Debe Seleccionar Los Productos de la Importación", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                Me.MousePointer = 1
                MsgBox "La Importación Nº " & txtnumero.Text & " ha sido Actualizada.", vbInformation, "Sistema de Logística"
            End If
            
            sw_nuevo_doc1 = False
            gfalta = falta
            gcanti = wcantord
            sw_nuevo_doc1 = False
            sw_detalle1 = False
            sw_cabecera1 = False
             
             '------------------- MUESTRA EL VALE DE INGRESO
             'If MsgBox("Desea Generar el Vale de Ingreso", vbInformation + vbYesNo, "Importacion ") = vbYes Then
             '   sw_nuevo_documento = True
             '   sw_importa_valedeingreso = True
             '   vale_ingreso.Show 1
             '   If sw_nuevo_documento = False Then
             '       cnn_dbbancos.Execute "update import_cab set f4numvale='" & cnumvale & "' where f4numimp='" & txtnumero.Text & "'"
             '   End If
             '   sw_importa_valedeingreso = False
             'End If
             '--------------------------------------------------
             
        End If
        Me.MousePointer = 1
        
    Case "ID_OrdendeCompra"
            Me.MousePointer = 11
            sw_ayuda_oc = True
'            If TxtNomPrv.Text = "" Then
'                wopcion = 2
'            Else
'                wopcion = 1
'            End If
            wtipoc = "I"
            importar_ocompra.Show 1
'            hlp_ocompra.wopcion = wopcion
'            hlp_ocompra.Show 1
'            Agrega_Items
            AGREGA_OCIMP
            TxtSerie.SetFocus
            Me.MousePointer = 1
    
    Case "idimprimir"
'        Me.MousePointer = vbHourglass
'        SQL = "Select * from TmpDet_Import order by f5marca"
'        RegImporta2.Caption = "Registro de New_Importaciones"
'        RegImporta2.DataControl1.ConnectionString = tempo
'        RegImporta2.DataControl1.Source = SQL
'        RegImporta2.Label1.Caption = wempresa
'        'RegImporta2.txtproforma.Text = New_Importaciones.txtnumproforma.Text
'        RegImporta2.Label27.Caption = aboFecha.Value
'        RegImporta2.Label28.Caption = New_Importaciones.TxtNomPrv.Text
'        RegImporta2.Label29.Caption = New_Importaciones.TxtDirPrv.Text
'        RegImporta2.Label30.Caption = New_Importaciones.TxtTelPrv.Text
'        RegImporta2.Label31.Caption = New_Importaciones.TxtRefere.Text
'        RegImporta2.Show vbModal
'        Me.MousePointer = vbDefault
'
    Case "idtraduccion"
        Me.MousePointer = vbHourglass
        sql = "Select * from TmpDet_Import order by f5marca"
        RegImporta.DataControl1.ConnectionString = tempo
        RegImporta.DataControl1.Source = sql
        RegImporta.Caption = "Traducción"
        RegImporta.Label1.Caption = wempresa
        'RegImporta.txtproforma.Text = New_Importaciones.txtnumproforma.Text
        RegImporta.Label27.Caption = aboFecha.Value
        RegImporta.Label28.Caption = New_Importaciones.LblNomPrv.Caption
        RegImporta.Label29.Caption = New_Importaciones.LblDirPrv.Caption
        RegImporta.Label30.Caption = New_Importaciones.LblTelPrv.Caption
        RegImporta.Label31.Caption = New_Importaciones.TxtRefere.Text
        RegImporta.Show vbModal
        Me.MousePointer = vbDefault
        
    Case "ID_Borrar"
        Me.MousePointer = 11
        elimina txtnumero.Text
        sw_nuevo_doc1 = True
        factor = Format(0, "0.00000000000000000000")
        nuevo
        dxDBProductos.Dataset.Close
        DELETEREC_N "TmpDet_Import", tempo
        AdicionaItem1
        Me.MousePointer = 1
    
    Case "ID_Calculadora"
        Me.MousePointer = 11
        Dim X As Variant
        X = Shell("Calc.exe", 1)
        Me.MousePointer = 1
        
    Case "ID_Salir"
        Me.MousePointer = 11
        If dxDBProductos.Dataset.State = dsEdit Then
            dxDBProductos.Dataset.Post
            sw_nuevo_item1 = True
        End If

        If sw_cabecera1 = True Or sw_detalle1 = True Then
            If MsgBox("Desea Grabar el Movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                GRABAR2
                sw_nuevo_doc1 = False
                sw_cabecera1 = False
                sw_detalle1 = False
            End If
        End If
        Me.MousePointer = 1
        
        '------------------ VERIFICA SI SE GRABO LA IMPORTACION
        If rsimport_cab.State = adStateOpen Then rsimport_cab.Close
        rsimport_cab.Open "SELECT F4NUMIMP FROM IMPORT_CAB WHERE F4NUMIMP = '" & txtnumero.Text & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
        If rsimport_cab.EOF Then
            
            csql = ("DELETE * FROM TB_COSTEOCAB WHERE F4NUMIMP='" & txtnumero.Text & "'")
            cnn_dbbancos.Execute cqsl
            'AlmacenaQuery_sql csql, cnn_dbbancos
            
            csql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP='" & txtnumero.Text & "'")
            cnn_dbbancos.Execute cqsl
            'AlmacenaQuery_sql csql, cnn_dbbancos
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos
            
        End If
        rsimport_cab.Close
        Set rsimport_cab = Nothing
        '-----------------------------------------------------------
        Unload Me

End Select
End Sub

Public Sub PROCEDIMIENTO_NUEVO()
Me.MousePointer = 11
sw_nuevo_doc1 = False
sw_detalle1 = False

'AdicionaItem1
'AdicionaItem1
llena_items
sw_nuevo_doc1 = True
Me.MousePointer = 1
End Sub

Public Sub nuevo()
'sw_nuevo_doc1 = True
If sw_nuevo_documento = True Then
'*****************
If tempo.State = 1 Then tempo.Close
tempo.Open
sql = "delete * from TMP_IMP_DET1"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
'*****************
sql = "delete * from TMP_IMP_DET2"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
'*****************
sql = "insert into TMP_IMP_DET1(t1numimp,T1ITEM) values ('" & Me.txtnumero.Text & "',1)"
tempo.Execute sql
'AlmacenaQuery_sql sql, tempo
'*****************
TxtRucPrv.Text = ""
LblNomPrv.Caption = ""
LblDirPrv.Caption = ""
LblTelPrv.Caption = ""
Txtcodcli.Text = ""
LblNomCli.Caption = ""
aboFecha.Value = Format(Now, "dd/mm/yyyy")
Me.TxtRefere.Text = ""
TxtSerie.Text = ""
TxtNumFac.Text = ""
TxtRefere.Text = ""
PnlFactorSinFin.Caption = Format(0, "0.00000000000000000000")
PnlFactorConFin.Caption = Format(0, "0.00000000000000000000")
'TxtRucPrv.SetFocus
Me.TxtZonaDet.Text = ""
TxtCant.Text = "0.00"
Me.LblGastoFin.Caption = "0.00"
Me.LblCostoI.Caption = "0.00"
Me.LblCostoS.Caption = "0.00"
Me.lblcosto.Caption = "0.00"
Me.LblIgvI.Caption = "0.00"
Me.LblIgvS.Caption = "0.00"
Me.LblIgv.Caption = "0.00"
Me.LblTotalI.Caption = "0.00"
Me.LblTotalS.Caption = "0.00"
Me.LblTot.Caption = "0.00"
TxtUtil.Text = "0.00"
Me.chkcerrar.Enabled = True
Me.chkcerrar.Value = 0
Me.ChkAprobar.Enabled = True
Me.ChkAprobar.Value = 0
TxtImportacion.Visible = False
LblImportacion.Visible = False
txtnumero.Visible = True
LblNumero.Visible = True
'*****************

'*****************
CargaUM
CargaVias
CargaZonas
CargaImportadores
CargaAgentes
'*****************
'If sw_nuevo_doc1 = True Then
    
'End If
'*****************
If Me.dxDBImport.Dataset.State = 1 Then Me.dxDBImport.Dataset.Close
'CargaImport
'GridImport

GridMercancias
CargaGastos
GridGastos
End If
'sw_nuevo_doc1 = False
sw_nuevo_documento = False
'dxDBImport.Dataset.ADODataset.Requery
End Sub

Private Sub GRABAR2()
Dim intcont  As Integer
Dim wtotord  As Double
Dim TOTA As Double

    Set RSDETALLE = New ADODB.Recordset
    Set TbCabImport1 = New ADODB.Recordset
    Set tbProducto1 = New ADODB.Recordset
    Set TbCabOrden1 = New ADODB.Recordset
    Set TbDetImport1 = New ADODB.Recordset
    Set TbDetOrden1 = New ADODB.Recordset
    Set TbDetTmpImp1 = New ADODB.Recordset
    
    wtotord = 0#
    intcont = 0#
    
    If sw_nuevo_doc1 = True Then
        ctipo = "A"
    Else
        ctipo = "M"
    End If
CLIENTE = ObtenerCampo("ef2clientes", "f2codcli", "f2newruc", Me.Txtcodcli.Text, "T", cnn_dbbancos)
'---------------------ASIGNA DATOS A LA CABECERA DE IMPORT_CAB
amovs_cab(0).campo = "F4NUMIMP": amovs_cab(0).valor = txtnumero.Text: amovs_cab(0).TIPO = "T"
amovs_cab(1).campo = "F4CODPRV": amovs_cab(1).valor = ObtenerCampo("EF2PROVEEDORES", "F2CODPROV", "F2NEWRUC", Trim(Me.TxtRucPrv.Text), "T", cnn_dbbancos): amovs_cab(1).TIPO = "T"
amovs_cab(2).campo = "F4FECHA": amovs_cab(2).valor = aboFecha.Value: amovs_cab(2).TIPO = "F"
amovs_cab(3).campo = "F4REFERE": amovs_cab(3).valor = IIf(Len(Trim(TxtRefere.Text)) = 0, " ", TxtRefere.Text): amovs_cab(3).TIPO = "T"
amovs_cab(4).campo = "F4CLIENTE": amovs_cab(4).valor = ObtenerCampo("ef2clientes", "f2codcli", "f2newruc", Me.Txtcodcli.Text, "T", cnn_dbbancos): amovs_cab(4).TIPO = "T"
amovs_cab(5).campo = "F4SERIE": amovs_cab(5).valor = IIf(Len(Trim(TxtSerie.Text)) = 0, " ", left(TxtSerie.Text, 3)): amovs_cab(5).TIPO = "T"
amovs_cab(6).campo = "F4NUMFAC": amovs_cab(6).valor = IIf(Len(Trim(TxtSerie.Text)) = 0, " ", Mid(TxtNumFac.Text, 1, 7)): amovs_cab(6).TIPO = "T"
amovs_cab(7).campo = "F4fobtot": amovs_cab(7).valor = dxDBMerca.Columns.ColumnByFieldName("T1fobtot").SummaryFooterValue: amovs_cab(7).TIPO = "N"
amovs_cab(8).campo = "F4ZONA": amovs_cab(8).valor = left(CboZona.Text, 2): amovs_cab(8).TIPO = "T"
amovs_cab(9).campo = "F4VIA": amovs_cab(9).valor = left(CboVia.Text, 1): amovs_cab(9).TIPO = "T"
amovs_cab(10).campo = "F4UNDMED": amovs_cab(10).valor = left(Me.CboUm.Text, 2): amovs_cab(10).TIPO = "T"
amovs_cab(11).campo = "F4MEDIDA": amovs_cab(11).valor = Val(Format(TxtCant.Text, "#0.00")): amovs_cab(11).TIPO = "N"
amovs_cab(12).campo = "F4ZONADES": amovs_cab(12).valor = Me.TxtZonaDet.Text: amovs_cab(12).TIPO = "T"
amovs_cab(13).campo = "F4costoi": amovs_cab(13).valor = Val(Format(LblCostoI.Caption, "#0.00")): amovs_cab(13).TIPO = "N"
amovs_cab(14).campo = "F4igvi": amovs_cab(14).valor = Val(Format(LblIgvI.Caption, "#0.00")): amovs_cab(14).TIPO = "N"
amovs_cab(15).campo = "F4toti": amovs_cab(15).valor = Val(Format(LblTotalI.Caption, "#0.00")): amovs_cab(15).TIPO = "N"
amovs_cab(16).campo = "F4Costos": amovs_cab(16).valor = Val(Format(LblCostoS.Caption, "#0.00")): amovs_cab(16).TIPO = "N"
amovs_cab(17).campo = "F4igvs": amovs_cab(17).valor = Val(Format(LblIgvS.Caption, "#0.00")): amovs_cab(17).TIPO = "N"
amovs_cab(18).campo = "F4tots": amovs_cab(18).valor = Val(Format(LblTotalS.Caption, "#0.00")): amovs_cab(18).TIPO = "N"
amovs_cab(19).campo = "F4gasfin": amovs_cab(19).valor = Format(LblGastoFin.Caption, "#0.00"): amovs_cab(19).TIPO = "N"
amovs_cab(20).campo = "F4IMPORTADOR": amovs_cab(20).valor = right(CboImporta.Text, 4): amovs_cab(20).TIPO = "T"
amovs_cab(21).campo = "F4factorsinfin": amovs_cab(21).valor = Val(PnlFactorSinFin.Caption): amovs_cab(21).TIPO = "N"
amovs_cab(22).campo = "F4factorconfin": amovs_cab(22).valor = Val(PnlFactorConFin.Caption): amovs_cab(22).TIPO = "N"
If Me.chkcerrar.Value = 1 Then
    westado = "S"
ElseIf Me.chkcerrar.Value = 0 Then
    westado = "N"
End If
amovs_cab(23).campo = "F4cerrado": amovs_cab(23).valor = westado: amovs_cab(23).TIPO = "T"
amovs_cab(24).campo = "F4utilidad": amovs_cab(24).valor = Format(TxtUtil.Text, "#0.00"): amovs_cab(24).TIPO = "N"
amovs_cab(25).campo = "F4numori": amovs_cab(25).valor = (TxtNumOri.Text): amovs_cab(25).TIPO = "T"
csql = "select * from TMP_IMP_DET2 where t2key='00010002'"
If tempo.State = 0 Then tempo.Open
If Rs.State = 1 Then Rs.Close
Rs.Open csql, tempo, 3, 1
If Rs.RecordCount > 0 Then
    nFlete = Val(Rs!t2total & "")
Else
    nFlete = 0
End If
amovs_cab(26).campo = "F4flete": amovs_cab(26).valor = nFlete: amovs_cab(26).TIPO = "N"
amovs_cab(27).campo = "F4AGENTE": amovs_cab(27).valor = right(CboAgente.Text, 4): amovs_cab(27).TIPO = "T"
'-----Graba Cabecera
sql = "select * from import_cab where f4numimp='" & txtnumero.Text & "'"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
If Rs.RecordCount > 0 Then
    ctipo = "M"
Else
    ctipo = "A"
End If
If ctipo = "A" Then '---Nuevo
    GRABA_REGISTRO amovs_cab(), "IMPORT_CAB", ctipo, 27, cnn_dbbancos, ""
    'WNUMERO1 = Val(Format(WNUMERO1, "0000000")) + 1
Else    '----------Modificacion
    GRABA_REGISTRO amovs_cab(), "IMPORT_CAB", ctipo, 27, cnn_dbbancos, "F4NUMIMP = '" & txtnumero.Text & "'"
End If
'*******************************************************************************
'*******************************************************************************
sql = "delete * from Import_Det1 where f4numimp='" & Me.txtnumero.Text & "'"
cnn_dbbancos.Execute sql
'AlmacenaQuery_sql sql, cnn_dbbancos
'GRABA DETALLE 1
sql = "Select * from TMP_IMP_DET1"
If TbCabOrden1.State = 1 Then TbCabOrden1.Close
TbCabOrden1.Open sql, tempo, 3, 1
If TbCabOrden1.RecordCount > 0 Then
    I = 1
    TbCabOrden1.MoveFirst
    Do While Not TbCabOrden1.EOF
        '---------------------ASIGNA DATOS AL DETALLE1
        amovs_det(0).campo = "F4NUMIMP": amovs_det(0).valor = txtnumero.Text & "": amovs_det(0).TIPO = "T"
        amovs_det(1).campo = "F3NUMORD": amovs_det(1).valor = TbCabOrden1!T1NumOrd & "": amovs_det(1).TIPO = "T"
        amovs_det(2).campo = "F5CODPRO": amovs_det(2).valor = TbCabOrden1!T1CodPro & "": amovs_det(2).TIPO = "T"
        amovs_det(3).campo = "F3CANTIDAD": amovs_det(3).valor = TbCabOrden1!T1Cantidad: amovs_det(3).TIPO = "N"
        amovs_det(4).campo = "F3FOBUNI": amovs_det(4).valor = TbCabOrden1!T1FobUni: amovs_det(4).TIPO = "N"
        amovs_det(5).campo = "F3FOBTOT": amovs_det(5).valor = TbCabOrden1!T1FobTot: amovs_det(5).TIPO = "N"
        amovs_det(6).campo = "F3COSUNI": amovs_det(6).valor = TbCabOrden1!T1CosUni: amovs_det(6).TIPO = "N"
        amovs_det(7).campo = "F3VTAUNI": amovs_det(7).valor = TbCabOrden1!t1vtauni: amovs_det(7).TIPO = "N"
        amovs_det(8).campo = "F3MARGEN": amovs_det(8).valor = TbCabOrden1!T1Margen: amovs_det(8).TIPO = "N"
        amovs_det(9).campo = "F3Merca": amovs_det(9).valor = TbCabOrden1!T1NomPro & "": amovs_det(9).TIPO = "T"
        amovs_det(10).campo = "F3item": amovs_det(10).valor = I: amovs_det(10).TIPO = "N"
        
        cvalores = "111111111111111111111111"
        cmes = Format(Month(aboFecha.Value), "00")
        
        '------- GRABA DETALLE '11
        'cnn_dbbancos.Execute ("DELETE * FROM IMPORT_DET WHERE F4NUMIMP = '" & txtnumero.Text & "'")
        'GRABA_REGISTRO_DET amovs_det(), "IMPORT_DET", "A", 8, cnn_dbbancos, "F4NUMIMP  = '" & txtnumero.Text & "'", Values(), nfila - 1, cvalores, cmes, ""
        GRABA_REGISTRO amovs_det(), "IMPORT_det1", "A", 10, cnn_dbbancos, txtnumero.Text

        TbCabOrden1.MoveNext: I = I + 1
    Loop
End If
'****************************************************************************
sql = "delete * from Import_Det2 where f4numimp='" & Me.txtnumero.Text & "'"
cnn_dbbancos.Execute sql
'AlmacenaQuery_sql sql, cnn_dbbancos
'GRABA DETALLE 2
sql = "Select * from TMP_IMP_DET2 WHERE VAL(T2TOTAL & '')>0"
If TbDetImport1.State = 1 Then TbDetImport1.Close
TbDetImport1.Open sql, tempo, 3, 1
If TbDetImport1.RecordCount > 0 Then
    TbDetImport1.MoveFirst
    Do While Not TbDetImport1.EOF
        '---------------------ASIGNA DATOS AL DETALLE2
        amovs_det1(0).campo = "F4NUMIMP": amovs_det1(0).valor = txtnumero.Text: amovs_det1(0).TIPO = "T"
        amovs_det1(1).campo = "F4Grupo": amovs_det1(1).valor = TbDetImport1!T2Grupo: amovs_det1(1).TIPO = "T"
        amovs_det1(2).campo = "F4SubGrupo": amovs_det1(2).valor = TbDetImport1!T2SUBGRUPO: amovs_det1(2).TIPO = "T"
        amovs_det1(3).campo = "F4Costo": amovs_det1(3).valor = IIf(Val(TbDetImport1!t2COSTO & "") = 0, "Null", TbDetImport1!t2COSTO): amovs_det1(3).TIPO = "N"
        amovs_det1(4).campo = "F4Igv": amovs_det1(4).valor = IIf(Val(TbDetImport1!t2igv & "") = 0, "Null", TbDetImport1!t2igv): amovs_det1(4).TIPO = "N"
        amovs_det1(5).campo = "F4Total": amovs_det1(5).valor = IIf(Val(TbDetImport1!t2total & "") = 0, "Null", TbDetImport1!t2total): amovs_det1(5).TIPO = "N"
        amovs_det1(6).campo = "F4Inciden": amovs_det1(6).valor = IIf(Val(TbDetImport1!T2INCIDEN & "") = 0, "NUll", TbDetImport1!T2INCIDEN): amovs_det1(6).TIPO = "N"
        amovs_det1(7).campo = "F4CALCULAR": amovs_det1(7).valor = IIf(TbDetImport1!T2Calcular = True, -1, 0): amovs_det1(7).TIPO = "N"
        amovs_det1(8).campo = "F4esp": amovs_det1(8).valor = IIf(Val(TbDetImport1!t2esp & "") = 0, "Null", TbDetImport1!t2esp): amovs_det1(8).TIPO = "N"
           
        GRABA_REGISTRO amovs_det1(), "IMPORT_det2", "A", 8, cnn_dbbancos, txtnumero.Text
        TbDetImport1.MoveNext
        graba = True
    Loop
End If
'****************************************************************************
sql = "delete * from Import_Det3 where f3numimp='" & Me.txtnumero.Text & "'"
cnn_dbbancos.Execute sql
'AlmacenaQuery_sql sql, cnn_dbbancos
'GRABA DETALLE 3
sql = "Select * from TMP_IMP_DET3 WHERE VAL(T3costo & '')>0"
If TbDetImport1.State = 1 Then TbDetImport1.Close
TbDetImport1.Open sql, tempo, 3, 1
If TbDetImport1.RecordCount > 0 Then
    TbDetImport1.MoveFirst
    Do While Not TbDetImport1.EOF
        '---------------------ASIGNA DATOS AL DETALLE2
        amovs_det1(0).campo = "F3NUMIMP": amovs_det1(0).valor = txtnumero.Text: amovs_det1(0).TIPO = "T"
        amovs_det1(1).campo = "F3Grupo": amovs_det1(1).valor = TbDetImport1!T3Grupo: amovs_det1(1).TIPO = "T"
        amovs_det1(2).campo = "F3SubGrupo": amovs_det1(2).valor = TbDetImport1!T3SUBGRUPO: amovs_det1(2).TIPO = "T"
        amovs_det1(3).campo = "F3Costo": amovs_det1(3).valor = IIf(Val(TbDetImport1!t3COSTO & "") = 0, "Null", TbDetImport1!t3COSTO): amovs_det1(3).TIPO = "N"
        amovs_det1(4).campo = "F3Igv": amovs_det1(4).valor = IIf(Val(TbDetImport1!t3igv & "") = 0, "Null", TbDetImport1!t3igv): amovs_det1(4).TIPO = "N"
        amovs_det1(5).campo = "F3Total": amovs_det1(5).valor = IIf(Val(TbDetImport1!T3TOTAL & "") = 0, "Null", TbDetImport1!T3TOTAL): amovs_det1(5).TIPO = "N"
        amovs_det1(6).campo = "F3Inciden": amovs_det1(6).valor = IIf(Val(TbDetImport1!T3INCIDEN & "") = 0, "NUll", TbDetImport1!T3INCIDEN): amovs_det1(6).TIPO = "N"
        amovs_det1(7).campo = "F3CALCULAR": amovs_det1(7).valor = IIf(TbDetImport1!T3Calcular = True, -1, 0): amovs_det1(7).TIPO = "N"
        amovs_det1(8).campo = "F3meses": amovs_det1(8).valor = IIf(Val(TbDetImport1!t3meses & "") = 0, "Null", TbDetImport1!t3meses): amovs_det1(8).TIPO = "N"
        amovs_det1(9).campo = "F3interes": amovs_det1(9).valor = IIf(Val(TbDetImport1!t3interes & "") = 0, "Null", TbDetImport1!t3interes): amovs_det1(9).TIPO = "N"
                
        GRABA_REGISTRO amovs_det1(), "IMPORT_det3", "A", 9, cnn_dbbancos, txtnumero.Text
        TbDetImport1.MoveNext
        graba = True
    Loop
End If
End Sub

Private Sub Setea_Import()

    For I% = 0 To 9
        wimporta(I%).Orden = ""
        wimporta(I%).f4falta = "0"
    Next I%

End Sub

Public Sub GRABACIONES()
Dim RSDETALLE As ADODB.Recordset
Set RSDETALLE = New ADODB.Recordset
Dim rspregunta As ADODB.Recordset
Set rspregunta = New ADODB.Recordset

sql = "Select * from TB_COSTEOCAB where F4NUMIMP = '" & txtnumero.Text & "'"
If rspregunta.State = adStateOpen Then rspregunta.Close
rspregunta.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic

If Not rspregunta.EOF Then
    ctipo = "M"
Else
    ctipo = "A"
   
End If

'Asignacion de Datos para la Cabecera
amovs_cab1(0).campo = "F4NUMIMP": amovs_cab1(0).valor = txtnumero.Text: amovs_cab1(0).TIPO = "T"
amovs_cab1(1).campo = "F4FACTOR": amovs_cab1(1).valor = factor: amovs_cab1(1).TIPO = "N"
amovs_cab1(2).campo = "F4TIPCAM": amovs_cab1(2).valor = WTipoCambio: amovs_cab1(2).TIPO = "N"
amovs_cab1(3).campo = "F4SERFAC": amovs_cab1(3).valor = IIf(Len(Trim(TxtSerie.Text)) = 0, " ", TxtSerie.Text): amovs_cab1(3).TIPO = "T"
amovs_cab1(4).campo = "F4NUMFAC": amovs_cab1(4).valor = IIf(Len(Trim(TxtNumFac.Text)) = 0, " ", TxtNumFac.Text): amovs_cab1(4).TIPO = "T"
amovs_cab1(5).campo = "F4FECHA": amovs_cab1(5).valor = aboFecha.Value: amovs_cab1(5).TIPO = "F"
amovs_cab1(6).campo = "F4TOTSOL": amovs_cab1(6).valor = dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").SummaryFooterValue: amovs_cab1(6).TIPO = "N"
amovs_cab1(7).campo = "F4TOTDOL": amovs_cab1(7).valor = dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").SummaryFooterValue: amovs_cab1(7).TIPO = "N"

'Asignacion de Datos para el Detalle
amovs_det1(0).campo = "F2CODIGO": amovs_det1(0).valor = "": amovs_det1(0).TIPO = "T"
amovs_det1(1).campo = "F3PRESUPUESTO": amovs_det1(1).valor = "": amovs_det1(1).TIPO = "N"
amovs_det1(2).campo = "F3SOLES": amovs_det1(2).valor = "": amovs_det1(2).TIPO = "N"
amovs_det1(3).campo = "F3DOLAR": amovs_det1(3).valor = "": amovs_det1(3).TIPO = "N"
amovs_det1(4).campo = "F4NUMIMP": amovs_det1(4).valor = "": amovs_det1(4).TIPO = "T"

'CALCULA NUMERO DE FILAS
nitems = 0
If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
'SQL = "SELECT COUNT(F5CODPRO) AS NTIPO FROM tmp_costos WHERE LEN(TRIM(F5CODPRO)) > 0 "
sql = "SELECT COUNT(F5CODPRO) AS NTIPO FROM tmp_costos WHERE F3CHECK=TRUE "
RSDETALLE.Open sql, Temp, adOpenDynamic, adLockOptimistic

If Not RSDETALLE.EOF Then
    ntipo = Val("" & RSDETALLE.Fields("NTIPO"))
End If
RSDETALLE.Close

ReDim Values(5, ntipo)

If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
RSDETALLE.Open "SELECT * FROM tmp_costos WHERE F3CHECK=TRUE", Temp
'RSDETALLE.Open "SELECT * FROM tmp_costos ", Temp
If Not RSDETALLE.EOF Then
     nfil = 0
     RSDETALLE.MoveFirst
     Do While Not RSDETALLE.EOF
         If Len(Trim(RSDETALLE.Fields("F5CODPRO") & "")) > 0 Then
             Values(0, nfil) = RSDETALLE.Fields("F5CODPRO") & ""
             Values(1, nfil) = RSDETALLE.Fields("F3PRESUPUESTO") & ""
             Values(2, nfil) = RSDETALLE.Fields("F3SOLES") & ""
             Values(3, nfil) = RSDETALLE.Fields("F3DOLARES") & ""
             Values(4, nfil) = txtnumero.Text
             nfil = nfil + 1
        End If
        RSDETALLE.MoveNext
     Loop
 End If

RSDETALLE.Close
cvalores = "11111"

If ctipo = "A" Then
    '------- GRABA CABECERA
    GRABA_REGISTRO amovs_cab1(), "TB_COSTEOCAB", ctipo, 7, cnn_dbbancos, ""

    If sw_graba_registro = True Then
        '------- GRABA DETALLE
        cmes = Format(Month(Date), "00") 'ojo esto va mas arriba
        
        GRABA_REGISTRO_DET amovs_det1(), "TB_COSTEODET", ctipo, 4, cnn_dbbancos, "", Values(), nfil - 1, cvalores, cmes, ""
    End If

Else
    '------- GRABA CABECERA
    GRABA_REGISTRO amovs_cab1(), "TB_COSTEOCAB", ctipo, 7, cnn_dbbancos, "F4NUMIMP = '" & txtnumero.Text & "'"

    '------- GRABA DETALLE
    sql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP = '" & txtnumero.Text & "'")
    cnn_dbbancos.Execute sql
    'AlmacenaQuery_sql sql, cnn_dbbancos
    
    GRABA_REGISTRO_DET amovs_det1(), "TB_COSTEODET", "A", 4, cnn_dbbancos, "F4NUMIMP  = '" & txtnumero.Text & "'", Values(), nfil - 1, cvalores, cmes, ""
End If
End Sub

Private Sub Calcula_Costos()
Dim TbDetTmpImp1 As New ADODB.Recordset
Dim wprecos, wprevta   As Double
Dim wutilidad          As Double
I% = 0
wprecos = 0#: wprevta = 0#: wutilidad = 0#

sql = "Select * from TmpDet_Import"
If TbDetTmpImp1.State = adStateOpen Then TbDetTmpImp1.Close
TbDetTmpImp1.Open sql, tempo, 3, 1
'TbDetTmpImp1.MoveFirst
Do While Not TbDetTmpImp1.EOF
    sql = "update TmpDet_Import set f3PRECOS="
    sql = sql & Format(Val(Format(TbDetTmpImp1!F3PREFOB, "0.0000")), "0.0000")
    sql = sql & " ,f3costototal="
    sql = sql & Format(TbDetTmpImp1!F3PRECOS * TbDetTmpImp1!F3CANTIDAD, "0.0000")
    sql = sql & " ,f3total="
    sql = sql & Format(TbDetTmpImp1!F3PRECOS * TbDetTmpImp1!F3CANTIDAD, "0.0000")
    sql = sql & " where f5codpro='" & TbDetTmpImp1!f5codpro & "'"
    tempo.Execute sql
    'AlmacenaQuery_sql sql, tempo
    'TbDetTmpImp1!f3PRECOS = Format(Val(Format(TbDetTmpImp1!F3PREFOB, "0.0000")) * Val(Format(PnlFactor.Caption, "0.000000000000000")), "###,##0.0000")
    'TbDetTmpImp1!f3costototal = Format(TbDetTmpImp1!f3PRECOS * TbDetTmpImp1!F3CANTIDAD, "0.0000")
    'TbDetTmpImp1!F3VALVTA = Format(Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * (1 + Val(Format(TbDetTmpImp1!F3MARGEN, "0.0000")) / 100), "###,##0.0000")
    'TbDetTmpImp1!F3VALVTA = Format(Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * (Val(Format(TbDetTmpImp1!F5FACTOR, "0.0000"))), "###,##0.0000")
    'TbDetTmpImp1!F3VTANET = Format(Val(Format(TbDetTmpImp1!F3VALVTA, "0.0000")) * (1 - Val(Format(TbDetTmpImp1!F3DSCTO, "0.0000")) / 100), "###,##0.0000")
    'TbDetTmpImp1!F3PREUNI = Format(Val(Format(TbDetTmpImp1!F3VTANET, "0.0000")) + (Val(Format(TbDetTmpImp1!F3VTANET, "0.0000")) * 0.19), "###,##0.0000")
    'TbDetTmpImp1.UpdateBatch
    'wprecos = wprecos + Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.0000"))
    'wprevta = wprevta + Val(Format(TbDetTmpImp1!F3VALVTA, "0.0000")) * Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.0000"))
    wprecos = wprecos + (TbDetTmpImp1!F3PRECOS * TbDetTmpImp1!F3CANTIDAD)
    wprevta = wprevta + (TbDetTmpImp1!F3VALVTA * TbDetTmpImp1!F3CANTIDAD)
    TbDetTmpImp1.MoveNext
Loop
dxDBProductos.Dataset.Edit
dxDBProductos.Dataset.Post
'PnlPreCost.Caption = Format$(wprecos, "###,##0.0000")
'PnlPreVta.Caption = Format$(wprevta, "###,##0.0000")
wutilidad = wprevta - wprecos
'PnlUtilidad.Caption = Format$(wutilidad, "###,##0.0000")
End Sub

Private Sub elimina(pnumero As String)
Dim cnumero As String
Set TbCabOrden1 = New ADODB.Recordset
Set TbDetOrden1 = New ADODB.Recordset
Set TbDetTmpImp1 = New ADODB.Recordset

If Len(Trim("" & txtnumero.Text)) = 0 Then
    MsgBox "El Numero de Importacion no ha sido grabado. Verifique", vbCritical, "Atencion"
    Exit Sub
End If

If MsgBox("Está seguro(a) de eliminar la Importacion ?", vbYesNo, "Atencion") = vbYes Then
    
    csql = ("DELETE * FROM IMPORT_CAB WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    
    '-------------Busca en Movimientos
    csql = ("DELETE * FROM IMPORT_DET1 WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    
    
    csql = ("DELETE * FROM IMPORT_DET2 WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos

    
    
    sw_ayuda_oc = True

End If
End Sub
Private Sub Calcula_New_Importaciones(pfoco As Integer)
Dim fob         As Double
Dim precos      As Double
Dim ValVta      As Double
Dim vtaneta     As Double
Dim preuni      As Double
Dim margen      As Double
Dim costo       As Double
    With dxDBProductos
        fob = Val(Format(.Columns.ColumnByFieldName("F3PREFOB").Value, "0.00"))        '8
        If fob > 0 Then
            Select Case pfoco
                Case 6
                    costo = Val(.Columns(7).Value) * (Val(.Columns(8).Value))
                    .Columns(12).Value = Format$(costo, "####,##0.0000")
                Case 7, 8, 10
                    'costo = Val(.Columns.ColumnByFieldName("F3CANTIDAD").Value) * (Val(.Columns.ColumnByFieldName("F5ADVALOREM").Value))           '7-8
                    wcantidad = Val(.Columns.ColumnByFieldName("F3CANTIDAD").Value)
                    wcosto = wcantidad * (Val(.Columns.ColumnByFieldName("F3PREFOB").Value))           '7-8
                    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(wcosto, "####,##0.0000")       '12
                    wprecos = Val(Format(.Columns.ColumnByFieldName("F3PREFOB").Value, "0.0000")) * Val(Format(PnlFactor.Caption, "0.00000000000000000000"))    '8
                    .Columns.ColumnByFieldName("F3PRECOS").Value = Format(wprecos, "###,##0.0000")     '13
                    wcostototal = wprecos * wcantidad
                    .Columns.ColumnByFieldName("F3COSTOTOTAL").Value = Format(wcostototal, "###,##0.0000")
                    
                    'ValVta = Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) * (1 + Val(Format(.Columns.ColumnByFieldName("F3MARGEN").Value, "0.00000")) / 100) '13-14
                    'ValVta = Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) * (Val(Format(.Columns.ColumnByFieldName("F5FACTOR").Value, "0.00000"))) '13-14
                    '.Columns.ColumnByFieldName("F3VALVTA").Value = Format(ValVta, "###,##0.0000")     '15
                    'vtaneta = Val(Format(.Columns.ColumnByFieldName("F3VALVTA").Value, "0.0000")) * (1 - Val(Format(.Columns.ColumnByFieldName("F3DSCTO").Value, "0.0000")) / 100) '15-16
                    '.Columns.ColumnByFieldName("F3VTANET").Value = Format(vtaneta, "###,##0.0000")        '17
                    'preuni = Val(Format(.Columns.ColumnByFieldName("F3VTANET").Value, "0.0000")) + (Val(Format(.Columns.ColumnByFieldName("F3VTANET").Value, "0.0000")) * (wigv / 100))   '17-17
                    '.Columns.ColumnByFieldName("F3PREUNI").Value = Format(preuni, "###,##0.0000")     '19
                Case 12
                    'ValVta = Val(Format(.Columns(13).Value, "0.0000")) * (1 + Val(Format(.Columns(14).Value, "0.0000")) / 100)
                    '.Columns(15).Value = Format(ValVta, "###,##0.0000")
                    'vtaneta = Val(Format(.Columns(15).Value, "0.0000")) * (1 - Val(Format(.Columns(16).Value, "0.0000")) / 100)
                    '.Columns(17).Value = Format(vtaneta, "###,##0.0000")
                    'preuni = Val(Format(.Columns(17).Value, "0.0000")) * (1 + (gigv / 100))
                    '.Columns(19).Value = Format(preuni, "###,##0.0000")
                Case 14
                    'If .Columns(16).Value > .Columns(14).Value Then
                    '    MsgBox "Error %Dscto debe ser menor al %Ganancia", vbInformation, "Atencion"
                    '    .Columns(16).Value = Format(0, "0.0000")
                    'Else
                    '    vtaneta = Val(Format(.Columns(15).Value, "0.0000")) * (1 - Val(Format(.Columns(16).Value, "0.0000")) / 100)
                    '    .Columns(17).Value = Format(vtaneta, "###,##0.0000")
                    '    preuni = Val(Format(.Columns(17).Value, "0.0000")) * (1 + (gigv / 100))
                    '    .Columns(19).Value = Format(preuni, "###,##0.0000")
                    'End If
                Case 15
                    'margen = (Val(Format(.Columns(15).Value, "0.0000")) / Val(Format(.Columns(13).Value, "0.0000")) - 1) * 100
                    '.Columns(14).Value = Format(margen, "###,##0.0000")
                    'ValVta = Val(Format(.Columns(17).Value, "0.0000")) / (1 - Val(Format(.Columns(16).Value, "0.0000")) / 100)
                    '.Columns(15).Value = Format(ValVta, "###,##0.0000")
                    'preuni = Val(Format(.Columns(17).Value, "0.0000")) * (1 + (gigv / 100))
                    '.Columns(19).Value = Format(preuni, "###,##0.0000")
                Case 16
                    'vtaneta = Val(Format(.Columns(19).Value, "0.0000")) / (1 + (gigv / 100))
                    '.Columns(17).Value = Format(vtaneta, "###,##0.0000")
                    'ValVta = Val(Format(.Columns(17).Value, "0.0000")) / (1 - Val(Format(.Columns(16).Value, "0.0000")) / 100)
                    '.Columns(15).Value = Format(ValVta, "###,##0.0000")
                    'margen = (Val(Format(.Columns(15).Value, "0.0000")) / Val(Format(.Columns(13).Value, "0.0000")) - 1) * 100
                    '.Columns(14).Value = Format(margen, "###,##0.0000")
            End Select
        End If
    End With
End Sub
Public Sub activar(Estado)
'Panel3D1.Enabled = Not estado
'Frame3D1.Enabled = Not Estado
chkcerrar.Enabled = Not Estado
ChkAprobar.Enabled = Not Estado


End Sub

Public Sub costeo()
With rptcosteo
    .datos.ConnectionString = cnn_dbbancos
    sql = "SELECT TB_COSTEODET.*, TB_COSTOSIMP.F2DESCRIPCION " _
    & "FROM TB_COSTEODET INNER JOIN TB_COSTOSIMP ON TB_COSTEODET.F2CODIGO = TB_COSTOSIMP.F2CODIGO " _
    & " where f4numimp='" & txtnumero.Text & "'"
    
    .Caption = "Resumen de Importación"
    .lblempresa.Caption = wempresa
    '.lblProforma.Caption = txtnumproforma.Text
    .lblFecha.Caption = aboFecha.Value
    .datos.Source = sql
    .Show vbModal
End With

End Sub

Private Sub TxtTotalI_Change()
CalculaMontos
End Sub

Private Sub TxtTotalS_Change()
CalculaMontos
End Sub

Private Sub TxtUtil_GotFocus()
TxtUtil.Text = Format(TxtUtil.Text, "#0.00")
TxtUtil.SelStart = 0: TxtUtil.SelLength = Len(TxtUtil.Text)
End Sub

Private Sub TxtUtil_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    SendKeys "{tab}"
End Select
End Sub

Private Sub TxtUtil_LostFocus()
TxtUtil.Text = Format(TxtUtil.Text, "#,#0.00")
If Val(Format(TxtUtil.Text, "#0.00")) > 0 Then
    CalculaMontos
    sql = "update TMP_IMP_DET1 set t1vtauni=t1COSuni*" & Val(1 + (Format(TxtUtil.Text, "#0.00") / 100))
    tempo.Execute sql
    'AlmacenaQuery_sql sql, tempo
    sql = "update TMP_IMP_DET1 set t1MARGEN=(t1vtauni-t1cosuni)*100/t1vtauni WHERE t1cosuni>0"
    tempo.Execute sql
    'AlmacenaQuery_sql sql, tempo
End If

End Sub

Private Sub TxtZonaDet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
